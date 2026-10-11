# PDF の暗号化の部品（shared\core\pdf_crypt.ps1）のテスト。SHA は公開されている正解の値、AES-256 の鍵の計算は別に書いた計算と照らす
BeforeDiscovery {
    $r56Cases = @(
        @{ revision = 5; salt = 1 }
        @{ revision = 6; salt = 1 }
        @{ revision = 6; salt = 7 }
        @{ revision = 6; salt = 90 }
    )
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\pdf.ps1"

    function toHexText([byte[]]$bytes) { return (($bytes | ForEach-Object { $_.ToString("x2") }) -join "") }
}

Describe "getPdfSha" -Tag "Unit" {
    It "SHA-256 の abc の値と合う" {
        toHexText (getPdfSha 256 (toPdfBytes "abc")) | Should -Be "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    }

    It "SHA-384・512 の abc の値と合う（公開されている値）" {
        toHexText (getPdfSha 384 (toPdfBytes "abc")) | Should -Be "cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7"
        toHexText (getPdfSha 512 (toPdfBytes "abc")) | Should -Be "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f"
    }
}

Describe "getPdfCryptHashR56" -Tag "Unit" {
    It "R<revision> の鍵のハッシュが、別に書いた計算と合う（塩 <salt>）" -TestCases $r56Cases {
        param ($revision, $salt)
        $saltBytes = [byte[]]($salt..($salt + 7))
        $empty = New-Object byte[] 0
        $actual = getPdfCryptHashR56 $revision $empty $saltBytes $empty
        $actual.Length | Should -Be 32
        toHexText $actual | Should -Be (toHexText (getTestHashR56 $revision $empty $saltBytes))
    }

    It "パスワードがあっても合う" {
        $password = toPdfBytes "secret"
        $saltBytes = [byte[]](1..8)
        toHexText (getPdfCryptHashR56 6 $password $saltBytes (New-Object byte[] 0)) | Should -Be (toHexText (getTestHashR56 6 $password $saltBytes))
    }
}

Describe "unprotectPdfData" -Tag "Unit" {
    It "AES の長さが 16 の倍数でないと、そのファイルだけ例外" {
        $crypt = @{ Key = [byte[]](1..32) }
        { unprotectPdfData $crypt "AESV3" ([byte[]](1..40)) } | Should -Throw "PDFの暗号化されたデータが壊れているため読めません。"
    }

    It "AES の詰め物が壊れていると例外" {
        $crypt = @{ Key = [byte[]](1..32) }
        $iv = [byte[]](1..16)
        $cipher = invokeTestAes $crypt.Key $iv ([byte[]]((1..15) + 0)) $false
        { unprotectPdfData $crypt "AESV3" ([byte[]]($iv + $cipher)) } | Should -Throw "PDFの暗号化されたデータが壊れているため読めません。"
    }

    It "AES で中身が 0 バイトの（空の）ストリームは、空のまま返す" {
        $crypt = @{ Key = [byte[]](1..16) }
        (unprotectPdfData $crypt "AESV3" (New-Object byte[] 0)).Length | Should -Be 0
    }

    It "AES で中身が IV の 16 バイトだけだと、そのファイルだけ例外" {
        $crypt = @{ Key = [byte[]](1..16) }
        { unprotectPdfData $crypt "AESV3" ([byte[]](1..16)) } | Should -Throw "PDFの暗号化されたデータが壊れているため読めません。"
    }

    It "None はそのまま返す" {
        $data = [byte[]](1..5)
        toHexText (unprotectPdfData @{ Key = [byte[]](1..5) } "None" $data) | Should -Be (toHexText $data)
    }
}

Describe "getPdfStreamCryptMethod" -Tag "Unit" {
    BeforeAll {
        $script:cryptDoc = @{
            Cache = @{}
            Crypt = @{
                StmMethod = "AESV3"
                Cf = @{ StdCF = @{ CFM = "/AESV3" }; NoneCF = @{ CFM = "/None" }; RcCF = @{ CFM = "/V2" } }
            }
        }
    }

    It "<name>" -TestCases @(
        @{ name = "ふつうのストリームは既定の方式"; dict = @{}; expected = "AESV3" }
        @{ name = "相互参照のストリームは復号しない"; dict = @{ Type = "/XRef" }; expected = "None" }
        @{ name = "/Crypt の /Identity は復号しない"; dict = @{ Filter = @("/Crypt", "/FlateDecode"); DecodeParms = @(@{ Name = "/Identity" }, $null) }; expected = "None" }
        @{ name = "/Crypt に名前が無ければ復号しない"; dict = @{ Filter = "/Crypt" }; expected = "None" }
        @{ name = "/Crypt の名前の方式にする（AES-256）"; dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = "/StdCF" }) }; expected = "AESV3" }
        @{ name = "/Crypt の名前が /None の方式なら復号しない"; dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = "/NoneCF" }) }; expected = "None" }
    ) {
        param ($name, $dict, $expected)
        getPdfStreamCryptMethod $script:cryptDoc @{ Dict = $dict } | Should -Be $expected
    }

    It "/Crypt の Name が名前でないと、方式が未対応の例外" {
        $stream = @{ Dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = 5.0 }) } }
        { getPdfStreamCryptMethod $script:cryptDoc $stream } | Should -Throw "PDFの暗号化の方式に対応していないため読めません。"
    }

    It "/Crypt が RC4 の方式（対応しない）を指すと、方式が未対応の例外" {
        $stream = @{ Dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = "/RcCF" }) } }
        { getPdfStreamCryptMethod $script:cryptDoc $stream } | Should -Throw "PDFの暗号化の方式に対応していないため読めません。"
    }

    It "/Crypt が知らない名前を指すと、方式が未対応の例外" {
        $stream = @{ Dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = "/Unknown" }) } }
        { getPdfStreamCryptMethod $script:cryptDoc $stream } | Should -Throw "PDFの暗号化の方式に対応していないため読めません。"
    }
}

Describe "getPdfKeyR56" -Tag "Unit" {
    It "R<revision> の鍵が、別に書いた計算（/UE の復号）と合う" -TestCases @(
        @{ revision = 5 }
        @{ revision = 6 }
    ) {
        param ($revision)
        $empty = New-Object byte[] 0
        $key = [byte[]](1..32 | ForEach-Object { ($_ * 11) % 256 })
        $vsalt = [byte[]](21..28)
        $ksalt = [byte[]](31..38)
        $u = joinPdfBytes @((getTestHashR56 $revision $empty $vsalt), $vsalt, $ksalt)
        $ue = invokeTestAes (getTestHashR56 $revision $empty $ksalt) (New-Object byte[] 16) $key $false
        toHexText (getPdfKeyR56 $revision $u $ue) | Should -Be (toHexText $key)
        $u[0] = $u[0] -bxor 1
        { getPdfKeyR56 $revision $u $ue } | Should -Throw "PDFにパスワードが必要なため読めません。"
    }
}
