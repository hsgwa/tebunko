# PDF の暗号化の部品（shared\core\pdf_crypt.ps1）のテスト。MD5・RC4 は自前なので、公開されている正解の値と .NET の結果で確かめる
BeforeDiscovery {
    $md5Cases = @(
        @{ text = ""; expected = "d41d8cd98f00b204e9800998ecf8427e" }
        @{ text = "abc"; expected = "900150983cd24fb0d6963f7d28e17f72" }
        @{ text = "message digest"; expected = "f96b697d7cb7938d525a2f31aaf161d0" }
        @{ text = ("1234567890" * 8); expected = "57edf4a22be3c955ac49da2e2107b67a" }
    )
    $rc4Cases = @(
        @{ key = "Key"; text = "Plaintext"; expected = "bbf316e8d940af0ad3" }
        @{ key = "Wiki"; text = "pedia"; expected = "1021bf0420" }
        @{ key = "Secret"; text = "Attack at dawn"; expected = "45a01f645fc35b383552544b9bf5" }
    )
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

Describe "getPdfRfc1321Hash" -Tag "Unit" {
    It "RFC 1321 の値と合う: <text>" -TestCases $md5Cases {
        param ($text, $expected)
        toHexText (getPdfRfc1321Hash (toPdfBytes $text)) | Should -Be $expected
    }

    It "詰め物の境目（0〜130 バイト）で .NET の MD5 と合う" {
        $bad = @()
        for ($n = 0; $n -le 130; $n++) {
            $data = New-Object byte[] $n
            for ($i = 0; $i -lt $n; $i++) { $data[$i] = ($i * 31 + 7) % 256 }
            $expected = toHexText (getTestDigest "MD5" $data)
            if ((toHexText (getPdfRfc1321Hash $data)) -ne $expected) { $bad += $n }
        }
        $bad | Should -BeNullOrEmpty
    }
}

Describe "invokePdfRc4" -Tag "Unit" {
    It "公開されている値と合う: <key>" -TestCases $rc4Cases {
        param ($key, $text, $expected)
        toHexText (invokePdfRc4 (toPdfBytes $key) (toPdfBytes $text)) | Should -Be $expected
    }

    It "もう一度かけると元に戻る" {
        $data = toPdfBytes "hello world"
        $key = toPdfBytes "k"
        toHexText (invokePdfRc4 $key (invokePdfRc4 $key $data)) | Should -Be (toHexText $data)
    }
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
        { unprotectPdfData $crypt "AESV3" 5 0 ([byte[]](1..40)) } | Should -Throw "PDFの暗号化されたデータが壊れているため読めません。"
    }

    It "AES の詰め物が壊れていると例外" {
        $crypt = @{ Key = [byte[]](1..32) }
        $iv = [byte[]](1..16)
        $cipher = invokeTestAes $crypt.Key $iv ([byte[]]((1..15) + 0)) $false
        { unprotectPdfData $crypt "AESV3" 5 0 ([byte[]]($iv + $cipher)) } | Should -Throw "PDFの暗号化されたデータが壊れているため読めません。"
    }

    It "AES で中身が 0 バイトの（空の）ストリームは、空のまま返す" {
        $crypt = @{ Key = [byte[]](1..16) }
        (unprotectPdfData $crypt "AESV2" 5 0 (New-Object byte[] 0)).Length | Should -Be 0
    }

    It "AES で中身が IV の 16 バイトだけだと、そのファイルだけ例外" {
        $crypt = @{ Key = [byte[]](1..16) }
        { unprotectPdfData $crypt "AESV2" 5 0 ([byte[]](1..16)) } | Should -Throw "PDFの暗号化されたデータが壊れているため読めません。"
    }

    It "None はそのまま返す" {
        $data = [byte[]](1..5)
        toHexText (unprotectPdfData @{ Key = [byte[]](1..5) } "None" 1 0 $data) | Should -Be (toHexText $data)
    }
}

Describe "getPdfStreamCryptMethod" -Tag "Unit" {
    BeforeAll {
        $script:cryptDoc = @{
            Cache = @{}
            Crypt = @{
                StmMethod = "AESV2"; Version = 4; EncryptMetadata = $false
                Cf = @{ StdCF = @{ CFM = "/AESV2" }; NoneCF = @{ CFM = "/None" }; RcCF = @{ CFM = "/V2" } }
            }
        }
    }

    It "<name>" -TestCases @(
        @{ name = "ふつうのストリームは既定の方式"; dict = @{}; expected = "AESV2" }
        @{ name = "相互参照のストリームは復号しない"; dict = @{ Type = "/XRef" }; expected = "None" }
        @{ name = "EncryptMetadata が false のメタデータは復号しない"; dict = @{ Type = "/Metadata" }; expected = "None" }
        @{ name = "/Crypt の /Identity は復号しない"; dict = @{ Filter = @("/Crypt", "/FlateDecode"); DecodeParms = @(@{ Name = "/Identity" }, $null) }; expected = "None" }
        @{ name = "/Crypt に名前が無ければ復号しない"; dict = @{ Filter = "/Crypt" }; expected = "None" }
        @{ name = "/Crypt の名前の方式にする（RC4）"; dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = "/RcCF" }) }; expected = "RC4" }
        @{ name = "/Crypt の名前が /None の方式なら復号しない"; dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = "/NoneCF" }) }; expected = "None" }
    ) {
        param ($name, $dict, $expected)
        getPdfStreamCryptMethod $script:cryptDoc @{ Dict = $dict } | Should -Be $expected
    }

    It "/Crypt の Name が名前でないと、方式が未対応の例外" {
        $stream = @{ Dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = 5.0 }) } }
        { getPdfStreamCryptMethod $script:cryptDoc $stream } | Should -Throw "PDFの暗号化の方式に対応していないため読めません。"
    }

    It "/Crypt が知らない名前を指すと、方式が未対応の例外" {
        $stream = @{ Dict = @{ Filter = @("/Crypt"); DecodeParms = @(@{ Name = "/Unknown" }) } }
        { getPdfStreamCryptMethod $script:cryptDoc $stream } | Should -Throw "PDFの暗号化の方式に対応していないため読めません。"
    }
}

Describe "getPdfKeyR234" -Tag "Unit" {
    # 別に書いた計算（.NET の MD5）で鍵と /U を作り、読む側の鍵と合うかを見る。公開されている鍵の正解の値は無い（仕様に例が無い）
    BeforeAll {
        $script:pad = [byte[]]@(0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41, 0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08, 0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80, 0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A)
        $script:owner = [byte[]](1..32 | ForEach-Object { $_ + 100 })
        $script:id = [byte[]](1..16 | ForEach-Object { $_ * 7 })
        function newExpectedKey($revision, $n, $withFf) {
            $parts = @($script:pad, $script:owner, [byte[]](0xFC, 0xFF, 0xFF, 0xFF), $script:id)
            if ($withFf) { $parts += , ([byte[]](255, 255, 255, 255)) }
            $h = getTestDigest 'MD5' (joinPdfBytes $parts)
            if ($revision -ge 3) { for ($i = 0; $i -lt 50; $i++) { $h = getTestDigest 'MD5' ([byte[]]$h[0..($n - 1)]) } }
            return , ([byte[]]$h[0..($n - 1)])
        }
        function newExpectedU($revision, $key) {
            if ($revision -eq 2) { return , (newTestRc4 $key $script:pad) }
            $x = newTestRc4 $key (getTestDigest 'MD5' (joinPdfBytes @($script:pad, $script:id)))
            for ($i = 1; $i -le 19; $i++) { $x = newTestRc4 ([byte[]]($key | ForEach-Object { $_ -bxor $i })) $x }
            return , (joinPdfBytes @($x, [byte[]](1..16)))
        }
    }

    It "R<revision>（鍵 <n> バイト・FF×4 <withFf>）の鍵が、別に書いた計算と合う" -TestCases @(
        @{ revision = 2; n = 5; withFf = $false }
        @{ revision = 3; n = 16; withFf = $false }
        @{ revision = 3; n = 8; withFf = $false }
        @{ revision = 4; n = 16; withFf = $false }
        @{ revision = 4; n = 16; withFf = $true }
    ) {
        param ($revision, $n, $withFf)
        $expected = newExpectedKey $revision $n $withFf
        $u = newExpectedU $revision $expected
        $key = getPdfKeyR234 $revision $n $script:owner $u (-4.0) $script:id (-not $withFf)
        toHexText $key | Should -Be (toHexText $expected)
    }

    It "/U と合わない（パスワードが要る）と例外" {
        $expected = newExpectedKey 3 16 $false
        $u = newExpectedU 3 $expected
        $u[0] = $u[0] -bxor 1
        { getPdfKeyR234 3 16 $script:owner $u (-4.0) $script:id $true } | Should -Throw "PDFにパスワードが必要なため読めません。"
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
