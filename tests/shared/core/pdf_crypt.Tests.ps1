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

    It "SHA-384・512 の長さ" {
        (getPdfSha 384 (toPdfBytes "abc")).Length | Should -Be 48
        (getPdfSha 512 (toPdfBytes "abc")).Length | Should -Be 64
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

    It "None はそのまま返す" {
        $data = [byte[]](1..5)
        toHexText (unprotectPdfData @{ Key = [byte[]](1..5) } "None" 1 0 $data) | Should -Be (toHexText $data)
    }
}
