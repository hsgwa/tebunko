# 暗号化されたファイルの種類の見分け（shared\office\office_protection_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function bytesOf([string]$hex) {
        return [byte[]](@($hex -split " " | Where-Object { $_ -ne "" } | ForEach-Object { [Convert]::ToByte($_, 16) }))
    }

    ${zipHead}      = bytesOf "50 4B 03 04 00 00"
    ${compoundHead} = bytesOf "D0 CF 11 E0 A1 B1 1A E1 00 00"
}

Describe "getOfficeProtectionKind" -Tag Unit {
    It "ZIP（PK 03 04）は Zip" {
        getOfficeProtectionKind ${zipHead} $null | Should -Be "Zip"
    }

    It "先頭バイト列が空・null は Text" {
        getOfficeProtectionKind ([byte[]]@()) $null | Should -Be "Text"
        getOfficeProtectionKind $null $null | Should -Be "Text"
    }

    # CFB のエントリ名の組み合わせごとの種類
    It "<name>" -TestCases @(
        @{ name = "新形式の権限保護（DRMEncryptedDataSpace・DRMEncryptedTransform）は Rights"
           names = @("Root Entry", ([char]6 + "DataSpaces"), "DRMEncryptedDataSpace", "DRMEncryptedTransform"); kind = "Rights" }
        @{ name = "旧形式の権限保護（\x09DRMContent が必須）は Rights"
           names = @("Root Entry", ([char]9 + "DRMContent"), ([char]9 + "DRMDataSpace"), ([char]9 + "DRMTransform")); kind = "Rights" }
        @{ name = "旧形式の権限保護（任意の \x09DRMViewerContent・LZX 系を含む）も Rights"
           names = @("Root Entry", ([char]9 + "DRMContent"), ([char]9 + "DRMViewerContent"), ([char]9 + "LZXDRMDataSpace"), ([char]9 + "LZXTransform")); kind = "Rights" }
        @{ name = "新形式のパスワード（StrongEncryptionDataSpace・EncryptionInfo・EncryptedPackage）は Password"
           names = @("Root Entry", ([char]6 + "DataSpaces"), "StrongEncryptionDataSpace", "EncryptionInfo", "EncryptedPackage"); kind = "Password" }
        @{ name = "EncryptionInfo だけでも Password"
           names = @("Root Entry", "EncryptionInfo", "EncryptedPackage"); kind = "Password" }
        @{ name = "\x06DataSpaces はあるが、権限保護・パスワードの名前を1つも取りこぼしたときは安全側で Rights"
           names = @("Root Entry", ([char]6 + "DataSpaces"), "DataSpaceMap"); kind = "Rights" }
        @{ name = "権限保護・パスワードの名前が無いCFBは Legacy（旧形式）"
           names = @("Root Entry", "WordDocument", "1Table"); kind = "Legacy" }
        @{ name = "エントリ名が読めない（`$null）ときも Legacy（今と同じ動き）"
           names = $null; kind = "Legacy" }
    ) {
        param ($name, $names, $kind)
        getOfficeProtectionKind ${compoundHead} $names | Should -Be $kind
    }

    # ZIPでもCFBでもない先頭バイト列
    It "<name>" -TestCases @(
        @{ name = "空ファイルは Text"; hex = ""; kind = "Text" }
        @{ name = "UTF-16(BOM) のテキストは Text（NULを含んでも）"; hex = "FF FE 54 00 45 00 53 00 54 00"; kind = "Text" }
        @{ name = "RTF（{\rtf）は Text"; hex = "7B 5C 72 74 66 31 5C 61 6E 73 69"; kind = "Text" }
        @{ name = "UTF-8(BOM) の HTML は Text"; hex = "EF BB BF 3C 68 74 6D 6C 3E 3C 68 65 61 64 3E"; kind = "Text" }
        @{ name = "UTF-8 のふつうの文章は Text"; hex = "E3 81 82 E3 81 84 E3 81 86"; kind = "Text" }
        @{ name = "NUL（0x00）を含むバイナリは Unknown（透過暗号化の製品の暗号文の見込み）"; hex = "12 34 00 56 78 9A"; kind = "Unknown" }
    ) {
        param ($name, $hex, $kind)
        $head = if ($hex -eq "") { [byte[]]@() } else { bytesOf $hex }
        getOfficeProtectionKind $head $null | Should -Be $kind
    }
}

Describe "getProtectionFailureText" -Tag Unit {
    It "<kind>" -TestCases @(
        @{ kind = "Password"; hasText = $true }
        @{ kind = "Rights"; hasText = $true }
        @{ kind = "Unknown"; hasText = $true }
        @{ kind = "Zip"; hasText = $false }
        @{ kind = "Legacy"; hasText = $false }
        @{ kind = "Text"; hasText = $false }
    ) {
        param ($kind, $hasText)
        $text = getProtectionFailureText $kind
        if ($hasText) {
            $text | Should -Not -BeNullOrEmpty
        } else {
            $text | Should -BeNullOrEmpty
        }
    }

    It "種類ごとに文言が違う" {
        (getProtectionFailureText "Password") | Should -Not -Be (getProtectionFailureText "Rights")
        (getProtectionFailureText "Rights") | Should -Not -Be (getProtectionFailureText "Unknown")
    }
}

Describe "testOfficeOutput" -Tag Unit {
    It "Zip を期待し、ZIPの先頭バイト列なら `$true" {
        testOfficeOutput ${zipHead} "Zip" | Should -Be $true
    }

    It "Zip を期待し、ZIPでなければ `$false" {
        testOfficeOutput (bytesOf "12 34 00 56") "Zip" | Should -Be $false
    }

    It "UnicodeText を期待し、BOM（FF FE）なら `$true" {
        testOfficeOutput (bytesOf "FF FE 41 00") "UnicodeText" | Should -Be $true
    }

    It "UnicodeText を期待し、BOMが無ければ `$false" {
        testOfficeOutput (bytesOf "41 00 42 00") "UnicodeText" | Should -Be $false
    }
}

Describe "testWorkbookFormat" -Tag Unit {
    It "<fileFormat>" -TestCases @(
        @{ fileFormat = 51; isWorkbook = $true }    # xlOpenXMLWorkbook
        @{ fileFormat = -4143; isWorkbook = $true }  # xlWorkbookNormal
        @{ fileFormat = 56; isWorkbook = $true }    # xlExcel8（97-2003）
        @{ fileFormat = 6; isWorkbook = $false }    # xlCSV
        @{ fileFormat = 44; isWorkbook = $false }   # xlHtml
        @{ fileFormat = 42; isWorkbook = $false }   # xlUnicodeText
        @{ fileFormat = -4158; isWorkbook = $false } # xlCurrentPlatformText
    ) {
        param ($fileFormat, $isWorkbook)
        testWorkbookFormat $fileFormat | Should -Be $isWorkbook
    }
}

Describe "getWordOpenFormat" -Tag Unit {
    It "<extension>" -TestCases @(
        @{ extension = ".doc"; value = 1 }
        @{ extension = ".DOC"; value = 1 }
        @{ extension = ".docx"; value = 9 }
        @{ extension = ".docm"; value = 10 }
        @{ extension = ".dotx"; value = 11 }
        @{ extension = ".xyz"; value = 0 }
    ) {
        param ($extension, $value)
        getWordOpenFormat $extension | Should -Be $value
    }
}
