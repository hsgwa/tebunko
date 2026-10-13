# 元のフォルダを記録できる名前の判断（tebunko\core\source_name.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "testSourceNameRecordable" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "ふつうの名前は記録できる"; name = "営業"; expected = $true }
        @{ label = "内側の空白は記録できる"; name = "営業 部"; expected = $true }
        @{ label = "幅ゼロの文字は削られないので記録できる"; name = ("営" + [char]0x200B + "業"); expected = $true }
        @{ label = "ソフトハイフン付きは削られないので記録できる"; name = ("営業" + [char]0x00AD); expected = $true }
        @{ label = "空は記録できない"; name = ""; expected = $false }
        @{ label = "空白だけは記録できない"; name = "  "; expected = $false }
        @{ label = "前に全角空白"; name = ([string][char]0x3000 + "営業"); expected = $false }
        @{ label = "後ろに全角空白"; name = ("営業" + [char]0x3000); expected = $false }
        @{ label = "後ろに NBSP"; name = ("営業" + [char]0x00A0); expected = $false }
        @{ label = "後ろにタブ"; name = ("営業" + [char]9); expected = $false }
        @{ label = "後ろに半角空白"; name = "営業 "; expected = $false }
    ) {
        param ($name, $expected)
        testSourceNameRecordable $name | Should -Be $expected
    }
}
