# 画面のスモークテストの共通の関数（tests\gui\gui_helpers.ps1）のうち、画面を起動せずに確かめられるもの。
# gui_helpers.ps1 は、読み込むだけでは画面も UI オートメーションも動かさない（アセンブリを読み込むだけ）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "getGuiCrashPatterns" -Tag Unit {
    It "プロセス ID を 10 進・16 進（0x 付き・無し）の 3 通りで返す" -TestCases @(
        @{ ProcessId = 4660;  Expected = @("4660", "0x1234", "1234") }
        @{ ProcessId = 255;   Expected = @("255", "0xff", "ff") }
        @{ ProcessId = 43981; Expected = @("43981", "0xabcd", "abcd") }
    ) {
        $patterns = @(getGuiCrashPatterns $ProcessId)
        $patterns.Count | Should -Be 3
        foreach ($expectedPattern in $Expected) {
            $patterns | Should -Contain $expectedPattern
        }
    }

    It "16 進の記録（大文字）も、-like で見つけられる" {
        $message = "Faulting process id: 0xABCD"
        @(getGuiCrashPatterns 43981 | Where-Object { $message -like "*$_*" }).Count | Should -BeGreaterThan 0
    }
}
