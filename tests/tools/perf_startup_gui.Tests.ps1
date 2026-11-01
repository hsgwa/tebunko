# 起動から最初の検索までの計測（tools\perf\measure_startup.ps1）を、本物の画面で 1 回流す確かめ（集計は perf_startup.Tests.ps1）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\..\gui\gui_helpers.ps1"
    . "$PSScriptRoot\..\..\tools\perf\measure_startup.ps1"
}

Describe "invokeStartupMeasure（本物の画面を起動して測る）" -Tag Gui {
    BeforeAll {
        $script:tool = newGuiTool $TestDrive
        newGuiSampleIndex $script:tool $TestDrive
    }

    It "記録つきで 1 回起動し、検索まで進んで、区切り 1〜11 の経過と窓の止まりが出る" {
        $run = invokeStartupMeasure $script:tool.Dir "単価" (Join-Path $script:tool.Work "trace") $script:tool.DefaultWorkspace 120

        $run.Failed | Should -BeFalse
        foreach ($def in @(getStartupTraceDefs | Where-Object { $_.Kind -eq "mark" -and $_.Key -notlike "10 *" })) {
            $run.Markers.ContainsKey($def.Key) | Should -BeTrue -Because $def.Key
        }
        $run.OutsideMs | Should -BeGreaterThan 0
        $run.SearchMarkers.ContainsKey("司令: 受け取り") | Should -BeTrue
    }
}
