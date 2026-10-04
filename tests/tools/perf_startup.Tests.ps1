# 起動と終了の速さの比べ方（tools\perf\startup_common.ps1）の判定。プロセスは起こさない
BeforeAll {
    . "$PSScriptRoot\..\..\tools\perf\startup_common.ps1"
}

Describe "getStartupMedian" {
    It "先頭の warmup 回を捨てた中央値を返す" -TestCases @(
        @{ Values = @(9000, 100, 300, 200); Warmup = 1; Expected = 200 }
        @{ Values = @(9000, 100, 200, 300, 400); Warmup = 1; Expected = 250 }
        @{ Values = @(100, 200, 300); Warmup = 0; Expected = 200 }
    ) {
        getStartupMedian ([double[]]$Values) $Warmup | Should -Be $Expected
    }

    It "捨てたあとに値が無ければ `$null" -TestCases @(
        @{ Values = @(); Warmup = 1 }
        @{ Values = @(100); Warmup = 1 }
    ) {
        getStartupMedian ([double[]]$Values) $Warmup | Should -BeNullOrEmpty
    }
}

Describe "compareStartup" {
    It "単一 / zip の中央値で倍率を出し、1.2 倍以内なら合格にする" -TestCases @(
        @{ Zip = @(5000, 1000, 1000, 1000); Single = @(9000, 1200, 1200, 1200); Ratio = 1.2; Pass = $true }
        @{ Zip = @(5000, 1000, 1000, 1000); Single = @(9000, 1300, 1300, 1300); Ratio = 1.3; Pass = $false }
        @{ Zip = @(5000, 1000, 2000, 3000); Single = @(9000, 500, 1500, 2500); Ratio = 0.75; Pass = $true }
    ) {
        $c = compareStartup ([double[]]$Zip) ([double[]]$Single)
        [Math]::Round($c.Ratio, 2) | Should -Be $Ratio
        $c.Pass | Should -Be $Pass
        $c.Limit | Should -Be 1.2
    }

    It "ウォームアップの値は比べない" {
        $c = compareStartup ([double[]]@(99999, 1000, 1000)) ([double[]]@(1, 1000, 1000))
        $c.Ratio | Should -Be 1
        $c.Pass | Should -BeTrue
    }

    It "中央値が出せない・zip 版が 0 のときは、倍率も合否も出さない（合格とは言わない）" -TestCases @(
        @{ Zip = @(1000); Single = @(1000, 1000) }
        @{ Zip = @(1000, 1000); Single = @() }
        @{ Zip = @(1000, 0, 0); Single = @(1000, 100, 100) }
    ) {
        $c = compareStartup ([double[]]$Zip) ([double[]]$Single)
        $c.Ratio | Should -BeNullOrEmpty
        $c.Pass | Should -BeNullOrEmpty
    }

    It "limit を変えられる" {
        (compareStartup ([double[]]@(0, 1000)) ([double[]]@(0, 1500)) 1 2.0).Pass | Should -BeTrue
    }
}

Describe "countNonZeroExit" {
    It "0 でない・終わらなかった（`$null）回を数える" -TestCases @(
        @{ Codes = @(0, 0, 0); Expected = 0 }
        @{ Codes = @(0, 5, 0, 5); Expected = 2 }
        @{ Codes = @(0, $null, 1); Expected = 2 }
        @{ Codes = @(); Expected = 0 }
    ) {
        countNonZeroExit $Codes | Should -Be $Expected
    }
}
