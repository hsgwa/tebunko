# 起動の時間の計測（tools\perf\startup_common.ps1）の集計のテスト。起動そのもの（measure_startup.ps1）は実機で確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\tools\perf\startup_common.ps1"

    function newTraceLines {
        param ([int[]]$offsets)
        # 段階 s1, s2, … と、終点 2 つの記録の行（経過ミリ秒は offsets）
        $names = @("PowerShell の起動", "s2", ${startupMarkWindow}, ${startupMarkList})
        return @(for ($i = 0; $i -lt $offsets.Count; $i++) { "$($names[$i])`t$($offsets[$i])" })
    }

    function newRun {
        param ([int[]]$offsets, $external = $null)
        $run = getStartupRunTimes (readStartupTrace (newTraceLines $offsets)).Rows
        $run.ExternalMs = $external
        return $run
    }
}

Describe "readStartupTrace" -Tag Unit {
    It "正常な行を、名前と経過ミリ秒に読む" {
        $read = readStartupTrace @("PowerShell の起動`t900", "lib.ps1`t1500.5")
        $read.Problems.Count | Should -Be 0
        $read.Rows.Count | Should -Be 2
        $read.Rows[1].Name | Should -Be "lib.ps1"
        $read.Rows[1].Ms | Should -Be 1500.5
    }

    It "<name>: 飛ばして、理由を知らせる" -TestCases @(
        @{ name = "タブが無い"; line = "段階だけ" }
        @{ name = "列が多い"; line = "a`t1`t2" }
        @{ name = "数でない"; line = "a`tabc" }
        @{ name = "名前が空"; line = "`t5" }
        @{ name = "負の数"; line = "a`t-3" }
    ) {
        $read = readStartupTrace @("先頭`t10", $line, "末尾`t20")
        $read.Rows.Count | Should -Be 2
        $read.Problems.Count | Should -Be 1
        $read.Problems[0] | Should -Match "^2 行目を飛ばした"
    }

    It "時刻が前の行より戻っている行は飛ばす。空行は問題にしない" {
        $read = readStartupTrace @("a`t100", "", "b`t50", "c`t200")
        @($read.Rows | ForEach-Object { $_.Name }) | Should -Be @("a", "c")
        $read.Problems.Count | Should -Be 1
        $read.Problems[0] | Should -Match "戻っている"
    }

    It "行が無くても止まらない" {
        (readStartupTrace @()).Rows.Count | Should -Be 0
        (readStartupTrace $null).Rows.Count | Should -Be 0
    }
}

Describe "getStartupRunTimes" -Tag Unit {
    It "段階の時間は隣り合う行の差で、最初の段階はプロセスの始まりから。終点の時刻は累計のまま" {
        $run = getStartupRunTimes (readStartupTrace (newTraceLines @(900, 1200, 2500, 3100))).Rows
        @($run.Stages | ForEach-Object { $_.Ms }) | Should -Be @(900, 300, 1300, 600)
        $run.WindowMs | Should -Be 2500
        $run.ListMs | Should -Be 3100
    }

    It "終点の行が無いときは、その値が `$null" {
        $run = getStartupRunTimes (readStartupTrace (newTraceLines @(900, 1200))).Rows
        $run.WindowMs | Should -BeNullOrEmpty
        $run.ListMs | Should -BeNullOrEmpty
        $run.Stages.Count | Should -Be 2
    }
}

Describe "getStartupSummary" -Tag Unit {
    It "段階ごとの最小・中央値・最大と、終点・外から測った値の統計を出す（段階は最初に現れた順）" {
        $summary = getStartupSummary @(
            (newRun @(900, 1200, 2500, 3100) 4000),
            (newRun @(1000, 1400, 2700, 3500) 4400),
            (newRun @(800, 1100, 2300, 2900) 3600)
        )
        $summary.Runs | Should -Be 3
        @($summary.Stages | ForEach-Object { $_.Name }) | Should -Be @("PowerShell の起動", "s2", ${startupMarkWindow}, ${startupMarkList})
        $first = $summary.Stages[0].Stats
        $first.Min | Should -Be 800
        $first.Median | Should -Be 900
        $first.Max | Should -Be 1000
        $summary.Window.Median | Should -Be 2500
        $summary.List.Max | Should -Be 3500
        $summary.External.Median | Should -Be 4000
    }

    It "段階が欠けた回があっても、あった回だけで集計する。回が無ければ統計は `$null" {
        $summary = getStartupSummary @((newRun @(900, 1200, 2500, 3100)), (newRun @(900, 1200)))
        ($summary.Stages | Where-Object { $_.Name -eq ${startupMarkList} }).Stats.Count | Should -Be 1
        $summary.External | Should -BeNullOrEmpty
        (getStartupSummary @()).Window | Should -BeNullOrEmpty
    }
}

Describe "formatStartupTable" -Tag Unit {
    It "後だけのとき: 段階の行・終点の行・外から測った行を、最小 / 中央値 / 最大で並べる" {
        $lines = @(formatStartupTable (getStartupSummary @((newRun @(900, 1200, 2500, 3100) 4000))))
        $lines[0] | Should -Be "| 段階 | 最小 / 中央値 / 最大（ms） |"
        $lines | Should -Contain "| s2 | 300 / 300 / 300 |"
        $lines | Should -Contain "| **(A) 画面が出る（プロセスの始まりから）** | 2500 / 2500 / 2500 |"
        $lines[-1] | Should -Be "| 外から測った、tebunko.bat を起動してから本体の窓が出るまで | 4000 / 4000 / 4000 |"
    }

    It "前と後を渡すと、中央値の差と割合の欄が付く。前に無い段階は - になる" {
        $before = getStartupSummary @((newRun @(1000, 1300, 3000) 5000))
        $after = getStartupSummary @((newRun @(900, 1100, 2400, 3000) 4000))
        $lines = @(formatStartupTable $after $before)
        $lines[0] | Should -Match "中央値の差 \|$"
        $lines | Should -Contain "| **(A) 画面が出る（プロセスの始まりから）** | 3000 / 3000 / 3000 | 2400 / 2400 / 2400 | -600 (-20%) |"
        $lines | Should -Contain "| 外から測った、tebunko.bat を起動してから本体の窓が出るまで | 5000 / 5000 / 5000 | 4000 / 4000 / 4000 | -1000 (-20%) |"
        $lines | Should -Contain "| **(B) 一覧が出る（プロセスの始まりから）** | - | 3000 / 3000 / 3000 | - |"
    }

    It "増えたときは + を付ける" {
        $before = getStartupSummary @((newRun @(1000, 2000, 3000) 5000))
        $after = getStartupSummary @((newRun @(1000, 2000, 3300) 5000))
        (@(formatStartupTable $after $before) -join "`n") | Should -Match "\+300 \(\+10%\)"
    }
}
