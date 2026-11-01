# 処理の記録の集計（tools\perf\read_trace.ps1）と、起動から最初の検索までの計測（tools\perf\measure_startup.ps1）の集計のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\tools\perf\measure_startup.ps1"

    # 記録の TSV を 1 つ書く（時刻・経過ミリ秒・区分・名前・スレッド）
    function writeTraceFile {
        param ([string]$Dir, [string]$FileName, [object[]]$Rows)

        [void][System.IO.Directory]::CreateDirectory($Dir)
        $lines = @("時刻`t経過ミリ秒`t区分`t名前`tスレッド")
        foreach ($row in $Rows) {
            $cells = $row -split "[|]"
            $lines += "00:00:00.0000000`t$($cells[0])`t$($cells[1])`t$($cells[2])`t1"
        }
        [System.IO.File]::WriteAllText((Join-Path $Dir $FileName), (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    }

    $script:startupKeys = @(
        "01 PowerShell の起動", "stepSplash 5", "stepSplash 50", "04 画面の組み立て（Loaded）", "05 窓が出る（ContentRendered）",
        "06 起動時の読み込み（画面のスレッドの分）", "07 インデックスのツリーが画面に入る", "08 検索の司令の用意が済む", "09 検索の送り",
        "10 表に最初の行が出る", "11 総数が決まる"
    )
    # 全部の区切りがある 1 回分の行。経過は 100・200・…・1100（ミリ秒）に Offset を足す
    function newStartupRows {
        param ([double]$Offset = 0)
        $rows = @()
        for ($i = 0; $i -lt $script:startupKeys.Count; $i++) {
            $rows += "$($Offset + 100 * ($i + 1))|startup|$($script:startupKeys[$i])"
        }
        return $rows
    }
}

Describe "readTraceRows" -Tag Unit {
    It "指定の区分の行だけを読み、ほかの区分・見出し・壊れた行は数えない" {
        $dir = Join-Path $TestDrive "read1"
        writeTraceFile $dir "20260101-000000_100_1.tsv" @("10|startup|a", "20|search|b", "30|startup|c")
        [System.IO.File]::AppendAllText((Join-Path $dir "20260101-000000_100_1.tsv"), "壊れた行`r`n")

        $rows = @(readTraceRows $dir "startup")

        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be "a"
        $rows[1].ElapsedMs | Should -Be 30
        $rows[0].Process | Should -Be 100
    }

    It "-ProcessId を渡すと、そのプロセスのファイルだけを読む。フォルダが無ければ空" {
        $dir = Join-Path $TestDrive "read2"
        writeTraceFile $dir "20260101-000000_100_1.tsv" @("10|startup|a")
        writeTraceFile $dir "20260101-000000_200_1.tsv" @("10|startup|b")

        @(readTraceRows $dir "startup" 200).Count | Should -Be 1
        @(readTraceRows $dir "startup" 200)[0].Name | Should -Be "b"
        @(readTraceRows (Join-Path $TestDrive "nothing") "startup").Count | Should -Be 0
    }
}

Describe "getTraceTable（区間ごとの時間と統計）" -Tag Unit {
    It "区間の長さは終点と始まりの経過の差で、最小・中央値・最大を出す" {
        $defs = @(
            @{ Key = "a"; Label = "A"; Previous = $null; Kind = "mark" }
            @{ Key = "b"; Label = "B"; Previous = "a"; Kind = "mark" }
        )
        $runs = @(
            @{ Markers = @{ a = 100.0; b = 150.0 }; Failed = $false }
            @{ Markers = @{ a = 200.0; b = 300.0 }; Failed = $false }
            @{ Markers = @{ a = 300.0; b = 450.0 }; Failed = $false }
        )

        $table = getTraceTable $runs $defs
        $rows = @($table -split "`r`n" | Where-Object { $_ -like "| [0-9]*" })

        $rows.Count | Should -Be 2
        # A: 長さ中央値 200・最大 300、経過 中央値 200・最大 300 / B: 長さ 50・100・150 → 中央値 100・最大 150、経過 150・300・450 → 中央値 300・最大 450
        $rows[0] | Should -Be "| 1 | A | 200 | 300 | 200 | 300 | 3 |"
        $rows[1] | Should -Be "| 2 | B | 100 | 150 | 300 | 450 | 3 |"
        $table | Should -Match "回数 3 回（失敗 0 回は除く）"
    }

    It "失敗の回は数えず、件数だけ添える。欠けた区切りは「–」にする" {
        $defs = @(
            @{ Key = "a"; Label = "A"; Previous = $null; Kind = "mark" }
            @{ Key = "b"; Label = "B"; Previous = "a"; Kind = "mark" }
            @{ Key = "z"; Label = "Z"; Previous = "b"; Kind = "mark" }
        )
        $runs = @(
            @{ Markers = @{ a = 100.0; b = 200.0 }; Failed = $false }
            @{ Markers = @{ a = 9999.0; b = 99999.0 }; Failed = $true }
        )

        $table = getTraceTable $runs $defs
        $rows = @($table -split "`r`n" | Where-Object { $_ -like "| [0-9]*" })

        $rows[1] | Should -Be "| 2 | B | 100 | 100 | 200 | 200 | 1 |"
        $rows[2] | Should -Be "| 3 | Z | – | – | – | – | 0 |"
        $table | Should -Match "回数 1 回（失敗 1 回は除く）"
    }

    It "Kind が duration の区切りは、値をそのまま長さにする" {
        $defs = @(@{ Key = "block"; Label = "止まり"; Previous = $null; Kind = "duration" })
        $runs = @(@{ Markers = @{ block = 40.0 }; Failed = $false }, @{ Markers = @{ block = 60.0 }; Failed = $false })

        $rows = @((getTraceTable $runs $defs) -split "`r`n" | Where-Object { $_ -like "| [0-9]*" })

        $rows[0] | Should -Be "| 1 | 止まり | 50 | 60 | – | – | 2 |"
    }

    It "定義を渡さないときは、記録に出てきた名前を経過の中央値の順に並べる" {
        $runs = @(
            @{ Markers = @{ later = 300.0; first = 100.0 }; Failed = $false }
            @{ Markers = @{ later = 500.0; first = 200.0 }; Failed = $false }
        )

        $rows = @((getTraceTable $runs) -split "`r`n" | Where-Object { $_ -like "| [0-9]*" })

        $rows[0] | Should -Match '^\| 1 \| first \|'
        $rows[1] | Should -Match '^\| 2 \| later \| 250 \|'
    }
}

Describe "getTraceRuns・getStartupRunFromRows（記録の行から実行ごとにまとめる）" -Tag Unit {
    It "ほかの区分の行を混ぜても、指定の区分だけで 1 回分を作り、同じ名前は最初の 1 回を取る" {
        $dir = Join-Path $TestDrive "runs1"
        $rows = @(newStartupRows)
        $rows += "5|search|混ざる"
        $rows += "9999|startup|11 総数が決まる"
        writeTraceFile $dir "20260101-000000_100_1.tsv" $rows

        $run = getStartupRunFromRows @(readTraceRows $dir "startup") 1500 33

        $run.Failed | Should -Be $false
        $run.Markers["11 総数が決まる"] | Should -Be 1100
        $run.Markers.ContainsKey("混ざる") | Should -Be $false
        $run.Markers["窓の最長の止まり"] | Should -Be 33
        $run.OutsideMs | Should -Be 1500
    }

    It "総数の区切りが無い（途中で終わった）回は、失敗にする" {
        $dir = Join-Path $TestDrive "runs2"
        writeTraceFile $dir "20260101-000000_100_1.tsv" @("100|startup|01 PowerShell の起動")

        (getStartupRunFromRows @(readTraceRows $dir "startup") 0 0).Failed | Should -Be $true
        (getStartupRunFromRows @() 0 0).Failed | Should -Be $true
    }
}

Describe "getStartupSummaryTable（区切り 1〜12 の表）" -Tag Unit {
    It "12 行の表になり、外から測った値と記録の区切り 11 の差を添える" {
        $dir = Join-Path $TestDrive "summary"
        $runs = @()
        foreach ($n in 1..3) {
            $file = "20260101-00000$n" + "_$($n)00_1.tsv"
            writeTraceFile $dir $file (newStartupRows ($n * 10))
        }
        foreach ($n in 1..3) {
            $runs += getStartupRunFromRows @(readTraceRows $dir "startup" ($n * 100)) 1200 20
        }

        $text = getStartupSummaryTable $runs
        $rows = @($text -split "`r`n" | Where-Object { $_ -like "| [0-9]*" })

        $rows.Count | Should -Be 12
        $rows[0] | Should -Match '^\| 1 \| PowerShell の起動'
        $rows[10] | Should -Match '^\| 11 \| 総数が決まる \| 100 \| 100 \| 1120 \|'
        $rows[11] | Should -Be "| 12 | 送りまでに、窓が入力を受け付けなかった最長の時間（外から測る） | 20 | 20 | – | – | 3 |"
        $text | Should -Match '外から測った 起動 → 記録が書かれる: 中央値 1200 ms'
        $text | Should -Match '記録の区切り 11 の経過: 中央値 1120 ms（差 6.7 %）'
    }
}

Describe "getStartupSearchTable（最初の検索の内訳）" -Tag Unit {
    It "search の行から、画面と司令の区切りを並べた表を作り、無い区切り（0 件の最初の行）は「–」にする" {
        $dir = Join-Path $TestDrive "searchtable"
        $startup = @(newStartupRows)
        $search = @("1000|search|画面: 送り", "1010|search|司令: 受け取り", "1100|search|司令: gram 索引で絞り終え", "1150|search|司令: ファイルの列挙の終わり",
            "1400|search|司令: searchContentIndex の終わり", "1450|search|司令: 終わり (done)", "1500|search|画面: 総数")
        writeTraceFile $dir "20260101-000000_100_1.tsv" ($startup + $search)
        $runs = @(getStartupRunFromRows @(readTraceRows $dir "startup" 100) 1700 10 @(readTraceRows $dir "search" 100))

        $runs[0].SearchMarkers["司令: 受け取り"] | Should -Be 1010
        $rows = @((getStartupSearchTable $runs) -split "`r`n" | Where-Object { $_ -like "| [0-9]*" })

        $rows.Count | Should -Be 7
        $rows[1] | Should -Be "| 2 | 司令: 受け取り | 10 | 10 | 1010 | 1010 | 1 |"
        $rows[4] | Should -Be "| 5 | 司令: 読み込みと照合（searchContentIndex）の終わり | 250 | 250 | 1400 | 1400 | 1 |"
        $rows[5] | Should -Be "| 6 | 画面: 表に最初の行（送りから） | – | – | – | – | 0 |"
        $rows[6] | Should -Be "| 7 | 画面: 総数が決まる（司令の終わりから） | 100 | 100 | 1500 | 1500 | 1 |"
    }

    It "search の行が無い回は、内訳の表では失敗の回として数える" {
        $runs = @(getStartupRunFromRows @() 0 0)

        $runs[0].SearchMarkers.Count | Should -Be 0
        (getStartupSearchTable $runs) | Should -Match "回数 0 回（失敗 1 回は除く）"
    }
}
