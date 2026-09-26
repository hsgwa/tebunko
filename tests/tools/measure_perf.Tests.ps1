# 性能の計測（tools\measure_perf.ps1・tools\perf\）のテスト
$measure = "$PSScriptRoot\..\..\tools\measure_perf.ps1"
. "$PSScriptRoot\..\..\tools\perf\perf_common.ps1"

function writeTsv([string]$path, [string[]]$lines) {
    [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

Describe "perf_common.ps1 の統計値" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "件数が奇数のときの最小・中央値・平均・最大"; values = @(5, 1, 3); min = 1; median = 3; mean = 3; max = 5; count = 3 }
        @{ name = "件数が偶数のときは、中央の 2 つの平均を中央値にする"; values = @(4, 1, 2, 3); min = 1; median = 2.5; mean = 2.5; max = 4; count = 4 }
        @{ name = "1 件だけのときは、どれも同じ値"; values = @(7); min = 7; median = 7; mean = 7; max = 7; count = 1 }
    ) {
        param ($values, $min, $median, $mean, $max, $count)
        $s = getStats ([double[]]$values)
        $s.Min | Should Be $min
        $s.Median | Should Be $median
        $s.Mean | Should Be $mean
        $s.Max | Should Be $max
        $s.Count | Should Be $count
    }

    It '値が無ければ $null を返す' {
        getStats ([double[]]@()) | Should BeNullOrEmpty
    }

    It "<name>" -TestCases @(
        @{ name = "最小二乗の傾き: 一直線に増える"; x = @(1, 2, 3, 4); y = @(10, 12, 14, 16); slope = 2 }
        @{ name = "最小二乗の傾き: 変わらない"; x = @(1, 2, 3); y = @(5, 5, 5); slope = 0 }
        @{ name = "最小二乗の傾き: ばらつきがある"; x = @(1, 2, 3); y = @(1, 4, 3); slope = 1 }
    ) {
        param ($x, $y, $slope)
        getSlope ([double[]]$x) ([double[]]$y) | Should Be $slope
    }

    It "<name>" -TestCases @(
        @{ name = '最小二乗の傾き: 2 点に満たないときは $null'; x = @(1); y = @(1) }
        @{ name = '最小二乗の傾き: x が同じ値だけのときは $null'; x = @(2, 2); y = @(1, 3) }
    ) {
        param ($x, $y)
        getSlope ([double[]]$x) ([double[]]$y) | Should BeNullOrEmpty
    }

    It "<name>" -TestCases @(
        @{ name = "点が多いときは間引き、最初と最後の点は残す"; count = 250; limit = 100; first = 1; last = 250; most = 101 }
        @{ name = "点が少ないときは間引かない"; count = 5; limit = 100; first = 1; last = 5; most = 5 }
    ) {
        param ($count, $limit, $first, $last, $most)
        $thin = thinOut @(1..$count) $limit
        $thin.Count | Should Not BeGreaterThan $most
        $thin[0] | Should Be $first
        $thin[$thin.Count - 1] | Should Be $last
    }

    It "折れ線グラフは、数をカルチャによらない書き方にし、色を指定できる" {
        $lines = newLineChart "題" "回" "[1, 2]" "ms" @(, @(1.5, 2000)) @("#e41a1c")
        ($lines -join "`n") | Should Match "line \[1\.5, 2000\]"
        ($lines -join "`n") | Should Match "plotColorPalette"
        $lines[0] | Should Be '```mermaid'
    }
}

Describe "measure_perf.ps1" -Tag Io {
    # 取り込みの一時置き場の形（<フォルダ>\<ファイル名.xlsx>\<場所>.tsv）の小さなインデックスと、検索する語の表
    $index = Join-Path $TestDrive "work\index"
    writeTsv "$index\営業\見積.xlsx\シート1.tsv" @("1`t見積書`t山田", "2`tりんご`t佐藤")
    writeTsv "$index\営業\見積.xlsx\シート2.tsv" @("3`tみかん")
    writeTsv "$index\総務\名簿.xlsx\シート1.tsv" @("10`t佐藤`t総務", "11`t鈴木`t総務")
    $words = Join-Path $TestDrive "words.tsv"
    writeTsv $words @("名前`t語`t正規表現`t件数", "無い`t該当無し`tfalse`t0", "まれ`t見積書`tfalse`t1", "数字`t^\d{2}\t`ttrue`t2")
    $out = Join-Path $TestDrive "work\result"
    & $measure -Index $index -Work (Join-Path $TestDrive "work") -Words $words -Count 3 -SampleMs 50 -Label "テスト" -RunId "1" -Ref "main" -Sha "abc1234" -Scale "0.1" -DataSeconds 1.5 6>$null | Out-Null
    $result = [System.IO.File]::ReadAllText("$out\result.json") | ConvertFrom-Json

    It "形式の版と実行の情報を書く" {
        $result.Schema | Should Be 1
        $result.Run.RunId | Should Be "1"
        $result.Run.Ref | Should Be "main"
        $result.Run.Count | Should Be 3
        $result.Run.DataSeconds | Should Be 1.5
        # main のコードは検索の司令のスレッド（SearchService）を持つので、画面と同じ流れで測る
        $result.Run.SearchMode | Should Be "service"
    }

    It "インデックス作成で pack を作り、フォルダ・ブック・TSV の数を数える" {
        $result.Index.Pack.Folders | Should Be 2
        $result.Index.Pack.Books | Should Be 2
        $result.Index.Pack.Tsv | Should Be 3
        $result.Index.Pack.Packs | Should Be 2
        @($result.Index.Resources | ForEach-Object { $_.Phase }) -contains "pack の作成" | Should Be $true
        # pack にまとめたので、ブックごとの TSV は残らない
        @([System.IO.Directory]::GetFiles($index, "シート*.tsv", "AllDirectories")).Count | Should Be 0
    }

    It "語ごとにヒット件数と、検索時間の最小・中央値・平均・最大を返す" {
        $byName = @{}
        foreach ($s in $result.Search) { $byName[$s.Name] = $s }
        $byName["無い"].Hits | Should Be 0
        $byName["まれ"].Hits | Should Be 1
        $byName["数字"].Hits | Should Be 2
        foreach ($s in $result.Search) {
            $s.TotalMs.Count | Should Be 3
            $s.TotalMs.Min | Should Not BeGreaterThan $s.TotalMs.Median
            $s.TotalMs.Median | Should Not BeGreaterThan $s.TotalMs.Max
            $s.TotalMs.Mean | Should Not BeGreaterThan $s.TotalMs.Max
            $s.TotalMs.Min | Should Not BeGreaterThan $s.TotalMs.Mean
            $s.WorkingSetPer10MB | Should Not BeNullOrEmpty
            # service の流れでは、列挙と照合に分け、lib.ps1 の読み込みは検索ごとに分けない
            $s.ListMs.Count | Should Be 3
            $s.MatchMs.Count | Should Be 3
            $s.LoadMs | Should BeNullOrEmpty
        }
    }

    It "検索 1 回ごとの記録と、1 行 1 指標の metrics.csv を書く" {
        @(Import-Csv "$out\searches.csv").Count | Should Be 9
        $metrics = @(Import-Csv "$out\metrics.csv")
        @($metrics | Where-Object { $_.metric -eq "search_total_ms" -and $_.word -eq "まれ" -and $_.stat -eq "median" }).Count | Should Be 1
        @($metrics | Where-Object { $_.metric -eq "index_seconds" }).Count | Should Be 1
        @($metrics | Where-Object { $_.run_id -ne "1" -or $_.sha -ne "abc1234" }).Count | Should Be 0
        (Get-Content "$out\metrics.csv" -TotalCount 1) | Should Match '"run_id","date","ref","sha","scale","metric","word","stat","value","unit"$'
    }

    It "summary.md に表とグラフを書き、パスを書かない" {
        $md = [System.IO.File]::ReadAllText("$out\summary.md")
        $md | Should Match "# tebunko の性能（テスト）"
        $md | Should Match "## インデックス作成"
        $md | Should Match "## 検索"
        $md | Should Match "検索 1 回ごとの時間"
        $md | Should Match "検索 1 回ごとのワーキングセット"
        $md.Contains($TestDrive) | Should Be $false
    }

    It "pack だけのインデックスは、作成を測らずに検索だけを測る" {
        $again = Join-Path $TestDrive "again"
        & $measure -Index $index -Work $again -Words $words -Count 1 -SampleMs 50 6>$null | Out-Null
        $second = [System.IO.File]::ReadAllText("$again\result\result.json") | ConvertFrom-Json
        $second.Index.Pack | Should BeNullOrEmpty
        @($second.Search).Count | Should Be 3
    }
}
