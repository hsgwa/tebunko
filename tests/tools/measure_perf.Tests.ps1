# 性能の計測（tools\measure_perf.ps1・tools\perf\）のテスト
BeforeAll {
    $measure = "$PSScriptRoot\..\..\tools\measure_perf.ps1"
    . "$PSScriptRoot\..\..\tools\perf\perf_common.ps1"

    function writeTsv([string]$path, [string[]]$lines) {
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
        [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    }
}

Describe "perf_common.ps1 の統計値" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "件数が奇数のときの最小・中央値・平均・最大"; values = @(5, 1, 3); min = 1; median = 3; mean = 3; max = 5; count = 3 }
        @{ name = "件数が偶数のときは、中央の 2 つの平均を中央値にする"; values = @(4, 1, 2, 3); min = 1; median = 2.5; mean = 2.5; max = 4; count = 4 }
        @{ name = "1 件だけのときは、どれも同じ値"; values = @(7); min = 7; median = 7; mean = 7; max = 7; count = 1 }
    ) {
        param ($values, $min, $median, $mean, $max, $count)
        $s = getStats ([double[]]$values)
        $s.Min | Should -Be $min
        $s.Median | Should -Be $median
        $s.Mean | Should -Be $mean
        $s.Max | Should -Be $max
        $s.Count | Should -Be $count
    }

    It '値が無ければ $null を返す' {
        getStats ([double[]]@()) | Should -BeNullOrEmpty
    }

    It "<name>" -TestCases @(
        @{ name = "最小二乗の傾き: 一直線に増える"; x = @(1, 2, 3, 4); y = @(10, 12, 14, 16); slope = 2 }
        @{ name = "最小二乗の傾き: 変わらない"; x = @(1, 2, 3); y = @(5, 5, 5); slope = 0 }
        @{ name = "最小二乗の傾き: ばらつきがある"; x = @(1, 2, 3); y = @(1, 4, 3); slope = 1 }
    ) {
        param ($x, $y, $slope)
        getSlope ([double[]]$x) ([double[]]$y) | Should -Be $slope
    }

    It "<name>" -TestCases @(
        @{ name = '最小二乗の傾き: 2 点に満たないときは $null'; x = @(1); y = @(1) }
        @{ name = '最小二乗の傾き: x が同じ値だけのときは $null'; x = @(2, 2); y = @(1, 3) }
    ) {
        param ($x, $y)
        getSlope ([double[]]$x) ([double[]]$y) | Should -BeNullOrEmpty
    }

    It "<name>" -TestCases @(
        @{ name = "点が多いときは間引き、最初と最後の点は残す"; count = 250; limit = 100; first = 1; last = 250; most = 101 }
        @{ name = "点が少ないときは間引かない"; count = 5; limit = 100; first = 1; last = 5; most = 5 }
    ) {
        param ($count, $limit, $first, $last, $most)
        $thin = thinOut @(1..$count) $limit
        $thin.Count | Should -Not -BeGreaterThan $most
        $thin[0] | Should -Be $first
        $thin[$thin.Count - 1] | Should -Be $last
    }

    It "折れ線グラフは、数をカルチャによらない書き方にし、色を指定できる" {
        $lines = newLineChart "題" "回" "[1, 2]" "ms" @(, @(1.5, 2000)) @("#e41a1c")
        ($lines -join "`n") | Should -Match "line \[1\.5, 2000\]"
        ($lines -join "`n") | Should -Match "plotColorPalette"
        $lines[0] | Should -Be '```mermaid'
    }
}

Describe "measure_perf.ps1" -Tag Io {
    # 取り込みの一時置き場の形（<フォルダ>\<ファイル名.xlsx>\<場所>.tsv）の小さなインデックスと、検索する語の表
    BeforeAll {
        $index = Join-Path $TestDrive "work\index"
        writeTsv "$index\営業\見積.xlsx\シート1.tsv" @("1`t見積書`t山田", "2`tりんご`t佐藤")
        writeTsv "$index\営業\見積.xlsx\シート2.tsv" @("3`tみかん")
        writeTsv "$index\総務\名簿.xlsx\シート1.tsv" @("10`t佐藤`t総務", "11`t鈴木`t総務")
        $words = Join-Path $TestDrive "words.tsv"
        writeTsv $words @("名前`t語`t正規表現`t件数", "無い`t該当無し`tfalse`t0", "まれ`t見積書`tfalse`t1", "数字`t^\d{2}\t`ttrue`t2")
        $out = Join-Path $TestDrive "work\result"
        & $measure -Index $index -Work (Join-Path $TestDrive "work") -Words $words -Count 3 -SampleMs 50 -Label "テスト" -RunId "1" -Ref "main" -Sha "abc1234" -Scale "0.1" -DataSeconds 1.5 6>$null | Out-Null
        $result = [System.IO.File]::ReadAllText("$out\result.json") | ConvertFrom-Json
    }

    It "形式の版と実行の情報を書く" {
        $result.Schema | Should -Be 1
        $result.Run.RunId | Should -Be "1"
        $result.Run.Ref | Should -Be "main"
        $result.Run.Count | Should -Be 3
        $result.Run.DataSeconds | Should -Be 1.5
        # main のコードは検索の司令のスレッド（SearchService）を持つので、画面と同じ流れで測る
        $result.Run.SearchMode | Should -Be "service"
    }

    It "インデックス作成で pack を作り、フォルダ・ブック・TSV の数を数える" {
        $result.Index.Pack.Folders | Should -Be 2
        $result.Index.Pack.Books | Should -Be 2
        $result.Index.Pack.Tsv | Should -Be 3
        $result.Index.Pack.Packs | Should -Be 2
        @($result.Index.Resources | ForEach-Object { $_.Phase }) -contains "pack の作成" | Should -Be $true
        # pack にまとめたので、ブックごとの TSV は残らない
        @([System.IO.Directory]::GetFiles($index, "シート*.tsv", "AllDirectories")).Count | Should -Be 0
    }

    It "語ごとにヒット件数と、検索時間の最小・中央値・平均・最大を返す" {
        $byName = @{}
        foreach ($s in $result.Search) { $byName[$s.Name] = $s }
        $byName["無い"].Hits | Should -Be 0
        $byName["まれ"].Hits | Should -Be 1
        $byName["数字"].Hits | Should -Be 2
        foreach ($s in $result.Search) {
            $s.TotalMs.Count | Should -Be 3
            $s.TotalMs.Min | Should -Not -BeGreaterThan $s.TotalMs.Median
            $s.TotalMs.Median | Should -Not -BeGreaterThan $s.TotalMs.Max
            $s.TotalMs.Mean | Should -Not -BeGreaterThan $s.TotalMs.Max
            $s.TotalMs.Min | Should -Not -BeGreaterThan $s.TotalMs.Mean
            $s.WorkingSetPer10MB | Should -Not -BeNullOrEmpty
            # service の流れでは、列挙と照合に分け、lib.ps1 の読み込みは検索ごとに分けない
            $s.ListMs.Count | Should -Be 3
            $s.MatchMs.Count | Should -Be 3
            $s.LoadMs | Should -BeNullOrEmpty
        }
    }

    It "検索 1 回ごとの記録と、1 行 1 指標の metrics.csv を書く" {
        @(Import-Csv "$out\searches.csv").Count | Should -Be 9
        $metrics = @(Import-Csv "$out\metrics.csv")
        @($metrics | Where-Object { $_.metric -eq "search_total_ms" -and $_.word -eq "まれ" -and $_.stat -eq "median" }).Count | Should -Be 1
        @($metrics | Where-Object { $_.metric -eq "index_seconds" }).Count | Should -Be 1
        @($metrics | Where-Object { $_.run_id -ne "1" -or $_.sha -ne "abc1234" }).Count | Should -Be 0
        (Get-Content "$out\metrics.csv" -TotalCount 1) | Should -Match '"run_id","date","ref","sha","scale","metric","word","stat","value","unit"$'
    }

    It "summary.md に表とグラフを書き、パスを書かない" {
        $md = [System.IO.File]::ReadAllText("$out\summary.md")
        $md | Should -Match "# tebunko の性能（テスト）"
        $md | Should -Match "## インデックス作成"
        $md | Should -Match "## 検索"
        $md | Should -Match "検索 1 回ごとの時間"
        $md | Should -Match "検索 1 回ごとのワーキングセット"
        $md.Contains($TestDrive) | Should -Be $false
    }

    It "pack だけのインデックスは、作成を測らずに検索だけを測る" {
        $again = Join-Path $TestDrive "again"
        & $measure -Index $index -Work $again -Words $words -Count 1 -SampleMs 50 6>$null | Out-Null
        $second = [System.IO.File]::ReadAllText("$again\result\result.json") | ConvertFrom-Json
        $second.Index.Pack | Should -BeNullOrEmpty
        @($second.Search).Count | Should -Be 3
        # -Index だけのときは、取り込みは測らず Ingest を null にする（キーは残す）
        @($second.PSObject.Properties | ForEach-Object { $_.Name }) | Should -Contain "Ingest"
        $second.Ingest | Should -BeNullOrEmpty
    }
}

Describe "取り込みの計測の部品（ingest_common.ps1）" -Tag Unit {
    BeforeAll {
        . "$PSScriptRoot\..\..\tools\perf\ingest_common.ps1"
        $header = "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー`t抽出版"
        function newStatusTsv([string[]]$rows) {
            $path = Join-Path $TestDrive ("status-" + [guid]::NewGuid().ToString("N") + ".tsv")
            writeTsv $path (@("クロール対象フォルダ`tC:\共有\営業部`t計測", $header) + $rows)
            return $path
        }
    }

    It "取り込み一覧は、相対パスごとに最後の行の状態で数える（フォルダの行と列の足りない行は外す）" {
        $path = newStatusTsv @(
            "計測\a.docx`t2026/01/01`t1`t未取り込み`t0`t`t`t2"
            "計測\a.docx`t2026/01/01`t1`t済`t3`t2026/01/02`t`t2"
            "計測\b.docx`t2026/01/01`t1`t失敗`t0`t2026/01/02`tエラー`t2"
            "計測\c.docx`t2026/01/01`t1`t未取り込み`t0`t`t`t2"
            "計測\d.docx`t途中"
        )
        $s = readIngestStatusCounts $path
        $s.Done | Should -Be 1
        $s.Failed | Should -Be 1
        $s.Other | Should -Be 1
    }

    It "取り込み一覧が無い・見出しが無いときは失敗にする" {
        { readIngestStatusCounts (Join-Path $TestDrive "無い.tsv") } | Should -Throw "*取り込み一覧がありません*"
        $path = Join-Path $TestDrive "見出し無し.tsv"
        writeTsv $path @("クロール対象フォルダ`tC:\共有`t計測")
        { readIngestStatusCounts $path } | Should -Throw "*見出し*"
    }

    It "成功 + 失敗がファイル数と合わなければ失敗にする" -TestCases @(
        @{ done = 4; failed = 0; total = 4; ok = $true }
        @{ done = 3; failed = 1; total = 4; ok = $true }
        @{ done = 3; failed = 0; total = 4; ok = $false }
        @{ done = 4; failed = 1; total = 4; ok = $false }
    ) {
        param ($done, $failed, $total, $ok)
        $call = { assertIngestCounts @{ Done = $done; Failed = $failed; Other = 0 } $total }
        if ($ok) { $call | Should -Not -Throw } else { $call | Should -Throw "*合いません*" }
    }

    It "段階ごとの秒を出す（同じ段階が続けば足す）" {
        $s = getIngestPhaseSeconds @(@{ Name = "クロール"; StartMs = 0; EndMs = 500 }, @{ Name = "確認"; StartMs = 500; EndMs = 700 }, @{ Name = "取り込み"; StartMs = 700; EndMs = 3200 }, @{ Name = "取り込み"; StartMs = 3200; EndMs = 4200 }, @{ Name = "仕上げ"; StartMs = 4200; EndMs = 4500 })
        $s["クロール"] | Should -Be 0.5
        $s["取り込み"] | Should -Be 3.5
        $s["仕上げ"] | Should -Be 0.3
    }

    It "<name>は失敗にする" -TestCases @(
        @{ name = "取り込みの段階が読めなかったとき"; phases = @(@{ Name = "クロール"; StartMs = 0; EndMs = 500 }, @{ Name = "仕上げ"; StartMs = 500; EndMs = 700 }); message = "*取り込みの段階が読めませんでした*" }
        @{ name = "段階が 1 つも読めなかったとき"; phases = @(); message = "*取り込みの段階が読めませんでした*" }
        @{ name = "知らない段階の名前が来たとき"; phases = @(@{ Name = "取り込み"; StartMs = 0; EndMs = 500 }, @{ Name = "整理"; StartMs = 500; EndMs = 700 }); message = "*知らない段階*" }
        @{ name = "閉じていない段階があるとき"; phases = @(@{ Name = "取り込み"; StartMs = 0; EndMs = $null }); message = "*閉じていません*" }
    ) {
        param ($phases, $message)
        { getIngestPhaseSeconds $phases } | Should -Throw $message
    }

    It "データの種類ごとの数と、使う Office のレーンを出す" -TestCases @(
        @{ files = @("a.docx", "b.PPTX"); lanes = "なし" }
        @{ files = @("a.xlsx", "b.doc"); lanes = "Excel・Word" }
        @{ files = @("a.ppt", "b.docx", "c.txt"); lanes = "PowerPoint" }
    ) {
        param ($files, $lanes)
        $dir = Join-Path $TestDrive ("data-" + [guid]::NewGuid().ToString("N"))
        foreach ($f in $files) { writeTsv "$dir\$f" @("x") }
        $c = getIngestKindCounts $dir
        $c.Total | Should -Be (@($files | Where-Object { $_ -notlike "*.txt" }).Count)
        getIngestLanes $c | Should -Be $lanes
    }

    It "何回か測った結果から、最小・中央値・平均・最大と、リソースを取る回（時間が中央値の回）を出す" {
        $run = { param ($sec, $ing) [pscustomobject]@{
            Files = [pscustomobject]@{ docx = 2 }; Total = 2; Done = 2; Failed = 0; Threads = 2; Lanes = "なし"; Seconds = $sec
            Phases = @([pscustomobject]@{ Phase = "取り込み"; Seconds = $ing }); PerFileMs = ($ing * 500); Resources = @([pscustomobject]@{ Phase = "取り込み"; Marker = $sec }); PeakWorkingSetMB = $sec * 10 } }
        $r = newIngestResult @((& $run 5 4), (& $run 3 2), (& $run 4 3))
        $r.Repeat | Should -Be 3
        $r.Seconds.Min | Should -Be 3
        $r.Seconds.Median | Should -Be 4
        $r.Seconds.Max | Should -Be 5
        $r.PhaseSeconds[0].Phase | Should -Be "取り込み"
        $r.PhaseSeconds[0].Stats.Median | Should -Be 3
        $r.PerFileMs.Median | Should -Be 1500
        $r.ResourceRun | Should -Be 3
        $r.Resources[0].Marker | Should -Be 4
        $r.PeakWorkingSetMB | Should -Be 50
    }
}

Describe "startResourceMonitor（受け渡しの口の段階を読む）" -Tag Io {
    It "口の段階が変わるたびに段階を記録し、止めるときに最後の段階を閉じる" {
        $channel = [hashtable]::Synchronized(@{ Progress = $null })
        $monitor = startResourceMonitor 50 $channel
        # 記録のスレッドが動き出す（最初の記録が入る）まで待つ
        for ($i = 0; $i -lt 100 -and $monitor.Samples.Count -eq 0; $i++) { Start-Sleep -Milliseconds 50 }
        $channel.Progress = @{ Phase = "クロール" }
        Start-Sleep -Milliseconds 150
        $channel.Progress = @{ Phase = "取り込み" }
        Start-Sleep -Milliseconds 150
        $channel.Progress = @{ Phase = "仕上げ" }
        Start-Sleep -Milliseconds 150
        $samples = stopResourceMonitor $monitor
        @($monitor.Phases | ForEach-Object { $_.Name }) | Should -Be @("クロール", "取り込み", "仕上げ")
        @($monitor.Phases | Where-Object { $null -eq $_.EndMs }).Count | Should -Be 0
        @($samples | ForEach-Object { $_.Phase } | Select-Object -Unique) | Should -Contain "取り込み"
    }
}

Describe "new_ingest_data.ps1" -Tag Io {
    BeforeAll {
        $newData = "$PSScriptRoot\..\..\tools\perf\new_ingest_data.ps1"
    }

    It "種類ごとの数のファイルを作り、50 ファイルごとにフォルダを分ける" {
        $dest = Join-Path $TestDrive "data1"
        & $newData -Dest $dest -Docx 30 -Pptx 25 -Doc 1 -Ppt 1 6>$null | Out-Null
        $expected = @{ docx = 30; pptx = 25; doc = 1; ppt = 1 }
        $all = [System.IO.Directory]::GetFiles($dest, "*", "AllDirectories")
        foreach ($ext in $expected.Keys) { @($all | Where-Object { $_.EndsWith(".$ext") }).Count | Should -Be $expected[$ext] }
        $folders = @([System.IO.Directory]::GetDirectories($dest) | Sort-Object)
        $folders.Count | Should -Be 2
        @([System.IO.Directory]::GetFiles($folders[0])).Count | Should -Be 50
        @([System.IO.Directory]::GetFiles($folders[1])).Count | Should -Be 7
        [System.IO.Path]::GetFileName($folders[0]) | Should -Be "フォルダ001"
        [System.IO.File]::Exists("$($folders[0])\資料0001.docx") | Should -Be $true
    }

    It "同じ引数からは、同じ構成・同じ中身になる" {
        $a = Join-Path $TestDrive "same-a"; $b = Join-Path $TestDrive "same-b"
        & $newData -Dest $a -Docx 3 -Pptx 2 6>$null | Out-Null
        & $newData -Dest $b -Docx 3 -Pptx 2 6>$null | Out-Null
        $listA = @(Get-ChildItem $a -Recurse -File | ForEach-Object { $_.FullName.Substring($a.Length) + ":" + (Get-FileHash $_.FullName).Hash })
        $listB = @(Get-ChildItem $b -Recurse -File | ForEach-Object { $_.FullName.Substring($b.Length) + ":" + (Get-FileHash $_.FullName).Hash })
        $listA.Count | Should -Be 5
        $listA | Should -Be $listB
    }

    It ".xlsx は、ブックのフォルダからパスの順に先頭の数冊を写す" {
        $books = Join-Path $TestDrive "books"
        foreach ($n in @("b\2.xlsx", "a\9.xlsx", "a\1.xlsx")) { writeTsv "$books\$n" @($n) }
        $dest = Join-Path $TestDrive "data-xlsx"
        & $newData -Dest $dest -Xlsx 2 -Books $books -Docx 0 -Pptx 0 6>$null | Out-Null
        $files = @([System.IO.Directory]::GetFiles($dest, "*.xlsx", "AllDirectories") | Sort-Object)
        $files.Count | Should -Be 2
        [System.IO.File]::ReadAllText($files[0]) | Should -Match "a\\1.xlsx"
        [System.IO.File]::ReadAllText($files[1]) | Should -Match "a\\9.xlsx"
    }

    It "<name>は失敗にする" -TestCases @(
        @{ name = "置き場所が空でないとき"; run = { $d = Join-Path $TestDrive "used"; writeTsv "$d\x.txt" @("x"); & $newData -Dest $d } ; message = "*空ではありません*" }
        @{ name = ".xlsx が足りないとき"; run = { $b = Join-Path $TestDrive "few"; writeTsv "$b\1.xlsx" @("x"); & $newData -Dest (Join-Path $TestDrive "few-out") -Xlsx 2 -Books $b } ; message = "*足りません*" }
        @{ name = "-Xlsx に -Books が無いとき"; run = { & $newData -Dest (Join-Path $TestDrive "nobooks") -Xlsx 1 } ; message = "*-Books*" }
    ) {
        param ($run, $message)
        { & $run 6>$null } | Should -Throw $message
    }
}

Describe "measure_perf.ps1 -Office" -Tag Io {
    BeforeAll {
        $measure = "$PSScriptRoot\..\..\tools\measure_perf.ps1"
        $newData = "$PSScriptRoot\..\..\tools\perf\new_ingest_data.ps1"
        $repoConfig = Join-Path (Resolve-Path "$PSScriptRoot\..\..").ProviderPath "setting.config"
        $configBefore = if (Test-Path -LiteralPath $repoConfig) { (Get-FileHash $repoConfig).Hash } else { $null }
        $data = Join-Path $TestDrive "office"
        & $newData -Dest $data -Docx 2 -Pptx 2 6>$null | Out-Null
        $work = Join-Path $TestDrive "ingest-work"
        & $measure -Office $data -Work $work -Threads 2 -Repeat 2 -SampleMs 20 -RunId "9" -Sha "abc1234" 6>$null | Out-Null
        $out = Join-Path $work "result"
        $result = [System.IO.File]::ReadAllText("$out\result.json") | ConvertFrom-Json
    }

    It "取り込みを測り、ファイル数・成功・失敗・スレッドの数・時間を書く" {
        $result.Schema | Should -Be 1
        $result.Ingest.Total | Should -Be 4
        $result.Ingest.Done | Should -Be 4
        $result.Ingest.Failed | Should -Be 0
        $result.Ingest.Threads | Should -Be 2
        $result.Ingest.Repeat | Should -Be 2
        $result.Ingest.Lanes | Should -Be "なし"
        $result.Ingest.Files.docx | Should -Be 2
        $result.Ingest.Files.pptx | Should -Be 2
        $result.Ingest.Seconds.Count | Should -Be 2
        $result.Ingest.Seconds.Min | Should -BeGreaterThan 0
        @($result.Ingest.PhaseSeconds | ForEach-Object { $_.Phase }) | Should -Contain "取り込み"
        $result.Ingest.PerFileMs.Median | Should -BeGreaterThan 0
        @($result.Ingest.Resources | ForEach-Object { $_.Phase }) | Should -Contain "取り込み"
        $result.Run.Threads | Should -Be 2
        $result.Run.Office.Excel | Should -Not -BeNullOrEmpty
    }

    It "-Office だけのときは、Index と Search のキーを残して null にする" {
        $names = @($result.PSObject.Properties | ForEach-Object { $_.Name })
        $names | Should -Contain "Index"
        $names | Should -Contain "Search"
        $result.Index | Should -BeNullOrEmpty
        $result.Search | Should -BeNullOrEmpty
        [System.IO.File]::Exists("$out\searches.csv") | Should -Be $false
    }

    It "metrics.csv に ingest_ の行を書く" {
        $metrics = @(Import-Csv "$out\metrics.csv")
        foreach ($stat in @("min", "median", "mean", "max")) {
            @($metrics | Where-Object { $_.metric -eq "ingest_seconds" -and $_.stat -eq $stat }).Count | Should -Be 1
            @($metrics | Where-Object { $_.metric -eq "ingest_per_file_ms" -and $_.stat -eq $stat }).Count | Should -Be 1
            @($metrics | Where-Object { $_.metric -eq "ingest_phase_seconds" -and $_.word -eq "取り込み" -and $_.stat -eq $stat }).Count | Should -Be 1
        }
        @($metrics | Where-Object { $_.metric -eq "ingest_files" }).value | Should -Be "4"
        @($metrics | Where-Object { $_.metric -like "index_*" -or $_.metric -like "search_*" }).Count | Should -Be 0
    }

    It "summary.md に「Office からの取り込み」の表を書き、パスを書かない" {
        $md = [System.IO.File]::ReadAllText("$out\summary.md")
        $md | Should -Match "## Office からの取り込み"
        $md | Should -Match "1 ファイルあたり（ms）"
        $md | Should -Match "EXCEL・WINWORD・POWERPNT のプロセスは含まない"
        $md | Should -Not -Match "## インデックス作成"
        $md | Should -Not -Match "## 検索"
        $md | Should -Not -Match "取り込みに失敗したファイルがある"
        $md.Contains($TestDrive) | Should -Be $false
        [System.IO.File]::ReadAllText("$out\result.json").Contains($TestDrive) | Should -Be $false
        [System.IO.File]::ReadAllText("$out\metrics.csv").Contains($TestDrive) | Should -Be $false
    }

    It "リポジトリの setting.config を作らず、変えない" {
        $after = if (Test-Path -LiteralPath $repoConfig) { (Get-FileHash $repoConfig).Hash } else { $null }
        $after | Should -Be $configBefore
    }

    It "取り込めないファイルは失敗として数え、summary.md に書く" {
        $bad = Join-Path $TestDrive "office-bad"
        & $newData -Dest $bad -Docx 1 -Pptx 0 6>$null | Out-Null
        # 失敗になるテストデータ（読み取りパスワード付きの .docx）を 1 つ足す
        [System.IO.File]::Copy("$PSScriptRoot\..\testdata\office\Word\異常系\読み取りパスワード付き.docx", "$bad\フォルダ001\資料0002.docx")
        $w = Join-Path $TestDrive "bad-work"
        & $measure -Office $bad -Work $w -Repeat 1 -SampleMs 20 6>$null | Out-Null
        $r = [System.IO.File]::ReadAllText("$w\result\result.json") | ConvertFrom-Json
        $r.Ingest.Total | Should -Be 2
        $r.Ingest.Done | Should -Be 1
        $r.Ingest.Failed | Should -Be 1
        [System.IO.File]::ReadAllText("$w\result\summary.md") | Should -Match "取り込みに失敗したファイルがある"
    }

    It "-Index と -Office の両方を指定すると、取り込みも pack の作成も検索も測る" {
        $index = Join-Path $TestDrive "both\index"
        writeTsv "$index\営業\見積.xlsx\シート1.tsv" @("1`t見積書")
        $w = Join-Path $TestDrive "both-work"
        & $measure -Index $index -Office $data -Work $w -Repeat 1 -Count 2 -SampleMs 20 6>$null | Out-Null
        $r = [System.IO.File]::ReadAllText("$w\result\result.json") | ConvertFrom-Json
        $r.Ingest.Total | Should -Be 4
        $r.Index.Pack.Books | Should -Be 1
        @($r.Search).Count | Should -Be 1
        [System.IO.File]::ReadAllText("$w\result\summary.md") | Should -Match "(?s)## Office からの取り込み.*## インデックス作成.*## 検索"
    }

    It "-Index も -Office も無いときは失敗にする" {
        { & $measure -Work (Join-Path $TestDrive "none") 6>$null } | Should -Throw "*-Index と -Office*"
    }
}