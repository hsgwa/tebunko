# 性能の計測（tools\measure_perf.ps1・tools\perf\）のテスト
BeforeAll {
    $measure = "$PSScriptRoot\..\..\tools\measure_perf.ps1"
    . "$PSScriptRoot\..\..\tools\perf\perf_common.ps1"

    function writeTsv([string]$path, [string[]]$lines) {
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
        [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    }
}

Describe "perf_common.ps1 の子のスクリプトの実行" -Tag Unit {
    BeforeAll {
        function runChild([string]$body, [int]$tailLines = 20) {
            $path = Join-Path $TestDrive "child_$([guid]::NewGuid().ToString('N')).ps1"
            [System.IO.File]::WriteAllText($path, $body, (New-Object System.Text.UTF8Encoding($true)))
            invokeChildScript $path @() $tailLines
        }
    }
    # CI で流しても本物の Application ログを読まないよう、既定では空を返す（読む It は自分で差し替える）
    BeforeEach { Mock getChildCrashEvents { @() } }

    It "0 以外で終わると、<name>" -TestCases @(
        @{ name = "出した文（標準出力・標準エラー）と終了コードが例外に入る"; body = "Write-Output 'out-line'; [Console]::Error.WriteLine('err-line'); exit 5"; contains = @("out-line", "err-line", "終了コード 5・0x00000005") }
        @{ name = "何も出さなくても終了コードが入り、出力が無かったと分かる"; body = "exit 5"; contains = @("終了コード 5・0x00000005", "何も出力しませんでした") }
        @{ name = "大きな終了コードは 16 進でも読める"; body = "exit -1073741819"; contains = @("終了コード -1073741819・0xC0000005") }
    ) {
        param ($body, $contains)
        $thrown = $null
        try { runChild $body | Out-Null } catch { $thrown = $_.Exception.Message }
        $thrown | Should -Not -BeNullOrEmpty
        foreach ($text in $contains) { $thrown | Should -BeLike "*$text*" }
    }

    # 本物のイベントログは読まず、getChildCrashEvents を差し替える
    It "0 以外で終わったとき、<name>" -TestCases @(
        @{ name = "CI のときは、取れたイベントが元の文のあとに入る"; ci = "true"; reader = { "10:00:00 Application Error (ID 1000): crash-detail" }; contains = @("終了コード 5", "crash-detail"); calls = 1 }
        @{ name = "CI でないときは、イベントログを読まない"; ci = ""; reader = { "10:00:00 Application Error (ID 1000): crash-detail" }; contains = @("終了コード 5"); calls = 0 }
        @{ name = "読むのに失敗しても、元の文（終了コードと出力）が出る"; ci = "true"; reader = { throw "読めません" }; contains = @("終了コード 5", "child-out"); calls = 1 }
    ) {
        param ($ci, $reader, $contains, $calls)
        $saved = $env:GITHUB_ACTIONS
        $env:GITHUB_ACTIONS = $ci
        try {
            Mock getChildCrashEvents $reader
            $thrown = $null
            try { runChild "Write-Output 'child-out'; exit 5" | Out-Null } catch { $thrown = $_.Exception.Message }
            foreach ($text in $contains) { $thrown | Should -BeLike "*$text*" }
            if ($calls -eq 0) { $thrown | Should -Not -BeLike "*crash-detail*" }
            Should -Invoke getChildCrashEvents -Times $calls -Exactly
        } finally {
            $env:GITHUB_ACTIONS = $saved
        }
    }

    It "例外に入れるのは、出力の終わりの指定行数だけ" {
        $thrown = $null
        try { runChild "1..10 | ForEach-Object { Write-Output ('line' + `$_) }; exit 3" 3 | Out-Null } catch { $thrown = $_.Exception.Message }
        $thrown | Should -BeLike "*line10*"
        $thrown | Should -BeLike "*line8*"
        $thrown | Should -Not -BeLike "*line7*"
    }

    It "成功したときは、子の出力をそのまま流し、例外にしない" {
        $out = @(runChild "Write-Output 'ok-line'; exit 0")
        $out | Should -Contain "ok-line"
    }
}

Describe "perf_common.ps1 の子の異常終了のイベント" -Tag Unit {
    # 提供元ごとの記録（id・m = 本文・s = 秒）を差し替える（本物のログは読まない）。"fail" の提供元は、読むと失敗する
    It "<name>" -TestCases @(
        @{ name = "powershell.exe に関するものだけを取る"; max = 5; len = 400
            by = @{ "Application Error" = @(@{ id = 1000; m = "faulting powershell.exe code c0000005"; s = 1 }, @{ id = 1000; m = "faulting other.exe"; s = 2 }) }
            expect = @("Application Error (ID 1000): faulting powershell.exe code c0000005"); absent = @("other.exe") }
        @{ name = "件数は上限までで、時刻の早い順"; max = 2; len = 400
            by = @{ "Application Error" = @(@{ id = 1000; m = "powershell.exe a"; s = 3 }, @{ id = 1000; m = "powershell.exe b"; s = 1 }); ".NET Runtime" = @(@{ id = 1026; m = "powershell.exe c"; s = 2 }) }
            expect = @("powershell.exe b", "powershell.exe c"); absent = @("powershell.exe a") }
        @{ name = "1 件の長さは上限で切る"; max = 5; len = 20
            by = @{ "Application Error" = @(@{ id = 1000; m = ("powershell.exe " + ("x" * 100)); s = 0 }) }
            expect = @("powershell.exe xxxxx…"); absent = @("xxxxxxxxxxxxxxxxxxxxxxxx") }
        @{ name = "改行や連続した空白は 1 行にまとめる"; max = 5; len = 400
            by = @{ "Application Error" = @(@{ id = 1000; m = "powershell.exe`r`n  Faulting   module`r`n code"; s = 0 }) }
            expect = @("powershell.exe Faulting module code"); absent = @() }
        @{ name = "1 つの提供元が読めなくても、残りの提供元のものを取る"; max = 5; len = 400
            by = @{ "Application Error" = "fail"; ".NET Runtime" = @(@{ id = 1026; m = "powershell.exe survived"; s = 0 }); "Windows Error Reporting" = "fail" }
            expect = @(".NET Runtime (ID 1026): powershell.exe survived"); absent = @() }
        @{ name = "どれも読めなければ空を返す"; max = 5; len = 400
            by = @{ "Application Error" = "fail"; ".NET Runtime" = "fail"; "Windows Error Reporting" = "fail" }
            expect = @(); absent = @() }
    ) {
        param ($by, $max, $len, $expect, $absent)
        $script:fakeEvents = @{}
        foreach ($provider in $by.Keys) {
            $spec = $by[$provider]
            $script:fakeEvents[$provider] = $(if ($spec -is [string]) { $spec } else { @($spec | ForEach-Object {
                [pscustomobject]@{ TimeCreated = [datetime]::new(2026, 1, 1, 10, 0, [int]$_.s); ProviderName = $provider; Id = [int]$_.id; Message = [string]$_.m }
            }) })
        }
        Mock Get-WinEvent {
            $value = $script:fakeEvents[$FilterHashtable.ProviderName]
            if ($null -eq $value -or $value -is [string]) { throw "読めません" }
            return $value
        }
        $got = @(getChildCrashEvents ([datetime]::new(2026, 1, 1, 10, 0, 0)) ([datetime]::new(2026, 1, 1, 10, 0, 10)) $max $len)
        $text = $got -join "`n"
        $got.Count | Should -Be @($expect).Count
        foreach ($part in $expect) { $text | Should -BeLike "*$part*" }
        foreach ($part in $absent) { $text | Should -Not -BeLike "*$part*" }
        $got | ForEach-Object { $_ | Should -Not -BeLike "*`n*" }
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

Describe "perf_common.ps1 の関数の名前の選び方（名前を改めた前後の版）" -Tag Unit {
    BeforeAll {
        # 測る tebunko の関数の代わりに、名前だけが違う関数を置く
        function global:newNameOnlyFunction { }
        function global:oldNameOnlyFunction { }
    }
    AfterAll {
        Remove-Item function:global:newNameOnlyFunction, function:global:oldNameOnlyFunction -ErrorAction SilentlyContinue
    }

    It "<name>" -TestCases @(
        @{ name = "新しい名前だけがあれば、新しい名前"; candidates = @("newNameOnlyFunction", "noSuchFunction"); expect = "newNameOnlyFunction" }
        @{ name = "古い名前だけがあれば、古い名前"; candidates = @("noSuchFunction", "oldNameOnlyFunction"); expect = "oldNameOnlyFunction" }
        @{ name = "両方あれば、先に書いた名前"; candidates = @("oldNameOnlyFunction", "newNameOnlyFunction"); expect = "oldNameOnlyFunction" }
    ) {
        param ($candidates, $expect)
        resolveTebunkoFunction $candidates | Should -Be $expect
    }

    It "どれも無ければ、先頭の名前を挙げて止まる" {
        { resolveTebunkoFunction @("noSuchFunction", "noSuchOldFunction") } | Should -Throw "*noSuchFunction がありません*"
    }

    It "<name>" -TestCases @(
        @{ name = "新しい項目名だけがあれば、その値"; table = @{ ContentIndexFiles = @(1, 2) }; expect = 2 }
        @{ name = "古い項目名だけがあれば、その値"; table = @{ Packs = @(1, 2, 3) }; expect = 3 }
    ) {
        param ($table, $expect)
        @(getTebunkoProperty $table @("ContentIndexFiles", "Packs")).Count | Should -Be $expect
    }
}

Describe "perf_common.ps1 の入力しながらの検索の部品" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "該当無し を -MinLength 1 で（4 つ）"; word = "該当無し"; isRegex = $false; minLength = 1; steps = @("該", "該当", "該当無", "該当無し") }
        @{ name = "該当無し を既定の 3 で（2 つ）"; word = "該当無し"; isRegex = $false; minLength = 3; steps = @("該当無", "該当無し") }
        @{ name = "見積書 を既定で（1 つ）"; word = "見積書"; isRegex = $false; minLength = 3; steps = @("見積書") }
        @{ name = "-MinLength より短い語は、語そのものだけ"; word = "の"; isRegex = $false; minLength = 3; steps = @("の") }
        @{ name = "正規表現 ^\d{5}\t を -MinLength 1 で（^・^\・^\d{5}\ を飛ばす）"; word = '^\d{5}\t'; isRegex = $true; minLength = 1
            steps = @('^\d', '^\d{', '^\d{5', '^\d{5}', '^\d{5}\t') }
        @{ name = "すべて飛ばす正規表現（空）"; word = "^"; isRegex = $true; minLength = 3; steps = @() }
    ) {
        param ($word, $isRegex, $minLength, $steps)
        @(getTypingSteps $word $isRegex $minLength) | Should -Be $steps
    }

    It "サロゲートペアを含む語は、途中で切れる長さを飛ばす" {
        $pair = [char]::ConvertFromUtf32(0x1F600)
        $word = "見" + $pair + "書"
        $withPair = "見" + $pair
        $result = @(getTypingSteps $word $false 1)
        $result.Count | Should -Be 3
        $result | Should -Be @("見", $withPair, $word)
    }

    It "測らなかった語は、Steps 0・Skipped ありで、ほかは null にする" {
        $r = newTypingWordResult -name "無い" -regex $false -steps 0 -attempts @() -skipped "送る語がありません" -peakWorkingSetMB 12.0
        $r.Name | Should -Be "無い"
        $r.Steps | Should -Be 0
        $r.Skipped | Should -Be "送る語がありません"
        $r.Hits | Should -BeNullOrEmpty
        $r.FirstHitMs | Should -BeNullOrEmpty
        $r.PeakWorkingSetMB | Should -Be 12.0
    }

    It "1 回ごとの記録から、2 回目以降の統計・全回の統計・1 回目の値・回数の合計を作る" {
        $attempts = @(
            @{ FirstHitMs = 50; FinishMs = 80; WaitPrevMs = 5; IntervalMs = @(140); LagMs = @(2, 10); RequestMs = @(1, 2); Hits = 1; Truncated = $false; Packs = 5; GramUsed = $false; CancelledSteps = 1 }
            @{ FirstHitMs = $null; FinishMs = 60; WaitPrevMs = 0; IntervalMs = @(145); LagMs = @(1, 3); RequestMs = @(1, 1); Hits = 0; Truncated = $false; Packs = 5; GramUsed = $true; CancelledSteps = 0 }
            @{ FirstHitMs = 40; FinishMs = 70; WaitPrevMs = 2; IntervalMs = @(150); LagMs = @(0, 5); RequestMs = @(1, 2); Hits = 1; Truncated = $false; Packs = 5; GramUsed = $true; CancelledSteps = 2 }
        )
        $r = newTypingWordResult -name "まれ" -regex $false -steps 2 -attempts $attempts
        $r.Skipped | Should -BeNullOrEmpty
        $r.Hits | Should -Be 1   # 最後（3 回目）の要求の結果
        $r.GramUsed | Should -Be 2   # 全回のうち gram 索引を使った回数
        $r.CancelledSteps | Should -Be 2   # 2 回目以降（0 + 2）の合計
        $r.FirstHitMs.Median | Should -Be 40   # 2 回目以降（$null を除く）
        $r.FirstHitMs.Count | Should -Be 1
        $r.FinishMs.Median | Should -Be 65   # 2 回目以降（60, 70）
        $r.WaitPrevMs.Median | Should -Be 1   # 2 回目以降（0, 2）
        $r.IntervalMs.Count | Should -Be 3   # 全回（140, 145, 150）
        $r.IntervalMs.Median | Should -Be 145
        $r.LagMs.Count | Should -Be 6   # 全回・要求ごと
        $r.RequestMs.Count | Should -Be 6
        $r.First.FirstHitMs | Should -Be 50   # 1 回目の値そのもの（統計には入らない）
        $r.First.FinishMs | Should -Be 80
    }

    It "ヒットの無い語（2 回目以降すべて null）は、FirstHitMs の統計が null になる" {
        $attempts = @(
            @{ FirstHitMs = $null; FinishMs = 20; WaitPrevMs = 0; IntervalMs = @(); LagMs = @(0); RequestMs = @(1); Hits = 0; Truncated = $false; Packs = 3; GramUsed = $false; CancelledSteps = 0 }
            @{ FirstHitMs = $null; FinishMs = 18; WaitPrevMs = 0; IntervalMs = @(); LagMs = @(0); RequestMs = @(1); Hits = 0; Truncated = $false; Packs = 3; GramUsed = $false; CancelledSteps = 0 }
        )
        $r = newTypingWordResult -name "無い" -regex $false -steps 1 -attempts $attempts
        $r.Hits | Should -Be 0
        $r.FirstHitMs | Should -BeNullOrEmpty
    }
}

Describe "perf_common.ps1 のプレビューの部品" -Tag Unit {
    BeforeAll {
        function newPreviewRaw([double[]]$uncachedMs, [double[]]$cachedMs, [double]$beforeManaged = 10.0, [double]$afterManaged = 10.4) {
            $rows = @()
            foreach ($v in $uncachedMs) { $rows += [pscustomobject]@{ Cache = "無し"; ReadMs = $v; ReadTextCalls = 1; BuildMs = 2.0; Cells = 400; FileMB = 1.5 } }
            foreach ($v in $cachedMs) { $rows += [pscustomobject]@{ Cache = "あり"; ReadMs = $v; ReadTextCalls = 0; BuildMs = 3.0; Cells = 800; FileMB = 1.5 } }
            return [pscustomobject]@{
                Name = "語"; Regex = $false; Count = $uncachedMs.Count; Previews = $rows
                Memory = [pscustomobject]@{ AfterSearchWorkingSetMB = 200.0; BeforePreviewManagedMB = $beforeManaged; AfterPreviewWorkingSetMB = 230.0
                    AfterGcWorkingSetMB = 210.0; AfterCacheDropWorkingSetMB = 150.0; AfterGcManagedMB = $afterManaged; CacheChars = 5000000 }
                Render = $null; PeakWorkingSetMB = 300.0
            }
        }
    }

    It "1 回ごとの記録から、キャッシュの有無ごとの読み込みの統計と、残ったメモリ・キャッシュの減りを作る" {
        $r = newPreviewResult (newPreviewRaw @(10, 30, 20) @(1, 3, 2))
        $r.Previews | Should -Be 6
        $r.ReadUncachedMs.Median | Should -Be 20
        $r.ReadUncachedMs.Max | Should -Be 30
        $r.ReadCachedMs.Median | Should -Be 2
        $r.BuildMs.Count | Should -Be 6
        $r.Cells.Max | Should -Be 800
        $r.FileMBMax | Should -Be 1.5
        $r.RetainedManagedMB | Should -Be 0.4
        $r.RetainedPerPreviewMB | Should -Be 0.067
        $r.CacheDropMB | Should -Be 60
        $r.CacheChars | Should -Be 5000000
        $r.ReadTextCallsUncached | Should -Be 3   # キャッシュ無しの 3 回で 1 回ずつ。キャッシュありの分は数えない
        $r.FullRead | Should -BeNullOrEmpty   # 前の読み方は、呼ぶ側が測ったときだけ入れる
        $r.Render | Should -BeNullOrEmpty
        # 表を作る時間は、最初の 1 回（キャッシュ無しの 1 件目）と、2 回目以降（残り 5 回）に分ける
        $r.BuildFirstMs | Should -Be 2
        $r.BuildRestMs.Count | Should -Be 5
        $r.BuildRestMs.Max | Should -Be 3
    }

    It "描画の結果は、1 回目と 2 回目以降の統計に分け、1 回だけのときは 2 回目以降が null になる" {
        $raw = newPreviewRaw @(10) @(1)
        $raw.Render = [pscustomobject]@{ Cells = 900; Rows = 30; Columns = 6; RenderMs = @(200.0, 40.0, 20.0, 30.0, 25.0); WorkingSetBeforeMB = 250.0; WorkingSetAfterMB = 270.0 }
        $r = newPreviewResult $raw
        $r.Render.FirstMs | Should -Be 200
        $r.Render.RestMs.Count | Should -Be 4
        $r.Render.RestMs.Median | Should -Be 27.5
        $r.Render.RestMs.Max | Should -Be 40
        @($r.Render.RenderMs).Count | Should -Be 5
        ($r.Render.Cells, $r.Render.Rows, $r.Render.Columns) | Should -Be @(900, 30, 6)
        $r.Render.WorkingSetAfterMB | Should -Be 270
        $single = newRenderResult ([pscustomobject]@{ Cells = 1; Rows = 1; Columns = 1; RenderMs = @(5.0); WorkingSetBeforeMB = 1.0; WorkingSetAfterMB = 1.0 })
        $single.FirstMs | Should -Be 5
        $single.RestMs | Should -BeNullOrEmpty
    }

    It "大きな本文の TSV は、指定した行だけに語が入り、行番号は 4 桁までで、1 行ずつ CRLF で終わる" {
        $text = newLargeContentTsvText 200 "架空語" @(10, 150, 200) 2
        $rows = @($text -split "`r`n" | Where-Object { $_ -ne "" })
        $rows.Count | Should -Be 200
        $hitRows = @($rows | Where-Object { $_.Contains("架空語") })
        $hitRows.Count | Should -Be 3
        @($hitRows | ForEach-Object { ($_ -split "`t")[0] }) | Should -Be @("10", "150", "200")
        $rows[0] | Should -Match '^1\t行1\t架空説明 '
        ($rows | Where-Object { $_ -match '^\d{5}\t' }).Count | Should -Be 0
        ($rows | Where-Object { $_ -match '見積書|の' }).Count | Should -Be 0   # perfdata の語に当たらない
        $text.EndsWith("`r`n") | Should -Be $true
        $text.Contains("`n") | Should -Be $true
        ([regex]::Matches($text, "(?<!`r)`n")).Count | Should -Be 0
    }

    It "大きな本文の TSV は、繰り返しで長くなり、行の数が 1〜9999 の外なら失敗にする" {
        (newLargeContentTsvText 5 "語" @(1) 10).Length | Should -BeGreaterThan (newLargeContentTsvText 5 "語" @(1) 1).Length
        { newLargeContentTsvText 0 "語" @(1) } | Should -Throw "*1〜9999*"
        { newLargeContentTsvText 10000 "語" @(1) } | Should -Throw "*1〜9999*"
    }

    It "横に広い本文の TSV は、指定した行の指定した列にだけ語が入り、1 行のセルの数は指定どおりになる" {
        $text = newWideContentTsvText 5 40 "架空語" @(2, 4) 30
        $rows = @($text -split "`r`n" | Where-Object { $_ -ne "" })
        $rows.Count | Should -Be 5
        foreach ($row in $rows) { @($row -split "`t").Count | Should -Be 41 }   # 行番号 + 40 セル
        $hitRows = @($rows | Where-Object { $_.Contains("架空語") })
        @($hitRows | ForEach-Object { ($_ -split "`t")[0] }) | Should -Be @("2", "4")
        foreach ($row in $hitRows) { ($row -split "`t")[30] | Should -Be "架空語" }
        $rows[0] | Should -Match '^1\t行1列1\t行1列2\t'
        $text.EndsWith("`r`n") | Should -Be $true
        ([regex]::Matches($text, "(?<!`r)`n")).Count | Should -Be 0
    }

    It "<name>" -TestCases @(
        @{ name = "横に広い本文の TSV は、行が 0 なら失敗にする"; call = { newWideContentTsvText 0 10 "語" @(1) }; message = "*1〜9999*" }
        @{ name = "横に広い本文の TSV は、列が 0 なら失敗にする"; call = { newWideContentTsvText 5 0 "語" @(1) 1 }; message = "*columns*" }
        @{ name = "横に広い本文の TSV は、語の列が列の数を超えれば失敗にする"; call = { newWideContentTsvText 5 10 "語" @(1) 11 }; message = "*hitColumn*" }
    ) {
        param ($call, $message)
        $call | Should -Throw $message
    }

    It "読み足しの記録は、回ごとの時間を統計にし、行・列・セルの数は最後の回のものにする" {
        $runs = @(
            [pscustomobject]@{ ReadMs = 10.0; BuildMs = 4.0; CellsMs = 2.0; Rows = 200; TableRows = 250; Columns = 2000; CreatedCells = 60000 }
            [pscustomobject]@{ ReadMs = 30.0; BuildMs = 6.0; CellsMs = 4.0; Rows = 200; TableRows = 450; Columns = 2000; CreatedCells = 64000 }
        )
        $r = newReadMoreResult $runs
        $r.Runs | Should -Be 2
        $r.Complete | Should -Be $false -Because "5 回に届かなければ、少ない回数の値だと分かるようにする"
        (newReadMoreResult @(1..5 | ForEach-Object { $runs[0] })).Complete | Should -Be $true
        $short = [pscustomobject]@{ ReadMs = 1.0; BuildMs = 1.0; CellsMs = 1.0; Rows = 0; TableRows = 250; Columns = 2000; CreatedCells = 60000 }
        (newReadMoreResult @($runs[0], $runs[0], $runs[0], $runs[0], $short)).Complete | Should -Be $false -Because "5 回目が空・短いときは、5 回分とは言えない"
        $r.ReadMs.Max | Should -Be 30
        $r.BuildMs.Median | Should -Be 5
        $r.CellsMs.Count | Should -Be 2
        ($r.TableRows, $r.Columns, $r.CreatedCells) | Should -Be @(450, 2000, 64000)
    }

    It "プレビューの結果は、作ったセルの数と読み足しを引き継ぎ、記録が無ければ null にする" {
        $raw = newPreviewRaw @(10) @(1)
        (newPreviewResult $raw).CreatedCells | Should -BeNullOrEmpty
        (newPreviewResult $raw).ReadMore | Should -BeNullOrEmpty
        foreach ($row in $raw.Previews) { $row | Add-Member -NotePropertyName CreatedCells -NotePropertyValue 500 }
        $raw | Add-Member -NotePropertyName ReadMore -NotePropertyValue @([pscustomobject]@{ ReadMs = 1.0; BuildMs = 1.0; CellsMs = 1.0; Rows = 200; TableRows = 250; Columns = 6; CreatedCells = 90 })
        $r = newPreviewResult $raw
        $r.CreatedCells.Max | Should -Be 500
        $r.ReadMore.Runs | Should -Be 1
        $r.ReadMore.CreatedCells | Should -Be 90
    }

    It "描画の結果は、部品の数・スクロールの時間・残るメモリを引き継ぐ" {
        $r = newRenderResult ([pscustomobject]@{ Cells = 1; Rows = 1; Columns = 1; Elements = 321; ScrollMs = 12.5; RenderRetainedMB = 3.2; RenderMs = @(5.0); WorkingSetBeforeMB = 1.0; WorkingSetAfterMB = 1.0 })
        ($r.Elements, $r.ScrollMs, $r.RenderRetainedMB) | Should -Be @(321, 12.5, 3.2)
    }

    It "<name>" -TestCases @(
        @{ name = "全部読む経路の記録で、キャッシュ無しの全部が全体を読んでいれば、問題は 0 件"; rows = @([pscustomobject]@{ N = 1; Cache = "無し"; ReadTextCalls = 1 }, [pscustomobject]@{ N = 2; Cache = "あり"; ReadTextCalls = 0 }); expect = @() }
        @{ name = "キャッシュ無しで全体を読んでいない回があれば、その番号を挙げる"; rows = @([pscustomobject]@{ N = 1; Cache = "無し"; ReadTextCalls = 1 }, [pscustomobject]@{ N = 2; Cache = "無し"; ReadTextCalls = 0 }); expect = @("N=2 で readContentIndexText が呼ばれていません") }
        @{ name = "記録が空なら、問題は 0 件"; rows = @(); expect = @() }
    ) {
        param ($rows, $expect)
        $problems = @(getFullReadPathProblems $rows)
        $problems.Count | Should -Be $expect.Count
        for ($i = 0; $i -lt $expect.Count; $i++) { $problems[$i] | Should -Be $expect[$i] }
    }

    It "<name>" -TestCases @(
        @{ name = "上限の内なら合格"; preview = { newPreviewResult (newPreviewRaw @(10, 12, 11) @(1, 1, 1)) }; read = 50; retained = 1; count = 3; expect = @() }
        @{ name = "読み込みの中央値が上限を超えたら、その数字を挙げる"; preview = { newPreviewResult (newPreviewRaw @(60, 70, 80) @(1, 1, 1)) }; read = 50; retained = 1; count = 3; expect = @("*中央値 70.0 ms（上限 50.0 ms）*") }
        @{ name = "残るメモリの合計が上限を超えたら、その数字を挙げる"; preview = { newPreviewResult (newPreviewRaw @(10, 12, 11) @(1, 1, 1) 10 70) }; read = 50; retained = 1; count = 3; expect = @("*60.0 MB（上限 1.0 MB）*") }
        @{ name = "測ったプレビューの数が期待と違えば挙げる"; preview = { newPreviewResult (newPreviewRaw @(10, 12) @(1, 1)) }; read = 50; retained = 1; count = 3; expect = @("*4 回（期待 6 回）*") }
        @{ name = "結果が無ければ、結果が無いと挙げる"; preview = { $null }; read = 50; retained = 1; count = 3; expect = @("*結果がありません*") }
    ) {
        param ($preview, $read, $retained, $count, $expect)
        # 空の配列は 1 つの値として返るので、@() で包まずに受けて展開する
        $problems = getPreviewPerfProblems (& $preview) $read $retained $count
        $problems.Count | Should -Be $expect.Count
        for ($i = 0; $i -lt $expect.Count; $i++) { $problems[$i] | Should -BeLike $expect[$i] }
    }

    # perf_preview.Tests.ps1（Slow）が失敗メッセージを作るのと同じ関数。合格のときに空の 1 件が入らないことを確かめる
    It "<name>" -TestCases @(
        @{ name = "合格なら、メッセージは 0 件"; preview = { newPreviewResult (newPreviewRaw @(10, 12, 11) @(1, 1, 1)) }; expect = @() }
        @{ name = "超えたら、語の名前を付けた文が 1 件ずつ入り、空の文は入らない"; preview = { newPreviewResult (newPreviewRaw @(60, 70, 80) @(1, 1, 1) 10 70) }; expect = @("プレビュー まれ: プレビューの読み込み*", "プレビュー まれ: プレビューのあとに残るメモリ（合計）*") }
    ) {
        param ($preview, $expect)
        $messages = @(getPreviewPerfMessages "まれ" (& $preview) 50 1 3)
        $messages.Count | Should -Be $expect.Count
        for ($i = 0; $i -lt $expect.Count; $i++) { $messages[$i] | Should -BeLike $expect[$i] }
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
        # -Typing 2 -TypingIntervalMs 50 -TypingMinLength 1 は、取り消しの起きる並びで流すため（既定の 3 文字からでは、この words.tsv の語は取り消しが起きない）
        & $measure -Index $index -Work (Join-Path $TestDrive "work") -Words $words -Count 3 -SampleMs 50 -Label "テスト" -RunId "1" -Ref "main" -Sha "abc1234" -Scale "0.1" -DataSeconds 1.5 `
            -Typing 2 -TypingIntervalMs 50 -TypingMinLength 1 6>$null | Out-Null
        $result = [System.IO.File]::ReadAllText("$out\result.json") | ConvertFrom-Json
    }

    It "形式の版と実行の情報を書く" {
        $result.Schema | Should -Be 5
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
        $md | Should -Match "## 入力しながらの検索"
        $md.Contains($TestDrive) | Should -Be $false
    }

    It "-Typing を付けると、語ごとの入力しながらの検索の値が出る" {
        $result.Typing.Mode | Should -Be "service"
        $result.Typing.MinLength | Should -Be 1
        $byName = @{}
        foreach ($t in $result.Typing.Words) { $byName[$t.Name] = $t }
        foreach ($t in $result.Typing.Words) {
            $t.FinishMs.Count | Should -Be 1   # 2 回目以降（-Typing 2 なので 1 回分）
            $t.LagMs.Min | Should -Not -BeLessThan 0   # 予定より早く渡さない
        }
        $byName["まれ"].Hits | Should -Be 1   # 一発の検索と同じ
        $byName["数字"].Hits | Should -Be 2
        $byName["無い"].FirstHitMs | Should -BeNullOrEmpty   # ヒットが無い語
        $byName["まれ"].FirstHitMs.Max | Should -Not -BeGreaterThan $byName["まれ"].FinishMs.Max
        @(Import-Csv "$out\typing.csv").Count | Should -Be 6   # 語の数（3）× -Typing（2）
        $metrics = @(Import-Csv "$out\metrics.csv")
        @($metrics | Where-Object { $_.metric -eq "typing_finish_ms" -and $_.stat -eq "median" }).Count | Should -Be 3
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
        # -Typing を渡さなかったときも、Typing のキーは残して null にする
        @($second.PSObject.Properties | ForEach-Object { $_.Name }) | Should -Contain "Typing"
        $second.Typing | Should -BeNullOrEmpty
        # -Preview を渡さなかったときも、Preview のキーは残して null にし、プレビューの表・ファイルは出さない
        @($second.PSObject.Properties | ForEach-Object { $_.Name }) | Should -Contain "Preview"
        $second.Preview | Should -BeNullOrEmpty
        [System.IO.File]::Exists("$again\result\preview.csv") | Should -Be $false
        [System.IO.File]::ReadAllText("$again\result\summary.md") | Should -Not -Match "## プレビュー"
        @(Import-Csv "$again\result\metrics.csv" | Where-Object { $_.metric -like "preview_*" }).Count | Should -Be 0
    }

    It "-Preview を付けると、ヒットのある語ごとにキャッシュの有無でプレビューを測り、ヒットの無い語は飛ばす" {
        $w = Join-Path $TestDrive "preview"
        & $measure -Index $index -Work $w -Words $words -Count 1 -Preview 3 -SampleMs 50 6>$null | Out-Null
        $r = [System.IO.File]::ReadAllText("$w\result\result.json") | ConvertFrom-Json
        $r.Preview.Count | Should -Be 3
        $r.Preview.Render | Should -Be $false
        @($r.Preview.Skipped) | Should -Be @("無い")
        $byName = @{}
        foreach ($p in $r.Preview.Words) { $byName[$p.Name] = $p }
        @($r.Preview.Words).Count | Should -Be 2
        # ヒットが 1 件・2 件の語は、その件数をキャッシュ無し・ありの 2 回ずつ測る
        $byName["まれ"].Previews | Should -Be 2
        $byName["数字"].Previews | Should -Be 4
        foreach ($p in $r.Preview.Words) {
            $p.ReadUncachedMs.Count | Should -Be ($p.Previews / 2)
            $p.ReadCachedMs.Count | Should -Be ($p.Previews / 2)
            $p.Cells.Min | Should -BeGreaterThan 0
            $p.CacheChars | Should -BeGreaterThan 0
            $p.RetainedPerPreviewMB | Should -Not -BeNullOrEmpty
            $p.Render | Should -BeNullOrEmpty
            $p.FullRead | Should -BeNullOrEmpty   # -PreviewFullRead を付けていない
            $p.BuildFirstMs | Should -BeGreaterThan 0   # 最初に作った表の時間
            $p.BuildRestMs.Count | Should -Be ($p.Previews - 1)   # 残りは 2 回目以降
        }
        @(Import-Csv "$w\result\preview.csv").Count | Should -Be 6
        @(Import-Csv "$w\result\metrics.csv" | Where-Object { $_.metric -eq "preview_build_first_ms" }).Count | Should -Be 2
        @(Import-Csv "$w\result\metrics.csv" | Where-Object { $_.metric -eq "preview_build_rest_ms" -and $_.stat -eq "median" }).Count | Should -Be 2
        [System.IO.File]::ReadAllText("$w\result\summary.md") | Should -Match "表を作る 1 回目"
        $metrics = @(Import-Csv "$w\result\metrics.csv")
        @($metrics | Where-Object { $_.metric -eq "preview_read_uncached_ms" -and $_.word -eq "まれ" -and $_.stat -eq "median" }).Count | Should -Be 1
        $md = [System.IO.File]::ReadAllText("$w\result\summary.md")
        $md | Should -Match "## プレビュー"
        $md | Should -Match "### プレビューのメモリ"
        $md | Should -Not -Match "### プレビューの描画"
        $md | Should -Not -Match "### 前の読み方"
        $md.Contains($TestDrive) | Should -Be $false
    }

    It "-ReadMore を付けると、ヒットのある語ごとに読み足しの記録（回数・行・列・セルの数）と、その表の行・指標が出る" {
        $w = Join-Path $TestDrive "preview-more"
        & $measure -Index $index -Work $w -Words $words -Count 1 -Preview 3 -ReadMore -SampleMs 50 6>$null | Out-Null
        $r = [System.IO.File]::ReadAllText("$w\result\result.json") | ConvertFrom-Json
        foreach ($p in $r.Preview.Words) {
            $p.ReadMore | Should -Not -BeNullOrEmpty
            $p.ReadMore.Runs | Should -BeGreaterThan 0   # 読み足す行が尽きたら 5 回に満たず止まる
            $p.ReadMore.Columns | Should -BeGreaterThan 0
            $p.CreatedCells.Count | Should -Be $p.Previews
        }
        @(Import-Csv "$w\result\metrics.csv" | Where-Object { $_.metric -eq "preview_read_more_ms" -and $_.stat -eq "median" }).Count | Should -Be 2
        [System.IO.File]::ReadAllText("$w\result\summary.md") | Should -Match "読み足し"
    }

    It "-PreviewFullRead を付けると、前の読み方（毎回全部読む）でも測り、全体を読んだ回数と並べて出す" {
        $w = Join-Path $TestDrive "preview-full"
        & $measure -Index $index -Work $w -Words $words -Count 1 -Preview 3 -PreviewFullRead -SampleMs 50 6>$null | Out-Null
        $r = [System.IO.File]::ReadAllText("$w\result\result.json") | ConvertFrom-Json
        $r.Schema | Should -Be 5
        foreach ($p in $r.Preview.Words) {
            $p.FullRead | Should -Not -BeNullOrEmpty
            $p.FullRead.Name | Should -Be $p.Name
            # 全部読む経路は、キャッシュ無しの選択のたびに全体を読む
            $p.FullRead.ReadTextCallsUncached | Should -Be ($p.Previews / 2)
            $p.FullRead.ReadUncachedMs.Count | Should -Be ($p.Previews / 2)
        }
        $metrics = @(Import-Csv "$w\result\metrics.csv")
        @($metrics | Where-Object { $_.metric -eq "preview_full_read_uncached_ms" -and $_.word -eq "まれ" -and $_.stat -eq "median" }).Count | Should -Be 1
        @($metrics | Where-Object { $_.metric -eq "preview_retained_mb" -and $_.word -eq "まれ" }).Count | Should -Be 1
        $md = [System.IO.File]::ReadAllText("$w\result\summary.md")
        $md | Should -Match "### 前の読み方"
        $md.Contains($TestDrive) | Should -Be $false
    }

    It "<name>は失敗にする" -TestCases @(
        @{ name = "-Preview -1"; extraArgs = @{ Preview = -1 }; expectedMessage = "*-Preview は 0 以上*" }
        @{ name = "-Render だけ（-Preview が無い）"; extraArgs = @{ Render = $true }; expectedMessage = "*-Render は -Preview*" }
        @{ name = "-PreviewFullRead だけ（-Preview が無い）"; extraArgs = @{ PreviewFullRead = $true }; expectedMessage = "*-PreviewFullRead は -Preview*" }
        @{ name = "-ReadMore だけ（-Preview が無い）"; extraArgs = @{ ReadMore = $true }; expectedMessage = "*-ReadMore は -Preview*" }
        @{ name = "-Typing 1"; extraArgs = @{ Typing = 1 }; expectedMessage = "*-Typing は 0 か 2 以上*" }
    ) {
        param ($extraArgs, $expectedMessage)
        $w = Join-Path $TestDrive ("typing-bad-" + [guid]::NewGuid().ToString("N"))
        { & $measure -Index $index -Work $w -Words $words -Count 1 -SampleMs 50 @extraArgs 6>$null | Out-Null } | Should -Throw -ExpectedMessage $expectedMessage
        # 始める前（インデックス作成・検索の前）に失敗するため、結果のフォルダは作られない
        (Test-Path -LiteralPath (Join-Path $w "result")) | Should -Be $false
    }

    It "実行の日時は、現在のカルチャ（暦・時刻の区切り）に左右されず ISO 8601 の形で書く" {
        # 仏暦（th-TH。年が 543 大きい）で、時刻の区切りも変えたカルチャにする
        $culture = New-Object System.Globalization.CultureInfo("th-TH", $false)
        $culture.DateTimeFormat.TimeSeparator = "."
        $thread = [System.Threading.Thread]::CurrentThread
        $saved = $thread.CurrentCulture
        $thread.CurrentCulture = $culture
        try {
            & $measure -Index $index -Work (Join-Path $TestDrive "culture") -Words $words -Count 1 -SampleMs 50 6>$null | Out-Null
        } finally {
            $thread.CurrentCulture = $saved
        }
        $r = [System.IO.File]::ReadAllText("$TestDrive\culture\result\result.json") | ConvertFrom-Json
        $r.Run.Date | Should -Match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'
        $year = [int]$r.Run.Date.Substring(0, 4)
        $year | Should -BeLessThan 2400
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
        @{ name = "最初に読めた段階がクロールでないとき（記録の開始が遅れた）"; phases = @(@{ Name = "取り込み"; StartMs = 0; EndMs = 500 }, @{ Name = "仕上げ"; StartMs = 500; EndMs = 700 }); message = "*最初に読めた段階*" }
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

Describe "measure_perf.ps1 -Xlsx" -Tag Io {
    BeforeAll {
        $measure = "$PSScriptRoot\..\..\tools\measure_perf.ps1"
        $repoConfig = Join-Path (Resolve-Path "$PSScriptRoot\..\..").ProviderPath "setting.config"
        $configBefore = if (Test-Path -LiteralPath $repoConfig) { (Get-FileHash $repoConfig).Hash } else { $null }
        $work = Join-Path $TestDrive "xlsx-work"
        & $measure -Xlsx 2 -XlsxRows 50 -XlsxManyRows 20 -Work $work -SampleMs 20 -RunId "9" -Sha "abc1234" 6>$null | Out-Null
        $out = Join-Path $work "result"
        $result = [System.IO.File]::ReadAllText("$out\result.json") | ConvertFrom-Json
    }

    It "Excel の流れ読みを測り、読んだ場所の数・行の数・失敗の数・時間を書く" {
        $result.Schema | Should -Be 5
        $result.Xlsx.Repeat | Should -Be 2
        $result.Xlsx.Rows | Should -Be 50
        $result.Xlsx.ManyRows | Should -Be 20
        $result.Xlsx.NormalUnits | Should -Be 1
        $result.Xlsx.NormalLines | Should -Be 2
        $result.Xlsx.ManyLines | Should -Be (2 * $result.Xlsx.ManySheets)
        $result.Xlsx.ManyUnits | Should -Be $result.Xlsx.ManySheets
        $result.Xlsx.NormalFailures | Should -Be 0
        $result.Xlsx.ManyFailures | Should -Be 0
        $result.Xlsx.NormalMs.Count | Should -Be 2
        $result.Xlsx.ManyMs | Should -BeGreaterOrEqual 0
        @($result.Xlsx.Resources | ForEach-Object { $_.Phase }) | Should -Contain "普通のブック"
        @($result.Xlsx.Resources | ForEach-Object { $_.Phase }) | Should -Contain "同じ部品を指すブック"
    }

    It "-Xlsx だけのときは、Index・Search・Ingest のキーを残して null にする" {
        $names = @($result.PSObject.Properties | ForEach-Object { $_.Name })
        foreach ($name in @("Index", "Search", "Ingest", "Xlsx")) { $names | Should -Contain $name }
        $result.Index | Should -BeNullOrEmpty
        $result.Search | Should -BeNullOrEmpty
        $result.Ingest | Should -BeNullOrEmpty
        [System.IO.File]::Exists("$out\searches.csv") | Should -Be $false
        [System.IO.File]::Exists("$out\resource-xlsx.csv") | Should -Be $true
    }

    It "metrics.csv に xlsx_ の行を書く" {
        $metrics = @(Import-Csv "$out\metrics.csv")
        foreach ($stat in @("min", "median", "mean", "max")) {
            @($metrics | Where-Object { $_.metric -eq "xlsx_normal_ms" -and $_.stat -eq $stat }).Count | Should -Be 1
        }
        foreach ($name in @("xlsx_normal_units", "xlsx_normal_failures", "xlsx_many_ms", "xlsx_many_units", "xlsx_many_failures", "xlsx_peak_working_set_mb")) {
            @($metrics | Where-Object { $_.metric -eq $name }).Count | Should -Be 1
        }
        @($metrics | Where-Object { $_.metric -eq "xlsx_normal_failures" }).value | Should -Be "0"
        @($metrics | Where-Object { $_.metric -like "index_*" -or $_.metric -like "search_*" -or $_.metric -like "ingest_*" }).Count | Should -Be 0
    }

    It "summary.md に「Excel の流れ読み」の表を書き、パスを書かない" {
        $md = [System.IO.File]::ReadAllText("$out\summary.md")
        $md | Should -Match "## Excel の流れ読み"
        $md | Should -Not -Match "## インデックス作成"
        $md | Should -Not -Match "## 検索"
        $md.Contains($TestDrive) | Should -Be $false
        [System.IO.File]::ReadAllText("$out\result.json").Contains($TestDrive) | Should -Be $false
        [System.IO.File]::ReadAllText("$out\metrics.csv").Contains($TestDrive) | Should -Be $false
    }

    It "リポジトリの setting.config を作らず、変えない" {
        $after = if (Test-Path -LiteralPath $repoConfig) { (Get-FileHash $repoConfig).Hash } else { $null }
        $after | Should -Be $configBefore
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
        $result.Schema | Should -Be 5
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
        { & $measure -Work (Join-Path $TestDrive "none") 6>$null } | Should -Throw "*-Index・-Office・-Xlsx*"
    }
}
