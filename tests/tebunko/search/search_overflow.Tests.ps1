# 件数の上限を超えた分の書き溜め（tebunko\search\search_overflow.ps1・content_index_search.ps1 の合図と数えるだけの照合・search_run.ps1 の置き場所の決め方）のテスト。
# 表（Runs）+ 溢れのファイルを全件の順に並べると、上限なし（-limit 0）の結果と同じになることを確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:toRows {
        # ヒットを溢れの 1 行と同じ形（RelDir<TAB>Book<TAB>Location<TAB>LineNumber<TAB>Line）にする
        param ($hits)
        return @($hits | ForEach-Object { "$($_.RelDir)`t$($_.Book)`t$($_.Location)`t$($_.LineNumber)`t$($_.Line)" })
    }

    function script:newOverflowRoot {
        # テストごとの tmp の根（TmpRoot）を作って返す
        param ([string]$name)
        $tmp = Join-Path $TestDrive $name
        [void][System.IO.Directory]::CreateDirectory($tmp)
        return $tmp
    }
}

Describe "上限を超えた分の書き溜め" -Tag Io {
    BeforeDiscovery {
        $cases = foreach ($limit in 1, 5, 6, 7, 11, 12, 13) {
            foreach ($mode in "single", "parallel") {
                foreach ($word in "多い", "後ろだけ", "同じ行に多い") {
                    @{ Limit = $limit; Mode = $mode; Word = $word; Name = "上限 $limit・$mode・$word" }
                }
            }
        }
    }

    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "ov_tsv"
        for ($i = 0; $i -lt 6; $i++) {
            $lines = @("多い $i", "x", "多い 多い $i", "同じ行に多い 多い 多い")
            if ($i -ge 4) { $lines += "後ろだけ $i" }
            newTsv "$tsvRoot\d$i\book$i.xlsx\S.tsv" $lines
        }
        $contentIndexFiles = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ov_content_index")
    }

    It "<Name>: 表 + 溢れのファイルが、上限なしの結果と同じ順・同じ件数になる" -ForEach $cases {
        $tmp = newOverflowRoot "t-$Limit-$Mode-$Word"
        $dir = getSearchOverflowDir $tmp 1
        $workers = if ($Mode -eq "single") { 1 } else { 3 }
        $expected = toRows (expandHitRuns (searchContentIndex $Word $contentIndexFiles $true 0 -workerCount $workers -taskBytes 1).Runs)
        $result = searchContentIndex $Word $contentIndexFiles $true $Limit -workerCount $workers -taskBytes 1 -checkEvery 1 -overflowDir $dir
        $rows = @(toRows (expandHitRuns $result.Runs)) + @(readSearchOverflowLines $dir)
        $rows.Count | Should -Be $expected.Count
        ($rows -join "`n") | Should -Be ($expected -join "`n")
        $result.HitTotal | Should -Be $expected.Count
        $result.Truncated | Should -Be ($expected.Count -gt $Limit)
        (getHitRunCount $result.Runs) | Should -Be ([Math]::Min($Limit, $expected.Count))
        $result.OverflowFailed | Should -BeFalse
    }

    It "正規表現（行頭・先読み・空行）でも、数えるだけの照合の件数と書き溜めた行が、上限なしの結果と同じになる" -ForEach @(
        @{ Word = "^多い"; Workers = 1 }
        @{ Word = "^多い"; Workers = 3 }
        @{ Word = "(?!x)多い"; Workers = 1 }
        @{ Word = "(?!x)多い"; Workers = 3 }
        @{ Word = "^$"; Workers = 1 }
        @{ Word = "^$"; Workers = 3 }
    ) {
        $tmp = newOverflowRoot "t-regex-$Workers-$($Word -replace '\W', '_')"
        $dir = getSearchOverflowDir $tmp 1
        $expected = toRows (expandHitRuns (searchContentIndex $Word $contentIndexFiles $false 0 -workerCount $Workers -taskBytes 1).Runs)
        $result = searchContentIndex $Word $contentIndexFiles $false 2 -workerCount $Workers -taskBytes 1 -checkEvery 1 -overflowDir $dir
        $rows = @(toRows (expandHitRuns $result.Runs)) + @(readSearchOverflowLines $dir)
        ($rows -join "`n") | Should -Be ($expected -join "`n")
        $result.HitTotal | Should -Be $expected.Count
    }

    It "上限に届かなければ、溢れのフォルダを作らない" {
        $tmp = newOverflowRoot "t-none"
        $dir = getSearchOverflowDir $tmp 1
        $result = searchContentIndex "後ろだけ" $contentIndexFiles $true 100 -workerCount 1 -overflowDir $dir
        $result.Truncated | Should -BeFalse
        $result.HitTotal | Should -Be 2
        [System.IO.Directory]::Exists($dir) | Should -BeFalse
    }

    It "溢れの置き場所が空なら、数えるだけで書かない（HitTotal は合う）" {
        $result = searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1
        (getHitRunCount $result.Runs) | Should -Be 3
        $result.Truncated | Should -BeTrue
        $result.HitTotal | Should -Be (getHitRunCount (searchContentIndex "多い" $contentIndexFiles $true 0).Runs)
    }

    It "書けない置き場所（同じ名前のファイルがある）は OverflowFailed にし、表と HitTotal はそのまま" {
        $tmp = newOverflowRoot "t-fail"
        $dir = getSearchOverflowDir $tmp 1
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($dir))
        Set-Content -LiteralPath $dir -Value "x"
        $all = (getHitRunCount (searchContentIndex "多い" $contentIndexFiles $true 0).Runs)
        $result = searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1 -overflowDir $dir
        $result.OverflowFailed | Should -BeTrue
        (getHitRunCount $result.Runs) | Should -Be 3
        $result.HitTotal | Should -Be $all
    }

    It "上限 0 は、溢れを作らず HitTotal は件数と同じ" {
        $result = searchContentIndex "多い" $contentIndexFiles $true 0 -workerCount 1
        $result.Truncated | Should -BeFalse
        $result.HitTotal | Should -Be (getHitRunCount $result.Runs)
    }

    It "進み具合に、今までの件数（HitTotal）が付く" {
        $script:totals = @()
        [void](searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1 -taskBytes 1 -onProgress { param($done, $total, $newHits, $hitTotal) $script:totals += $hitTotal })
        $script:totals.Count | Should -Be 6
        $script:totals[$script:totals.Count - 1] | Should -BeGreaterThan 3
        ($script:totals -join ",") | Should -Be ((@($script:totals | Sort-Object)) -join ",")
    }

    It "数えている途中で中止すると、HitTotal は -1 で Cancelled" {
        $stop = [hashtable]::Synchronized(@{ Stop = $false })
        $result = searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1 -taskBytes 1 -onProgress { param($done, $total, $newHits, $hitTotal) if ($done -ge 4) { $stop.Stop = $true } } -stopSignal $stop
        $result.Cancelled | Should -BeTrue
        $result.HitTotal | Should -Be -1
    }
}

Describe "溢れの書き溜め（照合の細部）" -Tag Io {
    It "Excel の図形でセル番地だけに一致した行は、表にも溢れにも入れない。長いテキストの行は表と同じに切る" {
        $tsvRoot = Join-Path $TestDrive "ov2_tsv"
        newTsv "$tsvRoot\a\図.xlsx\$(toIndexFileName "S[図形]")" @("F1`tx", "G1`tF1 あ", "F1`ty", "G1`tF1 い", "G1`tF1 う")
        newTsv "$tsvRoot\a\長い.txt\本文.tsv" @(("あ" * 300) + "検索語" + ("い" * 300), "検索語 短い")
        $files = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ov2_content_index")
        $tmp = newOverflowRoot "t-cell"
        foreach ($mode in "single", "parallel") {
            $dir = getSearchOverflowDir $tmp $(if ($mode -eq "single") { 1 } else { 2 })
            $workers = if ($mode -eq "single") { 1 } else { 2 }
            foreach ($word in "F1", "検索語") {
                $expected = toRows (expandHitRuns (searchContentIndex $word $files $true 0 -workerCount $workers -taskBytes 1).Runs)
                $result = searchContentIndex $word $files $true 1 -workerCount $workers -taskBytes 1 -overflowDir "$dir-$word" -checkEvery 1
                $rows = @(toRows (expandHitRuns $result.Runs)) + @(readSearchOverflowLines "$dir-$word")
                ($rows -join "`n") | Should -Be ($expected -join "`n")
                $result.HitTotal | Should -Be $expected.Count
            }
        }
    }

    It "正規表現で、ファイルの途中で数えるだけに切り替わっても、溢れの行が重ならず抜けない" {
        $tsvRoot = Join-Path $TestDrive "ov3_tsv"
        $mixed = @("aaab 1", "aaab 2", "aaab 3", "aaab 4")
        for ($i = 0; $i -lt 3; $i++) { newTsv "$tsvRoot\s$i\b$i.xlsx\S.tsv" $mixed }
        $files = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ov3_content_index")
        $expected = toRows (expandHitRuns (searchContentIndex "a+b" $files $false 0 -workerCount 1 -taskBytes 1).Runs)
        # 1 つ目のファイルの途中（2 件目）で数えるだけに切り替わる
        $tmp = newOverflowRoot "t-redo"
        $dir = getSearchOverflowDir $tmp 1
        $result = searchContentIndex "a+b" $files $false 5 -workerCount 1 -taskBytes 1 -checkEvery 1 -overflowDir $dir
        $rows = @(toRows (expandHitRuns $result.Runs)) + @(readSearchOverflowLines $dir)
        ($rows -join "`n") | Should -Be ($expected -join "`n")
    }

    It "上限を超えた後の照合が時間切れなら、CountTimeout にして表はそのまま（HitTotal は -1）" {
        $tsvRoot = Join-Path $TestDrive "ov4_tsv"
        newTsv "$tsvRoot\s0\b0.xlsx\S.tsv" @("a", "a", "a", "a")
        newTsv "$tsvRoot\s1\b1.xlsx\S.tsv" @(("a" * 40) + "!")
        $files = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ov4_content_index")
        $regexTimeout = [timespan]::FromMilliseconds(1)
        $result = searchContentIndex "(a+)+$" $files $false 2 -workerCount 1 -taskBytes 1
        $result.CountTimeout | Should -BeTrue
        $result.HitTotal | Should -Be -1
        (getHitRunCount $result.Runs) | Should -Be 2
        $result.Truncated | Should -BeTrue
    }

    It "上限の手前で時間切れなら、これまでどおり検索を中止する" {
        $tsvRoot = Join-Path $TestDrive "ov5_tsv"
        newTsv "$tsvRoot\s1\b1.xlsx\S.tsv" @(("a" * 40) + "!")
        $files = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ov5_content_index")
        $regexTimeout = [timespan]::FromMilliseconds(1)
        { searchContentIndex "(a+)+$" $files $false 5 -workerCount 1 } | Should -Throw -ExpectedMessage "*時間がかかりすぎる*"
    }

    It "合図（gate）があると、照合のスレッドのスクリプトは時間切れでも Hits を返し、時間切れの欄に 1 を書く" {
        $tsvRoot = Join-Path $TestDrive "ov6_tsv"
        newTsv "$tsvRoot\s0\b0.xlsx\S.tsv" @("a", "a")
        newTsv "$tsvRoot\s1\b1.xlsx\S.tsv" @(("a" * 40) + "!")
        $files = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ov6_content_index")
        $regexTimeout = [timespan]::FromMilliseconds(1)
        $search = newSearchRegex "(a+)+$" $false $false
        # 作業が 2 つある合図（[2]・[3] が件数の欄、[4]・[5] が時間切れの欄）の、番号 0 の作業として動かす
        $gate = [int[]]::new(2 + 2 * 2)
        $gate[1] = 10
        $output = & ${contentIndexWorkerScript} $files 0 $files.Count $search.Regex 10 $search.TextRegex $search.ScanMode $null $null $null $null $null $gate 0 64 ""
        $output.Timeout | Should -BeTrue
        @((expandHitRuns $output.Runs)).Count | Should -Be 2
        $gate[4] | Should -Be 1
    }

    It "全文への照合が時間切れになって 1 行ずつ照合し直しても、溢れの行が二重にならず抜けない（作業が 1 つ・2 つ）" -ForEach @(
        @{ Slots = 1 }
        @{ Slots = 2 }
    ) {
        $tsvRoot = Join-Path $TestDrive "ov8_tsv_$Slots"
        # 先頭の数行は速く当たり、その後ろの行で全文への照合だけが時間切れになる（1 行ずつなら速い）
        $rowsInFile = @(1..4 | ForEach-Object { "aaab $_" }) + @(1..8 | ForEach-Object { ("a" * 27) + "!" }) + @("aaab 5")
        for ($i = 0; $i -lt $Slots; $i++) { newTsv "$tsvRoot\s$i\b$i.xlsx\S.tsv" $rowsInFile }
        $files = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ov8_content_index_$Slots")
        $pattern = "^(a|aa)+b"
        $options = [System.Text.RegularExpressions.RegexOptions]::None
        $lineRegex = [regex]::new($pattern, $options, [timespan]::FromSeconds(5))
        $textRegex = [regex]::new($pattern, $options -bor [System.Text.RegularExpressions.RegexOptions]::Multiline, [timespan]::FromMilliseconds(1))
        # 手前の確かめ: 4 行の当たりの後ろから、全文への照合は時間切れになる
        $text = ($rowsInFile -join "`n")
        { [void]$textRegex.Match($text, 4 * ("aaab 1".Length + 1)) } | Should -Throw

        $tmp = newOverflowRoot "t-timeout-$Slots"
        $dir = getSearchOverflowDir $tmp 1
        $gate = [int[]]::new(2 + 2 * $Slots)
        $gate[1] = 1
        $expectedRows = @()
        $tableRows = @()
        for ($s = 0; $s -lt $Slots; $s++) {
            $expected = toRows (expandHitRuns (searchContentIndexFiles $files $s 1 $lineRegex -1 $textRegex "lines").Runs)
            $found = searchContentIndexFiles $files $s 1 $lineRegex 1 $textRegex "lines" $null $null $null $null $null $gate $s 1 $dir
            $found.Timeout | Should -BeFalse
            $found.OverflowFailed | Should -BeFalse
            $found.Count | Should -Be $expected.Count
            $own = toRows (expandHitRuns $found.Runs)
            # 表は上限 + 1 件（上限を超えたと分かる 1 件）まで。残りは溢れのファイルに入る
            # （2 つ目の作業は、前の作業の件数だけで上限を超えているので、表に行を作らず数えるだけ）
            $tableCount = if ($s -eq 0) { 2 } else { 0 }
            $own.Count | Should -Be $tableCount
            ($own -join "`n") | Should -Be (@($expected | Select-Object -First $tableCount) -join "`n")
            $tableRows += $own
            $expectedRows += @($expected | Select-Object -Skip $tableCount)
        }
        $overflow = @(readSearchOverflowLines $dir)
        $overflow.Count | Should -Be $expectedRows.Count
        ($overflow -join "`n") | Should -Be ($expectedRows -join "`n")
    }

    It "合図の止める印が立っていれば、数えている最中のスクリプトも止まる" {
        $tsvRoot = Join-Path $TestDrive "ov7_tsv"
        newTsv "$tsvRoot\s1\b1.xlsx\S.tsv" (1..3000 | ForEach-Object { "ヒット $_" })
        $files = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ov7_content_index")
        $search = newSearchRegex "ヒット" $true $false
        $gate = [int[]]::new(2 + 2 * 1)
        $gate[1] = 1
        $gate[0] = 1
        $output = & ${contentIndexWorkerScript} $files 0 $files.Count $search.Regex 1 $search.TextRegex $search.ScanMode $null $null $null $null $null $gate 0 64 ""
        $output.Count | Should -BeLessThan 3000
    }
}

Describe "溢れの置き場所と後始末" -Tag Io {
    It "置き場所は TmpRoot の下の、PC の鍵・search_overflow・PID・要求の番号の順のフォルダ" {
        $tmp = "C:\t"
        getSearchOverflowDir $tmp 7 | Should -Be (Join-Path (Join-Path (Join-Path (Join-Path $tmp (getMachineKey)) "search_overflow") $PID) 7)
    }

    It "根のリンク・相対パス・根でないものには、作らず消さない" {
        $tmp = newOverflowRoot "t-guard"
        $outside = Join-Path $TestDrive "outside"
        [void][System.IO.Directory]::CreateDirectory($outside)
        Set-Content -LiteralPath "$outside\keep.txt" -Value "x"
        testSearchOverflowRoot "relative\search_overflow" $tmp | Should -BeFalse
        testSearchOverflowRoot (Join-Path $TestDrive "other") $tmp | Should -BeFalse
        testSearchOverflowRoot (getSearchOverflowRoot $tmp) $tmp | Should -BeTrue
        # 根がリンクなら、辿って消さない
        [void][System.IO.Directory]::CreateDirectory((Join-Path $tmp (getMachineKey)))
        $root = getSearchOverflowRoot $tmp
        New-Item -ItemType Junction -Path $root -Target $outside | Out-Null
        testSearchOverflowRoot $root $tmp | Should -BeFalse
        removeSearchOverflow (getSearchOverflowPidDir $tmp) $tmp | Should -BeFalse
        Test-Path -LiteralPath "$outside\keep.txt" | Should -BeTrue
        [System.IO.Directory]::Delete($root)
    }

    It "PID のフォルダを中ごと消す。無ければ真" {
        $tmp = newOverflowRoot "t-remove"
        $dir = getSearchOverflowDir $tmp 3
        [void][System.IO.Directory]::CreateDirectory($dir)
        Set-Content -LiteralPath "$dir\0.tsv" -Value "x"
        removeSearchOverflow (getSearchOverflowPidDir $tmp) $tmp | Should -BeTrue
        [System.IO.Directory]::Exists((getSearchOverflowPidDir $tmp)) | Should -BeFalse
        removeSearchOverflow (getSearchOverflowPidDir $tmp) $tmp | Should -BeTrue
    }

    It "起動時の片付けは、動いていないプロセスの PID のフォルダだけを消し、今のプロセス・動いているプロセス・数字でない名前・ほかの PC の鍵は残す" {
        $tmp = newOverflowRoot "t-stale"
        $root = getSearchOverflowRoot $tmp
        foreach ($name in "$PID", "111111", "222222", "memo") {
            [void][System.IO.Directory]::CreateDirectory((Join-Path $root "$name\1"))
        }
        $other = Join-Path (Join-Path (Join-Path $tmp "OTHERPC") "search_overflow") "333333"
        [void][System.IO.Directory]::CreateDirectory($other)
        $removed = removeStaleSearchOverflow $tmp { param ($processId) $processId -eq 222222 }
        $removed | Should -Be 1
        Test-Path -LiteralPath (Join-Path $root "$PID") | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $root "222222") | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $root "111111") | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $root "memo") | Should -BeTrue
        Test-Path -LiteralPath $other | Should -BeTrue
    }

    It "根が無ければ、片付けは何もせず 0" {
        removeStaleSearchOverflow (newOverflowRoot "t-noroot") | Should -Be 0
    }

    It "溢れの読み手は、行の中の CR で行を分けない" {
        $dir = Join-Path $TestDrive "reader"
        [void][System.IO.Directory]::CreateDirectory($dir)
        $writer = newSearchOverflowWriter $dir "0.tsv"
        $writer.WriteLine("a`tb`tc`t1`tx`ry")
        $writer.WriteLine("a`tb`tc`t2`tz")
        $writer.Dispose()
        $lines = @(readSearchOverflowLines $dir)
        $lines.Count | Should -Be 2
        $lines[0] | Should -Be "a`tb`tc`t1`tx`ry"
    }

    It "作業の番号の順に、head → 作業のファイルの順に並べる" {
        $dir = Join-Path $TestDrive "order"
        [void][System.IO.Directory]::CreateDirectory($dir)
        foreach ($name in "10.tsv", "2.tsv", "2.head.tsv", "1.tsv", "memo.txt") { Set-Content -LiteralPath (Join-Path $dir $name) -Value "x" }
        (@(getSearchOverflowFiles $dir) | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join "," | Should -Be "1.tsv,2.head.tsv,2.tsv,10.tsv"
    }
}

Describe "invokeSearchRequest（溢れの置き場所・件数）" -Tag Io {
    BeforeAll {
        $wsDir = Join-Path $TestDrive "ws_run"
        $tsvRoot = Join-Path $TestDrive "run_tsv"
        for ($i = 0; $i -lt 4; $i++) { newTsv "$tsvRoot\d\b$i.xlsx\S.tsv" @("多い $i", "多い 多い") }
        $contentRoot = Join-Path $TestDrive "run_content_index"
        $null = newContentIndexFiles $tsvRoot $contentRoot
    }

    It "上限を超えると HitTotal・Truncated を要求に入れ、溢れを書き、次の検索で前の溢れを消す" {
        $own = [Workspace]::new($wsDir)
        $request = newSearchRequest "多い" $true @($contentRoot) 3 @{} $wsDir
        invokeSearchRequest $request $null $null
        $request.Error | Should -BeNullOrEmpty
        $request.Truncated | Should -BeTrue
        $request.HitTotal | Should -Be 8
        $request.OverflowDir | Should -Be (getSearchOverflowDir $own.TmpRoot $request.Number)
        @(readSearchOverflowLines $request.OverflowDir).Count | Should -Be 5
        $first = $request.OverflowDir

        $next = newSearchRequest "多い" $true @($contentRoot) 3 @{} $wsDir
        invokeSearchRequest $next $null $null
        Test-Path -LiteralPath $first | Should -BeFalse
        @(readSearchOverflowLines $next.OverflowDir).Count | Should -Be 5
    }

    It "上限なしなら、溢れの置き場所を決めない" {
        $request = newSearchRequest "多い" $true @($contentRoot) 0 @{} $wsDir
        invokeSearchRequest $request $null $null
        $request.OverflowDir | Should -Be ""
        $request.Truncated | Should -BeFalse
        $request.HitTotal | Should -Be 8
    }
}

Describe "splitContentIndexTasks（先頭の作業を小さくする）" -Tag Unit {
    It "<name>" -ForEach @(
        @{ name = "headBytes なしは、今までどおり taskBytes ずつ"; head = 0; sizes = @(10, 10, 10, 10, 10); expected = @(3, 2) }
        @{ name = "headBytes があると、先頭の作業を小さくし、作業ごとに 2 倍にして taskBytes まで大きくする"; head = 5; sizes = @(5, 10, 20, 20, 20, 20, 20); expected = @(1, 1, 1, 2, 2) }
        @{ name = "headBytes が taskBytes 以上なら使わない"; head = 30; sizes = @(10, 10, 10, 10, 10); expected = @(3, 2) }
    ) {
        $files = @($sizes | ForEach-Object { [pscustomobject]@{ Size = $_ } })
        $tasks = splitContentIndexTasks $files 30 $head
        @($tasks | ForEach-Object { $_.Count }) | Should -Be $expected
        # 先頭から隙間なく並ぶ
        $next = 0
        foreach ($task in $tasks) { $task.Start | Should -Be $next; $next += $task.Count }
        $next | Should -Be $files.Count
    }
}

Describe "resetSearchOverflowWriter" -Tag Io {
    It "書いた行を、指定した位置まで切り戻し、続きをそこから書ける" {
        $dir = Join-Path $TestDrive "reset_ok"
        $writer = newSearchOverflowWriter $dir "0.tsv"
        $writer.WriteLine("1行目")
        $writer.Flush()
        $pos = $writer.BaseStream.Position
        $writer.WriteLine("消える行")
        (resetSearchOverflowWriter $writer $pos) | Should -BeTrue
        $writer.WriteLine("2行目")
        $writer.Dispose()
        @(readSearchOverflowLines $dir) -join "," | Should -Be "1行目,2行目"
    }

    It "書けない（閉じた）writer は、閉じ直して偽を返す" {
        $dir = Join-Path $TestDrive "reset_ng"
        $writer = newSearchOverflowWriter $dir "0.tsv"
        $writer.Dispose()
        (resetSearchOverflowWriter $writer 0) | Should -BeFalse
    }
}

Describe "writeSearchOverflowHits" -Tag Io {
    It "まとまりの中の行を、1 行ずつ <RelDir>\t<Book>\t<Location>\t<行番号>\t<行> の形で書く" {
        $dir = Join-Path $TestDrive "write_runs"
        $run = @{ RelDir = "営業"; Book = "見積.xlsx"; Location = "4月"
            LineNumbers = [System.Collections.Generic.List[int]]::new(); Lines = [System.Collections.Generic.List[string]]::new() }
        foreach ($n in 2, 5) { $run.LineNumbers.Add($n); $run.Lines.Add("行$n") }
        $writer = newSearchOverflowWriter $dir "0.head.tsv"
        (writeSearchOverflowHits $writer @($run, $run)) | Should -BeTrue
        $writer.Dispose()
        @(readSearchOverflowLines $dir) -join "|" | Should -Be ((@("営業`t見積.xlsx`t4月`t2`t行2", "営業`t見積.xlsx`t4月`t5`t行5") * 2) -join "|")
    }
}
