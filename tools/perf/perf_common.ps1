# 性能の計測（tools\measure_perf.ps1）で使う共通の部品。計測の入口と、子のプロセス（measure_index.ps1・measure_search.ps1）が読み込む。
#   ・測る tebunko の読み込み口を探す
#   ・リソース（メモリ・CPU・スレッド・ハンドル・GC）を別のスレッドで記録する
#   ・統計値（最小・中央値・平均・最大）と増え方（最小二乗の傾き）を求める
#   ・Mermaid のグラフの行を作る

$perfInvariant = [System.Globalization.CultureInfo]::InvariantCulture

function resolveTebunkoLib {
    # 測る tebunko の読み込み口（lib.ps1）のパスを返す。
    # フォルダの名前を変える前（#99 より前）の版は scripts\tebunko_grep\lib.ps1
    param (
        [string]$tool
    )

    foreach ($rel in @("scripts\tebunko\lib.ps1", "scripts\tebunko_grep\lib.ps1")) {
        $path = Join-Path $tool $rel
        if (Test-Path -LiteralPath $path) { return (Resolve-Path -LiteralPath $path).ProviderPath }
    }
    throw "測る tebunko に scripts\tebunko\lib.ps1 がありません。-Tool に tebunko のフォルダを指定してください。"
}

function assertTebunkoFunctions {
    # 計測で呼ぶ関数が、読み込んだ tebunko にあるかを確かめる
    param (
        [string[]]$names
    )

    foreach ($name in $names) {
        if (!(Get-Command $name -CommandType Function -ErrorAction SilentlyContinue)) {
            throw (getMissingTebunkoFunctionMessage $name)
        }
    }
}

function getMissingTebunkoFunctionMessage {
    param (
        [string]$name
    )

    return "測る tebunko に $name がありません。pack 形式より前の版は測れません（関数の引数が変わった場合は、この計測スクリプトを直してください）。"
}

function resolveTebunkoFunction {
    # 候補の名前の並び（新しい名前から先）から、読み込んだ tebunko にある最初の関数の名前を返す。
    # 名前を改めた前の版（古い名前だけある）も、改めた後の版も測れるようにする。どれも無ければ assertTebunkoFunctions と同じ文言で止める
    param (
        [string[]]$names
    )

    foreach ($name in $names) {
        if (Get-Command $name -CommandType Function -ErrorAction SilentlyContinue) { return $name }
    }
    throw (getMissingTebunkoFunctionMessage $names[0])
}

function getTebunkoProperty {
    # 候補の名前の並び（新しい名前から先）から、ハッシュテーブルにある最初の項目の値を返す（getIndexPackFiles の戻り値の項目の新旧）
    param (
        [hashtable]$table,
        [string[]]$names
    )

    foreach ($name in $names) {
        if ($table.ContainsKey($name)) { return $table[$name] }
    }
    return $null
}

# リソースを 1 回記録する。記録用のスレッドと、段階を切り替える呼び出し元の両方で使う
$perfTakeSample = {
    param ($monitor, $cpuCounter, $memCounter)
    $p = [System.Diagnostics.Process]::GetCurrentProcess()
    $pcCpu = $null; $pcFree = $null
    if ($cpuCounter) { try { $pcCpu = [Math]::Round($cpuCounter.NextValue(), 1) } catch { } }
    if ($memCounter) { try { $pcFree = [int]$memCounter.NextValue() } catch { } }
    [void]$monitor.Samples.Add([pscustomobject]@{
        Ms = [int]$monitor.Clock.Elapsed.TotalMilliseconds; Phase = $monitor.Phase
        WorkingSetMB = [Math]::Round($p.WorkingSet64 / 1MB, 1); PrivateMB = [Math]::Round($p.PrivateMemorySize64 / 1MB, 1)
        ManagedMB = [Math]::Round([GC]::GetTotalMemory($false) / 1MB, 1); CpuMs = [int]$p.TotalProcessorTime.TotalMilliseconds
        Threads = $p.Threads.Count; Handles = $p.HandleCount
        Gc0 = [GC]::CollectionCount(0); Gc1 = [GC]::CollectionCount(1); Gc2 = [GC]::CollectionCount(2)
        PcCpuPercent = $pcCpu; PcFreeMB = $pcFree
    })
    $p.Dispose()
}

function startResourceMonitor {
    # 別のスレッドで sampleMs ごとにリソースを記録し始める。段階は setMonitorPhase で切り替える。
    # phaseChannel（インデクサの受け渡しの口）を渡すと、記録のスレッドが Progress.Phase を短い間隔で読み、
    # 段階が変わるたびにすぐ記録して、その段階を記録の段階にする（段階の切り替えは setMonitorPhase の代わりにここで行う）
    param (
        [int]$sampleMs = 200,
        $phaseChannel = $null
    )

    $monitor = [hashtable]::Synchronized(@{
        Stop = $false; Phase = "準備"; Clock = [System.Diagnostics.Stopwatch]::StartNew()
        Samples = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        Phases = New-Object System.Collections.Generic.List[hashtable]; PhaseChannel = $phaseChannel
    })
    $sampler = {
        param ($monitor, $sampleMs, $takeText, $phaseChannel)
        $take = [scriptblock]::Create($takeText)
        $cpuCounter = $null; $memCounter = $null
        try { $cpuCounter = New-Object System.Diagnostics.PerformanceCounter "Processor", "% Processor Time", "_Total"; [void]$cpuCounter.NextValue() } catch { $cpuCounter = $null }
        try { $memCounter = New-Object System.Diagnostics.PerformanceCounter "Memory", "Available MBytes" } catch { $memCounter = $null }
        if ($null -eq $phaseChannel) {
            while (!$monitor.Stop) {
                & $take $monitor $cpuCounter $memCounter
                Start-Sleep -Milliseconds $sampleMs
            }
        } else {
            $sinceSample = [System.Diagnostics.Stopwatch]::StartNew()
            & $take $monitor $cpuCounter $memCounter
            while (!$monitor.Stop) {
                $changed = $false
                $progress = $phaseChannel.Progress
                if ($null -ne $progress -and $progress.Phase -and $progress.Phase -ne $monitor.Phase) {
                    $now = [int]$monitor.Clock.Elapsed.TotalMilliseconds
                    if ($monitor.Phases.Count -gt 0) { $monitor.Phases[$monitor.Phases.Count - 1].EndMs = $now }
                    $monitor.Phases.Add(@{ Name = [string]$progress.Phase; StartMs = $now; EndMs = $null })
                    $monitor.Phase = [string]$progress.Phase
                    $changed = $true
                }
                if ($changed -or $sinceSample.ElapsedMilliseconds -ge $sampleMs) {
                    & $take $monitor $cpuCounter $memCounter
                    $sinceSample.Restart()
                }
                Start-Sleep -Milliseconds 10
            }
        }
    }
    $ps = [powershell]::Create()
    [void]$ps.AddScript($sampler).AddArgument($monitor).AddArgument($sampleMs).AddArgument($perfTakeSample.ToString()).AddArgument($phaseChannel)
    $monitor.PowerShell = $ps
    $monitor.Handle = $ps.BeginInvoke()
    return $monitor
}

function setMonitorPhase {
    # 段階を切り替える。切り替えの前後でも記録し、短い段階にも開始時と終了時の値が残るようにする。name が空なら段階を閉じるだけ
    param (
        $monitor,
        [string]$name
    )

    $now = [int]$monitor.Clock.Elapsed.TotalMilliseconds
    if ($monitor.Phases.Count -gt 0) {
        & $perfTakeSample $monitor $null $null
        $monitor.Phases[$monitor.Phases.Count - 1].EndMs = $now
    }
    $monitor.Phase = $name
    if ($name) {
        $monitor.Phases.Add(@{ Name = $name; StartMs = $now; EndMs = $null })
        & $perfTakeSample $monitor $null $null
    }
}

function stopResourceMonitor {
    # 記録を止め、記録した値の配列を返す
    param (
        $monitor
    )

    if ($null -eq $monitor.PhaseChannel) {
        setMonitorPhase $monitor ""
        $monitor.Stop = $true
    } else {
        # 記録のスレッドが段階を切り替えるとき（phaseChannel あり）は、閉じた段階をまた開かないよう、先に止めてから閉じる
        $monitor.Stop = $true
        [void]$monitor.PowerShell.EndInvoke($monitor.Handle)
        setMonitorPhase $monitor ""
        $monitor.PowerShell.Dispose()
        return , @($monitor.Samples.ToArray())
    }
    [void]$monitor.PowerShell.EndInvoke($monitor.Handle)
    $monitor.PowerShell.Dispose()
    return , @($monitor.Samples.ToArray())
}

function getProcessSnapshot {
    # このプロセスの今のメモリ・ハンドル・スレッドを返す（検索 1 回ごとの記録）
    $p = [System.Diagnostics.Process]::GetCurrentProcess()
    try {
        return [ordered]@{
            WorkingSetMB = [Math]::Round($p.WorkingSet64 / 1MB, 1); PrivateMB = [Math]::Round($p.PrivateMemorySize64 / 1MB, 1)
            ManagedMB = [Math]::Round([GC]::GetTotalMemory($false) / 1MB, 1); Handles = $p.HandleCount; Threads = $p.Threads.Count
        }
    } finally {
        $p.Dispose()
    }
}

function getStats {
    # 最小・中央値・平均・最大・件数を返す。値が無ければ $null
    param (
        [double[]]$values
    )

    if ($null -eq $values -or $values.Count -eq 0) { return $null }
    $sorted = [double[]]$values.Clone()
    [Array]::Sort($sorted)
    $n = $sorted.Count
    $mid = [int][Math]::Floor($n / 2)
    $median = if ($n % 2) { $sorted[$mid] } else { ($sorted[$mid - 1] + $sorted[$mid]) / 2 }
    $sum = 0.0
    foreach ($v in $sorted) { $sum += $v }
    return [ordered]@{ Min = $sorted[0]; Median = $median; Mean = $sum / $n; Max = $sorted[$n - 1]; Count = $n }
}

function getSlope {
    # 最小二乗法で求めた直線の傾き（y の x あたりの増え方）を返す。2 点に満たないときは $null
    param (
        [double[]]$x,
        [double[]]$y
    )

    $n = $x.Count
    if ($n -lt 2 -or $y.Count -ne $n) { return $null }
    $mx = 0.0; $my = 0.0
    for ($i = 0; $i -lt $n; $i++) { $mx += $x[$i]; $my += $y[$i] }
    $mx /= $n; $my /= $n
    $sxy = 0.0; $sxx = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $sxy += ($x[$i] - $mx) * ($y[$i] - $my)
        $sxx += ($x[$i] - $mx) * ($x[$i] - $mx)
    }
    if ($sxx -eq 0) { return $null }
    return $sxy / $sxx
}

function getPhaseResources {
    # 記録した値を段階ごとに集計する（時間・CPU 時間と使用率・メモリの最大・スレッドとハンドルの最大・GC の回数）
    param (
        [object[]]$samples,
        [int]$cores = [Environment]::ProcessorCount
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($name in @($samples | ForEach-Object { $_.Phase } | Where-Object { $_ } | Select-Object -Unique)) {
        $set = @($samples | Where-Object { $_.Phase -eq $name })
        $first = $set[0]; $last = $set[$set.Count - 1]
        $wallMs = [Math]::Max(1, $last.Ms - $first.Ms)
        $pcCpu = @($set | Where-Object { $null -ne $_.PcCpuPercent } | ForEach-Object { $_.PcCpuPercent })
        $rows.Add([ordered]@{
            Phase = $name; Seconds = [Math]::Round($wallMs / 1000, 1)
            CpuSeconds = [Math]::Round(($last.CpuMs - $first.CpuMs) / 1000, 1)
            CpuPercent = [Math]::Round(($last.CpuMs - $first.CpuMs) * 100 / $wallMs / $cores, 0)
            PcCpuPercent = $(if ($pcCpu.Count) { [Math]::Round(($pcCpu | Measure-Object -Average).Average, 0) } else { $null })
            WorkingSetMaxMB = ($set | Measure-Object WorkingSetMB -Maximum).Maximum
            PrivateMaxMB = ($set | Measure-Object PrivateMB -Maximum).Maximum
            ManagedMaxMB = ($set | Measure-Object ManagedMB -Maximum).Maximum
            ThreadsMax = ($set | Measure-Object Threads -Maximum).Maximum
            HandlesMax = ($set | Measure-Object Handles -Maximum).Maximum
            Gc0 = $last.Gc0 - $first.Gc0; Gc1 = $last.Gc1 - $first.Gc1; Gc2 = $last.Gc2 - $first.Gc2
        })
    }
    return , $rows.ToArray()
}

function formatPerfNumber {
    # グラフ・CSV に書く数（カルチャに左右されない書き方）
    param (
        [double]$value
    )

    return $value.ToString("0.###", $perfInvariant)
}

function newLineChart {
    # Mermaid の折れ線グラフ（xychart-beta）の行を返す。
    #   xAxis : x 軸の書き方（"0 --> 80" のような範囲、または "[1, 2, 3]" のような並び）
    #   lines : 線ごとの値の配列の配列
    #   colors: 線の色（"#rrggbb"）。凡例は xychart に無いので、色と名前の対応は呼び出し元が文で書く
    param (
        [string]$title,
        [string]$xLabel,
        [string]$xAxis,
        [string]$yLabel,
        [object[]]$lines,
        [string[]]$colors = @()
    )

    $out = New-Object System.Collections.Generic.List[string]
    $out.Add('```mermaid')
    if ($colors.Count) {
        $out.Add("%%{init: {`"themeVariables`": {`"xyChart`": {`"plotColorPalette`": `"$($colors -join ', ')`"}}}}%%")
    }
    $out.Add("xychart-beta")
    $out.Add("    title `"$title`"")
    $out.Add("    x-axis `"$xLabel`" $xAxis")
    $out.Add("    y-axis `"$yLabel`"")
    foreach ($line in $lines) {
        $out.Add("    line [" + ((@($line) | ForEach-Object { formatPerfNumber ([double]$_) }) -join ", ") + "]")
    }
    $out.Add('```')
    return , $out.ToArray()
}

function thinOut {
    # 点が多すぎるとグラフが読めないため、おおよそ max 個に間引く（最後の点は残す）
    param (
        [object[]]$items,
        [int]$max = 100
    )

    if ($items.Count -le $max) { return , $items }
    $step = [int][Math]::Ceiling($items.Count / $max)
    $list = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $items.Count; $i += $step) { $list.Add($items[$i]) }
    if ($list[$list.Count - 1] -ne $items[$items.Count - 1]) { $list.Add($items[$items.Count - 1]) }
    return , $list.ToArray()
}

function getSearchPerfProblems {
    # 検索と pack の作成の回帰テストの判定。合わないものの一覧（文字列の配列）を返す。空なら合格。
    #   result      : measure_perf.ps1 の result.json（ConvertFrom-Json したもの）
    #   rows        : searches.csv の行（Word・Hits・Truncated・Packs）。1 回ごとの値
    #   limits      : 語の名前 → 検索の中央値の上限（ms）
    #   expected    : 語の名前 → 期待する件数（数、または "10000+" = 1 万件で打ち切り）
    #   packLimitSeconds : pack の作成の秒の上限
    # 失敗のメッセージには、対象の名前・値・上限の数字だけを書く
    param (
        $result,
        [object[]]$rows,
        [hashtable]$limits,
        [hashtable]$expected,
        [double]$packLimitSeconds
    )

    $problems = New-Object System.Collections.Generic.List[string]
    $mode = [string]$result.Run.SearchMode
    if ($mode -ne "service") { $problems.Add("検索の流れが service ではありません（$mode）") }

    $pack = $result.Index.Pack
    $packCount = 0
    if ($null -eq $pack -or [int]$pack.Packs -le 0) {
        $problems.Add("pack がありません")
    } else {
        $packCount = [int]$pack.Packs
        if ([double]$pack.Seconds -gt $packLimitSeconds) {
            $problems.Add("pack の作成: $($pack.Seconds) 秒（上限 $packLimitSeconds 秒）")
        }
    }

    foreach ($name in @($limits.Keys | Sort-Object)) {
        $entry = @($result.Search | Where-Object { $_.Name -eq $name }) | Select-Object -First 1
        if ($null -eq $entry -or $null -eq $entry.TotalMs) {
            $problems.Add("検索 ${name}: 結果がありません")
            continue
        }
        if ([double]$entry.TotalMs.Median -gt [double]$limits[$name]) {
            $problems.Add([string]::Format([CultureInfo]::InvariantCulture, "検索 {0}: 中央値 {1:N0} ms（上限 {2:N0} ms）", $name, [double]$entry.TotalMs.Median, [double]$limits[$name]))
        }

        $mine = @($rows | Where-Object { $_.Word -eq $name })
        if ($mine.Count -eq 0) {
            $problems.Add("検索 ${name}: 1 回ごとの記録がありません")
            continue
        }
        $want = [string]$expected[$name]
        if ([string]::IsNullOrWhiteSpace($want)) {
            $problems.Add("検索 ${name}: 期待する件数がありません")
            continue
        }
        $bad = @($mine | Where-Object {
            $hits = [int]$_.Hits
            $truncated = ("$($_.Truncated)" -eq "True")
            if ($want -eq "10000+") { !($hits -eq 10000 -and $truncated) } else { $hits -ne [int]$want }
        })
        if ($bad.Count) {
            $problems.Add("検索 ${name}: 件数が期待の $want と違う回が $($bad.Count) 回（最初は $($bad[0].Hits) 件）")
        }
        $badPacks = @($mine | Where-Object { [int]$_.Packs -le 0 -or [int]$_.Packs -ne $packCount })
        if ($badPacks.Count) {
            $problems.Add("検索 ${name}: 照合した pack の数が $packCount と違う回が $($badPacks.Count) 回（最初は $($badPacks[0].Packs)）")
        }
    }
    return , $problems.ToArray()
}

function getIngestPerfProblems {
    # 取り込み（.docx・.pptx）の回帰テストの判定。合わないものの一覧（文字列の配列）を返す。空なら合格。
    #   ingest       : result.json の Ingest（無ければ $null）
    #   aggregate    : 最後の回のワークスペースにできた集約ファイル（content_index.*.tsv）の @{ Count; Bytes }
    #   perFileLimitMs・secondsLimit : 1 ファイルあたりの中央値（ms）・全体の中央値（秒）の上限
    #   expectedFiles: 取り込むファイルの数
    param (
        $ingest,
        $aggregate,
        [double]$perFileLimitMs,
        [double]$secondsLimit,
        [int]$expectedFiles
    )

    $problems = New-Object System.Collections.Generic.List[string]
    if ($null -eq $ingest) {
        $problems.Add("取り込みの結果がありません")
        return , $problems.ToArray()
    }
    if ([int]$ingest.Total -ne $expectedFiles) { $problems.Add("取り込んだファイルが $($ingest.Total) 件（期待 $expectedFiles 件）") }
    if ([int]$ingest.Done -ne $expectedFiles) { $problems.Add("取り込みが済んだのは $($ingest.Done) 件（期待 $expectedFiles 件）") }
    if ([int]$ingest.Failed -ne 0) { $problems.Add("取り込みに失敗したのは $($ingest.Failed) 件") }
    $perFile = if ($null -ne $ingest.PerFileMs) { $ingest.PerFileMs.Median } else { $null }
    if ($null -eq $perFile) {
        $problems.Add("1 ファイルあたりの時間がありません")
    } elseif ([double]$perFile -gt $perFileLimitMs) {
        $problems.Add([string]::Format([CultureInfo]::InvariantCulture, "取り込み 1 ファイルあたり: 中央値 {0:N0} ms（上限 {1:N0} ms）", [double]$perFile, $perFileLimitMs))
    }
    $seconds = if ($null -ne $ingest.Seconds) { $ingest.Seconds.Median } else { $null }
    if ($null -eq $seconds) {
        $problems.Add("全体の時間がありません")
    } elseif ([double]$seconds -gt $secondsLimit) {
        $problems.Add([string]::Format([CultureInfo]::InvariantCulture, "取り込み全体: 中央値 {0:N1} 秒（上限 {1:N0} 秒）", [double]$seconds, $secondsLimit))
    }
    if ($null -eq $aggregate -or [int]$aggregate.Count -le 0) {
        $problems.Add("集約ファイルがありません")
    } elseif ([long]$aggregate.Bytes -le 0) {
        $problems.Add("集約ファイルの合計の大きさが 0 です")
    }
    return , $problems.ToArray()
}

function getTypingSteps {
    # 入力しながらの検索で送る語の並びを作る（画面の入力しながらの検索と同じ、3 文字目から 1 文字ずつ伸ばす）。
    # tools\perf\measure_typing.ps1 が使う。
    #   word     : 検索する語
    #   isRegex  : 正規表現として送るか（words.tsv の 正規表現 列）
    #   minLength: 最初に送る語の長さ。語の長さがこれより短ければ、語そのものだけを返す
    # 文字数は String.Length（UTF-16 のコード単位）で数え、サロゲートペアの途中で切れる長さは飛ばす。
    # 正規表現のときは、正規表現として不正な文字列と、空の文字列に一致する文字列（画面が自動では検索しないもの）も飛ばす。
    # 送れる長さが 1 つも無ければ、空の配列を返す（呼び出し側は Skipped として扱う）
    param (
        [string]$word,
        [bool]$isRegex,
        [int]$minLength = 3
    )

    $len = $word.Length
    $lengths = if ($len -le $minLength) { @($len) } else { @($minLength..$len) }
    $steps = New-Object System.Collections.Generic.List[string]
    foreach ($length in $lengths) {
        if ($length -lt 1) { continue }
        if ($length -lt $len -and [char]::IsHighSurrogate($word[$length - 1])) {
            continue  # サロゲートペアの途中で切れる長さは飛ばす
        }
        $step = $word.Substring(0, $length)
        if ($isRegex) {
            $re = $null
            try { $re = [regex]::new($step) } catch { continue }
            if ($re.Match("").Success) { continue }  # 空の文字列に一致する語は、画面が自動では検索しない
        }
        $steps.Add($step)
    }
    return $steps.ToArray()
}

function newTypingWordResult {
    # 入力しながらの検索を、語ごとに 1 回ごとの記録（measure_typing.ps1 の Attempts）から result.json の Typing.Words の 1 件にまとめる。
    #   attempts: 1 回目が先頭の配列。1 回ごとに @{ FirstHitMs; FinishMs; WaitPrevMs（すべて 1 回目は含めても除いても呼び出し側が扱う）；
    #             IntervalMs; LagMs; RequestMs（要求ごとの値の配列）; Hits; Truncated; Packs; FastUsed; CancelledSteps }
    #   steps   : 送った語の数（getTypingSteps の結果の数）。並びが空だった語は 0
    #   skipped : 測らなかった理由（並びが空の語）。$null なら測った
    # 統計（Min・Median・Mean・Max・Count）は getStats を使う。FirstHitMs・FinishMs・WaitPrevMs は 2 回目以降（attempts の 2 件目から）、
    # IntervalMs・LagMs・RequestMs は全回（1 回目を含む）から作る。CancelledSteps は 2 回目以降の合計、FastUsed は全回の合計（回数）。
    param (
        [string]$name,
        [bool]$regex,
        [int]$steps,
        [object[]]$attempts = @(),
        [string]$skipped = $null,
        $workingSetMaxMB = $null,
        $cpuSeconds = $null,
        $gc0 = $null,
        $gc1 = $null,
        $gc2 = $null,
        $peakWorkingSetMB = $null
    )

    if ($skipped -or $attempts.Count -eq 0) {
        return [ordered]@{
            Name = $name; Regex = $regex; Steps = $steps
            Hits = $null; Truncated = $null; Packs = $null; FastUsed = $null; CancelledSteps = $null
            FirstHitMs = $null; FinishMs = $null; WaitPrevMs = $null; IntervalMs = $null; LagMs = $null; RequestMs = $null
            First = $null
            WorkingSetMaxMB = $null; CpuSeconds = $null; Gc0 = $null; Gc1 = $null; Gc2 = $null
            PeakWorkingSetMB = $peakWorkingSetMB
            Skipped = $(if ($skipped) { $skipped } else { "送る語がありません（すべて正規表現として不正、または空の文字列に一致する）" })
        }
    }

    function flattenTypingValues($rows, [string]$key) {
        $list = New-Object System.Collections.Generic.List[double]
        foreach ($row in $rows) {
            foreach ($x in @($row.$key)) {
                if ($null -ne $x) { [void]$list.Add([double]$x) }
            }
        }
        return , $list.ToArray()
    }

    $first = $attempts[0]
    $rest = @($attempts | Select-Object -Skip 1)
    $last = $attempts[$attempts.Count - 1]

    return [ordered]@{
        Name = $name; Regex = $regex; Steps = $steps
        Hits = [int]$last.Hits; Truncated = [bool]$last.Truncated; Packs = [int]$last.Packs
        FastUsed = @($attempts | Where-Object { $_.FastUsed }).Count
        CancelledSteps = [int]((@($rest | ForEach-Object { [int]$_.CancelledSteps }) | Measure-Object -Sum).Sum)
        FirstHitMs = getStats (flattenTypingValues $rest "FirstHitMs")
        FinishMs = getStats (flattenTypingValues $rest "FinishMs")
        WaitPrevMs = getStats (flattenTypingValues $rest "WaitPrevMs")
        IntervalMs = getStats (flattenTypingValues $attempts "IntervalMs")
        LagMs = getStats (flattenTypingValues $attempts "LagMs")
        RequestMs = getStats (flattenTypingValues $attempts "RequestMs")
        First = [ordered]@{ FirstHitMs = $first.FirstHitMs; FinishMs = $first.FinishMs; WaitPrevMs = $first.WaitPrevMs }
        WorkingSetMaxMB = $workingSetMaxMB
        CpuSeconds = $cpuSeconds; Gc0 = $gc0; Gc1 = $gc1; Gc2 = $gc2
        PeakWorkingSetMB = $peakWorkingSetMB
        Skipped = $null
    }
}
