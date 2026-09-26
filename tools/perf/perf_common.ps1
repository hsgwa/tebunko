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
            throw "測る tebunko に $name がありません。pack 形式より前の版は測れません（関数の引数が変わった場合は、この計測スクリプトを直してください）。"
        }
    }
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
