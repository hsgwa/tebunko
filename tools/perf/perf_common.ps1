# 性能の計測（tools\measure_perf.ps1）で使う共通の部品。計測の入口と、子のプロセス（measure_index.ps1・measure_search.ps1・measure_preview.ps1 など）が読み込む。
#   ・測る tebunko の読み込み口を探す
#   ・リソース（メモリ・CPU・スレッド・ハンドル・GC）を別のスレッドで記録する
#   ・統計値（最小・中央値・平均・最大）と増え方（最小二乗の傾き）を求める
#   ・Mermaid のグラフの行を作る

$perfInvariant = [System.Globalization.CultureInfo]::InvariantCulture
# プレビューの読み足しの計測で、1 回に読む行数（measure_preview.ps1 が頼む数。newReadMoreResult は、この数だけ読めた回を完全な 1 回と数える）
$readMoreChunkLines = 200

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

function getToolIndexDir {
    # 測る tebunko の、ワークスペース（work）の本文インデックスのフォルダを返す。
    # 置き場の名前は版によって違うため、測る版の Workspace（core\workspace.ps1。クラスの本体は .NET の型だけを使う）から取る
    param (
        [string]$tool,
        [string]$work
    )

    $workspaceScript = Join-Path (Split-Path -Parent (resolveTebunkoLib $tool)) "core\workspace.ps1"
    if (!(Test-Path -LiteralPath $workspaceScript)) {
        throw "測る tebunko に core\workspace.ps1 がありません。本文インデックスのフォルダを決められません。"
    }
    . ([WildcardPattern]::Escape($workspaceScript.Replace('`', '``')))
    return ([Workspace]::new($work)).IndexDir
}

function getPerfGramWorkDir {
    # 本文インデックスの置き場（-Index）が <W>\store\content\<名前> の形のとき、<W>（ワークスペース）を返す。違えば $null。
    # gram 索引は <W>\store\gram\<名前> に作られ、検索はワークスペースの置き場（store\content）の直下の名前で探すので、この形でなければ gram 索引は効かない
    param (
        [string]$index
    )

    if (!$index) { return $null }
    # Workspace クラスが無い版（置き場の名前を Workspace が持つ前）は、gram 索引の置き場を割り出せない
    if (-not ('Workspace' -as [type])) { return $null }
    # 置き場の名前は Workspace クラスだけが持つ（呼ぶ前に、測る tebunko の lib.ps1 かその core\workspace.ps1 を読み込んでおく）
    $content = [System.IO.Path]::GetDirectoryName($index.TrimEnd("\", "/"))
    if (!$content) { return $null }
    $store = [System.IO.Path]::GetDirectoryName($content)
    if (!$store) { return $null }
    $candidate = [System.IO.Path]::GetDirectoryName($store)
    if (!$candidate) { return $null }
    if (!([Workspace]::new($candidate)).IndexDir.Equals($content, [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
    return $candidate
}

function getChildCrashEvents {
    # from から to までの間に Application ログへ載った、powershell.exe の異常終了に関するイベント（Application Error・.NET Runtime・Windows Error Reporting）を、
    # 1 件 1 行の文字列で返す。件数は maxCount 件、1 件は maxLength 文字まで。読めなければ空を返す
    param (
        [datetime]$from,
        [datetime]$to,
        [int]$maxCount = 5,
        [int]$maxLength = 400
    )

    # 提供元は 1 つずつ読む（機械に登録が無い提供元があると、まとめて読む呼び出し全体が失敗するため。失敗した提供元は飛ばして、残りを取る）
    $found = New-Object System.Collections.Generic.List[object]
    foreach ($provider in @("Application Error", ".NET Runtime", "Windows Error Reporting")) {
        try {
            $filter = @{ LogName = "Application"; ProviderName = $provider; StartTime = $from.AddSeconds(-1); EndTime = $to.AddSeconds(5) }
            foreach ($entry in @(Get-WinEvent -FilterHashtable $filter -MaxEvents 50 -ErrorAction Stop)) {
                if ($entry.Message -match "powershell\.exe") { $found.Add($entry) }
            }
        } catch {
            continue
        }
    }
    return @($found | Sort-Object TimeCreated | Select-Object -First $maxCount | ForEach-Object {
        $oneLine = ($_.Message -replace "\s+", " ").Trim()
        if ($oneLine.Length -gt $maxLength) { $oneLine = $oneLine.Substring(0, $maxLength) + "…" }
        "$($_.TimeCreated.ToString('HH:mm:ss')) $($_.ProviderName) (ID $($_.Id)): $oneLine"
    })
}

function invokeChildScript {
    # 計測の子のスクリプトを別の powershell.exe で動かす。子の出力はそのまま流す（標準エラーも標準出力と同じ流れに入れる）。
    # 0 以外で終わったら、終了コード（10 進と 16 進。落ちたときの NTSTATUS が読める）と、子の出力の終わりの tailLines 行を、例外の文に入れて止める。
    # CI（GITHUB_ACTIONS が true）のときだけ、子が動いていた間の Application ログ（getChildCrashEvents）も入れる。手元の機械のログは読まない。
    # 出力は CI のログと例外の文にだけ出し、ファイルには残さない
    param (
        [string]$scriptPath,
        [string[]]$arguments = @(),
        [int]$tailLines = 20
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $saved = $ErrorActionPreference
    $ErrorActionPreference = "Continue"   # 子の標準エラーを、例外にせず文字として受ける
    $startedAt = Get-Date
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath @arguments 2>&1 | ForEach-Object {
            $text = [string]$_
            $lines.Add($text)
            $text
        }
    } finally {
        $ErrorActionPreference = $saved
    }
    $code = [int]$LASTEXITCODE
    $endedAt = Get-Date
    if ($code) {
        $name = [System.IO.Path]::GetFileName($scriptPath)
        $message = "$name が失敗しました（終了コード $code・0x$('{0:X8}' -f $code)）。"
        $shown = @($lines | Where-Object { $_.Trim() -ne "" } | Select-Object -Last $tailLines | ForEach-Object { if ($_.Length -gt 300) { $_.Substring(0, 300) + "…" } else { $_ } })
        if ($shown.Count -eq 0) {
            $message += "`n子は何も出力しませんでした。"
        } else {
            $message += "`n子の出力の終わり（最後の $($shown.Count) 行）:`n" + ($shown -join "`n")
        }
        if ($env:GITHUB_ACTIONS -eq "true") {
            $events = @()
            try { $events = @(getChildCrashEvents $startedAt $endedAt) } catch { $events = @() }
            if ($events.Count -gt 0) { $message += "`n子が動いていた間の Application ログ（powershell.exe に関するもの）:`n" + ($events -join "`n") }
        }
        throw $message
    }
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

function formatPerfPacks {
    # 検索の表の「照合した pack（gram で絞った数／全部）」の 1 つ分。gram 索引を使わなかった語（回数が 0）は、絞った数を「–」にして全部だけ出す
    param (
        $packs,
        $allPacks,
        [int]$gramUsed
    )

    $all = if ($null -ne $allPacks) { "{0:N0}" -f [int]$allPacks } else { "–" }
    if ($gramUsed -le 0 -or $null -eq $packs) { return "–／$all" }
    return ("{0:N0}／{1}" -f [int]$packs, $all)
}

function getPerfGramLines {
    # 「インデックス作成」の節に足す、gram 索引の作成の行。作らなかったときは理由を書く。gram の結果が無ければ空
    param (
        $gram
    )

    if ($null -eq $gram) { return @() }
    if ($gram.Skipped) {
        return @("gram 索引は作らなかった（$($gram.Skipped)）。この結果の検索は gram 索引なしの数字になる。")
    }
    return @(
        "pack の作成のあと、検索が使う gram 索引を作った（updateGramIndexes。作成の失敗は、そのインデックスを飛ばして続ける）。",
        "",
        "| gram 索引の作成 | ファイル | 合計 | 失敗したインデックス |",
        "|---|---|---|---|",
        ("| {0:N1} 秒 | {1:N0} | {2:N1} MB | {3} |" -f $gram.Seconds, $gram.Files, $gram.MB, $gram.Failed)
    )
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
    #   rows        : searches.csv の行（Word・Hits・HitTotal・Truncated・Packs・GramUsed・AllPacks）。1 回ごとの値
    #   limits      : 語の名前 → 検索の中央値の上限（ms）。比べる値は ShownMs（表がそろうまで）の中央値（無ければ TotalMs）
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
        # 語ごとの上限と比べるのは ShownMs（表がそろうまで）。TotalMs（総数が出るまで）は比べない（記録だけ）
        $shown = if ($null -ne $entry.ShownMs) { $entry.ShownMs } else { $entry.TotalMs }
        if ([double]$shown.Median -gt [double]$limits[$name]) {
            $problems.Add([string]::Format([CultureInfo]::InvariantCulture, "検索 {0}: 中央値 {1:N0} ms（上限 {2:N0} ms）", $name, [double]$shown.Median, [double]$limits[$name]))
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
        if ($want -eq "10000+") {
            # 総数は 1 万件より大きく、全回とも同じ（数え終えた総数）
            $totals = @($mine | ForEach-Object { if ([string]::IsNullOrWhiteSpace([string]$_.HitTotal)) { -1 } else { [int]$_.HitTotal } })
            $small = @($totals | Where-Object { $_ -le 10000 })
            if ($small.Count) {
                $problems.Add("検索 ${name}: 総数（HitTotal）が 10000 より大きくない回が $($small.Count) 回（最初は $($small[0])）")
            }
            if (@($totals | Select-Object -Unique).Count -gt 1) {
                $problems.Add("検索 ${name}: 総数（HitTotal）が回ごとに違う（$((@($totals | Select-Object -Unique)) -join '・')）")
            }
        }
        # gram 索引で絞った回（GramUsed）は、照合した pack が 0〜全部の間（0 件の語は全部が外れる）。全部の数（AllPacks）は pack の数と同じ。絞らない回は全部を照合する
        $gramRows = @($mine | Where-Object { "$($_.GramUsed)" -eq "True" })
        $badPacks = @($mine | Where-Object { "$($_.GramUsed)" -ne "True" -and ([int]$_.Packs -le 0 -or [int]$_.Packs -ne $packCount) })
        if ($badPacks.Count) {
            $problems.Add("検索 ${name}: 照合した pack の数が $packCount と違う回が $($badPacks.Count) 回（最初は $($badPacks[0].Packs)）")
        }
        $outside = @($gramRows | Where-Object { [int]$_.Packs -lt 0 -or [int]$_.Packs -gt $packCount })
        if ($outside.Count) {
            $problems.Add("検索 ${name}: gram 索引で絞った回の照合した pack が 0〜$packCount の外の回が $($outside.Count) 回（最初は $($outside[0].Packs)）")
        }
        $badAll = @($gramRows | Where-Object { $null -ne $_.AllPacks -and "$($_.AllPacks)" -ne "" -and [int]$_.AllPacks -ne $packCount })
        if ($badAll.Count) {
            $problems.Add("検索 ${name}: 全部の数（AllPacks）が $packCount と違う回が $($badAll.Count) 回（最初は $($badAll[0].AllPacks)）")
        }
    }
    return , $problems.ToArray()
}

function getGramPerfProblems {
    # 検索の回帰テストで、gram 索引が作られ、検索が使ったかの判定。合わないものの一覧を返す。空なら合格
    #   result : measure_perf.ps1 の result.json（Index.Gram を見る）
    #   rows   : searches.csv の行
    #   word   : gram で絞れる語の名前（既定は「まれ」）。この語の GramUsed が 1 回以上あること
    param (
        $result,
        [object[]]$rows,
        [string]$word = "まれ"
    )

    $problems = New-Object System.Collections.Generic.List[string]
    $gram = $result.Index.Gram
    if ($null -eq $gram) {
        $problems.Add("gram 索引の作成の結果がありません")
    } elseif ($gram.Skipped) {
        $problems.Add("gram 索引を作っていません（$($gram.Skipped)）")
    } elseif ([int]$gram.Failed -gt 0) {
        $problems.Add("gram 索引の作成に失敗したインデックスがあります（$($gram.Failed) 件）")
    }
    $used = @($rows | Where-Object { $_.Word -eq $word -and "$($_.GramUsed)" -eq "True" })
    if ($used.Count -eq 0) {
        $problems.Add("検索 ${word}: gram 索引を使った回がありません")
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

function newRenderResult {
    # measure_preview.ps1 の Render（-Render の結果）から、result.json の Render を作る。
    #   RenderMs … 描画（Measure・Arrange・UpdateLayout）の時間を測った順に並べたもの
    #   Elements … できた部品の数。ScrollMs … 下と右へスクロールして見える範囲を作り直す時間。HScrollMs … 30 px 刻みで続けて 1 画面ぶん右へ動かす時間（HScrollChanges はそのとき範囲を決め直した回数）。RenderRetainedMB … 描画のあとに残るマネージドヒープの増え
    #   （この 3 つは、測った版が記録していなければ $null）
    #   FirstMs … 1 回目（部品の初期化が入る）。RestMs … 2 回目以降の統計（1 回だけのときは $null）
    param (
        $render
    )

    $times = [double[]]@($render.RenderMs)
    $rest = $null
    if ($times.Count -gt 1) { $rest = getStats ([double[]]@($times[1..($times.Count - 1)])) }
    return [ordered]@{
        Cells = $render.Cells; Rows = $render.Rows; Columns = $render.Columns
        Elements = $render.Elements; ScrollMs = $render.ScrollMs; HScrollMs = $render.HScrollMs; HScrollChanges = $render.HScrollChanges; RenderRetainedMB = $render.RenderRetainedMB
        RenderMs = $times
        FirstMs = $(if ($times.Count) { $times[0] } else { $null })
        RestMs = $rest
        WorkingSetBeforeMB = $render.WorkingSetBeforeMB; WorkingSetAfterMB = $render.WorkingSetAfterMB
    }
}

function newReadMoreResult {
    # measure_preview.ps1 の ReadMore（-ReadMore の結果。下に 200 行ずつ読み足した回ごとの記録）から、result.json の ReadMore を作る。
    #   ReadMs・BuildMs・CellsMs … 読む・行を作る・見える範囲のセルを作る時間の統計。CreatedCells … 最後に作ってあったセルの数の合計
    param (
        $runs
    )

    $items = @($runs)
    $last = $items[$items.Count - 1]
    return [ordered]@{
        Runs = $items.Count
        # 場所の終わりで止まり、5 回に届かなかったとき・どれかの回が 200 行に満たなかった（短い・空の回があった）ときは $false（少ない行の値を、5 回分と見間違えないため）
        Complete = ($items.Count -ge 5 -and @($items | Where-Object { [int]$_.Rows -ne $readMoreChunkLines }).Count -eq 0)
        ReadMs = (getStats ([double[]]@($items | ForEach-Object { $_.ReadMs })))
        BuildMs = (getStats ([double[]]@($items | ForEach-Object { $_.BuildMs })))
        CellsMs = (getStats ([double[]]@($items | ForEach-Object { $_.CellsMs })))
        TableRows = [int]$last.TableRows; Columns = [int]$last.Columns; CreatedCells = [int]$last.CreatedCells
    }
}

function newPreviewResult {
    # measure_preview.ps1 の結果（ConvertFrom-Json したもの）から、result.json の Preview の 1 語分を作る。
    #   読み込みの時間はキャッシュの有無ごと、表を作る時間とセルの数は全部まとめて集計する。
    #   RetainedManagedMB … プレビューの前後で、GC のあとに残ったマネージドヒープの増え（プレビューが何かを持ち続けていないかを見る）
    #   RetainedPerPreviewMB … 上をプレビューの回数で割った 1 回あたり
    #   CacheDropMB … 検索のキャッシュを捨てて GC したあとの、ワーキングセットの減り（キャッシュの常駐の大きさ）
    #   ReadTextCallsUncached … キャッシュ無しのプレビューで readContentIndexText が呼ばれた回数の合計（ファイルの全体を読んだ回数）
    #   FullRead … 前の読み方（毎回ファイルの全体を読む）で測った同じ形の結果。測っていなければ $null（呼ぶ側が入れる）
    param (
        $raw
    )

    $rows = @($raw.Previews)
    $uncached = @($rows | Where-Object { $_.Cache -eq "無し" })
    $cached = @($rows | Where-Object { $_.Cache -eq "あり" })
    $mem = $raw.Memory
    $retained = [Math]::Round([double]$mem.AfterGcManagedMB - [double]$mem.BeforePreviewManagedMB, 1)
    $entry = [ordered]@{
        Name = $raw.Name; Regex = [bool]$raw.Regex; Count = [int]$raw.Count; Previews = $rows.Count
        ReadUncachedMs = (getStats ([double[]]@($uncached | ForEach-Object { $_.ReadMs })))
        ReadCachedMs = (getStats ([double[]]@($cached | ForEach-Object { $_.ReadMs })))
        BuildMs = (getStats ([double[]]@($rows | ForEach-Object { $_.BuildMs })))
        BuildFirstMs = $(if ($rows.Count) { [double]$rows[0].BuildMs } else { $null })
        BuildRestMs = $(if ($rows.Count -gt 1) { getStats ([double[]]@($rows[1..($rows.Count - 1)] | ForEach-Object { $_.BuildMs })) } else { $null })
        Cells = (getStats ([double[]]@($rows | ForEach-Object { $_.Cells })))
        CreatedCells = $(if (@($rows | Where-Object { $null -ne $_.CreatedCells }).Count -gt 0) { getStats ([double[]]@($rows | ForEach-Object { $_.CreatedCells })) } else { $null })
        ReadMore = $(if ($null -ne $raw.ReadMore -and @($raw.ReadMore).Count -gt 0) { newReadMoreResult $raw.ReadMore } else { $null })
        FileMBMax = $(if ($rows.Count) { ($rows | Measure-Object FileMB -Maximum).Maximum } else { $null })
        ReadTextCallsUncached = [int](($uncached | Measure-Object -Property ReadTextCalls -Sum).Sum)
        FullRead = $null
        RetainedManagedMB = $retained
        RetainedPerPreviewMB = $(if ($rows.Count) { [Math]::Round($retained / $rows.Count, 3) } else { $null })
        WorkingSetAfterSearchMB = $mem.AfterSearchWorkingSetMB; WorkingSetAfterPreviewMB = $mem.AfterPreviewWorkingSetMB
        WorkingSetAfterGcMB = $mem.AfterGcWorkingSetMB; WorkingSetAfterCacheDropMB = $mem.AfterCacheDropWorkingSetMB
        CacheDropMB = [Math]::Round([double]$mem.AfterGcWorkingSetMB - [double]$mem.AfterCacheDropWorkingSetMB, 1)
        CacheChars = [long]$mem.CacheChars
        Render = $(if ($null -eq $raw.Render) { $null } else { newRenderResult $raw.Render })
        PeakWorkingSetMB = $raw.PeakWorkingSetMB
    }
    return $entry
}

function newLargeContentTsvText {
    # プレビューの計測用に、大きな本文インデックスのファイル（TSV）の中身を作る。架空の文だけで、人名・社名・パスは入れない。
    #   lines    : 行の数（行番号は 1 から lines まで）
    #   hitWord  : 検索する語。hitLines の行にだけ入れる
    #   hitLines : 語を入れる行番号
    #   repeat   : 1 行の長さを決める、埋め草の文の繰り返し（約 95 バイト × repeat）
    # 行番号は 4 桁までにしておく（正規表現の語 ^\d{5}\t が、5 桁の行番号の行に当たらないように。lines は 9999 まで）。
    # 形は「行番号 TAB 行N TAB 語または架空説明 + 埋め草」を CRLF でつないだもの（BOM は呼ぶ側で付ける）。
    param (
        [int]$lines,
        [string]$hitWord,
        [int[]]$hitLines,
        [int]$repeat = 10
    )

    if ($lines -lt 1 -or $lines -gt 9999) { throw "lines は 1〜9999 にしてください。" }
    $filler = (@("下書き 確認 覚え書き 日程 調整 備品 発注 会議室 予約 資料 保管場所") * $repeat) -join " "
    $hits = New-Object System.Collections.Generic.HashSet[int]
    foreach ($hitLine in $hitLines) { [void]$hits.Add($hitLine) }
    $builder = New-Object System.Text.StringBuilder
    for ($lineNo = 1; $lineNo -le $lines; $lineNo++) {
        $note = if ($hits.Contains($lineNo)) { $hitWord } else { "架空説明" }
        [void]$builder.Append("$lineNo`t行$lineNo`t$note $filler`r`n")
    }
    return $builder.ToString()
}

function newWideContentTsvText {
    # プレビューの計測用に、横に広い Excel の本文インデックスのファイル（TSV）の中身を作る。架空の値だけで、人名・社名・パスは入れない。
    #   lines     : 行の数（行番号は 1 から lines まで。9999 まで）
    #   columns   : 行ごとのセルの数（2000 など）
    #   hitWord   : 検索する語。hitLines の行の hitColumn 列目（1 から）にだけ入れる
    #   hitLines  : 語を入れる行番号
    #   hitColumn : 語を入れる列（1 から）。既定は右寄りの 1500 列目（一致したセルが、最初に見える範囲の外にある表を作る）
    # 形は「行番号 TAB セル 1 TAB セル 2 …」を CRLF でつないだもの（BOM は呼ぶ側で付ける）。セルは「行 N 列 M」の短い文字
    param (
        [int]$lines,
        [int]$columns,
        [string]$hitWord,
        [int[]]$hitLines,
        [int]$hitColumn = 1500
    )

    if ($lines -lt 1 -or $lines -gt 9999) { throw "lines は 1〜9999 にしてください。" }
    if ($columns -lt 1) { throw "columns は 1 以上にしてください。" }
    if ($hitColumn -lt 1 -or $hitColumn -gt $columns) { throw "hitColumn は 1〜columns にしてください。" }
    $hits = New-Object System.Collections.Generic.HashSet[int]
    foreach ($hitLine in $hitLines) { [void]$hits.Add($hitLine) }
    $builder = New-Object System.Text.StringBuilder
    for ($lineNo = 1; $lineNo -le $lines; $lineNo++) {
        [void]$builder.Append($lineNo)
        for ($col = 1; $col -le $columns; $col++) {
            [void]$builder.Append("`t")
            if ($col -eq $hitColumn -and $hits.Contains($lineNo)) { [void]$builder.Append($hitWord) } else { [void]$builder.Append("行${lineNo}列$col") }
        }
        [void]$builder.Append("`r`n")
    }
    return $builder.ToString()
}

function getFullReadPathProblems {
    # -FullRead で測ったプレビューの記録（measure_preview.ps1 の Previews）が、本当に全部読む経路を通ったかを確かめる。
    # キャッシュ無しの選択で readContentIndexText が 1 回も呼ばれていないものがあれば、その番号を文字列で返す（無ければ何も返さない。呼ぶ側は @() で受ける）。
    param (
        $rows
    )

    foreach ($row in @($rows | Where-Object { $_.Cache -eq "無し" })) {
        if ([int]$row.ReadTextCalls -lt 1) { "N=$($row.N) で readContentIndexText が呼ばれていません" }
    }
}

function getPreviewPerfProblems {
    # プレビューの回帰テストの判定。合わないものの一覧（文字列の配列）を返す。空なら合格。
    #   preview      : result.json の Preview の 1 語分（newPreviewResult の結果。無ければ $null）
    #   readLimitMs  : キャッシュ無しの読み込みの中央値の上限（ms）
    #   retainedLimitMB : プレビューのあと、GC のあとも残るマネージドヒープの増えの合計の上限（MB。1 回あたりではなく合計で持つ）
    #   expectedCount: 測ったプレビューの数（キャッシュの有無ごと）
    param (
        $preview,
        [double]$readLimitMs,
        [double]$retainedLimitMB,
        [int]$expectedCount
    )

    $problems = New-Object System.Collections.Generic.List[string]
    if ($null -eq $preview) {
        $problems.Add("プレビューの結果がありません")
        return , $problems.ToArray()
    }
    if ([int]$preview.Previews -ne $expectedCount * 2) { $problems.Add("測ったプレビューが $($preview.Previews) 回（期待 $($expectedCount * 2) 回）") }
    if ($null -eq $preview.ReadUncachedMs) {
        $problems.Add("キャッシュ無しの読み込みの時間がありません")
    } elseif ([double]$preview.ReadUncachedMs.Median -gt $readLimitMs) {
        $problems.Add([string]::Format([CultureInfo]::InvariantCulture, "プレビューの読み込み（キャッシュ無し）: 中央値 {0:N1} ms（上限 {1:N1} ms）", [double]$preview.ReadUncachedMs.Median, $readLimitMs))
    }
    if ($null -eq $preview.Cells -or [double]$preview.Cells.Min -le 0) {
        $problems.Add("セルを作らなかったプレビューがあります")
    }
    if ($null -eq $preview.RetainedManagedMB) {
        $problems.Add("残ったメモリの増えがありません")
    } elseif ([double]$preview.RetainedManagedMB -gt $retainedLimitMB) {
        $problems.Add([string]::Format([CultureInfo]::InvariantCulture, "プレビューのあとに残るメモリ（合計）: {0:N1} MB（上限 {1:N1} MB）", [double]$preview.RetainedManagedMB, $retainedLimitMB))
    }
    return , $problems.ToArray()
}

function getPreviewPerfMessages {
    # getPreviewPerfProblems の結果に語の名前を付け、1 件ずつの文字列にして返す（回帰テストの失敗メッセージ用）。
    # getPreviewPerfProblems は空の配列を 1 つの値として返すので、いったん変数で受けて展開する。合格なら何も返さない（呼ぶ側は @() で受ける）。
    param (
        [string]$name,
        $preview,
        [double]$readLimitMs,
        [double]$retainedLimitMB,
        [int]$expectedCount
    )

    $problems = getPreviewPerfProblems $preview $readLimitMs $retainedLimitMB $expectedCount
    foreach ($problem in $problems) { "プレビュー ${name}: $problem" }
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
    #             IntervalMs; LagMs; RequestMs（要求ごとの値の配列）; Hits; Truncated; Packs; GramUsed; CancelledSteps }
    #   steps   : 送った語の数（getTypingSteps の結果の数）。並びが空だった語は 0
    #   skipped : 測らなかった理由（並びが空の語）。$null なら測った
    # 統計（Min・Median・Mean・Max・Count）は getStats を使う。FirstHitMs・FinishMs・WaitPrevMs は 2 回目以降（attempts の 2 件目から）、
    # IntervalMs・LagMs・RequestMs は全回（1 回目を含む）から作る。CancelledSteps は 2 回目以降の合計、GramUsed・Narrowed（最後の要求が前の結果から絞って答えた）は全回の合計（回数）。
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
            Hits = $null; Truncated = $null; Packs = $null; GramUsed = $null; Narrowed = $null; CancelledSteps = $null
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
        GramUsed = @($attempts | Where-Object { $_.GramUsed }).Count
        Narrowed = @($attempts | Where-Object { $_.Narrowed }).Count
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
