# pack の作成と検索の速さ、そのあいだのリソース（メモリ・CPU・スレッド・GC）を測る。GitHub Actions の perf.yml と手元で使う。
# 結果には時間・件数・大きさだけを出し、パス・ファイル名は出さない（そのまま共有できる）。
#
#   .\tools\measure_perf.ps1 -Index <TSV のインデックス> -Work <作業フォルダ> -Words <words.tsv>
#   -Index    … 場所ごとの TSV（取り込みの一時置き場の形。tebunko-perfdata の new_index.ps1 で作る）。
#                pack に書き換えて TSV を消すので、作り直したものを渡す。pack だけなら作成は測らず検索だけを測る
#   -Tool     … 測る tebunko のフォルダ（scripts\tebunko_grep\lib.ps1 を読む）。既定はこのスクリプトのリポジトリ
#   -Words    … 検索する語の表（名前・語・正規表現・件数。tebunko-perfdata の words.tsv）
#   -AddWords … 足して検索する語（文字どおり）
#   -Repeat   … 検索を繰り返す回数（中央値を出す）
#   -SampleMs … リソースを記録する間隔（ミリ秒）
#   -Label    … 結果の見出しに出す名前（測る ref など）
#   -DataSeconds … データ（TSV）の作成にかかった秒数。渡せば結果に書く
#   -Out      … 結果（summary.md・result.json・resource.csv）を書くフォルダ。既定は <Work>\result
#
# 段階:
#   1. 環境            … PowerShell・OS・CPU・メモリ・Defender のリアルタイム保護
#   2. pack の作成     … インデクサと同じ publishIndexFolders（pack・TSV の削除・システムインデックス）。Office の取り込みは含まない
#   3. 検索            … 画面の検索スレッドと同じ流れ（新しい Runspace で lib.ps1 を読む → 列挙 → 照合）。
#                        読んだ内容の入れ物（newTsvTextCache）を新しくしてから初回・2 回目を行い、これを -Repeat 回繰り返す
#   リソースは、別のスレッドで -SampleMs ごとにこのプロセスと PC 全体の値を記録し、段階ごとにまとめる
param (
    [Parameter(Mandatory = $true)][string]$Index,
    [Parameter(Mandatory = $true)][string]$Work,
    [string]$Tool = (Split-Path $PSScriptRoot -Parent),
    [string]$Words,
    [string[]]$AddWords = @(),
    [int]$Repeat = 3,
    [int]$SampleMs = 200,
    [string]$Label = "",
    [double]$DataSeconds = -1,
    [string]$Out
)

$ErrorActionPreference = "Stop"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$utf8 = New-Object System.Text.UTF8Encoding($false)
$Tool = (Resolve-Path -LiteralPath $Tool).ProviderPath
$lib = Join-Path $Tool "scripts\tebunko_grep\lib.ps1"
if (!(Test-Path -LiteralPath $lib)) {
    throw "測る tebunko に scripts\tebunko_grep\lib.ps1 がありません。-Tool に tebunko のフォルダを指定してください。"
}
. $lib
foreach ($name in @("findIndexFoldersWithBooks", "publishIndexFolders", "getIndexPackFiles", "searchPackIndex", "newTsvTextCache")) {
    if (!(Get-Command $name -CommandType Function -ErrorAction SilentlyContinue)) {
        throw "測る tebunko に $name がありません。pack 形式より前の版は測れません（この道具が呼ぶ関数の形が変わった場合は、道具を直してください）。"
    }
}
$Index = (Resolve-Path -LiteralPath $Index).ProviderPath.TrimEnd("\")
[void][System.IO.Directory]::CreateDirectory($Work)
$Work = (Resolve-Path -LiteralPath $Work).ProviderPath.TrimEnd("\")
if (!$Out) { $Out = Join-Path $Work "result" }
[void][System.IO.Directory]::CreateDirectory($Out)

# 検索する語（words.tsv の 1 行目は見出し）
$wordList = New-Object System.Collections.Generic.List[hashtable]
if (!$Words) {
    $wordList.Add(@{ Name = "0 件"; Word = "tebunko計測用の存在しない語"; Regex = $false; Expected = "0" })
} else {
    $lines = [System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath $Words).ProviderPath, [System.Text.Encoding]::UTF8)
    foreach ($line in @($lines | Select-Object -Skip 1)) {
        $f = $line.Split("`t")
        if ($f.Count -lt 3 -or $f[1] -eq "") { continue }
        $wordList.Add(@{ Name = $f[0]; Word = $f[1]; Regex = ($f[2] -eq "true"); Expected = $(if ($f.Count -gt 3) { $f[3] } else { "" }) })
    }
}
$n = 0
foreach ($w in @($AddWords | Where-Object { $_ })) {
    $n++
    $wordList.Add(@{ Name = "追加 $n"; Word = $w; Regex = $false; Expected = "" })
}

# リソースの記録。別のスレッドが -SampleMs ごとに記録し、段階（Phase）を切り替えるときは呼び出し元でも記録する
# （短い段階にも始まりと終わりの値が残るように）。時計は 1 つを共有する
$monitor = [hashtable]::Synchronized(@{ Stop = $false; Phase = "準備"; Clock = [System.Diagnostics.Stopwatch]::StartNew(); Samples = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList)) })
$takeSample = {
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
$sampler = {
    param ($monitor, $sampleMs, $takeText)
    $take = [scriptblock]::Create($takeText)
    $cpuCounter = $null; $memCounter = $null
    try { $cpuCounter = New-Object System.Diagnostics.PerformanceCounter "Processor", "% Processor Time", "_Total"; [void]$cpuCounter.NextValue() } catch { $cpuCounter = $null }
    try { $memCounter = New-Object System.Diagnostics.PerformanceCounter "Memory", "Available MBytes" } catch { $memCounter = $null }
    while (!$monitor.Stop) {
        & $take $monitor $cpuCounter $memCounter
        Start-Sleep -Milliseconds $sampleMs
    }
}
$samplerPs = [powershell]::Create()
[void]$samplerPs.AddScript($sampler).AddArgument($monitor).AddArgument($SampleMs).AddArgument($takeSample.ToString())
$samplerHandle = $samplerPs.BeginInvoke()
$phases = New-Object System.Collections.Generic.List[hashtable]
function setPhase([string]$name) {
    $now = [int]$monitor.Clock.Elapsed.TotalMilliseconds
    if ($phases.Count -gt 0) {
        & $takeSample $monitor $null $null
        $phases[$phases.Count - 1].EndMs = $now
    }
    $monitor.Phase = $name
    if ($name) {
        $phases.Add(@{ Name = $name; StartMs = $now; EndMs = $null })
        & $takeSample $monitor $null $null
    }
}
$result = [ordered]@{ Label = $Label; DataSeconds = $(if ($DataSeconds -ge 0) { [Math]::Round($DataSeconds, 1) } else { $null }); Date = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm 'UTC'") }
try {
    # 1. 環境
    setPhase "環境"
    $envInfo = [ordered]@{ PowerShell = $PSVersionTable.PSVersion.ToString(); OS = [Environment]::OSVersion.VersionString; Cores = [Environment]::ProcessorCount }
    try { $envInfo.Cpu = (@(Get-CimInstance Win32_Processor)[0].Name).Trim() } catch { $envInfo.Cpu = "不明" }
    try { $envInfo.MemoryGB = [Math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 1) } catch { $envInfo.MemoryGB = $null }
    try { $envInfo.DefenderRealtime = [string](Get-MpComputerStatus -ErrorAction Stop).RealTimeProtectionEnabled } catch { $envInfo.DefenderRealtime = "不明" }
    $result.Environment = $envInfo

    # 2. pack の作成（インデクサと同じ関数。TSV を置いたフォルダをまとめて書き出す）
    $systemRoot = Join-Path $Work "system_index"
    $statePath = Join-Path $Work "システムインデックスの状態.tsv"
    $folders = [string[]](findIndexFoldersWithBooks $Index)
    if ($folders.Count -gt 0) {
        $tsvFiles = [System.IO.Directory]::GetFiles($Index, "*.tsv", "AllDirectories")
        $tsvBytes = 0L
        foreach ($f in $tsvFiles) { $tsvBytes += ([System.IO.FileInfo]$f).Length }
        $bookCount = [System.IO.Directory]::GetDirectories($Index, "*.*", "AllDirectories").Where({ testIndexBookDir $_ }).Count
        Write-Host ("pack を作ります（フォルダ {0:N0}・ブック {1:N0}・TSV {2:N0} 件・{3:N0} MB）…" -f $folders.Count, $bookCount, $tsvFiles.Count, ($tsvBytes / 1MB))
        $pending = @{}
        foreach ($f in $folders) { $pending[$f] = [string[]]@() }
        setPhase "pack の作成"
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        [void](publishIndexFolders $pending $Index $systemRoot $statePath)
        $packMs = $watch.Elapsed.TotalMilliseconds
        setPhase "pack の集計"
        $packs = (getIndexPackFiles @($Index)).Packs
        $packBytes = 0L
        foreach ($p in $packs) { $packBytes += $p.Size }
        $systemBytes = 0L
        if ([System.IO.Directory]::Exists($systemRoot)) {
            foreach ($f in [System.IO.Directory]::GetFiles($systemRoot, "*", "AllDirectories")) { $systemBytes += ([System.IO.FileInfo]$f).Length }
        }
        $result.Pack = [ordered]@{ Folders = $folders.Count; Books = $bookCount; Tsv = $tsvFiles.Count; TsvMB = [Math]::Round($tsvBytes / 1MB, 1)
            Seconds = [Math]::Round($packMs / 1000, 2); Packs = $packs.Count; PackMB = [Math]::Round($packBytes / 1MB, 1); SystemIndexMB = [Math]::Round($systemBytes / 1MB, 1) }
        Write-Host ("  {0:N1} 秒" -f ($packMs / 1000))
    } else {
        Write-Host "TSV のフォルダがありません（pack だけのインデックス）。pack の作成は測りません。"
        $result.Pack = $null
    }

    # 3. 検索（画面の検索スレッドと同じ流れ。読んだ内容の入れ物は、画面と同じく 1 回の繰り返しのあいだ持ち続ける）
    $searchScript = {
        param ($lib, $index, $word, $simple, $cache)
        $total = [System.Diagnostics.Stopwatch]::StartNew(); $step = [System.Diagnostics.Stopwatch]::StartNew()
        . $lib
        $load = $step.Elapsed.TotalMilliseconds; $step.Restart()
        $found = getIndexPackFiles @($index)
        $list = $step.Elapsed.TotalMilliseconds; $step.Restart()
        $r = searchPackIndex $word $found.Packs $simple 10000 -cache $cache
        @{ Load = $load; List = $list; Search = $step.Elapsed.TotalMilliseconds; Total = $total.Elapsed.TotalMilliseconds; Packs = $found.Packs.Count; Hits = $r.Hits.Count; Truncated = [bool]$r.Truncated }
    }
    $runs = New-Object System.Collections.Generic.List[object]
    for ($rep = 1; $rep -le $Repeat; $rep++) {
        $cache = newTsvTextCache
        foreach ($round in @("初回", "2 回目")) {
            setPhase "検索 $rep $round"
            foreach ($w in $wordList) {
                $ps = [powershell]::Create()
                try {
                    [void]$ps.AddScript($searchScript.ToString()).AddArgument($lib).AddArgument($Index).AddArgument($w.Word).AddArgument(!$w.Regex).AddArgument($cache)
                    $r = $ps.Invoke()[0]
                } finally {
                    $ps.Dispose()
                }
                $runs.Add([pscustomobject]@{ Repeat = $rep; Round = $round; Name = $w.Name; Load = $r.Load; List = $r.List; Search = $r.Search; Total = $r.Total; Packs = $r.Packs; Hits = $r.Hits; Truncated = $r.Truncated })
                Write-Host ("  {0} {1} {2}: {3:N0} ms（{4}{5} 件）" -f $rep, $round, $w.Name, $r.Total, $r.Hits, $(if ($r.Truncated) { "+" } else { "" }))
            }
        }
    }
    setPhase ""
} finally {
    $monitor.Stop = $true
    [void]$samplerPs.EndInvoke($samplerHandle)
    $samplerPs.Dispose()
}

# まとめ
function median([double[]]$values) {
    $sorted = @($values | Sort-Object)
    if ($sorted.Count -eq 0) { return $null }
    $mid = [int][Math]::Floor($sorted.Count / 2)
    if ($sorted.Count % 2) { return $sorted[$mid] }
    return ($sorted[$mid - 1] + $sorted[$mid]) / 2
}
$searchRows = New-Object System.Collections.Generic.List[object]
foreach ($w in $wordList) {
    foreach ($round in @("初回", "2 回目")) {
        $set = @($runs | Where-Object { $_.Name -eq $w.Name -and $_.Round -eq $round })
        if ($set.Count -eq 0) { continue }
        $totals = [double[]]@($set | ForEach-Object { $_.Total })
        $searchRows.Add([ordered]@{
            Name = $w.Name; Regex = $w.Regex; Round = $round; Expected = $w.Expected
            Hits = "$($set[0].Hits)$(if ($set[0].Truncated) { '+' })"; Packs = $set[0].Packs
            TotalMs = [Math]::Round((median $totals), 0); MinMs = [Math]::Round(($totals | Measure-Object -Minimum).Minimum, 0); MaxMs = [Math]::Round(($totals | Measure-Object -Maximum).Maximum, 0)
            LoadMs = [Math]::Round((median ([double[]]@($set | ForEach-Object { $_.Load }))), 0)
            ListMs = [Math]::Round((median ([double[]]@($set | ForEach-Object { $_.List }))), 0)
            SearchMs = [Math]::Round((median ([double[]]@($set | ForEach-Object { $_.Search }))), 0)
        })
    }
}
$result.Search = $searchRows.ToArray()
$result.SearchRuns = $runs.ToArray()

# 段階ごとのリソース（記録の段階名でまとめる）
$samples = @($monitor.Samples.ToArray())
$cores = [Environment]::ProcessorCount
$phaseRows = New-Object System.Collections.Generic.List[object]
foreach ($name in @($samples | ForEach-Object { $_.Phase } | Where-Object { $_ } | Select-Object -Unique)) {
    $set = @($samples | Where-Object { $_.Phase -eq $name })
    $first = $set[0]; $last = $set[$set.Count - 1]
    $wallMs = [Math]::Max(1, $last.Ms - $first.Ms)
    $pcCpu = @($set | Where-Object { $null -ne $_.PcCpuPercent } | ForEach-Object { $_.PcCpuPercent })
    $phaseRows.Add([ordered]@{
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
$result.Resources = $phaseRows.ToArray()
$peak = [System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64
$result.PeakWorkingSetMB = [Math]::Round($peak / 1MB, 1)

# 書き出す: result.json・resource.csv・summary.md
[System.IO.File]::WriteAllText((Join-Path $Out "result.json"), ($result | ConvertTo-Json -Depth 6), $utf8)
$csv = @($samples | ConvertTo-Csv -NoTypeInformation)
[System.IO.File]::WriteAllLines((Join-Path $Out "resource.csv"), [string[]]$csv, $utf8Bom)

$md = New-Object System.Collections.Generic.List[string]
$md.Add("# tebunko の性能$(if ($Label) { "（$Label）" })")
$md.Add("")
$md.Add("$($result.Date)。時間・件数・大きさだけを出す。")
$md.Add("")
$md.Add("## 環境")
$md.Add("")
$md.Add("| PowerShell | OS | CPU | 論理コア | メモリ | Defender のリアルタイム保護 |")
$md.Add("|---|---|---|---|---|---|")
$md.Add("| $($envInfo.PowerShell) | $($envInfo.OS) | $($envInfo.Cpu) | $($envInfo.Cores) | $($envInfo.MemoryGB) GB | $($envInfo.DefenderRealtime) |")
$md.Add("")
$md.Add("## pack の作成")
$md.Add("")
if ($null -ne $result.DataSeconds) { $md.Add(("データ（TSV）の作成: {0:N1} 秒（tebunko-perfdata の new_index.ps1。tebunko の処理ではない）。" -f $result.DataSeconds)); $md.Add("") }
if ($result.Pack) {
    $pk = $result.Pack
    $md.Add("Office の取り込みを除いた、インデックス作成の後半（pack を書く・TSV を消す・システムインデックスを作る）。")
    $md.Add("")
    $md.Add("| フォルダ | ブック | TSV | TSV の合計 | 時間 | pack | pack の合計 | システムインデックス |")
    $md.Add("|---|---|---|---|---|---|---|---|")
    $md.Add(("| {0:N0} | {1:N0} | {2:N0} | {3:N1} MB | {4:N1} 秒 | {5:N0} | {6:N1} MB | {7:N1} MB |" -f $pk.Folders, $pk.Books, $pk.Tsv, $pk.TsvMB, $pk.Seconds, $pk.Packs, $pk.PackMB, $pk.SystemIndexMB))
} else {
    $md.Add("測っていない（渡したインデックスに TSV が無い）。")
}
$md.Add("")
$md.Add("## 検索")
$md.Add("")
$md.Add("画面の検索と同じ流れ（新しいスレッドで lib.ps1 を読む → 列挙 → 照合。上限 1 万件）。$Repeat 回の中央値（最小〜最大）。初回は読んだ内容の入れ物が空の検索（OS のファイルのキャッシュは効いている）。")
$md.Add("")
$md.Add("| 語 | 回 | ヒット | 合計 | lib.ps1 | 列挙 | 照合 |")
$md.Add("|---|---|---|---|---|---|---|")
foreach ($s in $searchRows) {
    $md.Add(("| {0} | {1} | {2} | {3:N0} ms（{4:N0}〜{5:N0}） | {6:N0} ms | {7:N0} ms | {8:N0} ms |" -f $s.Name, $s.Round, $s.Hits, $s.TotalMs, $s.MinMs, $s.MaxMs, $s.LoadMs, $s.ListMs, $s.SearchMs))
}
$md.Add("")
$md.Add("## リソース（段階ごと）")
$md.Add("")
$md.Add("このプロセス（計測と tebunko の処理）の値。CPU はこのプロセスの使用率（全コアで 100%）、PC は PC 全体の平均。ピークのワーキングセットは $($result.PeakWorkingSetMB) MB。")
$md.Add("")
$md.Add("| 段階 | 時間 | CPU 時間 | CPU | PC の CPU | ワーキングセット | プライベート | マネージド | スレッド | ハンドル | GC 0/1/2 |")
$md.Add("|---|---|---|---|---|---|---|---|---|---|---|")
foreach ($p in $phaseRows) {
    $md.Add(("| {0} | {1:N1} 秒 | {2:N1} 秒 | {3}% | {4} | {5:N0} MB | {6:N0} MB | {7:N0} MB | {8} | {9} | {10}/{11}/{12} |" -f $p.Phase, $p.Seconds, $p.CpuSeconds, $p.CpuPercent, $(if ($null -ne $p.PcCpuPercent) { "$($p.PcCpuPercent)%" } else { "–" }), $p.WorkingSetMaxMB, $p.PrivateMaxMB, $p.ManagedMaxMB, $p.ThreadsMax, $p.HandlesMax, $p.Gc0, $p.Gc1, $p.Gc2))
}

# 推移のグラフ（Mermaid。点は 100 個ほどに間引く）
if ($samples.Count -ge 2) {
    $step = [Math]::Max(1, [int][Math]::Ceiling($samples.Count / 100))
    $points = @(for ($i = 0; $i -lt $samples.Count; $i += $step) { $samples[$i] })
    $xs = ($points | ForEach-Object { [Math]::Round($_.Ms / 1000, 1).ToString([System.Globalization.CultureInfo]::InvariantCulture) }) -join ", "
    $cpu = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $points.Count; $i++) {
        if ($i -eq 0) { $cpu.Add("0"); continue }
        $dt = [Math]::Max(1, $points[$i].Ms - $points[$i - 1].Ms)
        $cpu.Add([string][Math]::Round(($points[$i].CpuMs - $points[$i - 1].CpuMs) * 100 / $dt / $cores, 0))
    }
    $md.Add("")
    $md.Add("## 推移")
    $md.Add("")
    $md.Add("段階の始まり（秒）: " + (($phases | ForEach-Object { "{0} {1:N1}" -f $_.Name, ($_.StartMs / 1000) }) -join "、"))
    foreach ($chart in @(
        @{ Title = "メモリ（MB）"; Lines = @(@($points | ForEach-Object { $_.WorkingSetMB }), @($points | ForEach-Object { $_.ManagedMB })); Legend = "ワーキングセット（上）・マネージド（下）" },
        @{ Title = "CPU（このプロセス、%）"; Lines = @(, $cpu.ToArray()); Legend = "" }
    )) {
        $md.Add("")
        if ($chart.Legend) { $md.Add($chart.Legend) ; $md.Add("") }
        $md.Add('```mermaid')
        $md.Add("xychart-beta")
        $md.Add("    title `"$($chart.Title)`"")
        $md.Add("    x-axis `"秒`" [$xs]")
        foreach ($line in $chart.Lines) {
            $md.Add("    line [" + (($line | ForEach-Object { ([double]$_).ToString([System.Globalization.CultureInfo]::InvariantCulture) }) -join ", ") + "]")
        }
        $md.Add('```')
    }
}
[System.IO.File]::WriteAllLines((Join-Path $Out "summary.md"), [string[]]$md, $utf8)
Write-Host "結果を書きました: summary.md・result.json・resource.csv"
