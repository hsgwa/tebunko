# インデックス作成（pack の作成）と検索の速さ、およびその間のリソース（メモリ・CPU・スレッド・ハンドル・GC）を測る。
# GitHub Actions の perf.yml からも、手元からも使う。結果には数字だけを書き、パスやファイル名は書かない（そのまま共有できるようにするため）。
#
#   .\tools\measure_perf.ps1 -Index <TSV のインデックス> -Work <作業フォルダ> -Words <words.tsv>
#   -Index    … 場所ごとの TSV（取り込みの一時置き場の形。tebunko-perfdata の new_index.ps1 で作る）。
#                TSV は pack に変換され、元の TSV は削除されるので、毎回作り直したものを渡す。pack しか無いときは、作成は測らず検索だけを測る
#   -Tool     … 測る tebunko のフォルダ。既定はこのスクリプトのリポジトリ
#   -Words    … 検索する語の表（名前・語・正規表現・件数。tebunko-perfdata の words.tsv）。省略すると、ヒットしない語 1 つだけを検索する
#   -Count    … 語ごとに、同じプロセスで続けて検索する回数
#   -SampleMs … リソースを記録する間隔（ミリ秒）
#   -Label    … 結果の見出しに出す名前
#   -RunId・-Ref・-Sha・-Scale・-DataSeconds … 実行の情報（result.json と metrics.csv に書く。-DataSeconds はデータの生成にかかった秒数）
#   -Out      … 結果を書くフォルダ。既定は <Work>\result
#
# 流れ（それぞれ別のプロセスで動かし、リソースが混ざらないようにする）:
#   1. インデックス作成 … tools\perf\measure_index.ps1（インデクサと同じ publishIndexFolders。Office からの取り込みは含まない）
#   2. 検索            … 語ごとに tools\perf\measure_search.ps1 を起動し、同じプロセスで -Count 回続けて検索する
# OS のファイルキャッシュは空にしない（pack は作成の直後なので、OS のキャッシュに載っている）。
#
# 書くもの:
#   summary.md          … 表と Mermaid のグラフ（GitHub Actions のジョブの Summary に出す）
#   result.json         … すべての数字（形式の版 Schema と実行の情報つき）
#   metrics.csv         … 1 行 1 指標の縦長の形（run_id, date, ref, sha, scale, metric, word, stat, value, unit）。後で Grafana などに入れるため
#   searches.csv        … 検索 1 回ごとの時間・ヒット件数・メモリ
#   resource-index.csv・resource-search.csv … リソースの記録（-SampleMs ごと）
param (
    [Parameter(Mandatory = $true)][string]$Index,
    [Parameter(Mandatory = $true)][string]$Work,
    [string]$Tool = (Split-Path $PSScriptRoot -Parent),
    [string]$Words,
    [int]$Count = 20,
    [int]$SampleMs = 200,
    [string]$Label = "",
    [string]$RunId = "",
    [string]$Ref = "",
    [string]$Sha = "",
    [string]$Scale = "",
    [double]$DataSeconds = -1,
    [string]$Out
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\perf\perf_common.ps1"
$Tool = (Resolve-Path -LiteralPath $Tool).ProviderPath
[void](resolveTebunkoLib $Tool)
$Index = (Resolve-Path -LiteralPath $Index).ProviderPath.TrimEnd("\")
[void][System.IO.Directory]::CreateDirectory($Work)
$Work = (Resolve-Path -LiteralPath $Work).ProviderPath.TrimEnd("\")
if (!$Out) { $Out = Join-Path $Work "result" }
[void][System.IO.Directory]::CreateDirectory($Out)
$raw = Join-Path $Work "raw"
[void][System.IO.Directory]::CreateDirectory($raw)
$utf8 = New-Object System.Text.UTF8Encoding($false)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$cores = [Environment]::ProcessorCount

# 検索する語（words.tsv の 1 行目は見出し）
$wordList = New-Object System.Collections.Generic.List[hashtable]
if ($Words) {
    $lines = [System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath $Words).ProviderPath, [System.Text.Encoding]::UTF8)
    foreach ($line in @($lines | Select-Object -Skip 1)) {
        $f = $line.Split("`t")
        if ($f.Count -lt 3 -or $f[1] -eq "") { continue }
        $wordList.Add(@{ Name = $f[0]; Word = $f[1]; Regex = ($f[2] -eq "true") })
    }
} else {
    $wordList.Add(@{ Name = "0 件"; Word = "tebunko計測用の存在しない語"; Regex = $false })
}

# 子のプロセスを起動する（失敗したら止める）
function invokeChild([string]$script, [string[]]$arguments) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path "$PSScriptRoot\perf" $script) @arguments
    if ($LASTEXITCODE) { throw "$script が失敗しました（終了コード $LASTEXITCODE）。" }
}
function readJson([string]$path) { [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json }

# 実行の情報
$run = [ordered]@{
    RunId = $RunId; Date = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"); Ref = $Ref; Sha = $Sha; Scale = $Scale; Label = $Label
    PowerShell = $PSVersionTable.PSVersion.ToString(); OS = [Environment]::OSVersion.VersionString; Cores = $cores
    Count = $Count; DataSeconds = $(if ($DataSeconds -ge 0) { [Math]::Round($DataSeconds, 1) } else { $null })
}
try { $run.Cpu = (@(Get-CimInstance Win32_Processor)[0].Name).Trim() } catch { $run.Cpu = "不明" }
try { $run.MemoryGB = [Math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 1) } catch { $run.MemoryGB = $null }
try { $run.DefenderRealtime = [string](Get-MpComputerStatus -ErrorAction Stop).RealTimeProtectionEnabled } catch { $run.DefenderRealtime = "不明" }

# 1. インデックス作成
Write-Host "インデックス作成を測ります。"
$indexJson = Join-Path $raw "index.json"
invokeChild "measure_index.ps1" @("-Tool", $Tool, "-Index", $Index, "-Work", $Work, "-Out", $indexJson, "-SampleMs", $SampleMs)
$indexRaw = readJson $indexJson
$indexSamples = @($indexRaw.Samples)
$indexResult = [ordered]@{ Pack = $indexRaw.Pack; Resources = (getPhaseResources $indexSamples $cores); PeakWorkingSetMB = $indexRaw.PeakWorkingSetMB }

# 2. 検索（語ごとに別のプロセス）
$search = New-Object System.Collections.Generic.List[object]
$searchRaw = New-Object System.Collections.Generic.List[object]
for ($w = 0; $w -lt $wordList.Count; $w++) {
    $spec = $wordList[$w]
    Write-Host "検索を測ります（$($spec.Name)）。"
    $wordFile = Join-Path $raw "word-$w.json"
    [System.IO.File]::WriteAllText($wordFile, ([pscustomobject]$spec | ConvertTo-Json), $utf8)
    $searchJson = Join-Path $raw "search-$w.json"
    invokeChild "measure_search.ps1" @("-Tool", $Tool, "-Index", $Index, "-WordFile", $wordFile, "-Count", $Count, "-Out", $searchJson, "-SampleMs", $SampleMs)
    $r = readJson $searchJson
    $searchRaw.Add($r)
    $rows = @($r.Searches)
    $n = [double[]]@($rows | ForEach-Object { $_.N })
    $ws = [double[]]@($rows | ForEach-Object { $_.WorkingSetMB })
    $slope = getSlope $n $ws
    $res = @(getPhaseResources @($r.Samples) $cores | Where-Object { $_.Phase -eq "検索" })
    $search.Add([ordered]@{
        Name = $r.Name; Regex = [bool]$r.Regex; Hits = $rows[0].Hits; Truncated = [bool]$rows[0].Truncated; Packs = $rows[0].Packs
        TotalMs = getStats ([double[]]@($rows | ForEach-Object { $_.TotalMs }))
        LoadMs = getStats ([double[]]@($rows | ForEach-Object { $_.LoadMs }))
        ListMs = getStats ([double[]]@($rows | ForEach-Object { $_.ListMs }))
        MatchMs = getStats ([double[]]@($rows | ForEach-Object { $_.MatchMs }))
        WorkingSetFirstMB = $ws[0]; WorkingSetLastMB = $ws[$ws.Count - 1]; WorkingSetMaxMB = ($ws | Measure-Object -Maximum).Maximum
        WorkingSetPer10MB = $(if ($null -ne $slope) { [Math]::Round($slope * 10, 2) } else { $null })
        HandlesFirst = $rows[0].Handles; HandlesLast = $rows[$rows.Count - 1].Handles
        ThreadsMax = ($rows | Measure-Object Threads -Maximum).Maximum
        CpuSeconds = $(if ($res.Count) { $res[0].CpuSeconds } else { $null })
        Gc0 = $(if ($res.Count) { $res[0].Gc0 } else { $null }); Gc1 = $(if ($res.Count) { $res[0].Gc1 } else { $null }); Gc2 = $(if ($res.Count) { $res[0].Gc2 } else { $null })
        PeakWorkingSetMB = $r.PeakWorkingSetMB
    })
}

# result.json（形式の版 1）
$result = [ordered]@{ Schema = 1; Run = $run; Index = $indexResult; Search = $search.ToArray() }
[System.IO.File]::WriteAllText((Join-Path $Out "result.json"), ($result | ConvertTo-Json -Depth 6), $utf8)

# metrics.csv（1 行 1 指標）
$metrics = New-Object System.Collections.Generic.List[object]
function addMetric([string]$metric, [string]$word, [string]$stat, $value, [string]$unit) {
    if ($null -eq $value) { return }
    $metrics.Add([pscustomobject][ordered]@{ run_id = $RunId; date = $run.Date; ref = $Ref; sha = $Sha; scale = $Scale
        metric = $metric; word = $word; stat = $stat; value = (formatPerfNumber ([double]$value)); unit = $unit })
}
addMetric "data_seconds" "" "value" $run.DataSeconds "s"
if ($indexResult.Pack) {
    addMetric "index_seconds" "" "value" $indexResult.Pack.Seconds "s"
    addMetric "index_books" "" "value" $indexResult.Pack.Books "count"
    addMetric "index_tsv_mb" "" "value" $indexResult.Pack.TsvMB "MB"
    addMetric "index_pack_mb" "" "value" $indexResult.Pack.PackMB "MB"
    addMetric "index_system_index_mb" "" "value" $indexResult.Pack.SystemIndexMB "MB"
}
$packPhase = @($indexResult.Resources | Where-Object { $_.Phase -eq "pack の作成" })
if ($packPhase.Count) {
    addMetric "index_cpu_seconds" "" "value" $packPhase[0].CpuSeconds "s"
    addMetric "index_cpu_percent" "" "value" $packPhase[0].CpuPercent "%"
    addMetric "index_working_set_mb" "" "max" $packPhase[0].WorkingSetMaxMB "MB"
    addMetric "index_managed_mb" "" "max" $packPhase[0].ManagedMaxMB "MB"
    addMetric "index_gc2" "" "value" $packPhase[0].Gc2 "count"
}
foreach ($s in $search) {
    foreach ($part in @(@("search_total_ms", $s.TotalMs), @("search_load_ms", $s.LoadMs), @("search_list_ms", $s.ListMs), @("search_match_ms", $s.MatchMs))) {
        foreach ($stat in @("Min", "Median", "Mean", "Max")) { addMetric $part[0] $s.Name $stat.ToLower() $part[1][$stat] "ms" }
    }
    addMetric "search_hits" $s.Name "value" $s.Hits "count"
    addMetric "search_working_set_mb" $s.Name "max" $s.WorkingSetMaxMB "MB"
    addMetric "search_working_set_per10_mb" $s.Name "value" $s.WorkingSetPer10MB "MB"
    addMetric "search_handles" $s.Name "last" $s.HandlesLast "count"
    addMetric "search_cpu_seconds" $s.Name "value" $s.CpuSeconds "s"
    addMetric "search_gc2" $s.Name "value" $s.Gc2 "count"
}
[System.IO.File]::WriteAllLines((Join-Path $Out "metrics.csv"), [string[]]@($metrics | ConvertTo-Csv -NoTypeInformation), $utf8Bom)

# searches.csv・resource-*.csv
$searchRows = foreach ($r in $searchRaw) { foreach ($row in @($r.Searches)) { $o = [ordered]@{ Word = $r.Name }; foreach ($p in $row.PSObject.Properties) { $o[$p.Name] = $p.Value }; [pscustomobject]$o } }
[System.IO.File]::WriteAllLines((Join-Path $Out "searches.csv"), [string[]]@($searchRows | ConvertTo-Csv -NoTypeInformation), $utf8Bom)
[System.IO.File]::WriteAllLines((Join-Path $Out "resource-index.csv"), [string[]]@($indexSamples | ConvertTo-Csv -NoTypeInformation), $utf8Bom)
$searchSamples = foreach ($r in $searchRaw) { foreach ($sm in @($r.Samples)) { $o = [ordered]@{ Word = $r.Name }; foreach ($p in $sm.PSObject.Properties) { $o[$p.Name] = $p.Value }; [pscustomobject]$o } }
[System.IO.File]::WriteAllLines((Join-Path $Out "resource-search.csv"), [string[]]@($searchSamples | ConvertTo-Csv -NoTypeInformation), $utf8Bom)

# summary.md
$palette = @("#e41a1c", "#377eb8", "#4daf4a", "#984ea3", "#ff7f00", "#a65628")
$colorNames = @("赤", "青", "緑", "紫", "橙", "茶")
$md = New-Object System.Collections.Generic.List[string]
function add([string]$text = "") { $md.Add($text) }
function ms($stats, [string]$key) { if ($null -eq $stats) { return "–" }; return ("{0:N0}" -f $stats[$key]) }
add "# tebunko の性能$(if ($Label) { "（$Label）" })"
add
add "計測日時: $($run.Date)。結果には数字だけを書き、パスやファイル名は書かない。"
add
add "## 環境"
add
add "| PowerShell | OS | CPU | 論理コア | メモリ | Defender のリアルタイム保護 | OS のファイルキャッシュ |"
add "|---|---|---|---|---|---|---|"
add "| $($run.PowerShell) | $($run.OS) | $($run.Cpu) | $($run.Cores) | $($run.MemoryGB) GB | $($run.DefenderRealtime) | 空にしていない |"
if ($null -ne $run.DataSeconds) {
    add
    add ("データ（TSV）の生成に {0:N1} 秒かかった（tebunko-perfdata の new_index.ps1 による。tebunko の処理時間には含まない）。" -f $run.DataSeconds)
}

add
add "## インデックス作成"
add
if ($indexResult.Pack) {
    $pk = $indexResult.Pack
    add "インデックス作成のうち、Office からの取り込みを除いた後半の処理（pack の書き出し、TSV の削除、システムインデックスの作成）。インデクサと同じ関数を、別のプロセスで動かして測った。"
    add
    add "| フォルダ | ブック | TSV | TSV の合計 | 時間 | pack | pack の合計 | システムインデックス |"
    add "|---|---|---|---|---|---|---|---|"
    add ("| {0:N0} | {1:N0} | {2:N0} | {3:N1} MB | {4:N1} 秒 | {5:N0} | {6:N1} MB | {7:N1} MB |" -f $pk.Folders, $pk.Books, $pk.Tsv, $pk.TsvMB, $pk.Seconds, $pk.Packs, $pk.PackMB, $pk.SystemIndexMB)
} else {
    add "渡したインデックスに TSV が無いため、測っていない。"
}
add
add "「CPU」はこのプロセスの使用率（全コアを使い切ると 100%）、「PC の CPU」は PC 全体の使用率の平均。ワーキングセットのピークは $($indexResult.PeakWorkingSetMB) MB。"
add
add "| 段階 | 時間 | CPU 時間 | CPU | PC の CPU | ワーキングセット | プライベート | マネージドヒープ | スレッド数 | ハンドル数 | GC 回数（0/1/2 世代） |"
add "|---|---|---|---|---|---|---|---|---|---|---|"
foreach ($p in $indexResult.Resources) {
    add ("| {0} | {1:N1} 秒 | {2:N1} 秒 | {3}% | {4} | {5:N0} MB | {6:N0} MB | {7:N0} MB | {8} | {9} | {10}/{11}/{12} |" -f $p.Phase, $p.Seconds, $p.CpuSeconds, $p.CpuPercent, $(if ($null -ne $p.PcCpuPercent) { "$($p.PcCpuPercent)%" } else { "–" }), $p.WorkingSetMaxMB, $p.PrivateMaxMB, $p.ManagedMaxMB, $p.ThreadsMax, $p.HandlesMax, $p.Gc0, $p.Gc1, $p.Gc2)
}
if ($indexSamples.Count -ge 2) {
    # 記録はほぼ等間隔なので、横軸は秒の範囲にして目盛りは Mermaid に任せる
    $points = thinOut $indexSamples 100
    $lastSecond = [Math]::Ceiling($points[$points.Count - 1].Ms / 1000)
    $cpu = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i -lt $points.Count; $i++) {
        if ($i -eq 0) { $cpu.Add(0); continue }
        $dt = [Math]::Max(1, $points[$i].Ms - $points[$i - 1].Ms)
        $cpu.Add([Math]::Round(($points[$i].CpuMs - $points[$i - 1].CpuMs) * 100 / $dt / $cores, 0))
    }
    add
    add ("各段階の開始時刻（秒）: " + ((@($indexSamples | Group-Object Phase | Where-Object { $_.Name }) | ForEach-Object { "{0} {1:N1}" -f $_.Name, ($_.Group[0].Ms / 1000) }) -join "、"))
    add
    add "赤の線がワーキングセット、青の線がマネージドヒープ。"
    add
    foreach ($l in (newLineChart "インデックス作成のメモリ" "秒" "0 --> $lastSecond" "MB" @(@($points | ForEach-Object { $_.WorkingSetMB }), @($points | ForEach-Object { $_.ManagedMB })) $palette[0..1])) { add $l }
    add
    foreach ($l in (newLineChart "インデックス作成の CPU 使用率（このプロセス）" "秒" "0 --> $lastSecond" "%" @(, $cpu.ToArray()) $palette[0..0])) { add $l }
}

add
add "## 検索"
add
add "語ごとに新しいプロセスを起動し、同じプロセスで $Count 回続けて検索した（画面の検索と同じ流れ: 新しいスレッドで lib.ps1 を読み込む → 集約ファイルを列挙する → 照合する。上限 1 万件）。1 回目は画面を開き直した直後の検索に当たり、たいてい最大の値になる。"
add
add "| 語 | ヒット件数 | 最小 | 中央値 | 平均 | 最大 | lib.ps1 の読み込み（中央値／最大） | 列挙（中央値／最大） | 照合（中央値／最大） |"
add "|---|---|---|---|---|---|---|---|---|"
foreach ($s in $search) {
    add ("| {0} | {1}{2} | {3} ms | {4} ms | {5} ms | {6} ms | {7}／{8} ms | {9}／{10} ms | {11}／{12} ms |" -f $s.Name, $s.Hits, $(if ($s.Truncated) { "+" } else { "" }),
        (ms $s.TotalMs "Min"), (ms $s.TotalMs "Median"), (ms $s.TotalMs "Mean"), (ms $s.TotalMs "Max"),
        (ms $s.LoadMs "Median"), (ms $s.LoadMs "Max"), (ms $s.ListMs "Median"), (ms $s.ListMs "Max"), (ms $s.MatchMs "Median"), (ms $s.MatchMs "Max"))
}
add
add "### 検索のリソース"
add
add "語ごとのプロセスの値。「10 回あたりの増え方」は、検索 1 回ごとのワーキングセットを最小二乗法で直線に当てはめた傾き。検索のキャッシュは上限（約 128MB）まで増えて止まるのが正しい動きで、それを超えて増え続けるときはリークを疑う。"
add
add "| 語 | ワーキングセット（最初／最後／最大） | 10 回あたりの増え方 | ハンドル数（最初／最後） | スレッド数の最大 | CPU 時間 | GC 回数（0/1/2 世代） |"
add "|---|---|---|---|---|---|---|"
foreach ($s in $search) {
    add ("| {0} | {1:N0}／{2:N0}／{3:N0} MB | {4} MB | {5}／{6} | {7} | {8} 秒 | {9}/{10}/{11} |" -f $s.Name, $s.WorkingSetFirstMB, $s.WorkingSetLastMB, $s.WorkingSetMaxMB,
        $(if ($null -ne $s.WorkingSetPer10MB) { "{0:N1}" -f $s.WorkingSetPer10MB } else { "–" }), $s.HandlesFirst, $s.HandlesLast, $s.ThreadsMax, $s.CpuSeconds, $s.Gc0, $s.Gc1, $s.Gc2)
}
if ($Count -ge 2) {
    $legend = (0..($search.Count - 1) | ForEach-Object { "$($colorNames[$_ % $colorNames.Count]): $($search[$_].Name)" }) -join "、"
    $xs = "[" + ((1..$Count) -join ", ") + "]"
    $colors = @(0..($search.Count - 1) | ForEach-Object { $palette[$_ % $palette.Count] })
    add
    add "線の色: $legend。横軸は検索の回数。"
    add
    foreach ($l in (newLineChart "検索 1 回ごとの時間" "回" $xs "ms" @($searchRaw | ForEach-Object { , @(@($_.Searches) | ForEach-Object { $_.TotalMs }) }) $colors)) { add $l }
    add
    foreach ($l in (newLineChart "検索 1 回ごとのワーキングセット" "回" $xs "MB" @($searchRaw | ForEach-Object { , @(@($_.Searches) | ForEach-Object { $_.WorkingSetMB }) }) $colors)) { add $l }
}
[System.IO.File]::WriteAllLines((Join-Path $Out "summary.md"), [string[]]$md, $utf8)
Write-Host "結果を書きました: summary.md・result.json・metrics.csv・searches.csv・resource-index.csv・resource-search.csv"
