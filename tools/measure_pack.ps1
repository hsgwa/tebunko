# 検索用のまとめファイルの速さを、今の TSV の検索と比べて測る（PoC 用）。結果には時間と件数だけを出す。
#
#   .\tools\measure_pack.ps1 -TsvRoot <元の index> -PackRoot <変換先の index> -Folders poc\部署0,poc\部署1 -Words 侔了
#   -Folders   … 測るフォルダ（index からの相対パス。省略するとすべて）
#   -SkipTsv   … TSV の検索を測らない（インデックスが大きく、TSV の検索が終わらないとき）
#   -OutFile   … 結果を書くファイル
param (
    [Parameter(Mandatory = $true)][string]$TsvRoot,
    [Parameter(Mandatory = $true)][string]$PackRoot,
    [string[]]$Folders = @(""),
    [string[]]$Words = @(),
    [switch]$SkipTsv,
    [string]$OutFile
)

$ErrorActionPreference = "Stop"
$lib = Join-Path (Split-Path $PSScriptRoot -Parent) "scripts\tebunko_grep\lib.ps1"
. $lib
$TsvRoot = (Resolve-Path -LiteralPath $TsvRoot).ProviderPath.TrimEnd("\")
$PackRoot = (Resolve-Path -LiteralPath $PackRoot).ProviderPath.TrimEnd("\")
$Words = @("tebunko計測用の存在しない語") + @($Words | Where-Object { $_ })

$report = New-Object System.Collections.Generic.List[string]
function out([string]$text) { $report.Add($text); Write-Host $text }
function toKeys($hits) { @($hits | ForEach-Object { "{0}|{1}|{2}|{3}|{4}" -f $_.RelDir, $_.Book, $_.Location, $_.LineNumber, $_.Line }) }
$mp = New-Object System.Diagnostics.PerformanceCounter "Process", "% Processor Time", "MsMpEng"

out "# まとめファイルの計測 $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"
out ""
out "## 1. 環境"
$os = Get-CimInstance Win32_OperatingSystem
out ("PowerShell {0} / 論理コア {1} / 物理メモリ {2:N1} GB（空き {3:N1} GB）" -f $PSVersionTable.PSVersion, [Environment]::ProcessorCount, ($os.TotalVisibleMemorySize / 1MB), ($os.FreePhysicalMemory / 1MB))

# 測る範囲（画面の検索対象ツリーで、フォルダ以下すべてを選んだのと同じ）
$tsvTargets = @($Folders | ForEach-Object { @{ Root = $TsvRoot; RelPath = $_; Recurse = $true } })
$packs = New-Object System.Collections.Generic.List[hashtable]
foreach ($f in $Folders) { $packs.AddRange([hashtable[]](getPackFiles $PackRoot $f $true)) }
$bytes = 0L; $maxBytes = 0L
foreach ($p in $packs) { $bytes += $p.Size; $maxBytes = [Math]::Max($maxBytes, $p.Size) }
out ""
out "## 2. まとめファイル"
out ("まとめファイル {0:N0} 件 / 合計 {1:N1} MB / 平均 {2:N2} MB / 最大 {3:N2} MB" -f $packs.Count, ($bytes / 1MB), ($bytes / [Math]::Max(1, $packs.Count) / 1MB), ($maxBytes / 1MB))

# 画面の検索スレッドと同じ流れ（新しいスレッドで lib.ps1 を読み込み、列挙して検索する）。読んだ内容の入れ物は画面と同じく持ち続ける
$tsvScript = {
    param ($lib, $targets, $word, $cache)
    $total = [System.Diagnostics.Stopwatch]::StartNew(); $step = [System.Diagnostics.Stopwatch]::StartNew()
    . $lib
    $load = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $index = getIndexTsvFiles $targets
    $list = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $r = searchIndex $word $index.Files $true 10000 50 -cache $cache
    @{ Load = $load; List = $list; Search = $step.Elapsed.TotalMilliseconds; Total = $total.Elapsed.TotalMilliseconds; Files = $index.Files.Count; Hits = $r.Hits; Truncated = $r.Truncated }
}
$packScript = {
    param ($lib, $root, $folders, $word, $cache)
    $total = [System.Diagnostics.Stopwatch]::StartNew(); $step = [System.Diagnostics.Stopwatch]::StartNew()
    . $lib
    $load = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $packs = New-Object System.Collections.Generic.List[hashtable]
    foreach ($f in $folders) { $packs.AddRange([hashtable[]](getPackFiles $root $f $true)) }
    $list = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $r = searchPackIndex $word $packs $true 10000 -cache $cache
    @{ Load = $load; List = $list; Search = $step.Elapsed.TotalMilliseconds; Total = $total.Elapsed.TotalMilliseconds; Files = $packs.Count; Hits = $r.Hits; Truncated = $r.Truncated }
}
function invokeRound($script, [object[]]$arguments) {
    $ps = [powershell]::Create()
    try {
        [void]$ps.AddScript($script.ToString())
        foreach ($a in $arguments) { [void]$ps.AddArgument($a) }
        $m0 = $mp.NextSample().RawValue
        $r = $ps.Invoke()[0]
        $r.MpMs = ($mp.NextSample().RawValue - $m0) / 10000
        return $r
    } finally {
        $ps.Dispose()
    }
}

out ""
out "## 3. 画面と同じ検索（スレッドの用意から検索の終わりまで。画面への表示は含まない）"
out "| 方式 | 回 | ワード | ファイル | ヒット | 合計 | lib.ps1 | 列挙 | 検索 | MsMpEng CPU | TSV と同じ結果 |"
out "|---|---|---|---|---|---|---|---|---|---|---|"
$tsvCache = newTsvTextCache
$packCache = newTsvTextCache ([long]([Math]::Floor($os.TotalVisibleMemorySize * 1KB / 16 / 2)))
$tsvKeys = @{}
foreach ($round in 1..2) {
    foreach ($word in $Words) {
        $label = if ($word -eq $Words[0]) { "（存在しない語）" } else { "ワード $([array]::IndexOf($Words, $word))" }
        $rows = @()
        if (!$SkipTsv) {
            $r = invokeRound $tsvScript @($lib, $tsvTargets, $word, $tsvCache)
            $tsvKeys[$word] = (toKeys $r.Hits | Sort-Object) -join "`n"
            $rows += , @("TSV", $r, "—")
        }
        $r = invokeRound $packScript @($lib, $PackRoot, $Folders, $word, $packCache)
        $same = if ($SkipTsv) { "—" } elseif (((toKeys $r.Hits | Sort-Object) -join "`n") -ceq $tsvKeys[$word]) { "同じ" } else { "**違う**" }
        $rows += , @("まとめ", $r, $same)
        foreach ($row in $rows) {
            $x = $row[1]
            $hits = if ($x.Truncated) { "$($x.Hits.Count)+" } else { "$($x.Hits.Count)" }
            out ("| {0} | {1} | {2} | {3:N0} | {4} | {5:N0} ms | {6:N0} ms | {7:N0} ms | {8:N0} ms | {9:N0} ms | {10} |" -f $row[0], $round, $label, $x.Files, $hits, $x.Total, $x.Load, $x.List, $x.Search, $x.MpMs, $row[2])
        }
    }
}

out ""
out "## 4. 書き直し（まとめファイル 1 件を、同じ中身で書き直す）"
$largest = @($packs | Sort-Object { $_.Size } -Descending)[0]
foreach ($p in @($packs[0], $largest)) {
    $text = readPackText $p.Path
    $m0 = $mp.NextSample().RawValue
    $w = [System.Diagnostics.Stopwatch]::StartNew()
    writePackFile (fromLongPath $p.Path) $text
    $ms = $w.Elapsed.TotalMilliseconds
    Start-Sleep -Milliseconds 1500
    out ("{0:N2} MB: {1:N0} ms / MsMpEng CPU {2:N0} ms" -f ($p.Size / 1MB), $ms, (($mp.NextSample().RawValue - $m0) / 10000))
}

if ($OutFile) {
    [System.IO.File]::WriteAllLines($OutFile, $report, [System.Text.UTF8Encoding]::new($true))
}
