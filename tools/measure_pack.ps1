# まとめファイル（本文.<拡張子>.tsv）の検索の速さを測る。結果には時間と件数だけを出す。
#
#   .\tools\measure_pack.ps1 -IndexRoot <index のフォルダ> -Folders poc\部署0,poc\部署1 -Words 見積
#   -Folders   … 測るフォルダ（index からの相対パス。省略するとすべて）
#   -OutFile   … 結果を書くファイル
param (
    [Parameter(Mandatory = $true)][string]$IndexRoot,
    [string[]]$Folders = @(""),
    [string[]]$Words = @(),
    [string]$OutFile
)

$ErrorActionPreference = "Stop"
$lib = Join-Path (Split-Path $PSScriptRoot -Parent) "scripts\tebunko_grep\lib.ps1"
. $lib
$IndexRoot = (Resolve-Path -LiteralPath $IndexRoot).ProviderPath.TrimEnd("\")
$Words = @("tebunko計測用の存在しない語") + @($Words | Where-Object { $_ })

$report = New-Object System.Collections.Generic.List[string]
function out([string]$text) { $report.Add($text); Write-Host $text }
$mp = New-Object System.Diagnostics.PerformanceCounter "Process", "% Processor Time", "MsMpEng"

out "# まとめファイルの計測 $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"
out ""
out "## 1. 環境"
$os = Get-CimInstance Win32_OperatingSystem
out ("PowerShell {0} / 論理コア {1} / 物理メモリ {2:N1} GB（空き {3:N1} GB）" -f $PSVersionTable.PSVersion, [Environment]::ProcessorCount, ($os.TotalVisibleMemorySize / 1MB), ($os.FreePhysicalMemory / 1MB))

# 測る範囲（画面の検索対象ツリーで、フォルダ以下すべてを選んだのと同じ）
$targets = @($Folders | ForEach-Object { @{ Root = $IndexRoot; RelPath = $_; Recurse = $true } })
$packs = (getIndexPackFiles $targets).Packs
$bytes = 0L; $maxBytes = 0L
foreach ($p in $packs) { $bytes += $p.Size; $maxBytes = [Math]::Max($maxBytes, $p.Size) }
out ""
out "## 2. まとめファイル"
out ("まとめファイル {0:N0} 件 / 合計 {1:N1} MB / 平均 {2:N2} MB / 最大 {3:N2} MB" -f $packs.Count, ($bytes / 1MB), ($bytes / [Math]::Max(1, $packs.Count) / 1MB), ($maxBytes / 1MB))

# 画面の検索スレッドと同じ流れ（新しいスレッドで lib.ps1 を読み込み、列挙して検索する）。読んだ内容の入れ物は画面と同じく持ち続ける
$searchScript = {
    param ($lib, $targets, $word, $cache)
    $total = [System.Diagnostics.Stopwatch]::StartNew(); $step = [System.Diagnostics.Stopwatch]::StartNew()
    . $lib
    $load = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $index = getIndexPackFiles $targets
    $list = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $r = searchPackIndex $word $index.Packs $true 10000 -cache $cache
    @{ Load = $load; List = $list; Search = $step.Elapsed.TotalMilliseconds; Total = $total.Elapsed.TotalMilliseconds; Files = $index.Packs.Count; Hits = $r.Hits.Count; Truncated = $r.Truncated }
}

out ""
out "## 3. 画面と同じ検索（スレッドの用意から検索の終わりまで。画面への表示は含まない）"
out "| 回 | ワード | まとめファイル | ヒット | 合計 | lib.ps1 | 列挙 | 検索 | MsMpEng CPU |"
out "|---|---|---|---|---|---|---|---|---|"
$cache = newTsvTextCache
foreach ($round in 1..2) {
    foreach ($word in $Words) {
        $label = if ($word -eq $Words[0]) { "（存在しない語）" } else { "ワード $([array]::IndexOf($Words, $word))" }
        $ps = [powershell]::Create()
        try {
            [void]$ps.AddScript($searchScript.ToString()).AddArgument($lib).AddArgument($targets).AddArgument($word).AddArgument($cache)
            $m0 = $mp.NextSample().RawValue
            $r = $ps.Invoke()[0]
            $mpMs = ($mp.NextSample().RawValue - $m0) / 10000
        } finally {
            $ps.Dispose()
        }
        $hits = if ($r.Truncated) { "$($r.Hits)+" } else { "$($r.Hits)" }
        out ("| {0} | {1} | {2:N0} | {3} | {4:N0} ms | {5:N0} ms | {6:N0} ms | {7:N0} ms | {8:N0} ms |" -f $round, $label, $r.Files, $hits, $r.Total, $r.Load, $r.List, $r.Search, $mpMs)
    }
}

if ($OutFile) {
    [System.IO.File]::WriteAllLines($OutFile, $report, [System.Text.UTF8Encoding]::new($true))
}
