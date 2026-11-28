# 検索のヒットと同じ形の PSCustomObject を、件数ぶん List[psobject] に作る時間を測る（perf.yml の入力 scattered が 1 のとき。手元でも使える）。
# 照合（searchContentIndexFiles）がヒット 1 件ごとに作るのと同じ 8 項目（Root・RelPath・RelDir・FileName・Book・Location・LineNumber・Line）にする。
# ヒットを場所ごとのまとまりで返す作り（#290 の子 2〜4）の許す幅の式で、「1 万件分の PSCustomObject を作る時間」に使う。
# tebunko の版によらないので、tebunko の lib は読まない。新しいプロセスで起動する（同じプロセスで前に作ったものの GC を混ぜないため）。
#
#   .\tools\perf\measure_hit_objects.ps1 -Out <結果のフォルダ> [-Hits 10000] [-Repeat 20]
#   -Hits   … 1 回に作る件数（既定 10000。検索の上限と同じ）
#   -Repeat … 測る回数（最小・中央値・平均・最大を出す）
#   -Out    … hit_objects.json・hit_objects.md を書くフォルダ。ほかには書かない
param (
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Hits = 10000,
    [int]$Repeat = 20
)

$ErrorActionPreference = "Stop"
if ($Hits -lt 1 -or $Repeat -lt 1) { throw "-Hits・-Repeat は 1 以上にしてください。" }
. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))
$Out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)
[void][System.IO.Directory]::CreateDirectory($Out)

function newHitObjects([int]$count) {
    # 照合と同じ形のヒットを count 件作る。値は照合が持つ文字列・数と同じ型にする
    $root = "C:\計測\本文インデックス"
    $relPath = "散在\content_index.docx.001.tsv"
    $line = "1-001-1`t散在計測語`t本文"
    $list = New-Object System.Collections.Generic.List[psobject]
    for ($i = 0; $i -lt $count; $i++) {
        $list.Add([pscustomobject]@{
            Root = $root; RelPath = $relPath; RelDir = "散在"; FileName = "資料01.docx"
            Book = "資料01.docx"; Location = "ページ001"; LineNumber = $i + 1; Line = $line
        })
    }
    return , $list
}

$ms = New-Object System.Collections.Generic.List[double]
$gcBefore = @(0..2 | ForEach-Object { [GC]::CollectionCount($_) })
$created = 0
for ($r = 0; $r -lt $Repeat; $r++) {
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $list = newHitObjects $Hits
    $ms.Add([Math]::Round($watch.Elapsed.TotalMilliseconds, 1))
    $created = $list.Count
    $properties = @($list[0].PSObject.Properties | ForEach-Object { $_.Name })
    $list = $null
}
$gc = @(0..2 | ForEach-Object { [GC]::CollectionCount($_) - $gcBefore[$_] })
$stats = getStats $ms.ToArray()
$result = [ordered]@{
    Hits = $Hits; Repeat = $Repeat; Created = $created; Properties = $properties; Ms = $stats; Runs = $ms.ToArray()
    Gc0 = $gc[0]; Gc1 = $gc[1]; Gc2 = $gc[2]
    PowerShell = $PSVersionTable.PSVersion.ToString(); Cores = [Environment]::ProcessorCount
}
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $Out "hit_objects.json"), ($result | ConvertTo-Json -Depth 4), $utf8)

$md = New-Object System.Collections.Generic.List[string]
$md.Add("## PSCustomObject 1 万件を作る時間")
$md.Add("")
$md.Add("検索のヒットと同じ 8 項目（$($properties -join '・')）の PSCustomObject を、$Hits 件 List[psobject] に作る時間を、新しいプロセスで $Repeat 回測った（tebunko の版によらない。ヒットを場所ごとのまとまりで返す作りの、許す幅の式に使う）。GC 回数は $Repeat 回の合計。")
$md.Add("")
$md.Add("| 件数 | 回数 | 最小 | 中央値 | 平均 | 最大 | GC 回数（0/1/2 世代） |")
$md.Add("|---|---|---|---|---|---|---|")
$md.Add(("| {0} | {1} | {2:N1} ms | {3:N1} ms | {4:N1} ms | {5:N1} ms | {6}/{7}/{8} |" -f $Hits, $Repeat, $stats.Min, $stats.Median, $stats.Mean, $stats.Max, $gc[0], $gc[1], $gc[2]))
[System.IO.File]::WriteAllText((Join-Path $Out "hit_objects.md"), (($md -join "`n") + "`n"), $utf8)
Write-Information ("PSCustomObject {0:N0} 件を作る時間: 中央値 {1:N1} ms（{2} 回）" -f $Hits, $stats.Median, $Repeat) -InformationAction Continue
