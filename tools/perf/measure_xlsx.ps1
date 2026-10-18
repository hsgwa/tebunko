# Excel のブックを、Office を使わずに ZIP から直接読む部分（readXlsxObjectUnits。シートの流れ読み）の速さを測る。
# tools\measure_perf.ps1 が別のプロセスで起動する。測るブックは、その場で作る（tools\perf\xlsx_data.ps1。Office は使わない）。
#   普通のブック … シート 1 枚。-Repeat 回続けて読み、1 回ごとの時間を出す（比べるのは中央値）
#   同じ部品を指すブック … <sheet> を -ManySheets 個並べ、すべて同じシートの部品を指す。1 回だけ読む
# 結果（時間・読んだ件数・リソースの記録）を -Out の JSON に書く。件数は、測る版どうしで同じ結果を返したかの確かめに使う。
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Repeat = 5,
    [int]$Rows = 200000,
    [int]$ManySheets = 200,
    [int]$ManyRows = 50000,
    [int]$SampleMs = 200
)

$ErrorActionPreference = "Stop"
. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))
. ([WildcardPattern]::Escape("$PSScriptRoot\xlsx_data.ps1".Replace('`', '``')))
# readXlsxObjectUnits は、インデックス作成の読み込み口（indexer\indexer_lib.ps1。lib.ps1 も読み込む）が読み込む
$indexerLib = Join-Path (Split-Path -Parent (resolveTebunkoLib $Tool)) "indexer\indexer_lib.ps1"
if (!(Test-Path -LiteralPath $indexerLib)) { throw "測る tebunko に indexer\indexer_lib.ps1 がありません。" }
. ([WildcardPattern]::Escape($indexerLib.Replace('`', '``')))
assertTebunkoFunctions @("readXlsxObjectUnits")

function measureXlsxRead([string]$path) {
    # 1 回読み、時間・場所の数・行の数・読めなかった部品の数を返す
    $failures = New-Object System.Collections.Generic.List[string]
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $units = readXlsxObjectUnits $path $failures
    $ms = $watch.Elapsed.TotalMilliseconds
    $lines = 0
    foreach ($key in @($units.Keys)) { $lines += @($units[$key]).Count }
    return [ordered]@{ Ms = [Math]::Round($ms, 1); Units = $units.Count; Lines = $lines; Failures = $failures.Count }
}

$dir = Join-Path $Work "xlsx"
[void][System.IO.Directory]::CreateDirectory($dir)
$normalPath = Join-Path $dir "normal.xlsx"
$manyPath = Join-Path $dir "many.xlsx"
Write-Host "ブックを作ります（シート 1 枚・$Rows 行／シート $ManySheets 個・$ManyRows 行）。"
newXlsxPerfBook $normalPath 1 $Rows
# 読む量の見積もり: 部品 1 つは 1 行約 120 バイトで、既定（50000 行）だと約 6 MB。<sheet> ごとに同じ部品を読む版は、200 個で約 1.2 GB を読む（流れ読みの上限 1000 MB を超える）。同じ部品は 1 回だけ読む版は、約 6 MB のまま
newXlsxPerfBook $manyPath $ManySheets $ManyRows

$monitor = startResourceMonitor $SampleMs
$result = [ordered]@{ Repeat = $Repeat; Rows = $Rows; ManySheets = $ManySheets; ManyRows = $ManyRows
    NormalMB = [Math]::Round(([System.IO.FileInfo]$normalPath).Length / 1MB, 2); ManyMB = [Math]::Round(([System.IO.FileInfo]$manyPath).Length / 1MB, 2); Normal = @(); Many = $null }
try {
    setMonitorPhase $monitor "普通のブック"
    $runs = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $Repeat; $i++) {
        $r = measureXlsxRead $normalPath
        Write-Host ("  普通のブック {0} 回目: {1:N0} ms" -f $i, $r.Ms)
        $runs.Add($r)
    }
    $result.Normal = $runs.ToArray()
    setMonitorPhase $monitor "同じ部品を指すブック"
    $result.Many = measureXlsxRead $manyPath
    Write-Host ("  同じ部品を指すブック: {0:N0} ms" -f $result.Many.Ms)
} finally {
    $samples = stopResourceMonitor $monitor
}
$result.Samples = $samples
$result.PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
[System.IO.File]::WriteAllText($Out, ($result | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
