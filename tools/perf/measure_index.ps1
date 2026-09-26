# インデックス作成（pack の作成）を測る。tools\measure_perf.ps1 が別のプロセスで起動する。
# インデクサと同じ publishIndexFolders で、TSV があるフォルダをまとめて pack にし、システムインデックスを作る
# （Office からの取り込みを除いた、インデックス作成の後半）。TSV は pack に変換され、元の TSV は削除される。
# 結果（件数・大きさ・時間・リソースの記録）を -Out の JSON に書く。
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Index,
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$SampleMs = 200
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\perf_common.ps1"
. (resolveTebunkoLib $Tool)
assertTebunkoFunctions @("findIndexFoldersWithBooks", "publishIndexFolders", "getIndexPackFiles", "testIndexBookDir")

$monitor = startResourceMonitor $SampleMs
$result = [ordered]@{ Pack = $null }
try {
    setMonitorPhase $monitor "準備"
    $folders = [string[]](findIndexFoldersWithBooks $Index)
    if ($folders.Count -gt 0) {
        $tsvFiles = [System.IO.Directory]::GetFiles($Index, "*.tsv", "AllDirectories")
        $tsvBytes = 0L
        foreach ($f in $tsvFiles) { $tsvBytes += ([System.IO.FileInfo]$f).Length }
        $bookCount = [System.IO.Directory]::GetDirectories($Index, "*.*", "AllDirectories").Where({ testIndexBookDir $_ }).Count
        Write-Host ("pack を作ります（フォルダ {0:N0}・ブック {1:N0}・TSV {2:N0} 件・{3:N0} MB）…" -f $folders.Count, $bookCount, $tsvFiles.Count, ($tsvBytes / 1MB))
        $pending = @{}
        foreach ($f in $folders) { $pending[$f] = [string[]]@() }

        setMonitorPhase $monitor "pack の作成"
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        [void](publishIndexFolders $pending $Index (Join-Path $Work "system_index") (Join-Path $Work "システムインデックスの状態.tsv"))
        $seconds = $watch.Elapsed.TotalSeconds

        setMonitorPhase $monitor "集計"
        $packs = (getIndexPackFiles @($Index)).Packs
        $packBytes = 0L
        foreach ($p in $packs) { $packBytes += $p.Size }
        $systemBytes = 0L
        $systemRoot = Join-Path $Work "system_index"
        if ([System.IO.Directory]::Exists($systemRoot)) {
            foreach ($f in [System.IO.Directory]::GetFiles($systemRoot, "*", "AllDirectories")) { $systemBytes += ([System.IO.FileInfo]$f).Length }
        }
        $result.Pack = [ordered]@{
            Folders = $folders.Count; Books = $bookCount; Tsv = $tsvFiles.Count; TsvMB = [Math]::Round($tsvBytes / 1MB, 1)
            Seconds = [Math]::Round($seconds, 2); Packs = $packs.Count; PackMB = [Math]::Round($packBytes / 1MB, 1)
            SystemIndexMB = [Math]::Round($systemBytes / 1MB, 1)
        }
        Write-Host ("  {0:N1} 秒" -f $seconds)
    } else {
        Write-Host "TSV が無い（pack だけの）インデックスなので、pack の作成は測りません。"
    }
} finally {
    $samples = stopResourceMonitor $monitor
}
$result.Samples = $samples
$result.PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
[System.IO.File]::WriteAllText($Out, ($result | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
