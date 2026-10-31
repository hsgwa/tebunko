# インデックス作成（pack の作成）を測る。tools\measure_perf.ps1 が別のプロセスで起動する。
# インデクサと同じ publishIndexFolders で、TSV があるフォルダをまとめて pack にする
# （Office からの取り込みを除いた、インデックス作成の後半）。TSV は pack に変換され、元の TSV は削除される。
# pack の作成のあと、検索が使う gram 索引を作る（updateGramIndexes。-Index が <W>\store\content\<名前> の形のときだけ。<W>\store\gram に作る）。
# 測る版に updateGramIndexes が無いか、置き場の形が違うときは作らず、Gram.Skipped に理由を書く。
# 結果（件数・大きさ・時間・リソースの記録）を -Out の JSON に書く。
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Index,
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$SampleMs = 200
)

$ErrorActionPreference = "Stop"
. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))
. ([WildcardPattern]::Escape((resolveTebunkoLib $Tool).Replace('`', '``')))
assertTebunkoFunctions @("findIndexFoldersWithBooks", "publishIndexFolders", "testIndexBookDir")
# 本文インデックスのファイルを列挙する関数は、名前を改めた後が getContentIndexFiles、前が getIndexPackFiles
$getFiles = resolveTebunkoFunction @("getContentIndexFiles", "getIndexPackFiles")

$monitor = startResourceMonitor $SampleMs
$result = [ordered]@{ Pack = $null; Gram = $null }
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
        [void](publishIndexFolders $pending $Index)
        $seconds = $watch.Elapsed.TotalSeconds

        setMonitorPhase $monitor "集計"
        $packs = @(getTebunkoProperty (& $getFiles @($Index)) @("ContentIndexFiles", "Packs"))
        $packBytes = 0L
        foreach ($p in $packs) { $packBytes += $p.Size }
        $result.Pack = [ordered]@{
            Folders = $folders.Count; Books = $bookCount; Tsv = $tsvFiles.Count; TsvMB = [Math]::Round($tsvBytes / 1MB, 1)
            Seconds = [Math]::Round($seconds, 2); Packs = $packs.Count; PackMB = [Math]::Round($packBytes / 1MB, 1)
        }
        Write-Host ("  {0:N1} 秒" -f $seconds)
    } else {
        Write-Host "TSV が無い（pack だけの）インデックスなので、pack の作成は測りません。"
    }

    # gram 索引の作成
    $gramWork = $null
    if (Get-Command updateGramIndexes -CommandType Function -ErrorAction SilentlyContinue) { $gramWork = getPerfGramWorkDir $Index }
    if (!(Get-Command updateGramIndexes -CommandType Function -ErrorAction SilentlyContinue)) {
        $result.Gram = [ordered]@{ Skipped = "測る版に gram 索引が無い" }
    } elseif (!$gramWork) {
        $result.Gram = [ordered]@{ Skipped = "置き場の形" }
    } else {
        $target = [Workspace]::new($gramWork)
        $gramRoot = $target.GramDir
        setMonitorPhase $monitor "gram 索引の作成"
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $gram = updateGramIndexes -indexRoot $target.IndexDir -gramRoot $gramRoot
        $gramSeconds = $watch.Elapsed.TotalSeconds
        $gramBytes = 0L
        if ([System.IO.Directory]::Exists($gramRoot)) {
            foreach ($f in [System.IO.Directory]::GetFiles($gramRoot, "*", "AllDirectories")) { $gramBytes += ([System.IO.FileInfo]$f).Length }
        }
        $result.Gram = [ordered]@{ Seconds = [Math]::Round($gramSeconds, 2); Files = [int]$gram.Files; MB = [Math]::Round($gramBytes / 1MB, 1); Failed = @($gram.Failed).Count; Skipped = $null }
    }
} finally {
    $samples = stopResourceMonitor $monitor
}
$result.Samples = $samples
$result.PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
[System.IO.File]::WriteAllText($Out, ($result | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
