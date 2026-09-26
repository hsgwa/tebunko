# Office からの取り込みを測る。tools\measure_perf.ps1 が別のプロセスで起動する（リソースが混ざらないようにする）。
# 1 回の実行で 1 回だけ取り込む（新しいプロセス・空のワークスペース。Office の起動を含む、初めての取り込みの時間）。回数は呼び出し側の -Repeat。
#
# 測る tebunko の scripts\ を <Work>\ingest\tool\scripts に写し、写したツールのフォルダに setting.config を書いて取り込む。
# 写したフォルダは書き込めるので設定はそこを使い、利用者の設定・既定のワークスペース・リポジトリの setting.config には触らない。
#
# 計測の口（この計測が頼るもの。refactor で変えるときは、先にこの計測を直し、main で測り直す。docs/design/testing/ci.md に同じ表がある）:
#   起動口       … scripts\tebunko\indexer.ps1 -Channel <口>。終了コード 0 = 完了
#   読み込み口   … scripts\tebunko\lib.ps1 の newIndexerChannel（引数 retryFailed・confirmTargets・workers）
#   受け渡しの口 … Progress.Phase と、その値 クロール・確認・取り込み・仕上げ
#   設定         … setting.config がツールのフォルダにあること。キー targetFolders（@{ name; path; enabled } の配列）・workspaceFolder・ingestThreads
#   取り込み一覧 … ワークスペース直下の 取り込み一覧.tsv。見出しの 相対パス・状態。状態の値 済・失敗
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Data,
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Threads = 2,
    [int]$SampleMs = 200
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\perf_common.ps1"
. "$PSScriptRoot\ingest_common.ps1"

$Tool = (Resolve-Path -LiteralPath $Tool).ProviderPath.TrimEnd("\")
$Data = (Resolve-Path -LiteralPath $Data).ProviderPath.TrimEnd("\")
if (!(Test-Path -LiteralPath (Join-Path $Tool "scripts\tebunko\indexer.ps1"))) {
    throw "測る tebunko に scripts\tebunko\indexer.ps1 がありません。取り込みを測れるのは、この起動口がある版です。"
}
$counts = getIngestKindCounts $Data
if ($counts.Total -eq 0) { throw "データのフォルダに取り込むファイルがありません。" }

# 写したツールと空のワークスペースを用意する
[void][System.IO.Directory]::CreateDirectory($Work)
$ingestRoot = Join-Path (Resolve-Path -LiteralPath $Work).ProviderPath "ingest"
if ([System.IO.Directory]::Exists($ingestRoot)) { [System.IO.Directory]::Delete($ingestRoot, $true) }
$toolCopy = Join-Path $ingestRoot "tool"
$workspaceDir = Join-Path $ingestRoot "ws"
[void][System.IO.Directory]::CreateDirectory($toolCopy)
[void][System.IO.Directory]::CreateDirectory($workspaceDir)
Copy-Item -LiteralPath (Join-Path $Tool "scripts") -Destination (Join-Path $toolCopy "scripts") -Recurse
$settings = [ordered]@{
    targetFolders   = @([ordered]@{ name = "計測"; path = $Data; enabled = $true })
    workspaceFolder = $workspaceDir
    ingestThreads   = $Threads
}
[System.IO.File]::WriteAllText((Join-Path $toolCopy "setting.config"), ($settings | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($false)))

. (Join-Path $toolCopy "scripts\tebunko\lib.ps1")
assertTebunkoFunctions @("newIndexerChannel")
$channel = newIndexerChannel -retryFailed $false -confirmTargets $false -workers -1

Write-Host ("取り込みを測ります（{0} ファイル・読み取りのスレッド {1}・Office: {2}）…" -f $counts.Total, $Threads, (getIngestLanes $counts))
$monitor = startResourceMonitor $SampleMs $channel
$exitCode = -1
$watch = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $null = & (Join-Path $toolCopy "scripts\tebunko\indexer.ps1") -Channel $channel
    $exitCode = $LASTEXITCODE
} finally {
    $seconds = $watch.Elapsed.TotalSeconds
    $samples = stopResourceMonitor $monitor
}
if ($exitCode -ne 0) { throw "取り込みが終了コード $exitCode で終わりました。$($channel.Error)" }

$phaseSeconds = getIngestPhaseSeconds $monitor.Phases.ToArray()
$status = readIngestStatusCounts (Join-Path $workspaceDir $ingestStatusFileName)
assertIngestCounts $status $counts.Total
$ingestSeconds = $phaseSeconds[$ingestPhaseIngest]
$perFileMs = if ($ingestSeconds -gt 0) { $ingestSeconds * 1000 / $counts.Total } else { $null }

$result = [ordered]@{
    Files = $counts; Total = $counts.Total; Done = $status.Done; Failed = $status.Failed
    Threads = $Threads; Lanes = (getIngestLanes $counts)
    Seconds = [Math]::Round($seconds, 2)
    Phases = @($phaseSeconds.Keys | ForEach-Object { [ordered]@{ Phase = $_; Seconds = [Math]::Round($phaseSeconds[$_], 2) } })
    IngestSeconds = [Math]::Round($ingestSeconds, 2)
    PerFileMs = $(if ($null -ne $perFileMs) { [Math]::Round($perFileMs, 1) } else { $null })
    Resources = (getPhaseResources $samples)
    PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
    Samples = $samples
}
[System.IO.File]::WriteAllText($Out, ($result | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("  {0:N1} 秒（取り込み {1:N1} 秒・成功 {2}・失敗 {3}）" -f $seconds, $ingestSeconds, $status.Done, $status.Failed)
