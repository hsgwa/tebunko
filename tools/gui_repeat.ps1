# 画面のスモークテストを繰り返し流し、回ごとの結果（成功・終了コード 5・閉じない・記録の待ち切れ・そのほか）と失敗の材料を残す。
# gui.yml の gui-repeat が、shard ごとに呼ぶ（docs/design/testing/ci.md「gui.yml」）。途中で落ちても続け、終了コードは常に 0。
param (
    [Parameter(Mandatory)][int]$Shard,
    [Parameter(Mandatory)][int]$Total,   # 全体の回数。5 本の shard で分ける
    [Parameter(Mandatory)][string]$Paths, # 流すテストのファイル（カンマ区切り）
    [string]$OutDir = "repeat-results"
)

$repoRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $repoRoot

# 1 回分の出力から、落ち方を分ける
function getFailureKind {
    param ([int]$Code, [string]$Text)
    if ($Code -eq 0) { return "成功" }
    if ($Text -match '画面の終了コードが 0 ではない（5）') { return "終了コード 5" }
    if ($Text -match '画面が終了する が 30 秒以内に見つからなかった') { return "閉じない" }
    if ($Text -match '記録のファイルができなかった') { return "記録の待ち切れ" }
    return "そのほか"
}

$times = [int][Math]::Ceiling($Total / 5)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$rows = @()
for ($i = 1; $i -le $times; $i++) {
    $log = Join-Path $OutDir "run-$Shard-$i.log"
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    & powershell -NoProfile -ExecutionPolicy Bypass -File "$repoRoot\tests\run.ps1" -Tag Gui -Path $Paths *> $log
    $code = $LASTEXITCODE
    $watch.Stop()
    $seconds = [int]$watch.Elapsed.TotalSeconds
    $kind = getFailureKind $code (Get-Content -LiteralPath $log -Raw)
    $rows += "$Shard,$i,$code,$kind,$seconds"
    Write-Host "[$Shard-$i] $kind（終了コード $code・$seconds 秒）"
    if ($code -ne 0 -and (Test-Path "work\test\gui")) {
        Copy-Item -Recurse -Force "work\test\gui" (Join-Path $OutDir "evidence-$Shard-$i")
    }
}
$rows | Set-Content -LiteralPath (Join-Path $OutDir "result-$Shard.csv") -Encoding UTF8
exit 0
