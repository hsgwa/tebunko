# テストの実行（Pester 3.4）
#
#   .\tests\run.ps1              既定（Unit・Io・Meta。Office と Slow は除く）
#   .\tests\run.ps1 -Tag Unit    速い確認だけ
#   .\tests\run.ps1 -All         Office・Slow も含めて全部（Office が必要）
#   .\tests\run.ps1 -Ci          結果の XML とカバレッジを出し、失敗数を終了コードにする
param (
    [string[]]$Tag,
    [string[]]$ExcludeTag,
    [switch]$All,    # Office・Slow も実行する（Excel・Word・PowerPoint が必要）
    [switch]$Ci,     # 結果の XML とカバレッジを出し、失敗数で終了する
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"

$testsDir = $PSScriptRoot
$rootDir  = Split-Path $testsDir -Parent
$outDir   = "$rootDir\work\test"

# 既定で外すタグ。Office は COM が要るもの、Slow は時間がかかるもの、Manual は手で確かめるもの
$defaultExclude = @("Office", "Slow", "Manual")
if (!$PSBoundParameters.ContainsKey("ExcludeTag")) {
    $ExcludeTag = if ($All) { @("Manual") } else { $defaultExclude }
}

$arguments = @{
    Script     = $testsDir
    PassThru   = $true
    ExcludeTag = $ExcludeTag
}
if ($Tag)   { $arguments.Tag = $Tag }
if ($Quiet) { $arguments.Quiet = $true }

if ($Ci) {
    [System.IO.Directory]::CreateDirectory($outDir) | Out-Null
    $arguments.OutputFile   = "$outDir\results.xml"
    $arguments.OutputFormat = "NUnitXml"
    # カバレッジの対象は判断層・状態層だけにする（画面層は自動テストの対象外）
    $arguments.CodeCoverage = @(Get-ChildItem "$rootDir\scripts" -Recurse -Filter "*.ps1" |
        Where-Object { $_.Name -notmatch "^(gui|shell|app_host)\.ps1$" -and $_.Name -notmatch "_tab\.ps1$" -and $_.Name -notmatch "_dialog\.ps1$" } |
        ForEach-Object { $_.FullName })
}

$result = Invoke-Pester @arguments

if ($Ci) {
    $covered = 0
    $total = 0
    if ($result.CodeCoverage) {
        $covered = @($result.CodeCoverage.HitCommands).Count
        $total   = $covered + @($result.CodeCoverage.MissedCommands).Count
    }
    $percent = if ($total -gt 0) { [Math]::Round(100.0 * $covered / $total, 1) } else { 0 }
    "テスト {0} 件中 {1} 件成功 / {2} 件失敗　カバレッジ {3}%（{4}/{5}）" -f `
        $result.TotalCount, $result.PassedCount, $result.FailedCount, $percent, $covered, $total | Write-Host

    # カバレッジの下限（tests\coverage.baseline。下回ったら失敗にする）
    $baselineFile = "$testsDir\coverage.baseline"
    if (Test-Path -LiteralPath $baselineFile) {
        $baseline = [double]((Get-Content -LiteralPath $baselineFile -Raw).Trim())
        if ($percent -lt $baseline) {
            Write-Host "カバレッジが基準 $baseline% を下回りました（$percent%）。" -ForegroundColor Red
            exit 1
        }
    }
    exit $result.FailedCount
}