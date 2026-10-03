# 前の版のファイル（見本・golden。tests/testdata/compat/settings/<見本の名前>/setting.config）を作る手伝い。
#
#   .\tools\make_settings_golden.ps1 -Tree C:\tebunko_golden\v0.3.1 -Name v0.3.1
#
# 手順（詳しくは tests/testdata/README.md の「前の版のファイル（compat\）」）:
#   1. -Tree に、見本にしたい版をチェックアウトした別のツリー（例: git worktree add --detach）を指す
#   2. 別の Windows PowerShell 5.1 プロセスで、そのツリーの shared.ps1・settings.ps1 と、
#      今のツリーの tests/helpers/settings_golden.ps1 を読み込み、writeGoldenSettings を呼んで
#      tests/testdata/compat/settings/<Name>/setting.config に書く
#   3. expected.json は手で書く（このスクリプトは作らない）
#
# tests/helpers/settings_golden.ps1 は load.ps1 に頼らず、settings.ps1 が公開する書く関数だけを呼ぶ。
# そのため、見本にしたい版の settings.ps1 がそれらの関数を持ってさえいれば、そのまま上で動く。
param (
    [Parameter(Mandatory = $true)]
    [string]$Tree,
    [Parameter(Mandatory = $true)]
    [string]$Name
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$sharedPath = Join-Path $Tree "scripts\shared\shared.ps1"
$settingsScriptPath = Join-Path $Tree "scripts\tebunko\core\settings.ps1"
$goldenHelperPath = Join-Path $repoRoot "tests\helpers\settings_golden.ps1"

foreach ($p in @($sharedPath, $settingsScriptPath, $goldenHelperPath)) {
    if (!(Test-Path -LiteralPath $p)) {
        throw "見つかりません: ${p}"
    }
}

$sampleDir = Join-Path $repoRoot "tests\testdata\compat\settings\${Name}"
New-Item -ItemType Directory -Force -Path $sampleDir | Out-Null
$settingsPath = Join-Path $sampleDir "setting.config"
if (Test-Path -LiteralPath $settingsPath) {
    Remove-Item -LiteralPath $settingsPath -Force
}

$runnerPath = Join-Path ([System.IO.Path]::GetTempPath()) "make_settings_golden_$([guid]::NewGuid().ToString("N")).ps1"
$runner = @"
`$ErrorActionPreference = "Stop"
. '$sharedPath'
. '$settingsScriptPath'
. '$goldenHelperPath'
writeGoldenSettings -path '$settingsPath'
"@
[System.IO.File]::WriteAllText($runnerPath, $runner, (New-Object System.Text.UTF8Encoding($true)))

try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $runnerPath
    if ($LASTEXITCODE -ne 0) {
        throw "見本を書くプロセスが失敗しました（終了コード ${LASTEXITCODE}）"
    }
} finally {
    Remove-Item -LiteralPath $runnerPath -Force -ErrorAction SilentlyContinue
}

if (!(Test-Path -LiteralPath $settingsPath)) {
    throw "${settingsPath} が書かれませんでした"
}

Write-Host "見本を作りました: ${sampleDir}"
Write-Host "  setting.config"
Write-Host "  expected.json は手で書いてください"
