# 開発の準備を 1 回で済ませる。clone したら実行する（何度流しても壊れない）。
#
#   .\tools\setup_dev.ps1              準備する
#   .\tools\setup_dev.ps1 -SkipDocs    MkDocs の準備を飛ばす（Python が無いとき）
#   .\tools\setup_dev.ps1 -DryRun      何も入れず、やることを出すだけ
#
# やること（版は .github\workflows\test.yml の PESTER_VERSION・PSSA_VERSION を読んで固定する。ここには書かない）:
#   1. Pester（テストの実行）を入れる
#   2. PSScriptAnalyzer（コミット前の静的解析）を入れる
#   3. tools\install_hooks.ps1 でコミット前の検査を有効にする
#   4. MkDocs 用の venv（work\mkdocs-venv）を作り、tools\mkdocs\requirements.txt を入れる
param (
    [switch]$DryRun,
    [switch]$SkipDocs,
    [string]$WorkflowPath = ""
)

$ErrorActionPreference = "Stop"

# test.yml のテキストから、<名前>: <版> の行の版を取り出す。無ければ空
function readPinnedVersion([string]$text, [string]$name) {
    foreach ($line in ($text -split "\r?\n")) {
        if ($line -match ('^\s*' + [regex]::Escape($name) + ':\s*([0-9][0-9.]*)\s*$')) { return $Matches[1] }
    }
    return ""
}

# 進み具合を出す（Write-Host を使わず、取り込めるよう情報ストリームに出す）
function say([string]$message) {
    Write-Information $message -InformationAction Continue
}

# モジュールを、その版が入っていなければ入れる。DryRun のときは入れない
function ensureModule([string]$name, [string]$version, [bool]$dryRun) {
    $installed = @(Get-Module -ListAvailable -Name $name | Where-Object { $_.Version -eq [version]$version })
    if ($installed.Count -gt 0) {
        say "${name} ${version}: 入っています（何もしません）"
        return
    }
    if ($dryRun) {
        say "${name} ${version}: 入れます（Install-Module -Scope CurrentUser）"
        return
    }
    Install-Module -Name $name -RequiredVersion $version -Scope CurrentUser -Force -SkipPublisherCheck
    say "${name} ${version}: 入れました"
}

function runHooksInstall([string]$root) {
    & (Join-Path $PSScriptRoot "install_hooks.ps1")
}

# Python の実行ファイルのパス。無ければ空
function findPython() {
    foreach ($name in @("python", "py")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $cmd) { return $cmd.Source }
    }
    return ""
}

function setupDocsVenv([string]$root) {
    $venv = Join-Path (Join-Path $root "work") "mkdocs-venv"
    $python = findPython
    if ($python -eq "") {
        Write-Warning "MkDocs: Python が見つからないため飛ばします（Python 3.12 を入れて、もう一度実行してください）"
        return
    }
    if (!(Test-Path -LiteralPath $venv)) { & $python -m venv $venv }
    $venvPython = Join-Path (Join-Path $venv "Scripts") "python.exe"
    & $venvPython -m pip install --disable-pip-version-check --require-hashes -r (Join-Path (Join-Path (Join-Path $root "tools") "mkdocs") "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "MkDocs のパッケージを入れられませんでした。" }
    say "MkDocs: work\mkdocs-venv に入れました（work\mkdocs-venv\Scripts\mkdocs で使います）"
}

# 準備の全体。戻り値は終了コード
function runSetupDev([string]$root, [string]$workflowPath, [bool]$dryRun, [bool]$skipDocs) {
    if (!(Test-Path -LiteralPath $workflowPath)) {
        Write-Warning "版の読み元がありません: $workflowPath"
        return 1
    }
    $text = [System.IO.File]::ReadAllText($workflowPath)
    $pester = readPinnedVersion $text "PESTER_VERSION"
    $pssa = readPinnedVersion $text "PSSA_VERSION"
    if ($pester -eq "" -or $pssa -eq "") {
        Write-Warning "test.yml から PESTER_VERSION・PSSA_VERSION を読めません。版を決められないため止めます。"
        return 1
    }
    say "[1/4] Pester $pester"
    ensureModule "Pester" $pester $dryRun
    say "[2/4] PSScriptAnalyzer $pssa"
    ensureModule "PSScriptAnalyzer" $pssa $dryRun
    say "[3/4] install_hooks（コミット前の検査）"
    if ($dryRun) { say "install_hooks.ps1 を実行します" } else { runHooksInstall $root }
    say "[4/4] mkdocs（MkDocs 用の venv）"
    if ($skipDocs) { say "-SkipDocs のため飛ばします" }
    elseif ($dryRun) { say "work\mkdocs-venv を作り、tools\mkdocs\requirements.txt を入れます" }
    else { setupDocsVenv $root }
    return 0
}

# 読み込んだだけ（. .\setup_dev.ps1）のときは関数だけ使えるようにして、何も実行しない
if ($MyInvocation.InvocationName -eq ".") { return }

$rootDir = Split-Path $PSScriptRoot -Parent
if ($WorkflowPath -eq "") { $WorkflowPath = Join-Path (Join-Path (Join-Path $rootDir ".github") "workflows") "test.yml" }
exit (runSetupDev $rootDir $WorkflowPath $DryRun.IsPresent $SkipDocs.IsPresent)
