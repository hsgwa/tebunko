# コミット前のフック（tools/hooks/pre-commit）で流すテストを、ステージした変更から選んで流す。
#
#   .\tools\run_commit_tests.ps1                      ステージした変更に応じたテストを流す
#   .\tools\run_commit_tests.ps1 -List -Files a,b     変更したファイルを指定し、選んだテストを出すだけ（テスト用）
#
# 選び方（パスはリポジトリからの相対パス）:
#   - scripts/<パス>.ps1 は tests/<パス>.Tests.ps1 と、構成を守るテスト（structure・layers）
#     scripts/tebunko/indexer.ps1 は tests/tebunko/indexer/indexer.Tests.ps1
#   - scripts/ の .xaml は structure（XML として読めるか・画面の部品の名前）と table_border（表に外枠を付けない）
#   - tests/gui/（本物の画面を開くテスト。タグ Gui）は選ばない。CI（gui.yml）と手元の -Tag Gui で流す
#   - tests/ の *.Tests.ps1 はそのテスト自身
#   - tools/<名前>.ps1 は tests/tools/<名前>.Tests.ps1（無ければ流さない）。check_markdown_links.ps1 は links も
#   - tools/isolation/*.ps1（既定のワークスペースを守る共通の関数）は tests/tools/run_isolated.Tests.ps1
#   - tests/testdata/scrub_personal.ps1 は tests/testdata/scrub_personal.Tests.ps1
#   - .md と docs/ の中は links（リンク先のファイル・見出し）
#   - 次のものは、どのテストに効くか分からないため速いテスト（Unit・Meta）を全部流す
#     対応するテストが無い・消した scripts/ の .ps1、tests/ の上のどれにも当たらないもの（run.ps1・helpers・testdata など）
#   - それ以外（.github/・画像・設定など）はテストを流さない
# 選んだテストファイルは既定のタグ（Unit・Io・Meta）で流す。文字コードは tools/check_commit.ps1 -Staged が確かめる。
# あわせて、ステージした scripts・tools・tests の .ps1 だけに PSScriptAnalyzer をかけ（tools/check_analyzer.ps1 -Path）、
# ベースラインより指摘が増えたら失敗にする。PSScriptAnalyzer が無い手元・Windows PowerShell 以外で動いたときは、注意を出して飛ばす。
# 全テストと全体の静的解析は CI（test）が確かめる。
param (
    [string[]]$Files,  # 変更したファイル（既定はステージした変更）
    [switch]$List,     # テストを流さず、選んだテストファイル（全部なら "all"）を出す
    [switch]$ListAnalyzer,  # 静的解析をかけるファイルを出すだけ（テスト用）
    [switch]$ListAnalyzerVersion,  # 静的解析に使う PSScriptAnalyzer の版（CI が固定した版）を出すだけ（テスト用）
    [int]$TestExit = -1,   # 0 以上なら、テストを流さず、そのテストの終了コードとして扱う（テスト用）
    [string]$AnalyzerVersion,  # PSScriptAnalyzer の版（既定は .github/workflows/test.yml の PSSA_VERSION。テストで差し替える）
    [string]$AnalyzerScript = "$PSScriptRoot\check_analyzer.ps1",  # 静的解析のスクリプト（テストで差し替える）
    [string]$AnalyzerModule = "PSScriptAnalyzer",  # 静的解析のモジュール名（テストで差し替える）
    [string]$Edition = $PSVersionTable.PSEdition   # PowerShell の版（テストで差し替える）
)

$ErrorActionPreference = "Stop"

$rootDir = [System.IO.Path]::GetFullPath("$PSScriptRoot\..")

# 変更したファイルから流すテストを選ぶ。All が $true なら速いテストを全部流す。Paths は tests/ からの相対パス
function selectCommitTests([string[]]$files, [string]$root) {
    $paths = New-Object System.Collections.Generic.List[string]
    $add = {
        param([string]$path)
        if (!$paths.Contains($path)) { $paths.Add($path) }
    }
    foreach ($file in @($files | Where-Object { $_ })) {
        $file = $file.Replace('\', '/')
        $exists = Test-Path -LiteralPath (Join-Path $root $file)
        if ($file -match '(?i)\.md$' -or $file -match '^docs/') {
            & $add "tests/meta/links.Tests.ps1"
            continue
        }
        if ($file -match '(?i)^scripts/.+\.xaml$') {
            & $add "tests/meta/structure.Tests.ps1"
            & $add "tests/meta/table_border.Tests.ps1"
            & $add "tests/meta/settings_file_text.Tests.ps1"
            continue
        }
        if ($file -match '(?i)^scripts/(.+)\.ps1$') {
            $test = if ($file -ieq "scripts/tebunko/indexer.ps1") { "tests/tebunko/indexer/indexer.Tests.ps1" } else { "tests/$($Matches[1]).Tests.ps1" }
            if (!$exists -or !(Test-Path -LiteralPath (Join-Path $root $test))) {
                return @{ All = $true; Paths = @() }
            }
            & $add $test
            & $add "tests/meta/structure.Tests.ps1"
            & $add "tests/meta/layers.Tests.ps1"
            continue
        }
        if ($file -match '(?i)^tests/gui/') {
            continue
        }
        if ($file -match '(?i)^tests/.+\.Tests\.ps1$') {
            if ($exists) { & $add $file }
            continue
        }
        if ($file -ieq "tests/testdata/scrub_personal.ps1") {
            & $add "tests/testdata/scrub_personal.Tests.ps1"
            continue
        }
        if ($file -match '^tests/') {
            return @{ All = $true; Paths = @() }
        }
        if ($file -match '(?i)^tools/isolation/[^/]+\.ps1$') {
            & $add "tests/tools/run_isolated.Tests.ps1"
            continue
        }
        if ($file -match '(?i)^tools/([^/]+)\.ps1$') {
            $test = "tests/tools/$($Matches[1]).Tests.ps1"
            if (Test-Path -LiteralPath (Join-Path $root $test)) { & $add $test }
            if ($Matches[1] -ieq "check_markdown_links") { & $add "tests/meta/links.Tests.ps1" }
            continue
        }
    }
    return @{ All = $false; Paths = $paths.ToArray() }
}

# CI が固定した PSScriptAnalyzer の版（.github/workflows/test.yml の PSSA_VERSION）。無ければ空
function readPinnedAnalyzerVersion([string]$root) {
    $yml = Join-Path $root ".github\workflows\test.yml"
    if (!(Test-Path -LiteralPath $yml)) { return "" }
    foreach ($line in [System.IO.File]::ReadAllLines($yml)) {
        if ($line -match '^\s*PSSA_VERSION:\s*([0-9][0-9.]*)\s*$') { return $Matches[1] }
    }
    return ""
}

# テストと静的解析の終了コードから、フックの終了コードを決める（テストの失敗を先に返す）
function decideCommitExit([int]$testExit, [int]$analyzerExit) {
    if ($testExit -ne 0) { return $testExit }
    return $analyzerExit
}

# 選んだファイルに check_analyzer.ps1 -Path をかける。戻り値は終了コード（飛ばしたときは 0）。
# ベースラインは CI が固定した版の件数なので、その版が手元に無いときも飛ばす（版が違うと件数がずれ、コードを変えなくても止まるため）
function runCommitAnalyzer([string[]]$targets, [string]$script, [string]$module, [string]$version, [string]$edition) {
    if ($targets.Count -eq 0) { return 0 }
    if ($edition -ne "Desktop") {
        Write-Warning "Windows PowerShell 以外で動いているため、静的解析を飛ばします（CI が検査します）。powershell.exe で実行してください。"
        return 0
    }
    $installed = @(Get-Module -ListAvailable -Name $module)
    if ($installed.Count -eq 0) {
        Write-Warning "$module が入っていないため、静的解析を飛ばします（CI が検査します）。"
        return 0
    }
    if ($version -and !(@($installed | Where-Object { $_.Version.ToString() -eq $version }).Count)) {
        Write-Warning "$module $version が入っていないため、静的解析を飛ばします（ベースラインはこの版の件数。CI が検査します）。"
        return 0
    }
    "静的解析をかけます: $($targets -join ', ')" | Out-Host
    $was = $env:PSSA_VERSION
    try {
        if ($version) { $env:PSSA_VERSION = $version }
        & powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $script -Path ($targets -join ",") | Out-Host
        return $LASTEXITCODE
    } finally {
        $env:PSSA_VERSION = $was
    }
}

if ($PSBoundParameters.ContainsKey("Files")) {
    # powershell.exe -File で呼ぶと、-Files a,b は配列にならず 1 つの文字列で渡るため、カンマで分ける
    $Files = @($Files | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() })
} else {
    # 名前を変えたファイルは、消したファイルと足したファイルに分けて受け取る（消した側も選ぶ材料にするため）
    $Files = @(git -C $rootDir -c core.quotepath=off diff --cached --name-only --no-renames)
    if ($LASTEXITCODE -ne 0) { throw "git diff --cached が失敗しました" }
}

. (Join-Path $PSScriptRoot "analyzer_baseline.ps1")
if (!$PSBoundParameters.ContainsKey("AnalyzerVersion")) { $AnalyzerVersion = readPinnedAnalyzerVersion $rootDir }
$selected = selectCommitTests $Files $rootDir
$analyzerTargets = @(selectAnalyzerFiles $Files $rootDir)

if ($List) {
    if ($selected.All) { "all" } else { $selected.Paths }
    exit 0
}
if ($ListAnalyzerVersion) {
    $AnalyzerVersion
    exit 0
}
if ($ListAnalyzer) {
    $analyzerTargets
    exit 0
}

$run = Join-Path $rootDir "tests\run.ps1"
$testCode = 0
if ($TestExit -ge 0) {
    $testCode = $TestExit
} elseif ($selected.All) {
    Write-Host "速いテスト（Unit・Meta）を全部流します。"
    & $run -Tag Unit, Meta -Quiet
    $testCode = $LASTEXITCODE
} elseif ($selected.Paths.Count -eq 0) {
    Write-Host "ステージした変更に対応するテストはありません（テストは CI で流します）。"
} else {
    Write-Host "変更に対応するテストを流します: $(($selected.Paths | ForEach-Object { $_ -replace '^tests/', '' }) -join ', ')"
    & $run -Path @($selected.Paths | ForEach-Object { Join-Path $rootDir $_ }) -Quiet
    $testCode = $LASTEXITCODE
}

# テストが落ちていても静的解析は流し、両方の結果を見せる
$analyzerExit = runCommitAnalyzer $analyzerTargets $AnalyzerScript $AnalyzerModule $AnalyzerVersion $Edition
exit (decideCommitExit $testCode $analyzerExit)
