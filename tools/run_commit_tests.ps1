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
    [switch]$AnalyzerOnly,  # テストを流さず、静的解析だけを流す（テスト用）
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

# 選んだファイルに check_analyzer.ps1 -Path をかける。戻り値は終了コード（飛ばしたときは 0）
function runCommitAnalyzer([string[]]$targets, [string]$script, [string]$module, [string]$edition) {
    if ($targets.Count -eq 0) { return 0 }
    if ($edition -ne "Desktop") {
        Write-Warning "Windows PowerShell 以外で動いているため、静的解析を飛ばします（CI が検査します）。powershell.exe で実行してください。"
        return 0
    }
    if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Warning "$module が入っていないため、静的解析を飛ばします（CI が検査します）。"
        return 0
    }
    "静的解析をかけます: $($targets -join ', ')" | Out-Host
    & powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $script -Path ($targets -join ",") | Out-Host
    return $LASTEXITCODE
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
$selected = selectCommitTests $Files $rootDir
$analyzerTargets = @(selectAnalyzerFiles $Files $rootDir)

if ($List) {
    if ($selected.All) { "all" } else { $selected.Paths }
    exit 0
}
if ($ListAnalyzer) {
    $analyzerTargets
    exit 0
}

$run = Join-Path $rootDir "tests\run.ps1"
$testExit = 0
if ($AnalyzerOnly) {
    # 静的解析だけ（テスト用）
} elseif ($selected.All) {
    Write-Host "速いテスト（Unit・Meta）を全部流します。"
    & $run -Tag Unit, Meta -Quiet
    $testExit = $LASTEXITCODE
} elseif ($selected.Paths.Count -eq 0) {
    Write-Host "ステージした変更に対応するテストはありません（テストは CI で流します）。"
} else {
    Write-Host "変更に対応するテストを流します: $(($selected.Paths | ForEach-Object { $_ -replace '^tests/', '' }) -join ', ')"
    & $run -Path @($selected.Paths | ForEach-Object { Join-Path $rootDir $_ }) -Quiet
    $testExit = $LASTEXITCODE
}

# テストが落ちていても静的解析は流し、両方の結果を見せる
$analyzerExit = runCommitAnalyzer $analyzerTargets $AnalyzerScript $AnalyzerModule $Edition
if ($testExit -ne 0) { exit $testExit }
exit $analyzerExit
