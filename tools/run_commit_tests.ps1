# コミット前のフック（tools/hooks/pre-commit）で流すテストを、ステージした変更から選んで流す。
#
#   .\tools\run_commit_tests.ps1                      ステージした変更に応じたテストを流す
#   .\tools\run_commit_tests.ps1 -List -Files a,b     変更したファイルを指定し、選んだテストを出すだけ（テスト用）
#
# 選び方（パスはリポジトリからの相対パス）:
#   - scripts/<パス>.ps1 は tests/<パス>.Tests.ps1 と、構成を守るテスト（structure・layers）
#     scripts/tebunko/indexer.ps1 は tests/tebunko/indexer/indexer.Tests.ps1
#   - scripts/ の .xaml は structure（XML として読めるか・画面の部品の名前）
#   - tests/ の *.Tests.ps1 はそのテスト自身
#   - tools/<名前>.ps1 は tests/tools/<名前>.Tests.ps1（無ければ流さない）。check_markdown_links.ps1 は links も
#   - tests/testdata/scrub_personal.ps1 は tests/testdata/scrub_personal.Tests.ps1
#   - .md と docs/ の中は links（リンク先のファイル・見出し）
#   - 次のものは、どのテストに効くか分からないため速いテスト（Unit・Meta）を全部流す
#     対応するテストが無い・消した scripts/ の .ps1、tests/ の上のどれにも当たらないもの（run.ps1・helpers・testdata など）
#   - それ以外（.github/・画像・設定など）はテストを流さない
# 選んだテストファイルは既定のタグ（Unit・Io・Meta）で流す。文字コードは tools/check_commit.ps1 -Staged が、
# 全テスト・静的解析は CI（test）が確かめる。
param (
    [string[]]$Files,  # 変更したファイル（既定はステージした変更）
    [switch]$List      # テストを流さず、選んだテストファイル（全部なら "all"）を出す
)

$ErrorActionPreference = "Stop"

$rootDir = (Resolve-Path "$PSScriptRoot\..").Path

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
        if ($file -match '(?i)^tools/([^/]+)\.ps1$') {
            $test = "tests/tools/$($Matches[1]).Tests.ps1"
            if (Test-Path -LiteralPath (Join-Path $root $test)) { & $add $test }
            if ($Matches[1] -ieq "check_markdown_links") { & $add "tests/meta/links.Tests.ps1" }
            continue
        }
    }
    return @{ All = $false; Paths = $paths.ToArray() }
}

if ($PSBoundParameters.ContainsKey("Files")) {
    # powershell.exe -File で呼ぶと、-Files a,b は配列にならず 1 つの文字列で渡るため、カンマで分ける
    $Files = @($Files | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() })
} else {
    # 名前を変えたファイルは、消したファイルと足したファイルに分けて受け取る（消した側も選ぶ材料にするため）
    $Files = @(git -C $rootDir -c core.quotepath=off diff --cached --name-only --no-renames)
    if ($LASTEXITCODE -ne 0) { throw "git diff --cached が失敗しました" }
}

$selected = selectCommitTests $Files $rootDir

if ($List) {
    if ($selected.All) { "all" } else { $selected.Paths }
    exit 0
}

$run = Join-Path $rootDir "tests\run.ps1"
if ($selected.All) {
    Write-Host "速いテスト（Unit・Meta）を全部流します。"
    & $run -Tag Unit, Meta -Quiet
    exit $LASTEXITCODE
}
if ($selected.Paths.Count -eq 0) {
    Write-Host "ステージした変更に対応するテストはありません（テストは CI で流します）。"
    exit 0
}
Write-Host "変更に対応するテストを流します: $(($selected.Paths | ForEach-Object { $_ -replace '^tests/', '' }) -join ', ')"
& $run -Path @($selected.Paths | ForEach-Object { Join-Path $rootDir $_ }) -Quiet
exit $LASTEXITCODE
