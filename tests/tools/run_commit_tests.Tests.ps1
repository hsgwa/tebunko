# コミット前に流すテストを選ぶスクリプト（tools\run_commit_tests.ps1）のテスト。
# 対応するテストがあるかは、リポジトリにある本物のファイルで調べる
BeforeAll {
    $select = "$PSScriptRoot\..\..\tools\run_commit_tests.ps1"

    function selectTests([string[]]$files) {
        return @(& $select -List -Files $files)
    }
}

Describe "run_commit_tests.ps1 の選び方" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "scripts の .ps1 は対応するテストと構成のテスト"; files = @("scripts/tebunko/search/search_query.ps1"); expected = @("tests/tebunko/search/search_query.Tests.ps1", "tests/meta/structure.Tests.ps1", "tests/meta/layers.Tests.ps1") }
        @{ name = "インデクサの起動口は indexer フォルダのテスト"; files = @("scripts/tebunko/indexer.ps1"); expected = @("tests/tebunko/indexer/indexer.Tests.ps1", "tests/meta/structure.Tests.ps1", "tests/meta/layers.Tests.ps1") }
        @{ name = "\ 区切りのパスも同じに扱う"; files = @("scripts\shared\core\fs.ps1"); expected = @("tests/shared/core/fs.Tests.ps1", "tests/meta/structure.Tests.ps1", "tests/meta/layers.Tests.ps1") }
        @{ name = "カンマ区切りの 1 つの文字列も分ける（powershell.exe -File の渡し方）"; files = @("tests/shared/core/text.Tests.ps1,README.md"); expected = @("tests/shared/core/text.Tests.ps1", "tests/meta/links.Tests.ps1") }
        @{ name = "XAML は structure"; files = @("scripts/tebunko/xaml/tab_search.xaml"); expected = @("tests/meta/structure.Tests.ps1") }
        @{ name = "テストを変えたらそのテスト"; files = @("tests/shared/core/text.Tests.ps1"); expected = @("tests/shared/core/text.Tests.ps1") }
        @{ name = "Markdown と docs の中は links"; files = @("README.md", "docs/images/none.png"); expected = @("tests/meta/links.Tests.ps1") }
        @{ name = "道具はそのテスト。リンクの検査は links も"; files = @("tools/check_signoff.ps1", "tools/check_markdown_links.ps1"); expected = @("tests/tools/check_signoff.Tests.ps1", "tests/tools/check_markdown_links.Tests.ps1", "tests/meta/links.Tests.ps1") }
        @{ name = "個人情報の除去はそのテスト"; files = @("tests/testdata/scrub_personal.ps1"); expected = @("tests/testdata/scrub_personal.Tests.ps1") }
        @{ name = "同じテストは 1 回だけ"; files = @("scripts/shared/core/fs.ps1", "tests/shared/core/fs.Tests.ps1", "scripts/shared/core/text.ps1"); expected = @("tests/shared/core/fs.Tests.ps1", "tests/meta/structure.Tests.ps1", "tests/meta/layers.Tests.ps1", "tests/shared/core/text.Tests.ps1") }
    ) {
        param($name, $files, $expected)
        (selectTests $files) -join "|" | Should -Be ($expected -join "|")
    }

    It "<name>" -TestCases @(
        @{ name = "対応するテストが無い scripts の .ps1 は全部"; files = @("scripts/tebunko/lib.ps1") }
        @{ name = "消した scripts の .ps1 は全部"; files = @("scripts/tebunko/search/none.ps1") }
        @{ name = "テストの実行口は全部"; files = @("tests/run.ps1") }
        @{ name = "テストの共通の準備は全部"; files = @("tests/helpers/load.ps1") }
        @{ name = "ほかの変更があっても全部"; files = @("README.md", "tests/testdata/README.txt") }
    ) {
        param($name, $files)
        (selectTests $files) -join "|" | Should -Be "all"
    }

    It "<name>" -TestCases @(
        @{ name = "CI の設定・画像はテストを流さない"; files = @(".github/workflows/test.yml", "tebunko.bat", "tools/hooks/pre-commit") }
        @{ name = "テストの無い道具は流さない"; files = @("tools/new_icon.ps1") }
        @{ name = "消したテストは流さない"; files = @("tests/shared/core/none.Tests.ps1") }
        @{ name = "変更が無ければ流さない"; files = @() }
    ) {
        param($name, $files)
        @(selectTests $files).Count | Should -Be 0
    }
}
