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
        @{ name = "XAML は構成・表の外枠・設定ファイルの欄のテスト"; files = @("scripts/tebunko/xaml/search/search.xaml"); expected = @("tests/meta/structure.Tests.ps1", "tests/meta/table_border.Tests.ps1", "tests/meta/settings_file_text.Tests.ps1") }
        @{ name = "テストを変えたらそのテスト"; files = @("tests/shared/core/text.Tests.ps1"); expected = @("tests/shared/core/text.Tests.ps1") }
        @{ name = "Markdown と docs の中は links"; files = @("README.md", "docs/images/none.png"); expected = @("tests/meta/links.Tests.ps1") }
        @{ name = "道具はそのテスト。リンクの検査は links も"; files = @("tools/check_signoff.ps1", "tools/check_markdown_links.ps1"); expected = @("tests/tools/check_signoff.Tests.ps1", "tests/tools/check_markdown_links.Tests.ps1", "tests/meta/links.Tests.ps1") }
        @{ name = "既定のワークスペースを守る共通の関数は run_isolated のテスト"; files = @("tools/isolation/isolation_common.ps1"); expected = @("tests/tools/run_isolated.Tests.ps1") }
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
        @{ name = "画面のテスト（tests/gui/）は選ばない。CI（gui.yml）と手元の -Tag Gui で流す"; files = @("tests/gui/smoke.Tests.ps1", "tests/gui/gui_helpers.ps1") }
    ) {
        param($name, $files)
        @(selectTests $files).Count | Should -Be 0
    }
}

Describe "run_commit_tests.ps1 の静的解析の対象" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "scripts・tools・tests の .ps1 だけが対象（/ 区切りにそろえる）"; files = @("tools\check_analyzer.ps1", "tests/tools/check_analyzer.Tests.ps1", "README.md", "tools/hooks/pre-commit"); expected = "tools/check_analyzer.ps1|tests/tools/check_analyzer.Tests.ps1" }
        @{ name = ".md と docs だけなら対象なし"; files = @("README.md", "docs/design/testing/ci.md"); expected = "" }
        @{ name = "消した・名前を変えて元が無い .ps1 は飛ばす"; files = @("tools/none.ps1"); expected = "" }
        @{ name = "変更が無ければ対象なし"; files = @(); expected = "" }
    ) {
        param($name, $files, $expected)
        $out = if ($files.Count -gt 0) { @(& $select -ListAnalyzer -Files $files) } else { @(& $select -ListAnalyzer -Files @()) }
        ($out -join "|") | Should -Be $expected
    }
}

# 静的解析の呼び出し。check_analyzer.ps1 の代わりに、受け取った -Path を出して終了コードを環境変数で返す偽物を使う
Describe "run_commit_tests.ps1 の静的解析の呼び出し" -Tag Io {
    BeforeAll {
        $script:work = Join-Path ([System.IO.Path]::GetTempPath()) ("commit_analyzer_" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $script:work | Out-Null
        $script:fake = Join-Path $script:work "fake_analyzer.ps1"
        Set-Content -LiteralPath $script:fake -Value 'param([string[]]$Path)
"CALLED:" + ($Path -join ",") + " V:" + $env:PSSA_VERSION
exit [int]$env:FAKE_ANALYZER_EXIT'
        $script:runner = (Resolve-Path $select).Path

        # PSScriptAnalyzer の入っていない手元でも流せるよう、既定ではどの機械にも入っている Pester を「静的解析のモジュール」にする
        $script:pesterVersion = @(Get-Module -ListAvailable -Name Pester | Select-Object -First 1)[0].Version.ToString()

        function runHook([string[]]$files, [string[]]$extra, [int]$fakeExit = 0, [int]$testExit = 0) {
            $env:FAKE_ANALYZER_EXIT = "$fakeExit"
            $given = @($extra)
            if ($given -notcontains "-AnalyzerModule") { $given += @("-AnalyzerModule", "Pester") }
            if ($given -notcontains "-AnalyzerVersion") { $given += @("-AnalyzerVersion", $script:pesterVersion) }
            try {
                $out = & powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $script:runner -TestExit $testExit -Files ($files -join ",") -AnalyzerScript $script:fake @given 2>&1
                return @{ Code = $LASTEXITCODE; Text = ($out | Out-String) }
            } finally {
                Remove-Item Env:\FAKE_ANALYZER_EXIT -ErrorAction SilentlyContinue
            }
        }
    }
    AfterAll { Remove-Item -LiteralPath $script:work -Recurse -Force -ErrorAction SilentlyContinue }

    It "ステージに .ps1 があれば、そのファイルを -Path に付けて呼ぶ" {
        $r = runHook @("tools/check_analyzer.ps1", "README.md")
        $r.Code | Should -Be 0
        $r.Text | Should -Match "CALLED:tools/check_analyzer\.ps1"
    }

    It "<name>" -TestCases @(
        @{ name = "呼んだ静的解析が失敗したら、同じ終了コードで失敗する"; testExit = 0; fakeExit = 1; code = 1 }
        @{ name = "テストが落ちても静的解析は流し、テストの終了コードを返す"; testExit = 2; fakeExit = 0; code = 2 }
        @{ name = "テストも静的解析も落ちたら、テストの終了コードを返す"; testExit = 2; fakeExit = 1; code = 2 }
        @{ name = "どちらも通れば 0"; testExit = 0; fakeExit = 0; code = 0 }
    ) {
        param($name, $testExit, $fakeExit, $code)
        $r = runHook @("tools/check_analyzer.ps1") @() $fakeExit $testExit
        $r.Code | Should -Be $code
        $r.Text | Should -Match "CALLED"
    }

    It "既定の版は CI が固定した版（test.yml の PSSA_VERSION）" {
        $pinned = @(& powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $script:runner -ListAnalyzerVersion -Files "README.md")[0]
        $pinned | Should -Match '^\d+(\.\d+)+$'
        @(Select-String -LiteralPath "$PSScriptRoot\..\..\.github\workflows\test.yml" -Pattern "PSSA_VERSION:\s*$([regex]::Escape($pinned))\s*$").Count | Should -Be 1
    }

    It "入っている版を PSSA_VERSION に入れて子に渡す" {
        $r = runHook @("tools/check_analyzer.ps1")
        $r.Code | Should -Be 0
        $r.Text | Should -Match "CALLED:tools/check_analyzer\.ps1 V:$([regex]::Escape($script:pesterVersion))"
    }

    It "固定した版が手元に入っていないときは、呼ばずに注意を出して 0" {
        $r = runHook @("tools/check_analyzer.ps1") @("-AnalyzerVersion", "0.0.1") 1
        $r.Code | Should -Be 0
        $r.Text | Should -Not -Match "CALLED"
        $r.Text | Should -Match "飛ばします"
    }

    It ".md だけなら呼ばない" {
        $r = runHook @("README.md", "docs/design/testing/ci.md") @() 1
        $r.Code | Should -Be 0
        $r.Text | Should -Not -Match "CALLED"
    }

    It "<name>" -TestCases @(
        @{ name = "PSScriptAnalyzer が無いときは、呼ばずに注意を出して 0"; extra = @("-AnalyzerModule", "NoSuchModule_for_test") }
        @{ name = "Windows PowerShell 以外で動いたときは、呼ばずに注意を出して 0"; extra = @("-Edition", "Core") }
    ) {
        param($name, $extra)
        $r = runHook @("tools/check_analyzer.ps1") $extra 1
        $r.Code | Should -Be 0
        $r.Text | Should -Not -Match "CALLED"
        $r.Text | Should -Match "飛ばします"
    }
}
