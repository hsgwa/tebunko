# 層の境界のテスト。境界の向こう側を、決まった答えを返すスタブに替えて、こちら側の関数が
# 「向こう側の答えのとおりに動くこと」と「向こう側へ渡すものが正しいこと」だけを確かめる。
#   判断層 → 状態層: 判断層の関数が呼ぶ状態層の関数をスタブにする（状態層の中身が変わっても、判断層のテストは変わらない）
#   画面層 → 判断層: 画面層の関数が呼ぶ判断層の関数をスタブにする（画面の部品は偽物。判断の中身は判断層のテストが見る）
# 依存の向きそのものは tests\meta\layers.Tests.ps1 が見る。ここは、その向きの境界で受け渡しが守られているかを見る。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationCore, WindowsBase
    . "${scriptsDir}\tebunko\ui\search\search_bar_view.ps1"
    . "${scriptsDir}\tebunko\ui\settings\settings_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_archive_view.ps1"

    # 画面層のファイルは、読み込むと画面の部品を作るため、使う関数だけを取り出す
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\search\search_bar.ps1", [ref]$tokens, [ref]$errors)
    foreach ($name in "updateFastSearchView", "getCurrentWordNotice") {
        $definition = $ast.Find({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }
}

Describe "判断層 → 状態層（状態層をスタブにする）" -Tag Unit {
    Context "getWordNotice（search_query.ps1 の isValidRegex）" {
        It "正規表現を使うときは、状態層の答えが偽なら注意を返し、真なら空を返す。調べるのは入力した語" -ForEach @(
            @{ valid = $false; expected = "正規表現が正しくありません" }
            @{ valid = $true; expected = "" }
        ) {
            Mock isValidRegex { return $valid }
            getWordNotice "(語" $true | Should -Be $expected
            Should -Invoke isValidRegex -Times 1 -Exactly -ParameterFilter { $pattern -eq "(語" }
        }

        It "正規表現を使わないとき・語が空のときは、状態層に聞かない" -ForEach @(
            @{ word = "(語"; useRegex = $false }
            @{ word = ""; useRegex = $true }
        ) {
            Mock isValidRegex { return $false }
            getWordNotice $word $useRegex | Should -Be ""
            Should -Invoke isValidRegex -Times 0 -Exactly
        }
    }

    Context "getFastSearchView（search_gram.ps1 の testFastSearchUsable）" {
        It "使えるかは状態層の答えのとおりにし、使えないとき（Windows Search あり）は理由を正規表現・語の長さで分ける" -ForEach @(
            @{ usable = $true; searchable = $true; regex = $false; goToIndex = $false; text = "高速検索：使用可" }
            @{ usable = $false; searchable = $true; regex = $true; goToIndex = $false; text = "高速検索：使用不可" }
            @{ usable = $false; searchable = $false; regex = $false; goToIndex = $true; text = "高速検索：使用不可" }
        ) {
            Mock testFastSearchUsable { return $usable }
            $view = getFastSearchView $searchable $regex "語"
            $view.Usable | Should -Be $usable
            $view.Text | Should -Be $text
            $view.GoToIndex | Should -Be $goToIndex
            $view.Tip | Should -Not -BeNullOrEmpty
            Should -Invoke testFastSearchUsable -Times 1 -Exactly -ParameterFilter { $available -eq $searchable -and $useRegex -eq $regex -and $word -eq "語" }
        }

        It "まだ確かめていない（`$null）ときは、使えるものとして状態層に渡す" {
            Mock testFastSearchUsable { return $true }
            (getFastSearchView $null $false "語").Usable | Should -Be $true
            Should -Invoke testFastSearchUsable -Times 1 -Exactly -ParameterFilter { $available -eq $true }
        }
    }

    Context "getCountedEntryPaths（core/settings.ps1 の testSettingsFileName）" {
        It "既定のワークスペースでは、状態層が設定ファイルの名前と答えたものを数えない。渡すのは名前だけ（フォルダは付けない）" {
            Mock testSettingsFileName { return ($name -eq "setting.config") }
            $kept = @(getCountedEntryPaths @("C:\ws\setting.config", "C:\ws\営業") $true)
            $kept | Should -Be @("C:\ws\営業")
            Should -Invoke testSettingsFileName -Times 1 -Exactly -ParameterFilter { $name -eq "setting.config" }
            Should -Invoke testSettingsFileName -Times 1 -Exactly -ParameterFilter { $name -eq "営業" }
        }

        It "既定のワークスペースでないときは、状態層に聞かず、そのまま返す" {
            Mock testSettingsFileName { return $true }
            @(getCountedEntryPaths @("C:\ws\setting.config", "C:\ws\営業") $false).Count | Should -Be 2
            Should -Invoke testSettingsFileName -Times 0 -Exactly
        }
    }

    Context "testIndexImportInput（index_name.ps1 の testIndexName・index_archive_rules.ps1 の getIndexFolderConflict）" {
        It "名前の判定が理由を返したら、それをそのまま返し、フォルダの重なりは調べない" {
            Mock testIndexName { return "名前が使えません" }
            Mock getIndexFolderConflict { return "重なっています" }
            testIndexImportInput "C:\data\営業" " 営業 " @() | Should -Be "名前が使えません"
            Should -Invoke testIndexName -Times 1 -Exactly -ParameterFilter { $name -eq "営業" }
            Should -Invoke getIndexFolderConflict -Times 0 -Exactly
        }

        It "名前が通れば、フォルダの重なりの判定の答えをそのまま返す" -ForEach @(
            @{ conflict = "" }
            @{ conflict = "重なっています" }
        ) {
            Mock testIndexName { return "" }
            Mock getIndexFolderConflict { return $conflict }
            testIndexImportInput "C:\data\営業" "営業" @() | Should -Be $conflict
            Should -Invoke getIndexFolderConflict -Times 1 -Exactly -ParameterFilter { $folder -eq "C:\data\営業" }
        }
    }
}

Describe "画面層 → 判断層（判断層をスタブにする）" -Tag Unit {
    BeforeAll {
        function getWordText { return $script:word }
        function newFakeUi {
            return @{
                RegexCheck = [pscustomobject]@{ IsChecked = $true }
                FastSearchText = [pscustomobject]@{ Text = "" }
                FastBadge = [pscustomobject]@{ Tag = ""; ToolTip = ""; Cursor = $null }
            }
        }
    }

    BeforeEach {
        $script:word = "語"
        $script:fastAvailable = $null
        $ui = newFakeUi
    }

    Context "updateFastSearchView（search_bar_view.ps1 の getFastSearchView）" {
        It "判断層の答えを、そのまま画面の部品に置く。判断層には、確かめた結果・［正規表現］・検索ワードを渡す" -ForEach @(
            @{ usable = $true; goToIndex = $false; tag = "ok"; cursorName = "Arrow" }
            @{ usable = $false; goToIndex = $true; tag = "off"; cursorName = "Hand" }
        ) {
            $script:fastAvailable = $false
            Mock getFastSearchView { return @{ Usable = $usable; Text = "判断層の文言"; Tip = "判断層の説明"; GoToIndex = $goToIndex } }

            updateFastSearchView

            $ui.FastSearchText.Text | Should -Be "判断層の文言"
            $ui.FastBadge.ToolTip | Should -Be "判断層の説明"
            $ui.FastBadge.Tag | Should -Be $tag
            [object]::ReferenceEquals($ui.FastBadge.Cursor, [System.Windows.Input.Cursors]::$cursorName) | Should -Be $true
            $script:fastGoToIndex | Should -Be $goToIndex
            Should -Invoke getFastSearchView -Times 1 -Exactly -ParameterFilter { $available -eq $false -and $useRegex -eq $true -and $word -eq "語" }
        }
    }

    Context "getCurrentWordNotice（search_bar_view.ps1 の getWordNotice）" {
        It "判断層の答えをそのまま返す。判断層には、検索ワードと［正規表現］の状態を渡す" {
            Mock getWordNotice { return "判断層の注意" }
            $ui.RegexCheck.IsChecked = $false

            getCurrentWordNotice | Should -Be "判断層の注意"
            Should -Invoke getWordNotice -Times 1 -Exactly -ParameterFilter { $word -eq "語" -and $useRegex -eq $false }
        }
    }
}
