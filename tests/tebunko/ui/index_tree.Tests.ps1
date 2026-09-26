# ［2 検索］の検索対象インデックスのツリー（tebunko\ui\index_tree.ps1）のテスト。
# 画面の部品（$ui.IndexTree など）は偽物にし、インデックス・設定ファイルは TestDrive に作って確かめる。
. "$PSScriptRoot\..\..\helpers\load.ps1"
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
. "${scriptsDir}\shared\ui\types.ps1"
. "${scriptsDir}\tebunko\ui\types.ps1"

# ---- 画面の偽物 ----
# 読み込み時に登録されるイベントの処理は $handlers に、呼ばれた操作は $fake に取っておく
$handlers = @{}
$fake = @{ TargetUpdates = 0 }

function newFakeButton([string]$name) {
    $button = [pscustomobject]@{ PartName = $name }
    $button | Add-Member ScriptMethod Add_Click { param ($block) $handlers["$($this.PartName).Click"] = $block }
    return $button
}

$tree = [pscustomobject]@{ ItemsSource = $null; SelectedItem = $null }
$tree | Add-Member ScriptMethod AddHandler { param ($event, $handler) $handlers["IndexTree.$($event.Name)"] = $handler }
$tree | Add-Member ScriptMethod Add_PreviewKeyDown { param ($block) $handlers["IndexTree.PreviewKeyDown"] = $block }

$ui = [pscustomobject]@{
    IndexTree             = $tree
    IndexTreePlaceholder  = [pscustomobject]@{ Visibility = "Collapsed" }
    CheckAllIndexButton   = newFakeButton "CheckAll"
    UncheckAllIndexButton = newFakeButton "UncheckAll"
}

# 画面の共通部品（shared\ui\shell.ps1）と［2 検索］タブ（search_tab.ps1）の代わり
function safe {
    param ([scriptblock]$block)
    & $block
}
function updateSearchTarget {
    $fake.TargetUpdates++
}

. "${scriptsDir}\tebunko\ui\index_tree.ps1"

# 設定の保存先がリポジトリの setting.config・work\ にならないよう、どこにも無い場所にしておく（各 Describe で TestDrive に向け直す）
$safeDir = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko_index_tree_test_$([guid]::NewGuid())"
$indexDir = "$safeDir\index"
$settingsFile = "$safeDir\setting.config"
$statusFile = "$safeDir\取り込み一覧.tsv"
$workspace = newTestWorkspace @{ IndexDir = $indexDir; StatusFile = $statusFile }

# ---- テストの準備 ----

function newIndexFiles {
    # インデックス（work\index）を TestDrive に作る。
    #   営業部\直下.xlsx          … インデックス直下のファイル
    #   営業部\2024\見積.xlsx      … サブフォルダのファイル
    #   営業部\2024\東京\議事録.docx
    #   総務部\規程.docx
    param ([string]$dir)

    newTsv "$dir\営業部\直下.xlsx\Sheet1.tsv" @("`t見積")
    newTsv "$dir\営業部\2024\見積.xlsx\4月.tsv" @("`t見積")
    newTsv "$dir\営業部\2024\東京\議事録.docx\ページ001.tsv" @("見積")
    newTsv "$dir\総務部\規程.docx\ページ001.tsv" @("規程")
}

function writeSettingsJson {
    param ([string]$path, [string]$json)
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding $false))
}

function getRoot([string]$name) {
    return @($script:indexRoots | Where-Object { $_.Name -eq $name })[0]
}

Describe "読み込み" -Tag Unit {
    It "ツリーに一覧をつなぐ" {
        [object]::ReferenceEquals($ui.IndexTree.ItemsSource, $script:indexRoots) | Should Be $true
    }
}

Describe "describeSearchTargets" -Tag Unit {
    It "対象が無ければ空" {
        describeSearchTargets @() | Should Be ""
    }

    It "インデックス名からの相対パスをそのまま出し、直下だけのものは印を付ける" {
        $targets = @(
            [SearchTarget]@{ Root = "C:\tebunko\work\index"; RelPath = "営業部"; Recurse = $true }
            [SearchTarget]@{ Root = "C:\tebunko\work\index"; RelPath = "総務部\2024\"; Recurse = $false }
        )
        describeSearchTargets $targets | Should Be "営業部、総務部\2024（直下のファイル）"
    }

    It "<name>" -TestCases @(
        @{ name = "3 件までは並べる"; names = @("A", "B", "C"); expected = "A、B、C" }
        @{ name = "4 件以上は先頭の 3 件と残りの数"; names = @("A", "B", "C", "D", "E"); expected = "A、B、C ほか 2 か所" }
    ) {
        param ($name, $names, $expected)
        $targets = @($names | ForEach-Object { [SearchTarget]@{ RelPath = $_; Recurse = $true } })
        describeSearchTargets $targets | Should Be $expected
    }
}

Describe "loadIndexTree" -Tag Io {
    $indexDir = "$TestDrive\index"
    $settingsFile = "$TestDrive\setting.config"
    $statusFile = "$TestDrive\取り込み一覧.tsv"
    $workspace = newTestWorkspace @{ IndexDir = $indexDir; StatusFile = $statusFile }

    BeforeEach {
        Remove-Item -LiteralPath $indexDir, $settingsFile -Recurse -Force -ErrorAction SilentlyContinue
        $script:indexRoots.Clear()
        $fake.TargetUpdates = 0
    }

    It "インデックスが無ければ案内を出す" {
        loadIndexTree

        $script:indexRoots.Count | Should Be 0
        $ui.IndexTreePlaceholder.Visibility | Should Be "Visible"
        $fake.TargetUpdates | Should Be 1
        @(getSearchTargets).Count | Should Be 0
    }

    It "インデックスを 1 件ずつ一番上に並べ、設定の順にする" {
        newIndexFiles $indexDir
        writeSettingsJson $settingsFile '{ "targetFolders": [ { "name": "総務部", "path": "C:\\共有\\総務部", "enabled": true }, { "name": "営業部", "path": "C:\\共有\\営業部", "enabled": true } ] }'

        loadIndexTree

        (@($script:indexRoots | ForEach-Object { $_.Name }) -join ",") | Should Be "総務部,営業部"
        $ui.IndexTreePlaceholder.Visibility | Should Be "Collapsed"
        (getRoot "営業部").SourcePath | Should Be "C:\共有\営業部"
        (getRoot "営業部").ToolTip | Should Match "^元のフォルダ：C:\\共有\\営業部"
        isAllIndexChecked | Should Be $true
    }

    It "すべてチェックしていれば、インデックスごとに下のフォルダも含めて検索する" {
        newIndexFiles $indexDir
        loadIndexTree

        $targets = @(getSearchTargets)

        $targets.Count | Should Be 2
        (@($targets | ForEach-Object { "$($_.RelPath):$($_.Recurse)" }) -join ",") | Should Be "営業部:True,総務部:True"
        $targets[0].Root | Should Be (Resolve-Path -LiteralPath $indexDir).ProviderPath
    }

    It "保存したチェックなしのフォルダを戻す（下のフォルダごと・直下のファイルだけ）" {
        newIndexFiles $indexDir
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "$root\営業部\2024"; Subfolders = $true }
            [pscustomobject]@{ Path = "$root\総務部"; Subfolders = $false }
        ) $settingsFile

        loadIndexTree

        isAllIndexChecked | Should Be $false
        $null -eq (getRoot "営業部").IsChecked | Should Be $true
        # 総務部は直下のファイルしか無い（サブフォルダが無い）ので、外すとインデックスごと外れる
        (getRoot "総務部").IsChecked | Should Be $false
        $targets = @(getSearchTargets)
        (@($targets | ForEach-Object { "$($_.RelPath):$($_.Recurse)" }) -join ",") | Should Be "営業部:False"
        describeSearchTargets $targets | Should Be "営業部（直下のファイル）"
    }

    It "読み込み直しても、展開していたフォルダは展開したままにする" {
        newIndexFiles $indexDir
        loadIndexTree
        $sales = getRoot "営業部"
        $sales.SetExpanded($true)
        $year = @($sales.Children | Where-Object { $_.Name -eq "2024" })[0]
        $year.SetExpanded($true)

        loadIndexTree

        $sales = getRoot "営業部"
        $sales.IsExpanded | Should Be $true
        @($sales.Children | Where-Object { $_.Name -eq "2024" })[0].IsExpanded | Should Be $true
    }

    It "元のフォルダが分からないインデックスは、インデックスのフォルダを出す" {
        newIndexFiles $indexDir
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath

        loadIndexTree

        $sales = getRoot "営業部"
        $sales.SourcePath | Should Be ""
        $sales.ToolTip | Should Be "$root\営業部"
        $sales.LoadChildren()
        (@($sales.Children | ForEach-Object { $_.ToolTip }) -join "|") | Should Be "サブフォルダを除く、$root\営業部 の直下のファイル|$root\営業部\2024"
    }

    It "インデックスのフォルダが無くなっていても読み込める" {
        newIndexFiles $indexDir
        loadIndexTree
        Remove-Item -LiteralPath "$indexDir\総務部" -Recurse -Force

        loadIndexTree

        (@($script:indexRoots | ForEach-Object { $_.Name }) -join ",") | Should Be "営業部"
    }
}

Describe "saveSearchExcludes・setAllIndexChecked" -Tag Io {
    $indexDir = "$TestDrive\index"
    $settingsFile = "$TestDrive\setting.config"
    $statusFile = "$TestDrive\取り込み一覧.tsv"
    $workspace = newTestWorkspace @{ IndexDir = $indexDir; StatusFile = $statusFile }

    BeforeEach {
        Remove-Item -LiteralPath $indexDir, $settingsFile -Recurse -Force -ErrorAction SilentlyContinue
        $script:indexRoots.Clear()
        $fake.TargetUpdates = 0
        newIndexFiles $indexDir
        loadIndexTree
        $fake.TargetUpdates = 0
    }

    It "チェックを外したフォルダを保存し、検索対象の表示を更新する" {
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath
        (getRoot "総務部").SetChecked($false)

        onIndexTreeChecked

        $saved = @(readSearchExcludes $settingsFile)
        $saved.Count | Should Be 1
        $saved[0].Path | Should Be "$root\総務部"
        $saved[0].Subfolders | Should Be $true
        $fake.TargetUpdates | Should Be 1
    }

    It "チェックを戻すと、保存した記録も消す" {
        (getRoot "総務部").SetChecked($false)
        onIndexTreeChecked
        (getRoot "総務部").SetChecked($true)

        onIndexTreeChecked

        @(readSearchExcludes $settingsFile).Count | Should Be 0
    }

    It "見つからないインデックスのフォルダの記録は残す" {
        writeSearchExcludes @([pscustomobject]@{ Path = "\\server\共有\index\営業部"; Subfolders = $true }) $settingsFile

        saveSearchExcludes

        $saved = @(readSearchExcludes $settingsFile)
        $saved.Count | Should Be 1
        $saved[0].Path | Should Be "\\server\共有\index\営業部"
    }

    It "ツリーにあるインデックスの古い記録は、今のチェックの状態で置き換える" {
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath
        writeSearchExcludes @([pscustomobject]@{ Path = "$root\営業部\2024"; Subfolders = $true }) $settingsFile

        saveSearchExcludes

        @(readSearchExcludes $settingsFile).Count | Should Be 0
    }

    It "すべて外す・すべてチェックする" {
        setAllIndexChecked $false
        @(getSearchTargets).Count | Should Be 0
        @(readSearchExcludes $settingsFile).Count | Should Be 2
        isAllIndexChecked | Should Be $false

        setAllIndexChecked $true
        @(getSearchTargets).Count | Should Be 2
        @(readSearchExcludes $settingsFile).Count | Should Be 0
        isAllIndexChecked | Should Be $true
        $fake.TargetUpdates | Should Be 2
    }

    It "［すべて外す］［すべてチェック］のボタンから切り替える" {
        & $handlers["UncheckAll.Click"]
        isAllIndexChecked | Should Be $false

        & $handlers["CheckAll.Click"]
        isAllIndexChecked | Should Be $true
    }
}

Describe "イベント" -Tag Io {
    $indexDir = "$TestDrive\index"
    $settingsFile = "$TestDrive\setting.config"
    $statusFile = "$TestDrive\取り込み一覧.tsv"
    $workspace = newTestWorkspace @{ IndexDir = $indexDir; StatusFile = $statusFile }

    BeforeEach {
        Remove-Item -LiteralPath $indexDir, $settingsFile -Recurse -Force -ErrorAction SilentlyContinue
        $script:indexRoots.Clear()
        newIndexFiles $indexDir
        loadIndexTree
        $fake.TargetUpdates = 0
        $ui.IndexTree.SelectedItem = $null
    }

    It "スペースで選んでいるフォルダのチェックを切り替えて保存する" {
        $ui.IndexTree.SelectedItem = getRoot "総務部"
        $e = [pscustomobject]@{ Key = "Space"; Handled = $false }

        & $handlers["IndexTree.PreviewKeyDown"] $ui.IndexTree $e

        $e.Handled | Should Be $true
        (getRoot "総務部").IsChecked | Should Be $false
        @(readSearchExcludes $settingsFile).Count | Should Be 1
        $fake.TargetUpdates | Should Be 1
    }

    It "スペース以外のキー・何も選んでいないとき・「読み込み中…」では切り替えない" {
        $ui.IndexTree.SelectedItem = getRoot "総務部"
        $other = [pscustomobject]@{ Key = "Enter"; Handled = $false }
        & $handlers["IndexTree.PreviewKeyDown"] $ui.IndexTree $other

        $ui.IndexTree.SelectedItem = $null
        $none = [pscustomobject]@{ Key = "Space"; Handled = $false }
        & $handlers["IndexTree.PreviewKeyDown"] $ui.IndexTree $none

        $ui.IndexTree.SelectedItem = [IndexNode]::NewPlaceholder((getRoot "営業部"))
        $placeholder = [pscustomobject]@{ Key = "Space"; Handled = $false }
        & $handlers["IndexTree.PreviewKeyDown"] $ui.IndexTree $placeholder

        $other.Handled | Should Be $false
        $none.Handled | Should Be $false
        $placeholder.Handled | Should Be $false
        isAllIndexChecked | Should Be $true
        $fake.TargetUpdates | Should Be 0
    }

    It "フォルダを展開したら子を読み込む" {
        $sales = getRoot "営業部"
        $sales.Children.Count | Should Be 1
        $sales.Children[0].IsPlaceholder | Should Be $true
        $e = New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.TreeViewItem]::ExpandedEvent)
        $e.Source = [pscustomobject]@{ DataContext = $sales }

        $handlers["IndexTree.Expanded"].Invoke($ui.IndexTree, $e)

        (@($sales.Children | ForEach-Object { $_.Name }) -join ",") | Should Be "（このフォルダ直下のファイル）,2024"
    }

    It "チェックボックス以外のクリックではチェックを変えない" {
        $e = New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
        $e.Source = [pscustomobject]@{ DataContext = (getRoot "総務部") }

        $handlers["IndexTree.Click"].Invoke($ui.IndexTree, $e)

        isAllIndexChecked | Should Be $true
        $fake.TargetUpdates | Should Be 0
    }
}
