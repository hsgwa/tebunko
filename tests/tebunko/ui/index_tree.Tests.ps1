# ［検索］の検索対象インデックスのツリー（tebunko\ui\index_tree.ps1）のテスト。
# 画面の部品（$ui.IndexTree など）は偽物にし、インデックス・設定ファイルは TestDrive に作って確かめる。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\target_types.ps1"
    . "${scriptsDir}\tebunko\ui\search\search_bar_view.ps1"

    # ---- 画面の偽物 ----
    # 読み込み時に登録されるイベントの処理は $handlers に、呼ばれた操作は $fake に取っておく
    $handlers = @{}
    $script:sourceFolderMaps = @{}   # 元のフォルダの対応のキャッシュ（［検索］の画面の search_session.ps1 が持つ）
    $fake = @{ TargetUpdates = 0; Statuses = (New-Object System.Collections.Generic.List[string]); Jobs = (New-Object System.Collections.Generic.List[object]); ErrorLogs = (New-Object System.Collections.Generic.List[string]) }

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
        CheckAllIndexButton   = newFakeButton "CheckAll"
        UncheckAllIndexButton = newFakeButton "UncheckAll"
    }

    # 画面の共通部品（shared\ui\shell.ps1）と［検索］の検索条件の画面（ui\search\search_bar.ps1）の代わり
    function safe {
        param ([scriptblock]$block)
        & $block
    }
    function updateSearchTarget {
        $fake.TargetUpdates++
    }
    function setStatus {
        param ([string]$text)
        $fake.Statuses.Add($text)
    }
    function writeErrorLog {
        param ([string]$context, $record)
        $fake.ErrorLogs.Add("$context : $($record.Exception.Message)")
    }
    function startJob {
        # 裏の列には出さず、依頼を取っておく（テストが onDone を呼んで、終わったことにする）
        param ([scriptblock]$scriptBlock, [object[]]$arguments, [scriptblock]$onDone, [string]$queue = "default")
        $fake.Jobs.Add(@{ ScriptBlock = $scriptBlock; Arguments = $arguments; OnDone = $onDone; Queue = $queue })
    }

    # 画面のスレッドの分割（startUiWork ほか）。タイマーは偽物で、結果を見る前に completeUiWork で終わらせる
    . "$PSScriptRoot\..\..\helpers\ui_work.ps1"
    function reportUnexpectedError {
        param ([string]$context, $record)
        $fake.ErrorLogs.Add("report : $($record.Exception.Message)")
    }

    . "${scriptsDir}\tebunko\ui\index_tree.ps1"

    function loadIndexTreeNow {
        # loadIndexTree は裏の列に仕事を出すだけなので、出された仕事をその場で流し、結果を onDone に渡して終わったことにする
        loadIndexTree
        $job = $fake.Jobs[$fake.Jobs.Count - 1]
        $jobArguments = $job.Arguments
        $output = @(& $job.ScriptBlock @jobArguments)
        & $job.OnDone $output ""
        [void]$fake.Jobs.Remove($job)
        completeUiWork indexTree
    }

    # 設定の保存先がリポジトリの setting.config・work\ にならないよう、どこにも無い場所にしておく（各 Describe で TestDrive に向け直す）
    $safeDir = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko_index_tree_test_$([guid]::NewGuid())"
    $indexDir = "$safeDir\index"
    $settingsFile = "$safeDir\setting.config"
    $statusFile = "$safeDir\ingest_status.tsv"
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

    function newTreeData([string[]]$names) {
        return @{
            State = ${pathStateFound}; Message = ""; Root = "C:\work\index"; Sources = @{}; OldIndexCount = 0; Children = @{}
            Indexes = @($names | ForEach-Object { @{ Name = $_; SourcePath = ""; Exists = $false; HasSubfolders = $false } })
        }
    }

    function getRoot([string]$name) {
        return @($script:indexRoots | Where-Object { $_.Name -eq $name })[0]
    }
}

Describe "読み込み" -Tag Unit {
    It "ツリーに一覧をつなぐ" {
        [object]::ReferenceEquals($ui.IndexTree.ItemsSource, $script:indexRoots) | Should -Be $true
    }
}

Describe "describeSearchTargets" -Tag Unit {
    It "対象が無ければ空" {
        describeSearchTargets @() | Should -Be ""
    }

    It "インデックス名からの相対パスをそのまま出し、直下だけのものは印を付ける" {
        $targets = @(
            [SearchTarget]@{ Root = "C:\tebunko\work\index"; RelPath = "営業部"; Recurse = $true }
            [SearchTarget]@{ Root = "C:\tebunko\work\index"; RelPath = "総務部\2024\"; Recurse = $false }
        )
        describeSearchTargets $targets | Should -Be "営業部、総務部\2024（直下のファイル）"
    }

    It "<name>" -TestCases @(
        @{ name = "3 件までは並べる"; names = @("A", "B", "C"); expected = "A、B、C" }
        @{ name = "4 件以上は先頭の 3 件と残りの数"; names = @("A", "B", "C", "D", "E"); expected = "A、B、C ほか 2 か所" }
    ) {
        param ($name, $names, $expected)
        $targets = @($names | ForEach-Object { [SearchTarget]@{ RelPath = $_; Recurse = $true } })
        describeSearchTargets $targets | Should -Be $expected
    }
}

Describe "loadIndexTree" -Tag Io {
    BeforeAll {
        $indexDir = "$TestDrive\index"
        $settingsFile = "$TestDrive\setting.config"
        $statusFile = "$TestDrive\ingest_status.tsv"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir; StatusFile = $statusFile }
    }

    BeforeEach {
        Remove-Item -LiteralPath $indexDir, $settingsFile -Recurse -Force -ErrorAction SilentlyContinue
        $script:indexRoots.Clear()
        $fake.TargetUpdates = 0
    }

    It "インデックスが無ければ空のまま対象を更新する" {
        loadIndexTreeNow

        $script:indexRoots.Count | Should -Be 0
        # 読み込みを始めたとき（読み込み中の表示）と、結果を反映したときの 2 回
        $fake.TargetUpdates | Should -Be 2
        @(getSearchTargets).Count | Should -Be 0
    }

    It "インデックスを 1 件ずつ一番上に並べ、設定の順にする" {
        newIndexFiles $indexDir
        writeSettingsJson $settingsFile '{ "targetFolders": [ { "name": "総務部", "path": "C:\\共有\\総務部", "enabled": true }, { "name": "営業部", "path": "C:\\共有\\営業部", "enabled": true } ] }'

        loadIndexTreeNow

        (@($script:indexRoots | ForEach-Object { $_.Name }) -join ",") | Should -Be "総務部,営業部"
        (getRoot "営業部").SourcePath | Should -Be "C:\共有\営業部"
        (getRoot "営業部").ToolTip | Should -Match "^元のフォルダ：C:\\共有\\営業部"
        isAllIndexChecked | Should -Be $true
    }

    It "すべてチェックしていれば、インデックスごとに下のフォルダも含めて検索する" {
        newIndexFiles $indexDir
        loadIndexTreeNow

        $targets = @(getSearchTargets)

        $targets.Count | Should -Be 2
        (@($targets | ForEach-Object { "$($_.RelPath):$($_.Recurse)" }) -join ",") | Should -Be "営業部:True,総務部:True"
        $targets[0].Root | Should -Be (Resolve-Path -LiteralPath $indexDir).ProviderPath
    }

    It "保存したチェックなしのフォルダを戻す（下のフォルダごと・直下のファイルだけ）" {
        newIndexFiles $indexDir
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "$root\営業部\2024"; Subfolders = $true }
            [pscustomobject]@{ Path = "$root\総務部"; Subfolders = $false }
        ) $settingsFile

        loadIndexTreeNow

        isAllIndexChecked | Should -Be $false
        $null -eq (getRoot "営業部").IsChecked | Should -Be $true
        # 総務部は直下のファイルしか無い（サブフォルダが無い）ので、外すとインデックスごと外れる
        (getRoot "総務部").IsChecked | Should -Be $false
        $targets = @(getSearchTargets)
        (@($targets | ForEach-Object { "$($_.RelPath):$($_.Recurse)" }) -join ",") | Should -Be "営業部:False"
        describeSearchTargets $targets | Should -Be "営業部（直下のファイル）"
    }

    It "読み込み直しても、展開していたフォルダは展開したままにする" {
        newIndexFiles $indexDir
        loadIndexTreeNow
        $sales = getRoot "営業部"
        $sales.SetExpanded($true)
        expandIndexNode $sales
        $year = @($sales.Children | Where-Object { $_.Name -eq "2024" })[0]
        $year.SetExpanded($true)

        loadIndexTreeNow

        $sales = getRoot "営業部"
        $sales.IsExpanded | Should -Be $true
        @($sales.Children | Where-Object { $_.Name -eq "2024" })[0].IsExpanded | Should -Be $true
    }

    It "元のフォルダが分からないインデックスは、インデックスのフォルダを出す" {
        newIndexFiles $indexDir
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath

        loadIndexTreeNow

        $sales = getRoot "営業部"
        $sales.SourcePath | Should -Be ""
        $sales.ToolTip | Should -Be "$root\営業部"
        expandIndexNode $sales
        (@($sales.Children | ForEach-Object { $_.ToolTip }) -join "|") | Should -Be "サブフォルダを除く、$root\営業部 の直下のファイル|$root\営業部\2024"
    }

    It "インデックスのフォルダが無くなっていても読み込める" {
        newIndexFiles $indexDir
        loadIndexTreeNow
        Remove-Item -LiteralPath "$indexDir\総務部" -Recurse -Force

        loadIndexTreeNow

        (@($script:indexRoots | ForEach-Object { $_.Name }) -join ",") | Should -Be "営業部"
    }

    Context "反映の分割" {
        BeforeEach {
            $script:sourceFolderMaps = @{}
            $script:indexTreeLoading = $false
            $fake.ErrorLogs.Clear()
            $script:clock = 0
            # 1 件回すたびに予算切れにする（1 回の予算で 1 件だけ回る）
            Mock getUiWorkMilliseconds { $script:clock += 100; $script:clock }
        }

        AfterEach {
            stopUiWork indexTree
        }

        It "applyIndexTreeData のあと completeUiWork で、根が設定の順にそろう" {
            applyIndexTreeData (newTreeData @("a", "b", "c")) @()

            $script:indexRoots.Count | Should -Be 1
            completeUiWork indexTree

            (@($script:indexRoots | ForEach-Object { $_.Name }) -join ",") | Should -Be "a,b,c"
        }

        It "終わるまで indexTreeLoading は真で、終わると偽になり、対象を更新する" {
            $before = $fake.TargetUpdates

            applyIndexTreeData (newTreeData @("a", "b")) @()

            $script:indexTreeLoading | Should -Be $true
            testUiWorkRunning indexTree | Should -Be $true
            $fake.TargetUpdates | Should -Be $before
            completeUiWork indexTree
            $script:indexTreeLoading | Should -Be $false
            $fake.TargetUpdates | Should -Be ($before + 1)
        }

        It "途中で呼び直すと、前の分の根は残らない" {
            applyIndexTreeData (newTreeData @("a", "b", "c")) @()

            applyIndexTreeData (newTreeData @("x", "y")) @()
            completeUiWork indexTree

            (@($script:indexRoots | ForEach-Object { $_.Name }) -join ",") | Should -Be "x,y"
        }

        It "OnItem が例外を投げたら、知らせて止まり、indexTreeLoading が偽に戻る" {
            Mock applyIndexTreeChildren { throw "壊れた" }

            applyIndexTreeData (newTreeData @("a", "b")) @()
            completeUiWork indexTree

            $script:indexTreeLoading | Should -Be $false
            testUiWorkRunning indexTree | Should -Be $false
            @($fake.ErrorLogs | Where-Object { $_ -like "report*壊れた" }).Count | Should -Be 1
        }
    }
}

Describe "saveSearchExcludes・setAllIndexChecked" -Tag Io {
    BeforeAll {
        $indexDir = "$TestDrive\index"
        $settingsFile = "$TestDrive\setting.config"
        $statusFile = "$TestDrive\ingest_status.tsv"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir; StatusFile = $statusFile }
    }

    BeforeEach {
        Remove-Item -LiteralPath $indexDir, $settingsFile -Recurse -Force -ErrorAction SilentlyContinue
        $script:indexRoots.Clear()
        $fake.TargetUpdates = 0
        newIndexFiles $indexDir
        loadIndexTreeNow
        $fake.TargetUpdates = 0
    }

    It "チェックを外したフォルダを保存し、検索対象の表示を更新する" {
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath
        (getRoot "総務部").SetChecked($false)

        onIndexTreeChecked

        $saved = @(readSearchExcludes $settingsFile)
        $saved.Count | Should -Be 1
        $saved[0].Path | Should -Be "$root\総務部"
        $saved[0].Subfolders | Should -Be $true
        $fake.TargetUpdates | Should -Be 1
    }

    It "チェックを戻すと、保存した記録も消す" {
        (getRoot "総務部").SetChecked($false)
        onIndexTreeChecked
        (getRoot "総務部").SetChecked($true)

        onIndexTreeChecked

        @(readSearchExcludes $settingsFile).Count | Should -Be 0
    }

    It "見つからないインデックスのフォルダの記録は残す" {
        writeSearchExcludes @([pscustomobject]@{ Path = "\\server\共有\index\営業部"; Subfolders = $true }) $settingsFile

        saveSearchExcludes

        $saved = @(readSearchExcludes $settingsFile)
        $saved.Count | Should -Be 1
        $saved[0].Path | Should -Be "\\server\共有\index\営業部"
    }

    It "ツリーにあるインデックスの古い記録は、今のチェックの状態で置き換える" {
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath
        writeSearchExcludes @([pscustomobject]@{ Path = "$root\営業部\2024"; Subfolders = $true }) $settingsFile

        saveSearchExcludes

        @(readSearchExcludes $settingsFile).Count | Should -Be 0
    }

    It "名前が前方一致する別の根の記録は、存在しない根の下のものを消さない" {
        $root = (Resolve-Path -LiteralPath $indexDir).ProviderPath
        newTsv "$indexDir\営業\規程.docx\ページ001.tsv" @("規程")
        $script:indexRoots.Clear()
        loadIndexTreeNow
        # 営業部 は存在しない根（ネットワークが切れているなど）。営業 が存在する根
        (getRoot "営業部").Exists = $false
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "$root\営業部\2024"; Subfolders = $true }
            [pscustomobject]@{ Path = "$root\営業\古い"; Subfolders = $true }
        ) $settingsFile
        (getRoot "営業").SetChecked($false)

        saveSearchExcludes

        $saved = @(readSearchExcludes $settingsFile | ForEach-Object { $_.Path })
        ($saved -join "|") | Should -Be "$root\営業部\2024|$root\営業"
    }

    It "保存する一覧が今の保存と同じなら、設定ファイルを書き直さない" {
        (getRoot "総務部").SetChecked($false)
        onIndexTreeChecked
        Mock writeSearchExcludes {}

        saveSearchExcludes

        Should -Invoke writeSearchExcludes -Times 0 -Exactly
    }

    It "すべて外す・すべてチェックする" {
        setAllIndexChecked $false
        @(getSearchTargets).Count | Should -Be 0
        @(readSearchExcludes $settingsFile).Count | Should -Be 2
        isAllIndexChecked | Should -Be $false

        setAllIndexChecked $true
        @(getSearchTargets).Count | Should -Be 2
        @(readSearchExcludes $settingsFile).Count | Should -Be 0
        isAllIndexChecked | Should -Be $true
        $fake.TargetUpdates | Should -Be 2
    }

    It "［すべて外す］［すべてチェック］のボタンから切り替える" {
        & $handlers["UncheckAll.Click"]
        isAllIndexChecked | Should -Be $false

        & $handlers["CheckAll.Click"]
        isAllIndexChecked | Should -Be $true
    }
}

Describe "applyIndexTreeData（根ごとの振り分け）" -Tag Io {
    BeforeAll {
        $settingsFile = "$TestDrive\setting.config"
        $workspace = newTestWorkspace @{ IndexDir = "C:\work\index"; StatusFile = "C:\work\ingest_status.tsv" }
    }

    BeforeEach {
        $script:indexRoots.Clear()
        $script:sourceFolderMaps = @{}
    }

    It "除外と展開は、パスの属する根にだけ効く（名前が前方一致する根・存在しない根を含む）" {
        $r = "C:\work\index"
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "$r\営業部"; Subfolders = $true }
            [pscustomobject]@{ Path = "$r\閉鎖"; Subfolders = $true }
            [pscustomobject]@{ Path = "D:\ほか\営業"; Subfolders = $true }
        ) $settingsFile
        $data = @{
            State = ${pathStateFound}; Message = ""; Root = $r; Sources = @{}; OldIndexCount = 0; Children = @{}
            Indexes = @(
                @{ Name = "営業"; SourcePath = ""; Exists = $true; HasSubfolders = $false }
                @{ Name = "営業部"; SourcePath = ""; Exists = $true; HasSubfolders = $false }
                @{ Name = "閉鎖"; SourcePath = ""; Exists = $false; HasSubfolders = $false }
            )
        }

        applyIndexTreeData $data @("$r\営業", "$r\閉鎖", "D:\ほか\営業")
        completeUiWork indexTree

        (getRoot "営業").IsChecked | Should -Be $true
        (getRoot "営業部").IsChecked | Should -Be $false
        (getRoot "閉鎖").IsChecked | Should -Be $false
        (getRoot "営業").IsExpanded | Should -Be $true
        (getRoot "営業部").IsExpanded | Should -Be $false
        (getRoot "閉鎖").IsExpanded | Should -Be $true
    }
}

Describe "イベント" -Tag Io {
    BeforeAll {
        $indexDir = "$TestDrive\index"
        $settingsFile = "$TestDrive\setting.config"
        $statusFile = "$TestDrive\ingest_status.tsv"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir; StatusFile = $statusFile }
    }

    BeforeEach {
        Remove-Item -LiteralPath $indexDir, $settingsFile -Recurse -Force -ErrorAction SilentlyContinue
        $script:indexRoots.Clear()
        newIndexFiles $indexDir
        loadIndexTreeNow
        $fake.TargetUpdates = 0
        $ui.IndexTree.SelectedItem = $null
    }

    It "スペースで選んでいるフォルダのチェックを切り替えて保存する" {
        $ui.IndexTree.SelectedItem = getRoot "総務部"
        $e = [pscustomobject]@{ Key = "Space"; Handled = $false }

        & $handlers["IndexTree.PreviewKeyDown"] $ui.IndexTree $e

        $e.Handled | Should -Be $true
        (getRoot "総務部").IsChecked | Should -Be $false
        @(readSearchExcludes $settingsFile).Count | Should -Be 1
        $fake.TargetUpdates | Should -Be 1
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

        $other.Handled | Should -Be $false
        $none.Handled | Should -Be $false
        $placeholder.Handled | Should -Be $false
        isAllIndexChecked | Should -Be $true
        $fake.TargetUpdates | Should -Be 0
    }

    It "フォルダを展開したら子を読み込む" {
        $sales = getRoot "営業部"
        $sales.Children.Count | Should -Be 1
        $sales.Children[0].IsPlaceholder | Should -Be $true
        $e = New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.TreeViewItem]::ExpandedEvent)
        $e.Source = [pscustomobject]@{ DataContext = $sales }

        $handlers["IndexTree.Expanded"].Invoke($ui.IndexTree, $e)

        (@($sales.Children | ForEach-Object { $_.Name }) -join ",") | Should -Be "（このフォルダ直下のファイル）,2024"
    }

    It "チェックボックス以外のクリックではチェックを変えない" {
        $e = New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
        $e.Source = [pscustomobject]@{ DataContext = (getRoot "総務部") }

        $handlers["IndexTree.Click"].Invoke($ui.IndexTree, $e)

        isAllIndexChecked | Should -Be $true
        $fake.TargetUpdates | Should -Be 0
    }
}

Describe "ネットワークのワークスペース" -Tag Unit {
    BeforeAll {
        $netDir = "\\fileserver\共有\work\index"
        $settingsFile = "$TestDrive\setting.config"
        $workspace = newTestWorkspace @{ IndexDir = $netDir; StatusFile = "\\fileserver\共有\work\ingest_status.tsv" } "\\fileserver\共有\work"

        function newNetData {
            # 裏で集めた結果の形（getIndexTreeData）
            return @{
                State = ${pathStateFound}; Message = ""; Root = $netDir; Sources = @{ "営業部" = "C:\共有\営業部" }
                Indexes = @(
                    @{ Name = "営業部"; SourcePath = "C:\共有\営業部"; Exists = $true; HasSubfolders = $true }
                    @{ Name = "総務部"; SourcePath = ""; Exists = $true; HasSubfolders = $false }
                )
                Children = @{
                    "$netDir\営業部" = @{ HasFiles = $true; Folders = @(@{ Name = "2024"; HasSubfolders = $false }); Error = "" }
                }
            }
        }
    }

    BeforeEach {
        $script:indexRoots.Clear()
        $script:indexTreeLoading = $false
        $script:indexTreeConnectError = ""
        $script:sourceFolderMaps = @{}
        $fake.TargetUpdates = 0
        $fake.Statuses.Clear()
        $fake.Jobs.Clear()
        $fake.ErrorLogs.Clear()
    }

    It "画面のスレッドでは読まず、裏の列（network）に出して、読み込み中にする" {
        loadIndexTree

        $fake.Jobs.Count | Should -Be 1
        $fake.Jobs[0].Queue | Should -Be "network"
        $script:indexTreeLoading | Should -Be $true
        $script:indexRoots.Count | Should -Be 0
        $fake.TargetUpdates | Should -Be 1
    }

    It "裏の結果が届いたら、ツリーを組み立てて読み込み中を終える" {
        loadIndexTree

        & $fake.Jobs[0].OnDone @(newNetData) ""
        completeUiWork indexTree

        $script:indexTreeLoading | Should -Be $false
        (@($script:indexRoots | ForEach-Object { $_.Name }) -join ",") | Should -Be "営業部,総務部"
        # 展開の材料が一緒に届いていれば、展開する前から子が入っている
        (@((getRoot "営業部").Children | ForEach-Object { $_.Name }) -join ",") | Should -Be "（このフォルダ直下のファイル）,2024"
        $script:sourceFolderMaps[$netDir]["営業部"] | Should -Be "C:\共有\営業部"
    }

    It "やり直した読み込みがあれば、前の依頼の結果は捨てる" {
        loadIndexTree
        loadIndexTree

        & $fake.Jobs[0].OnDone @(newNetData) ""

        $script:indexRoots.Count | Should -Be 0
        $script:indexTreeLoading | Should -Be $true
        & $fake.Jobs[1].OnDone @(newNetData) ""
        completeUiWork indexTree
        $script:indexRoots.Count | Should -Be 2
    }

    It "<name>のときは、ツリーを空にして、接続できないことを知らせる" -TestCases @(
        @{ name = "届かない"; state = "Unreachable" }
        @{ name = "その他の失敗"; state = "Other" }
    ) {
        param ($name, $state)
        loadIndexTree

        & $fake.Jobs[0].OnDone @(@{ State = (Get-Variable "pathState$state").Value; Message = "x"; Root = $netDir; Sources = @{}; Indexes = @(); Children = @{} }) ""

        $script:indexTreeLoading | Should -Be $false
        $script:indexTreeConnectError | Should -Be $workspace.Dir
        $script:indexRoots.Count | Should -Be 0
        $fake.Statuses[$fake.Statuses.Count - 1] | Should -Be "ワークスペースに接続できません：$($workspace.Dir)"
    }

    It "裏の仕事が失敗したときも、接続できないことを知らせる" {
        loadIndexTree

        & $fake.Jobs[0].OnDone @() "失敗しました"

        $fake.ErrorLogs | Should -Be @("検索対象のツリーの読み込み : 失敗しました")
        $script:indexTreeConnectError | Should -Be $workspace.Dir
        $script:indexTreeLoading | Should -Be $false
    }

    It "ネットワークのフォルダを展開すると、裏で読む間は読み込み中のまま、読めたら子を入れる" {
        loadIndexTree
        & $fake.Jobs[0].OnDone @((newNetData)) ""
        $fake.Jobs.Clear()
        $node = [IndexNode]::CreateRoot($netDir, "技術部", "技術部", $null, $true, $true)

        expandIndexNode $node
        expandIndexNode $node

        $fake.Jobs.Count | Should -Be 1
        $fake.Jobs[0].Queue | Should -Be "network"
        $node.IsLoading | Should -Be $true
        $node.Children[0].IsPlaceholder | Should -Be $true
        & $fake.Jobs[0].OnDone @(@{ HasFiles = $false; Folders = @(@{ Name = "設計"; HasSubfolders = $false }); Error = "" }) ""
        $node.IsLoading | Should -Be $false
        (@($node.Children | ForEach-Object { $_.Name }) -join ",") | Should -Be "設計"
    }

    It "読めなかったときは、フォルダを閉じて読み込み前に戻し、フォルダを知らせる" {
        $node = [IndexNode]::CreateRoot($netDir, "技術部", "技術部", $null, $true, $true)
        $node.SetExpanded($true)
        expandIndexNode $node

        & $fake.Jobs[0].OnDone @(@{ HasFiles = $false; Folders = @(); Error = "アクセスできません" }) ""

        $node.NeedsLoad() | Should -Be $true
        $node.IsExpanded | Should -Be $false
        $node.Children[0].IsPlaceholder | Should -Be $true
        $fake.Statuses[$fake.Statuses.Count - 1] | Should -Be "フォルダを読み込めませんでした：技術部（アクセスできません）"
        # もう一度開けば読み直す
        expandIndexNode $node
        $fake.Jobs.Count | Should -Be 2
    }

    It "展開の途中でツリーを読み込み直したら、前の結果は入れない" {
        $node = [IndexNode]::CreateRoot($netDir, "技術部", "技術部", $null, $true, $true)
        expandIndexNode $node
        loadIndexTree

        & $fake.Jobs[0].OnDone @(@{ HasFiles = $false; Folders = @(@{ Name = "設計"; HasSubfolders = $false }); Error = "" }) ""

        $node.Children[0].IsPlaceholder | Should -Be $true
    }
}

Describe "ローカルのワークスペース" -Tag Unit {
    BeforeAll {
        $localDir = "C:\Users\test\tebunko_ws\work\index"
        $settingsFile = "$TestDrive\setting.config"
        $workspace = newTestWorkspace @{ IndexDir = $localDir; StatusFile = "C:\Users\test\tebunko_ws\work\ingest_status.tsv" } "C:\Users\test\tebunko_ws\work"
    }

    BeforeEach {
        $script:indexRoots.Clear()
        $script:indexTreeLoading = $false
        $script:indexTreeConnectError = ""
        $fake.Jobs.Clear()
    }

    It "ローカルも画面のスレッドでは読まず、裏の列（default）に出して、読み込み中にする（インデックスが多いと止まって見えるため）" {
        Mock getIndexTreeData { throw "画面の流れで読んだ" }

        loadIndexTree

        Should -Invoke getIndexTreeData -Times 0
        $fake.Jobs.Count | Should -Be 1
        $fake.Jobs[0].Queue | Should -Be "default"
        $script:indexTreeLoading | Should -Be $true
        $script:indexRoots.Count | Should -Be 0
    }
}
