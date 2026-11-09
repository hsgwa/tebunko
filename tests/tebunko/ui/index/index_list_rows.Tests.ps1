# インデックス一覧の行の置き直し（tebunko\ui\index\index_list.ps1 の updateIndexingButton・updateIndexRowsProgress・updateIndexListView など）のテスト。
# 画面は偽物にする。置き直す回数を数えて、変わらないときは行に触らないことを確かめる。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_types.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_detail_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit_view.ps1"

    $script:calls = New-Object System.Collections.Generic.List[string]
    function updateIndexFooter { }
    function updateIndexDetailPanel { param ([object[]]$checkedItems) $script:calls.Add("updateIndexDetailPanel") }
    function refreshFolderStatus { }
    function refreshIndexingState { }
    function updateIndexSourceFile { }
    function saveTargets { param ([switch]$Quiet) }
    function setStatus { param ([string]$text) }
    function getIndexAddedStatus { param ($name) "" }
    function refreshLeftoverIndexes { }
    function testIndexSubfoldersPending { param ($a, $b) $false }
    function testIndexItemTextExcludesPending { param ($item) $false }
    function loadTheme { @{} }
    function getTargetFolders { @() }
    function getTargetsKey { param ($folders) "" }
    function newFolderItem {
        param ([string]$path, [bool]$enabled, [string]$name, [bool]$subfolders = $true, [bool]$ingestedSubfolders = $true)
        $item = [FolderItem]::new(); $item.Path = $path; $item.Enabled = $enabled; $item.Name = $name; $item
    }

    # 画面の型を使うファイルは読み込まず、必要な関数だけを取り出す（dot-source は、この場の scope で行う）
    foreach ($source in @(
            @{ File = "${scriptsDir}\tebunko\ui\index\index_list.ps1"; Names = @("isIndexing", "isIndexingOrPreparing", "getCurrentIndexJobBlocker", "testIndexInRound", "testIndexExporting", "getIndexRowActionsKey", "setIndexRowActions", "resolveIndexCheckedItems", "updateIndexSelectionView", "updateIndexRowsProgress", "updateIndexListView", "fillTargetItems", "applyIndexStats", "setExportingRows") }
            @{ File = "${scriptsDir}\tebunko\ui\index\index_detail.ps1"; Names = @("updateIndexingButton", "getIndexingRatio") }
            @{ File = "${scriptsDir}\tebunko\ui\index\index_store.ps1"; Names = @("addIndexItem") }
        )) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($source.File, [ref]$null, [ref]$null)
        foreach ($functionName in $source.Names) {
            $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $functionName }, $true)
            . ([scriptblock]::Create($node.Extent.Text))
        }
    }

    function newFakeUi {
        $ui = @{}
        foreach ($name in @("IndexingButton", "IndexingHint", "NewIndexButton", "IndexDetailPathButton", "IndexDetailName", "IndexDetailPath", "IndexDetailSubfolders", "ChangeIndexTextExtensionsButton", "RowMenuUpdate", "RowMenuOpenFolder", "ExportIndexButton", "RemoveIndexButton", "ActionsButton", "ActionUpdate", "ActionExport", "ActionImport", "ActionDelete", "SelectionCountText", "SelectAllCheckBox", "IndexGridPlaceholder", "IndexDetailHost", "IndexSplitter", "IndexingProgress")) {
            $ui[$name] = [pscustomobject]@{ Content = $null; IsEnabled = $true; ToolTip = $null; Text = $null; Visibility = $null; IsChecked = $null; IsIndeterminate = $false; Value = 0.0 }
        }
        $ui.IndexGrid = [pscustomobject]@{ SelectedItem = $null; ItemsSource = $null }
        $ui.IndexDetailRow = [pscustomobject]@{ Height = $null; MinHeight = 0 }
        return $ui
    }

    function newFakeRow {
        # SetRowActions・SetIndexProgress を数える偽物の行
        param ([string]$name, [string]$level = "Ok", [bool]$enabled = $true)
        $row = [pscustomobject]@{ Name = $name; IndexLevel = $level; Enabled = $enabled; Checked = $false; IndexSub = "旧"; IndexPercent = -1.0; IndexText = "旧"; FailedCount = 0; PendingCount = 0; Subfolders = $true; IngestedSubfolders = $true; RowActionCalls = 0; ProgressCalls = 0 }
        $row | Add-Member ScriptMethod SetRowActions { param ($action, $updateEnabled) $this.RowActionCalls++ }
        $row | Add-Member ScriptMethod SetIndexProgress { param ($sub, $percent) $this.ProgressCalls++; $this.IndexSub = $sub; $this.IndexPercent = $percent }
        $row | Add-Member ScriptMethod SetStats { param ($a, $b, $c) }
        $row | Add-Member ScriptMethod SetIndexState { param ($text, $toolTip, $level, $sub, $percent) }
        return $row
    }
}

Describe "行のボタンの置き直し" -Tag Unit {
    BeforeEach {
        $script:ui = newFakeUi
        $script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        $script:exportingNames = @()
        $script:indexingSession = $null
        $script:indexingState = $null
        $script:indexBusy = $false; $script:archiveBusy = $false; $script:targetsNaming = $false; $script:workspaceMoving = $false; $script:indexingPreparing = $false
        $script:indexRowsVersion = 0
        $script:indexRowActionsKey = $null
        $script:targetItems.Add((newFakeRow "営業"))
        $script:targetItems.Add((newFakeRow "技術"))
        Mock getIndexCheckedItems { @() }
    }

    It "動いている処理・エクスポート中の名前・行が前回と同じなら、updateIndexingButton は行の SetRowActions を呼ばない" {
        updateIndexingButton
        $script:targetItems[0].RowActionCalls | Should -Be 1
        $script:targetItems[1].RowActionCalls | Should -Be 1

        updateIndexingButton

        $script:targetItems[0].RowActionCalls | Should -Be 1
        $script:targetItems[1].RowActionCalls | Should -Be 1
    }

    It "<case> が変わったら、全行に呼ぶ" -TestCases @(
        @{ case = "動いている処理"; change = { $script:indexBusy = $true } }
        @{ case = "エクスポート中の名前"; change = { $script:exportingNames = @("営業") } }
    ) {
        updateIndexingButton
        & $change

        updateIndexingButton

        $script:targetItems[0].RowActionCalls | Should -Be 2
        $script:targetItems[1].RowActionCalls | Should -Be 2
    }

    It "行を足したら、動いている処理が同じでも、新しい行に SetRowActions を呼び、ボタンの値が getIndexRowActions の値になる" {
        $script:ui.IndexGrid | Add-Member ScriptMethod ScrollIntoView { param ($i) }
        $script:indexingState = $null
        Mock updateIndexListView { }
        updateIndexingButton

        addIndexItem "C:\共有\営業部" "総務"
        $newItem = $script:targetItems[2]
        $newItem.IndexLevel = "Ok"
        updateIndexingButton

        $expected = getIndexRowActions "Ok" "" $false
        $newItem.RowAction | Should -Be $expected.Action
        $newItem.RowUpdateEnabled | Should -Be $expected.UpdateEnabled
        $script:indexRowsVersion | Should -Be 1
    }

    It "fillTargetItems で行を置き換えたら（数が同じでも）、次の updateIndexingButton で新しい行に呼ぶ" {
        $script:targetItems.Clear()
        $first = newFakeRow "営業"
        $script:targetItems.Add($first)
        Mock updateIndexListView { }
        updateIndexingButton
        $first.RowActionCalls | Should -Be 1
        Mock getTargetFolders { @() }
        Mock newFolderItem { newFakeRow $name }

        fillTargetItems @([pscustomobject]@{ Name = "営業"; Path = "C:\共有\営業部"; Enabled = $true; Subfolders = $true }) @()
        $replaced = $script:targetItems[0]
        updateIndexingButton

        $replaced | Should -Not -Be $first
        $replaced.RowActionCalls | Should -Be 1
    }

    It "SetRowActions は同じ値なら通知しない" {
        $item = [FolderItem]::new()
        $names = New-Object System.Collections.Generic.List[string]
        $item.add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]{ param($sender, $e) $names.Add($e.PropertyName) }.GetNewClosure())
        $item.SetRowActions("Stop", $true)
        $item.SetRowActions("Stop", $true)
        @($names) | Should -Be @("RowAction", "RowUpdateEnabled")
    }

    It "setIndexRowActions は、blocker を渡せば求めず、省けば求める" {
        $item = newFakeRow "営業"
        Mock getCurrentIndexJobBlocker { "" }

        setIndexRowActions $item "indexing"
        Should -Invoke getCurrentIndexJobBlocker -Times 0

        setIndexRowActions $item
        Should -Invoke getCurrentIndexJobBlocker -Times 1
    }
}

Describe "updateIndexRowsProgress" -Tag Unit {
    BeforeEach {
        $script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        $script:indexingOnlyNames = @()
        $script:indexRowsPercent = $null
        $script:run1 = newFakeRow "営業" "Run"
        $script:run2 = newFakeRow "技術" "Run"
        $script:idle = newFakeRow "総務" "Ok"
        foreach ($row in @($script:run1, $script:run2, $script:idle)) { $script:targetItems.Add($row) }
        $script:ratio = 0.45
        Mock getIndexingRatio { $script:ratio }
        Mock getIndexRowView { @{ Sub = "$([int]($ratio * 100))%"; Percent = [int]($ratio * 100) } }
        Mock updateIndexDetailPanel { }
    }

    It "割合が同じ 2 回目は、行の文言を求めず、行に置かない" {
        updateIndexRowsProgress
        updateIndexRowsProgress

        Should -Invoke getIndexRowView -Times 1 -Exactly
        $script:run1.ProgressCalls | Should -Be 1
    }

    It "割合が変わったら、文言は 4 回以下で求め、更新中の行の補足と棒だけを置く（更新中でない行は変えない）" {
        updateIndexRowsProgress
        $script:ratio = 0.5

        updateIndexRowsProgress

        Should -Invoke getIndexRowView -Times 2 -Exactly
        $script:run1.IndexSub | Should -Be "50%"
        $script:run1.IndexPercent | Should -Be 50
        $script:run2.IndexSub | Should -Be "50%"
        $script:idle.IndexSub | Should -Be "旧"
        $script:idle.IndexPercent | Should -Be -1
        $script:idle.IndexText | Should -Be "旧"
    }

    It "Enabled・今の回の組が違う行は、組ごとに 1 回ずつ文言を求める" {
        $script:run2.Enabled = $false
        updateIndexRowsProgress

        Should -Invoke getIndexRowView -Times 2 -Exactly
    }

    It "詳細欄は置かない" {
        updateIndexRowsProgress

        Should -Invoke updateIndexDetailPanel -Times 0
    }

    It "控えを消したあと（回の始め・終わり・集計の反映）は、割合が同じでも行に置く" {
        updateIndexRowsProgress
        $script:indexRowsPercent = $null

        updateIndexRowsProgress

        $script:run1.ProgressCalls | Should -Be 2
    }

    It "applyIndexStats は控えた割合を消し、blocker を 1 回だけ求めて全行に渡し、控えのキーを置く" {
        $script:ui = newFakeUi
        $script:indexRowsPercent = 45
        $script:exportingNames = @()
        $script:indexBusy = $false; $script:archiveBusy = $false; $script:targetsNaming = $false; $script:workspaceMoving = $false
        $script:indexRowsVersion = 0
        $script:indexRowActionsKey = $null
        $script:indexFileTotal = 0
        $script:textExtState = @{ Excludes = @{}; Ingested = @{}; Counts = @{} }
        Mock getCurrentIndexJobBlocker { "" }
        Mock getIndexRowView { @{ Text = "最新"; ToolTip = ""; Level = "Ok"; Sub = ""; Percent = 0 } }
        Mock updateIndexFooter { }
        foreach ($row in $script:targetItems) { $row | Add-Member ScriptMethod SetStats { param ($a, $b, $c) } -Force }

        applyIndexStats $null $true $null

        Should -Invoke getCurrentIndexJobBlocker -Times 1 -Exactly
        $script:indexRowsPercent | Should -BeNullOrEmpty
        $script:indexRowActionsKey | Should -Be (getIndexRowActionsKey "")
        $script:run1.RowActionCalls | Should -Be 1
    }
}

Describe "updateIndexListView" -Tag Unit {
    It "1 回の操作で getIndexCheckedItems は 1 回だけ呼ばれる" {
        $script:ui = newFakeUi
        $script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        $script:targetItems.Add((newFakeRow "営業"))
        $script:indexListEmpty = $false
        $script:exportingNames = @()
        $script:indexingSession = $null
        $script:indexBusy = $false; $script:archiveBusy = $false; $script:targetsNaming = $false; $script:workspaceMoving = $false; $script:indexingPreparing = $false
        $script:indexRowsVersion = 0
        $script:indexRowActionsKey = $null
        $script:indexingState = $null
        Mock getIndexCheckedItems { @() }
        Mock updateIndexDetailPanel { }

        updateIndexListView

        Should -Invoke getIndexCheckedItems -Times 1 -Exactly
        Should -Invoke updateIndexDetailPanel -Times 1 -Exactly
    }
}
