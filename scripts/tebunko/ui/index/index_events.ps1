# インデックス管理の画面のイベントの登録（ボタン・一覧の操作・行のメニュー）。

# 頭の ⓘ と、ステータス列の見出しの説明（文言は index_view.ps1 が決める）
$ui.IndexScreenInfo.ToolTip = (getIndexScreenInfoText) -join "`n"
[System.Windows.Automation.AutomationProperties]::SetHelpText($ui.IndexScreenInfo, $ui.IndexScreenInfo.ToolTip)
$ui.StatusColumnHeader.ToolTip = (getIndexStatusHelpText) -join "`n"
$ui.NewIndexButton.Add_Click({ safe { newIndex } })
$ui.IndexEmptyAddButton.Add_Click({ safe { newIndex } })
$ui.EditIndexButton.Add_Click({ safe { editIndex } })
# 行の右クリックの［更新］は、行の右端の［更新］と同じ処理（押した行だけを更新する）
$ui.RowMenuUpdate.Add_Click({
    safe {
        $item = getIndexTargetItem
        if ($item) { updateSelectedIndexes @($item.Name) }
    }
})
$ui.RowMenuOpenFolder.Add_Click({ safe { openIndexSourceFolder } })
$ui.IndexDetailPathButton.Add_Click({ safe { changeIndexFolder } })
$ui.RemoveIndexButton.Add_Click({ safe { deleteIndex } })
$ui.ExportIndexButton.Add_Click({ safe { newExportIndex } })
# ［アクション ▾］: 押したら、ボタンの下にメニューを開く（項目の可否は updateIndexingButton が決める）
$ui.ActionsButton.Add_Click({
    safe {
        $ui.ActionsMenu.PlacementTarget = $ui.ActionsButton
        $ui.ActionsMenu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
        $ui.ActionsMenu.IsOpen = $true
    }
})
$ui.ActionImport.Add_Click({ safe { newImportIndex } })
# ［エクスポート…］［削除…］は、チェックを付けた行に対して動く。1 件なら既存の 1 件ずつの処理（一覧で選んでいる行に効くため、
# その行を選んでから呼ぶ）、2 件以上ならまとめての処理（書き出し先・削除の確認は 1 回）
$ui.ActionExport.Add_Click({
    safe {
        $checked = @(getIndexCheckedItems @($script:targetItems))
        if ($checked.Count -eq 1) {
            $ui.IndexGrid.SelectedItem = $checked[0]
            newExportIndex
        } elseif ($checked.Count -ge 2) {
            newBulkExportIndexes @($checked | ForEach-Object { $_.Name })
        }
    }
})
# ［編集…］も同じ（1 件にだけ効く）。チェックが 1 件ならその行を選んでから、チェックが無ければ押した行に対して動く。2 件以上は使えない
$ui.ActionEdit.Add_Click({
    safe {
        $checked = @(getIndexCheckedItems @($script:targetItems))
        if ($checked.Count -eq 1) {
            $ui.IndexGrid.SelectedItem = $checked[0]
            editIndex
        } elseif ($checked.Count -eq 0) {
            editIndex
        }
    }
})
$ui.ActionDelete.Add_Click({
    safe {
        $checked = @(getIndexCheckedItems @($script:targetItems))
        if ($checked.Count -eq 1) {
            $ui.IndexGrid.SelectedItem = $checked[0]
            deleteIndex
        } elseif ($checked.Count -ge 2) {
            deleteIndexes @($checked | ForEach-Object { $_.Name })
        }
    }
})
# 一覧の名前・パスは、「…」で切れている行だけ全文をツールチップで見せる（行ごとに 1 回だけ、行の Loaded で印の付いた TextBlock に付ける）
$ui.IndexGrid.Add_LoadingRow({
    param ($sender, $e)
    addTrimmedToolTipToRows $e.Row "TrimTip"
})
$ui.ActionUpdate.Add_Click({ safe { updateSelectedIndexes @(getIndexCheckedItems @($script:targetItems) | ForEach-Object { $_.Name }) } })
# 行の［更新］［中止］は行ごとの部品なので、一覧の Click で受ける
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
    param ($sender, $e)
    safe {
        $button = $e.OriginalSource
        if ($button -isnot [System.Windows.Controls.Button]) {
            return
        }
        if ($button.Tag -eq "RowUpdate" -and $button.DataContext -is [FolderItem]) {
            updateSelectedIndexes @($button.DataContext.Name)
            $e.Handled = $true
        } elseif ($button.Tag -eq "RowStop") {
            stopIndexing
            $e.Handled = $true
        }
    }
})
# 見出しの全選択: 1 件でも付いていないものがあれば全部付け、全部付いていれば全部外す（一部だけのときは付ける側）。
# チェックはその場だけの選びなので、保存しない
$ui.SelectAllCheckBox.Add_Click({
    safe {
        $items = @($script:targetItems)
        $checkedCount = @(getIndexCheckedItems $items).Count
        $value = ($checkedCount -lt $items.Count)
        foreach ($item in $items) {
            $item.SetRowChecked($value)
        }
        updateIndexListView
    }
})
# 行を右クリックしたら、その行を選んでからメニューを開く（選んでいる別の行にメニューが効かないようにする）。
# 行の外（列見出し・余白）では開かない
$ui.IndexGrid.Add_PreviewMouseRightButtonDown({
    param ($sender, $e)
    safe {
        $row = getDataGridRowAt $e.OriginalSource
        if ($row) {
            $ui.IndexGrid.SelectedItem = $row.Item
        }
    }
})
$ui.IndexGrid.Add_ContextMenuOpening({
    param ($sender, $e)
    safe {
        if (!(getDataGridRowAt $e.OriginalSource)) {
            $e.Handled = $true
        }
    }
})
$ui.IndexGrid.Add_SelectionChanged({ safe { updateIndexListView } })
$ui.IndexGrid.Add_MouseDoubleClick({ safe { editIndex } })
$ui.IndexGrid.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Delete") {
        safe { deleteIndex }
    }
})
$ui.IndexGrid.Add_PreviewDragOver({ onFolderDragOver @args })
$ui.IndexGrid.Add_PreviewDrop({
    param ($sender, $e)
    safe {
        foreach ($folder in (getDroppedFolders $e)) {
            addIndexForFolder $folder
        }
    }
    $e.Handled = $true
})
