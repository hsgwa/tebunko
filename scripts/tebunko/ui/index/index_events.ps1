# インデックス管理の画面のイベントの登録（ボタン・一覧の操作・行のメニュー）。

$ui.NewIndexButton.Add_Click({ safe { newIndex } })
$ui.IndexEmptyAddButton.Add_Click({ safe { newIndex } })
$ui.EditIndexButton.Add_Click({ safe { editIndex } })
$ui.RemoveIndexButton.Add_Click({ safe { deleteIndex } })
$ui.ExportIndexButton.Add_Click({ safe { newExportIndex } })
$ui.ImportIndexButton.Add_Click({ safe { newImportIndex } })
# 行の［⋯］を押したら、その行を選んで行のメニュー（編集・エクスポート・削除）を［⋯］の下に開く
# （［⋯］は行ごとの部品なので、一覧の Click で受ける。メニューの項目の可否は updateIndexingButton が決める）
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
    param ($sender, $e)
    safe {
        $button = $e.OriginalSource
        if ($button -isnot [System.Windows.Controls.Button] -or $button.Tag -ne "IndexRowMenu") {
            return
        }
        $ui.IndexGrid.SelectedItem = $button.DataContext
        $ui.IndexRowMenu.PlacementTarget = $button
        $ui.IndexRowMenu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
        $ui.IndexRowMenu.IsOpen = $true
        $e.Handled = $true
    }
})
function getIndexGridRowAt {
    # 一覧の中の、クリックした場所の行（DataGridRow）。行の外（列見出し・余白・スクロールバー）なら $null
    param ($source)

    $element = $source
    while ($element -and $element -isnot [System.Windows.Controls.DataGridRow]) {
        if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
            return $null
        }
        $element = if ($element -is [System.Windows.Media.Visual]) { [System.Windows.Media.VisualTreeHelper]::GetParent($element) } else { $null }
    }
    return $element
}
# 行を右クリックしたら、その行を選んでからメニューを開く（選んでいる別の行にメニューが効かないようにする）。
# 行の外（列見出し・余白）では開かない
$ui.IndexGrid.Add_PreviewMouseRightButtonDown({
    param ($sender, $e)
    safe {
        $row = getIndexGridRowAt $e.OriginalSource
        if ($row) {
            $ui.IndexGrid.SelectedItem = $row.Item
        }
    }
})
$ui.IndexGrid.Add_ContextMenuOpening({
    param ($sender, $e)
    safe {
        if (!(getIndexGridRowAt $e.OriginalSource)) {
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
