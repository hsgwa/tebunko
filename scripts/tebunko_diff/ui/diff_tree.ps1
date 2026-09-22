# フォルダのツリー（左右に並べる）。行の組み立ては判断層の buildTreeRows（開いているフォルダの中だけの平らな一覧）。
# WPF の TreeView は、行が多いと仮想化が効きにくく遅いため使わず、仮想化した ListBox に並べる

$script:tree = @{
    Expanded  = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Collapsed = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Rows      = @()
    Selected  = ""     # 選んでいる行の相対パス
    Dirty     = $false # 状態が変わったので、次のタイマーで並べ直す
    Syncing   = $false # 並べ直している間（選択の変更を扱わない）
}

# 選んだファイルの差分は、止まってから 150 ms たったものだけを出す（↑↓ で次々に選んでも、たまらないように）
$script:treeSelectTimer = newTimer 150 {
    $script:treeSelectTimer.Stop()
    safe { onDiffFileSelected $script:tree.Selected }
}

function resetTree {
    $script:tree.Expanded.Clear()
    $script:tree.Collapsed.Clear()
    $script:tree.Rows = @()
    $script:tree.Selected = ""
    $ui.TreeList.ItemsSource = $null
}

function refreshTree {
    # ツリーを並べ直す（開いたフォルダ・選んだ行は保つ）
    $entries = @($script:dt.Entries)
    $rows = buildTreeRows $entries $script:tree.Expanded $script:tree.Collapsed ([bool]$ui.HideSameCheck.IsChecked)
    $script:tree.Rows = $rows
    $script:tree.Syncing = $true
    try {
        $ui.TreeList.ItemsSource = $rows
        if ($script:tree.Selected) {
            for ($i = 0; $i -lt $rows.Count; $i++) {
                if ([string]::Equals($rows[$i].RelPath, $script:tree.Selected, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $ui.TreeList.SelectedIndex = $i
                    break
                }
            }
        }
    } finally {
        $script:tree.Syncing = $false
    }
    $script:tree.Dirty = $false
}

function toggleTreeFolder {
    # フォルダを開く・閉じる（open: $true = 開く、$false = 閉じる、$null = 切り替え）
    param (
        $row,
        $open = $null
    )

    if (!$row -or !$row.IsFolder) {
        return
    }
    $willOpen = if ($null -eq $open) { !$row.Expanded } else { [bool]$open }
    if ($willOpen) {
        [void]$script:tree.Collapsed.Remove($row.RelPath)
        [void]$script:tree.Expanded.Add($row.RelPath)
    } else {
        [void]$script:tree.Expanded.Remove($row.RelPath)
        [void]$script:tree.Collapsed.Add($row.RelPath)
    }
    $script:tree.Selected = $row.RelPath
    refreshTree
}

function setAllTreeFolders {
    # すべて開く・閉じる
    param (
        [bool]$open
    )

    $script:tree.Expanded.Clear()
    $script:tree.Collapsed.Clear()
    foreach ($entry in @($script:dt.Entries)) {
        foreach ($folder in (getFolderPathsToOpen $entry.RelPath)) {
            if ($open) { [void]$script:tree.Expanded.Add($folder) } else { [void]$script:tree.Collapsed.Add($folder) }
        }
    }
    refreshTree
}

function selectTreeFile {
    # 相対パスのファイルを選ぶ（閉じたフォルダの中なら開く）
    param (
        [string]$relPath
    )

    foreach ($folder in (getFolderPathsToOpen $relPath)) {
        [void]$script:tree.Collapsed.Remove($folder)
        [void]$script:tree.Expanded.Add($folder)
    }
    $script:tree.Selected = $relPath
    refreshTree
    if ($ui.TreeList.SelectedItem) {
        $ui.TreeList.ScrollIntoView($ui.TreeList.SelectedItem)
    }
    $script:treeSelectTimer.Stop()
    onDiffFileSelected $relPath
}

function moveDiffFile {
    # 次（direction = 1）・前（-1）の違うファイルへ移る。移れたら $true
    param (
        [int]$direction
    )

    if ($script:dt.Mode -ne "folder" -or @($script:dt.Entries).Count -eq 0) {
        return $false
    }
    $next = findNextDiffEntry @($script:dt.Entries) $script:tree.Selected $direction
    if (!$next) {
        return $false
    }
    selectTreeFile $next
    return $true
}

# ---- イベント ----

$ui.TreeList.Add_SelectionChanged({
    safe {
        if ($script:tree.Syncing) { return }
        $row = $ui.TreeList.SelectedItem
        if (!$row) { return }
        $script:tree.Selected = $row.RelPath
        if ($row.IsFolder) {
            $script:treeSelectTimer.Stop()
            onDiffFolderSelected $row
            return
        }
        $script:treeSelectTimer.Stop()
        $script:treeSelectTimer.Start()
    }
})

$ui.TreeList.Add_PreviewMouseLeftButtonUp({
    param ($sender, $e)
    safe {
        $item = [System.Windows.Controls.ItemsControl]::ContainerFromElement($ui.TreeList, $e.OriginalSource)
        if ($item -and $item.DataContext -and $item.DataContext.IsFolder) {
            toggleTreeFolder $item.DataContext
        }
    }
})

$ui.TreeList.Add_PreviewKeyDown({
    param ($sender, $e)
    safe {
        $row = $ui.TreeList.SelectedItem
        $modifiers = [System.Windows.Input.Keyboard]::Modifiers
        if ($e.Key -eq "Right" -and $row -and $row.IsFolder) {
            toggleTreeFolder $row $true
            $e.Handled = $true
        } elseif ($e.Key -eq "Left" -and $row -and $row.IsFolder) {
            toggleTreeFolder $row $false
            $e.Handled = $true
        } elseif ($e.Key -eq "Enter" -and $row -and $row.IsFolder) {
            toggleTreeFolder $row
            $e.Handled = $true
        } elseif ($e.Key -eq "Tab" -and $modifiers -eq "None") {
            # ツリーから差分へ
            $target = if ($ui.ChangeList.Visibility -eq "Visible") { $ui.ChangeList } else { $ui.SideList }
            [void]$target.Focus()
            $e.Handled = $true
        }
    }
})

$ui.ExpandTreeButton.Add_Click({ safe { setAllTreeFolders $true } })
$ui.CollapseTreeButton.Add_Click({ safe { setAllTreeFolders $false } })
