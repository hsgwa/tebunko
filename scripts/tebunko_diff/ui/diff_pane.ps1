# ファイルの差分の表示（左右に並べる・一覧）。［2 比較］タブ（diff_tab.ps1）から呼ぶ。
# 行の組み立て（左右の対応・たたみ・文字の区切り）は裏の runspace で済ませてある。ここは表に渡し、選ぶ・移る・開くだけを行う

$script:pane = @{
    Diff       = $null   # 表示中の FileDiff
    LeftPath   = ""      # 元のファイル（開くときに使う）
    RightPath  = ""
    PlaceIndex = -1
    Rows       = @()     # 左右の表示に出している行（たたんだ行を開くと変わる）
    State      = $null   # DiffViewState（Excel のセルの横の位置）
    PendingJump = 0      # 表示した後に移る変更（1 = 最初の変更、-1 = 最後の変更。F8 でファイルをまたいだとき）
}
$script:pane.State = [DiffViewState]::new()
$ui.SideList.Tag = $script:pane.State

function clearFileDiff {
    # 差分を消し、代わりに message を出す
    param (
        [string]$message = ""
    )

    $script:pane.Diff = $null
    $script:pane.Rows = @()
    $script:pane.PlaceIndex = -1
    $ui.SideList.ItemsSource = $null
    $ui.ChangeList.ItemsSource = $null
    $ui.PlaceList.ItemsSource = $null
    $ui.GridHeader.Visibility = "Collapsed"
    $ui.GridHScroll.Visibility = "Collapsed"
    $ui.OverviewBar.Children.Clear()
    $ui.PlaceNoteText.Visibility = "Collapsed"
    $ui.DiffMessage.Text = $message
    $ui.DiffMessage.Visibility = if ($message) { "Visible" } else { "Collapsed" }
    $ui.DetailText.Text = ""
    $ui.OpenLeftButton.IsEnabled = $false
    $ui.OpenRightButton.IsEnabled = $false
    $ui.PrevChangeButton.IsEnabled = $false
    $ui.NextChangeButton.IsEnabled = $false
}

function showFileDiff {
    # 1 ファイルの比較の結果を出す。title は差分の上に出す名前（フォルダのときは相対パス）
    param (
        $fileDiff,
        [string]$title,
        [string]$leftPath,
        [string]$rightPath
    )

    $script:pane.Diff = $fileDiff
    $script:pane.LeftPath = $leftPath
    $script:pane.RightPath = $rightPath
    $ui.FileTitleText.Text = $title
    $ui.FileTitleText.ToolTip = "比較元: $leftPath`n比較先: $rightPath"
    $tabs = New-Object System.Collections.Generic.List[object]
    $places = @($fileDiff.Places)
    for ($i = 0; $i -lt $places.Count; $i++) {
        $tab = [PlaceTab]::new()
        $tab.Text = getPlaceTabText $places[$i]
        $tab.Status = $places[$i].Status
        $tab.Index = $i
        $tabs.Add($tab)
    }
    $ui.PlaceList.ItemsSource = $tabs
    if ($places.Count -eq 0) {
        clearFileDiff "比べる文字がありません。（どちらのファイルにも文字が見つかりませんでした）"
        $ui.FileTitleText.Text = $title
        return
    }
    $index = getDefaultPlaceIndex $places
    $ui.PlaceList.SelectedIndex = $index
    showPlace $index
}

function showPlace {
    # 場所（シート・本文など）を 1 つ出す
    param (
        [int]$index
    )

    $places = @($script:pane.Diff.Places)
    if ($index -lt 0 -or $index -ge $places.Count) {
        return
    }
    $script:pane.PlaceIndex = $index
    $place = $places[$index]
    $rows = if ($ui.FoldCheck.IsChecked -and $place.FoldedRows) { $place.FoldedRows } else { $place.Rows }
    $script:pane.Rows = @($rows)
    $script:pane.State.SetOffset(0)
    $ui.GridHScroll.Value = 0

    if ($place.IsGrid) {
        # 左右の列はそろえて並べてある。相手側にだけある列は見出しを空きにし、追加・削除した列は色を付ける
        $leftHeads = New-Object System.Collections.Generic.List[object]
        $rightHeads = New-Object System.Collections.Generic.List[object]
        $total = 0.0
        for ($c = 0; $c -lt @($place.ColumnWidths).Count; $c++) {
            $total += $place.ColumnWidths[$c]
            foreach ($side in @("Left", "Right")) {
                $head = [ColumnHead]::new()
                $head.Text = $place."${side}ColumnNames"[$c]
                $head.Width = $place.ColumnWidths[$c]
                $head.Kind = if (!$head.Text) { "empty" } else { $place.ColumnKinds[$c] }
                if ($side -eq "Left") { $leftHeads.Add($head) } else { $rightHeads.Add($head) }
            }
        }
        $ui.LeftColumnHeader.ItemsSource = $leftHeads
        $ui.RightColumnHeader.ItemsSource = $rightHeads
        $ui.GridHeader.Visibility = "Visible"
        $script:pane.GridWidth = $total
        updateGridScroll
    } else {
        $ui.GridHeader.Visibility = "Collapsed"
        $ui.GridHScroll.Visibility = "Collapsed"
    }

    # シート名を変えた・行が多すぎる等の注意と、追加・削除・移動した列の説明。
    # 行があるときは表の上の欄に、行が無いときは表の真ん中に出す
    $note = (@($place.Note, $place.ColumnNote) | Where-Object { $_ }) -join "　"
    $showNote = ($note -and $script:pane.Rows.Count -gt 0)
    $ui.PlaceNoteText.Text = if ($showNote) { $note } else { "" }
    $ui.PlaceNoteText.Visibility = if ($showNote) { "Visible" } else { "Collapsed" }
    $message = if ($showNote) { "" } else { $note }
    if (!$message -and $script:pane.Rows.Count -eq 0) {
        $message = "文字がありません。"
    }
    $ui.DiffMessage.Text = $message
    $ui.DiffMessage.Visibility = if ($message) { "Visible" } else { "Collapsed" }

    $ui.SideList.ItemsSource = $script:pane.Rows
    $ui.ChangeList.ItemsSource = @($place.Rows | Where-Object { ($_.Type -eq "Line" -or $_.Type -eq "Header") -and $_.Kind -ne "same" })
    $hasChange = @($place.Rows | Where-Object { ($_.Type -eq "Line" -or $_.Type -eq "Header") -and $_.Kind -ne "same" }).Count -gt 0
    $ui.PrevChangeButton.IsEnabled = $true
    $ui.NextChangeButton.IsEnabled = $true
    $ui.DetailText.Text = if ($hasChange -or $place.Status -ne "same") { "" } else { "この場所に違いはありません。" }
    $ui.OpenLeftButton.IsEnabled = [bool]$script:pane.LeftPath -and [bool]$place.LeftName
    $ui.OpenRightButton.IsEnabled = [bool]$script:pane.RightPath -and [bool]$place.RightName
    updateOverviewBar

    # F8 でファイルをまたいできたときは、最初（または最後）の変更へ
    if ($script:pane.PendingJump -ne 0) {
        $direction = $script:pane.PendingJump
        $script:pane.PendingJump = 0
        $start = if ($direction -gt 0) { -1 } else { $script:pane.Rows.Count }
        $next = findNextChangeRow $script:pane.Rows $start $direction
        if ($next -ge 0) { selectSideRow $next }
    } elseif ($script:pane.Rows.Count -gt 0) {
        $ui.SideList.ScrollIntoView($script:pane.Rows[0])
    }
}

function updateGridScroll {
    # Excel の表の横のスクロールの範囲（表の幅 - 見えている幅）
    if ($ui.GridHeader.Visibility -ne "Visible") {
        return
    }
    $visible = [Math]::Max(0, $ui.SideList.ActualWidth / 2 - 44 - 20)
    $max = [Math]::Max(0, $script:pane.GridWidth - $visible)
    $ui.GridHScroll.Maximum = $max
    $ui.GridHScroll.ViewportSize = $visible
    $ui.GridHScroll.LargeChange = [Math]::Max(40, $visible * 0.8)
    $ui.GridHScroll.Visibility = if ($max -gt 0) { "Visible" } else { "Collapsed" }
}

function updateOverviewBar {
    # 右の端に、場所全体での変更の位置を印で出す（変更の塊の先頭ごと。多すぎれば間引く）
    $bar = $ui.OverviewBar
    $bar.Children.Clear()
    $rows = $script:pane.Rows
    $height = $bar.ActualHeight
    if ($rows.Count -eq 0 -or $height -le 0) {
        return
    }
    $brushes = @{
        change = [System.Windows.Media.Brushes]::RoyalBlue
        insert = [System.Windows.Media.Brushes]::SeaGreen
        delete = [System.Windows.Media.Brushes]::IndianRed
    }
    $count = 0
    $last = -10.0
    for ($i = findNextChangeRow $rows -1 1; $i -ge 0; $i = findNextChangeRow $rows $i 1) {
        $top = [Math]::Floor($height * $i / $rows.Count)
        if ($top - $last -lt 3) { continue }
        $last = $top
        $mark = New-Object System.Windows.Shapes.Rectangle
        $mark.Width = 8
        $mark.Height = 4
        $mark.Fill = $brushes[$rows[$i].Kind]
        $mark.Tag = $i
        [System.Windows.Controls.Canvas]::SetLeft($mark, 3)
        [System.Windows.Controls.Canvas]::SetTop($mark, $top)
        [void]$bar.Children.Add($mark)
        $count++
        if ($count -ge 400) { break }
    }
}

function selectSideRow {
    # 左右の差分の行を選び、見える位置へ出す
    param (
        [int]$index
    )

    if ($index -lt 0 -or $index -ge $script:pane.Rows.Count) {
        return
    }
    $row = $script:pane.Rows[$index]
    $ui.SideList.SelectedIndex = $index
    $ui.SideList.ScrollIntoView($row)
    if ($ui.ChangeList.Visibility -eq "Visible") {
        $ui.ChangeList.SelectedItem = $row
        if ($ui.ChangeList.SelectedItem) { $ui.ChangeList.ScrollIntoView($row) }
    }
}

function moveChange {
    # 次（direction = 1）・前（-1）の変更へ移る。ファイルの最後まで来たら、フォルダのときは次の違うファイルへ
    param (
        [int]$direction
    )

    if ($null -eq $script:pane.Diff) {
        moveDiffFile $direction
        return
    }
    $rows = $script:pane.Rows
    $from = $ui.SideList.SelectedIndex
    if ($ui.ChangeList.Visibility -eq "Visible" -and $ui.ChangeList.SelectedItem) {
        $from = [Array]::IndexOf($rows, $ui.ChangeList.SelectedItem)
    }
    $next = findNextChangeRow $rows $from $direction
    if ($next -ge 0) {
        selectSideRow $next
        return
    }
    # この場所の最後: 次の違いのある場所へ
    $places = @($script:pane.Diff.Places)
    $p = $script:pane.PlaceIndex + $direction
    while ($p -ge 0 -and $p -lt $places.Count) {
        if ($places[$p].Status -ne "same") {
            $script:pane.PendingJump = $direction
            $ui.PlaceList.SelectedIndex = $p
            return
        }
        $p += $direction
    }
    # ファイルの最後: フォルダのときは次の違うファイルへ
    if (moveDiffFile $direction) {
        $script:pane.PendingJump = $direction
    } else {
        setStatus $(if ($direction -gt 0) { "最後の変更です。" } else { "最初の変更です。" })
    }
}

function expandFold {
    # たたんだ行を開く（その行の代わりに、たたんだ行を並べる）
    param (
        $fold
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $script:pane.Rows) {
        if ([object]::ReferenceEquals($row, $fold)) {
            foreach ($hidden in @($fold.FoldRows)) { $rows.Add($hidden) }
        } else {
            $rows.Add($row)
        }
    }
    $script:pane.Rows = $rows.ToArray()
    $ui.SideList.ItemsSource = $script:pane.Rows
    if (@($fold.FoldRows).Count -gt 0) {
        $ui.SideList.ScrollIntoView($fold.FoldRows[0])
    }
    updateOverviewBar
}

function getSelectedDiffRow {
    if ($ui.ChangeList.Visibility -eq "Visible") {
        return $ui.ChangeList.SelectedItem
    }
    return $ui.SideList.SelectedItem
}

function updateDetail {
    # 選んだ行の説明（Excel の違うセル）と、開くボタン
    $row = getSelectedDiffRow
    if ($null -eq $row) {
        return
    }
    $ui.DetailText.Text = if ($row.Detail) { $row.Detail } elseif ($row.Kind -ne "same" -and $row.Type -eq "Line") { getKindLabel $row.Kind } else { "" }
}

function openDiffSide {
    # 比較元（left）・比較先（right）のファイルを、選んだ行の位置で開く
    param (
        [string]$side,
        $row = $null
    )

    $path = if ($side -eq "left") { $script:pane.LeftPath } else { $script:pane.RightPath }
    if (!$path) {
        return
    }
    if (!(Test-Path -LiteralPath $path)) {
        showMessage "ファイルが見つかりません。`n$path" "OK" "Warning" | Out-Null
        return
    }
    $mode = getDiffOpenMode
    $place = ""
    $cell = ""
    if ($row) {
        $place = if ($side -eq "left") { $row.LeftPlace } else { $row.RightPlace }
        $cell = if ($side -eq "left") { $row.LeftCell } else { $row.RightCell }
    }
    if (!$place -and $script:pane.Diff -and $script:pane.PlaceIndex -ge 0) {
        $current = @($script:pane.Diff.Places)[$script:pane.PlaceIndex]
        $place = if ($side -eq "left") { $current.LeftName } else { $current.RightName }
    }
    if ((getDiffKind $path) -eq "Excel") {
        try {
            openInExcel $path $place $cell $mode
            setStatus "Excel で開きました：$([System.IO.Path]::GetFileName($path))$(if ($cell) { "（$place!$cell）" })"
            return
        } catch {
            # Excel がダイアログを出している等で操作できないときは、既定のアプリで開く
        }
    }
    [void](openWithShell $path $mode)
    setStatus "開きました：$([System.IO.Path]::GetFileName($path))"
}

function getDiffOpenMode {
    # 開き方は、検索の［開き方］と同じ設定（setting.config の openMode）を使う
    try {
        $data = readSettingsData
        if ($data -and ${openModes} -contains [string]$data.openMode) {
            return [string]$data.openMode
        }
    } catch { }
    return ${openModeNormal}
}

function setDiffView {
    # 左右に並べる・一覧を切り替える
    param (
        [string]$view
    )

    $side = ($view -ne "list")
    $ui.SideList.Visibility = if ($side) { "Visible" } else { "Collapsed" }
    $ui.ChangeList.Visibility = if ($side) { "Collapsed" } else { "Visible" }
    $ui.GridHeader.Visibility = if ($side -and $script:pane.Diff -and $script:pane.PlaceIndex -ge 0 -and @($script:pane.Diff.Places)[$script:pane.PlaceIndex].IsGrid) { "Visible" } else { "Collapsed" }
    if (!$side) { $ui.GridHScroll.Visibility = "Collapsed" } else { updateGridScroll }
}

# ---- イベント ----

$ui.PlaceList.Add_SelectionChanged({
    safe {
        $tab = $ui.PlaceList.SelectedItem
        if ($tab -and $tab.Index -ne $script:pane.PlaceIndex) {
            showPlace $tab.Index
        }
    }
})

$ui.FoldCheck.Add_Click({
    safe {
        updateDiffSettingsSafe ([ordered]@{ diffFoldSame = [bool]$ui.FoldCheck.IsChecked })
        if ($script:pane.PlaceIndex -ge 0) { showPlace $script:pane.PlaceIndex }
    }
})

$ui.ViewSideButton.Add_Checked({ safe { setDiffView "side"; updateDiffSettingsSafe ([ordered]@{ diffView = "side" }) } })
$ui.ViewListButton.Add_Checked({ safe { setDiffView "list"; updateDiffSettingsSafe ([ordered]@{ diffView = "list" }) } })
$ui.NextChangeButton.Add_Click({ safe { moveChange 1 } })
$ui.PrevChangeButton.Add_Click({ safe { moveChange -1 } })

$ui.GridHScroll.Add_ValueChanged({ $script:pane.State.SetOffset(-$ui.GridHScroll.Value) })
$ui.SideList.Add_SizeChanged({ safe { updateGridScroll; updateOverviewBar } })
$ui.OverviewBar.Add_SizeChanged({ safe { updateOverviewBar } })
$ui.OverviewBar.Add_MouseLeftButtonDown({
    param ($sender, $e)
    safe {
        $rows = $script:pane.Rows
        if ($rows.Count -eq 0) { return }
        $y = $e.GetPosition($ui.OverviewBar).Y
        $index = [int][Math]::Floor($rows.Count * $y / [Math]::Max(1, $ui.OverviewBar.ActualHeight))
        $next = findNextChangeRow $rows ([Math]::Max(-1, $index - 1)) 1
        if ($next -lt 0) { $next = findNextChangeRow $rows $rows.Count -1 }
        if ($next -ge 0) { selectSideRow $next }
    }
})

$ui.SideList.Add_SelectionChanged({ safe { updateDetail } })
$ui.ChangeList.Add_SelectionChanged({ safe { updateDetail } })

# たたんだ行はクリックで開く
$ui.SideList.Add_PreviewMouseLeftButtonUp({
    param ($sender, $e)
    safe {
        $item = [System.Windows.Controls.ItemsControl]::ContainerFromElement($ui.SideList, $e.OriginalSource)
        if ($item -and $item.DataContext -and $item.DataContext.Type -eq "Fold") {
            expandFold $item.DataContext
        }
    }
})

# ダブルクリック・Enter: 押した側（左半分 = 比較元・右半分 = 比較先）のファイルを、その位置で開く
$ui.SideList.Add_MouseDoubleClick({
    param ($sender, $e)
    safe {
        $item = [System.Windows.Controls.ItemsControl]::ContainerFromElement($ui.SideList, $e.OriginalSource)
        if (!$item -or !$item.DataContext -or $item.DataContext.Type -eq "Fold") { return }
        $x = $e.GetPosition($ui.SideList).X
        $side = if ($x -lt $ui.SideList.ActualWidth / 2) { "left" } else { "right" }
        $row = $item.DataContext
        if ($side -eq "left" -and $row.LeftEmpty) { $side = "right" }
        if ($side -eq "right" -and $row.RightEmpty) { $side = "left" }
        openDiffSide $side $row
    }
})
$ui.ChangeList.Add_MouseDoubleClick({
    param ($sender, $e)
    safe {
        $row = $ui.ChangeList.SelectedItem
        if ($row) { openDiffSide $(if ($row.Kind -eq "delete") { "left" } else { "right" }) $row }
    }
})
$ui.OpenLeftButton.Add_Click({ safe { openDiffSide "left" (getSelectedDiffRow) } })
$ui.OpenRightButton.Add_Click({ safe { openDiffSide "right" (getSelectedDiffRow) } })
