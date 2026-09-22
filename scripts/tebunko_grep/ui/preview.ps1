# ［2 検索］タブの、選択した行の前後を表で見せるプレビュー。

# 選択行のプレビューは、↑↓で続けて選択が変わったときは最後の1回だけ読む（巨大なTSVでも操作が重くならないようにする）。
# プレビューの高さを変えたときも、入る行数に合わせて読み直すためにこのタイマーを使う
$script:detailTimer = newTimer 120 {
    $script:detailTimer.Stop()
    safe { showDetail }
}
# 高さを変えただけのときは、横スクロールの位置をそのままにする（ドラッグのたびに左へ戻らないように）
$script:detailKeepScroll = $false

# 表示中の行（getViewRows）・選んでいる行（getSelectedRows）・選んでいる行 1 つ（getCurrentHitRow）は result_list.ps1

function clearDetail {
    # 行を選んでいないときのプレビュー。枠（と高さ）はそのままにし、中身を空にして案内を出す
    $ui.PreviewHeader.ItemsSource = $null
    $ui.PreviewRows.ItemsSource = $null
    $script:previewTable = $null
    $ui.DetailTitle.Text = ""
    $ui.DetailTitle.ToolTip = $null
    $ui.PreviewNote.Visibility = "Collapsed"
    $ui.PreviewHeaderScroll.Visibility = "Collapsed"
    $ui.PreviewPlaceholder.Visibility = "Visible"
    $ui.OpenButton.IsEnabled = $false
    $ui.OpenFolderButton.IsEnabled = $false
}

function getPreviewContextLines {
    # プレビューの高さを測り、選択行の前後に読む行数を決める（決め方は preview_view.ps1）
    $height = $ui.PreviewScroll.ViewportHeight
    if ($height -le 0) {
        $height = $ui.PreviewScroll.ActualHeight - ${previewScrollBarSize}
    }
    return (getPreviewRowCounts $height ${previewRowHeight} ${maxPreviewRows})
}

function showDetail {
    $row = getCurrentHitRow
    if ($null -eq $row) {
        clearDetail
        return
    }
    $ui.PreviewPlaceholder.Visibility = "Collapsed"
    $ui.PreviewHeaderScroll.Visibility = "Visible"
    $ui.OpenButton.IsEnabled = $true
    $ui.OpenFolderButton.IsEnabled = $true
    $path = if ($row.RelDir) { "$($row.RelDir)\$($row.Book)" } else { $row.Book }
    $place = if ($row.MatchCell) { "セル $($row.MatchCell)" } else { "$($row.LineNumber) 行目" }
    $ui.OpenButton.Content = if ($row.IsExcel) { "Excel で開く" } else { "開く" }

    # 前後の行をインデックスのTSVから読む（読めなければ選択行だけを出す）。行数はプレビューの高さに合わせる
    $lines = getPreviewContextLines
    $context = @(readTsvContext ([System.IO.Path]::Combine($row.Root, $row.RelPath)) $row.LineNumber $lines[0] $lines[1])
    $table = $row.BuildPreview([int[]]@($context | ForEach-Object { $_.LineNumber }), [string[]]@($context | ForEach-Object { $_.Line }))

    $title = "${path} ・ $($row.Location) ・ ${place}"
    $ui.DetailTitle.Text = $title
    $ui.DetailTitle.ToolTip = $title

    # 横に長い行は一部の列だけを表示するため、その範囲を知らせる
    if ($table.TotalColumns -gt $table.ShownColumns) {
        $note = "表示は $($table.RangeLabel) の $($table.ShownColumns.ToString('N0')) 列（全 $($table.TotalColumns.ToString('N0')) 列）"
        $ui.PreviewNote.Text = $note
        $ui.PreviewNote.ToolTip = "$note　一致したセルを中心に表示しています。ほかの列は元のファイルで確認してください。"
        $ui.PreviewNote.Visibility = "Visible"
    } else {
        $ui.PreviewNote.Visibility = "Collapsed"
    }
    $ui.PreviewHeader.ItemsSource = $table.Columns
    $ui.PreviewRows.ItemsSource = $table.Rows
    $script:previewTable = $table

    # 一致したセルが見えるよう横にスクロールする（左端から見えていればそのまま）。
    # 高さを変えただけのときは、見ていた横の位置をそのままにする
    $ui.PreviewScroll.UpdateLayout()
    if ($script:detailKeepScroll) {
        $script:detailKeepScroll = $false
        return
    }
    $offset = 0
    if ($table.HitOffset + $table.HitWidth -gt $ui.PreviewScroll.ViewportWidth) {
        $offset = [math]::Max(0, $table.HitOffset - 120)
    }
    $ui.PreviewScroll.ScrollToHorizontalOffset($offset)
}

function getPreviewCell {
    # マウスの下（またはイベントの発生元）のプレビューのセル。セルの上でなければ $null
    param (
        $source
    )

    $element = $source
    while ($element) {
        if ($element -is [System.Windows.FrameworkElement] -and $element.DataContext -is [PreviewCell]) {
            return $element.DataContext
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    return $null
}

function copyPreviewSelection {
    # プレビューで選んだセルの値をクリップボードに入れる（1 セルならその値のまま、複数ならタブ区切り）
    if ($null -eq $script:previewTable -or !$script:previewTable.HasSelection()) {
        setStatus "プレビューでコピーするセルをクリックしてください（Shift＋クリック・ドラッグで複数選べます）。"
        return
    }
    $text = $script:previewTable.GetSelectionText()
    if ($text -eq "") {
        [System.Windows.Clipboard]::Clear()
    } else {
        [System.Windows.Clipboard]::SetText($text)
    }
    $count = $script:previewTable.SelectedCount()
    if ($count -le 1) {
        setStatus "セルの値をコピーしました：$(toStatusText $text)"
    } else {
        setStatus "${count} 個のセルをコピーしました（Excel に貼り付けると、元の位置に並びます）"
    }
}

$ui.ResultGrid.Add_SelectionChanged({
    # 別の行を選んだときは、一致したセルが見える位置まで横にスクロールし直す
    $script:detailKeepScroll = $false
    $script:detailTimer.Stop()
    $script:detailTimer.Start()
})

# 列見出しは行とは別のスクロールに置いている（縦に隠れないようにするため）ので、横位置を行に合わせる
$ui.PreviewScroll.Add_ScrollChanged({
    $ui.PreviewHeaderScroll.ScrollToHorizontalOffset($ui.PreviewScroll.HorizontalOffset)
})
# プレビューの高さを変えたら（GridSplitter のドラッグ）、入る行数に合わせて前後の行を読み直す。
# ドラッグ中は何度も起きるので、ほかと同じタイマーでまとめて 1 回だけ読む
$ui.PreviewScroll.Add_SizeChanged({
    param ($sender, $e)
    if ($e.HeightChanged -and $ui.ResultGrid.SelectedItem) {
        $script:detailKeepScroll = $true
        $script:detailTimer.Stop()
        $script:detailTimer.Start()
    }
})
# プレビューのセルをクリックすると、その値をコピーできるように選ぶ（Shift＋クリック・ドラッグで範囲、Ctrl+C でコピー）
$script:previewTable = $null
$ui.PreviewRows.Add_PreviewMouseLeftButtonDown({
    param ($sender, $e)
    safe {
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable) {
            $extend = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift
            $script:previewTable.Select($cell, [bool]$extend)
            [void]$ui.PreviewScroll.Focus()
        }
    }
})
$ui.PreviewRows.Add_MouseMove({
    param ($sender, $e)
    if ($e.LeftButton -ne "Pressed") {
        return
    }
    safe {
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable) {
            $script:previewTable.Select($cell, $true)
        }
    }
})
$ui.PreviewRows.Add_PreviewMouseRightButtonDown({
    param ($sender, $e)
    safe {
        # 右クリックしたセルが選ばれていなければ、そのセルだけを選ぶ
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable -and !$cell.IsSelected) {
            $script:previewTable.Select($cell, $false)
        }
    }
})
$ui.PreviewScroll.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "C" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") {
        safe { copyPreviewSelection }
        $e.Handled = $true
    }
})
$ui.MenuPreviewCopy.Add_Click({ safe { copyPreviewSelection } })
$ui.MenuPreviewCopyRow.Add_Click({
    safe {
        if ($script:previewTable -and $script:previewTable.HasSelection()) {
            # 選んでいるセルのある行をすべて選んでからコピーする
            $selected = $null
            foreach ($row in $script:previewTable.Rows) {
                $selected = @($row.Cells | Where-Object { $_.IsSelected })[0]
                if ($selected) {
                    break
                }
            }
            if ($selected) {
                $script:previewTable.SelectRow($selected)
            }
        }
        copyPreviewSelection
    }
})
# プレビューの列見出しの右端をドラッグすると、その列（PreviewColumn）の幅が変わる（各行のセルも同じ列を参照しているため一緒に変わる）
$ui.PreviewHeader.AddHandler(
    [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent,
    [System.Windows.Controls.Primitives.DragDeltaEventHandler] {
        param ($sender, $e)
        safe {
            $column = $e.OriginalSource.DataContext
            if ($column -is [PreviewColumn]) {
                $column.SetWidth($column.Width + $e.HorizontalChange)
            }
        }
    })
