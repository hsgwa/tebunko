# ［検索］タブの、選択した行の前後を表で見せるプレビュー。

# 選択行のプレビューは、↑↓で続けて選択が変わったときは最後の1回だけ読む（巨大なTSVでも操作が重くならないようにする）。
$script:detailTimer = newTimer 120 {
    $script:detailTimer.Stop()
    safe { showDetail }
}
# 読み込みを頼んだプレビューの番号と行（別の行を選んだ後に届いた古い結果を捨てるため）
$script:previewRequest = 0
$script:previewRow = $null
# 最後に読んだ 1 つのファイルの控え（readContentIndexLines の Marks）。次に読むときに渡し、同じファイルなら先頭から読み直さない。
# ファイルのパス・更新日時・大きさが変わったら readContentIndexLines が使わずに作り直す（初めて開くファイルの最初の読みは、ファイル全体を読む）
$script:previewMarks = $null
# 列見出しをドラッグしている間 $true（見えている列の範囲を決め直さない。updatePreviewWindow）
$script:previewDragging = $false

# 表示中の行（getViewRows）・出力する素のヒット（getExportHits）・選んでいる行（getSelectedRows）・選んでいる行 1 つ（getCurrentHitRow）は result_list.ps1

function clearDetail {
    # 行を選んでいないときのプレビュー。枠（と高さ）はそのままにし、中身を空にして案内を出す
    $script:previewRequest++   # 読んでいる途中の読み足しの結果は捨てる
    $script:previewDragging = $false
    $script:previewMarks = $null
    $ui.PreviewHeader.ItemsSource = $null
    $ui.PreviewHeaderPanel.DataContext = $null
    $ui.PreviewRows.ItemsSource = $null
    $script:previewTable = $null
    $ui.DetailTitle.Text = ""
    $ui.DetailTitle.ToolTip = $null
    $ui.PreviewHeaderScroll.Visibility = "Collapsed"
    $ui.PreviewPlaceholder.Visibility = "Visible"
    $ui.OpenButton.IsEnabled = $false
    $ui.OpenMenuButton.IsEnabled = $false
    $ui.OpenFolderButton.IsEnabled = $false
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
    $ui.OpenMenuButton.IsEnabled = $true
    $ui.OpenFolderButton.IsEnabled = $true
    # 閉じている見出しを選んだときの先頭の行は、まだ画面に出ていない（LoadingRow で Prepare されていない）ため、
    # セル番地（MatchCell）・「場所」の列の表記（PlaceDisplay）が空のままになる。ここで作っておく（作り済みなら何もしない）
    prepareHitRow $row

    # 前後の行を本文インデックスのファイルから読む。最初は選択行の前後に ${previewFirstLines} 行ずつ（高さを変えても読み直さない。
    # 表は行を仮想化して出すので、高さに入る分より多く読んでも重くならず、縦にスクロールして見られる。端に近づいたら読み足す。updatePreviewReadMore）。
    # 検索で読んだ内容があれば使う。読み込みは画面のスレッドでは行わない（キャッシュに無いと本文インデックスのファイル全体を読むため）。
    # 読み終わったら applyDetail で表にする。読んでいる間に別の行を選んだら、古い結果は捨てる（番号で見分ける）
    $script:previewRequest++
    $script:previewRow = $row
    $first = [Math]::Max(1, $row.LineNumber - ${previewFirstLines})
    $count = $row.LineNumber - $first + 1 + ${previewFirstLines}
    startJob {
        param ($id, $path, $book, $location, $first, $count, $cache, $marks)
        $result = readContentIndexLinesAcrossParts $path $book $location $first $count $cache $marks
        @{ Id = $id; Context = @($result.Rows); End = $result.End; Marks = $result.Marks }
    } @($script:previewRequest, [System.IO.Path]::Combine($row.Root, $row.RelPath), $row.Book, $row.Location, $first, $count, $script:tsvCache, $script:previewMarks) {
        param ($output, $errorText)
        if ($output -and $output.Count -gt 0 -and $output[0].Id -eq $script:previewRequest) {
            $script:previewMarks = $output[0].Marks
            applyDetail $script:previewRow $output[0].Context $output[0].Marks $output[0].End
        } elseif ($errorText -and $null -ne $script:previewRow) {
            # 読めなければ選択行だけを出す（読み足しもしない）
            applyDetail $script:previewRow @([pscustomobject]@{ LineNumber = $script:previewRow.LineNumber; Line = $script:previewRow.Line }) $null $true -noMore
        }
    }
}

function applyDetail {
    # 読んだ前後の行（readContentIndexLines の結果）でプレビューの表を作る。marks は読み足しに使う控え、
    # end は下にもう行が無いか（読んだ数が頼んだ数より少なかった）。noMore なら、読み足しをしない
    param (
        $row,
        [object[]]$context,
        $marks = $null,
        [bool]$end = $false,
        [switch]$noMore
    )

    if (!$row.Equals((getCurrentHitRow))) {
        return
    }
    $path = if ($row.RelDir) { "$($row.RelDir)\$($row.Book)" } else { $row.Book }
    $table = $row.BuildPreview([int[]]@($context | ForEach-Object { $_.LineNumber }), [string[]]@($context | ForEach-Object { $_.Line }))

    $title = formatDisplayText "${path} ・ $($row.PlaceDisplay) ・ $($row.Kind)"
    $ui.DetailTitle.Text = $title
    $ui.DetailTitle.ToolTip = $title

    $table.Marks = $marks
    $table.Request = $script:previewRequest
    $table.BottomDone = $end -or $noMore
    if ($noMore) { $table.TopDone = $true }
    $script:previewTable = $table
    # 見出しをドラッグしている途中で表が替わると、ドラッグ中の部品が木から外れて DragCompleted が届かない。旗を戻して、新しい表の範囲を決める
    $script:previewDragging = $false
    $script:previewBaseWidth = if ($table.Columns.Count -eq 1) { $table.Columns[0].Width } else { 0 }
    $script:previewFitWidth = $script:previewBaseWidth
    # 列は全部の列を持ち、横にスクロールしたとき見えている列の範囲のセルだけを作る（updatePreviewWindow）
    $ui.PreviewHeader.ItemsSource = $null
    $ui.PreviewHeaderPanel.DataContext = $table.Window
    $ui.PreviewRows.ItemsSource = $table.Rows

    # 一致したセルが見えるよう横にスクロールする（左端から見えていればそのまま）。
    # 列の幅は、行を入れて縦のバーが出るかどうかが決まってから（見える幅が定まってから）合わせる
    $ui.PreviewScroll.UpdateLayout()
    fitPreviewWidth
    $offset = 0
    if ($table.HitOffset + $table.HitWidth -gt $ui.PreviewScroll.ViewportWidth) {
        $offset = [math]::Max(0, $table.HitOffset - 120)
    }
    $ui.PreviewScroll.ScrollToHorizontalOffset($offset)
    # 選択行が上から 3 行目に来るよう縦にスクロールする（縦の単位は行。位置は getPreviewTopRow）
    $hitIndex = -1
    for ($i = 0; $i -lt $table.Rows.Count; $i++) {
        if ($table.Rows[$i].IsHitRow) {
            $hitIndex = $i
            break
        }
    }
    $ui.PreviewScroll.ScrollToVerticalOffset((getPreviewTopRow $hitIndex))
    $ui.PreviewScroll.UpdateLayout()
    updatePreviewWindow $offset
    updatePreviewReadMore
}

function updatePreviewWindow {
    # 横の位置・見える幅・列の幅から、見えている列の範囲を決め直す（判断は getPreviewColumnWindow）。
    # 範囲が変わったら、列見出しを入れ替え、見えている行のセルを今の範囲で作り直す。
    # offset は、スクロールした直後でまだ配置が終わっていないときに、行き先の位置を渡す
    param (
        $offset = $null
    )

    $table = $script:previewTable
    if ($null -eq $table -or $table.Columns.Count -eq 0) { return }
    # 列見出しをドラッグしている間は、範囲を決め直さない（見出しを入れ替えると、ドラッグしている見出しの部品が作り直されて途切れる）。
    # 縦にスクロールして見えた行のセルだけ作り、ドラッグが終わったら決め直す
    if ($script:previewDragging) {
        $table.UpdateRowWidth()   # 行の幅は、広げた列の分だけ今のうちに合わせる（右側が切れないように）
        updatePreviewCells
        return
    }
    if ($null -eq $offset) { $offset = $ui.PreviewScroll.HorizontalOffset }
    $window = getPreviewColumnWindow $table.Prefix() $offset $ui.PreviewScroll.ViewportWidth ([HitRow]::NumberWidth) ${previewColumnMargin} $table.Window.First $table.Window.Last ${previewColumnSpread}
    if ($table.SetWindow($window.First, $window.Last, $window.Left, $window.Right)) {
        $ui.PreviewHeader.ItemsSource = $table.WindowColumns()
    }
    updatePreviewCells
}

function updatePreviewCells {
    # 見えている行（と上下の余白の行）に、今の列の範囲のセルが無ければ作る。縦にスクロールして新しく見えた行・読み足した行のため
    $table = $script:previewTable
    if ($null -eq $table -or $table.DataCount() -le 0) { return }
    $visible = getPreviewVisibleRows
    $lead = $table.LeadCount()
    $table.EnsureCells($lead + $visible.First - ${previewCellRowMargin}, $lead + $visible.Last + ${previewCellRowMargin})
}

function getPreviewVisibleRows {
    # 今見えている本物の行（メッセージの行を除く）の範囲。@{ First; Last }（表の本物の行の中の位置。0 から）
    $table = $script:previewTable
    $scroll = $ui.PreviewScroll
    $lead = $table.LeadCount()
    $last = $table.DataCount() - 1
    $first = [Math]::Max(0, [int][Math]::Floor($scroll.VerticalOffset) - $lead)
    $bottom = [Math]::Min($last, [int][Math]::Ceiling($scroll.VerticalOffset + $scroll.ViewportHeight) - 1 - $lead)
    return @{ First = $first; Last = [Math]::Max($first, $bottom) }
}

function updatePreviewReadMore {
    # スクロールの位置が読んである範囲の端に近づいていたら、その向きに読み足す（判断は getPreviewReadMore）
    $table = $script:previewTable
    if ($null -eq $table -or $null -eq $script:previewRow -or $table.DataCount() -le 0) { return }
    $visible = getPreviewVisibleRows
    # メッセージの行が出ている間は、その向きを読んでいる途中（読めなかったときは、その向きは done）
    $requests = @(getPreviewReadMore $visible.First $visible.Last $table.DataCount() $table.FirstLine() $table.LastLine() `
            $table.TopDone $table.BottomDone ($null -ne $table.UpMessage -and !$table.TopDone) ($null -ne $table.DownMessage -and !$table.BottomDone) `
            ${previewReadAheadLines} ${previewChunkLines})
    foreach ($request in $requests) {
        startReadMore $request
    }
}

function startReadMore {
    # 上または下に行を読み足す。読んでいる間は、その端に「読み込み中…」の行を出す
    param (
        $request
    )

    # 読む行・番号は、読み足す表のもの（行を切り替えている間も、古い表に別の行の番号を混ぜない）
    $table = $script:previewTable
    $row = $table.Owner
    $table.ShowMessage($request.Side, "読み込み中…")
    startJob {
        param ($id, $side, $path, $book, $location, $first, $count, $cache, $marks)
        try {
            $result = readContentIndexLinesAcrossParts $path $book $location $first $count $cache $marks
            @{ Id = $id; Side = $side; Failed = $false; First = $first; Rows = @($result.Rows); End = $result.End; Marks = $result.Marks }
        } catch {
            @{ Id = $id; Side = $side; Failed = $true }
        }
    } @($table.Request, $request.Side, [System.IO.Path]::Combine($row.Root, $row.RelPath), $row.Book, $row.Location, $request.FirstLine, $request.Count, $script:tsvCache, $table.Marks) {
        param ($output, $errorText)
        applyReadMore $output $errorText
    }
}

function stopReadMore {
    # 読み足しを読めなかったとき、その向きのメッセージを「読めませんでした」にして、その向きの読み足しをやめる
    param (
        $table,
        [string]$side
    )

    if ($side -eq "up") { $table.TopDone = $true } else { $table.BottomDone = $true }
    $table.ShowMessage($side, "読めませんでした")
}

function applyReadMore {
    # 読み足した行を表に足す（上に足したときは、見ていた行が同じ所に見えたままになるよう縦の位置を直す）。
    # 持つ行が上限を超えたら、見ている所から遠い側を捨てる（判断は getPreviewTrim）
    param (
        $output,
        $errorText
    )

    $table = $script:previewTable
    if ($null -eq $table) { return }
    $result = $(if ($output -and $output.Count -gt 0) { $output[0] } else { $null })
    if ($null -eq $result) {
        # 読み足しの仕事そのものが失敗した（どの向きか分からないので、読んでいる向きすべて）
        foreach ($side in @("up", "down")) {
            if ($null -ne $(if ($side -eq "up") { $table.UpMessage } else { $table.DownMessage })) { stopReadMore $table $side }
        }
        return
    }
    # 別の行を選んで表が替わったあとに届いた古い結果は捨てる（表が持つ番号で見分ける）
    if ($result.Id -ne $table.Request) { return }
    $side = $result.Side
    if ($result.Failed) {
        stopReadMore $table $side
        return
    }
    $scroll = $ui.PreviewScroll
    $oldOffset = [int][Math]::Floor($scroll.VerticalOffset)
    $lead = $table.LeadCount()
    $numbers = [int[]]@($result.Rows | ForEach-Object { $_.LineNumber })
    $lines = [string[]]@($result.Rows | ForEach-Object { $_.Line })
    $rows = $table.Owner.BuildPreviewRows($table, $numbers, $lines)
    $table.Marks = $result.Marks
    $script:previewMarks = $result.Marks
    # 今ある行とつながらない結果（読んでいる間に反対側を捨てたなど）は足さず、また近づいたときに読み直す
    $joins = $rows.Count -gt 0 -and $(if ($side -eq "up") { $rows[$rows.Count - 1].LineNumber -eq $table.FirstLine() - 1 } else { $rows[0].LineNumber -eq $table.LastLine() + 1 })
    $table.HideMessage($side)
    if ($side -eq "up") {
        if ($joins) {
            $table.InsertRows("up", $rows)
            $scroll.UpdateLayout()
            # 見ていた先頭が「読み込み中…」の行のとき（oldOffset < lead）も、その下の行が同じ所に見えるようにする
            $scroll.ScrollToVerticalOffset([Math]::Max(0, $oldOffset - $lead + $rows.Count))
        } else {
            $scroll.ScrollToVerticalOffset([Math]::Max(0, $oldOffset - $lead))
        }
        $scroll.UpdateLayout()
        # つながらなかった結果は、場所の終わりまで読んだ（End）なら、その向きを終わりにする（終わりにしないと、同じ範囲を頼み直し続ける）。
        # 終わりまで読んでいなければ終わりとは言えない（空で返ったときだけ、その向きにもう行が無い）。
        # 下側と違って頼んだ先頭を確かめないのは、上の読みは今ある先頭の直前までを頼むので、End・空になるのは元のファイルが変わった・読めないときだけだから
        if ($rows.Count -eq 0 -or $result.End -or ($joins -and $table.FirstLine() -le 1)) { $table.TopDone = $true }
    } else {
        # 終わりと見るのは、頼んだ先頭の行が（足す前の）下の端の次のままの結果だけ。頼んだあとに下の端が動いた（下を捨てた）古い結果の End・空は、
        # 今の端の下のことではないので信じない（信じると、捨てて戻した終わりの印を立て直し、その先が読めなくなる）
        $fresh = $result.First -eq $table.LastLine() + 1
        if ($joins) { $table.InsertRows("down", $rows) }
        if ($fresh -and ($rows.Count -eq 0 -or $result.End)) { $table.BottomDone = $true }
    }
    # 持つ行が上限を超えたら、遠い側を捨てる
    $visible = getPreviewVisibleRows
    $trim = getPreviewTrim $visible.First $visible.Last $table.DataCount() ${previewMaxRows} ${previewChunkLines}
    if ($trim.Side -ne "none") {
        $offset = [int][Math]::Floor($scroll.VerticalOffset)
        $table.TrimRows($trim.Side, $trim.Count)
        if ($trim.Side -eq "top") {
            $table.TopDone = $false
            $scroll.UpdateLayout()
            $scroll.ScrollToVerticalOffset([Math]::Max(0, $offset - $trim.Count))
            $scroll.UpdateLayout()
        } else {
            $table.BottomDone = $false
        }
    }
    # 読み足した行に列が増えていれば、列の範囲を決め直す（右の空きの幅が変わる）。見えている行のセルも作る
    updatePreviewWindow
    updatePreviewReadMore
}

function fitPreviewWidth {
    # Excel 以外は列が 1 つなので、その幅を枠の幅いっぱいにする（段落は幅いっぱいに折り返して読めるようにする）。
    # 見える幅が変わるたびに呼ぶ（縦のバーの出入り・窓や欄の幅の変更）。幅の決め方は getPreviewFillWidth
    if ($null -eq $script:previewTable -or $script:previewTable.Columns.Count -ne 1 -or $script:previewTable.IsExcel) {
        return
    }
    $column = $script:previewTable.Columns[0]
    $width = getPreviewFillWidth $ui.PreviewScroll.ViewportWidth ([HitRow]::NumberWidth) $script:previewBaseWidth $column.Width $script:previewFitWidth ([PreviewColumn]::MinWidth)
    # 手で変えた幅のときは、今の幅がそのまま返る。そのときは「前に合わせた幅」を変えず、手で変えたことを覚えておく
    if ($width -ne $column.Width) {
        $column.SetWidth($width)
        $script:previewFitWidth = $column.Width
        updatePreviewWindow
    }
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
        setStatus "プレビューでコピーするセルをクリックしてください。"
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
    $script:detailTimer.Stop()
    $script:detailTimer.Start()
})

# 列見出しは行とは別のスクロールに置いている（縦に隠れないようにするため）ので、横位置を行に合わせる
$ui.PreviewScroll.Add_ScrollChanged({
    param ($sender, $e)
    # 横に動いた・見える幅が変わったときは、見えている列の範囲を決め直す（セルは見えている列の分だけ作る）
    if ($e.HorizontalChange -ne 0 -or $e.ViewportWidthChange -ne 0 -or $e.ExtentWidthChange -ne 0) {
        safe { updatePreviewWindow }
    }
    $ui.PreviewHeaderScroll.ScrollToHorizontalOffset($ui.PreviewScroll.HorizontalOffset)
    if ($e.ViewportWidthChange -ne 0) {
        fitPreviewWidth
    }
    # 縦に動いた・行が増減した・見える高さが変わったときは、新しく見えた行のセルを作り、端に近づいていれば読み足す
    if ($e.VerticalChange -ne 0 -or $e.ExtentHeightChange -ne 0 -or $e.ViewportHeightChange -ne 0) {
        safe { updatePreviewCells; updatePreviewReadMore }
    }
})
# 結果の表とプレビューの境目。結果が少ないときは表を内容の高さまで縮めてプレビューを広げる（判断は getPreviewRowHeight）。
# 利用者が境目をドラッグしている間は、自動で高さを書き戻さない（書き戻すと境目がマウスに付いて来ない）。
# ドラッグして実際に動かしたら、その高さを優先して、それ以後は自動で変えない（クリックしただけでは止めない）
$script:previewHeightDragged = $false
$ui.PreviewSplitter.Add_DragCompleted({
    param ($sender, $e)
    if (-not $e.Canceled -and $e.VerticalChange -ne 0) {
        $script:previewHeightDragged = $true
    }
})

function updateResultPaneHeight {
    # 結果の表の中身が変わった・窓の大きさが変わったあとに、プレビューの高さを決め直す。
    #   itemCount: 表の行の数（多いときは内容を測らない）
    param (
        [int]$itemCount = 0
    )

    if ($script:previewHeightDragged -or $ui.PreviewSplitter.IsDragging) { return }
    $available = $ui.ResultListHost.ActualHeight + $ui.DetailRow.ActualHeight
    $width = $ui.ResultListHost.ActualWidth
    if ($available -le 0 -or $width -le 0) { return }   # まだ画面に出ていない
    $content = $null
    # 件数が多いと全部の行を測ることになるので、少ないときだけ測る（多いときは表がいっぱいになるので測る必要が無い）
    if ($itemCount -le 40) {
        # 行を足し引きした直後は、表の中身がまだ測り直されていない。先に配置を終えてから測る
        $ui.ResultListHost.UpdateLayout()
        $root = $ui.ResultListHost.Content
        $root.Measure([System.Windows.Size]::new($width, [double]::PositiveInfinity))
        $content = $root.DesiredSize.Height
        $root.InvalidateMeasure()
    }
    $height = getPreviewRowHeight $available $content -itemCount $itemCount
    if ([Math]::Abs($ui.DetailRow.Height.Value - $height) -ge 1) {
        $ui.DetailRow.Height = [System.Windows.GridLength]::new($height)
    }
}

# 窓の大きさが変わったら決め直す（高さが同じなら何もしないので、自分で変えた高さで繰り返さない）
$ui.ResultListHost.Add_SizeChanged({
    safe { updateResultPaneHeight $script:resultItems.Count }
})

# プレビューのセルをクリックすると、その値をコピーできるように選ぶ（Shift＋クリック・ドラッグで範囲、Ctrl+C でコピー）
$script:previewTable = $null
$script:previewBaseWidth = 0
$script:previewFitWidth = 0
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
# プレビューが無いとき（結果を選んでいない・読めなかったとき）はメニューを出さない
$ui.PreviewRows.Add_ContextMenuOpening({
    param ($sender, $e)
    safe {
        if ($null -eq $script:previewTable) {
            $e.Handled = $true
            return
        }
        $parts = @{ openHere = $ui.MenuPreviewOpen; copyCell = $ui.MenuPreviewCopy; copyRow = $ui.MenuPreviewCopyRow }
        setContextMenuItems $ui.PreviewMenu (getPreviewMenuItems) $parts
    }
})
$ui.MenuPreviewOpen.Add_Click({ safe { openSource } })
$ui.MenuPreviewCopy.Add_Click({ safe { copyPreviewSelection } })
$ui.MenuPreviewCopyRow.Add_Click({
    safe {
        if ($script:previewTable -and $script:previewTable.HasSelection()) {
            # 選んでいるセルのある行をすべて選んでからコピーする
            $line = $script:previewTable.SelectedFirstLine()
            if ($line -gt 0) {
                $script:previewTable.SelectLine($line)
            }
        }
        copyPreviewSelection
    }
})
# プレビューの列見出しの右端をドラッグすると、その列（PreviewColumn）の幅が変わる（各行のセルも同じ列を参照しているため一緒に変わる）
$ui.PreviewHeader.AddHandler(
    [System.Windows.Controls.Primitives.Thumb]::DragStartedEvent,
    [System.Windows.Controls.Primitives.DragStartedEventHandler] {
        param ($sender, $e)
        $script:previewDragging = $true
    })
$ui.PreviewHeader.AddHandler(
    [System.Windows.Controls.Primitives.Thumb]::DragCompletedEvent,
    [System.Windows.Controls.Primitives.DragCompletedEventHandler] {
        param ($sender, $e)
        safe {
            $script:previewDragging = $false
            updatePreviewWindow   # ドラッグ中に変わった幅で、見えている列の範囲と左右の空きを決め直す
        }
    })
$ui.PreviewHeader.AddHandler(
    [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent,
    [System.Windows.Controls.Primitives.DragDeltaEventHandler] {
        param ($sender, $e)
        safe {
            $column = $e.OriginalSource.DataContext
            if ($column -is [PreviewColumn]) {
                $column.SetWidth($column.Width + $e.HorizontalChange)
                updatePreviewWindow   # ドラッグ中は範囲を決め直さない（終わったときに決め直す）
            }
        }
    })
