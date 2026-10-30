# インデックスごとの「取り込む拡張子」を選ぶダイアログ（画面層）。screens.ps1 が読み込む。
# 文言・行・一覧の見せ方・設定に書く値の作り方は判断層（text_ext_view.ps1）。ここは、それらをつなぎ、ダイアログ（dialog_index_text_extensions.xaml）を開くだけ。
# 設定への保存はここではしない（呼び元が、戻り値を writeIndexTextExcludes で書く）。入口は詳細の「取り込む拡張子」の［変更…］と、確認ダイアログの［変更…］。
# 1 行はチェックと拡張子・件数（右揃え）・割合の棒（最大の拡張子を 240 とした長さ（棒の列に収まる長さ））。外した行は拡張子・件数を薄くし、棒は Border.Divider、チェックの行の棒は Accent.Soft。
# 件数が 0 の拡張子は、一覧の末尾の 1 行にまとめて折りたたむ（押すと開いて選べる）。件数が無いとき（控え無し）は、件数の列・棒・0 件のまとめを出さない。

function addIndexTextExtensionRow {
    # 一覧の 1 行（変えた印の点・チェックと拡張子・件数・棒）を panel に足し、行の部品を返す:
    # @{ Extension; RowCount（件数。無ければ $null）; Zero（件数が 0 か）; Line; Dot; Check; CountText; Bar }
    param (
        $dialog,
        $panel,
        $row,   # getIndexTextExtensionRows の 1 行
        [bool]$hasCounts,
        [int]$maxCount
    )

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Height = 30
    $grid.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
    foreach ($width in @(20, 242, 100)) {
        $column = New-Object System.Windows.Controls.ColumnDefinition
        $column.Width = New-Object System.Windows.GridLength($width)
        [void]$grid.ColumnDefinitions.Add($column)
    }
    [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))

    $dot = New-Object System.Windows.Shapes.Ellipse
    $dot.Width = 6
    $dot.Height = 6
    $dot.Fill = $dialog.FindResource("Accent")
    $dot.Opacity = 0
    $dot.HorizontalAlignment = "Left"
    $dot.VerticalAlignment = "Center"
    $check = New-Object System.Windows.Controls.CheckBox
    $check.Content = $row.Extension
    $check.IsChecked = $row.Checked
    $check.VerticalAlignment = "Center"
    $check.SetValue([System.Windows.Automation.AutomationProperties]::NameProperty, $row.Extension)
    [System.Windows.Controls.Grid]::SetColumn($check, 1)
    $children = @($dot, $check)

    $countText = $null
    $bar = $null
    if ($hasCounts) {
        $countText = New-Object System.Windows.Controls.TextBlock
        $countText.Text = formatTextExtensionCount $row.Count
        $countText.TextAlignment = "Right"
        $countText.VerticalAlignment = "Center"
        $countText.Margin = New-Object System.Windows.Thickness(0, 0, 12, 0)
        [System.Windows.Controls.Grid]::SetColumn($countText, 2)
        $bar = New-Object System.Windows.Controls.Border
        $bar.Height = 10
        $bar.CornerRadius = New-Object System.Windows.CornerRadius(3)
        $bar.HorizontalAlignment = "Left"
        $bar.VerticalAlignment = "Center"
        $bar.Width = $(if ($maxCount -gt 0) { [Math]::Round(240.0 * [int]$row.Count / $maxCount) } else { 0 })
        [System.Windows.Controls.Grid]::SetColumn($bar, 3)
        $children += @($countText, $bar)
    }
    foreach ($child in $children) {
        [void]$grid.Children.Add($child)
    }
    $line = New-Object System.Windows.Controls.Border
    $line.BorderBrush = $dialog.FindResource("Border.Row")
    $line.BorderThickness = New-Object System.Windows.Thickness(0, 0, 0, 1)
    $line.Child = $grid
    [void]$panel.Children.Add($line)
    return @{ Extension = $row.Extension; Zero = ($hasCounts -and [int]$row.Count -eq 0); Line = $line; Dot = $dot; Check = $check; CountText = $countText; Bar = $bar }
}

function newIndexTextExtensionZeroHeader {
    # 件数 0 の拡張子をまとめた行（押すと開閉する）。文は setIndexTextExtensionZeroHeader が入れる
    param (
        $dialog
    )

    $text = New-Object System.Windows.Controls.TextBlock
    $text.Style = $dialog.FindResource("Meta")
    $text.VerticalAlignment = "Center"
    $border = New-Object System.Windows.Controls.Border
    $border.Background = $dialog.FindResource("Bg.Subtle")
    $border.Padding = New-Object System.Windows.Thickness(10, 8, 0, 8)
    $border.Cursor = [System.Windows.Input.Cursors]::Hand
    $border.Child = $text
    return @{ Border = $border; Text = $text }
}

function updateIndexTextExtensionDialogView {
    # 選びが変わったとき・絞り込みが変わったとき・0 件の行を開閉したときに、一覧の見せ方（行の表示・変えた印・外した行の色・要約・一致なしの文）を直す。
    # ctx: @{ Rows; Controls; ZeroHeader; HasCounts; State（@{ ZeroExpanded }） }
    param (
        $ctx
    )

    $current = @($ctx.Rows | ForEach-Object { @{ Extension = $_.Extension; Count = $_.RowCount; Checked = [bool]$_.Check.IsChecked } })
    $filterText = [string]$ctx.Controls.TextExtFilterBox.Text
    $changedOnly = [bool]$ctx.Controls.TextExtChangedOnly.IsChecked
    $view = getIndexTextExtensionListView $current $filterText $changedOnly
    $filtering = ($filterText.Trim() -ne "" -or $changedOnly)
    $muted = $ctx.Controls.Dialog.FindResource("Ink.Muted")
    $normal = $ctx.Controls.Dialog.FindResource("Ink.Body")
    $accentSoft = $ctx.Controls.Dialog.FindResource("Accent.Soft")
    $divider = $ctx.Controls.Dialog.FindResource("Border.Divider")
    $anyVisible = $false
    foreach ($rowControl in $ctx.Rows) {
        $checked = [bool]$rowControl.Check.IsChecked
        $shown = ($view.Visible -contains $rowControl.Extension) -and (-not $rowControl.Zero -or $ctx.State.ZeroExpanded -or $filtering)
        $rowControl.Line.Visibility = $(if ($shown) { "Visible" } else { "Collapsed" })
        $anyVisible = $anyVisible -or $shown
        $rowControl.Dot.Opacity = $(if ($checked) { 0 } else { 1 })
        $rowControl.Check.Foreground = $(if ($checked) { $normal } else { $muted })
        if ($ctx.HasCounts) {
            $rowControl.CountText.Foreground = $(if ($checked) { $normal } else { $muted })
            $rowControl.Bar.Background = $(if ($checked) { $accentSoft } else { $divider })
            $rowControl.Bar.Visibility = $(if ($rowControl.Zero) { "Collapsed" } else { "Visible" })
        }
    }
    if ($null -ne $ctx.ZeroHeader) {
        $ctx.ZeroHeader.Text.Text = getIndexTextExtensionZeroHeader $view.ZeroCount $ctx.State.ZeroExpanded
        $ctx.ZeroHeader.Border.Visibility = $(if ($filtering) { "Collapsed" } else { "Visible" })
    }
    $ctx.Controls.TextExtSummary.Text = $view.Summary
    $ctx.Controls.TextExtEmpty.Visibility = $(if ($view.Empty -or -not $anyVisible) { "Visible" } else { "Collapsed" })
    $ctx.Controls.TextExtFilterPlaceholder.Visibility = $(if ($filterText -eq "") { "Visible" } else { "Collapsed" })
}

function newIndexTextExtensionDialogControls {
    # ダイアログを読み込み、文言を入れて、部品の表（名前 → 部品）を返す
    param (
        $owner,
        $texts
    )

    $dialog = loadWindow "${xamlDir}\dialog_index_text_extensions.xaml"
    $dialog.Owner = $owner
    $dialog.Title = $texts.Title
    $ctrl = @{ Dialog = $dialog }
    foreach ($name in @("TextExtTitle", "NoteText", "ExtensionHeader", "CountHeader", "RowsPanel", "TextExtList", "TextExtEmpty", "TextExtSummary", "TextExtChangedOnly",
            "TextExtFilterBox", "TextExtFilterPlaceholder", "TextExtResetButton", "TextExtUncheckAllButton", "TextExtCancelButton", "TextExtOkButton")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $ctrl.TextExtTitle.Text = $texts.Title
    $ctrl.NoteText.Text = $texts.Note
    $ctrl.ExtensionHeader.Text = $texts.Extension
    $ctrl.CountHeader.Text = $texts.Count
    $ctrl.TextExtEmpty.Text = $texts.Empty
    $ctrl.TextExtFilterPlaceholder.Text = $texts.FilterPlaceholder
    $ctrl.TextExtChangedOnly.Content = $texts.RemovedOnly
    $ctrl.TextExtResetButton.Content = $texts.Reset
    $ctrl.TextExtUncheckAllButton.Content = $texts.UncheckAll
    $ctrl.TextExtCancelButton.Content = $texts.Cancel
    $ctrl.TextExtOkButton.Content = $texts.Ok
    return $ctrl
}

function showIndexTextExtensionsDialog {
    # 「取り込む拡張子 - 〈インデックス名〉」を開く。戻り値: ［OK］なら新しい取り込まない拡張子の配列（変わらなくても配列。空配列も有効）、［キャンセル］・閉じたときは $null。
    #   Excluded    今の取り込まない拡張子
    #   Counts      拡張子 → 件数のハッシュテーブル（件数なしは $null）
    #   CountsSource 件数の出どころ（previous = 詳細から・前回の件数、now = 確認ダイアログから・今数えた件数、none = 件数なし）
    #   CountsTime  前回の控えの日時の文字列（previous のとき注に入れる）
    param (
        $Owner,
        [string]$IndexName,
        [object[]]$Excluded = @(),
        $Counts = $null,
        [string]$CountsSource = "none",
        [string]$CountsTime = ""
    )

    $hasCounts = ($CountsSource -ne "none" -and $null -ne $Counts)
    $texts = getIndexTextExtensionDialogTexts $IndexName $(if ($hasCounts) { $CountsSource } else { "none" }) $CountsTime
    $rows = @(getIndexTextExtensionRows $(if ($hasCounts) { $Counts } else { $null }) $Excluded)
    $ctrl = newIndexTextExtensionDialogControls $Owner $texts
    $dialog = $ctrl.Dialog
    if (-not $hasCounts) {
        $ctrl.CountHeader.Visibility = "Collapsed"
    }
    $maxCount = $(if ($hasCounts) { [int](@($rows | ForEach-Object { [int]$_.Count } | Measure-Object -Maximum).Maximum) } else { 0 })

    # 行を足す（件数 0 の拡張子は末尾に並び、その前にまとめの行を置く）
    $rowControls = New-Object System.Collections.Generic.List[object]
    $zeroHeader = $null
    foreach ($row in $rows) {
        if ($hasCounts -and $null -eq $zeroHeader -and [int]$row.Count -eq 0) {
            $zeroHeader = newIndexTextExtensionZeroHeader $dialog
            [void]$ctrl.RowsPanel.Children.Add($zeroHeader.Border)
        }
        $control = addIndexTextExtensionRow $dialog $ctrl.RowsPanel $row $hasCounts $maxCount
        $control.RowCount = $row.Count
        $rowControls.Add($control)
    }
    $ctx = @{ Rows = $rowControls.ToArray(); Controls = $ctrl; ZeroHeader = $zeroHeader; HasCounts = $hasCounts; State = @{ ZeroExpanded = $false } }

    # 選び・絞り込み・開閉が変わったら、一覧の見せ方を直す（ハンドラは関数を名前で呼ばず、実体を変数に取って & で呼ぶ）
    $refresh = ${function:updateIndexTextExtensionDialogView}
    $onChange = { & $refresh $ctx }.GetNewClosure()
    foreach ($rowControl in $ctx.Rows) {
        $rowControl.Check.Add_Checked($onChange)
        $rowControl.Check.Add_Unchecked($onChange)
    }
    $ctrl.TextExtFilterBox.Add_TextChanged($onChange)
    $ctrl.TextExtChangedOnly.Add_Checked($onChange)
    $ctrl.TextExtChangedOnly.Add_Unchecked($onChange)
    if ($null -ne $zeroHeader) {
        $zeroHeader.Border.Add_MouseLeftButtonUp({
            $ctx.State.ZeroExpanded = -not $ctx.State.ZeroExpanded
            & $refresh $ctx
        }.GetNewClosure())
    }

    # ［既定に戻す］は全部チェック、［すべて外す］は全部外す（どちらも保存は［OK］）
    $ctrl.TextExtResetButton.Add_Click({
        foreach ($rowControl in $ctx.Rows) { $rowControl.Check.IsChecked = $true }
    }.GetNewClosure())
    $ctrl.TextExtUncheckAllButton.Add_Click({
        foreach ($rowControl in $ctx.Rows) { $rowControl.Check.IsChecked = $false }
    }.GetNewClosure())
    $ctrl.TextExtOkButton.Add_Click({ $dialog.DialogResult = $true }.GetNewClosure())
    & $refresh $ctx

    if ($true -ne (showOwnedDialog $dialog)) {
        return $null
    }
    $chosen = @($ctx.Rows | ForEach-Object { @{ Extension = $_.Extension; Checked = [bool]$_.Check.IsChecked } })
    return , @(convertIndexTextExtensionRows $chosen)
}
