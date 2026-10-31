# インデックスごとの「取り込み対象の拡張子」を選ぶダイアログ（画面層）。screens.ps1 が読み込む。
# 文言・行・一覧の見せ方・設定に書く値の作り方・種類の見出しの状態は判断層（text_ext_view.ps1）。ここは、それらをつなぎ、ダイアログ（dialog_index_text_extensions.xaml）を開くだけ。
# 設定への保存はここではしない（呼び元が、戻り値を writeIndexTextExcludes で書く）。入口は詳細の「取り込み対象の拡張子」の［変更…］と、確認ダイアログの［変更…］。
# 一覧は Excel・Word・PowerPoint・PDF・テキストの見出しで分ける（常に展開）。見出しの行は 3 状態のチェック（全部＝チェック・一部＝「－」・全部外し＝空）で、押すとその種類をまとめて取り込み対象にする・外す。
# 1 行はチェックと拡張子・件数（右揃え）・割合の棒（最大の拡張子を 240 とした長さ（棒の列に収まる長さ））。外した行は全体を薄く（透明度 0.55）し、棒は Border.Divider、チェックの行の棒は Accent.Soft。
# 件数が 0 のテキストの拡張子は、テキストの種類の末尾の 1 行にまとめて折りたたむ（押すと開いて選べる）。件数が無いとき（控え無し）は、件数の列・棒・0 件のまとめを出さない。

function addIndexExtensionHeadingRow {
    # 種類の見出しの行（3 状態のチェックと種類名）を panel に足し、部品を返す: @{ Heading = $true; Group; Line; Check }
    param (
        $dialog,
        $panel,
        $row   # getIndexTextExtensionRows の見出しの行
    )

    $check = New-Object System.Windows.Controls.CheckBox
    $check.Content = $row.Group
    $check.IsThreeState = $true
    $check.FontWeight = "SemiBold"
    $check.Foreground = $dialog.FindResource("Ink.Strong")
    $check.VerticalAlignment = "Center"
    $check.SetValue([System.Windows.Automation.AutomationProperties]::NameProperty, $row.Group)
    $check.SetValue([System.Windows.Automation.AutomationProperties]::AutomationIdProperty, "TextExtGroup_$($row.Group)")
    $line = New-Object System.Windows.Controls.Border
    $line.Background = $dialog.FindResource("Bg.Subtle")
    $line.BorderBrush = $dialog.FindResource("Border.Row")
    $line.BorderThickness = New-Object System.Windows.Thickness(0, 0, 0, 1)
    $line.Padding = New-Object System.Windows.Thickness(10, 6, 0, 6)
    $line.Child = $check
    [void]$panel.Children.Add($line)
    return @{ Heading = $true; Group = $row.Group; Line = $line; Check = $check }
}

function addIndexTextExtensionRow {
    # 一覧の 1 行（変えた印の点・チェックと拡張子・件数・棒）を panel に足し、行の部品を返す:
    # @{ Heading = $false; Group; Extension; RowCount（件数。無ければ $null）; Zero（折りたたむ件数 0 の行か。テキストの種類の中だけ）; NoCount（件数が 0 か）; Line; Dot; Check; CountText; Bar }
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
    $noCount = ($hasCounts -and [int]$row.Count -eq 0)
    return @{ Heading = $false; Group = $row.Group; Extension = $row.Extension; Zero = ($noCount -and $row.Group -eq "テキスト"); NoCount = $noCount; Line = $line; Dot = $dot; Check = $check; CountText = $countText; Bar = $bar }
}

function newIndexTextExtensionZeroHeader {
    # 件数 0 のテキストの拡張子をまとめた行（押すと開閉する）。文は setIndexTextExtensionZeroHeader が入れる
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

function getIndexExtensionGroupMembers {
    # 種類（group）の拡張子の行の部品（隠れた行も含む）
    param (
        $ctx,
        [string]$group
    )

    return @($ctx.Rows | Where-Object { $_.Group -eq $group })
}

function setIndexTextExtensionChecks {
    # 行のチェックをまとめて変える（checkedFor は拡張子の行の部品を受け取り、チェックにするかを返すブロック）。
    # 1 行ずつ一覧の見せ方を直すと遅いので、途中は止めて、最後に 1 回だけ直す
    param (
        $ctx,
        [object[]]$members,
        [scriptblock]$checkedFor
    )

    $ctx.State.Suspend = $true
    try {
        foreach ($member in $members) { $member.Check.IsChecked = [bool](& $checkedFor $member) }
    } finally {
        $ctx.State.Suspend = $false
    }
    updateIndexTextExtensionDialogView $ctx
}

function toggleIndexExtensionGroup {
    # 見出しのチェックを押したときの動き。今の状態（行のチェックから数える）に応じて、その種類を全部取り込み対象にする・全部外す
    # （向きは getIndexExtensionGroupToggle、外した後の配列は setIndexExtensionGroup。ほかの種類の行は変えない）
    param (
        $ctx,
        [string]$group
    )

    $members = getIndexExtensionGroupMembers $ctx $group
    $extensions = @($members | ForEach-Object { $_.Extension })
    $excluded = @($ctx.Rows | Where-Object { -not [bool]$_.Check.IsChecked } | ForEach-Object { $_.Extension })
    $include = getIndexExtensionGroupToggle (getIndexExtensionGroupState $extensions $excluded)
    $after = @(setIndexExtensionGroup $excluded $extensions $include)
    $checkedFor = { param ($member) $after -notcontains $member.Extension }.GetNewClosure()
    setIndexTextExtensionChecks $ctx $members $checkedFor
}

function updateIndexTextExtensionDialogView {
    # 選びが変わったとき・絞り込みが変わったとき・0 件の行を開閉したときに、一覧の見せ方（行の表示・変えた印・外した行の薄さ・見出しの状態と表示・要約・一致なしの文）を直す。
    # ctx: @{ Rows（拡張子の行の部品）; Headings（見出しの行の部品）; Controls; ZeroHeader; HasCounts; State（@{ ZeroExpanded; Suspend（まとめて変える間、一覧の見せ方の更新を止める）; Busy（見出しのチェックを入れ直している間。押された扱いにしない） }） }
    param (
        $ctx
    )

    if ($ctx.State.Suspend) {
        return
    }
    $current = @($ctx.Rows | ForEach-Object { @{ Group = $_.Group; Extension = $_.Extension; Count = $_.RowCount; Checked = [bool]$_.Check.IsChecked } })
    $filterText = [string]$ctx.Controls.TextExtFilterBox.Text
    $keptOnly = [bool]$ctx.Controls.TextExtChangedOnly.IsChecked
    $view = getIndexTextExtensionListView $current $filterText $keptOnly
    $filtering = ($filterText.Trim() -ne "" -or $keptOnly)
    $accentSoft = $ctx.Controls.Dialog.FindResource("Accent.Soft")
    $divider = $ctx.Controls.Dialog.FindResource("Border.Divider")
    $anyVisible = $false
    foreach ($rowControl in $ctx.Rows) {
        $checked = [bool]$rowControl.Check.IsChecked
        $shown = ($view.Visible -contains $rowControl.Extension) -and (-not $rowControl.Zero -or $ctx.State.ZeroExpanded -or $filtering)
        $rowControl.Line.Visibility = $(if ($shown) { "Visible" } else { "Collapsed" })
        $anyVisible = $anyVisible -or $shown
        $rowControl.Dot.Opacity = $(if ($checked) { 0 } else { 1 })
        $rowControl.Line.Opacity = $(if ($checked) { 1.0 } else { 0.55 })
        if ($ctx.HasCounts) {
            $rowControl.Bar.Background = $(if ($checked) { $accentSoft } else { $divider })
            $rowControl.Bar.Visibility = $(if ($rowControl.NoCount) { "Collapsed" } else { "Visible" })
        }
    }
    $ctx.State.Busy = $true
    try {
        foreach ($heading in $ctx.Headings) {
            # 見出しの状態は、絞り込みで隠れた行も含めて数える。見せるのは、見せる行が 1 つ以上ある種類だけ
            $members = getIndexExtensionGroupMembers $ctx $heading.Group
            $state = getIndexExtensionGroupState @($members | ForEach-Object { $_.Extension }) @($members | Where-Object { -not [bool]$_.Check.IsChecked } | ForEach-Object { $_.Extension })
            $heading.Check.IsChecked = $(switch ($state) { "All" { $true } "None" { $false } default { $null } })
            $heading.Line.Visibility = $(if ($view.VisibleGroups -contains $heading.Group) { "Visible" } else { "Collapsed" })
        }
    } finally {
        $ctx.State.Busy = $false
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
    foreach ($name in @("TextExtTitle", "NoteText", "ExtensionHeader", "CountHeader", "BarHeader", "RowsPanel", "TextExtList", "TextExtEmpty", "TextExtSummary", "TextExtChangedOnly",
            "TextExtFilterBox", "TextExtFilterPlaceholder", "TextExtResetButton", "TextExtUncheckAllButton", "TextExtCancelButton", "TextExtOkButton")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $ctrl.TextExtTitle.Text = $texts.Title
    $ctrl.NoteText.Text = $texts.Note
    $ctrl.ExtensionHeader.Text = $texts.Extension
    $ctrl.CountHeader.Text = $texts.Count
    $ctrl.TextExtEmpty.Text = $texts.Empty
    $ctrl.TextExtFilterPlaceholder.Text = $texts.FilterPlaceholder
    $ctrl.TextExtChangedOnly.Content = $texts.KeptOnly
    $ctrl.TextExtResetButton.Content = $texts.Reset
    $ctrl.TextExtUncheckAllButton.Content = $texts.UncheckAll
    $ctrl.TextExtCancelButton.Content = $texts.Cancel
    $ctrl.TextExtOkButton.Content = $texts.Ok
    return $ctrl
}

function addIndexTextExtensionDialogRows {
    # 一覧の行を RowsPanel に足す（見出しの行を挟む。テキストの中の件数 0 の拡張子は末尾に並び、その前にまとめの行を置く）。
    # 戻り値: @{ Rows（拡張子の行の部品）; Headings（見出しの行の部品）; ZeroHeader（まとめの行。無ければ $null） }
    param (
        $dialog,
        $panel,
        [object[]]$rows,
        [bool]$hasCounts,
        [int]$maxCount
    )

    $rowControls = New-Object System.Collections.Generic.List[object]
    $headingControls = New-Object System.Collections.Generic.List[object]
    $zeroHeader = $null
    foreach ($row in $rows) {
        if ($row.Heading) {
            $headingControls.Add((addIndexExtensionHeadingRow $dialog $panel $row))
            continue
        }
        if ($hasCounts -and $null -eq $zeroHeader -and $row.Group -eq "テキスト" -and [int]$row.Count -eq 0) {
            $zeroHeader = newIndexTextExtensionZeroHeader $dialog
            [void]$panel.Children.Add($zeroHeader.Border)
        }
        $control = addIndexTextExtensionRow $dialog $panel $row $hasCounts $maxCount
        $control.RowCount = $row.Count
        $rowControls.Add($control)
    }
    return @{ Rows = $rowControls.ToArray(); Headings = $headingControls.ToArray(); ZeroHeader = $zeroHeader }
}

function showIndexTextExtensionsDialog {
    # 「取り込み対象の拡張子 - 〈インデックス名〉」を開く。戻り値: ［OK］なら新しい取り込まない拡張子の配列（Office・PDF・テキストを含む。変わらなくても配列。空配列も有効）、［キャンセル］・閉じたときは $null。
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
        $ctrl.BarHeader.Visibility = "Collapsed"
    }
    $maxCount = $(if ($hasCounts) { [int](@($rows | Where-Object { -not $_.Heading } | ForEach-Object { [int]$_.Count } | Measure-Object -Maximum).Maximum) } else { 0 })

    $built = addIndexTextExtensionDialogRows $dialog $ctrl.RowsPanel $rows $hasCounts $maxCount
    $zeroHeader = $built.ZeroHeader
    $ctx = @{ Rows = $built.Rows; Headings = $built.Headings; Controls = $ctrl; ZeroHeader = $zeroHeader; HasCounts = $hasCounts; State = @{ ZeroExpanded = $false; Suspend = $false; Busy = $false } }

    # 選び・絞り込み・開閉が変わったら、一覧の見せ方を直す（ハンドラは関数を名前で呼ばず、実体を変数に取って & で呼ぶ）
    $refresh = ${function:updateIndexTextExtensionDialogView}
    $toggleGroup = ${function:toggleIndexExtensionGroup}
    $setChecks = ${function:setIndexTextExtensionChecks}
    $allOn = { param ($member) $true }
    $allOff = { param ($member) $false }
    $onChange = { & $refresh $ctx }.GetNewClosure()
    foreach ($rowControl in $ctx.Rows) {
        $rowControl.Check.Add_Checked($onChange)
        $rowControl.Check.Add_Unchecked($onChange)
    }
    # 見出しのチェックは、押した回数ではなく行の今の状態から向きを決める（全部 → 外す・一部／全部外し → 全部取り込み対象）。
    # 押すと WPF が 3 状態を回すので、その変わった印は読まず、行から決め直した状態を見出しに入れ直す（見せ方を直すときの入れ直しでは動かない: Busy）。
    # 支援技術の切り替え（Click を通らない）でも動くよう、Click でなく Checked・Unchecked・Indeterminate で受ける
    foreach ($heading in $ctx.Headings) {
        $groupName = $heading.Group
        $onHeading = { if (-not $ctx.State.Busy) { & $toggleGroup $ctx $groupName } }.GetNewClosure()
        $heading.Check.Add_Checked($onHeading)
        $heading.Check.Add_Unchecked($onHeading)
        $heading.Check.Add_Indeterminate($onHeading)
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
        & $setChecks $ctx $ctx.Rows $allOn
    }.GetNewClosure())
    $ctrl.TextExtUncheckAllButton.Add_Click({
        & $setChecks $ctx $ctx.Rows $allOff
    }.GetNewClosure())
    $ctrl.TextExtOkButton.Add_Click({ $dialog.DialogResult = $true }.GetNewClosure())
    & $refresh $ctx

    if ($true -ne (showOwnedDialog $dialog)) {
        return $null
    }
    $chosen = @($ctx.Rows | ForEach-Object { @{ Extension = $_.Extension; Checked = [bool]$_.Check.IsChecked } })
    return , @(convertIndexTextExtensionRows $chosen)
}
