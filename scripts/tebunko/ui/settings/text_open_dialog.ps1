# テキストの開き方（拡張子ごとに「既定のアプリで開く」か「エディタで開く」か）を選ぶダイアログと、［設定］の枠の表示（画面層）。screens.ps1 が読み込む。
# 文言・行・一覧の見せ方・設定に書く値の作り方は判断層（text_open_view.ps1）、設定の読み書きは core\settings.ps1。
# ここは、それらをつなぎ、ダイアログ（dialog_text_open.xaml）を開いて［OK］で保存するだけ。入口は［設定］の「テキストファイル」の枠の［変更…］（ChangeTextExtensionsButton）。
# 開き方は行ごとに 2 つのボタン（RadioButton。Tag が開き方の値）を 1 回押して選ぶ。選んだ値は、その行の RadioButton の IsChecked で読む。
# エディタの行: ［変更…］で .exe を選ぶ（pickAndSaveTextEditor。保存の前に getTextEditorPickResult で確かめる）。［メモ帳に戻す］は resetTextEditor。
# 選んだエディタのフルパスと、選べなかった理由は、行の下の段に出す（［設定］の行と同じ。updateTextEditorPathView・setTextEditorError）。

function updateTextEditorPathView {
    # エディタのフルパスの段（選んだエディタがあるときだけ出す）と、［メモ帳に戻す］の押せる・押せないを直す
    param (
        $pathBlock,
        $resetButton,
        $configured
    )

    $pathText = getTextEditorPathText $configured
    $pathBlock.Text = $pathText
    $pathBlock.Visibility = $(if ($pathText -ne "") { "Visible" } else { "Collapsed" })
    $resetButton.IsEnabled = testTextEditorResetEnabled $configured
}

function setTextEditorError {
    # エディタを選べなかった理由の段（赤字）。空なら隠す
    param (
        $errorBlock,
        [string]$message
    )

    $errorBlock.Text = $message
    $errorBlock.Visibility = $(if ($message -ne "") { "Visible" } else { "Collapsed" })
}

function updateTextOpenSummary {
    # ［設定］の枠に、開き方の要約とエディタの名前・パスを出す
    $texts = getTextOpenDialogTexts
    $ui.TextExtensionHint.Text = $texts.SettingsHint
    $editorPath = ""
    try {
        $current = readTextExtensionSettings
        $ui.TextExtensionSummary.Text = getTextOpenSummary $current.Notepad
        $editorPath = $current.EditorPath
    } catch {
        $ui.TextExtensionSummary.Text = ""
    }
    $ui.EditorHint.Text = ($texts.EditorSettingsHint -f (getTextEditorLabel $editorPath))
    updateTextEditorPathView $ui.EditorPathText $ui.ResetEditorButton $editorPath
}

function pickAndSaveTextEditor {
    # ［変更…］（［設定］の行・ダイアログの行の両方）から呼ぶ。.exe を選んでもらい、getTextEditorPickResult が通ったときだけ textEditorPath に保存する。
    # 返す値: @{ Picked（選んだか。キャンセルは $false）; Saved; Message（選べなかった理由。無ければ空） }
    param (
        [System.Windows.Window]$owner = $window
    )

    $texts = getTextOpenDialogTexts
    $initial = ""
    try {
        $dir = [System.IO.Path]::GetDirectoryName([string](readTextExtensionSettings).EditorPath)
        if ($dir -and (Test-Path -LiteralPath $dir -PathType Container)) {
            $initial = $dir
        }
    } catch {
        $initial = ""
    }
    if ($initial -eq "") {
        # 設定のエディタが無いときは、Windows のあるドライブの Program Files から始める（無ければ指定しない。OS の既定の場所）
        $programFiles = Join-Path ([System.IO.Path]::GetPathRoot($env:SystemRoot)) "Program Files"
        if (Test-Path -LiteralPath $programFiles -PathType Container) {
            $initial = $programFiles
        }
    }
    $chosen = selectExeFile $texts.EditorPickTitle $initial $owner
    if (-not $chosen) {
        return @{ Picked = $false; Saved = $false; Message = "" }
    }
    # ファイルを調べるのは、ドライブ文字から始まる形のときだけ（UNC など、届かないと止まりうる場所は調べず、形の違いとして断る）
    $exists = $false
    $versionNames = $null
    if (testTextEditorPath $chosen) {
        $exists = Test-Path -LiteralPath $chosen -PathType Leaf
        if ($exists) {
            $versionNames = getTextEditorVersionNames $chosen
        }
    }
    $result = getTextEditorPickResult $chosen $exists $versionNames
    if (-not $result.Ok) {
        return @{ Picked = $true; Saved = $false; Message = $result.Message }
    }
    try {
        writeTextEditorPath $chosen
    } catch {
        return @{ Picked = $true; Saved = $false; Message = "設定を保存できませんでした。もう一度お試しください。" }
    }
    return @{ Picked = $true; Saved = $true; Message = "" }
}

function resetTextEditor {
    # ［メモ帳に戻す］: textEditorPath を空にして保存する。保存できたか（できなければ知らせる）を返す
    try {
        writeTextEditorPath ""
    } catch {
        showMessage "設定を保存できませんでした。もう一度お試しください。" "OK" "Warning" | Out-Null
        return $false
    }
    return $true
}

function changeTextEditorFromSettings {
    # ［設定］の「エディタ」の［変更…］
    setTextEditorError $ui.EditorErrorText ""
    $result = pickAndSaveTextEditor
    if (-not $result.Picked) {
        return
    }
    updateTextOpenSummary
    setTextEditorError $ui.EditorErrorText $result.Message
}

function resetTextEditorFromSettings {
    # ［設定］の「エディタ」の［メモ帳に戻す］
    setTextEditorError $ui.EditorErrorText ""
    if (resetTextEditor) {
        updateTextOpenSummary
    }
}

function updateTextOpenDialogEditor {
    # ダイアログの下のエディタの行（名前・パス・［メモ帳に戻す］）を、今の設定で直す
    param (
        $ctrl
    )

    $editorPath = (readTextExtensionSettings).EditorPath
    $ctrl.TextExtEditorName.Text = getTextEditorLabel $editorPath
    updateTextEditorPathView $ctrl.TextExtEditorPath $ctrl.TextExtEditorResetButton $editorPath
}

function changeTextEditorInDialog {
    # ダイアログの下のエディタの行の［変更…］。保存は押した時点（［OK］を待たない）。ダイアログの行と［設定］の行の両方を直す
    param (
        $ctrl
    )

    setTextEditorError $ctrl.TextExtEditorError ""
    $result = pickAndSaveTextEditor $ctrl.Dialog
    if (-not $result.Picked) {
        return
    }
    updateTextOpenDialogEditor $ctrl
    updateTextOpenSummary
    setTextEditorError $ctrl.TextExtEditorError $result.Message
}

function resetTextEditorInDialog {
    # ダイアログの下のエディタの行の［メモ帳に戻す］
    param (
        $ctrl
    )

    setTextEditorError $ctrl.TextExtEditorError ""
    if (resetTextEditor) {
        updateTextOpenDialogEditor $ctrl
        updateTextOpenSummary
    }
}

function newTextOpenWayRadio {
    # 開き方を選ぶボタン 1 つ（RadioButton。見た目は dialog_text_open.xaml の WayRadio）
    param (
        $dialog,
        [string]$group,
        [string]$way,
        [bool]$isLast
    )

    $radio = New-Object System.Windows.Controls.RadioButton
    $radio.Style = $dialog.FindResource("WayRadio")
    $radio.GroupName = $group
    $radio.Tag = $way
    $radio.Content = getTextOpenWayText $way
    if ($isLast) {
        $radio.BorderThickness = New-Object System.Windows.Thickness(0)
        $radio.Width = 104
    }
    return $radio
}

function newTextOpenRowGrid {
    # 1 行の土台（変えた印の点・拡張子・開き方の欄・注意書きの 4 列）
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Height = 33
    foreach ($width in @(20, 110, 224)) {
        $column = New-Object System.Windows.Controls.ColumnDefinition
        $column.Width = New-Object System.Windows.GridLength($width)
        [void]$grid.ColumnDefinitions.Add($column)
    }
    [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
    $grid.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
    return $grid
}

function addTextOpenRow {
    # ダイアログの 1 行（変えた印の点・拡張子・開き方の 2 つのボタン・注意書き）を panel に足し、行の部品を返す:
    # @{ Extension; Line; Dot; Radios（開き方の値 → RadioButton） }
    param (
        $dialog,
        $panel,
        $row,   # getTextOpenRows の 1 行
        $texts
    )

    $grid = newTextOpenRowGrid
    $dot = New-Object System.Windows.Shapes.Ellipse
    $dot.Width = 6
    $dot.Height = 6
    $dot.Fill = $dialog.FindResource("Accent")
    $dot.Opacity = 0
    $dot.HorizontalAlignment = "Left"
    $dot.VerticalAlignment = "Center"
    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = $row.Extension
    $label.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($label, 1)

    $ways = New-Object System.Windows.Controls.StackPanel
    $ways.Orientation = "Horizontal"
    $ways.Height = 24
    $container = New-Object System.Windows.Controls.Border
    $container.CornerRadius = New-Object System.Windows.CornerRadius(6)
    $container.BorderThickness = New-Object System.Windows.Thickness(1)
    $container.BorderBrush = $dialog.FindResource("Border.Normal")
    $container.ClipToBounds = $true
    $container.Height = 26
    $container.HorizontalAlignment = "Left"
    $container.VerticalAlignment = "Center"
    $container.Child = $ways
    [System.Windows.Controls.Grid]::SetColumn($container, 2)

    $radios = @{}
    foreach ($way in @("default", "editor")) {
        $radio = newTextOpenWayRadio $dialog "way_$($row.Extension)" $way ($way -eq "editor")
        # 選べない開き方（実行・登録になる形式の「既定のアプリで開く」）は押せなくする
        $radio.IsEnabled = (@($row.Options | Where-Object { $_.Value -eq $way }).Count -gt 0)
        $radio.IsChecked = ($way -eq $row.Way)
        $radio.SetValue([System.Windows.Automation.AutomationProperties]::NameProperty, "$($row.Extension) ： $($radio.Content)")
        [void]$ways.Children.Add($radio)
        $radios[$way] = $radio
    }
    foreach ($child in @($dot, $label, $container)) {
        [void]$grid.Children.Add($child)
    }
    if (${textNotepadExtensions} -contains $row.Extension) {
        $note = New-Object System.Windows.Controls.TextBlock
        $note.Text = $texts.RestrictedNote
        $note.Style = $dialog.FindResource("Micro")
        $note.VerticalAlignment = "Center"
        $note.Margin = New-Object System.Windows.Thickness(12, 0, 0, 0)
        [System.Windows.Controls.Grid]::SetColumn($note, 3)
        [void]$grid.Children.Add($note)
    }
    $line = New-Object System.Windows.Controls.Border
    $line.BorderBrush = $dialog.FindResource("Border.Row")
    $line.BorderThickness = New-Object System.Windows.Thickness(0, 0, 0, 1)
    $line.Child = $grid
    [void]$panel.Children.Add($line)
    return @{ Extension = $row.Extension; Line = $line; Dot = $dot; Label = $label; Radios = $radios }
}

function getTextOpenControlWay {
    # 行の部品で、今選ばれている開き方の値
    param (
        $rowControl
    )

    foreach ($way in $rowControl.Radios.Keys) {
        if ($rowControl.Radios[$way].IsChecked) {
            return [string]$way
        }
    }
    return "default"
}

function updateTextOpenDialogView {
    # 選びが変わったとき・絞り込みが変わったときに、一覧の見せ方（行の表示・変えた印・要約・一致なしの文）を直す。
    # ctx: @{ Rows（addTextOpenRow の返す部品の並び）; Controls; RestrictedHeader }
    param (
        $ctx
    )

    $current = @($ctx.Rows | ForEach-Object { @{ Extension = $_.Extension; Way = (getTextOpenControlWay $_) } })
    $view = getTextOpenListView $current ([string]$ctx.Controls.TextExtFilterBox.Text) ([bool]$ctx.Controls.TextExtChangedOnly.IsChecked)
    foreach ($rowControl in $ctx.Rows) {
        $way = getTextOpenControlWay $rowControl
        $rowControl.Line.Visibility = $(if ($view.Visible -contains $rowControl.Extension) { "Visible" } else { "Collapsed" })
        $rowControl.Dot.Opacity = $(if ($way -ne (getTextOpenDefaultWay $rowControl.Extension)) { 1 } else { 0 })
    }
    $ctx.RestrictedHeader.Visibility = $(if ($view.ShowRestrictedHeader) { "Visible" } else { "Collapsed" })
    $ctx.Controls.TextExtSummary.Text = $view.Summary
    $ctx.Controls.TextExtEmpty.Visibility = $(if ($view.Empty) { "Visible" } else { "Collapsed" })
    $ctx.Controls.TextExtFilterPlaceholder.Visibility = $(if ([string]$ctx.Controls.TextExtFilterBox.Text -eq "") { "Visible" } else { "Collapsed" })
}

function newTextOpenDialogControls {
    # ダイアログを読み込み、文言を入れて、部品の表（名前 → 部品）を返す
    param (
        $texts,
        $editorPath
    )

    $dialog = loadWindow "${xamlDir}\dialog_text_open.xaml"
    $dialog.Owner = $window
    $dialog.Title = $texts.Title
    $ctrl = @{ Dialog = $dialog }
    foreach ($name in @("NoteText", "ExtensionHeader", "WayHeader", "RowsPanel", "TextExtList", "TextExtEmpty", "TextExtSummary", "TextExtChangedOnly",
            "TextExtFilterBox", "TextExtFilterPlaceholder", "TextExtResetButton", "TextExtCancelButton", "TextExtOkButton",
            "TextExtEditorTitle", "TextExtEditorName", "TextExtEditorHint", "TextExtEditorChangeButton", "TextExtEditorResetButton", "TextExtEditorPath", "TextExtEditorError")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $ctrl.NoteText.Text = $texts.Note
    $ctrl.ExtensionHeader.Text = $texts.Extension
    $ctrl.WayHeader.Text = $texts.Way
    $ctrl.TextExtEmpty.Text = $texts.Empty
    $ctrl.TextExtFilterPlaceholder.Text = $texts.FilterPlaceholder
    $ctrl.TextExtChangedOnly.Content = $texts.ChangedOnly
    $ctrl.TextExtResetButton.Content = $texts.Reset
    $ctrl.TextExtCancelButton.Content = $texts.Cancel
    $ctrl.TextExtOkButton.Content = $texts.Ok
    $ctrl.TextExtEditorTitle.Text = $texts.EditorTitle
    $ctrl.TextExtEditorName.Text = getTextEditorLabel $editorPath
    $ctrl.TextExtEditorHint.Text = $texts.EditorHint
    $ctrl.TextExtEditorChangeButton.Content = $texts.Change
    $ctrl.TextExtEditorResetButton.Content = $texts.ResetEditor
    updateTextEditorPathView $ctrl.TextExtEditorPath $ctrl.TextExtEditorResetButton $editorPath
    return $ctrl
}

function addTextOpenRows {
    # 一覧の行を足す（実行・登録になる形式は末尾にまとめ、その前に見出しの行を置く）。行の部品の並びと、見出しの部品を返す
    param (
        $ctrl,
        $rows,
        $texts
    )

    $rowControls = New-Object System.Collections.Generic.List[object]
    $header = New-Object System.Windows.Controls.TextBlock
    $header.Text = $texts.RestrictedHeader
    $header.Style = $ctrl.Dialog.FindResource("Meta")
    $header.FontWeight = "SemiBold"
    $header.Padding = New-Object System.Windows.Thickness(10, 8, 0, 8)
    $header.Background = $ctrl.Dialog.FindResource("Bg.Subtle")
    $headerAdded = $false
    foreach ($row in $rows) {
        if (-not $headerAdded -and (${textNotepadExtensions} -contains $row.Extension)) {
            [void]$ctrl.RowsPanel.Children.Add($header)
            $headerAdded = $true
        }
        $rowControls.Add((addTextOpenRow $ctrl.Dialog $ctrl.RowsPanel $row $texts))
    }
    return @{ Rows = $rowControls.ToArray(); Header = $header }
}

function showTextOpenDialog {
    # ダイアログを開く。［OK］で、選んだ開き方を設定に保存して要約を更新する。保存できなければ知らせて、設定は変えない
    $texts = getTextOpenDialogTexts
    $current = readTextExtensionSettings
    $rows = @(getTextOpenRows $current.Notepad)
    $ctrl = newTextOpenDialogControls $texts $current.EditorPath
    $added = addTextOpenRows $ctrl $rows $texts
    $ctx = @{ Rows = $added.Rows; Controls = $ctrl; RestrictedHeader = $added.Header }

    # 選び・絞り込みが変わったら、一覧の見せ方を直す（ハンドラは関数を名前で呼ばず、実体を変数に取って & で呼ぶ）
    $refresh = ${function:updateTextOpenDialogView}
    $onChange = { & $refresh $ctx }.GetNewClosure()
    foreach ($rowControl in $ctx.Rows) {
        foreach ($radio in $rowControl.Radios.Values) {
            $radio.Add_Checked($onChange)
        }
    }
    $ctrl.TextExtFilterBox.Add_TextChanged($onChange)
    # マウス・キーボードでも UI オートメーションの Toggle でも IsChecked が変わるので、Click ではなく Checked / Unchecked で受ける
    $ctrl.TextExtChangedOnly.Add_Checked($onChange)
    $ctrl.TextExtChangedOnly.Add_Unchecked($onChange)

    # ［既定に戻す］: 行を既定の開き方にするだけ（保存は［OK］）
    $defaults = @(getTextOpenRows @())
    $ctrl.TextExtResetButton.Add_Click({
        for ($i = 0; $i -lt $ctx.Rows.Count; $i++) {
            $ctx.Rows[$i].Radios[$defaults[$i].Way].IsChecked = $true
        }
    }.GetNewClosure())
    # エディタの行（ハンドラは関数を名前で呼ばず、実体を変数に取って & で呼ぶ）
    $changeEditor = ${function:changeTextEditorInDialog}
    $resetEditor = ${function:resetTextEditorInDialog}
    $ctrl.TextExtEditorChangeButton.Add_Click({ & $changeEditor $ctrl }.GetNewClosure())
    $ctrl.TextExtEditorResetButton.Add_Click({ & $resetEditor $ctrl }.GetNewClosure())
    $dialog = $ctrl.Dialog
    $ctrl.TextExtOkButton.Add_Click({ $dialog.DialogResult = $true }.GetNewClosure())
    & $refresh $ctx

    if ($true -ne (showOwnedDialog $dialog)) {
        return
    }
    $chosen = @($ctx.Rows | ForEach-Object { @{ Extension = $_.Extension; Way = (getTextOpenControlWay $_) } })
    try {
        writeTextOpenSettings @(addOutOfListExtensions @(convertTextOpenRows $chosen) (readTextExtensionSettings).Notepad)
    } catch {
        showMessage "設定を保存できませんでした。もう一度お試しください。" "OK" "Warning" | Out-Null
        return
    }
    updateTextOpenSummary
}
