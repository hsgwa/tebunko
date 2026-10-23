# テキストファイルの拡張子の扱いを選ぶダイアログ（画面層）。screens.ps1 が読み込む。
# 文言・行・設定に書く値の作り方は判断層（text_ext_view.ps1）、設定の読み書きは core\settings.ps1。
# ここは、それらをつなぎ、ダイアログ（dialog_text_extensions.xaml）を開いて［OK］で保存するだけ。
# 開く入口は 2 つ: ［インデックス管理］の［アクション］メニュー（ActionTextExtensions）と、［設定］の枠（ChangeTextExtensionsButton）。

function updateTextExtensionSummary {
    # ［設定］の枠に、今の状態の要約を出す
    try {
        $current = readTextExtensionSettings
        $ui.TextExtensionSummary.Text = getTextExtensionSummary $current.Excluded $current.Notepad
    } catch {
        $ui.TextExtensionSummary.Text = ""
    }
}

function addTextExtensionRow {
    # ダイアログの 1 行（拡張子の文字と、扱いを選ぶ欄）を panel に足し、選ぶ欄（ComboBox。Tag に拡張子）を返す
    param (
        $panel,
        $row   # getTextExtensionRows の 1 行
    )

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(10, 4, 10, 4)
    $first = New-Object System.Windows.Controls.ColumnDefinition
    $first.Width = New-Object System.Windows.GridLength(130)
    [void]$grid.ColumnDefinitions.Add($first)
    [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = $row.Extension
    $label.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($label, 0)
    $combo = New-Object System.Windows.Controls.ComboBox
    $combo.Tag = $row.Extension
    # 項目は ComboBoxItem（Content = 文言、Tag = 扱いの値）。選んだ値は SelectedItem の Tag で読む
    foreach ($option in $row.Options) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $option.Text
        $item.Tag = $option.Value
        [void]$combo.Items.Add($item)
        if ($option.Value -eq $row.Way) {
            $combo.SelectedItem = $item
        }
    }
    $combo.Width = 200
    $combo.HorizontalAlignment = "Left"
    $combo.SetValue([System.Windows.Automation.AutomationProperties]::NameProperty, "$($row.Extension) の扱い")
    [System.Windows.Controls.Grid]::SetColumn($combo, 1)
    [void]$grid.Children.Add($label)
    [void]$grid.Children.Add($combo)
    [void]$panel.Children.Add($grid)
    return $combo
}

function showTextExtensionDialog {
    # ダイアログを開く。［OK］で、選んだ扱いを設定に保存して要約を更新する。保存できなければ知らせて、設定は変えない
    $texts = getTextExtensionDialogTexts
    $current = readTextExtensionSettings
    $rows = @(getTextExtensionRows $current.Excluded $current.Notepad)

    $dialog = loadWindow "${xamlDir}\dialog_text_extensions.xaml"
    $dialog.Owner = $window
    $dialog.Title = $texts.Title
    $ctrl = @{}
    foreach ($name in @("NoteText", "ExtensionHeader", "WayHeader", "RowsPanel", "TextExtResetButton", "TextExtCancelButton", "TextExtOkButton")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $ctrl.NoteText.Text = $texts.Note
    $ctrl.ExtensionHeader.Text = $texts.Extension
    $ctrl.WayHeader.Text = $texts.Way
    $ctrl.TextExtResetButton.Content = $texts.Reset
    $ctrl.TextExtCancelButton.Content = $texts.Cancel
    $ctrl.TextExtOkButton.Content = $texts.Ok

    # 行ごとに、拡張子の文字と、扱いを選ぶ欄（ComboBox。Tag に拡張子）を並べる
    $combos = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $combos.Add((addTextExtensionRow $ctrl.RowsPanel $row))
    }

    # ［既定に戻す］: 欄を既定の扱いにするだけ（保存は［OK］）
    $defaults = @(getTextExtensionRows @() @())
    $ctrl.TextExtResetButton.Add_Click({
        for ($i = 0; $i -lt $combos.Count; $i++) {
            foreach ($item in $combos[$i].Items) {
                if ($item.Tag -eq $defaults[$i].Way) {
                    $combos[$i].SelectedItem = $item
                }
            }
        }
    }.GetNewClosure())
    $ctrl.TextExtOkButton.Add_Click({ $dialog.DialogResult = $true }.GetNewClosure())

    if ($true -ne (showOwnedDialog $dialog)) {
        return
    }
    $chosen = @($combos | ForEach-Object { @{ Extension = [string]$_.Tag; Way = [string]$_.SelectedItem.Tag } })
    $values = convertTextExtensionRows $chosen
    try {
        writeTextExtensionSettings $values.Excluded $values.Notepad
    } catch {
        showMessage "設定を保存できませんでした。もう一度お試しください。" "OK" "Warning" | Out-Null
        return
    }
    updateTextExtensionSummary
}
