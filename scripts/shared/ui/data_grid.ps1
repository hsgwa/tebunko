# DataGrid の共通の部品（画面層）。このツール固有の名前・文言は持たない。

function getDataGridRowAt {
    # マウスの下（またはイベントの発生元）の行（DataGridRow）。列見出し・スクロールバー・行の外（余白）では $null
    param (
        $source
    )

    $element = $source
    while ($element -and $element -isnot [System.Windows.Controls.DataGridRow]) {
        if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
            return $null
        }
        $element = $(if ($element -is [System.Windows.Media.Visual]) { [System.Windows.Media.VisualTreeHelper]::GetParent($element) } else { $null })
    }
    return $element
}
