# 文字が省略（…）で切れているときだけ、全文をツールチップで見せる（画面層）。
# 切れているかの判定は shared\core\text.ps1 の testTextTrimmed。ここは TextBlock の幅を測って渡すだけで、このツール固有の文言は持たない。
# ToolTip には全文を入れておき、開く直前（ToolTipOpening）に測って、切れていなければ e.Handled = $true にして開かない。

function testTextBlockTrimmed {
    # TextBlock が今の幅で文字を省略しているか。要る幅は、同じ書体・大きさで折り返さずに並べたときの幅
    param (
        [System.Windows.Controls.TextBlock]$textBlock
    )

    $text = [string]$textBlock.Text
    if ($text -eq "" -or $textBlock.ActualWidth -le 0) {
        return $false
    }
    $typeface = New-Object System.Windows.Media.Typeface($textBlock.FontFamily, $textBlock.FontStyle, $textBlock.FontWeight, $textBlock.FontStretch)
    $dpi = [System.Windows.Media.VisualTreeHelper]::GetDpi($textBlock).PixelsPerDip
    $formatted = New-Object System.Windows.Media.FormattedText($text, [System.Globalization.CultureInfo]::CurrentUICulture,
        [System.Windows.FlowDirection]::LeftToRight, $typeface, $textBlock.FontSize, [System.Windows.Media.Brushes]::Black, $dpi)
    return (testTextTrimmed $textBlock.ActualWidth $formatted.WidthIncludingTrailingWhitespace)
}

function addTrimmedToolTip {
    # TextBlock に、切れているときだけ開くツールチップ（全文）を付ける。Text が変わっても、開く直前の中身を使う。
    # DataGrid などの行の中の TextBlock は、行の Loaded で渡して付ける（addTrimmedToolTip $cell）
    param (
        [System.Windows.Controls.TextBlock]$textBlock
    )

    $textBlock.ToolTip = [string]$textBlock.Text
    $textBlock.Add_ToolTipOpening({
        param ($sender, $e)
        $sender.ToolTip = [string]$sender.Text
        if (!(testTextBlockTrimmed $sender)) {
            $e.Handled = $true
        }
    })
}

function addTrimmedToolTipToRows {
    # 行の Loaded（DataGridRow など）で、その行の中の Tag が tag の TextBlock に addTrimmedToolTip を付ける（付けたら Tag を「<tag>Set」にして、二重に付けない）。
    # 行の Tag に tag を入れ（Loaded のハンドラはここから読む。GetNewClosure にすると画面の関数が見えなくなる）、
    # 同じ行に 2 回呼ぶとハンドラが重なるので、呼ぶ側が行の Tag が tag かで見て、行ごとに 1 回だけにする
    param (
        [System.Windows.FrameworkElement]$row,
        [string]$tag
    )

    $row.Tag = $tag
    $row.Add_Loaded({
        param ($sender, $e)
        $tag = [string]$sender.Tag
        safe {
            $stack = New-Object System.Collections.Generic.Stack[System.Windows.DependencyObject]
            $stack.Push($sender)
            while ($stack.Count -gt 0) {
                $node = $stack.Pop()
                for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($node); $i++) {
                    $child = [System.Windows.Media.VisualTreeHelper]::GetChild($node, $i)
                    if ($child -is [System.Windows.Controls.TextBlock] -and $child.Tag -eq $tag) {
                        addTrimmedToolTip $child
                        $child.Tag = "${tag}Set"
                    }
                    $stack.Push($child)
                }
            }
        }
    })
}
