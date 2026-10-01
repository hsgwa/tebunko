# poc/wpf-figma/show.ps1
#
# 26 枚のフレームを選んで、実物の画面で見た目を確かめるための小さなビューア。
# - フレームを選ぶ一覧（左）
# - 選んだフレームを原寸で表示する窓（右）。「重ねて見る」で参照 PNG を半透明に重ね、
#   「並べて見る」で参照 PNG を横に並べる。
#
# 使い方: pwsh -File .\show.ps1

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$root = $PSScriptRoot
. (Join-Path $root "dummy.ps1")

# ---- フォント: Rethink Sans をインストールせずにファイルから使う ----
$fontsDir = (Join-Path $root "fonts") -replace "\\","/"
$fontUri = "file:///$fontsDir/#Rethink Sans, Yu Gothic UI"

function New-FrameRoot([string]$FrameName) {
    $frame = Get-FigmaFrame $FrameName
    if ($null -eq $frame) { throw "不明なフレーム: $FrameName" }
    $xamlPath = Join-Path $root $frame.Xaml
    $reader = [System.Xml.XmlReader]::Create($xamlPath)
    $visual = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()

    if ($frame.Xaml -eq "xaml\search.xaml") {
        Set-FigmaFrameState $visual $FrameName
    }

    # theme.xaml の Font.UI（既定は Yu Gothic UI）を、実物のフォントファイルに差し替える。
    $realFont = New-Object System.Windows.Media.FontFamily($fontUri)
    Set-RealFont $visual $realFont

    $visual.Measure((New-Object System.Windows.Size($frame.Width, $frame.Height)))
    $visual.Arrange((New-Object System.Windows.Rect(0, 0, $frame.Width, $frame.Height)))
    return @{ Visual = $visual; Frame = $frame }
}

function Get-ReferenceImage([string]$FrameName) {
    $path = Join-Path $root "reference\$FrameName.png"
    if (-not (Test-Path $path)) { return $null }
    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
    $bmp.BeginInit()
    $bmp.UriSource = New-Object System.Uri($path)
    $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bmp.EndInit()
    $bmp.Freeze()
    return $bmp
}

# ---- 一覧の窓 ----
$picker = New-Object System.Windows.Window
$picker.Title = "poc/wpf-figma — フレームを選ぶ"
$picker.Width = 260
$picker.Height = 640
$picker.WindowStartupLocation = "Manual"
$picker.Left = 20
$picker.Top = 40

$dock = New-Object System.Windows.Controls.DockPanel
$picker.Content = $dock

$modePanel = New-Object System.Windows.Controls.StackPanel
$modePanel.Orientation = "Horizontal"
$modePanel.Margin = "8"
[System.Windows.Controls.DockPanel]::SetDock($modePanel, "Top")
$dock.Children.Add($modePanel) | Out-Null

$btnOverlay = New-Object System.Windows.Controls.RadioButton
$btnOverlay.Content = "重ねて見る"
$btnOverlay.GroupName = "mode"
$btnOverlay.IsChecked = $true
$btnOverlay.Margin = "0,0,10,0"
$modePanel.Children.Add($btnOverlay) | Out-Null

$btnSide = New-Object System.Windows.Controls.RadioButton
$btnSide.Content = "並べて見る"
$btnSide.GroupName = "mode"
$modePanel.Children.Add($btnSide) | Out-Null

$list = New-Object System.Windows.Controls.ListBox
foreach ($f in Get-FigmaFrames) { $list.Items.Add($f.Name) | Out-Null }
$dock.Children.Add($list) | Out-Null

# ---- 表示する窓（選んだフレームを原寸で、必要なら参照 PNG と一緒に出す） ----
$viewerWindow = $null

function Show-Frame([string]$FrameName) {
    $result = New-FrameRoot $FrameName
    $visual = $result.Visual
    $frame = $result.Frame
    $refImage = Get-ReferenceImage $FrameName

    if ($null -ne $script:viewerWindow) { $script:viewerWindow.Close() }

    $win = New-Object System.Windows.Window
    $win.Title = "$FrameName（$($frame.Width)x$($frame.Height)）"
    $win.WindowStartupLocation = "Manual"
    $win.Left = 300
    $win.Top = 40
    $win.SizeToContent = "WidthAndHeight"
    $win.ResizeMode = "CanMinimize"

    if ($btnSide.IsChecked -eq $true -and $null -ne $refImage) {
        # 並べて見る: 左に実物、右に参照 PNG。
        $sidePanel = New-Object System.Windows.Controls.StackPanel
        $sidePanel.Orientation = "Horizontal"

        $leftHost = New-Object System.Windows.Controls.Border
        $leftHost.Width = $frame.Width
        $leftHost.Height = $frame.Height
        $leftHost.Child = $visual
        $sidePanel.Children.Add($leftHost) | Out-Null

        $img = New-Object System.Windows.Controls.Image
        $img.Source = $refImage
        $img.Width = $frame.Width
        $img.Height = $frame.Height
        $sidePanel.Children.Add($img) | Out-Null

        $win.Content = $sidePanel
    } else {
        # 重ねて見る: 実物の上に参照 PNG を半透明で重ねる（ズレが目で分かる）。
        $grid = New-Object System.Windows.Controls.Grid
        $grid.Width = $frame.Width
        $grid.Height = $frame.Height
        $grid.Children.Add($visual) | Out-Null
        if ($null -ne $refImage) {
            $img = New-Object System.Windows.Controls.Image
            $img.Source = $refImage
            $img.Width = $frame.Width
            $img.Height = $frame.Height
            $img.Opacity = 0.5
            $img.IsHitTestVisible = $false
            $grid.Children.Add($img) | Out-Null
        }
        $win.Content = $grid
    }

    $win.Show()
    $script:viewerWindow = $win
}

$list.Add_SelectionChanged({
    if ($list.SelectedItem) { Show-Frame $list.SelectedItem }
})

$picker.Add_Closed({
    if ($null -ne $script:viewerWindow) { $script:viewerWindow.Close() }
})

$picker.ShowDialog() | Out-Null
