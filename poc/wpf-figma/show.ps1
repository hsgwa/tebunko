# poc/wpf-figma/show.ps1
#
# 26 枚のフレームを選んで、実物の画面で見た目を確かめるための小さなビューア。
# - フレームを選ぶ一覧（左）
# - 選んだフレームを原寸で表示する窓（右）。WPF の画面だけを表示する。
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

$list = New-Object System.Windows.Controls.ListBox
foreach ($f in Get-FigmaFrames) { $list.Items.Add($f.Name) | Out-Null }
$dock.Children.Add($list) | Out-Null

# ---- 表示する窓（選んだフレームを原寸の WPF 画面だけで出す） ----
$viewerWindow = $null

function Show-Frame([string]$FrameName) {
    $result = New-FrameRoot $FrameName
    $visual = $result.Visual
    $frame = $result.Frame

    if ($null -ne $script:viewerWindow) { $script:viewerWindow.Close() }

    $win = New-Object System.Windows.Window
    $win.Title = "$FrameName（$($frame.Width)x$($frame.Height)）"
    $win.WindowStartupLocation = "Manual"
    $win.Left = 300
    $win.Top = 40
    $win.SizeToContent = "WidthAndHeight"
    $win.ResizeMode = "CanMinimize"
    $win.Content = $visual

    # search.xaml は画面の中に自前のタイトルバー（TitleBarGrid。32px・「tebunko」の文字と
    # 最小化／最大化／閉じるの絵）を描いている。既定の WindowStyle のままだと、OS の枠のバーと
    # この自前のバーが縦に二重に並んで見える（メンテナ指摘 2026-10-03・動作確認で）。
    # search.xaml を使うフレームだけ OS の枠を消し、自前のバーだけを見せる。
    # 自前のバーを持たない見本（splash.xaml など）は、従来どおり OS の枠のバーのままにする。
    if ($frame.Xaml -eq "xaml\search.xaml") {
        $win.WindowStyle = "None"
        $win.ResizeMode = "NoResize"
        $titleBar = $visual.FindName("TitleBarGrid")
        if ($null -ne $titleBar) {
            # OS の枠を消すとドラッグでの移動もできなくなるため、自前のバーの上で
            # 左ボタンを押したときだけ DragMove で移動できるようにする。
            $titleBar.Add_MouseLeftButtonDown({
                param($sender, $e)
                $w = [System.Windows.Window]::GetWindow($sender)
                if ($null -ne $w) { $w.DragMove() }
            })
        }
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
