# poc/wpf-figma/screenshots.ps1
#
# 確かめ用のスクリーンショットを out\screens\ に書き出す（.gitignore 済み）。
# - 26 枚のフレームを、それぞれの既定の大きさ（多くは 1280x820）で 1 枚ずつ。
# - H・H0・H-row3 の 3 枚だけ、窓の最小 1024x640 と 1600x1000 でも追加で撮る
#   （figma_wpf_map.md「レイアウトとリサイズ」の見本どおり）。
#
# 使い方: pwsh -File .\screenshots.ps1

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$root = $PSScriptRoot
. (Join-Path $root "dummy.ps1")

$outDir = Join-Path $root "out\screens"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$fontsDir = (Join-Path $root "fonts") -replace "\\","/"
$fontUri = "file:///$fontsDir/#Rethink Sans, Yu Gothic UI"
$realFont = New-Object System.Windows.Media.FontFamily($fontUri)

# Popup（開くメニュー）は、ヘッドレスの RenderTargetBitmap では自前の層が写らないため、
# Child を手で Measure/Arrange して別に描き、DrawingVisual で本体の絵に重ねる
# （diff_round3.md 3 ★。PlacementTarget が隠れている open-menu フレームだけ、
# 旧来の固定座標 1000,645 にフォールバックする）。
function Merge-PopupOverlay($visual, $rtb, [int]$Width, [int]$Height) {
    $popup = $visual.FindName("ContextMenuPopup")
    if ($null -eq $popup -or $popup.Tag -ne "open") { return $rtb }
    $child = $popup.Child
    if ($null -eq $child) { return $rtb }

    $child.Measure((New-Object System.Windows.Size([double]::PositiveInfinity, [double]::PositiveInfinity)))
    $cw = [Math]::Max(1, [Math]::Ceiling($child.DesiredSize.Width))
    $ch = [Math]::Max(1, [Math]::Ceiling($child.DesiredSize.Height))
    $child.Arrange((New-Object System.Windows.Rect(0, 0, $cw, $ch)))
    $child.UpdateLayout()

    $childRtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(
        $cw, $ch, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $childRtb.Render($child)

    $target = $popup.PlacementTarget
    if ($null -ne $target -and $target.ActualWidth -gt 0 -and $target.ActualHeight -gt 0) {
        $topLeft = $target.TransformToAncestor($visual).Transform((New-Object System.Windows.Point(0, 0)))
        $posX = $topLeft.X + $popup.HorizontalOffset
        $posY = $topLeft.Y + $target.ActualHeight + $popup.VerticalOffset
    } else {
        $posX = 1000
        $posY = 645
    }

    $dv = New-Object System.Windows.Media.DrawingVisual
    $ctx = $dv.RenderOpen()
    $ctx.DrawImage($rtb, (New-Object System.Windows.Rect(0, 0, $Width, $Height)))
    $ctx.DrawImage($childRtb, (New-Object System.Windows.Rect($posX, $posY, $cw, $ch)))
    $ctx.Close()

    $finalRtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(
        $Width, $Height, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $finalRtb.Render($dv)
    return $finalRtb
}

function Save-Shot([string]$XamlPath, [string]$FrameName, [int]$Width, [int]$Height, [string]$OutPath) {
    $reader = [System.Xml.XmlReader]::Create($XamlPath)
    $visual = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()

    if ($XamlPath -like "*xaml\search.xaml") {
        Set-FigmaFrameState $visual $FrameName
    }
    Set-RealFont $visual $realFont

    $visual.Measure((New-Object System.Windows.Size($Width, $Height)))
    $visual.Arrange((New-Object System.Windows.Rect(0, 0, $Width, $Height)))
    $visual.UpdateLayout()

    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(
        $Width, $Height, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($visual)

    if ($XamlPath -like "*xaml\search.xaml") {
        $popupEl = $visual.FindName("ContextMenuPopup")
        if ($null -ne $popupEl) { $popupEl.PlacementTarget = $visual.FindName("PreviewOpenArrow") }
        $rtb = Merge-PopupOverlay $visual $rtb $Width $Height
    }
    $rtb.Freeze()

    $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $stream = [System.IO.File]::Open($OutPath, [System.IO.FileMode]::Create)
    try { $encoder.Save($stream) } finally { $stream.Close() }
    Write-Output "撮った: $OutPath"
}

# ---- 既定の大きさで 26 枚すべて ----
foreach ($f in Get-FigmaFrames) {
    $xamlPath = Join-Path $root $f.Xaml
    $outPath = Join-Path $outDir "$($f.Name).png"
    Save-Shot $xamlPath $f.Name $f.Width $f.Height $outPath
}

# ---- リサイズの見本（H・H0・H-row3 を 1024x640・1600x1000 でも） ----
$resizeTargets = @("H", "H0", "H-row3")
$resizeSizes = @(
    @{ Width = 1024; Height = 640 }
    @{ Width = 1600; Height = 1000 }
)
foreach ($name in $resizeTargets) {
    $f = Get-FigmaFrame $name
    if ($null -eq $f) { continue }
    $xamlPath = Join-Path $root $f.Xaml
    foreach ($size in $resizeSizes) {
        $outPath = Join-Path $outDir "$($name)_$($size.Width)x$($size.Height).png"
        Save-Shot $xamlPath $name $size.Width $size.Height $outPath
    }
}

"out\screens\ に書き出した。"
