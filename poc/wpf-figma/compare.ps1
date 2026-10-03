# poc/wpf-figma/compare.ps1
#
# 26 枚のフレームを 1 枚ずつレンダリングし、reference/*.png とピクセル単位で比べる。
# 結果は out/<フレーム名>.diff.png（違うピクセルを赤で塗った画像）と、画面に出す表にまとめる。
#
# 許容差: 色チャンネルごとの差が ±8 以内のピクセルは「同じ」とみなす（アンチエイリアスの端の
# わずかなにじみを誤差として数えないため）。
#
# 参照 PNG の背景は透明（アルファ 0）で書き出されているものがあり（Figma のフレーム書き出しの
# 仕様）、実機の画面は不透明な白地のため、そのまま比べるとキャンバスの余白全体が「違う」と
# 判定されてしまう。比べる前に、参照画像を不透明な白の下地に合成してから色だけを見る
# （アルファそのものの差は見ない）。
#
# 使い方: pwsh -File .\compare.ps1

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$root = $PSScriptRoot
. (Join-Path $root "dummy.ps1")

$outDir = Join-Path $root "out"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$fontsDir = (Join-Path $root "fonts") -replace "\\","/"
$fontUri = "file:///$fontsDir/#Rethink Sans, Yu Gothic UI"
$realFont = New-Object System.Windows.Media.FontFamily($fontUri)

$Tolerance = 8

# Popup（開くメニュー）は RenderTargetBitmap に写らないため、Save-Shot（screenshots.ps1）と
# 同じやり方で Child を別に描いて重ねる（diff_round3.md 3 ★）。
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

function Render-Frame($Frame) {
    $xamlPath = Join-Path $root $Frame.Xaml
    $reader = [System.Xml.XmlReader]::Create($xamlPath)
    $visual = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()

    if ($Frame.Xaml -eq "xaml\search.xaml") {
        Set-FigmaFrameState $visual $Frame.Name
    }
    Set-RealFont $visual $realFont

    $visual.Measure((New-Object System.Windows.Size($Frame.Width, $Frame.Height)))
    $visual.Arrange((New-Object System.Windows.Rect(0, 0, $Frame.Width, $Frame.Height)))
    $visual.UpdateLayout()

    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(
        $Frame.Width, $Frame.Height, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($visual)

    if ($Frame.Xaml -eq "xaml\search.xaml") {
        $popupEl = $visual.FindName("ContextMenuPopup")
        if ($null -ne $popupEl) { $popupEl.PlacementTarget = $visual.FindName("PreviewOpenArrow") }
        $rtb = Merge-PopupOverlay $visual $rtb $Frame.Width $Frame.Height
    }
    $rtb.Freeze()
    return $rtb
}

function Load-ReferenceBitmap([string]$FrameName) {
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

function Get-PixelBytes([System.Windows.Media.Imaging.BitmapSource]$Bitmap) {
    $converted = $Bitmap
    if ($Bitmap.Format -ne [System.Windows.Media.PixelFormats]::Bgra32) {
        $converted = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap($Bitmap, [System.Windows.Media.PixelFormats]::Bgra32, $null, 0)
    }
    $stride = $converted.PixelWidth * 4
    $bytes = New-Object byte[] ($stride * $converted.PixelHeight)
    $converted.CopyPixels($bytes, $stride, 0)
    return @{ Bytes = $bytes; Stride = $stride; Width = $converted.PixelWidth; Height = $converted.PixelHeight }
}

# 1 枚のフレームを比べ、差分情報と diff.png を返す。
function Compare-Frame($Frame) {
    $rendered = Render-Frame $Frame
    $reference = Load-ReferenceBitmap $Frame.Name

    $result = [pscustomobject]@{
        Name       = $Frame.Name
        Width      = $Frame.Width
        Height     = $Frame.Height
        RefFound   = ($null -ne $reference)
        SizeMatch  = $true
        DiffPixels = 0
        TotalPixels = $Frame.Width * $Frame.Height
        DiffRatio  = 0.0
    }

    if ($null -eq $reference) {
        return $result
    }
    if ($reference.PixelWidth -ne $Frame.Width -or $reference.PixelHeight -ne $Frame.Height) {
        $result.SizeMatch = $false
        return $result
    }

    $a = Get-PixelBytes $rendered
    $b = Get-PixelBytes $reference
    $diffBytes = New-Object byte[] $a.Bytes.Length
    [System.Array]::Copy($b.Bytes, $diffBytes, $b.Bytes.Length)

    $diffCount = 0
    for ($i = 0; $i -lt $a.Bytes.Length; $i += 4) {
        # 参照画像の画素を、自分のアルファで不透明な白の下地に合成してから比べる
        # （実機の画面は不透明な白地のため、参照の透明な余白を違いとして数えない）。
        $alpha = [int]$b.Bytes[$i+3]
        $refB = [int]$b.Bytes[$i]   * $alpha / 255 + 255 * (255 - $alpha) / 255
        $refG = [int]$b.Bytes[$i+1] * $alpha / 255 + 255 * (255 - $alpha) / 255
        $refR = [int]$b.Bytes[$i+2] * $alpha / 255 + 255 * (255 - $alpha) / 255

        $db = [int]$a.Bytes[$i]   - $refB
        $dg = [int]$a.Bytes[$i+1] - $refG
        $dr = [int]$a.Bytes[$i+2] - $refR
        if ([Math]::Abs($db) -gt $Tolerance -or [Math]::Abs($dg) -gt $Tolerance -or
            [Math]::Abs($dr) -gt $Tolerance) {
            $diffCount++
            # 違うピクセルは赤で塗る（B,G,R,A）。
            $diffBytes[$i]   = 0
            $diffBytes[$i+1] = 0
            $diffBytes[$i+2] = 255
            $diffBytes[$i+3] = 255
        } else {
            # 合成した参照の色をそのまま出す（透明な余白も白地として見える形にする）。
            $diffBytes[$i]   = [byte]$refB
            $diffBytes[$i+1] = [byte]$refG
            $diffBytes[$i+2] = [byte]$refR
            $diffBytes[$i+3] = 255
        }
    }

    $result.DiffPixels = $diffCount
    $result.DiffRatio = if ($result.TotalPixels -gt 0) { [Math]::Round(100.0 * $diffCount / $result.TotalPixels, 2) } else { 0 }

    $diffBitmap = [System.Windows.Media.Imaging.BitmapSource]::Create(
        $Frame.Width, $Frame.Height, 96, 96, [System.Windows.Media.PixelFormats]::Bgra32, $null,
        $diffBytes, $a.Stride)
    $outPath = Join-Path $outDir "$($Frame.Name).diff.png"
    $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($diffBitmap))
    $stream = [System.IO.File]::Open($outPath, [System.IO.FileMode]::Create)
    try { $encoder.Save($stream) } finally { $stream.Close() }

    return $result
}

$results = @()
foreach ($f in Get-FigmaFrames) {
    $results += Compare-Frame $f
}

$results | ForEach-Object {
    [pscustomobject]@{
        フレーム    = $_.Name
        参照画像   = if ($_.RefFound) { "あり" } else { "なし" }
        サイズ一致 = if ($_.SizeMatch) { "OK" } else { "NG" }
        差分ピクセル = $_.DiffPixels
        差分率     = "$($_.DiffRatio)%"
    }
} | Format-Table -AutoSize

$total = $results.Count
$bad = ($results | Where-Object { -not $_.RefFound -or -not $_.SizeMatch -or $_.DiffRatio -gt 1.0 }).Count
"合計 $total 枚中、差分率 1% を超える・参照がない・サイズが違うもの: $bad 枚"
"diff.png は out\ に保存した（.gitignore 済み）。"
