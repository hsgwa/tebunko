# 画面のアイコン（scripts\tebunko_grep\tebunko_grep.ico）を、元データの SVG（docs\images\logo.svg）から作る。
#
#   .\tools\new_icon.ps1                     docs\images\logo.svg から scripts\tebunko_grep\tebunko_grep.ico を作る
#
# 手順:
#   1. SVG を Microsoft Edge（Windows に入っているもの）のヘッドレスモードで、背景を透明にして 1024 px の PNG に描く
#   2. .NET Framework の System.Drawing で各サイズ（16〜256 px）に縮小する
#   3. 各サイズの PNG をまとめて .ico にする（Windows Vista 以降の PNG 形式のアイコン。WPF の Window.Icon も読める）
#
# 第三者のツール（ImageMagick・Inkscape など）は使わない。途中のファイルは work\icon\ に置く。
param (
    [string]$Svg,
    [string]$OutFile
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$rootDir = Split-Path $PSScriptRoot -Parent
if (!$Svg) {
    $Svg = Join-Path $rootDir "docs\images\logo.svg"
}
if (!$OutFile) {
    $OutFile = Join-Path $rootDir "scripts\tebunko_grep\tebunko_grep.ico"
}
$Svg = (Resolve-Path -LiteralPath $Svg).Path

# .ico に入れるサイズ（タイトルバー・タスクバー・Alt+Tab・エクスプローラーの各表示と高 DPI で使われるもの）
$sizes = @(16, 20, 24, 32, 48, 64, 128, 256)
# 縮小の元にする大きさ
$sourceSize = 1024

$edge = @(
    (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"),
    (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe")
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (!$edge) {
    throw "Microsoft Edge が見つかりません。"
}

$workDir = Join-Path $rootDir "work\icon"
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
$htmlPath = Join-Path $workDir "icon.html"
$pngPath = Join-Path $workDir "icon_$sourceSize.png"
if (Test-Path -LiteralPath $pngPath) {
    Remove-Item -LiteralPath $pngPath -Force
}

# 1. SVG を透明の背景で描く
$svgUri = ([System.Uri]$Svg).AbsoluteUri
$html = "<!DOCTYPE html><html><body style=`"margin:0;background:transparent`"><img src=`"$svgUri`" width=`"$sourceSize`" height=`"$sourceSize`" style=`"display:block`"></body></html>"
[System.IO.File]::WriteAllText($htmlPath, $html, (New-Object System.Text.UTF8Encoding $false))
$edgeArgs = @(
    "--headless=new",
    "--disable-gpu",
    "--hide-scrollbars",
    "--force-device-scale-factor=1",
    "--default-background-color=00000000",
    "--window-size=$sourceSize,$sourceSize",
    "--user-data-dir=`"$(Join-Path $workDir 'edge')`"",
    "--screenshot=`"$pngPath`"",
    "`"$(([System.Uri]$htmlPath).AbsoluteUri)`""
)
Start-Process -FilePath $edge -ArgumentList $edgeArgs -Wait -NoNewWindow -RedirectStandardError (Join-Path $workDir "edge.log")
if (!(Test-Path -LiteralPath $pngPath)) {
    throw "SVG を PNG に描けませんでした（$(Join-Path $workDir 'edge.log')）。"
}

# 2. 各サイズに縮小して PNG のバイト列にする
$source = New-Object System.Drawing.Bitmap($pngPath)
$images = @()
try {
    if ($source.Width -ne $sourceSize -or $source.Height -ne $sourceSize) {
        throw "描いた PNG の大きさが $($source.Width)x$($source.Height) です（$sourceSize x $sourceSize のはず）。"
    }
    if ($source.GetPixel(0, 0).A -ne 0) {
        throw "描いた PNG の背景が透明になっていません。"
    }
    foreach ($size in $sizes) {
        $bitmap = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $stream = New-Object System.IO.MemoryStream
        try {
            $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.DrawImage($source, 0, 0, $size, $size)
            $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
            $images += , @{ Size = $size; Bytes = $stream.ToArray() }
        } finally {
            $stream.Dispose()
            $graphics.Dispose()
            $bitmap.Dispose()
        }
    }
} finally {
    $source.Dispose()
}

# 3. .ico にまとめる（ヘッダー 6 バイト + 画像ごとの項目 16 バイト + 各 PNG）
$out = New-Object System.IO.MemoryStream
$writer = New-Object System.IO.BinaryWriter($out)
try {
    $writer.Write([UInt16]0)                # 予約
    $writer.Write([UInt16]1)                # 種類（1 = アイコン）
    $writer.Write([UInt16]$images.Count)
    $offset = 6 + 16 * $images.Count
    foreach ($image in $images) {
        # 幅・高さは 1 バイトで、256 は 0 と書く
        $side = if ($image.Size -ge 256) { 0 } else { $image.Size }
        $writer.Write([Byte]$side)          # 幅
        $writer.Write([Byte]$side)          # 高さ
        $writer.Write([Byte]0)              # パレットの色数（使わない）
        $writer.Write([Byte]0)              # 予約
        $writer.Write([UInt16]1)            # カラープレーン
        $writer.Write([UInt16]32)           # 1 ピクセルのビット数
        $writer.Write([UInt32]$image.Bytes.Length)
        $writer.Write([UInt32]$offset)
        $offset += $image.Bytes.Length
    }
    foreach ($image in $images) {
        $writer.Write($image.Bytes)
    }
    $writer.Flush()
    [System.IO.File]::WriteAllBytes($OutFile, $out.ToArray())
} finally {
    $writer.Dispose()
    $out.Dispose()
}

Write-Host "アイコン: $OutFile（$($sizes -join '・') px、$((Get-Item -LiteralPath $OutFile).Length) バイト）"
