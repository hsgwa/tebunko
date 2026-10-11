# PDF の読み取りの速さを測る PDF を作る（tools\perf\measure_pdf.ps1 が使う）。外から PDF を持ってこず、テストの組み立て（tests\helpers\pdf.ps1）で作る。
#
#   .\tools\perf\new_pdf_data.ps1 -Dest <作る PDF のパス> -Scale 1
#   -Scale … 量の倍率（既定 1。1 で数 MB 規模。小さくすると速く作れる）
#
# 中身（同じ引数からは、同じ PDF ができる）:
#   - 文字の多いページ（TJ・Tj・T* の行が並ぶ。フォントは WinAnsi の Type1）
#   - 図形の演算子の多いページ（色・パス・行列の演算子、埋め込み画像、少しの文字）
#   - 大きな ToUnicode の表を持つ Type0 フォントのページ（表の違うフォントを 2 つ。16 進の文字列）
#   - 相互参照のストリーム（Predictor 付き）と、オブジェクトのストリームにまとめた多数の小さなオブジェクト
param (
    [Parameter(Mandatory = $true)][string]$Dest,
    [double]$Scale = 1.0
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\..\..\tests\helpers\pdf.ps1"

$inv = [System.Globalization.CultureInfo]::InvariantCulture
$rng = New-Object System.Random(20240601)

function newWord([int]$min, [int]$max) {
    $len = $rng.Next($min, $max + 1)
    $chars = New-Object char[] $len
    for ($i = 0; $i -lt $len; $i++) { $chars[$i] = [char]$rng.Next(97, 123) }
    return (New-Object string (, $chars))
}

function formatNumber([double]$v) { return $v.ToString("0.##", $inv) }

function compressForPdf([byte[]]$bytes) {
    # zlib 形式。末尾の検査値は読む側が見ないので 0 にする（作るのを速くするため）
    $ms = New-Object System.IO.MemoryStream
    $ms.WriteByte(0x78); $ms.WriteByte(0x9C)
    $ds = New-Object System.IO.Compression.DeflateStream($ms, [System.IO.Compression.CompressionMode]::Compress, $true)
    $ds.Write($bytes, 0, $bytes.Length)
    $ds.Dispose()
    for ($i = 0; $i -lt 4; $i++) { $ms.WriteByte(0) }
    return , ($ms.ToArray())
}

function newTextPage([int]$lines) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("BT /F1 10 Tf 12 TL 1 0 0 1 72 760 Tm`n")
    for ($i = 0; $i -lt $lines; $i++) {
        switch ($i % 4) {
            0 { [void]$sb.Append("[($(newWord 3 9)) -300 ($(newWord 3 9)) 20 ($(newWord 3 9)) -280 ($(newWord 2 8) \($(newWord 2 5)\))] TJ`n") }
            1 { [void]$sb.Append("($(newWord 3 9) $(newWord 3 9) $(newWord 3 9) $(newWord 3 9)) Tj T*`n") }
            2 { [void]$sb.Append("0 -14 Td ($(newWord 4 10) $(newWord 4 10)) Tj 40 0 Td ($(newWord 3 8)) Tj`n") }
            3 { [void]$sb.Append("1 0 0 1 72 $(formatNumber (740 - $i * 5.5)) Tm [(c) 10 (af\351) -400 ($(newWord 3 9))] TJ`n") }
        }
    }
    [void]$sb.Append("ET`n")
    return $sb.ToString()
}

function newGraphicsPage([int]$groups) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $groups; $i++) {
        $x = formatNumber ($rng.NextDouble() * 500)
        $y = formatNumber ($rng.NextDouble() * 700)
        $w = formatNumber ($rng.NextDouble() * 100)
        $r = formatNumber $rng.NextDouble()
        $g = formatNumber $rng.NextDouble()
        $b = formatNumber $rng.NextDouble()
        [void]$sb.Append("q $r $g $b rg $x $y $w $w re f Q`n")
        [void]$sb.Append("$r $g $b RG 0.5 w $x $y m $w $y l $w $w l h S`n")
        [void]$sb.Append("1 0 0 1 $x $y cm 0 0 m 10 20 30 40 50 60 c 70 80 l S`n")
        if ($i % 40 -eq 0) {
            [void]$sb.Append("BT /F1 9 Tf 1 0 0 1 $x $y Tm ($(newWord 3 9) $(newWord 3 9)) Tj ET`n")
        }
        if ($i % 100 -eq 0) {
            $image = New-Object byte[] 64
            $rng.NextBytes($image)
            [void]$sb.Append("q 20 0 0 20 $x $y cm BI /W 8 /H 8 /BPC 8 /CS /G ID ")
            [void]$sb.Append($script:pdfTestLatin1.GetString($image).Replace("E", "e"))
            [void]$sb.Append(" EI Q`n")
        }
    }
    return $sb.ToString()
}

function newCidPage([int]$lines, [string]$font, [int]$codeMax) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("BT /$font 10 Tf 12 TL 1 0 0 1 72 760 Tm`n")
    for ($i = 0; $i -lt $lines; $i++) {
        $hex = New-Object System.Text.StringBuilder
        for ($k = 0; $k -lt 24; $k++) { [void]$hex.Append(("{0:X4}" -f $rng.Next(1, $codeMax))) }
        if ($i % 2 -eq 0) { [void]$sb.Append("<$($hex.ToString())> Tj T*`n") }
        else { [void]$sb.Append("[<$($hex.ToString(0, 48))> -300 <$($hex.ToString(48, 48))>] TJ T*`n") }
    }
    [void]$sb.Append("ET`n")
    return $sb.ToString()
}

function newCmap([int]$chars, [int]$ranges) {
    # コード 1 .. chars を bfchar で、その後ろを bfrange で、ひらがな・漢字の範囲に対応させる
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("/CIDInit /ProcSet findresource begin 12 dict begin begincmap`n/CMapName /Adobe-Identity-UCS def /CMapType 2 def`n1 begincodespacerange <0000> <FFFF> endcodespacerange`n")
    $code = 1
    while ($code -le $chars) {
        $n = [Math]::Min(100, $chars - $code + 1)
        [void]$sb.Append("$n beginbfchar`n")
        for ($i = 0; $i -lt $n; $i++) {
            $target = 0x3041 + $rng.Next(0, 6000)
            [void]$sb.Append(("<{0:X4}> <{1:X4}>`n" -f $code, $target))
            $code++
        }
        [void]$sb.Append("endbfchar`n")
    }
    $left = $ranges
    while ($left -gt 0) {
        $n = [Math]::Min(100, $left)
        [void]$sb.Append("$n beginbfrange`n")
        for ($i = 0; $i -lt $n; $i++) {
            $lo = $code
            $hi = $code + 15
            [void]$sb.Append(("<{0:X4}> <{1:X4}> <{2:X4}>`n" -f $lo, $hi, (0x4E00 + $rng.Next(0, 4000))))
            $code = $hi + 1
        }
        [void]$sb.Append("endbfrange`n")
        $left -= $n
    }
    [void]$sb.Append("endcmap end end`n")
    return $sb.ToString()
}

$textPages = [Math]::Max(1, [int][Math]::Round(60 * $Scale))
$graphicsPages = [Math]::Max(1, [int][Math]::Round(30 * $Scale))
$cidPages = [Math]::Max(1, [int][Math]::Round(30 * $Scale))
$fillers = [Math]::Max(10, [int][Math]::Round(1500 * $Scale))
$cmapChars = [Math]::Max(100, [int][Math]::Round(8000 * $Scale))
$cmapRanges = [Math]::Max(10, [int][Math]::Round(400 * $Scale))
$pageCount = $textPages + $graphicsPages + $cidPages
$codeMax = $cmapChars + 1

$normal = @{}
$packed = @{}
$packed[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>'
$packed[4] = '<< /Type /Font /Subtype /Type0 /BaseFont /Test /Encoding /Identity-H /DescendantFonts [5 0 R] /ToUnicode 6 0 R >>'
$packed[5] = '<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Test /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> >>'
$normal[6] = newPdfStreamObject '/Filter /FlateDecode' (compressForPdf (toPdfBytes (newCmap $cmapChars $cmapRanges)))
$packed[7] = '<< /Type /Font /Subtype /Type0 /BaseFont /Test2 /Encoding /Identity-H /DescendantFonts [5 0 R] /ToUnicode 8 0 R >>'
$normal[8] = newPdfStreamObject '/Filter /FlateDecode' (compressForPdf (toPdfBytes (newCmap $cmapChars $cmapRanges)))

$kids = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $pageCount; $i++) {
    $pageId = 10 + $i * 2
    $contentId = $pageId + 1
    if ($i -lt $textPages) { $content = newTextPage 120 }
    elseif ($i -lt $textPages + $graphicsPages) { $content = newGraphicsPage 300 }
    else { $content = newCidPage 80 $(if ($i % 2 -eq 0) { "F2" } else { "F3" }) $codeMax }
    $packed[$pageId] = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 7 0 R >> >> /Contents $contentId 0 R >>"
    $normal[$contentId] = newPdfStreamObject '/Filter /FlateDecode' (compressForPdf (toPdfBytes $content))
    $kids.Add("$pageId 0 R")
}
$packed[1] = '<< /Type /Catalog /Pages 2 0 R >>'
$packed[2] = "<< /Type /Pages /Kids [$($kids -join ' ')] /Count $pageCount >>"
$fillerBase = 10 + $pageCount * 2 + 10
for ($i = 0; $i -lt $fillers; $i++) {
    $packed[$fillerBase + $i] = "<< /Type /Annot /Subtype /Link /Rect [$($rng.Next(0, 500)) $($rng.Next(0, 700)) $($rng.Next(500, 600)) $($rng.Next(700, 790))] /Border [0 0 0] /A << /S /URI /URI (http://example.com/$(newWord 5 12)) >> >>"
}

$bytes = newPdfWithXrefStream -normal $normal -packed $packed -predictor
$full = [System.IO.Path]::GetFullPath($Dest)
[void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($full))
[System.IO.File]::WriteAllBytes($full, $bytes)
