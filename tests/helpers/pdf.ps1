# テスト用の PDF を組み立てる（外から PDF を持ってこない。作成者などのメタデータは入れない）。
#   . "$PSScriptRoot\..\helpers\pdf.ps1"
# 文字列はすべて Latin-1 の 1 バイト = 1 文字として扱う。

$script:pdfTestLatin1 = [System.Text.Encoding]::GetEncoding(28591)

function toPdfBytes([string]$text) {
    return , ($script:pdfTestLatin1.GetBytes($text))
}

function joinPdfBytes([object[]]$parts) {
    $ms = New-Object System.IO.MemoryStream
    foreach ($p in $parts) {
        if ($null -eq $p) { continue }
        if ($p -is [byte[]]) { $ms.Write($p, 0, $p.Length) }
        else { $b = $script:pdfTestLatin1.GetBytes([string]$p); $ms.Write($b, 0, $b.Length) }
    }
    return , ($ms.ToArray())
}

function compressPdfBytes([byte[]]$bytes) {
    # zlib 形式（2 バイトの先頭 + deflate + Adler-32）
    $ms = New-Object System.IO.MemoryStream
    $ms.WriteByte(0x78); $ms.WriteByte(0x9C)
    $ds = New-Object System.IO.Compression.DeflateStream($ms, [System.IO.Compression.CompressionMode]::Compress, $true)
    $ds.Write($bytes, 0, $bytes.Length)
    $ds.Dispose()
    [uint32]$a = 1; [uint32]$b = 0
    foreach ($x in $bytes) {
        $a = ($a + $x) % 65521
        $b = ($b + $a) % 65521
    }
    $adler = ([uint64]$b -shl 16) -bor $a
    foreach ($shift in 24, 16, 8, 0) { $ms.WriteByte([byte](($adler -shr $shift) -band 0xFF)) }
    return , ($ms.ToArray())
}

function newPdfStreamObject([string]$dictBody, [byte[]]$data) {
    # /Length を足した辞書と本体
    return @{ Dict = "<< $dictBody /Length $($data.Length) >>"; Stream = $data }
}

function newPdfFromObjects {
    # objects: 番号 -> 文字列（辞書など）か、newPdfStreamObject の戻り値。xref は表形式
    param (
        [System.Collections.IDictionary]$objects,
        [string]$trailerExtra = '/Root 1 0 R',
        [string]$header = "%PDF-1.4`n"
    )

    $ms = New-Object System.IO.MemoryStream
    $h = toPdfBytes $header
    $ms.Write($h, 0, $h.Length)
    $offsets = @{}
    $maxId = ($objects.Keys | Measure-Object -Maximum).Maximum
    foreach ($id in ($objects.Keys | Sort-Object)) {
        $offsets[[int]$id] = [int]$ms.Length
        $o = $objects[$id]
        if ($o -is [string]) {
            $b = joinPdfBytes @("$id 0 obj`n", $o, "`nendobj`n")
        }
        else {
            $b = joinPdfBytes @("$id 0 obj`n", $o.Dict, "`nstream`n", $o.Stream, "`nendstream`nendobj`n")
        }
        $ms.Write($b, 0, $b.Length)
    }
    $xrefPos = [int]$ms.Length
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("xref`n0 $($maxId + 1)`n")
    [void]$sb.Append("0000000000 65535 f `n")
    for ($i = 1; $i -le $maxId; $i++) {
        if ($offsets.ContainsKey($i)) { [void]$sb.Append(('{0:D10} 00000 n ' -f $offsets[$i]) + "`n") }
        else { [void]$sb.Append("0000000000 00000 f `n") }
    }
    [void]$sb.Append("trailer`n<< /Size $($maxId + 1) $trailerExtra >>`nstartxref`n$xrefPos`n%%EOF`n")
    $t = toPdfBytes $sb.ToString()
    $ms.Write($t, 0, $t.Length)
    return , ($ms.ToArray())
}

function newSimplePdf {
    # pages: 各ページの内容ストリーム（文字列）。フォント F1 は Helvetica（WinAnsi）。compress なら Flate にする
    param (
        [string[]]$pages,
        [switch]$compress,
        [string]$fontObject = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>',
        [System.Collections.IDictionary]$extraObjects = $null
    )

    $objs = @{}
    $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
    $kids = (0..($pages.Count - 1) | ForEach-Object { "$(10 + $_ * 2) 0 R" }) -join ' '
    $objs[2] = "<< /Type /Pages /Kids [$kids] /Count $($pages.Count) >>"
    $objs[3] = $fontObject
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $pageId = 10 + $i * 2
        $contentId = $pageId + 1
        $objs[$pageId] = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents $contentId 0 R >>"
        $data = toPdfBytes $pages[$i]
        if ($compress) {
            $objs[$contentId] = newPdfStreamObject '/Filter /FlateDecode' (compressPdfBytes $data)
        }
        else {
            $objs[$contentId] = newPdfStreamObject '' $data
        }
    }
    if ($null -ne $extraObjects) { foreach ($k in $extraObjects.Keys) { $objs[[int]$k] = $extraObjects[$k] } }
    return , (newPdfFromObjects $objs)
}

function newPdfWithXrefStream {
    # normal: 番号 -> 文字列か newPdfStreamObject（そのまま書く）。packed: 番号 -> 辞書の文字列（オブジェクトのストリームにまとめる）
    # 相互参照は xref ストリーム（Flate。predictor なら PNG の Up を使う）
    param (
        [System.Collections.IDictionary]$normal,
        [System.Collections.IDictionary]$packed = @{},
        [switch]$predictor,
        [string]$trailerExtra = '/Root 1 0 R'
    )

    $objs = @{}
    foreach ($k in $normal.Keys) { $objs[[int]$k] = $normal[$k] }
    $maxId = ($objs.Keys + $packed.Keys | Measure-Object -Maximum).Maximum
    $objStmId = $maxId + 1
    $xrefId = $maxId + 2
    $packedIds = @($packed.Keys | ForEach-Object { [int]$_ } | Sort-Object)
    if ($packedIds.Count -gt 0) {
        $head = New-Object System.Text.StringBuilder
        $body = New-Object System.Text.StringBuilder
        foreach ($id in $packedIds) {
            [void]$head.Append("$id $($body.Length) ")
            [void]$body.Append($packed[$id] + "`n")
        }
        $data = toPdfBytes ($head.ToString() + $body.ToString())
        $objs[$objStmId] = newPdfStreamObject "/Type /ObjStm /N $($packedIds.Count) /First $($head.Length) /Filter /FlateDecode" (compressPdfBytes $data)
    }

    $ms = New-Object System.IO.MemoryStream
    $h = toPdfBytes "%PDF-1.5`n"
    $ms.Write($h, 0, $h.Length)
    $offsets = @{}
    foreach ($id in ($objs.Keys | Sort-Object)) {
        $offsets[[int]$id] = [int]$ms.Length
        $o = $objs[$id]
        if ($o -is [string]) { $b = joinPdfBytes @("$id 0 obj`n", $o, "`nendobj`n") }
        else { $b = joinPdfBytes @("$id 0 obj`n", $o.Dict, "`nstream`n", $o.Stream, "`nendstream`nendobj`n") }
        $ms.Write($b, 0, $b.Length)
    }
    $xrefPos = [int]$ms.Length
    $offsets[$xrefId] = $xrefPos

    $rows = New-Object System.Collections.ArrayList
    for ($i = 0; $i -le $xrefId; $i++) {
        $idx = [array]::IndexOf($packedIds, $i)
        if ($idx -ge 0) { [void]$rows.Add([byte[]]@(2, 0, 0, 0, 0, 0, 0)); $r = $rows[$rows.Count - 1]; $r[1] = ($objStmId -shr 24) -band 255; $r[2] = ($objStmId -shr 16) -band 255; $r[3] = ($objStmId -shr 8) -band 255; $r[4] = $objStmId -band 255; $r[5] = ($idx -shr 8) -band 255; $r[6] = $idx -band 255 }
        elseif ($offsets.ContainsKey($i)) { $o = $offsets[$i]; [void]$rows.Add([byte[]]@(1, (($o -shr 24) -band 255), (($o -shr 16) -band 255), (($o -shr 8) -band 255), ($o -band 255), 0, 0)) }
        else { [void]$rows.Add([byte[]]@(0, 0, 0, 0, 0, 0, 0)) }
    }
    $raw = New-Object System.IO.MemoryStream
    $prev = New-Object byte[] 7
    foreach ($row in $rows) {
        if ($predictor) {
            $raw.WriteByte(2)
            $d = New-Object byte[] 7
            for ($j = 0; $j -lt 7; $j++) { $d[$j] = [byte](($row[$j] - $prev[$j] + 256) % 256) }
            $raw.Write($d, 0, 7)
            $prev = $row
        }
        else { $raw.Write($row, 0, 7) }
    }
    $parms = if ($predictor) { ' /DecodeParms << /Predictor 12 /Columns 7 >>' } else { '' }
    $xs = newPdfStreamObject "/Type /XRef /W [1 4 2] /Size $($xrefId + 1) /Filter /FlateDecode$parms $trailerExtra" (compressPdfBytes $raw.ToArray())
    $b = joinPdfBytes @("$xrefId 0 obj`n", $xs.Dict, "`nstream`n", $xs.Stream, "`nendstream`nendobj`n", "startxref`n$xrefPos`n%%EOF`n")
    $ms.Write($b, 0, $b.Length)
    return , ($ms.ToArray())
}
