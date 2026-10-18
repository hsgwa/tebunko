# Excel の流れ読みの計測（measure_xlsx.ps1）で使うブックを、その場で作る（Office は使わない。ZIP を直接組み立てる）。
# テストデータには置かない。中身は架空の文字と数だけ。

function newXlsxPerfBook {
    # 最小のブック（.xlsx）を作る。シートの部品は 1 つで、<sheet> を $sheets 個並べて、すべて同じ部品を指す。
    # シートの部品は、行（sheetData）を $rows 行持ち、その後ろにヘッダーとフッターを持つ（読む側は sheetData を飛ばしてヘッダーを読む）。
    #   path   : 書き出すファイルのパス（あれば上書きする）
    #   sheets : <sheet> の数（1 なら普通のブック。名前は Sheet1 …）
    #   rows   : シートの部品の行の数
    param (
        [string]$path,
        [int]$sheets = 1,
        [int]$rows = 1000
    )

    if ($sheets -lt 1) { throw "sheets は 1 以上にしてください。" }
    if ($rows -lt 0) { throw "rows は 0 以上にしてください。" }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    $nsPkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    if ([System.IO.File]::Exists($path)) { [System.IO.File]::Delete($path) }
    $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $addText = {
            param ([string]$name, [string]$text)
            $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $writer = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
            try { $writer.Write($text) } finally { $writer.Dispose() }
        }
        & $addText "[Content_Types].xml" ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        & $addText "_rels/.rels" ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $nsPkg + '"><Relationship Id="rId1" Type="' + $nsRel + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')

        $sheetTags = New-Object System.Text.StringBuilder
        $relTags = New-Object System.Text.StringBuilder
        for ($i = 1; $i -le $sheets; $i++) {
            [void]$sheetTags.Append('<sheet name="Sheet' + $i + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>')
            [void]$relTags.Append('<Relationship Id="rId' + $i + '" Type="' + $nsRel + '/worksheet" Target="worksheets/sheet1.xml"/>')
        }
        & $addText "xl/workbook.xml" ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + $nsMain + '" xmlns:r="' + $nsRel + '"><sheets>' + $sheetTags.ToString() + '</sheets></workbook>')
        & $addText "xl/_rels/workbook.xml.rels" ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $nsPkg + '">' + $relTags.ToString() + '</Relationships>')

        $entry = $zip.CreateEntry("xl/worksheets/sheet1.xml", [System.IO.Compression.CompressionLevel]::Optimal)
        $writer = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
        try {
            $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="' + $nsMain + '"><sheetData>')
            for ($r = 1; $r -le $rows; $r++) {
                $writer.Write('<row r="' + $r + '"><c r="A' + $r + '" t="inlineStr"><is><t>row ' + $r + ' sample text</t></is></c><c r="B' + $r + '"><v>' + $r + '</v></c></row>')
            }
            $writer.Write('</sheetData><headerFooter><oddHeader>&amp;CPerf Header</oddHeader><oddFooter>&amp;CPerf Footer</oddFooter></headerFooter></worksheet>')
        } finally {
            $writer.Dispose()
        }
    } finally {
        $zip.Dispose()
    }
}
