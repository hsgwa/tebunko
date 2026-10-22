# 最小の .xlsx（Office Open XML）を、行ごとに書き出す。Excel は使わない。
# ZIP（System.IO.Compression）に部品を直接書く。シートの行は 1 行ずつ ZIP のエントリーへ流し、全件を文字列に溜めない。
# ツールを知らない部品: 見出しと行（文字列の配列）を受け取るだけで、列の意味は知らない。
#   newXlsxWriter   … 出力先を開き、書き手（状態）を返す
#   addXlsxRow      … 1 行足す（行の上限を超える分は書かず $false）
#   closeXlsxWriter … 完成させて閉じる（書いた行数・切り詰めたセルの数・落とした行数を返す）
#   abortXlsxWriter … 完成させずに閉じる（失敗・中止のとき。ファイルを消すのは呼ぶ側）
#   limitXlsxCellText … セルの文字を XML に入れられる形にし、セルの文字数の上限に収める（純粋な関数）
# 全セルを文字（inlineStr）で書くため、1-2・007・=… も文字のままになる。数式（<f>）・共有文字列・マクロ・外部リンクは入れない。
# 共通の部品（shared.ps1）を先に読み込んでおくこと。

Add-Type -AssemblyName System.IO.Compression

$script:XlsxSheetName = "検索結果"
$script:XlsxMaxCellLength = 32767
$script:XlsxNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
$script:XlsxXmlHead = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n"

function limitXlsxCellText {
    # XML に入れられない文字をスペースにし（CR は LF にそろえる）、32,767 文字を超えるなら 32,766 文字で切って末尾に … を付ける
    param (
        [string]$text
    )

    if (!$text) {
        return @{ Text = ""; Truncated = $false }
    }
    $text = [regex]::Replace($text, "\r\n?", "`n")
    $text = [regex]::Replace($text, "[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]", " ")
    if ($text.Length -le $script:XlsxMaxCellLength) {
        return @{ Text = $text; Truncated = $false }
    }
    $cut = $script:XlsxMaxCellLength - 1
    if ([char]::IsHighSurrogate($text[$cut - 1])) {
        $cut--
    }
    return @{ Text = $text.Substring(0, $cut) + [string][char]0x2026; Truncated = $true }
}

function getXlsxColumnName {
    # 1 始まりの列番号を A, B, ..., Z, AA の形にする
    param (
        [int]$index
    )

    $name = ""
    while ($index -gt 0) {
        $index--
        $name = [string][char](65 + ($index % 26)) + $name
        $index = [int][math]::Floor($index / 26)
    }
    return $name
}

function escapeXlsxText {
    param (
        [string]$text
    )

    return $text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
}

function addXlsxEntry {
    # ZIP に部品を 1 つ書く（小さい部品用。エントリー名の区切りは /）
    param (
        $zip,
        [string]$name,
        [string]$content
    )

    $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Fastest)
    $stream = $entry.Open()
    try {
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($content)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
    }
}

function writeXlsxRowXml {
    # 1 行分の XML を書く。見出しの行は見出しのスタイル、改行を含むセルには折り返しのスタイルを当てる
    param (
        $writer,
        [string[]]$cells,
        [int]$rowNumber,
        [bool]$isHeader
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<row r="').Append($rowNumber).Append('">')
    for ($i = 0; $i -lt $writer.Headers.Count; $i++) {
        $limited = limitXlsxCellText $(if ($i -lt $cells.Count) { $cells[$i] } else { "" })
        if (!$isHeader -and $limited.Truncated) {
            $writer.TruncatedCells++
        }
        $style = 0
        if ($isHeader) {
            $style = 1
        } elseif ($limited.Text.Contains("`n")) {
            $style = 2
        }
        [void]$sb.Append('<c r="').Append($writer.Columns[$i]).Append($rowNumber).Append('" t="inlineStr"')
        if ($style) {
            [void]$sb.Append(' s="').Append($style).Append('"')
        }
        [void]$sb.Append('><is><t xml:space="preserve">').Append((escapeXlsxText $limited.Text)).Append('</t></is></c>')
    }
    [void]$sb.Append('</row>')
    $writer.Sheet.Write($sb.ToString())
}

function newXlsxWriter {
    # 出力先を開く（FileMode.Create・FileShare.None）。開かれていれば IOException、読み取り専用なら UnauthorizedAccessException をそのまま投げる。
    # maxRows は見出しを含む行の上限
    param (
        [string]$path,
        [string[]]$headers,
        [int]$maxRows = 1048576
    )

    try {
        $file = New-Object System.IO.FileStream((toLongPath $path), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    } catch {
        # New-Object が包む MethodInvocationException ではなく、元の IOException・UnauthorizedAccessException を投げる
        if ($_.Exception.InnerException) { throw $_.Exception.InnerException }
        throw
    }
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($file, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        $entry = $zip.CreateEntry("xl/worksheets/sheet1.xml", [System.IO.Compression.CompressionLevel]::Fastest)
        $sheet = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        $file.Dispose()
        throw
    }

    $columns = @(for ($i = 1; $i -le $headers.Count; $i++) { getXlsxColumnName $i })
    $writer = @{
        File           = $file
        Zip            = $zip
        Sheet          = $sheet
        Headers        = $headers
        Columns        = $columns
        MaxRows        = $maxRows
        RowCount       = 0
        TruncatedCells = 0
        DroppedRows    = 0
    }

    $sheet.Write($script:XlsxXmlHead)
    $sheet.Write('<worksheet xmlns="' + $script:XlsxNs + '">')
    $sheet.Write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
    $sheet.Write('<cols>')
    for ($i = 1; $i -le $headers.Count; $i++) {
        $width = if ($i -eq $headers.Count) { 80 } else { 20 }
        $sheet.Write('<col min="' + $i + '" max="' + $i + '" width="' + $width + '" customWidth="1"/>')
    }
    $sheet.Write('</cols><sheetData>')
    writeXlsxRowXml $writer $headers 1 $true
    $writer.RowCount = 1
    return $writer
}

function addXlsxRow {
    # 1 行足す。書けたら $true。行の上限を超える分は書かず $false（DroppedRows に数える）
    param (
        $writer,
        [string[]]$cells
    )

    if ($writer.RowCount -ge $writer.MaxRows) {
        $writer.DroppedRows++
        return $false
    }
    $writer.RowCount++
    writeXlsxRowXml $writer $cells $writer.RowCount $false
    return $true
}

function closeXlsxWriter {
    # シートを閉じ、残りの部品を書いて完成させる。ファイルは閉じる
    param (
        $writer
    )

    $lastRow = $writer.RowCount
    $lastColumn = $writer.Columns[$writer.Columns.Count - 1]
    try {
        $writer.Sheet.Write('</sheetData><autoFilter ref="A1:' + $lastColumn + $lastRow + '"/></worksheet>')
        $writer.Sheet.Dispose()

        $ns = $script:XlsxNs
        $rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        $head = $script:XlsxXmlHead
        $name = $script:XlsxSheetName
        addXlsxEntry $writer.Zip "[Content_Types].xml" ($head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
        addXlsxEntry $writer.Zip "_rels/.rels" ($head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + $rel + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        addXlsxEntry $writer.Zip "xl/workbook.xml" ($head + '<workbook xmlns="' + $ns + '" xmlns:r="' + $rel + '"><sheets><sheet name="' + $name + '" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">''' + $name + '''!$A$1:$' + $lastColumn + '$' + $lastRow + '</definedName></definedNames></workbook>')
        addXlsxEntry $writer.Zip "xl/_rels/workbook.xml.rels" ($head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + $rel + '/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="' + $rel + '/styles" Target="styles.xml"/></Relationships>')
        addXlsxEntry $writer.Zip "xl/styles.xml" ($head + '<styleSheet xmlns="' + $ns + '"><fonts count="2"><font><sz val="11"/><name val="Meiryo UI"/></font><font><b/><sz val="11"/><name val="Meiryo UI"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    } finally {
        $writer.Zip.Dispose()
        $writer.File.Dispose()
    }
    return @{
        Count          = $lastRow - 1
        TruncatedCells = $writer.TruncatedCells
        DroppedRows    = $writer.DroppedRows
    }
}

function abortXlsxWriter {
    # 完成させずに閉じる（失敗・中止のとき。ファイルを消すのは呼ぶ側）。閉じる途中の例外は握りつぶす
    param (
        $writer
    )

    foreach ($target in @($writer.Sheet, $writer.Zip, $writer.File)) {
        try {
            $target.Dispose()
        } catch {
            $null = $_
        }
    }
}
