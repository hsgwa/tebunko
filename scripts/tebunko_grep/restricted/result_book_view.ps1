# 制限モードの検索結果のブック（xlsx）の中身（XML）を組み立てる（判断層。ファイルに触らない）。
#
# 制限言語モードでは Excel の COM も System.IO.Compression も使えないため、xlsx の部品（XML）を文字列で組み立て、
# result_book.ps1 が tar.exe で ZIP にまとめる。制限言語モードで動く書き方だけで書く。
#
# 行の組み立て方（何を 1 行にするか）は result_book.ps1 の newResultRows が決め、ここでは次の形の行を XML にする:
#   @{ Style（行の書式。下の ${resultStyle*}）; Level（グループの深さ 0〜2）; Hidden（畳んで隠す）; Collapsed（下の行を畳んだ）;
#      Texts（セルの文字の配列。A 列から）; Styles（列の番号 → 書式。行の書式と違う列だけ）;
#      Links（列の番号 → @{ Target（ファイル・フォルダのパス）; Location（ブックの中の場所 'シート'!C3。無くてもよい） }）;
#      HighlightFrom（この列の番号から右のセルは、一致した部分に色を付ける。無ければ付けない） }
# セルはすべて文字列（inlineStr）で書き、数式として扱わせない（元の文書の "=SUM(...)" などもそのまま文字として出す）。

# 書式（styles.xml の cellXfs の番号）
${resultStyleNormal}  = 0  # ふつう
${resultStyleHeader}  = 1  # 見出し（太字・背景色）
${resultStyleGroup}   = 2  # ファイルごとのまとまりの行（太字・薄い背景色）
${resultStyleLink}    = 3  # リンク（青・下線）
${resultStyleContext} = 4  # 前後の行（灰色）
${resultStyleGroupLink} = 5  # まとまりの行のリンク（太字・青・下線・薄い背景色）

# セルに入れられる文字数の上限（Excel の仕様）
${resultCellMaxLength} = 32767

# 一致した部分の文字の書式（濃い赤・太字）
${resultHighlightRunProperties} = '<rPr><b/><color rgb="FFC00000"/></rPr>'

function atLeastOne {
    # 1 以上にそろえる（[Math]::Max は制限言語モードで使えないため）
    param ([int]$value)
    if ($value -lt 1) { return 1 }
    return $value
}

function escapeXlsxText {
    # セルの文字を XML に書ける形にする。
    #   ・XML に書けない制御文字（タブ・改行以外）は、Excel と同じ _xHHHH_ の形にする（元から "_x0041_" のような文字列は _x005F_ で守る）
    #   ・& < > " はエンティティにする
    # 上限（32767 文字）を超える文字は切り詰める（Excel は超える文字を持てない）
    param (
        [string]$text
    )

    if ($text.Length -gt ${resultCellMaxLength}) {
        $text = $text.Substring(0, ${resultCellMaxLength})
    }
    $text = $text -creplace '_(x[0-9A-Fa-f]{4}_)', '_x005F_$1'
    if ($text -match '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]') {
        $parts = [regex]::Split($text, '([\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF])')
        for ($i = 1; $i -lt $parts.Count; $i += 2) {
            $parts[$i] = '_x{0:X4}_' -f [int][char]$parts[$i]
        }
        $text = $parts -join ""
    }
    return $text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace('"', "&quot;")
}

function getXlsxTextXml {
    # セルの文字（<is> の中身）。regex に一致した部分だけを別の書式（${resultHighlightRunProperties}）の区切り（<r>）にする
    param (
        [string]$text,
        [regex]$regex = $null
    )

    if ($text.Length -gt ${resultCellMaxLength}) {
        $text = $text.Substring(0, ${resultCellMaxLength})
    }
    $found = @()
    if ($regex -and $text -ne "") {
        $found = @($regex.Matches($text) | Where-Object { $_.Length -gt 0 })
    }
    if ($found.Count -eq 0) {
        return '<t xml:space="preserve">' + (escapeXlsxText $text) + '</t>'
    }
    $runs = @()
    $pos = 0
    foreach ($m in $found) {
        if ($m.Index -gt $pos) {
            $runs += '<r><t xml:space="preserve">' + (escapeXlsxText $text.Substring($pos, $m.Index - $pos)) + '</t></r>'
        }
        $runs += '<r>' + ${resultHighlightRunProperties} + '<t xml:space="preserve">' + (escapeXlsxText $m.Value) + '</t></r>'
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $text.Length) {
        $runs += '<r><t xml:space="preserve">' + (escapeXlsxText $text.Substring($pos)) + '</t></r>'
    }
    return ($runs -join "")
}

function toXlsxLinkTarget {
    # ハイパーリンクの行き先（リレーションの Target）。Excel が書く形に合わせ、ローカル・ネットワークのパスは file:/// に続けて書く。
    # URI として特別な意味を持つ文字（% # 空白）は %XX にする（日本語はそのまま。XML のエスケープは呼び出し側）
    param (
        [string]$path
    )

    $escaped = $path.Replace("%", "%25").Replace("#", "%23").Replace(" ", "%20")
    return "file:///" + $escaped
}

function toXlsxLocation {
    # ブックの中の場所（hyperlink の location）。シート名は ' で囲み、中の ' は重ねる: 'シート 1'!C12
    param (
        [string]$sheet,
        [string]$cell
    )

    return "'" + $sheet.Replace("'", "''") + "'!" + $cell
}

function getResultSheetXml {
    # 検索結果のシート（sheet1.xml）とそのリレーション（sheet1.xml.rels）を @{ Sheet; Rels; LinkCount } で返す。
    #   rows       : 行の配列（このファイルの先頭の形）。1 行目を見出しとし、固定して、オートフィルターを付ける
    #   columnCount: 列の数
    #   widths     : 列の幅（文字数。足りない列は既定の幅）
    #   regex      : 一致した部分に色を付ける正規表現（行の HighlightFrom 列より右のセルに使う）
    # 結果は 1 万件（前後の行を含めて 5 万行）になるため、セルごとに関数を呼ばない（1 回 0.1〜0.3 ms かかる）。
    # 関数を呼ぶのは、一致した部分があるセルと、XML に書けない文字を含むセルだけにする
    param (
        [object[]]$rows,
        [int]$columnCount,
        [double[]]$widths = @(),
        [regex]$regex = $null
    )

    $columns = atLeastOne $columnCount
    $names = @(for ($c = 1; $c -le $columns; $c++) { toColumnName $c })
    $unsafePattern = '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]|_x[0-9A-Fa-f]{4}_'
    $unsafe = New-Object regex -ArgumentList @($unsafePattern, 'CultureInvariant')
    # リンクは番号 → 内容のハッシュテーブルに入れる（配列に += で足すと、件数が多いと遅いため）
    $links = @{}
    $linkIds = @{}   # 行き先 → リレーションの番号（同じファイルへのリンクは 1 つにまとめる）
    $maxLevel = 0
    $rowXml = @(for ($r = 0; $r -lt $rows.Count; $r++) {
            $row = $rows[$r]
            $number = [string]($r + 1)
            $attributes = ' r="' + $number + '"'
            if ($row.Level -gt 0) {
                $attributes += ' outlineLevel="' + $row.Level + '"'
                if ($row.Level -gt $maxLevel) { $maxLevel = $row.Level }
            }
            if ($row.Hidden) { $attributes += ' hidden="1"' }
            if ($row.Collapsed) { $attributes += ' collapsed="1"' }
            $texts = $row.Texts
            $rowStyle = [int]$row.Style
            $styles = $row.Styles
            $rowLinks = $row.Links
            $from = if ($null -ne $row.HighlightFrom) { [int]$row.HighlightFrom } else { -1 }
            if (!$styles -and !$rowLinks -and $from -lt 0) {
                # 書式・リンク・色付けの無い行（前後の行・見出しなど。行の大半）は、セルごとに回さずにまとめて書く。
                # セルの番地（r）は省く（省くと左から順に並ぶ）。空のセルも書いて位置をそろえる
                # コマンドレット（Select-Object・Where-Object）は 1 回ごとに時間がかかるため、演算子で調べる
                $values = [string[]]$texts
                if ($values.Count -eq 0) {
                    '<row' + $attributes + '/>'
                    continue
                }
                if ($values.Count -gt $columns) {
                    $values = $values[0..($columns - 1)]
                }
                if (@($values -match $unsafePattern).Count -gt 0 -or ($values -join "").Length -gt ${resultCellMaxLength}) {
                    $escaped = @(foreach ($value in $values) { escapeXlsxText $value })
                } else {
                    $escaped = $values -replace "&", "&amp;" -replace "<", "&lt;" -replace ">", "&gt;" -replace '"', "&quot;"
                }
                $open = '<c' + $(if ($rowStyle) { ' s="' + $rowStyle + '"' } else { "" }) + ' t="inlineStr"><is><t xml:space="preserve">'
                '<row' + $attributes + '>' + $open + ($escaped -join ('</t></is></c>' + $open)) + '</t></is></c></row>'
                continue
            }
            $cellXml = @(for ($c = 0; $c -lt $texts.Count -and $c -lt $columns; $c++) {
                    $text = [string]$texts[$c]
                    $style = $rowStyle
                    if ($styles -and $styles.ContainsKey($c)) { $style = [int]$styles[$c] }
                    $link = $null
                    if ($rowLinks -and $rowLinks.ContainsKey($c)) { $link = $rowLinks[$c] }
                    # 書式の無い空のセルは書かない
                    if ($text -eq "" -and $style -eq 0 -and !$link) { continue }
                    $ref = $names[$c] + $number
                    if ($link) {
                        $target = toXlsxLinkTarget $link.Target
                        if (!$linkIds.ContainsKey($target)) {
                            $linkIds[$target] = "rId" + ($linkIds.Count + 1)
                        }
                        $links[$links.Count] = @{ Ref = $ref; Id = $linkIds[$target]; Location = $link.Location }
                    }
                    if ($regex -and $from -ge 0 -and $c -ge $from -and $text -ne "" -and $regex.IsMatch($text)) {
                        $inner = getXlsxTextXml $text $regex
                    } elseif ($text.Length -gt ${resultCellMaxLength} -or $unsafe.IsMatch($text)) {
                        $inner = '<t xml:space="preserve">' + (escapeXlsxText $text) + '</t>'
                    } else {
                        $inner = '<t xml:space="preserve">' + $text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace('"', "&quot;") + '</t>'
                    }
                    $s = if ($style) { ' s="' + $style + '"' } else { "" }
                    '<c r="' + $ref + '"' + $s + ' t="inlineStr"><is>' + $inner + '</is></c>'
                })
            '<row' + $attributes + '>' + ($cellXml -join "") + '</row>'
        })

    $lastColumn = $names[$columns - 1]
    $lastRow = atLeastOne $rows.Count
    $cols = @(for ($c = 1; $c -le $columns; $c++) {
            $width = if ($c -le $widths.Count) { $widths[$c - 1] } else { 20 }
            '<col min="' + $c + '" max="' + $c + '" width="' + $width + '" customWidth="1"/>'
        })

    $sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetPr><outlinePr summaryBelow="0"/></sheetPr>' +
        '<dimension ref="A1:' + $lastColumn + $lastRow + '"/>' +
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="18.75"' + $(if ($maxLevel -gt 0) { ' outlineLevelRow="' + $maxLevel + '"' } else { "" }) + '/>' +
        '<cols>' + ($cols -join "") + '</cols>' +
        '<sheetData>' + ($rowXml -join "") + '</sheetData>' +
        '<autoFilter ref="A1:' + $lastColumn + $lastRow + '"/>'
    if ($links.Count -gt 0) {
        $linkXml = @(for ($i = 0; $i -lt $links.Count; $i++) {
                $item = $links[$i]
                $location = if ($item.Location) { ' location="' + (escapeXlsxText $item.Location) + '"' } else { "" }
                '<hyperlink ref="' + $item.Ref + '" r:id="' + $item.Id + '"' + $location + '/>'
            })
        $sheet += '<hyperlinks>' + ($linkXml -join "") + '</hyperlinks>'
    }
    $sheet += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'

    $relXml = @(foreach ($target in $linkIds.Keys) {
            '<Relationship Id="' + $linkIds[$target] + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + (escapeXlsxText $target) + '" TargetMode="External"/>'
        })
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ($relXml -join "") + '</Relationships>'
    return @{ Sheet = $sheet; Rels = $rels; LinkCount = $links.Count }
}

function getResultInfoSheetXml {
    # 検索の条件のシート（sheet2.xml）。items は @( @("項目", "値"), ... )
    param (
        [object[]]$items
    )

    $rows = @(@{ Style = ${resultStyleHeader}; Texts = @("項目", "内容") })
    foreach ($item in $items) {
        $rows += @{ Style = ${resultStyleNormal}; Texts = @([string]$item[0], [string]$item[1]) }
    }
    $rowXml = @(for ($r = 0; $r -lt $rows.Count; $r++) {
            $cells = @(for ($c = 0; $c -lt $rows[$r].Texts.Count; $c++) {
                    $s = if ($rows[$r].Style) { ' s="' + $rows[$r].Style + '"' } else { "" }
                    '<c r="' + (toColumnName ($c + 1)) + ($r + 1) + '"' + $s + ' t="inlineStr"><is>' + (getXlsxTextXml ([string]$rows[$r].Texts[$c])) + '</is></c>'
                })
            '<row r="' + ($r + 1) + '">' + ($cells -join "") + '</row>'
        })
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
        '<cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="80" customWidth="1"/></cols>' +
        '<sheetData>' + ($rowXml -join "") + '</sheetData>' +
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
}

function getResultBookParts {
    # xlsx の部品（ZIP の中のパス → 中身）を返す。sheet・rels・info は getResultSheetXml・getResultInfoSheetXml の結果。
    #   resultRowCount: 検索結果のシートの行数（オートフィルターの範囲の名前に使う）
    param (
        [hashtable]$sheet,
        [string]$info,
        [int]$rowCount,
        [int]$columnCount
    )

    $lastColumn = toColumnName ((atLeastOne $columnCount))
    return [ordered]@{
        "[Content_Types].xml"        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
            '</Types>'
        "_rels/.rels"                = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>'
        "xl/workbook.xml"            = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<bookViews><workbookView/></bookViews>' +
            '<sheets><sheet name="検索結果" sheetId="1" r:id="rId1"/><sheet name="条件" sheetId="2" r:id="rId2"/></sheets>' +
            '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">検索結果!$A$1:$' + $lastColumn + '$' + (atLeastOne $rowCount) + '</definedName></definedNames>' +
            '</workbook>'
        "xl/_rels/workbook.xml.rels" = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
            '</Relationships>'
        "xl/styles.xml"              = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<fonts count="5">' +
            '<font><sz val="11"/><name val="Yu Gothic"/><family val="3"/><charset val="128"/></font>' +
            '<font><b/><sz val="11"/><name val="Yu Gothic"/><family val="3"/><charset val="128"/></font>' +
            '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Yu Gothic"/><family val="3"/><charset val="128"/></font>' +
            '<font><sz val="11"/><color rgb="FF808080"/><name val="Yu Gothic"/><family val="3"/><charset val="128"/></font>' +
            '<font><b/><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Yu Gothic"/><family val="3"/><charset val="128"/></font>' +
            '</fonts>' +
            '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill></fills>' +
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="6">' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '</cellXfs>' +
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
            '</styleSheet>'
        "xl/worksheets/sheet1.xml"   = $sheet.Sheet
        "xl/worksheets/_rels/sheet1.xml.rels" = $sheet.Rels
        "xl/worksheets/sheet2.xml"   = $info
    }
}
