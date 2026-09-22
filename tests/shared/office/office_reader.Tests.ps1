# Pester 3.4 以降で実行: Invoke-Pester .\tests
# Word・PowerPointは使わず、最小限の .docx / .pptx（ZIP）をテスト内で作成して検証する
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\office\office_reader.ps1"

function newZip {
    # ZIP内のパス → 内容 の辞書からZIPファイルを作成する
    param (
        [string]$path,
        [hashtable]$entries
    )

    $stream = [System.IO.File]::Create($path)
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $entries.Keys) {
            $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open(), (New-Object System.Text.UTF8Encoding($false)))
            $writer.Write($entries[$name])
            $writer.Dispose()
        }
    } finally {
        $zip.Dispose()
        $stream.Dispose()
    }
}

$wNs = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
$pNs = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$relNs = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'

function wPara([string]$text) {
    return "<w:p><w:r><w:t xml:space=`"preserve`">${text}</w:t></w:r></w:p>"
}

function pShape([string]$text, [string]$placeholder = "") {
    $ph = $(if ($placeholder) { "<p:ph type=`"${placeholder}`"/>" } else { "" })
    return "<p:sp><p:nvSpPr><p:cNvPr id=`"2`" name=`"s`"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>"
}

Describe "isZipFile" -Tag Io {
    It "ZIPなら `$true" {
        $path = "$TestDrive\zip.docx"
        newZip $path @{ "a.txt" = "a" }
        isZipFile $path | Should Be $true
    }

    It "旧形式・パスワード付き（複合ドキュメント形式）なら `$false" {
        $path = "$TestDrive\cfb.docx"
        [System.IO.File]::WriteAllBytes($path, [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1))
        isZipFile $path | Should Be $false
    }

    It "空ファイルなら `$false" {
        $path = "$TestDrive\empty.docx"
        [System.IO.File]::WriteAllBytes($path, [byte[]]@())
        isZipFile $path | Should Be $false
    }
}

Describe "isCompoundFile" -Tag Io {
    It "複合ドキュメント形式（旧形式・パスワード付き）なら `$true" {
        $path = "$TestDrive\cfb.ppt"
        [System.IO.File]::WriteAllBytes($path, [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00))
        isCompoundFile $path | Should Be $true
    }

    It "ZIP・テキスト・空ファイルなら `$false" {
        $zip = "$TestDrive\zip.pptx"
        newZip $zip @{ "a.txt" = "a" }
        isCompoundFile $zip | Should Be $false

        $text = "$TestDrive\text.pptx"
        [System.IO.File]::WriteAllText($text, "壊れたファイル", (New-Object System.Text.UTF8Encoding($false)))
        isCompoundFile $text | Should Be $false

        $empty = "$TestDrive\empty.ppt"
        [System.IO.File]::WriteAllBytes($empty, [byte[]]@())
        isCompoundFile $empty | Should Be $false
    }
}

Describe "resolveZipPath" -Tag Io {
    It "相対パスを解決する" {
        resolveZipPath "ppt/slides" "../notesSlides/notesSlide1.xml" | Should Be "ppt/notesSlides/notesSlide1.xml"
        resolveZipPath "ppt" "slides/slide1.xml" | Should Be "ppt/slides/slide1.xml"
    }

    It "/ で始まるパスはZIPのルートからとする" {
        resolveZipPath "ppt/slides" "/ppt/media/a.png" | Should Be "ppt/media/a.png"
    }
}

Describe "readDocxUnits" -Tag Io {
    $body = @(
        (wPara "見出し"),
        # タブはスペースに、変更履歴の削除・フィールドコードは読まない
        '<w:p><w:r><w:t xml:space="preserve">本文 </w:t></w:r><w:r><w:tab/><w:t>りんご</w:t></w:r><w:del><w:r><w:delText>削除済み</w:delText></w:r></w:del><w:r><w:instrText>PAGE</w:instrText></w:r></w:p>',
        # 表: 1行 = セルのタブ区切り。セル内の複数段落はスペース区切り。末尾の空セルは除く
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>品名</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>価格</w:t></w:r></w:p></w:tc></w:tr>',
        '<w:tr><w:tc><w:p><w:r><w:t>複数</w:t></w:r></w:p><w:p><w:r><w:t>段落</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>',
        '<w:p><w:r><w:lastRenderedPageBreak/><w:t>2ページ目</w:t></w:r></w:p>',
        # テキストボックス: 互換用の代替表示（mc:Fallback）は読まない
        '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><w:txbxContent>', (wPara "テキストボックス"), '</w:txbxContent></w:drawing></mc:Choice>',
        '<mc:Fallback><w:pict><w:txbxContent>', (wPara "テキストボックス"), '</w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>',
        # 手動の改ページ + 直後の lastRenderedPageBreak は1回の改ページ
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:lastRenderedPageBreak/><w:t>3ページ目</w:t></w:r></w:p>',
        # 手動の改ページの後に lastRenderedPageBreak が無い場合も改ページ
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p>', (wPara "4ページ目")
    ) -join ""

    $path = "$TestDrive\報告書[1].docx"
    newZip $path @{
        "word/document.xml"  = "<w:document ${wNs}><w:body>${body}<w:sectPr/></w:body></w:document>"
        "word/header1.xml"   = "<w:hdr ${wNs}>$(wPara '社外秘')</w:hdr>"
        "word/header2.xml"   = "<w:hdr ${wNs}>$(wPara '社外秘')</w:hdr>"
        "word/footer1.xml"   = "<w:ftr ${wNs}>$(wPara 'フッター')</w:ftr>"
        "word/footnotes.xml" = "<w:footnotes ${wNs}><w:footnote w:type=`"separator`" w:id=`"-1`"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id=`"1`">$(wPara '脚注本文')</w:footnote></w:footnotes>"
    }
    $units = readDocxUnits $path

    It "ページ・ヘッダー/フッター・脚注の順に分ける" {
        @($units.Keys) -join "|" | Should Be "ページ001|ページ002|ページ003|ページ004|ページ002[図形]|ヘッダー・フッター|脚注"
    }

    It "手動の改ページは、直後の保存時のページ区切りの有無にかかわらず1ページとして数える" {
        @($units["ページ003"]) -join "|" | Should Be "3ページ目"
        @($units["ページ004"]) -join "|" | Should Be "4ページ目"
    }

    It "段落と表の行を1行ずつにする" {
        @($units["ページ001"]) -join "|" | Should Be "見出し|本文  りんご|品名`t価格|複数 段落"
    }

    It "保存時のページ区切りで次のページにし、テキストボックスは本文から分けて、そのページの図形として1回だけ読む" {
        @($units["ページ002"]) -join "|" | Should Be "2ページ目"
        @($units["ページ002[図形]"]) -join "|" | Should Be "テキストボックス"
    }

    It "ヘッダー・フッターの重複を除く" {
        @($units["ヘッダー・フッター"]) -join "|" | Should Be "社外秘|フッター"
    }

    It "脚注の区切り線は読まない" {
        @($units["脚注"]) -join "|" | Should Be "脚注本文"
    }
}

Describe "readDocxUnits（保存時のページ区切りが無い文書）" -Tag Io {
    $body = @(
        '<w:p><w:r><w:t>A</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>',
        (wPara "B"),
        '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>C</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>D</w:t></w:r></w:p>',
        (wPara "E"),
        '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>F</w:t></w:r></w:p>'
    ) -join ""

    $path = "$TestDrive\explicit.docx"
    newZip $path @{ "word/document.xml" = "<w:document ${wNs}><w:body>${body}</w:body></w:document>" }
    $units = readDocxUnits $path

    It "手動の改ページ・セクション区切り（連続以外）・段落前で改ページ でページを数える" {
        @($units.Keys) -join "|" | Should Be "ページ001|ページ002|ページ003|ページ004"
        @($units["ページ001"]) -join "|" | Should Be "A"
        @($units["ページ002"]) -join "|" | Should Be "B|C|D"
        @($units["ページ003"]) -join "|" | Should Be "E"
        @($units["ページ004"]) -join "|" | Should Be "F"
    }
}

Describe "readPptxUnits" -Tag Io {
    $slideRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    $notesRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"

    $table = '<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>項目</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>値</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    $group = "<p:grpSp>$(pShape 'グループ内')</p:grpSp>"

    $path = "$TestDrive\提案.pptx"
    newZip $path @{
        # 表示順は slide2.xml → slide1.xml
        "ppt/presentation.xml"            = "<p:presentation ${pNs}><p:sldIdLst><p:sldId id=`"256`" r:id=`"rId3`"/><p:sldId id=`"257`" r:id=`"rId2`"/></p:sldIdLst></p:presentation>"
        "ppt/_rels/presentation.xml.rels" = "<Relationships ${relNs}><Relationship Id=`"rId2`" Type=`"${slideRel}`" Target=`"slides/slide1.xml`"/><Relationship Id=`"rId3`" Type=`"${slideRel}`" Target=`"slides/slide2.xml`"/></Relationships>"
        "ppt/slides/slide2.xml"           = "<p:sld ${pNs}><p:cSld><p:spTree>$(pShape '表紙' 'title')${group}$(pShape '1' 'sldNum')$(pShape '社外秘' 'ftr')${table}</p:spTree></p:cSld></p:sld>"
        "ppt/slides/slide1.xml"           = "<p:sld ${pNs} show=`"0`"><p:cSld><p:spTree>$(pShape '非表示の内容')$(pShape '社外秘' 'ftr')$(pShape '2026/09/19' 'dt')</p:spTree></p:cSld></p:sld>"
        "ppt/slides/_rels/slide1.xml.rels" = "<Relationships ${relNs}><Relationship Id=`"rId1`" Type=`"${notesRel}`" Target=`"../notesSlides/notesSlide1.xml`"/></Relationships>"
        "ppt/notesSlides/notesSlide1.xml" = "<p:notes ${pNs}><p:cSld><p:spTree>$(pShape 'ノート本文' 'body')$(pShape '2' 'sldNum')</p:spTree></p:cSld></p:notes>"
    }
    $units = readPptxUnits $path

    It "スライドの表示順に番号を付け、非表示スライド・ノート・フッターを分ける" {
        @($units.Keys) -join "|" | Should Be "スライド001|スライド002（非表示）|スライド002_ノート|ヘッダー・フッター"
    }

    It "図形・グループ内の図形・表を読み、スライド番号・フッターは読まない" {
        @($units["スライド001"]) -join "|" | Should Be "表紙|グループ内|項目`t値"
        @($units["スライド002（非表示）"]) -join "|" | Should Be "非表示の内容"
    }

    It "スライドのフッターは重複を除いてまとめ、日付は読まない" {
        @($units["ヘッダー・フッター"]) -join "|" | Should Be "社外秘"
    }

    It "ノートの本文を読み、スライド番号は読まない" {
        @($units["スライド002_ノート"]) -join "|" | Should Be "ノート本文"
    }
}

Describe "writeUnits" -Tag Io {
    It "場所ごとにTSVを出力し、空の場所は出力しない" {
        $outDir = "$TestDrive\out[1]"
        [System.IO.Directory]::CreateDirectory($outDir) | Out-Null

        $units = New-Object System.Collections.Specialized.OrderedDictionary
        addUnitLines $units "ページ001" @("a`tb`t", "  ")
        addUnitLines $units "ページ002" @("", " ")
        addUnitLines $units "スライド001_ノート" @("メモ")

        # ファイル名はフォルダ名（インデクサが作業フォルダから移す先）になるため、TSVの名前は場所だけ
        writeUnits $units $outDir | Should Be 2
        [System.IO.File]::ReadAllText("$outDir\ページ001.tsv") | Should Be "a`tb`r`n"
        Test-Path -LiteralPath "$outDir\ページ002.tsv" | Should Be $false
        # 場所の _ は符号化する（以前の形式のTSVと見分けるため）
        Test-Path -LiteralPath "$outDir\スライド001%5Fノート.tsv" | Should Be $true
    }
}

$xNs = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$xdrNs = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
$officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

function xAnchor([int]$col, [int]$row, [string]$shapes) {
    # 左上が (col, row)（0 から数える）の図形の位置
    return "<xdr:twoCellAnchor><xdr:from><xdr:col>$col</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>$row</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
        "<xdr:to><xdr:col>$($col + 2)</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>$($row + 2)</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
        "${shapes}<xdr:clientData/></xdr:twoCellAnchor>"
}

function xSp([string[]]$paragraphs) {
    $body = ($paragraphs | ForEach-Object { "<a:p><a:r><a:t>$_</a:t></a:r></a:p>" }) -join ""
    return "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=`"2`" name=`"s`"/><xdr:cNvSpPr txBox=`"1`"/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/>${body}</xdr:txBody></xdr:sp>"
}

Describe "readXlsxObjectUnits" -Tag Io {
    $path = "$TestDrive\objects.xlsx"
    $group = xAnchor 0 9 ("<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=`"9`" name=`"g`"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>" +
        (xSp @("グループ1")) + (xSp @("グループ2")) + "</xdr:grpSp>")
    $alternate = "<mc:AlternateContent><mc:Choice Requires=`"a14`">$(xAnchor 0 20 (xSp @('代替表示あり')))</mc:Choice>" +
        "<mc:Fallback>$(xAnchor 0 20 (xSp @('代替表示あり')))</mc:Fallback></mc:AlternateContent>"
    $chart = xAnchor 3 30 "<xdr:graphicFrame macro=`"`"><xdr:nvGraphicFramePr><xdr:cNvPr id=`"3`" name=`"c`"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri=`"x`"/></a:graphic></xdr:graphicFrame>"
    newZip $path @{
        "xl/workbook.xml" = "<workbook $xNs><sheets>" +
            "<sheet name=`"売上`" sheetId=`"1`" r:id=`"rId1`"/>" +
            "<sheet name=`"隠し`" sheetId=`"2`" state=`"hidden`" r:id=`"rId2`"/>" +
            "<sheet name=`"グラフ`" sheetId=`"3`" r:id=`"rId3`"/>" +
            "<sheet name=`"空`" sheetId=`"4`" r:id=`"rId4`"/>" +
            "</sheets></workbook>"
        "xl/_rels/workbook.xml.rels" = "<Relationships $relNs>" +
            "<Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
            "<Relationship Id=`"rId2`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet2.xml`"/>" +
            "<Relationship Id=`"rId3`" Type=`"$officeRel/chartsheet`" Target=`"chartsheets/sheet1.xml`"/>" +
            "<Relationship Id=`"rId4`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet3.xml`"/>" +
            "</Relationships>"
        "xl/worksheets/sheet1.xml" = "<worksheet $xNs/>"
        "xl/worksheets/_rels/sheet1.xml.rels" = "<Relationships $relNs>" +
            "<Relationship Id=`"rId1`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing1.xml`"/>" +
            "<Relationship Id=`"rId2`" Type=`"$officeRel/comments`" Target=`"../comments1.xml`"/>" +
            "<Relationship Id=`"rId3`" Type=`"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment`" Target=`"../threadedComments/threadedComment1.xml`"/>" +
            "<Relationship Id=`"rId4`" Type=`"$officeRel/hyperlink`" Target=`"https://example.com/`" TargetMode=`"External`"/>" +
            "</Relationships>"
        # 下の図形を先に書いておき、上の行からの順に並べ直されることを確かめる
        "xl/drawings/drawing1.xml" = "<xdr:wsDr $xdrNs>$(xAnchor 1 4 (xSp @('下の図形')))$(xAnchor 5 1 (xSp @('1段落目', '2段落目')))" +
            "$(xAnchor 2 1 (xSp @('引用&quot;あり')))$group$alternate$chart</xdr:wsDr>"
        "xl/comments1.xml" = "<comments $xNs><authors><author>test</author></authors><commentList>" +
            "<comment ref=`"B3`" authorId=`"0`"><text><r><t>test:</t></r><r><t xml:space=`"preserve`">`n価格は税抜</t></r><rPh sb=`"0`" eb=`"1`"><t>ふりがな</t></rPh></text></comment>" +
            "<comment ref=`"A2`" authorId=`"0`"><text><t>メモ</t></text></comment>" +
            "<comment ref=`"C1`" authorId=`"0`"><text><t>[スレッド化されたコメント] 古い版向けの案内文</t></text></comment>" +
            "</commentList></comments>"
        "xl/threadedComments/threadedComment1.xml" = "<ThreadedComments xmlns=`"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments`">" +
            "<threadedComment ref=`"C1`" id=`"{1}`"><text>確認してください</text></threadedComment>" +
            "<threadedComment ref=`"C1`" id=`"{2}`" parentId=`"{1}`"><text>確認しました</text></threadedComment>" +
            "</ThreadedComments>"
        "xl/worksheets/sheet2.xml" = "<worksheet $xNs/>"
        "xl/worksheets/_rels/sheet2.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing2.xml`"/></Relationships>"
        "xl/drawings/drawing2.xml" = "<xdr:wsDr $xdrNs>$(xAnchor 0 0 (xSp @('非表示シートの図形')))</xdr:wsDr>"
        "xl/worksheets/sheet3.xml" = "<worksheet $xNs/>"
    }
    $units = readXlsxObjectUnits $path
    $ls = [char]0x2028

    It "表示シートの図形・コメントだけを、シートごとの場所にする（非表示シート・グラフシート・何も無いシートは出さない）" {
        @($units.Keys) -join "|" | Should Be "売上[図形]|売上[コメント]"
    }

    It "図形は左上のセル番地と文字を並べ、上の行から（同じ行は左から）の順にする。文字の無い図形（グラフなど）は出さない" {
        $lines = @($units["売上[図形]"])
        $lines[0] | Should Be "C2`t`"引用`"`"あり`""
        $lines[1] | Should Be "F2`t`"1段落目${ls}2段落目`""
        $lines[2] | Should Be "B5`t下の図形"
        $lines.Count | Should Be 5
    }

    It "グループ化した図形はまとめて 1 つにし、互換用の代替表示（mc:Fallback）は読まない" {
        $lines = @($units["売上[図形]"])
        $lines[3] | Should Be "A10`t`"グループ1${ls}グループ2`""
        $lines[4] | Should Be "A21`t代替表示あり"
    }

    It "コメントはセルの順に並べ、ふりがなは読まない。スレッド形式のコメントがあるセルは、その文字（返信を含む）を使う" {
        @($units["売上[コメント]"]) -join "|" |
            Should Be "C1`t`"確認してください${ls}確認しました`"|A2`tメモ|B3`t`"test:${ls}価格は税抜`""
    }

    It "同じセルに左上がある図形が多数あっても、XML の順（作った順）を保つ" {
        $same = "$TestDrive\same_cell.xlsx"
        $anchors = (1..30 | ForEach-Object { xAnchor 1 1 (xSp @("図形$_")) }) -join ""
        newZip $same @{
            "xl/workbook.xml" = "<workbook $xNs><sheets><sheet name=`"S`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs/>"
            "xl/worksheets/_rels/sheet1.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing1.xml`"/></Relationships>"
            "xl/drawings/drawing1.xml" = "<xdr:wsDr $xdrNs>$anchors</xdr:wsDr>"
        }
        @((readXlsxObjectUnits $same)["S[図形]"]) -join "|" | Should Be ((1..30 | ForEach-Object { "B2`t図形$_" }) -join "|")
    }

    It "Excel のブックでない ZIP（.xlsb など）は何も返さない" {
        $xlsb = "$TestDrive\binary.xlsb"
        newZip $xlsb @{ "xl/workbook.bin" = "binary" }
        (readXlsxObjectUnits $xlsb).Count | Should Be 0
    }
}

# Word・PowerPoint の図形（テキストボックス・SmartArt・グラフ）とコメント
$dgmNs = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$cNs = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$docRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

function wTextBox([string[]]$paragraphs) {
    # 段落の中のテキストボックス（互換用の代替表示付き。代替表示は読まない）
    $body = ($paragraphs | ForEach-Object { wPara $_ }) -join ""
    return "<w:p><w:r><mc:AlternateContent><mc:Choice Requires=`"wps`"><w:drawing><w:txbxContent>$body</w:txbxContent></w:drawing></mc:Choice>" +
        "<mc:Fallback><w:pict><w:txbxContent>$body</w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>"
}

function smartArt([string]$relId) {
    return "<a:graphic xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`"><a:graphicData><dgm:relIds $dgmNs r:dm=`"$relId`" r:lo=`"x`" r:qs=`"x`" r:cs=`"x`"/></a:graphicData></a:graphic>"
}

function chartRef([string]$relId) {
    return "<a:graphic xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`"><a:graphicData><c:chart $cNs r:id=`"$relId`"/></a:graphicData></a:graphic>"
}

$diagramXml = "<dgm:dataModel $dgmNs xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`"><dgm:ptLst>" +
    "<dgm:pt modelId=`"1`" type=`"doc`"><dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t></dgm:pt>" +
    "<dgm:pt modelId=`"2`"><dgm:t><a:bodyPr/><a:p><a:r><a:t>企画</a:t></a:r></a:p></dgm:t></dgm:pt>" +
    "<dgm:pt modelId=`"3`"><dgm:t><a:bodyPr/><a:p><a:r><a:t>設計</a:t></a:r></a:p></dgm:t></dgm:pt>" +
    "</dgm:ptLst></dgm:dataModel>"
$chartXml = "<c:chartSpace $cNs><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>月別売上</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
    "<c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:f>S!B1</c:f><c:strCache><c:ptCount val=`"1`"/><c:pt idx=`"0`"><c:v>東京支店</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
    "<c:cat><c:strRef><c:f>S!A2:A3</c:f><c:strCache><c:pt idx=`"0`"><c:v>4月</c:v></c:pt><c:pt idx=`"1`"><c:v>5月</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
    "<c:val><c:numRef><c:f>S!B2:B3</c:f><c:numCache><c:pt idx=`"0`"><c:v>1200</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>" +
    "<c:ser><c:tx><c:strRef><c:strCache><c:pt idx=`"0`"><c:v>大阪支店</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
    "<c:cat><c:strRef><c:strCache><c:pt idx=`"0`"><c:v>4月</c:v></c:pt></c:strCache></c:strRef></c:cat></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"

Describe "readDocxUnits（図形・コメント）" -Tag Io {
    $path = "$TestDrive\objects.docx"
    $table = "<w:tbl><w:tr><w:tc>$(wPara '表のセル')$(wTextBox @('セルの中の箱'))</w:tc><w:tc>$(wPara '右のセル')</w:tc></w:tr></w:tbl>"
    $boxTable = "<w:p><w:r><w:drawing><w:txbxContent><w:tbl><w:tr><w:tc>$(wPara '箱の表A')</w:tc><w:tc>$(wPara '箱の表B')</w:tc></w:tr></w:tbl></w:txbxContent></w:drawing></w:r></w:p>"
    $body = (wPara "1ページ目の本文") + (wTextBox @("箱の1段落目", "箱の2段落目")) + $table + $boxTable +
        "<w:p><w:r><w:t>コメントを付けた段落</w:t></w:r><w:r><w:commentReference w:id=`"0`"/></w:r><w:r><w:commentReference w:id=`"1`"/></w:r></w:p>" +
        "<w:p><w:r><w:br w:type=`"page`"/></w:r></w:p>" +
        "<w:p><w:r><w:drawing>$(smartArt 'rId10')</w:drawing></w:r></w:p>" +
        "<w:p><w:r><w:drawing>$(chartRef 'rId11')</w:drawing></w:r><w:r><w:commentReference w:id=`"2`"/></w:r></w:p>"
    newZip $path @{
        "word/document.xml" = "<w:document $wNs><w:body>$body</w:body></w:document>"
        "word/_rels/document.xml.rels" = "<Relationships $relNs>" +
            "<Relationship Id=`"rId10`" Type=`"$docRel/diagramData`" Target=`"diagrams/data1.xml`"/>" +
            "<Relationship Id=`"rId11`" Type=`"$docRel/chart`" Target=`"charts/chart1.xml`"/></Relationships>"
        "word/diagrams/data1.xml" = $diagramXml
        "word/charts/chart1.xml" = $chartXml
        "word/comments.xml" = "<w:comments $wNs>" +
            "<w:comment w:id=`"0`" w:author=`"test`"><w:p><w:r><w:t>確認してください</w:t></w:r></w:p><w:p><w:r><w:t>2段落目</w:t></w:r></w:p></w:comment>" +
            "<w:comment w:id=`"1`" w:author=`"test`"><w:p><w:r><w:t>返信です</w:t></w:r></w:p></w:comment>" +
            "<w:comment w:id=`"2`" w:author=`"test`"><w:p><w:r><w:t>グラフの数字を確認</w:t></w:r></w:p></w:comment>" +
            "<w:comment w:id=`"3`" w:author=`"test`"><w:p><w:r><w:t>参照の無いコメント</w:t></w:r></w:p></w:comment></w:comments>"
    }
    $units = readDocxUnits $path

    It "テキストボックスの文字は本文から分け、段落をスペースでつないで図形 1 つにする（代替表示は読まない）" {
        @($units["ページ001"]) -join "|" | Should Be "1ページ目の本文|表のセル`t右のセル|コメントを付けた段落"
        @($units["ページ001[図形]"]) -join "|" | Should Be "箱の1段落目 箱の2段落目|セルの中の箱|箱の表A`t箱の表B"
    }

    It "SmartArt・グラフの文字は、そのページの図形にする（グラフはタイトル・系列名・項目名。数値と重複は読まない）" {
        @($units["ページ002[図形]"]) -join "|" | Should Be "企画 設計|月別売上 東京支店 4月 5月 大阪支店"
    }

    It "コメントは付けた所のページに、返信も 1 件ずつ入れる。本文に参照の無いコメントは「文書」にまとめる" {
        @($units["ページ001[コメント]"]) -join "|" | Should Be "確認してください 2段落目|返信です"
        @($units["ページ002[コメント]"]) -join "|" | Should Be "グラフの数字を確認"
        @($units["文書[コメント]"]) -join "|" | Should Be "参照の無いコメント"
    }
}

Describe "readPptxUnits（図形・コメント）" -Tag Io {
    $path = "$TestDrive\objects.pptx"
    $slideRel = "$docRel/slide"
    $frame = { param($graphic) "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=`"9`" name=`"f`"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/>$graphic</p:graphicFrame>" }
    newZip $path @{
        "ppt/presentation.xml" = "<p:presentation $pNs><p:sldIdLst><p:sldId id=`"256`" r:id=`"rId2`"/><p:sldId id=`"257`" r:id=`"rId3`"/></p:sldIdLst></p:presentation>"
        "ppt/_rels/presentation.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId2`" Type=`"$slideRel`" Target=`"slides/slide1.xml`"/><Relationship Id=`"rId3`" Type=`"$slideRel`" Target=`"slides/slide2.xml`"/></Relationships>"
        "ppt/slides/slide1.xml" = "<p:sld $pNs><p:cSld><p:spTree>$(pShape 'スライドの本文')$(& $frame (smartArt 'rId5'))$(& $frame (chartRef 'rId6'))</p:spTree></p:cSld></p:sld>"
        "ppt/slides/_rels/slide1.xml.rels" = "<Relationships $relNs>" +
            "<Relationship Id=`"rId5`" Type=`"$docRel/diagramData`" Target=`"../diagrams/data1.xml`"/>" +
            "<Relationship Id=`"rId6`" Type=`"$docRel/chart`" Target=`"../charts/chart1.xml`"/>" +
            "<Relationship Id=`"rId7`" Type=`"$docRel/comments`" Target=`"../comments/comment1.xml`"/></Relationships>"
        "ppt/diagrams/data1.xml" = $diagramXml
        "ppt/charts/chart1.xml" = $chartXml
        "ppt/comments/comment1.xml" = "<p:cmLst $pNs><p:cm authorId=`"0`" idx=`"1`"><p:pos x=`"0`" y=`"0`"/><p:text>旧形式のコメント</p:text></p:cm></p:cmLst>"
        "ppt/slides/slide2.xml" = "<p:sld $pNs show=`"0`"><p:cSld><p:spTree>$(pShape '非表示の本文')</p:spTree></p:cSld></p:sld>"
        "ppt/slides/_rels/slide2.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"http://schemas.microsoft.com/office/2018/10/relationships/comments`" Target=`"../comments/modernComment_1.xml`"/></Relationships>"
        "ppt/comments/modernComment_1.xml" = "<p188:cmLst xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`" xmlns:p188=`"http://schemas.microsoft.com/office/powerpoint/2018/8/main`">" +
            "<p188:cm id=`"{1}`" authorId=`"{9}`"><p188:replyLst><p188:reply id=`"{2}`" authorId=`"{9}`"><p188:txBody><a:bodyPr/><a:p><a:r><a:t>返信です</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst>" +
            "<p188:txBody><a:bodyPr/><a:p><a:r><a:t>新形式の</a:t></a:r></a:p><a:p><a:r><a:t>コメント</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"
    }
    $units = readPptxUnits $path

    It "テキストボックス・図形はスライドの本文のまま、SmartArt・グラフの文字はスライドの図形にする" {
        @($units["スライド001"]) -join "|" | Should Be "スライドの本文"
        @($units["スライド001[図形]"]) -join "|" | Should Be "企画 設計|月別売上 東京支店 4月 5月 大阪支店"
    }

    It "旧形式・新形式のコメントを読み、コメントの後に返信を 1 件ずつ入れる（非表示のスライドは場所の名前に付く）" {
        @($units["スライド001[コメント]"]) -join "|" | Should Be "旧形式のコメント"
        @($units["スライド002（非表示）[コメント]"]) -join "|" | Should Be "新形式の コメント|返信です"
    }
}
