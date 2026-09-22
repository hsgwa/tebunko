# 制限言語モード用の Office ファイルの読み取り（shared\office\office_reader_clm.ps1）のテスト。
# いつものインデクサの読み取り（office_reader.ps1）と、同じファイルから同じ場所・同じ行を返すことを突き合わせる。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\office\office_reader.ps1"
. "${scriptsDir}\shared\office\office_numfmt.ps1"
. "${scriptsDir}\shared\office\office_reader_clm.ps1"

# ユニット（場所 → 行）を比べられる文字列にする（OrderedDictionary の List と、順序付き辞書の配列のどちらでもよい）
function formatUnits {
    param ($units)
    return (@(foreach ($key in $units.Keys) { "#$key"; @($units[$key]) }) -join "`n")
}

function formatLines {
    param ([object[]]$lines)
    return (@($lines | ForEach-Object { "$($_.Page)|$($_.Text)" }) -join "`n")
}

Describe "office_reader_clm.ps1（テスト用のファイルで office_reader.ps1 と同じ）" -Tag Io {
    $work = Join-Path $TestDrive "zip"
    New-Item -ItemType Directory -Path $work | Out-Null
    $files = @(Get-ChildItem "${testDataDir}\office" -Recurse -File |
        Where-Object { $_.Extension -match '^\.(docx|docm|pptx|pptm|xlsx|xlsm)$' -and (isZipFile $_.FullName) } | Sort-Object FullName)

    It "調べるファイルがある（Word・PowerPoint・Excel）" {
        @($files | Where-Object { $_.Extension -match 'doc' }).Count -gt 5 | Should Be $true
        @($files | Where-Object { $_.Extension -match 'ppt' }).Count -gt 5 | Should Be $true
        @($files | Where-Object { $_.Extension -match 'xls' }).Count -gt 5 | Should Be $true
    }

    foreach ($file in $files) {
        $relative = $file.FullName.Substring($testDataDir.Length + 1)
        It "$relative" {
            if ($file.Extension -match 'doc') {
                $expected = readDocxUnits $file.FullName
                $actual = readDocxUnitsClm $file.FullName $work
            } elseif ($file.Extension -match 'ppt') {
                $expected = readPptxUnits $file.FullName
                $actual = readPptxUnitsClm $file.FullName $work
            } else {
                $expected = readXlsxObjectUnits $file.FullName
                $actual = readXlsxObjectUnitsClm $file.FullName $work
            }
            (formatUnits $actual) | Should BeExactly (formatUnits $expected)
            # 作業フォルダを残さない
            @(Get-ChildItem -LiteralPath $work).Count | Should Be 0
        }
    }
}

Describe "readXmlLinesClm（XML の書き方の違い）" -Tag Unit {
    $w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    It "接頭辞が違っても・既定の名前空間でも、名前空間で読む（readXmlLines と同じ）" {
        foreach ($xml in @(
                "<x:document xmlns:x=`"$w`"><x:body><x:p><x:r><x:t>本文</x:t></x:r></x:p></x:body></x:document>",
                "<document xmlns=`"$w`"><body><p><r><t>本文</t></r></p></body></document>",
                "<w:document xmlns:w=`"$w`"><w:body><w:p><w:r><w:t>本文</w:t></w:r></w:p><w:p xmlns:w=`"urn:other`"><w:r><w:t>別の名前空間</w:t></w:r></w:p></w:body></w:document>"
            )) {
            (formatLines (readXmlLinesClm $xml $w)) | Should BeExactly (formatLines (readXmlLines $xml $w))
        }
    }

    It "文字参照・CDATA・コメント・処理命令・改行の書き方" {
        $xml = "<?xml version=`"1.0`"?><!-- c --><w:document xmlns:w=`"$w`"><w:body><w:p><w:r><w:t>a&amp;b&lt;c&gt;&quot;&apos;&#12354;&#x1F600;&#x20BB7;</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t><![CDATA[<cdata> & x]]></w:t></w:r></w:p><w:p><w:r><w:t xml:space=`"preserve`">  前後`r`n改行`r</w:t></w:r></w:p></w:body></w:document>"
        (formatLines (readXmlLinesClm $xml $w)) | Should BeExactly (formatLines (readXmlLines $xml $w))
    }

    It "読み飛ばす要素（互換用の代替表示・移動元・文字の書式）と、入れ子の表・テキストボックス・改ページ" {
        $mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
        $xml = "<w:document xmlns:w=`"$w`" xmlns:mc=`"$mc`"><w:body>" +
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t><w:br w:type=`"page`"/><w:t>2</w:t></w:r></w:p>" +
            "<mc:AlternateContent><mc:Choice><w:p><w:r><w:t>選ぶ</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>代替</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>" +
            "<w:moveFrom><w:p><w:r><w:t>移動元</w:t></w:r></w:p></w:moveFrom>" +
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>入れ子</w:t></w:r></w:p></w:tc><w:tc/></w:tr></w:tbl></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>次</w:t><w:tab/><w:t>タブ</w:t><w:noBreakHyphen/></w:r><w:r><w:txbxContent><w:p><w:r><w:t>箱</w:t></w:r></w:p></w:txbxContent></w:r></w:p>" +
            "<w:p><w:pPr><w:sectPr><w:type w:val=`"continuous`"/></w:sectPr></w:pPr><w:r><w:t>連続</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:lastRenderedPageBreak/><w:t>節</w:t></w:r></w:p><w:p><w:r><w:t>最後</w:t></w:r></w:p>" +
            "</w:body></w:document>"
        foreach ($mode in @("none", "explicit", "rendered")) {
            $objectsA = New-Object System.Collections.Generic.List[object]
            $objectsB = @{}
            $expected = readXmlLines $xml $w $mode $null $objectsA
            $actual = readXmlLinesClm $xml $w $mode $null $objectsB
            (formatLines $actual) | Should BeExactly (formatLines $expected)
            (@(for ($i = 0; $i -lt $objectsB.Count; $i++) { "$($objectsB[$i].Kind)|$($objectsB[$i].Page)|$($objectsB[$i].Text)" }) -join ",") |
                Should BeExactly (@($objectsA | ForEach-Object { "$($_.Kind)|$($_.Page)|$($_.Text)" }) -join ",")
        }
    }
}

Describe "decodeXmlText / getXmlAttribute / getXmlDescendants" -Tag Unit {
    It "実体参照と文字参照（サロゲートペアを含む）を戻す" {
        decodeXmlText "a&amp;b&lt;&gt;&quot;&apos;&#65;&#x42;&#x20BB7;&unknown;" | Should BeExactly "a&b<>`"'AB𠮷&unknown;"
        decodeXmlText "そのまま" | Should BeExactly "そのまま"
    }

    It "属性を名前空間つきで読み、子孫を文書の順に返す" {
        [xml]$doc = '<a xmlns:w="u"><w:b w:id="3" id="4"><c>1</c><w:c>2</w:c><d><c>3</c></d></w:b></a>'
        $b = @($doc.SelectNodes("//*"))[1]
        getXmlAttribute $b "id" "u" | Should Be "3"
        getXmlAttribute $b "id" | Should Be "4"
        getXmlAttribute $b "none" | Should Be ""
        (@(getXmlDescendants $b "c") | ForEach-Object { $_.InnerText }) -join "," | Should Be "1,2,3"
        (@(getXmlDescendants $b "c" "") | ForEach-Object { $_.InnerText }) -join "," | Should Be "1,3"
    }
}

Describe "openOfficeZip / writeUnitsClm" -Tag Io {
    $work = Join-Path $TestDrive "zip2"
    New-Item -ItemType Directory -Path $work | Out-Null

    It "CP932 に無い文字を含むパスのファイルも読める" {
        $source = @(Get-ChildItem "${testDataDir}\office\Word" -Recurse -Filter "*.docx")[0]
        $copy = Join-Path $TestDrive "𠮷野家 😀.docx"
        Copy-Item -LiteralPath $source.FullName -Destination $copy
        (formatUnits (readDocxUnitsClm $copy $work)) | Should BeExactly (formatUnits (readDocxUnits $source.FullName))
    }

    It "ZIP でなければ、分かるメッセージで例外にし、作業フォルダを残さない" {
        $text = Join-Path $TestDrive "壊れた.docx"
        Set-Content -LiteralPath $text -Value "ZIP ではない"
        { readDocxUnitsClm $text $work } | Should Throw "ZIP として読めません"
        @(Get-ChildItem -LiteralPath $work).Count | Should Be 0
    }

    It "本文・プレゼンテーション情報が無ければ、いつもと同じメッセージで例外にする" {
        $dir = Join-Path $TestDrive "empty_zip"
        New-Item -ItemType Directory -Path "$dir\x" -Force | Out-Null
        Set-Content -LiteralPath "$dir\x\a.xml" -Value "<a/>"
        $zip = Join-Path $TestDrive "empty.zip"
        Compress-Archive -Path "$dir\x" -DestinationPath $zip
        { readDocxUnitsClm $zip $work } | Should Throw "word/document.xml"
        { readPptxUnitsClm $zip $work } | Should Throw "ppt/presentation.xml"
        (readXlsxObjectUnitsClm $zip $work).Count | Should Be 0
    }

    It "writeUnitsClm は writeUnits と同じ中身のファイルを書く（空の行・空のユニットは書かない）" {
        $expectedDir = Join-Path $TestDrive "units_a"
        $actualDir = Join-Path $TestDrive "units_b"
        New-Item -ItemType Directory -Path $expectedDir, $actualDir | Out-Null
        $a = New-Object System.Collections.Specialized.OrderedDictionary
        $a["ページ001"] = New-Object System.Collections.Generic.List[string]
        foreach ($line in @("1 行目  ", "", "𠮷")) { $a["ページ001"].Add($line) }
        $a["空"] = New-Object System.Collections.Generic.List[string]
        $a["空"].Add("   ")
        $a["シート_1[図形]"] = New-Object System.Collections.Generic.List[string]
        $a["シート_1[図形]"].Add("A1`t図形")
        $b = [ordered]@{ "ページ001" = [string[]]@("1 行目  ", "", "𠮷"); "空" = [string[]]@("   "); "シート_1[図形]" = [string[]]@("A1`t図形") }
        writeUnits $a $expectedDir | Should Be 2
        writeUnitsClm $b $actualDir | Should Be 2
        foreach ($name in @(Get-ChildItem -LiteralPath $expectedDir | ForEach-Object { $_.Name })) {
            ([System.IO.File]::ReadAllBytes((Join-Path $actualDir $name)) -join ",") | Should Be ([System.IO.File]::ReadAllBytes((Join-Path $expectedDir $name)) -join ",")
        }
        @(Get-ChildItem -LiteralPath $actualDir).Count | Should Be 2
    }
}

# --- 本文以外（SmartArt・グラフ・コメント）のテスト用のファイル ---
function newTestZip {
    # ZIP 内のパス → 内容 の辞書から ZIP を作る（office_reader.Tests.ps1 の newZip と同じ）
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

Describe "本文以外（SmartArt・グラフ・コメント）も office_reader.ps1 と同じ" -Tag Io {
    $work = Join-Path $TestDrive "zip3"
    New-Item -ItemType Directory -Path $work | Out-Null
    $relNs2 = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
    $docRel2 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    $wNs2 = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
        'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $pNs2 = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $aNs2 = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    $cNs2 = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    $chartXml = "<c:chartSpace $cNs2><c:chart><c:plotArea><c:barChart><c:ser>" +
        "<c:tx><c:strRef><c:strCache><c:pt idx=`"0`"><c:v>売上</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
        "<c:cat><c:multiLvlStrRef><c:multiLvlStrCache>" +
        "<c:lvl><c:pt idx=`"0`"><c:v>上期</c:v></c:pt></c:lvl><c:lvl><c:pt idx=`"0`"><c:v>2024年</c:v></c:pt></c:lvl>" +
        "</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>" +
        "<c:val><c:numRef><c:numCache><c:pt idx=`"0`"><c:v>100</c:v></c:pt></c:numCache></c:numRef></c:val>" +
        "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
    $diagramXml = "<dsp:drawing xmlns:dsp=`"http://schemas.microsoft.com/office/drawing/2008/diagram`" $aNs2>" +
        "<dsp:spTree><dsp:sp><dsp:txBody><a:bodyPr/><a:p><a:r><a:t>手順1</a:t></a:r></a:p>" +
        "<a:p><a:r><a:t>手順2</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>"

    It "Word: SmartArt・グラフ・コメント（返信つき）を同じ場所・同じ行で読む" {
        $path = Join-Path $TestDrive "本文以外.docx"
        newTestZip $path @{
            "word/document.xml" = "<w:document $wNs2><w:body>" +
                "<w:p><w:r><w:t>本文</w:t></w:r><w:commentReference w:id=`"1`"/></w:p>" +
                "<w:p><w:r><w:drawing><dgm:relIds r:dm=`"rId5`"/></w:drawing></w:r></w:p>" +
                "<w:p><w:r><w:drawing><c:chart r:id=`"rId6`"/></w:drawing></w:r></w:p>" +
                "</w:body></w:document>"
            "word/_rels/document.xml.rels" = "<Relationships $relNs2>" +
                "<Relationship Id=`"rId5`" Type=`"$docRel2/diagramDrawing`" Target=`"diagrams/drawing1.xml`"/>" +
                "<Relationship Id=`"rId6`" Type=`"$docRel2/chart`" Target=`"charts/chart1.xml`"/>" +
                "<Relationship Id=`"rId7`" Type=`"$docRel2/comments`" Target=`"comments.xml`"/>" +
                "</Relationships>"
            "word/diagrams/drawing1.xml" = $diagramXml
            "word/charts/chart1.xml" = $chartXml
            "word/comments.xml" = "<w:comments $wNs2><w:comment w:id=`"1`"><w:p><w:r><w:t>直してください</w:t></w:r></w:p></w:comment></w:comments>"
        }
        (formatUnits (readDocxUnitsClm $path $work)) | Should BeExactly (formatUnits (readDocxUnits $path))
        @(Get-ChildItem -LiteralPath $work).Count | Should Be 0
    }

    It "Word: 参照先が無い SmartArt・グラフは空にする（例外にしない）" {
        $path = Join-Path $TestDrive "参照なし.docx"
        newTestZip $path @{
            "word/document.xml" = "<w:document $wNs2><w:body>" +
                "<w:p><w:r><w:drawing><dgm:relIds r:dm=`"rId9`"/></w:drawing></w:r></w:p>" +
                "<w:p><w:r><w:drawing><c:chart r:id=`"rId8`"/></w:drawing></w:r></w:p>" +
                "</w:body></w:document>"
            "word/_rels/document.xml.rels" = "<Relationships $relNs2>" +
                "<Relationship Id=`"rId8`" Type=`"$docRel2/chart`" Target=`"charts/無い.xml`"/></Relationships>"
        }
        (formatUnits (readDocxUnitsClm $path $work)) | Should BeExactly (formatUnits (readDocxUnits $path))
    }

    It "PowerPoint: グラフとコメント（返信つき・以前の形式）を同じ場所・同じ行で読む" {
        $path = Join-Path $TestDrive "本文以外.pptx"
        $threadNs = 'xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" ' +
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        newTestZip $path @{
            "ppt/presentation.xml" = "<p:presentation $pNs2><p:sldIdLst><p:sldId id=`"256`" r:id=`"rId1`"/></p:sldIdLst></p:presentation>"
            "ppt/_rels/presentation.xml.rels" = "<Relationships $relNs2>" +
                "<Relationship Id=`"rId1`" Type=`"$docRel2/slide`" Target=`"slides/slide1.xml`"/></Relationships>"
            "ppt/slides/slide1.xml" = "<p:sld $pNs2><p:cSld><p:spTree>" +
                "<p:sp><p:nvSpPr><p:cNvPr id=`"2`" name=`"s`"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
                "<p:txBody><a:bodyPr/><a:p><a:r><a:t>表題</a:t></a:r></a:p></p:txBody></p:sp>" +
                "<p:graphicFrame><a:graphic><a:graphicData><c:chart r:id=`"rId3`"/></a:graphicData></a:graphic></p:graphicFrame>" +
                "</p:spTree></p:cSld></p:sld>"
            "ppt/slides/_rels/slide1.xml.rels" = "<Relationships $relNs2>" +
                "<Relationship Id=`"rId3`" Type=`"$docRel2/chart`" Target=`"../charts/chart1.xml`"/>" +
                "<Relationship Id=`"rId4`" Type=`"http://schemas.microsoft.com/office/2018/10/relationships/comments`" Target=`"../comments/modernComment_256_1.xml`"/>" +
                "</Relationships>"
            "ppt/charts/chart1.xml" = $chartXml
            "ppt/comments/modernComment_256_1.xml" = "<p188:cmLst $threadNs><p188:cm><p188:txBody><a:bodyPr/><a:p><a:r><a:t>ここを直す</a:t></a:r></a:p></p188:txBody>" +
                "<p188:replyLst><p188:reply><p188:txBody><a:bodyPr/><a:p><a:r><a:t>直しました</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst>" +
                "</p188:cm></p188:cmLst>"
        }
        (formatUnits (readPptxUnitsClm $path $work)) | Should BeExactly (formatUnits (readPptxUnits $path))
    }

    It "Excel: スレッド化されたコメント（返信つき）を同じ場所・同じ行で読む" {
        $path = Join-Path $TestDrive "コメント.xlsx"
        $xNs2 = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        newTestZip $path @{
            "xl/workbook.xml" = "<workbook $xNs2><sheets><sheet name=`"売上`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs2>" +
                "<Relationship Id=`"rId1`" Type=`"$docRel2/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs2/>"
            "xl/worksheets/_rels/sheet1.xml.rels" = "<Relationships $relNs2>" +
                "<Relationship Id=`"rId2`" Type=`"$docRel2/comments`" Target=`"../comments1.xml`"/>" +
                "<Relationship Id=`"rId3`" Type=`"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment`" Target=`"../threadedComments/threadedComment1.xml`"/>" +
                "</Relationships>"
            "xl/comments1.xml" = "<comments $xNs2><authors><author>test</author></authors><commentList>" +
                "<comment ref=`"C1`" authorId=`"0`"><text><t>[スレッド化されたコメント] 古い版向けの案内文</t></text></comment>" +
                "<comment ref=`"A2`" authorId=`"0`"><text><t>メモ</t></text></comment>" +
                "</commentList></comments>"
            "xl/threadedComments/threadedComment1.xml" = "<ThreadedComments xmlns=`"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments`">" +
                "<threadedComment ref=`"C1`" id=`"{1}`"><text>確認してください</text></threadedComment>" +
                "<threadedComment ref=`"C1`" id=`"{2}`" parentId=`"{1}`"><text>確認しました</text></threadedComment>" +
                "<threadedComment ref=`"C2`" id=`"{3}`"><text>  </text></threadedComment>" +
                "</ThreadedComments>"
        }
        (formatUnits (readXlsxObjectUnitsClm $path $work)) | Should BeExactly (formatUnits (readXlsxObjectUnits $path))
    }

    It "Word: 互換用の代替表示（mc:Fallback）は読まない" {
        $path = Join-Path $TestDrive "代替表示.docx"
        newTestZip $path @{
            "word/document.xml" = "<w:document $wNs2><w:body><w:p><w:r><mc:AlternateContent>" +
                "<mc:Choice Requires=`"wps`"><w:drawing><w:txbxContent><w:p><w:r><w:t>図形の文字</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice>" +
                "<mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>図形の文字</w:t></w:r></w:p></w:txbxContent></w:pict></mc:Fallback>" +
                "</mc:AlternateContent></w:r></w:p></w:body></w:document>"
        }
        (formatUnits (readDocxUnitsClm $path $work)) | Should BeExactly (formatUnits (readDocxUnits $path))
    }
}

Describe "Excel のセル（Excel を使わずに読む）" -Tag Io {
    $work = Join-Path $TestDrive "cells"
    New-Item -ItemType Directory -Path $work | Out-Null
    $xNs = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $relNs3 = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
    $docRel3 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    function newCellBook {
        # シートの中身（<sheetData> の中）から、セルを読むための最小のブックを作る
        param (
            [string]$path,
            [string]$sheetData,
            [string]$sharedStrings = "",
            [string]$styles = "",
            [string]$workbookPr = "",
            [hashtable]$extra = @{}
        )

        $entries = @{
            "xl/workbook.xml" = "<workbook $xNs>${workbookPr}<sheets><sheet name=`"売上`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs3><Relationship Id=`"rId1`" Type=`"$docRel3/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs><sheetData>${sheetData}</sheetData></worksheet>"
        }
        if ($sharedStrings) { $entries["xl/sharedStrings.xml"] = $sharedStrings }
        if ($styles) { $entries["xl/styles.xml"] = $styles }
        foreach ($key in $extra.Keys) { $entries[$key] = $extra[$key] }
        newTestZip $path $entries
    }

    It "共有文字列・数値・数式の結果・エラー・真偽値を、Excel のテキスト保存と同じセルにする" {
        $path = Join-Path $TestDrive "cells.xlsx"
        newCellBook $path (
            "<row r=`"1`"><c r=`"A1`" t=`"s`"><v>0</v></c><c r=`"B1`" t=`"s`"><v>1</v></c><c r=`"C1`"><v>1234567</v></c></row>" +
            "<row r=`"2`"><c r=`"A2`" t=`"str`"><f>1+1</f><v>結果</v></c><c r=`"B2`" t=`"e`"><v>#DIV/0!</v></c><c r=`"C2`" t=`"b`"><v>1</v></c></row>" +
            "<row r=`"4`"><c r=`"B4`" t=`"inlineStr`"><is><t>じか書き</t></is></c></row>"
        ) ("<sst $xNs><si><t>品名</t></si><si><t>りんご`tと`"引用`"</t></si></sst>")
        $texts = readXlsxCellTextsClm $path $work
        @($texts.Keys) -join "|" | Should Be "売上"
        ($texts["売上"] -split "`r`n") -join "|" | Should Be (@(
                "品名`t`"りんご`tと`"`"引用`"`"`"`t1234567",
                "結果`t#DIV/0!`tTRUE",
                "",
                "`tじか書き"
            ) -join "|")
    }

    It "表示形式（styles.xml）を当て、日付は 1904 年方式も読む" {
        $styles = "<styleSheet $xNs><numFmts count=`"1`"><numFmt numFmtId=`"179`" formatCode=`"#,##0&quot;円&quot;`"/></numFmts>" +
            "<cellXfs count=`"3`"><xf numFmtId=`"0`"/><xf numFmtId=`"179`"/><xf numFmtId=`"14`"/></cellXfs></styleSheet>"
        $sheet = "<row r=`"1`"><c r=`"A1`" s=`"1`"><v>1500</v></c><c r=`"B1`" s=`"2`"><v>45383</v></c><c r=`"C1`"><v>0.5</v></c></row>"

        $path = Join-Path $TestDrive "fmt.xlsx"
        newCellBook $path $sheet "" $styles
        (readXlsxCellTextsClm $path $work)["売上"] | Should Be "`"1,500円`"`t4/1/2024`t0.5"

        $path1904 = Join-Path $TestDrive "fmt1904.xlsx"
        newCellBook $path1904 $sheet "" $styles "<workbookPr date1904=`"1`"/>"
        (readXlsxCellTextsClm $path1904 $work)["売上"] | Should Be "`"1,500円`"`t4/2/2028`t0.5"
    }

    It "非表示シート・グラフシート・値の無いシートは返さない" {
        $path = Join-Path $TestDrive "sheets.xlsx"
        newTestZip $path @{
            "xl/workbook.xml" = "<workbook $xNs><sheets>" +
                "<sheet name=`"表`" sheetId=`"1`" r:id=`"rId1`"/>" +
                "<sheet name=`"隠し`" sheetId=`"2`" state=`"hidden`" r:id=`"rId2`"/>" +
                "<sheet name=`"グラフ`" sheetId=`"3`" r:id=`"rId3`"/>" +
                "<sheet name=`"空`" sheetId=`"4`" r:id=`"rId4`"/>" +
                "</sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs3>" +
                "<Relationship Id=`"rId1`" Type=`"$docRel3/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$docRel3/worksheet`" Target=`"worksheets/sheet2.xml`"/>" +
                "<Relationship Id=`"rId3`" Type=`"$docRel3/chartsheet`" Target=`"chartsheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId4`" Type=`"$docRel3/worksheet`" Target=`"worksheets/sheet3.xml`"/>" +
                "</Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs><sheetData><row r=`"1`"><c r=`"A1`"><v>1</v></c></row></sheetData></worksheet>"
            "xl/worksheets/sheet2.xml" = "<worksheet $xNs><sheetData><row r=`"1`"><c r=`"A1`"><v>2</v></c></row></sheetData></worksheet>"
            "xl/chartsheets/sheet1.xml" = "<chartsheet $xNs/>"
            "xl/worksheets/sheet3.xml" = "<worksheet $xNs><sheetData/></worksheet>"
        }
        @((readXlsxCellTextsClm $path $work).Keys) -join "|" | Should Be "表"
    }

    It "シート名のタブ・改行（_x0009_）と、セルの _x000D_ を元の文字に戻す" {
        $path = Join-Path $TestDrive "escape.xlsx"
        newTestZip $path @{
            "xl/workbook.xml" = "<workbook $xNs><sheets><sheet name=`"タブ_x0009_あり`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs3><Relationship Id=`"rId1`" Type=`"$docRel3/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs><sheetData><row r=`"1`"><c r=`"A1`" t=`"s`"><v>0</v></c></row></sheetData></worksheet>"
            "xl/sharedStrings.xml" = "<sst $xNs><si><t>CR_x000D_で改行</t></si></sst>"
        }
        $texts = readXlsxCellTextsClm $path $work
        @($texts.Keys)[0] | Should Be "タブ`tあり"
        $texts["タブ`tあり"] | Should Be "`"CR$(${cellNewLine})で改行`""
    }

    It "行・列の番号どおりの位置に置き、間は空セル・空行にする" {
        $path = Join-Path $TestDrive "position.xlsx"
        newCellBook $path "<row r=`"3`"><c r=`"B3`"><v>1</v></c><c r=`"D3`"><v>2</v></c></row><row r=`"5`"><c r=`"A5`"><v>3</v></c></row>"
        ((readXlsxCellTextsClm $path $work)["売上"] -split "`r`n") -join "|" | Should Be "||`t1`t`t2||3"
    }

    It "writeXlsxTsvClm は、セル・図形・コメントの TSV をまとめて書く" {
        $path = Join-Path $TestDrive "all.xlsx"
        $xdrNs2 = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        newCellBook $path "<row r=`"1`"><c r=`"A1`" t=`"s`"><v>0</v></c></row>" ("<sst $xNs><si><t>品名</t></si></sst>") "" "" @{
            "xl/worksheets/_rels/sheet1.xml.rels" = "<Relationships $relNs3>" +
                "<Relationship Id=`"rId1`" Type=`"$docRel3/drawing`" Target=`"../drawings/drawing1.xml`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$docRel3/comments`" Target=`"../comments1.xml`"/></Relationships>"
            "xl/drawings/drawing1.xml" = "<xdr:wsDr $xdrNs2><xdr:twoCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
                "<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
                "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=`"2`" name=`"s`"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>図形の文字</a:t></a:r></a:p></xdr:txBody></xdr:sp>" +
                "<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"
            "xl/comments1.xml" = "<comments $xNs><authors><author>test</author></authors><commentList>" +
                "<comment ref=`"B3`" authorId=`"0`"><text><t>コメント</t></text></comment></commentList></comments>"
        }
        $outDir = Join-Path $TestDrive "all_out"
        New-Item -ItemType Directory -Path $outDir | Out-Null
        writeXlsxTsvClm $path $work $outDir | Should Be 3
        @(Get-ChildItem -LiteralPath $outDir | ForEach-Object { $_.Name } | Sort-Object) -join "|" |
            Should Be "売上.tsv|売上[コメント].tsv|売上[図形].tsv"
        # セルの TSV は BOM 付き UTF-8・CRLF・末尾に改行 1 つ（いつものインデクサと同じ）
        $bytes = [System.IO.File]::ReadAllBytes((Join-Path $outDir "売上.tsv"))
        @($bytes[0], $bytes[1], $bytes[2]) -join "," | Should Be "239,187,191"
        [System.Text.Encoding]::UTF8.GetString($bytes) | Should Be "$([char]0xFEFF)品名`r`n"
    }
}

Describe "Excel のセル（足りない部品・変わった書き方）" -Tag Io {
    $work4 = Join-Path $TestDrive "cells4"
    New-Item -ItemType Directory -Path $work4 | Out-Null
    $xNs4 = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $relNs4 = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
    $docRel4 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    It "ブックの情報が無い ZIP は、セルを返さない" {
        $path = Join-Path $TestDrive "nocells.xlsx"
        newTestZip $path @{ "xl/other.xml" = "<a/>" }
        (readXlsxCellTextsClm $path $work4).Count | Should Be 0
        # 書くものが無ければ 0 件
        $outDir = Join-Path $TestDrive "nocells_out"
        New-Item -ItemType Directory -Path $outDir | Out-Null
        writeXlsxTsvClm $path $work4 $outDir | Should Be 0
    }

    It "共有文字列・表示形式が無くても読め、シートが無ければ飛ばす" {
        $path = Join-Path $TestDrive "bare.xlsx"
        newTestZip $path @{
            "xl/workbook.xml" = "<workbook $xNs4><sheets>" +
                "<sheet name=`"あり`" sheetId=`"1`" r:id=`"rId1`"/><sheet name=`"無い`" sheetId=`"2`" r:id=`"rId2`"/>" +
                "</sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs4>" +
                "<Relationship Id=`"rId1`" Type=`"$docRel4/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$docRel4/worksheet`" Target=`"worksheets/無い.xml`"/></Relationships>"
            # r 属性の無いセル（左から順に並べる）と、共有文字列の番号が範囲の外のセル
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs4><sheetData><row r=`"1`"><c><v>1</v></c><c><v>2</v></c>" +
                "<c t=`"s`"><v>9</v></c><c s=`"7`"><v>3</v></c></row></sheetData></worksheet>"
        }
        $texts = readXlsxCellTextsClm $path $work4
        @($texts.Keys) -join "|" | Should Be "あり"
        $texts["あり"] | Should Be "1`t2`t`t3"
    }

    It "共有文字列の <r>（書式で分かれた文字）はつなぎ、ふりがな（rPh）は読まない" {
        $path = Join-Path $TestDrive "runs.xlsx"
        newTestZip $path @{
            "xl/workbook.xml" = "<workbook $xNs4><sheets><sheet name=`"売上`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs4><Relationship Id=`"rId1`" Type=`"$docRel4/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs4><sheetData><row r=`"1`"><c r=`"A1`" t=`"s`"><v>0</v></c></row></sheetData></worksheet>"
            "xl/sharedStrings.xml" = "<sst $xNs4><si><r><t>東京</t></r><r><t>都</t></r><rPh sb=`"0`" eb=`"2`"><t>とうきょう</t></rPh></si></sst>"
        }
        (readXlsxCellTextsClm $path $work4)["売上"] | Should Be "東京都"
    }
}
