# 制限言語モード用の Office ファイルの読み取り（shared\office\office_reader_clm.ps1）のテスト。
# いつものインデクサの読み取り（office_reader.ps1）と、同じファイルから同じ場所・同じ行を返すことを突き合わせる。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\office\office_reader.ps1"
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
