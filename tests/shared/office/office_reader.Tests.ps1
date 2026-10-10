# Pester 5 で実行: .\tests\run.ps1
# Word・PowerPointは使わず、最小限の .docx / .pptx（ZIP）をテスト内で作成して検証する
BeforeAll {
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

    function newRawZip {
        # 複数の部品にバイト列をそのまま書き込むZIPを作る（無圧縮）。部品の fakeSize を渡すと、
        # 中央ディレクトリの「展開後の大きさ」フィールドだけを書き換え、偽りのヘッダー（実際の大きさと異なる申告）を作る。
        # entries: 名前 -> @{ bytes = [byte[]]; fakeSize = [uint32]（省略可） }
        param (
            [string]$path,
            [hashtable]$entries
        )

        $stream = [System.IO.File]::Create($path)
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($name in $entries.Keys) {
            # compress = $true の部品は Deflate で圧縮する（既定は無圧縮）
            $level = if ($entries[$name].ContainsKey("compress") -and $entries[$name].compress) { [System.IO.Compression.CompressionLevel]::Optimal } else { [System.IO.Compression.CompressionLevel]::NoCompression }
            $entry = $zip.CreateEntry($name, $level)
            $entryStream = $entry.Open()
            $bytes = $entries[$name].bytes
            $entryStream.Write($bytes, 0, $bytes.Length)
            $entryStream.Dispose()
        }
        $zip.Dispose()
        $stream.Dispose()

        $fakeNames = @($entries.Keys | Where-Object { $entries[$_].ContainsKey("fakeSize") -and $null -ne $entries[$_].fakeSize })
        if ($fakeNames.Count -eq 0) {
            return
        }
        $raw = [System.IO.File]::ReadAllBytes($path)
        foreach ($name in $fakeNames) {
            $fakeSize = [uint32]$entries[$name].fakeSize
            $nameBytes = [System.Text.Encoding]::UTF8.GetBytes($name)
            for ($i = 0; $i -le $raw.Length - 46; $i++) {
                if ($raw[$i] -eq 0x50 -and $raw[$i + 1] -eq 0x4b -and $raw[$i + 2] -eq 0x01 -and $raw[$i + 3] -eq 0x02) {
                    $nameLen = [BitConverter]::ToUInt16($raw, $i + 28)
                    if ($nameLen -eq $nameBytes.Length) {
                        $match = $true
                        for ($j = 0; $j -lt $nameLen; $j++) {
                            if ($raw[$i + 46 + $j] -ne $nameBytes[$j]) { $match = $false; break }
                        }
                        if ($match) {
                            [Array]::Copy([BitConverter]::GetBytes($fakeSize), 0, $raw, $i + 24, 4)
                            break
                        }
                    }
                }
            }
        }
        [System.IO.File]::WriteAllBytes($path, $raw)
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

    function xChartFrame([string]$graphic) {
        return "<xdr:graphicFrame macro=`"`"><xdr:nvGraphicFramePr><xdr:cNvPr id=`"3`" name=`"c`"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/>${graphic}</xdr:graphicFrame>"
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
}

Describe "isZipFile" -Tag Io {
    # bytes = ファイルの中身。ZIP は先頭のシグネチャ（PK\x03\x04）だけで見分ける
    It "<name>" -TestCases @(
        @{ name = "ZIPなら `$true"; bytes = [byte[]](0x50, 0x4B, 0x03, 0x04, 0x14, 0x00); expected = $true }
        @{ name = "旧形式・パスワード付き（複合ドキュメント形式）なら `$false"; bytes = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1); expected = $false }
        @{ name = "空ファイルなら `$false"; bytes = [byte[]]@(); expected = $false }
    ) {
        param ($name, $bytes, $expected)
        $path = "$TestDrive\zip.docx"
        [System.IO.File]::WriteAllBytes($path, $bytes)
        isZipFile $path | Should -Be $expected
    }
}

Describe "isCompoundFile" -Tag Io {
    It "<name>" -TestCases @(
        @{ name = "複合ドキュメント形式（旧形式・パスワード付き）なら `$true"; bytes = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00); expected = $true }
        @{ name = "ZIPなら `$false"; bytes = [byte[]](0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08); expected = $false }
        @{ name = "テキストなら `$false"; bytes = [System.Text.Encoding]::UTF8.GetBytes("壊れたファイル"); expected = $false }
        @{ name = "空ファイルなら `$false"; bytes = [byte[]]@(); expected = $false }
    ) {
        param ($name, $bytes, $expected)
        $path = "$TestDrive\cfb.ppt"
        [System.IO.File]::WriteAllBytes($path, $bytes)
        isCompoundFile $path | Should -Be $expected
    }
}

Describe "resolveZipPath" -Tag Io {
    It "相対パスを解決する" {
        resolveZipPath "ppt/slides" "../notesSlides/notesSlide1.xml" | Should -Be "ppt/notesSlides/notesSlide1.xml"
        resolveZipPath "ppt" "slides/slide1.xml" | Should -Be "ppt/slides/slide1.xml"
    }

    It "/ で始まるパスはZIPのルートからとする" {
        resolveZipPath "ppt/slides" "/ppt/media/a.png" | Should -Be "ppt/media/a.png"
    }
}

Describe "readDocxUnits" -Tag Io {
    BeforeAll {
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
    }

    It "ページ・ヘッダー/フッター・脚注の順に分ける" {
        @($units.Keys) -join "|" | Should -Be "ページ001|ページ002|ページ003|ページ004|ページ002[図形]|ヘッダー・フッター|脚注"
    }

    It "手動の改ページは、直後の保存時のページ区切りの有無にかかわらず1ページとして数える" {
        @($units["ページ003"]) -join "|" | Should -Be "3ページ目"
        @($units["ページ004"]) -join "|" | Should -Be "4ページ目"
    }

    It "段落と表の行を1行ずつにする" {
        @($units["ページ001"]) -join "|" | Should -Be "見出し|本文  りんご|品名`t価格|複数 段落"
    }

    It "保存時のページ区切りで次のページにし、テキストボックスは本文から分けて、そのページの図形として1回だけ読む" {
        @($units["ページ002"]) -join "|" | Should -Be "2ページ目"
        @($units["ページ002[図形]"]) -join "|" | Should -Be "テキストボックス"
    }

    It "ヘッダー・フッターの重複を除く" {
        @($units["ヘッダー・フッター"]) -join "|" | Should -Be "社外秘|フッター"
    }

    It "脚注の区切り線は読まない" {
        @($units["脚注"]) -join "|" | Should -Be "脚注本文"
    }
}

Describe "readDocxUnits（保存時のページ区切りが無い文書）" -Tag Io {
    BeforeAll {
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
    }

    It "手動の改ページ・セクション区切り（連続以外）・段落前で改ページ でページを数える" {
        @($units.Keys) -join "|" | Should -Be "ページ001|ページ002|ページ003|ページ004"
        @($units["ページ001"]) -join "|" | Should -Be "A"
        @($units["ページ002"]) -join "|" | Should -Be "B|C|D"
        @($units["ページ003"]) -join "|" | Should -Be "E"
        @($units["ページ004"]) -join "|" | Should -Be "F"
    }
}

Describe "readPptxUnits" -Tag Io {
    BeforeAll {
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
    }

    It "スライドの表示順に番号を付け、非表示スライド・ノート・フッターを分ける" {
        @($units.Keys) -join "|" | Should -Be "スライド001|スライド002（非表示）|スライド002_ノート|ヘッダー・フッター"
    }

    It "図形・グループ内の図形・表を読み、スライド番号・フッターは読まない" {
        @($units["スライド001"]) -join "|" | Should -Be "表紙|グループ内|項目`t値"
        @($units["スライド002（非表示）"]) -join "|" | Should -Be "非表示の内容"
    }

    It "スライドのフッターは重複を除いてまとめ、日付は読まない" {
        @($units["ヘッダー・フッター"]) -join "|" | Should -Be "社外秘"
    }

    It "ノートの本文を読み、スライド番号は読まない" {
        @($units["スライド002_ノート"]) -join "|" | Should -Be "ノート本文"
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
        writeUnits $units $outDir | Should -Be 2
        [System.IO.File]::ReadAllText("$outDir\page_001.tsv") | Should -Be "a`tb`r`n"
        Test-Path -LiteralPath "$outDir\page_002.tsv" | Should -Be $false
        # 固定名に当てはまる場所は英語の固定名にする（toIndexFileName）
        Test-Path -LiteralPath "$outDir\slide_001_notes.tsv" | Should -Be $true
    }

    It "testdata の docx・pptx すべてで、出力するファイル名がASCIIになる（再発防止。場所や種類を足して英語の名前を忘れると、ここで落ちる）" {
        $root = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\..\testdata\office")
        $files = @(Get-ChildItem -Path $root -Recurse -File -Include "*.docx", "*.pptx")
        $files.Count | Should -BeGreaterThan 0

        $checked = 0
        foreach ($file in $files) {
            $units = $null
            try {
                $units = $(if ($file.Extension -ieq ".docx") { readDocxUnits $file.FullName } else { readPptxUnits $file.FullName })
            } catch {
                continue
            }
            if ($null -eq $units -or $units.Count -eq 0) {
                continue
            }

            $sweepDir = "$TestDrive\ascii_sweep\$([Guid]::NewGuid().ToString('N'))"
            [System.IO.Directory]::CreateDirectory($sweepDir) | Out-Null
            writeUnits $units $sweepDir | Out-Null
            foreach ($written in Get-ChildItem -Path $sweepDir -File) {
                $written.Name | Should -Match "^[\x20-\x7E]+$"
            }
            $checked++
        }

        $checked | Should -BeGreaterThan 0
    }
}

Describe "readXlsxObjectUnits" -Tag Io {
    BeforeAll {
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
    }

    It "表示シートの図形・コメントだけを、シートごとの場所にする（非表示シート・グラフシート・何も無いシートは出さない）" {
        @($units.Keys) -join "|" | Should -Be "売上[図形]|売上[コメント]"
    }

    It "図形は左上のセル番地と文字を並べ、上の行から（同じ行は左から）の順にする。文字の無い図形（グラフなど）は出さない" {
        $lines = @($units["売上[図形]"])
        $lines[0] | Should -Be "C2`t`"引用`"`"あり`""
        $lines[1] | Should -Be "F2`t`"1段落目${ls}2段落目`""
        $lines[2] | Should -Be "B5`t下の図形"
        $lines.Count | Should -Be 5
    }

    It "グループ化した図形はまとめて 1 つにし、互換用の代替表示（mc:Fallback）は読まない" {
        $lines = @($units["売上[図形]"])
        $lines[3] | Should -Be "A10`t`"グループ1${ls}グループ2`""
        $lines[4] | Should -Be "A21`t代替表示あり"
    }

    It "コメントはセルの順に並べ、ふりがなは読まない。スレッド形式のコメントがあるセルは、その文字（返信を含む）を使う" {
        @($units["売上[コメント]"]) -join "|" |
            Should -Be "C1`t`"確認してください${ls}確認しました`"|A2`tメモ|B3`t`"test:${ls}価格は税抜`""
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
        @((readXlsxObjectUnits $same)["S[図形]"]) -join "|" | Should -Be ((1..30 | ForEach-Object { "B2`t図形$_" }) -join "|")
    }

    It "Excel のブックでない ZIP（.xlsb など）は何も返さない" {
        $xlsb = "$TestDrive\binary.xlsb"
        newZip $xlsb @{ "xl/workbook.bin" = "binary" }
        (readXlsxObjectUnits $xlsb).Count | Should -Be 0
    }
}

Describe "readXlsxObjectUnits（グラフ・SmartArt・グラフシート）" -Tag Io {
    BeforeAll {
        function xAbsoluteAnchor([string]$shapes) {
            # 位置をセルで持たない図形（グラフシートに置いたグラフなど）
            return "<xdr:absoluteAnchor><xdr:pos x=`"0`" y=`"0`"/><xdr:ext cx=`"100`" cy=`"100`"/>${shapes}<xdr:clientData/></xdr:absoluteAnchor>"
        }

        $groupChart = "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=`"9`" name=`"g2`"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>" +
            (xSp @("グループのテキスト")) + (xChartFrame (chartRef 'rId22')) + "</xdr:grpSp>"

        $chart2Xml = "<c:chartSpace $cNs><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>グループ内グラフ</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
            "<c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>"
        $chart3Xml = "<c:chartSpace $cNs><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>TCg グラフシートのタイトル</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
            "<c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>"

        $path = "$TestDrive\charts.xlsx"
        newZip $path @{
            "xl/workbook.xml" = "<workbook $xNs><sheets>" +
                "<sheet name=`"S`" sheetId=`"1`" r:id=`"rId1`"/>" +
                "<sheet name=`"隠しグラフ`" sheetId=`"2`" state=`"hidden`" r:id=`"rId2`"/>" +
                "<sheet name=`"グラフ2ページ`" sheetId=`"3`" r:id=`"rId3`"/>" +
                "<sheet name=`"隠しグラフシート`" sheetId=`"4`" state=`"hidden`" r:id=`"rId4`"/>" +
                "<sheet name=`"壊れたrels`" sheetId=`"5`" r:id=`"rId5`"/>" +
                "</sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet2.xml`"/>" +
                "<Relationship Id=`"rId3`" Type=`"$officeRel/chartsheet`" Target=`"chartsheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId4`" Type=`"$officeRel/chartsheet`" Target=`"chartsheets/sheet2.xml`"/>" +
                "<Relationship Id=`"rId5`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet3.xml`"/>" +
                "</Relationships>"

            "xl/worksheets/sheet1.xml" = "<worksheet $xNs/>"
            "xl/worksheets/_rels/sheet1.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing1.xml`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$officeRel/comments`" Target=`"../comments1.xml`"/>" +
                "</Relationships>"
            "xl/comments1.xml" = "<comments $xNs><authors><author>test</author></authors><commentList>" +
                "<comment ref=`"D1`" authorId=`"0`"><text><t>コメントは影響を受けない</t></text></comment>" +
                "</commentList></comments>"
            # A2: グラフ単体、A6: SmartArt 単体、C9: グループ（テキストボックス→グラフの順）、
            # A13: 参照先（リレーションシップ）が無いグラフ + テキストボックス、
            # A16: 部品（XML）が壊れたグラフ + テキストボックス、A19: 通常の図形（グラフ・SmartArt が読めなくても出る）
            "xl/drawings/drawing1.xml" = "<xdr:wsDr $xdrNs>" +
                "$(xAnchor 0 1 (xChartFrame (chartRef 'rId20')))" +
                "$(xAnchor 0 5 (xChartFrame (smartArt 'rId21')))" +
                "$(xAnchor 2 8 $groupChart)" +
                "$(xAnchor 0 12 ((xSp @('テキストボックス1')) + (xChartFrame (chartRef 'rIdMissing'))))" +
                "$(xAnchor 0 15 ((xSp @('テキストボックス2')) + (xChartFrame (chartRef 'rId23'))))" +
                "$(xAnchor 0 18 (xSp @('通常の図形')))" +
                "</xdr:wsDr>"
            "xl/drawings/_rels/drawing1.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId20`" Type=`"$docRel/chart`" Target=`"../charts/chart1.xml`"/>" +
                "<Relationship Id=`"rId21`" Type=`"$docRel/diagramData`" Target=`"../diagrams/data1.xml`"/>" +
                "<Relationship Id=`"rId22`" Type=`"$docRel/chart`" Target=`"../charts/chart2.xml`"/>" +
                "<Relationship Id=`"rId23`" Type=`"$docRel/chart`" Target=`"../charts/broken.xml`"/>" +
                "</Relationships>"
            "xl/charts/chart1.xml" = $chartXml
            "xl/charts/chart2.xml" = $chart2Xml
            "xl/charts/broken.xml" = "<c:chartSpace $cNs><c:chart>"  # 閉じタグが無い壊れたXML
            "xl/diagrams/data1.xml" = $diagramXml

            "xl/worksheets/sheet2.xml" = "<worksheet $xNs/>"
            "xl/worksheets/_rels/sheet2.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing2.xml`"/></Relationships>"
            "xl/drawings/drawing2.xml" = "<xdr:wsDr $xdrNs>$(xAnchor 0 0 (xChartFrame (chartRef 'rId1')))</xdr:wsDr>"
            "xl/drawings/_rels/drawing2.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$docRel/chart`" Target=`"../charts/chart1.xml`"/></Relationships>"

            "xl/chartsheets/sheet1.xml" = "<chartsheet $xNs/>"
            "xl/chartsheets/_rels/sheet1.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing3.xml`"/></Relationships>"
            "xl/drawings/drawing3.xml" = "<xdr:wsDr $xdrNs>$(xAbsoluteAnchor (xChartFrame (chartRef 'rId1')))</xdr:wsDr>"
            "xl/drawings/_rels/drawing3.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$docRel/chart`" Target=`"../charts/chart3.xml`"/></Relationships>"
            "xl/charts/chart3.xml" = $chart3Xml

            # 図形の部品自身のリレーションシップ（drawingN.xml.rels）の XML が壊れている場合
            "xl/worksheets/sheet3.xml" = "<worksheet $xNs/>"
            "xl/worksheets/_rels/sheet3.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing4.xml`"/></Relationships>"
            "xl/drawings/drawing4.xml" = "<xdr:wsDr $xdrNs>" +
                "$(xAnchor 0 1 (xChartFrame (chartRef 'rId1')))" +
                "$(xAnchor 0 5 (xSp @('壊れたrelsでも出る図形')))" +
                "</xdr:wsDr>"
            "xl/drawings/_rels/drawing4.xml.rels" = "<Relationships $relNs><Relationship"  # 閉じタグの無い壊れたXML
        }

        $failures = New-Object System.Collections.Generic.List[string]
        $units = readXlsxObjectUnits $path $failures
        $ls = [char]0x2028
    }

    It "グラフはタイトル・系列名を読み、項目名・数値は読まない（1 アンカー 1 行）" {
        @($units["S[図形]"])[0] | Should -Be "A2`t月別売上 東京支店 大阪支店"
    }

    It "SmartArt は data の文字を読む（drawing の文字とは重ならない）" {
        @($units["S[図形]"])[1] | Should -Be "A6`t企画 設計"
    }

    It "グループの中はテキストボックスの段落 → グラフ・SmartArt の文字の順で 1 行にする" {
        @($units["S[図形]"])[2] | Should -Be "C9`t`"グループのテキスト${ls}グループ内グラフ`""
    }

    It "参照先・リレーションシップが無いグラフは空にし、同じアンカーのテキストボックスは出す" {
        @($units["S[図形]"])[3] | Should -Be "A13`tテキストボックス1"
    }

    It "部品（XML）が壊れたグラフは空にし、同じアンカーのテキストボックスは出す。読めなかった部品を $failures に返す" {
        @($units["S[図形]"])[4] | Should -Be "A16`tテキストボックス2"
        # $failures はこのフィクスチャー全体（シート S と 壊れたrels）で読めなかった部品を集めたもの。
        # 期待する失敗の一覧を全部並べて比べる（一部だけの確かめだと、ほかの失敗が紛れ込んでも気付けない）
        @($failures) | Should -Be @("xl/charts/broken.xml", "xl/drawings/drawing4.xml")
    }

    It "グラフ・SmartArt が読めなくても、同じシートのほかの図形・コメントは出る" {
        @($units["S[図形]"])[5] | Should -Be "A19`t通常の図形"
        @($units["S[コメント]"]) -join "|" | Should -Be "D1`tコメントは影響を受けない"
    }

    It "非表示シートのグラフは出ない" {
        $units.Contains("隠しグラフ[図形]") | Should -Be $false
    }

    It "表示のグラフシートは、位置をセルで持たない（absoluteAnchor）ため A1 になる" {
        @($units["グラフ2ページ[図形]"]) -join "|" | Should -Be "A1`tTCg グラフシートのタイトル"
    }

    It "非表示のグラフシートは出ない" {
        $units.Contains("隠しグラフシート[図形]") | Should -Be $false
    }

    It "図形の部品自身のリレーションシップ（drawingN.xml.rels）が壊れていても、そのグラフだけを空にし、ほかの図形は出す" {
        @($units["壊れたrels[図形]"]) -join "|" | Should -Be "A6`t壊れたrelsでも出る図形"
        @($failures) | Should -Be @("xl/charts/broken.xml", "xl/drawings/drawing4.xml")
    }
}

Describe "getHeaderFooterLines" -Tag Unit {
    It "<Name>" -TestCases @(
        @{ Name = "左・中央・右を、この順の 3 行にする"; Code = '&L左&C中&R右'; Expected = "左|中|右" }
        @{ Name = "区切りが無いときは中央 1 行"; Code = '中央だけ'; Expected = "中央だけ" }
        @{ Name = "&& は & 1 文字"; Code = '&L会社&&部'; Expected = "会社&部" }
        @{ Name = "ページ番号・総ページ・ページの加減を取り除く"; Code = '&C&P / &N ページ'; Expected = "/  ページ" }
        @{ Name = "&P+1 の数字も取り除く"; Code = '&C&P+1 &P-2ページ'; Expected = "ページ" }
        @{ Name = "日付・時刻・パス・ファイル名・シート名・画像は文字にしない"; Code = '&L&D &T&C&Z&F&R&A&G'; Expected = "" }
        @{ Name = "フォントと文字の大きさを取り除く"; Code = '&C&"ＭＳ ゴシック,太字"&12見出し'; Expected = "見出し" }
        @{ Name = "大きさの後のスペースは、前後の空白として落とす"; Code = '&C&11 2024年度'; Expected = "2024年度" }
        @{ Name = "色（16 進 6 桁・テーマ色）を取り除く"; Code = '&C&KFF0000赤&K04+000青'; Expected = "赤青" }
        @{ Name = "太字・斜体・下線などを取り除く"; Code = '&C&B&I&U&E&S&X&Y&O&H太字'; Expected = "太字" }
        @{ Name = "知らない & の続きはそのまま残す"; Code = '&C&Q残る'; Expected = "&Q残る" }
        @{ Name = "末尾の & は残す"; Code = '&C末尾&'; Expected = "末尾&" }
        @{ Name = "閉じていないフォントの指定は残す"; Code = '&C&"壊れ'; Expected = '"&""壊れ"' }
        @{ Name = "空の区分は出さない"; Code = '&L&C&R右だけ'; Expected = "右だけ" }
        @{ Name = "区分の中の改行は別の行にする"; Code = "&C1行目`n2行目`r`n3行目"; Expected = "1行目|2行目|3行目" }
        @{ Name = "タブはスペースにし、前後の空白を落とす"; Code = "&C  会社`t名  "; Expected = "会社 名" }
        @{ Name = "`" を含む行は toObjectCellText の形（囲んで `"`" にする）"; Code = '&C"至急" と "確認"'; Expected = '"""至急"" と ""確認"""' }
    ) {
        (@(getHeaderFooterLines $Code) -join "|") | Should -BeExactly $Expected
    }
}

Describe "readXlsxObjectUnits（ヘッダー・フッター）" -Tag Io {
    BeforeAll {
        function hf([string]$body, [string]$attributes = "") {
            # ヘッダー・フッターの XML。文字は & を含むためエスケープする（子の要素は名前 → 書式コード付きの文字）
            $children = ""
            foreach ($name in @("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter")) {
                $m = [regex]::Match($body, "(?:^|\|)${name}=([^|]*)")
                if ($m.Success) { $children += "<$name>$([System.Security.SecurityElement]::Escape($m.Groups[1].Value))</$name>" }
            }
            return "<headerFooter${attributes}>${children}</headerFooter>"
        }

        $manyRows = (1..300 | ForEach-Object { "<row r=`"$_`"><c r=`"A$_`" t=`"inlineStr`"><is><t>行$_</t></is></c></row>" }) -join ""
        $sales = "<worksheet $xNs><sheetData>$manyRows</sheetData>" +
            (hf 'oddHeader=&L社外秘&C&"ＭＳ ゴシック,太字"&12見出し&R会議|oddFooter=&C&P / &N ページ|evenHeader=&C偶数ヘッダー|evenFooter=&C偶数の注記|firstHeader=&C表紙の見出し|firstFooter=&C社外秘' ' differentOddEven="1" differentFirst="true"') +
            "<drawing r:id=`"rId1`"/></worksheet>"
        $plain = "<worksheet $xNs><sheetData/>" +
            (hf 'oddHeader=&C通常のみ|firstHeader=&C読まない先頭|evenFooter=&C読まない偶数') + "</worksheet>"
        $prefixed = '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/>' +
            '<x:headerFooter><x:oddHeader>&amp;C接頭辞付き</x:oddHeader></x:headerFooter></x:worksheet>'

        $path = "$TestDrive\headerfooter.xlsx"
        newZip $path @{
            "xl/workbook.xml" = "<workbook $xNs><sheets>" +
                "<sheet name=`"売上`" sheetId=`"1`" r:id=`"rId1`"/>" +
                "<sheet name=`"通常`" sheetId=`"2`" r:id=`"rId2`"/>" +
                "<sheet name=`"隠し`" sheetId=`"3`" state=`"hidden`" r:id=`"rId3`"/>" +
                "<sheet name=`"グラフ`" sheetId=`"4`" r:id=`"rId4`"/>" +
                "<sheet name=`"なし`" sheetId=`"5`" r:id=`"rId5`"/>" +
                "<sheet name=`"空`" sheetId=`"6`" r:id=`"rId6`"/>" +
                "<sheet name=`"接頭辞`" sheetId=`"7`" r:id=`"rId7`"/>" +
                "<sheet name=`"引用`" sheetId=`"8`" r:id=`"rId8`"/>" +
                "<sheet name=`"隠しグラフ`" sheetId=`"9`" state=`"hidden`" r:id=`"rId9`"/>" +
                "</sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet2.xml`"/>" +
                "<Relationship Id=`"rId3`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet3.xml`"/>" +
                "<Relationship Id=`"rId4`" Type=`"$officeRel/chartsheet`" Target=`"chartsheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId5`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet4.xml`"/>" +
                "<Relationship Id=`"rId6`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet5.xml`"/>" +
                "<Relationship Id=`"rId7`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet6.xml`"/>" +
                "<Relationship Id=`"rId8`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet7.xml`"/>" +
                "<Relationship Id=`"rId9`" Type=`"$officeRel/chartsheet`" Target=`"chartsheets/sheet2.xml`"/>" +
                "</Relationships>"
            "xl/worksheets/sheet1.xml" = $sales
            "xl/worksheets/_rels/sheet1.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing1.xml`"/></Relationships>"
            "xl/drawings/drawing1.xml" = "<xdr:wsDr $xdrNs>$(xAnchor 0 0 (xSp @('図形の文字')))</xdr:wsDr>"
            "xl/worksheets/sheet2.xml" = $plain
            "xl/worksheets/sheet3.xml" = "<worksheet $xNs><sheetData/>$(hf 'oddHeader=&C隠しシートのヘッダー')</worksheet>"
            "xl/chartsheets/sheet1.xml" = "<chartsheet $xNs><sheetViews/>$(hf 'oddHeader=&Cグラフの見出し|oddFooter=&R&Pページ')</chartsheet>"
            "xl/chartsheets/sheet2.xml" = "<chartsheet $xNs>$(hf 'oddHeader=&C隠しグラフシートのヘッダー')</chartsheet>"
            "xl/worksheets/sheet4.xml" = "<worksheet $xNs><sheetData/></worksheet>"
            "xl/worksheets/sheet5.xml" = "<worksheet $xNs><sheetData/><headerFooter/></worksheet>"
            "xl/worksheets/sheet6.xml" = $prefixed
            "xl/worksheets/sheet7.xml" = "<worksheet $xNs><sheetData/>$(hf 'oddHeader=&C"至急" と "確認"')</worksheet>"
        }
        $failures = New-Object System.Collections.Generic.List[string]
        $units = readXlsxObjectUnits $path $failures
    }

    It "表示のワークシート・グラフシートの <シート名>[ヘッダー・フッター] を、図形・コメントの後に返す（非表示シート・ヘッダーの無いシートは出さない）" {
        @($units.Keys) -join "|" | Should -Be "売上[図形]|売上[ヘッダー・フッター]|通常[ヘッダー・フッター]|グラフ[ヘッダー・フッター]|接頭辞[ヘッダー・フッター]|引用[ヘッダー・フッター]"
        @($failures).Count | Should -Be 0
    }

    It "sheetData（多数の行）の後にある headerFooter を、ヘッダー → フッター、先頭 → 奇数 → 偶数、左 → 中央 → 右の順に読む。書式コードは文字にせず、同じ文字は 1 回だけ" {
        @($units["売上[ヘッダー・フッター]"]) -join "|" | Should -Be "表紙の見出し|社外秘|見出し|会議|偶数ヘッダー|/  ページ|偶数の注記"
    }

    It "differentFirst・differentOddEven が無いシートは、先頭・偶数ページのものを読まない" {
        @($units["通常[ヘッダー・フッター]"]) -join "|" | Should -Be "通常のみ"
    }

    It "表示のグラフシートのヘッダー・フッターを読む" {
        @($units["グラフ[ヘッダー・フッター]"]) -join "|" | Should -Be "グラフの見出し|ページ"
    }

    It "接頭辞付き（x:headerFooter）で書いたシートも読む" {
        @($units["接頭辞[ヘッダー・フッター]"]) -join "|" | Should -Be "接頭辞付き"
    }

    It "`" を含む行は、Excel のテキスト保存と同じ形（囲んで `"`" にする）で返す" {
        @($units["引用[ヘッダー・フッター]"]) -join "|" | Should -Be '"""至急"" と ""確認"""'
    }

    It "XML が壊れたシートのヘッダー・フッターだけを飛ばし、ほかのシートの図形・ヘッダー・フッターは返す。読めなかった部品を `$failures に返す" {
        $broken = "$TestDrive\broken_hf.xlsx"
        newZip $broken @{
            "xl/workbook.xml" = "<workbook $xNs><sheets><sheet name=`"壊れ`" sheetId=`"1`" r:id=`"rId1`"/><sheet name=`"正常`" sheetId=`"2`" r:id=`"rId2`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet2.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs><sheetData/><headerFooter><oddHeader>&amp;C途中まで"  # 閉じタグの無い壊れた XML
            "xl/worksheets/sheet2.xml" = "<worksheet $xNs><sheetData/>$(hf 'oddHeader=&C正常なヘッダー')</worksheet>"
        }
        $fails = New-Object System.Collections.Generic.List[string]
        $result = readXlsxObjectUnits $broken $fails
        @($result.Keys) -join "|" | Should -Be "正常[ヘッダー・フッター]"
        @($result["正常[ヘッダー・フッター]"]) -join "|" | Should -Be "正常なヘッダー"
        @($fails) | Should -Be @("xl/worksheets/sheet1.xml")
    }

    It "DTD を含むシートは読まず（外部のものを読み込まない）、部品として読めなかったことにする" {
        $dtd = "$TestDrive\dtd_hf.xlsx"
        newZip $dtd @{
            "xl/workbook.xml" = "<workbook $xNs><sheets><sheet name=`"S`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<!DOCTYPE worksheet [<!ENTITY x `"x`">]><worksheet $xNs><sheetData/><headerFooter><oddHeader>&amp;C&x;</oddHeader></headerFooter></worksheet>"
        }
        $fails = New-Object System.Collections.Generic.List[string]
        (readXlsxObjectUnits $dtd $fails).Count | Should -Be 0
        @($fails) | Should -Be @("xl/worksheets/sheet1.xml")
    }

    It "部品の上限を超えて大きいシートは、流れで読んでも上限で打ち切り、そのシートのヘッダー・フッターだけ読めなかったことにする" {
        $big = "$TestDrive\big_hf.xlsx"
        $rows = "<row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>$('あ' * 3000)</t></is></c></row>"
        newZip $big @{
            "xl/workbook.xml" = "<workbook $xNs><sheets><sheet name=`"大`" sheetId=`"1`" r:id=`"rId1`"/><sheet name=`"小`" sheetId=`"2`" r:id=`"rId2`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet2.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs><sheetData>$rows</sheetData>$(hf 'oddHeader=&C大きいシート')</worksheet>"
            "xl/worksheets/sheet2.xml" = "<worksheet $xNs><sheetData/>$(hf 'oddHeader=&C小さいシート')</worksheet>"
        }
        $orig = $script:zipSheetStreamMaxChars
        $script:zipSheetStreamMaxChars = 2000
        try {
            $script:zipTotalReadBytes = 0
            $fails = New-Object System.Collections.Generic.List[string]
            $result = readXlsxObjectUnits $big $fails
            @($result.Keys) -join "|" | Should -Be "小[ヘッダー・フッター]"
            @($fails).Count | Should -Be 1
            @($fails)[0] | Should -BeLike "xl/worksheets/sheet1.xml*流れ読みの上限*"
        } finally {
            $script:zipSheetStreamMaxChars = $orig
        }
    }

    It "流れ読みの上限は部品ごとの上限とは別で、部品の上限を超えるシートでもヘッダー・フッターを読める。申告より実際が大きいシートも上限で打ち切る" {
        $path = "$TestDrive"+[char]92+"stream_limit.xlsx"
        $rows = "<row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>$('あ' * 3000)</t></is></c></row>"
        newZip $path @{
            "xl/workbook.xml" = "<workbook $xNs><sheets><sheet name=`"大`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs><sheetData>$rows</sheetData>$(hf 'oddHeader=&C大きいシート')</worksheet>"
        }
        $origPart = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = 2000
        try {
            $script:zipTotalReadBytes = 0
            $fails = New-Object System.Collections.Generic.List[string]
            $result = readXlsxObjectUnits $path $fails
            @($result["大[ヘッダー・フッター]"]) -join "|" | Should -Be "大きいシート"
            @($fails).Count | Should -Be 0
        } finally {
            $script:zipPartMaxBytes = $origPart
        }
        # 申告（entry.Length）は上限内でも、読んだ文字数が上限を超えたら打ち切る
        $origChars = $script:zipSheetStreamMaxChars
        $script:zipSheetStreamMaxChars = 5000
        try {
            $script:zipTotalReadBytes = 0
            $fails = New-Object System.Collections.Generic.List[string]
            $null = readXlsxObjectUnits $path $fails
            @($fails).Count | Should -Be 1
        } finally {
            $script:zipSheetStreamMaxChars = $origChars
        }
    }

    It "シートに関係のない大きな部品（背景の画像など）がつながっていても、図形・コメントは読める" {
        $path = "$TestDrive\big_unrelated.xlsx"
        newZip $path @{
            "xl/workbook.xml" = "<workbook $xNs><sheets><sheet name=`"売上`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs/>"
            "xl/worksheets/_rels/sheet1.xml.rels" = "<Relationships $relNs>" +
                "<Relationship Id=`"rId1`" Type=`"$officeRel/image`" Target=`"../media/bg.bin`"/>" +
                "<Relationship Id=`"rId2`" Type=`"$officeRel/drawing`" Target=`"../drawings/drawing1.xml`"/>" +
                "<Relationship Id=`"rId3`" Type=`"$officeRel/comments`" Target=`"../comments1.xml`"/></Relationships>"
            "xl/media/bg.bin" = ('x' * 5000)
            "xl/drawings/drawing1.xml" = "<xdr:wsDr $xdrNs>$(xAnchor 0 0 (xSp @('図形あり')))</xdr:wsDr>"
            "xl/comments1.xml" = "<comments $xNs><authors><author>test</author></authors><commentList><comment ref=`"A2`" authorId=`"0`"><text><t>メモあり</t></text></comment></commentList></comments>"
        }
        $orig = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = 2000
        try {
            $script:zipTotalReadBytes = 0
            $result = readXlsxObjectUnits $path
            @($result["売上[図形]"]) -join "|" | Should -Be "A1`t図形あり"
            @($result["売上[コメント]"]) -join "|" | Should -Be "A2`tメモあり"
        } finally {
            $script:zipPartMaxBytes = $orig
        }
    }
}

Describe "readXlsxObjectUnits（実物のブック・ヘッダー）" -Tag Io {
    It "Excel で付けたヘッダー・フッター（オブジェクトのシート。左・中央・右、ページ番号の差し込みを含む）を読む" {
        $path = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\..\testdata\office\Excel\セル内容.xlsx")
        $units = readXlsxObjectUnits $path
        @($units["オブジェクト[ヘッダー・フッター]"]) -join "|" | Should -Be "TC10-O05 左ヘッダー|TC10 ヘッダーのテキスト|TC10-O05 右ヘッダー|TC10-O05 左フッター|TC10-O05 通常フッター|/  ページ"
    }
}

Describe "readXlsxObjectUnits（実物のブック）" -Tag Io {
    # Excel で作った実物のブック（tests/testdata/README.md「グラフ・SmartArt」）。
    # readXlsxObjectUnits 自体は Excel を使わないため、Io のタグで CI でも流れる
    BeforeAll {
        $path = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\..\testdata\office\Excel\グラフとSmartArt.xlsx")
        $units = readXlsxObjectUnits $path
        $ls = [char]0x2028
    }

    It "埋め込みグラフのタイトル・軸ラベル・系列名を読み、項目名・数値（987654）は読まない" {
        @($units["表[図形]"])[0] | Should -Be "D3`tTC30 グラフのタイトル TC30 横軸 TC30 縦軸 TC30 系列名"
        (@($units["表[図形]"]) -join "|") | Should -Not -Match "987654|TC30 項目"
    }

    It "SmartArt の文字を読む" {
        @($units["表[図形]"])[1] | Should -Be "H3`tTC30 SmartArt のテキスト"
    }

    It "グループの中はテキストボックス → グラフの順で 1 行になる" {
        @($units["表[図形]"])[2] | Should -Be "A17`t`"TC30 グループ内のテキストボックス${ls}TC30 グループ内グラフ TC30 系列名`""
    }

    It "非表示シートに置いたグラフは読まない" {
        $units.Contains("非表示グラフ[図形]") | Should -Be $false
    }

    It "表示のグラフシートは A1 で読む" {
        @($units["TC30グラフシート[図形]"]) -join "|" | Should -Be "A1`tTC30 グラフシートのタイトル TC30 系列名"
    }
}

Describe "readDocxUnits（図形・コメント）" -Tag Io {
    BeforeAll {
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
    }

    It "テキストボックスの文字は本文から分け、段落をスペースでつないで図形 1 つにする（代替表示は読まない）" {
        @($units["ページ001"]) -join "|" | Should -Be "1ページ目の本文|表のセル`t右のセル|コメントを付けた段落"
        @($units["ページ001[図形]"]) -join "|" | Should -Be "箱の1段落目 箱の2段落目|セルの中の箱|箱の表A`t箱の表B"
    }

    It "SmartArt・グラフの文字は、そのページの図形にする（グラフはタイトル・系列名。項目名・数値は読まず、重複も読まない）" {
        @($units["ページ002[図形]"]) -join "|" | Should -Be "企画 設計|月別売上 東京支店 大阪支店"
    }

    It "コメントは付けた所のページに、返信も 1 件ずつ入れる。本文に参照の無いコメントは「文書」にまとめる" {
        @($units["ページ001[コメント]"]) -join "|" | Should -Be "確認してください 2段落目|返信です"
        @($units["ページ002[コメント]"]) -join "|" | Should -Be "グラフの数字を確認"
        @($units["文書[コメント]"]) -join "|" | Should -Be "参照の無いコメント"
    }
}

Describe "readPptxUnits（図形・コメント）" -Tag Io {
    BeforeAll {
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
    }

    It "テキストボックス・図形はスライドの本文のまま、SmartArt・グラフの文字はスライドの図形にする" {
        @($units["スライド001"]) -join "|" | Should -Be "スライドの本文"
        @($units["スライド001[図形]"]) -join "|" | Should -Be "企画 設計|月別売上 東京支店 大阪支店"
    }

    It "旧形式・新形式のコメントを読み、コメントの後に返信を 1 件ずつ入れる（非表示のスライドは場所の名前に付く）" {
        @($units["スライド001[コメント]"]) -join "|" | Should -Be "旧形式のコメント"
        @($units["スライド002（非表示）[コメント]"]) -join "|" | Should -Be "新形式の コメント|返信です"
    }
}

Describe "readDocxUnits（表・変更履歴・書式・改行の細かい扱い）" -Tag Io {
    BeforeAll {
        $body = @(
            # 入れ子の表は外側のセルの中でスペース区切り。空のセル（<w:tc/>）も列として数え、後ろの列がずれない
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>外</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>内1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>内2</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc><w:tc/><w:tc><w:p><w:r><w:t>右</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
            '<w:tbl><w:tr><w:tc/><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
            # 文字の書式は読み飛ばしても文字は残す。変更履歴の移動元は読まない（移動先と重複する）
            '<w:p><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>太字</w:t></w:r><w:moveFrom><w:r><w:t>移動元</w:t></w:r></w:moveFrom><w:moveTo><w:r><w:t>移動先</w:t></w:r></w:moveTo></w:p>',
            # 行内の改行・改行文字はスペース、改行しないハイフンは -
            '<w:p><w:r><w:t>改行</w:t><w:br/><w:t>後</w:t><w:cr/><w:t>ハイフン</w:t><w:noBreakHyphen/><w:t>X</w:t></w:r></w:p>'
        ) -join ""
        $path = "$TestDrive\details.docx"
        newZip $path @{ "word/document.xml" = "<w:document $wNs><w:body>$body</w:body></w:document>" }
        $units = readDocxUnits $path
    }

    It "入れ子の表・空のセル・書式・移動・行内の改行を扱う" {
        @($units["ページ001"]) -join "|" | Should -Be "外 内1 内2`t`t右|`tB|太字移動先|改行 後 ハイフン-X"
    }
}

Describe "readDocxUnits / readPptxUnits（必要な部品が無い ZIP）" -Tag Io {
    It "Word の本文が無ければ、分かるメッセージで例外にする" {
        $path = "$TestDrive\本文なし.docx"
        newZip $path @{ "word/other.xml" = "x" }
        { readDocxUnits $path } | Should -Throw -ExpectedMessage "*word/document.xml*"
    }

    It "PowerPoint のプレゼンテーション情報が無ければ、分かるメッセージで例外にする" {
        $path = "$TestDrive\情報なし.pptx"
        newZip $path @{ "ppt/other.xml" = "x" }
        { readPptxUnits $path } | Should -Throw -ExpectedMessage "*ppt/presentation.xml*"
    }
}

Describe "readObjectText / readChartText" -Tag Io {
    BeforeAll {
        $path = "$TestDrive\objects_missing.docx"
        newZip $path @{ "word/document.xml" = "<w:document $wNs/>" }
    }

    It "参照（RelId）が無い・参照先のファイルが無ければ空" {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $rels = @{ rId1 = @{ Type = "$docRel/chart"; Target = "word/charts/無い.xml" } }
            readObjectText $zip $rels @{ Kind = "chart"; RelId = "rId9" } | Should -Be ""
            readObjectText $zip $rels @{ Kind = "chart"; RelId = "rId1" } | Should -Be ""
        } finally {
            $zip.Dispose()
        }
    }

    It "項目名（c:cat。多段の multiLvlStrCache を含む）は読まない。系列名（c:ser/c:tx）は読み、数値も読まない" {
        $xml = "<c:chartSpace $cNs><c:chart><c:plotArea><c:barChart><c:ser>" +
            "<c:tx><c:strRef><c:strCache><c:pt idx=`"0`"><c:v>TC 系列名</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
            "<c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val=`"1`"/>" +
            "<c:lvl><c:pt idx=`"0`"><c:v>上期</c:v></c:pt></c:lvl><c:lvl><c:pt idx=`"0`"><c:v>2024年</c:v></c:pt></c:lvl>" +
            "</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>" +
            "<c:val><c:numRef><c:numCache><c:pt idx=`"0`"><c:v>100</c:v></c:pt></c:numCache></c:numRef></c:val>" +
            "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
        readChartText $xml | Should -Be "TC 系列名"
    }

    It "系列名は、セル参照ではない直値（c:tx の直下の c:v）でも読む" {
        $xml = "<c:chartSpace $cNs><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:v>直値の系列名</c:v></c:tx>" +
            "<c:cat><c:strRef><c:strCache><c:pt idx=`"0`"><c:v>項目A</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
            "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
        readChartText $xml | Should -Be "直値の系列名"
    }

    It "タイトル・軸の名前は、セル参照（c:title/c:tx/c:strRef/c:strCache/c:v）でも読む" {
        # 直値のタイトル（c:rich の a:p）は readXmlLines 側で読めるが、セル参照（Excel の「=Sheet1!\$A\$1」のような
        # タイトル）は c:rich を持たず、系列名と同じ形（c:tx/c:strRef/c:strCache/c:pt/c:v）になる
        $xml = "<c:chartSpace $cNs><c:chart>" +
            "<c:title><c:tx><c:strRef><c:strCache><c:pt idx=`"0`"><c:v>TC セル参照のタイトル</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>" +
            "<c:plotArea><c:barChart><c:ser><c:tx><c:v>TC 系列名</c:v></c:tx></c:ser></c:barChart>" +
            "<c:catAx><c:title><c:tx><c:strRef><c:strCache><c:pt idx=`"0`"><c:v>TC セル参照の軸の名前</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title></c:catAx>" +
            "</c:plotArea></c:chart></c:chartSpace>"
        readChartText $xml | Should -Be "TC セル参照のタイトル TC セル参照の軸の名前 TC 系列名"
    }

    It "項目の点数が多いグラフでも、項目名を読まないため速く終わる（回帰の確かめ）" {
        # 項目名を辿って祖先を確かめる古い実装は、点数の多いグラフで二次関数的に遅くなっていた
        # （実測: 2 万点で約 46 秒、5 万点で約 230 秒）。系列（c:ser）の数だけ調べる今の実装は、項目名の祖先をたどる
        # 処理が無いため二次関数的には遅くならないが、XML 自体は点の数だけ大きくなるので読み込みの時間は点数にほぼ比例して増える
        # （実測: -Ci のカバレッジ計測下で 2 万点は約 7 秒、5 万点は約 19 秒）。
        # 古い実装に戻ってもしきい値ぎりぎりで通ってしまわないよう、点数を 5 万に増やす。
        # しきい値は、今の実装の実測（約 19 秒）に機械の負荷やカバレッジ計測の変動の余裕を持たせつつ、
        # 古い実装（5 万点で約 230 秒）とは十分に区別できる 60 秒とする
        $pts = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt 50000; $i++) {
            [void]$pts.Append("<c:pt idx=`"$i`"><c:v>項目$i</c:v></c:pt>")
        }
        $xml = "<c:chartSpace $cNs><c:chart><c:plotArea><c:barChart><c:ser>" +
            "<c:tx><c:strRef><c:strCache><c:pt idx=`"0`"><c:v>TC 大きい系列名</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
            "<c:cat><c:strRef><c:strCache><c:ptCount val=`"50000`"/>$($pts.ToString())</c:strCache></c:strRef></c:cat>" +
            "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
        $result = $null
        $elapsed = (Measure-Command { $result = readChartText $xml }).TotalSeconds
        $result | Should -Be "TC 大きい系列名"
        $elapsed | Should -BeLessThan 60
    }
}

Describe "getCellPosition" -Tag Unit {
    It "セル番地を @(行, 列) にする（$ 付き・小文字も読む）" {
        (getCellPosition "AB12") -join "," | Should -Be "12,28"
        (getCellPosition '$A$1') -join "," | Should -Be "1,1"
        (getCellPosition "xfd1048576") -join "," | Should -Be "1048576,16384"
    }

    It "セル番地として読めなければ @(0, 0)" {
        (getCellPosition "") -join "," | Should -Be "0,0"
        (getCellPosition "1A") -join "," | Should -Be "0,0"
        (getCellPosition "A1:B2") -join "," | Should -Be "0,0"
    }
}

Describe "readZipEntry（部品・合計のサイズの上限、偽りのヘッダー）" -Tag Io {
    BeforeEach {
        # 1ファイル分の合計を数える $script:zipTotalReadBytes は、readDocxUnits 等ではなくここでは
        # readZipEntry を直接呼ぶため、テストごとに 0 から数え直す
        $script:zipTotalReadBytes = 0
    }

    It "部品1つの展開後の大きさ（entry.Length）が上限を超えたら、ZipSizeLimitException（部品名・大きさ・Part）にする" {
        $orig = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = 5
        try {
            $path = "$TestDrive\part_over.zip"
            newRawZip $path @{ "a.xml" = @{ bytes = [System.Text.Encoding]::UTF8.GetBytes("123456") } }
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                { readZipEntry $zip "a.xml" } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"
                $caught = $null
                try { readZipEntry $zip "a.xml" } catch { $caught = $_.Exception }
                # [ZipSizeLimitException] の型リテラルは、Pester の It ブロックからは解決できない
                # （BeforeAll で dot-source したクラスでも、実行スコープが分かれるため）ため、
                # GetType().Name で型名を確かめる
                $caught.GetType().Name | Should -Be "ZipSizeLimitException"
                $caught.PartName | Should -Be "a.xml"
                $caught.MeasuredBytes | Should -Be 6
                $caught.LimitKind | Should -Be "Part"
            } finally {
                $zip.Dispose()
            }
        } finally {
            $script:zipPartMaxBytes = $orig
        }
    }

    It "1ファイルの中で読む合計（複数の部品の展開後の大きさの和）が上限を超えたら、ZipSizeLimitException（Total）にする" {
        $orig = $script:zipTotalMaxBytes
        $script:zipTotalMaxBytes = 8
        try {
            $path = "$TestDrive\total_over.zip"
            newZip $path @{ "a.xml" = "12345"; "b.xml" = "12345" }  # 1つ5バイト、合計10バイト
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                readZipEntry $zip "a.xml" | Out-Null  # 5バイト。まだ上限(8)以下
                $caught = $null
                try { readZipEntry $zip "b.xml" } catch { $caught = $_.Exception }  # 合計10 > 8
                $caught.GetType().Name | Should -Be "ZipSizeLimitException"
                $caught.Message | Should -BeLike "*大きすぎるため更新できません*"
                $caught.PartName | Should -Be "b.xml"
                # MeasuredBytes は Total のときは超えた時点の「申告の合計」（b.xml 自身の大きさ5ではなく、a.xml と合わせた10）
                $caught.MeasuredBytes | Should -Be 10
                $caught.LimitKind | Should -Be "Total"
            } finally {
                $zip.Dispose()
            }
        } finally {
            $script:zipTotalMaxBytes = $orig
        }
    }

    It "ヘッダーに書かれた大きさより実際の中身が大きい（偽りのヘッダー）ときは、ZipSizeLimitException（Part）にする" {
        $path = "$TestDrive\fake_header.zip"
        newRawZip $path @{ "a.xml" = @{ bytes = [System.Text.Encoding]::UTF8.GetBytes("0123456789"); fakeSize = 3 } }  # 実際は10バイトだが3バイトと偽る
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            { readZipEntry $zip "a.xml" } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"
            $caught = $null
            try { readZipEntry $zip "a.xml" } catch { $caught = $_.Exception }
            $caught.GetType().Name | Should -Be "ZipSizeLimitException"
            $caught.PartName | Should -Be "a.xml"
            # 実際の大きさまでは分からないため、検出できた時点（申告の3バイト + 1）の値になる
            $caught.MeasuredBytes | Should -Be 4
            $caught.LimitKind | Should -Be "Part"
        } finally {
            $zip.Dispose()
        }
    }

    It "<name>: 多バイト文字の途中に境界が来ても、文字単位ではなくバイト単位で偽りのヘッダーと見分ける" -TestCases @(
        @{ name = "1バイト文字（ASCII）" ; char = "9" }
        @{ name = "3バイト文字" ; char = "あ" }
        @{ name = "4バイト文字（サロゲートペア）" ; char = "😀" }
    ) {
        param ($name, $char)
        $prefix = [System.Text.Encoding]::UTF8.GetBytes("0123456789")
        $charBytes = [System.Text.Encoding]::UTF8.GetBytes($char)
        $bytes = $prefix + $charBytes
        $path = "$TestDrive\fake_header_$([guid]::NewGuid()).zip"
        # 実際の大きさより1バイト少なく申告する（多バイト文字の途中で境界が来ても見つけられることを確かめる。
        # 文字単位で数えていると、最後の文字が丸ごと余分と判定され、この1バイト差の偽りを見逃すおそれがある）
        newRawZip $path @{ "a.xml" = @{ bytes = $bytes; fakeSize = ($bytes.Length - 1) } }
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            { readZipEntry $zip "a.xml" } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"
        } finally {
            $zip.Dispose()
        }
    }

    It "<kind> の上限ちょうどは読め、1 バイト超えると ZipSizeLimitException にする" -TestCases @(
        @{ kind = "Part" }
        @{ kind = "Total" }
    ) {
        param ($kind)
        $origPart = $script:zipPartMaxBytes
        $origTotal = $script:zipTotalMaxBytes
        try {
            if ($kind -eq "Part") { $script:zipPartMaxBytes = 5 } else { $script:zipTotalMaxBytes = 5 }
            $path = "$TestDrive\edge_$kind.zip"
            newRawZip $path @{ "ok.xml" = @{ bytes = [System.Text.Encoding]::UTF8.GetBytes("12345") }; "ng.xml" = @{ bytes = [System.Text.Encoding]::UTF8.GetBytes("123456") } }
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                readZipEntry $zip "ok.xml" | Should -Be "12345"
                $script:zipTotalReadBytes = 0
                $caught = $null
                try { readZipEntry $zip "ng.xml" } catch { $caught = $_.Exception }
                $caught.GetType().Name | Should -Be "ZipSizeLimitException"
                $caught.LimitKind | Should -Be $kind
            } finally {
                $zip.Dispose()
            }
        } finally {
            $script:zipPartMaxBytes = $origPart
            $script:zipTotalMaxBytes = $origTotal
        }
    }

    It "中身が空の部品は、空の文字列で読める" {
        $path = "$TestDrive\empty_part.zip"
        newRawZip $path @{ "a.xml" = @{ bytes = [byte[]]@() } }
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            readZipEntry $zip "a.xml" | Should -Be ""
        } finally {
            $zip.Dispose()
        }
    }

    It "圧縮（Deflate）した部品でも、偽りのヘッダー（実際より小さい申告）を見つける" {
        $path = "$TestDrive\deflate_fake.zip"
        $bytes = [System.Text.Encoding]::UTF8.GetBytes("0123456789" * 20)
        newRawZip $path @{ "a.xml" = @{ bytes = $bytes; compress = $true; fakeSize = 10 } }
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $caught = $null
            try { readZipEntry $zip "a.xml" } catch { $caught = $_.Exception }
            $caught.GetType().Name | Should -Be "ZipSizeLimitException"
            $caught.LimitKind | Should -Be "Part"
        } finally {
            $zip.Dispose()
        }
    }

    It "申告の大きさが大きくても、実際の中身が小さければその中身を返す（申告の大きさぶんを先に確保しない）" {
        $orig = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = 1000000
        try {
            $path = "$TestDrive\declared_big.zip"
            newRawZip $path @{ "a.xml" = @{ bytes = [System.Text.Encoding]::UTF8.GetBytes("a"); fakeSize = 900000 } }
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                readZipEntry $zip "a.xml" | Should -Be "a"
            } finally {
                $zip.Dispose()
            }
        } finally {
            $script:zipPartMaxBytes = $orig
        }
    }

    It "<name>（宣言どおりの大きさ）は今までどおり読める（回帰）" -TestCases @(
        @{ name = "BOM付きUTF-8"; bytes = [byte[]](0xEF, 0xBB, 0xBF) + [System.Text.Encoding]::UTF8.GetBytes("あ") }
        @{ name = "UTF-16 LE（BOM付き）"; bytes = [System.Text.Encoding]::Unicode.GetPreamble() + [System.Text.Encoding]::Unicode.GetBytes("あ") }
        @{ name = "UTF-16 BE（BOM付き）"; bytes = [System.Text.Encoding]::BigEndianUnicode.GetPreamble() + [System.Text.Encoding]::BigEndianUnicode.GetBytes("あ") }
        @{ name = "BOM無し（UTF-8として読む）"; bytes = [System.Text.Encoding]::UTF8.GetBytes("あ") }
    ) {
        param ($name, $bytes)
        $path = "$TestDrive\bom_$([guid]::NewGuid()).zip"
        newRawZip $path @{ "a.xml" = @{ bytes = $bytes } }
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            readZipEntry $zip "a.xml" | Should -Be "あ"
        } finally {
            $zip.Dispose()
        }
    }

    It "<name>: 入れ物の境目を BOM・多バイト文字がまたいでも、同じ文字列に戻る。偽りのヘッダーも見つかる" -TestCases @(
        @{ name = "BOM付きUTF-8"; size = 3; encoding = [System.Text.UTF8Encoding]::new($true) }
        @{ name = "BOM付きUTF-8（入れ物 5 バイト）"; size = 5; encoding = [System.Text.UTF8Encoding]::new($true) }
        @{ name = "UTF-16 LE（BOM付き）"; size = 3; encoding = [System.Text.UnicodeEncoding]::new($false, $true) }
        @{ name = "UTF-16 BE（BOM付き）"; size = 3; encoding = [System.Text.UnicodeEncoding]::new($true, $true) }
        @{ name = "BOM無し（UTF-8）"; size = 4; encoding = [System.Text.UTF8Encoding]::new($false) }
    ) {
        param ($name, $size, $encoding)
        $text = "a" + "あ" + "😀" + "bcd"
        $bytes = $encoding.GetPreamble() + $encoding.GetBytes($text)
        $origBuffer = $script:zipReadBuffer
        $origSize = $script:zipReadBufferSize
        $script:zipReadBufferSize = $size
        $script:zipReadBuffer = $null
        try {
            $path = "$TestDrive\boundary_$([guid]::NewGuid()).zip"
            newRawZip $path @{ "a.xml" = @{ bytes = $bytes } }
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                readZipEntry $zip "a.xml" | Should -Be $text
            } finally {
                $zip.Dispose()
            }

            $script:zipTotalReadBytes = 0
            $path = "$TestDrive\boundary_fake_$([guid]::NewGuid()).zip"
            newRawZip $path @{ "a.xml" = @{ bytes = $bytes; fakeSize = ($bytes.Length - 1) } }
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                { readZipEntry $zip "a.xml" } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"
            } finally {
                $zip.Dispose()
            }
        } finally {
            $script:zipReadBufferSize = $origSize
            $script:zipReadBuffer = $origBuffer
        }
    }
}

Describe "newXmlDocument" -Tag Unit {
    It "名前空間付きのXMLを、今までどおりXmlDocumentとして読める" {
        $doc = newXmlDocument '<a:x xmlns:a="urn:test"><a:y>value</a:y></a:x>'
        $doc | Should -BeOfType ([System.Xml.XmlDocument])
        $doc.DocumentElement.LocalName | Should -Be "x"
        $doc.DocumentElement.FirstChild.InnerText | Should -Be "value"
    }

    It "制御文字（&#x1;）を含む文字も、空白を挟んでも今までどおり読める（回帰。文字の検査はしない）" {
        $doc = newXmlDocument "<a>x&#x1; y</a>"
        $doc.DocumentElement.InnerText | Should -Be "x$([char]1) y"
    }

    It "DTD宣言（DOCTYPE）を含むXMLは例外にする（実体参照を入れ子にして膨張させる攻撃を防ぐ）" {
        { newXmlDocument '<?xml version="1.0"?><!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>' } | Should -Throw
    }

    It "readXmlLines も、今までどおりDTDを拒否する（LoadXml経由だけでなく両方の読み込みで防ぐ）" {
        { readXmlLines '<?xml version="1.0"?><!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>' $wNs } | Should -Throw
    }
}

Describe "readDocxUnits・readPptxUnits・readXlsxObjectUnits（部品がサイズの上限を超えたZIP）" -Tag Io {
    It "readDocxUnits: 本文（word/document.xml）が部品の上限を超えたら、大きすぎるという例外にする" {
        $orig = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = 10
        try {
            $path = "$TestDrive\big.docx"
            newZip $path @{ "word/document.xml" = "<w:document ${wNs}><w:body>$('あ' * 20)</w:body></w:document>" }
            { readDocxUnits $path } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"
        } finally {
            $script:zipPartMaxBytes = $orig
        }
    }

    It "readPptxUnits: プレゼンテーション情報（ppt/presentation.xml）が部品の上限を超えたら例外にする" {
        $orig = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = 10
        try {
            $path = "$TestDrive\big.pptx"
            newZip $path @{ "ppt/presentation.xml" = "<p:presentation ${pNs}><p:sldIdLst>$('<!-- big -->' * 5)</p:sldIdLst></p:presentation>" }
            { readPptxUnits $path } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"
        } finally {
            $script:zipPartMaxBytes = $orig
        }
    }

    It "readXlsxObjectUnits: 図形（drawingN.xml）が部品の上限を超えたら例外にする（Excelのセルの取り込みは、呼び出し元のextract_office.ps1が別に続ける）" {
        # drawing1.xml だけが上限を超えるように、ほかの部品（workbook.xml 等）の大きさから上限を決める
        # （上限を固定の小さい値にすると、drawing1.xml の前に読む workbook.xml 等が先に超えてしまい、
        # 「図形が超えた」ことを確かめられない）
        $workbookXml = "<workbook ${xNs}><sheets><sheet name=`"S`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
        $workbookRelsXml = "<Relationships ${relNs}><Relationship Id=`"rId1`" Type=`"${officeRel}/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
        $sheetXml = "<worksheet ${xNs}/>"
        $sheetRelsXml = "<Relationships ${relNs}><Relationship Id=`"rId1`" Type=`"${officeRel}/drawing`" Target=`"../drawings/drawing1.xml`"/></Relationships>"
        $drawingXml = "<xdr:wsDr ${xdrNs}>$('<!-- big -->' * 50)</xdr:wsDr>"

        $otherParts = @($workbookXml, $workbookRelsXml, $sheetXml, $sheetRelsXml)
        $maxOtherBytes = ($otherParts | ForEach-Object { [System.Text.Encoding]::UTF8.GetByteCount($_) } | Measure-Object -Maximum).Maximum
        $drawingBytes = [System.Text.Encoding]::UTF8.GetByteCount($drawingXml)
        $drawingBytes | Should -BeGreaterThan $maxOtherBytes  # テストの前提（drawing1.xml がほかの部品より大きい）

        $orig = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = $maxOtherBytes + 1
        try {
            $path = "$TestDrive\big.xlsx"
            newZip $path @{
                "xl/workbook.xml" = $workbookXml
                "xl/_rels/workbook.xml.rels" = $workbookRelsXml
                "xl/worksheets/sheet1.xml" = $sheetXml
                "xl/worksheets/_rels/sheet1.xml.rels" = $sheetRelsXml
                "xl/drawings/drawing1.xml" = $drawingXml
            }
            { readXlsxObjectUnits $path } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"
        } finally {
            $script:zipPartMaxBytes = $orig
        }
    }

    It "readXlsxObjectUnits: グラフ（xl/charts/chart1.xml）だけが部品の上限を超えても、そのグラフだけを空にしてほかは読む。超えた部品は `$sizeFailures に ZipSizeLimitException として返す" {
        # chart1.xml だけが上限を超えるように、ほかの部品の大きさから上限を決める（上の drawing1.xml のテストと同じやり方）
        $workbookXml = "<workbook ${xNs}><sheets><sheet name=`"S`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
        $workbookRelsXml = "<Relationships ${relNs}><Relationship Id=`"rId1`" Type=`"${officeRel}/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
        $sheetXml = "<worksheet ${xNs}/>"
        $sheetRelsXml = "<Relationships ${relNs}><Relationship Id=`"rId1`" Type=`"${officeRel}/drawing`" Target=`"../drawings/drawing1.xml`"/></Relationships>"
        $drawingXml = "<xdr:wsDr ${xdrNs}>$(xAnchor 0 0 (xChartFrame (chartRef 'rId1')))</xdr:wsDr>"
        $drawingRelsXml = "<Relationships ${relNs}><Relationship Id=`"rId1`" Type=`"${docRel}/chart`" Target=`"../charts/chart1.xml`"/></Relationships>"
        $chartXml = "<c:chartSpace ${cNs}>$('<!-- big -->' * 100)</c:chartSpace>"

        $otherParts = @($workbookXml, $workbookRelsXml, $sheetXml, $sheetRelsXml, $drawingXml, $drawingRelsXml)
        $maxOtherBytes = ($otherParts | ForEach-Object { [System.Text.Encoding]::UTF8.GetByteCount($_) } | Measure-Object -Maximum).Maximum
        $chartBytes = [System.Text.Encoding]::UTF8.GetByteCount($chartXml)
        $chartBytes | Should -BeGreaterThan $maxOtherBytes  # テストの前提（chart1.xml がほかの部品より大きい）

        $orig = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = $maxOtherBytes + 1
        try {
            $path = "$TestDrive\big_chart.xlsx"
            newZip $path @{
                "xl/workbook.xml" = $workbookXml
                "xl/_rels/workbook.xml.rels" = $workbookRelsXml
                "xl/worksheets/sheet1.xml" = $sheetXml
                "xl/worksheets/_rels/sheet1.xml.rels" = $sheetRelsXml
                "xl/drawings/drawing1.xml" = $drawingXml
                "xl/drawings/_rels/drawing1.xml.rels" = $drawingRelsXml
                "xl/charts/chart1.xml" = $chartXml
            }
            $failures = New-Object System.Collections.Generic.List[string]
            $sizeFailures = New-Object System.Collections.Generic.List[object]
            $units = readXlsxObjectUnits $path $failures $sizeFailures

            # グラフだけが空になり、ブック自体は読める（例外にならない）
            $units.Contains("S[図形]") | Should -Be $false
            @($failures) | Should -Be @("xl/charts/chart1.xml")

            # $sizeFailures に、原因を調べるための部品名・大きさ・部品ごとか合計かを持つ例外が入る
            $sizeFailures.Count | Should -Be 1
            $sizeFailures[0].GetType().Name | Should -Be "ZipSizeLimitException"
            $sizeFailures[0].PartName | Should -Be "xl/charts/chart1.xml"
            $sizeFailures[0].LimitKind | Should -Be "Part"
            $sizeFailures[0].MeasuredBytes | Should -BeGreaterThan $maxOtherBytes
        } finally {
            $script:zipPartMaxBytes = $orig
        }
    }

    It "部品ごとは上限以下でも、申告した大きさ（展開後の大きさ）の合計が上限を超えたら例外にする（実際の中身は小さいまま）" {
        $origPart = $script:zipPartMaxBytes
        $origTotal = $script:zipTotalMaxBytes
        $script:zipPartMaxBytes = 10
        $script:zipTotalMaxBytes = 15
        $script:zipTotalReadBytes = 0
        try {
            $path = "$TestDrive\inflated_total.zip"
            # 部品ごとの中身は1バイトで小さいが、申告する大きさ（8バイト）は part の上限(10)以下。
            # 2部品分の申告の合計(16)が total の上限(15)を超える（ZIP爆弾のように、申告だけ大きく中身は小さい部品が
            # 複数集まっても、合計の上限で止めることを確かめる）
            newRawZip $path @{
                "a.xml" = @{ bytes = [System.Text.Encoding]::UTF8.GetBytes("1"); fakeSize = 8 }
                "b.xml" = @{ bytes = [System.Text.Encoding]::UTF8.GetBytes("2"); fakeSize = 8 }
            }
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                readZipEntry $zip "a.xml" | Out-Null  # 申告8バイト。まだ上限(15)以下
                { readZipEntry $zip "b.xml" } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"  # 申告の合計16 > 15
            } finally {
                $zip.Dispose()
            }
        } finally {
            $script:zipPartMaxBytes = $origPart
            $script:zipTotalMaxBytes = $origTotal
        }
    }

    It "合計の上限を超えて失敗した後も、別のファイルを読むときは合計のカウントが1ファイルごとに0から数え直される" {
        $orig = $script:zipTotalMaxBytes
        try {
            # 1つ目: 上限を1バイトにして、必ず合計の上限超過で失敗させる（$script:zipTotalReadBytes がリセットされず
            # 残ったままだと、2つ目のファイルの判定に混ざってしまうことを確かめるため、残り値をわざと作る）
            $firstXml = "<w:document ${wNs}><w:body>$('x' * 10)</w:body></w:document>"
            $overPath = "$TestDrive\total_over_first.docx"
            newZip $overPath @{ "word/document.xml" = $firstXml }
            $script:zipTotalMaxBytes = 1
            { readDocxUnits $overPath } | Should -Throw -ExpectedMessage "*大きすぎるため更新できません*"

            # 2つ目: この内容だけなら収まるが、1つ目の残り（10バイト超）が足されたままだと超えてしまう上限にする。
            # readDocxUnits が呼び出しのたびに $script:zipTotalReadBytes を0から数え直していれば、収まって読める
            $secondXml = "<w:document ${wNs}><w:body>$(wPara '小さい本文')</w:body></w:document>"
            $secondBytes = [System.Text.Encoding]::UTF8.GetByteCount($secondXml)
            $smallPath = "$TestDrive\small_after_total_over.docx"
            newZip $smallPath @{ "word/document.xml" = $secondXml }
            $script:zipTotalMaxBytes = $secondBytes + 1
            { readDocxUnits $smallPath } | Should -Not -Throw
        } finally {
            $script:zipTotalMaxBytes = $orig
        }
    }
}
