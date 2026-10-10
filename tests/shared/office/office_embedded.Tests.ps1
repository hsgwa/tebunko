# Pester 5 で実行: .\tests\run.ps1
# Word・PowerPoint は使わず、埋め込みのあるファイルをテスト内で作成して検証する（作成者名などは入れない）
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\shared\office\office_reader.ps1"
    . "${scriptsDir}\shared\office\office_embedded.ps1"

    function newZipBytes {
        # ZIP内のパス → 内容（文字列かバイト列）の辞書から、ZIP のバイト列を作る（部品は無圧縮）
        param (
            [hashtable]$entries
        )

        $memory = New-Object System.IO.MemoryStream
        $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        foreach ($name in $entries.Keys) {
            $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::NoCompression)
            $stream = $entry.Open()
            $value = $entries[$name]
            $bytes = $(if ($value -is [byte[]]) { $value } else { (New-Object System.Text.UTF8Encoding($false)).GetBytes([string]$value) })
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Dispose()
        }
        $zip.Dispose()
        return , $memory.ToArray()
    }

    function newZipFile {
        param (
            [string]$path,
            [hashtable]$entries
        )

        [System.IO.File]::WriteAllBytes($path, (newZipBytes $entries))
    }

    $script:xNs = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $script:wNs = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $script:pNs = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $script:relNs = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
    $script:officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    function newWorkbookEntries {
        # 表示の 2 つのシートと、非表示のシート 1 つがあるブック。文字のセルは共有文字列・インライン文字列・数式の文字
        return @{
            "xl/workbook.xml"            = "<workbook $script:xNs><sheets><sheet name=`"売上`" sheetId=`"1`" r:id=`"rId1`"/><sheet name=`"隠し`" sheetId=`"2`" state=`"hidden`" r:id=`"rId2`"/><sheet name=`"別`" sheetId=`"3`" r:id=`"rId3`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $script:relNs><Relationship Id=`"rId1`" Type=`"$script:officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/><Relationship Id=`"rId2`" Type=`"$script:officeRel/worksheet`" Target=`"worksheets/sheet2.xml`"/><Relationship Id=`"rId3`" Type=`"$script:officeRel/worksheet`" Target=`"worksheets/sheet3.xml`"/></Relationships>"
            "xl/sharedStrings.xml"       = "<sst $script:xNs><si><t>共有A</t></si><si><r><t>リッチ</t></r><r><t>文字</t></r><rPh sb=`"0`" eb=`"1`"><t>よみ</t></rPh></si></sst>"
            "xl/worksheets/sheet1.xml"   = "<worksheet $script:xNs><sheetData><row r=`"1`"><c r=`"A1`" t=`"s`"><v>0</v></c><c r=`"B1`"><v>123</v></c><c r=`"C1`" t=`"inlineStr`"><is><t>インライン</t></is></c></row><row r=`"2`"><c r=`"A2`" t=`"str`"><f>A1&amp;`"x`"</f><v>式の文字</v></c><c r=`"B2`" t=`"s`"><v>1</v></c></row></sheetData></worksheet>"
            "xl/worksheets/sheet2.xml"   = "<worksheet $script:xNs><sheetData><row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>隠しシートの文字</t></is></c></row></sheetData></worksheet>"
            "xl/worksheets/sheet3.xml"   = "<worksheet $script:xNs><sheetData><row r=`"40`"><c r=`"Z40`" t=`"inlineStr`"><is><t>遠いセル</t></is></c></row></sheetData></worksheet>"
        }
    }

    function newDocumentEntries {
        # 埋め込み 1 つ（RelId rId5）を持つ Word 文書の部品。$embeddedXml は w:body の中に入れる
        param (
            [string]$bodyInner,
            [hashtable]$embeddings
        )

        $entries = @{
            "word/document.xml"            = "<w:document $script:wNs><w:body>$bodyInner</w:body></w:document>"
            "word/_rels/document.xml.rels" = "<Relationships $script:relNs>" + (($embeddings.Keys | Sort-Object | ForEach-Object {
                "<Relationship Id=`"$_`" Type=`"$script:officeRel/package`" Target=`"embeddings/$($embeddings[$_].Name)`"/>"
            }) -join "") + "</Relationships>"
        }
        foreach ($id in $embeddings.Keys) {
            $entries["word/embeddings/$($embeddings[$id].Name)"] = $embeddings[$id].Bytes
        }
        return $entries
    }

    function embedPara([string]$relId) {
        return "<w:p><w:r><w:object><o:OLEObject Type=`"Embed`" ProgID=`"Excel.Sheet.12`" r:id=`"$relId`"/></w:object></w:r></w:p>"
    }
}

Describe "readXlsxCellLines" -Tag Unit {
    It "表示のシートの文字のセルを行ごとにタブでつなぎ、数値・非表示のシート・読み仮名は読まない" {
        $bytes = newZipBytes (newWorkbookEntries)
        $memory = New-Object System.IO.MemoryStream (, $bytes)
        $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $script:zipTotalReadBytes = 0
            $lines = @(readXlsxCellLines $zip)
        } finally {
            $zip.Dispose()
        }
        $lines -join "|" | Should -Be "共有A`tインライン|式の文字`tリッチ文字|遠いセル"
    }

    It "シートの合計が流れ読みの上限を超えたら、そこから先は読まず、失敗に部品の名前を足す" {
        $orig = $script:zipSheetStreamMaxBytes
        $script:zipSheetStreamMaxBytes = 400
        try {
            $bytes = newZipBytes (newWorkbookEntries)
            $memory = New-Object System.IO.MemoryStream (, $bytes)
            $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
            try {
                $script:zipTotalReadBytes = 0
                $failures = New-Object System.Collections.Generic.List[string]
                $lines = @(readXlsxCellLines $zip $failures)
            } finally {
                $zip.Dispose()
            }
            $failures.Count | Should -Be 1
            $failures[0] | Should -BeLike "xl/worksheets/sheet*.xml*上限*"
        } finally {
            $script:zipSheetStreamMaxBytes = $orig
        }
    }
}

Describe "Word の埋め込み" -Tag Unit {
    It "埋め込んだ Excel のブックの文字を、ページの [埋め込み1] に入れる（本文は変わらない）" {
        $path = "$TestDrive\doc_xlsx.docx"
        $embeddings = @{ rId5 = @{ Name = "Microsoft_Excel_Worksheet.xlsx"; Bytes = (newZipBytes (newWorkbookEntries)) } }
        newZipFile $path (newDocumentEntries ("<w:p><w:r><w:t>本文</w:t></w:r></w:p>" + (embedPara "rId5")) $embeddings)

        $units = readDocxUnits $path
        @($units.Keys) -join "," | Should -Be "ページ001,ページ001[埋め込み1]"
        @($units["ページ001"]) -join "|" | Should -Be "本文"
        @($units["ページ001[埋め込み1]"]) -join "|" | Should -Be "共有A`tインライン|式の文字`tリッチ文字|遠いセル"
    }

    It "埋め込んだ Word 文書の文字を読み、その中の埋め込みは読まない（深さは 1 段まで）" {
        $inner = newDocumentEntries ("<w:p><w:r><w:t>入れ子の本文</w:t></w:r></w:p>" + (embedPara "rId5")) @{
            rId5 = @{ Name = "deep.xlsx"; Bytes = (newZipBytes (newWorkbookEntries)) }
        }
        $path = "$TestDrive\doc_docx.docx"
        $embeddings = @{ rId5 = @{ Name = "Microsoft_Word_Document.docx"; Bytes = (newZipBytes $inner) } }
        newZipFile $path (newDocumentEntries (embedPara "rId5") $embeddings)

        $units = readDocxUnits $path
        @($units.Keys) -join "," | Should -Be "ページ001[埋め込み1]"
        @($units["ページ001[埋め込み1]"]) -join "|" | Should -Be "入れ子の本文"
    }

    It "同じ部品を指す参照は 1 回だけ読み、別の部品は読んだ順に 2 番から番号を振る" {
        $path = "$TestDrive\doc_two.docx"
        $other = newWorkbookEntries
        $other["xl/sharedStrings.xml"] = "<sst $script:xNs><si><t>別ブック</t></si><si><t>x</t></si></sst>"
        $embeddings = @{
            rId5 = @{ Name = "a.xlsx"; Bytes = (newZipBytes (newWorkbookEntries)) }
            rId6 = @{ Name = "b.xlsx"; Bytes = (newZipBytes $other) }
        }
        newZipFile $path (newDocumentEntries ((embedPara "rId5") + (embedPara "rId5") + (embedPara "rId6")) $embeddings)

        $units = readDocxUnits $path
        @($units.Keys) -join "," | Should -Be "ページ001[埋め込み1],ページ001[埋め込み2]"
        @($units["ページ001[埋め込み2]"])[0] | Should -Be "別ブック`tインライン"
    }

    It "ZIP でない埋め込み・どの形式でもない ZIP は、読まずに続ける（失敗にもしない）" {
        $path = "$TestDrive\doc_unknown.docx"
        $embeddings = @{
            rId5 = @{ Name = "legacy.bin"; Bytes = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 1, 2, 3, 4) }
            rId6 = @{ Name = "other.zip"; Bytes = (newZipBytes @{ "readme.txt" = "文字" }) }
        }
        newZipFile $path (newDocumentEntries ("<w:p><w:r><w:t>本文</w:t></w:r></w:p>" + (embedPara "rId5") + (embedPara "rId6")) $embeddings)

        $failures = New-Object System.Collections.Generic.List[string]
        $sizeFailures = New-Object System.Collections.Generic.List[object]
        $units = readDocxUnits $path $failures $sizeFailures
        @($units.Keys) -join "," | Should -Be "ページ001"
        $failures.Count | Should -Be 0
        $sizeFailures.Count | Should -Be 0
    }

    It "壊れた埋め込みは、その埋め込みだけを読まず、部品の名前を失敗に足して、ほかの埋め込みと本文を出す" {
        $broken = newWorkbookEntries
        $broken["xl/worksheets/sheet1.xml"] = "<worksheet $script:xNs><sheetData><row"
        $path = "$TestDrive\doc_broken.docx"
        $embeddings = @{
            rId5 = @{ Name = "broken.xlsx"; Bytes = (newZipBytes $broken) }
            rId6 = @{ Name = "ok.xlsx"; Bytes = (newZipBytes (newWorkbookEntries)) }
        }
        newZipFile $path (newDocumentEntries ("<w:p><w:r><w:t>本文</w:t></w:r></w:p>" + (embedPara "rId5") + (embedPara "rId6")) $embeddings)

        $failures = New-Object System.Collections.Generic.List[string]
        $units = readDocxUnits $path $failures
        @($failures) -join "|" | Should -Be "word/embeddings/broken.xlsx"
        @($units["ページ001"]) -join "|" | Should -Be "本文"
        @($units.Keys) -contains "ページ001[埋め込み2]" | Should -Be $true
        @($units.Keys) -contains "ページ001[埋め込み1]" | Should -Be $false
    }

    It "部品ごとの上限を超えた埋め込みは読まず、例外を失敗に足して続ける。1ファイルの合計の上限は投げ直す" {
        $big = newWorkbookEntries
        $big["xl/media/pad.bin"] = [byte[]]::new(3000)
        $path = "$TestDrive\doc_big.docx"
        $embeddings = @{ rId5 = @{ Name = "big.xlsx"; Bytes = (newZipBytes $big) } }
        newZipFile $path (newDocumentEntries ("<w:p><w:r><w:t>本文</w:t></w:r></w:p>" + (embedPara "rId5")) $embeddings)

        $origPart = $script:zipPartMaxBytes
        $origTotal = $script:zipTotalMaxBytes
        try {
            $script:zipPartMaxBytes = 2500
            $failures = New-Object System.Collections.Generic.List[string]
            $sizeFailures = New-Object System.Collections.Generic.List[object]
            $units = readDocxUnits $path $failures $sizeFailures
            @($failures) -join "|" | Should -Be "word/embeddings/big.xlsx"
            $sizeFailures.Count | Should -Be 1
            $sizeFailures[0].LimitKind | Should -Be "Part"
            $sizeFailures[0].PartName | Should -Be "word/embeddings/big.xlsx"
            @($units["ページ001"]) -join "|" | Should -Be "本文"

            $script:zipPartMaxBytes = $origPart
            $script:zipTotalMaxBytes = 4000
            { readDocxUnits $path (New-Object System.Collections.Generic.List[string]) (New-Object System.Collections.Generic.List[object]) } |
                Should -Throw -ExceptionType ([ZipSizeLimitException])
        } finally {
            $script:zipPartMaxBytes = $origPart
            $script:zipTotalMaxBytes = $origTotal
        }
    }
}

Describe "PowerPoint の埋め込み" -Tag Unit {
    It "スライドに埋め込んだ Excel のブックの文字を、スライドの [埋め込み1] に入れる" {
        $slide = "<p:sld $script:pNs><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>題</a:t></a:r></a:p></p:txBody></p:sp>" +
            "<p:graphicFrame><a:graphic><a:graphicData><p:oleObj r:id=`"rId2`" progId=`"Excel.Sheet.12`"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"
        $path = "$TestDrive\slide_xlsx.pptx"
        newZipFile $path @{
            "ppt/presentation.xml"            = "<p:presentation $script:pNs><p:sldIdLst><p:sldId id=`"256`" r:id=`"rId1`"/></p:sldIdLst></p:presentation>"
            "ppt/_rels/presentation.xml.rels" = "<Relationships $script:relNs><Relationship Id=`"rId1`" Type=`"$script:officeRel/slide`" Target=`"slides/slide1.xml`"/></Relationships>"
            "ppt/slides/slide1.xml"           = $slide
            "ppt/slides/_rels/slide1.xml.rels" = "<Relationships $script:relNs><Relationship Id=`"rId2`" Type=`"$script:officeRel/package`" Target=`"../embeddings/Microsoft_Excel_Worksheet.xlsx`"/></Relationships>"
            "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx" = (newZipBytes (newWorkbookEntries))
        }

        $units = readPptxUnits $path
        @($units.Keys) -join "," | Should -Be "スライド001,スライド001[埋め込み1]"
        @($units["スライド001"]) -join "|" | Should -Be "題"
        @($units["スライド001[埋め込み1]"]) -join "|" | Should -Be "共有A`tインライン|式の文字`tリッチ文字|遠いセル"
    }
}
