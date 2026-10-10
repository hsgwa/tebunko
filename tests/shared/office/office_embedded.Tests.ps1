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

    function setDeclaredSize {
        # ZIP のバイト列の中央ディレクトリの、部品の「展開後の大きさ」だけを書き換える（偽りのヘッダー）
        param (
            [byte[]]$bytes,
            [string]$name,
            [uint32]$fakeSize
        )

        $copy = [byte[]]$bytes.Clone()
        $nameBytes = [System.Text.Encoding]::UTF8.GetBytes($name)
        $patched = $false
        for ($i = 0; $i -le $copy.Length - 46; $i++) {
            if ($copy[$i] -eq 0x50 -and $copy[$i + 1] -eq 0x4b -and $copy[$i + 2] -eq 0x01 -and $copy[$i + 3] -eq 0x02 -and
                [BitConverter]::ToUInt16($copy, $i + 28) -eq $nameBytes.Length) {
                $same = $true
                for ($j = 0; $j -lt $nameBytes.Length; $j++) {
                    if ($copy[$i + 46 + $j] -ne $nameBytes[$j]) { $same = $false; break }
                }
                if ($same) {
                    [Array]::Copy([BitConverter]::GetBytes($fakeSize), 0, $copy, $i + 24, 4)
                    $patched = $true
                    break
                }
            }
        }
        if (-not $patched) { throw "部品が見つかりません: $name" }
        return , $copy
    }

    function newSingleSheetEntries {
        # 表示のシートが 1 つのブック。$sheetData は sheetData の中身、$strings は共有文字列（si の t の中身の配列）
        param (
            [string]$sheetData,
            [string[]]$strings = @()
        )

        $sst = "<sst $script:xNs>" + (($strings | ForEach-Object { "<si><t>$_</t></si>" }) -join "") + "</sst>"
        return @{
            "xl/workbook.xml"            = "<workbook $script:xNs><sheets><sheet name=`"S`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $script:relNs><Relationship Id=`"rId1`" Type=`"$script:officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/sharedStrings.xml"       = $sst
            "xl/worksheets/sheet1.xml"   = "<worksheet $script:xNs><sheetData>$sheetData</sheetData></worksheet>"
        }
    }

    function invokeReadXlsxCellLines {
        param (
            [byte[]]$bytes,
            [hashtable]$state
        )

        $memory = New-Object System.IO.MemoryStream (, $bytes)
        $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $script:zipTotalReadBytes = 0
            return @(readXlsxCellLines $zip $state)
        } finally {
            $zip.Dispose()
        }
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

    It "出す文字数の上限: 共有文字列 1 つを多数のセルが参照すると、解いた後の文字数で数えて上限を超えた時点で打ち切る" {
        $orig = $script:embeddedOutputMaxChars
        try {
            $script:embeddedOutputMaxChars = 1000
            # 100 文字の共有文字列を 50 個のセルが参照する（入力は小さいが、出す行は 5000 文字を超える）
            $cells = (1..50 | ForEach-Object { "<c r=`"A$_`" t=`"s`"><v>0</v></c>" }) -join ""
            $bytes = newZipBytes (newSingleSheetEntries "<row r=`"1`">$cells</row>" @("あ" * 100))
            $state = newEmbeddedState
            $thrown = $null
            try { invokeReadXlsxCellLines $bytes $state | Out-Null } catch { $thrown = $_.Exception }
            $thrown | Should -BeOfType ([ZipSizeLimitException])
            $thrown.LimitKind | Should -Be "Output"
            $state.OutputChars | Should -BeGreaterThan 1000
        } finally {
            $script:embeddedOutputMaxChars = $orig
        }
    }

    It "出す文字数の上限: 大きなインライン文字列 1 つでも、上限を超えたら打ち切る" {
        $orig = $script:embeddedOutputMaxChars
        try {
            $script:embeddedOutputMaxChars = 1000
            $bytes = newZipBytes (newSingleSheetEntries "<row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>$("い" * 2000)</t></is></c></row>")
            { invokeReadXlsxCellLines $bytes (newEmbeddedState) } | Should -Throw -ExceptionType ([ZipSizeLimitException])
        } finally {
            $script:embeddedOutputMaxChars = $orig
        }
    }

    It "出す文字数は、上限の内なら行をそのまま返し、出した文字数（区切り 1 文字を含む）を状態に足す" {
        $bytes = newZipBytes (newSingleSheetEntries "<row r=`"1`"><c r=`"A1`" t=`"s`"><v>0</v></c><c r=`"B1`" t=`"s`"><v>1</v></c></row>" @("甲乙", "丙"))
        $state = newEmbeddedState
        $lines = invokeReadXlsxCellLines $bytes $state
        $lines -join "|" | Should -Be "甲乙`t丙"
        $state.OutputChars | Should -Be 5
    }

    It "申告の大きさを偽ったシートは、申告の大きさで打ち切って部品ごとの上限（Part）の超過にする" {
        $bytes = newZipBytes (newSingleSheetEntries "<row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>$("う" * 500)</t></is></c></row>")
        $forged = setDeclaredSize $bytes "xl/worksheets/sheet1.xml" 50
        $thrown = $null
        try { invokeReadXlsxCellLines $forged (newEmbeddedState) | Out-Null } catch { $thrown = $_.Exception }
        $thrown | Should -BeOfType ([ZipSizeLimitException])
        $thrown.LimitKind | Should -Be "Part"
        $thrown.PartName | Should -Be "xl/worksheets/sheet1.xml"
    }

    It "シートは読んだ大きさが 1 ファイルの合計に数わる（シートごとに数え直さない）" {
        $bytes = newZipBytes (newWorkbookEntries)
        $memory = New-Object System.IO.MemoryStream (, $bytes)
        $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $script:zipTotalReadBytes = 0
            readXlsxCellLines $zip (newEmbeddedState) | Out-Null
            $expected = 0L
            foreach ($name in @("xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet3.xml")) {
                $expected += $zip.GetEntry($name).Length
            }
            $script:zipTotalReadBytes | Should -Be $expected  # 非表示の sheet2 は読まない
        } finally {
            $zip.Dispose()
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

    It "出す文字数の上限を超えた埋め込みは読まず、失敗に足し（Output）、その分は合計から戻して、ほかの埋め込みと本文を出す" {
        $cells = (1..50 | ForEach-Object { "<c r=`"A$_`" t=`"s`"><v>0</v></c>" }) -join ""
        $big = newSingleSheetEntries "<row r=`"1`">$cells</row>" @("あ" * 100)
        $path = "$TestDrive\doc_output.docx"
        $embeddings = @{
            rId5 = @{ Name = "big.xlsx"; Bytes = (newZipBytes $big) }
            rId6 = @{ Name = "ok.xlsx"; Bytes = (newZipBytes (newSingleSheetEntries "<row r=`"1`"><c r=`"A1`" t=`"s`"><v>0</v></c></row>" @("後ろ"))) }
        }
        newZipFile $path (newDocumentEntries ("<w:p><w:r><w:t>本文</w:t></w:r></w:p>" + (embedPara "rId5") + (embedPara "rId6")) $embeddings)

        $orig = $script:embeddedOutputMaxChars
        try {
            $script:embeddedOutputMaxChars = 1000
            $failures = New-Object System.Collections.Generic.List[string]
            $sizeFailures = New-Object System.Collections.Generic.List[object]
            $units = readDocxUnits $path $failures $sizeFailures
        } finally {
            $script:embeddedOutputMaxChars = $orig
        }
        @($failures) -join "|" | Should -Be "word/embeddings/big.xlsx"
        $sizeFailures.Count | Should -Be 1
        $sizeFailures[0].LimitKind | Should -Be "Output"
        @($units["ページ001"]) -join "|" | Should -Be "本文"
        @($units.Keys) -join "," | Should -Be "ページ001,ページ001[埋め込み2]"
        @($units["ページ001[埋め込み2]"]) -join "|" | Should -Be "後ろ"
    }

    It "申告を偽ったシートを持つ埋め込みは読まず、失敗に足し（Part）、本文は出す" {
        $sheet = newSingleSheetEntries "<row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>$("う" * 500)</t></is></c></row>"
        $forged = setDeclaredSize (newZipBytes $sheet) "xl/worksheets/sheet1.xml" 50
        $path = "$TestDrive\doc_forged.docx"
        newZipFile $path (newDocumentEntries ("<w:p><w:r><w:t>本文</w:t></w:r></w:p>" + (embedPara "rId5")) @{ rId5 = @{ Name = "forged.xlsx"; Bytes = $forged } })

        $failures = New-Object System.Collections.Generic.List[string]
        $sizeFailures = New-Object System.Collections.Generic.List[object]
        $units = readDocxUnits $path $failures $sizeFailures
        @($failures) -join "|" | Should -Be "word/embeddings/forged.xlsx"
        $sizeFailures[0].LimitKind | Should -Be "Part"
        @($units.Keys) -join "," | Should -Be "ページ001"
    }

    It "想定外の例外でも、その埋め込みだけを読まず（失敗に足す）、親の本文は出す。中止の例外は投げ直す" {
        $path = "$TestDrive\doc_unexpected.docx"
        $embeddings = @{ rId5 = @{ Name = "a.xlsx"; Bytes = (newZipBytes (newWorkbookEntries)) } }
        newZipFile $path (newDocumentEntries ("<w:p><w:r><w:t>本文</w:t></w:r></w:p>" + (embedPara "rId5")) $embeddings)

        Mock readEmbeddedPackageLines { throw [System.InvalidOperationException]::new("想定外") }
        $failures = New-Object System.Collections.Generic.List[string]
        $units = readDocxUnits $path $failures
        @($failures) -join "|" | Should -Be "word/embeddings/a.xlsx"
        @($units["ページ001"]) -join "|" | Should -Be "本文"

        Mock readEmbeddedPackageLines { throw [System.OperationCanceledException]::new("中止") }
        { readDocxUnits $path (New-Object System.Collections.Generic.List[string]) } | Should -Throw "*中止*"
    }

    It "埋め込みの書き方 <name>: 読む（Type が Embed 以外の OLEObject は読まない）" -TestCases @(
        @{ name = "o:OLEObject（Type=Embed）"; markup = "<w:p><w:r><w:object><o:OLEObject Type=`"Embed`" ProgID=`"Excel.Sheet.12`" r:id=`"rId5`"/></w:object></w:r></w:p>"; expected = "ページ001[埋め込み1]" }
        @{ name = "w:objectEmbed"; markup = "<w:p><w:r><w:object><w:objectEmbed r:id=`"rId5`"/></w:object></w:r></w:p>"; expected = "ページ001[埋め込み1]" }
        @{ name = "o:OLEObject（Type=Link）"; markup = "<w:p><w:r><w:object><o:OLEObject Type=`"Link`" ProgID=`"Excel.Sheet.12`" r:id=`"rId5`"/></w:object></w:r></w:p>"; expected = "" }
    ) {
        $path = "$TestDrive\doc_form_$([guid]::NewGuid().ToString('N')).docx"
        $embeddings = @{ rId5 = @{ Name = "a.xlsx"; Bytes = (newZipBytes (newWorkbookEntries)) } }
        newZipFile $path (newDocumentEntries $markup $embeddings)
        $units = readDocxUnits $path
        @($units.Keys) -join "," | Should -Be $expected
    }

    It "リレーションシップの種類が oleObject（旧形式の .bin）の参照は読み飛ばし、番号も使わない" {
        $entries = newDocumentEntries ((embedPara "rId4") + (embedPara "rId5")) @{ rId5 = @{ Name = "a.xlsx"; Bytes = (newZipBytes (newWorkbookEntries)) } }
        $entries["word/_rels/document.xml.rels"] = "<Relationships $script:relNs>" +
            "<Relationship Id=`"rId4`" Type=`"$script:officeRel/oleObject`" Target=`"embeddings/oleObject1.bin`"/>" +
            "<Relationship Id=`"rId5`" Type=`"$script:officeRel/package`" Target=`"embeddings/a.xlsx`"/></Relationships>"
        $entries["word/embeddings/oleObject1.bin"] = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 1, 2, 3, 4)
        $path = "$TestDrive\doc_ole.docx"
        newZipFile $path $entries

        $failures = New-Object System.Collections.Generic.List[string]
        $units = readDocxUnits $path $failures
        @($units.Keys) -join "," | Should -Be "ページ001[埋め込み1]"
        $failures.Count | Should -Be 0
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

    It "スライドに埋め込んだ PowerPoint の文字を読み、その中の埋め込みは読まない（深さは 1 段まで）" {
        $innerSlide = "<p:sld $script:pNs><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>入れ子の題</a:t></a:r></a:p></p:txBody></p:sp>" +
            "<p:graphicFrame><a:graphic><a:graphicData><p:oleObj r:id=`"rId2`" progId=`"Excel.Sheet.12`"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"
        $inner = @{
            "ppt/presentation.xml"            = "<p:presentation $script:pNs><p:sldIdLst><p:sldId id=`"256`" r:id=`"rId1`"/></p:sldIdLst></p:presentation>"
            "ppt/_rels/presentation.xml.rels" = "<Relationships $script:relNs><Relationship Id=`"rId1`" Type=`"$script:officeRel/slide`" Target=`"slides/slide1.xml`"/></Relationships>"
            "ppt/slides/slide1.xml"           = $innerSlide
            "ppt/slides/_rels/slide1.xml.rels" = "<Relationships $script:relNs><Relationship Id=`"rId2`" Type=`"$script:officeRel/package`" Target=`"../embeddings/deep.xlsx`"/></Relationships>"
            "ppt/embeddings/deep.xlsx"        = (newZipBytes (newWorkbookEntries))
        }
        $slide = "<p:sld $script:pNs><p:cSld><p:spTree><p:graphicFrame><a:graphic><a:graphicData><p:oleObj r:id=`"rId2`" progId=`"PowerPoint.Show.12`"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"
        $path = "$TestDrive\slide_pptx.pptx"
        newZipFile $path @{
            "ppt/presentation.xml"            = "<p:presentation $script:pNs><p:sldIdLst><p:sldId id=`"256`" r:id=`"rId1`"/></p:sldIdLst></p:presentation>"
            "ppt/_rels/presentation.xml.rels" = "<Relationships $script:relNs><Relationship Id=`"rId1`" Type=`"$script:officeRel/slide`" Target=`"slides/slide1.xml`"/></Relationships>"
            "ppt/slides/slide1.xml"           = $slide
            "ppt/slides/_rels/slide1.xml.rels" = "<Relationships $script:relNs><Relationship Id=`"rId2`" Type=`"$script:officeRel/package`" Target=`"../embeddings/Microsoft_PowerPoint_Presentation.pptx`"/></Relationships>"
            "ppt/embeddings/Microsoft_PowerPoint_Presentation.pptx" = (newZipBytes $inner)
        }

        $units = readPptxUnits $path
        @($units.Keys) -contains "スライド001[埋め込み2]" | Should -Be $false  # 入れ子の中の埋め込みは読まない
        @($units["スライド001[埋め込み1]"]) -join "|" | Should -Be "入れ子の題"
    }
}
