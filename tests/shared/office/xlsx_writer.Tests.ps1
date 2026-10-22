# .xlsx の書き手（shared\office\xlsx_writer.ps1）のテスト。
# 書いたファイルを ZIP と XML として開いて確かめる（Excel は使わないので CI でも流れる）。
BeforeDiscovery {
    $high = [string][char]0xD83D
    $low = [string][char]0xDE00
    $ellipsis = [string][char]0x2026
    $limitCases = @(
        @{ name = "32,766 文字はそのまま"; text = ("a" * 32766); expectedLength = 32766; truncated = $false }
        @{ name = "32,767 文字はそのまま"; text = ("a" * 32767); expectedLength = 32767; truncated = $false }
        @{ name = "32,768 文字は切って … を付ける（合計 32,767）"; text = ("a" * 32768); expectedLength = 32767; truncated = $true }
        @{ name = "切る位置がサロゲートペアの途中なら 1 文字手前で切る"; text = (("a" * 32765) + $high + $low + "bb"); expectedLength = 32766; truncated = $true }
        @{ name = "置き換えで長さが変わる（CRLF が LF 1 つになり上限に収まる）"; text = (("a" * 32766) + "`r`n"); expectedLength = 32767; truncated = $false }
        @{ name = "置き換えのあとの長さで判定する（CRLF を LF にしても超えれば切る）"; text = (("a" * 32767) + "`r`n"); expectedLength = 32767; truncated = $true }
    )
    $replaceCases = @(
        @{ name = "CR は LF にそろえる"; text = "a`rb"; expected = "a`nb" }
        @{ name = "CRLF は LF 1 つ"; text = "a`r`nb"; expected = "a`nb" }
        @{ name = "タブ・LF は残す"; text = "a`tb`nc"; expected = "a`tb`nc" }
        @{ name = "制御文字はスペース"; text = ("a" + [char]0 + "b" + [char]0x1F + "c" + [char]0x0B); expected = "a b c " }
        @{ name = "U+FFFE・U+FFFF はスペース"; text = ("a" + [char]0xFFFE + [char]0xFFFF + "b"); expected = "a  b" }
        @{ name = "対の無い上位サロゲートはスペース"; text = ("a" + $high + "b"); expected = "a b" }
        @{ name = "対の無い下位サロゲートはスペース"; text = ("a" + $low + "b"); expected = "a b" }
        @{ name = "対のサロゲートは残す"; text = ("a" + $high + $low + "b"); expected = ("a" + $high + $low + "b") }
        @{ name = "空文字"; text = ""; expected = "" }
        @{ name = "null"; text = $null; expected = "" }
    )
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function readXlsxParts {
        # 部品名 → 中身（文字列）の辞書
        param ([string]$path)
        $parts = @{}
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            foreach ($entry in $zip.Entries) {
                $reader = New-Object System.IO.StreamReader($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
                try { $parts[$entry.FullName] = $reader.ReadToEnd() } finally { $reader.Dispose() }
            }
        } finally {
            $zip.Dispose()
        }
        return $parts
    }

    function readSheetRows {
        # sheet1.xml の各行を、セルの文字の配列にして返す
        param ($sheetXml)
        $rows = @()
        foreach ($row in $sheetXml.worksheet.sheetData.row) {
            $rows += , @(@($row.c) | ForEach-Object { [string]$_.is.t.'#text' })
        }
        return , $rows
    }

    function writeXlsx {
        param ([string]$path, [string[]]$headers, $rows, [int]$maxRows = 1048576)
        $w = newXlsxWriter -path $path -headers $headers -maxRows $maxRows
        foreach ($r in $rows) { [void](addXlsxRow $w $r) }
        return closeXlsxWriter $w
    }
}

Describe "limitXlsxCellText" -Tag Unit {
    It "<name>" -TestCases $limitCases {
        $r = limitXlsxCellText $text
        $r.Text.Length | Should -Be $expectedLength
        $r.Truncated | Should -Be $truncated
        if ($truncated) { $r.Text.EndsWith([string][char]0x2026) | Should -BeTrue }
    }

    It "切った末尾は、サロゲートの途中で終わらない" {
        $r = limitXlsxCellText (("a" * 32765) + [char]0xD83D + [char]0xDE00 + "bb")
        $r.Text | Should -Be (("a" * 32765) + [char]0x2026)
    }

    It "<name>" -TestCases $replaceCases {
        (limitXlsxCellText $text).Text | Should -Be $expected
    }
}

Describe "xlsx の書き手" -Tag Io {
    BeforeAll {
        $headers = @("ファイル", "場所", "行", "備考", "内容")
        $special = @(
            @("a.xlsx", "Sheet1", "1-2", "007", "=SUM(A1)"),
            @("b.xlsx", "x`ny", 'say "hi" <b> & c', "  前後の空白  ", ("c" + [char]0x1 + "d")),
            @("c.xlsx")
        )
        $path = Join-Path $TestDrive "basic.xlsx"
        $result = writeXlsx $path $headers $special
        $parts = readXlsxParts $path
        [xml]$sheet = $parts["xl/worksheets/sheet1.xml"]
        $rows = readSheetRows $sheet
    }

    It "部品がそろい、区切りは / で、共有文字列は無い" {
        $parts.Keys | Sort-Object | Should -Be @("[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml")
    }

    It "全部品が XML として読める" {
        foreach ($k in $parts.Keys) { { [xml]$parts[$k] } | Should -Not -Throw }
    }

    It "戻り値: 見出しを除く行数・切ったセル 0・落とした行 0" {
        $result.Count | Should -Be 3
        $result.TruncatedCells | Should -Be 0
        $result.DroppedRows | Should -Be 0
    }

    It "全セルが inlineStr で、数式を含まない" {
        $cells = @($sheet.worksheet.sheetData.row | ForEach-Object { $_.c })
        $cells.Count | Should -Be (4 * 5)
        @($cells | Where-Object { $_.t -ne "inlineStr" }).Count | Should -Be 0
        $parts["xl/worksheets/sheet1.xml"] | Should -Not -Match "<f[ >]"
    }

    It "sheet1.xml の要素の順は sheetViews → cols → sheetData → autoFilter" {
        @($sheet.worksheet.ChildNodes | ForEach-Object { $_.LocalName }) | Should -Be @("sheetViews", "cols", "sheetData", "autoFilter")
        $sheet.worksheet.cols.col.Count | Should -Be 5
        $sheet.worksheet.cols.col[4].width | Should -Be "80"
    }

    It "どの行も見出しと同じ数のセルを持つ（足りない分は空）" {
        foreach ($r in $rows) { $r.Count | Should -Be 5 }
        $rows[3] | Should -Be @("c.xlsx", "", "", "", "")
    }

    It "値が期待どおり（見出し・改行・記号・先頭の =・007・前後の空白・入れられない文字）" {
        $rows[0] | Should -Be $headers
        $rows[1] | Should -Be @("a.xlsx", "Sheet1", "1-2", "007", "=SUM(A1)")
        $rows[2] | Should -Be @("b.xlsx", "x`ny", 'say "hi" <b> & c', "  前後の空白  ", "c d")
    }

    It "見出しは見出しのスタイル、改行を含むセルにだけ折り返しのスタイル" {
        $r = $sheet.worksheet.sheetData.row
        @($r[0].c | ForEach-Object { $_.s }) | Should -Be @("1", "1", "1", "1", "1")
        @($r[1].c | ForEach-Object { $_.s }) | Should -Be @($null, $null, $null, $null, $null)
        $r[2].c[1].s | Should -Be "2"
        @($r[2].c | Where-Object { $_.s }).Count | Should -Be 1
        $x = [xml]$parts["xl/styles.xml"]
        $x.styleSheet.cellXfs.count | Should -Be "3"
        $x.styleSheet.cellXfs.xf[2].alignment.wrapText | Should -Be "1"
    }

    It "見出しの行の固定・オートフィルター・隠した定義名がある" {
        $pane = $sheet.worksheet.sheetViews.sheetView.pane
        $pane.ySplit | Should -Be "1"
        $pane.state | Should -Be "frozen"
        $sheet.worksheet.autoFilter.ref | Should -Be "A1:E4"
        $wb = [xml]$parts["xl/workbook.xml"]
        $wb.workbook.sheets.sheet.name | Should -Be "検索結果"
        $dn = $wb.workbook.definedNames.definedName
        $dn.name | Should -Be "_xlnm._FilterDatabase"
        $dn.localSheetId | Should -Be "0"
        $dn.hidden | Should -Be "1"
        $dn.'#text' | Should -Be "'検索結果'!`$A`$1:`$E`$4"
    }

    It "32,767 文字を超えるセルは切り、行の数は入力どおり" {
        $long = "x" * 32768
        $big = "y" * 40000
        $longRows = @(
            @("ok", "s", "1", "n", $long),
            @($big, $big, $big, $big, $big),
            @("ok2", "s", "2", "n", ("z" * 32767))
        )
        $p = Join-Path $TestDrive "long.xlsx"
        $res = writeXlsx $p $headers $longRows
        [xml]$s = (readXlsxParts $p)["xl/worksheets/sheet1.xml"]
        $r = readSheetRows $s
        $r.Count | Should -Be 4
        $res.Count | Should -Be 3
        $res.TruncatedCells | Should -Be 6
        $truncated = 0
        foreach ($row in $r) {
            foreach ($cell in $row) {
                $cell.Length | Should -BeLessOrEqual 32767
                if ($cell.EndsWith([string][char]0x2026)) { $truncated++ }
            }
        }
        $truncated | Should -Be $res.TruncatedCells
    }

    It "行の上限を超えた分は書かず、DroppedRows に数える" {
        $p = Join-Path $TestDrive "max.xlsx"
        $w = newXlsxWriter -path $p -headers @("a", "b") -maxRows 3
        (addXlsxRow $w @("1", "1")) | Should -BeTrue
        (addXlsxRow $w @("2", "2")) | Should -BeTrue
        (addXlsxRow $w @("3", "3")) | Should -BeFalse
        (addXlsxRow $w @("4", "4")) | Should -BeFalse
        $res = closeXlsxWriter $w
        $res.Count | Should -Be 2
        $res.DroppedRows | Should -Be 2
        [xml]$s = (readXlsxParts $p)["xl/worksheets/sheet1.xml"]
        (readSheetRows $s).Count | Should -Be 3
        $s.worksheet.autoFilter.ref | Should -Be "A1:B3"
    }

    It "保存先が共有なしで開かれていると IOException" {
        $p = Join-Path $TestDrive "locked.xlsx"
        $fs = New-Object System.IO.FileStream($p, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            { newXlsxWriter -path $p -headers @("a") } | Should -Throw -ExceptionType ([System.IO.IOException])
        } finally {
            $fs.Dispose()
        }
    }

    It "読み取り専用だと UnauthorizedAccessException" {
        $p = Join-Path $TestDrive "readonly.xlsx"
        [System.IO.File]::WriteAllText($p, "x")
        Set-ItemProperty -LiteralPath $p -Name IsReadOnly -Value $true
        try {
            { newXlsxWriter -path $p -headers @("a") } | Should -Throw -ExceptionType ([System.UnauthorizedAccessException])
        } finally {
            Set-ItemProperty -LiteralPath $p -Name IsReadOnly -Value $false
        }
    }

    It "abortXlsxWriter はファイルを閉じる（その後に消せる）" {
        $p = Join-Path $TestDrive "abort.xlsx"
        $w = newXlsxWriter -path $p -headers @("a")
        [void](addXlsxRow $w @("1"))
        abortXlsxWriter $w
        { Remove-Item -LiteralPath $p -ErrorAction Stop } | Should -Not -Throw
        Test-Path -LiteralPath $p | Should -BeFalse
    }
}

Describe "xlsx の部品の組み立て" -Tag Unit {
    It "getXlsxColumnName は列番号を A, B, ..., Z, AA の形にする" -TestCases @(
        @{ index = 1; expected = "A" }
        @{ index = 5; expected = "E" }
        @{ index = 26; expected = "Z" }
        @{ index = 27; expected = "AA" }
        @{ index = 52; expected = "AZ" }
        @{ index = 703; expected = "AAA" }
    ) {
        getXlsxColumnName $index | Should -Be $expected
    }

    It "escapeXlsxText は XML の特別な文字だけを置き換える" {
        escapeXlsxText "a&b<c>d`"e" | Should -Be "a&amp;b&lt;c&gt;d`"e"
    }

    It "addXlsxEntry は ZIP に UTF-8 の部品を 1 つ書く" {
        $stream = New-Object System.IO.MemoryStream
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        addXlsxEntry $zip "xl/a.xml" "あ<x/>"
        $zip.Dispose()
        $read = New-Object System.IO.Compression.ZipArchive((New-Object System.IO.MemoryStream(, $stream.ToArray())), [System.IO.Compression.ZipArchiveMode]::Read)
        $reader = New-Object System.IO.StreamReader($read.GetEntry("xl/a.xml").Open(), (New-Object System.Text.UTF8Encoding($false)))
        try { $reader.ReadToEnd() | Should -Be "あ<x/>" } finally { $reader.Dispose(); $read.Dispose() }
    }

    It "writeXlsxRowXml は見出しに見出しのスタイル、改行を含むセルに折り返しのスタイルを当て、足りない列は空にする" {
        $sheet = New-Object System.IO.StringWriter
        $writer = @{ Headers = @("a", "b", "c"); Columns = @("A", "B", "C"); Sheet = $sheet; TruncatedCells = 0 }
        writeXlsxRowXml $writer @("x", "1`n2") 1 $true
        writeXlsxRowXml $writer @("x&", "1`n2") 2 $false
        $xml = $sheet.ToString()
        $xml | Should -Match '<row r="1"><c r="A1" t="inlineStr" s="1">'
        $xml | Should -Match '<row r="2"><c r="A2" t="inlineStr"><is><t xml:space="preserve">x&amp;</t>'
        $xml | Should -Match '<c r="B2" t="inlineStr" s="2">'
        $xml | Should -Match '<c r="C2" t="inlineStr"><is><t xml:space="preserve"></t>'
    }
}
