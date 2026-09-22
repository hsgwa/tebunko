# 1 ファイルの抽出（tebunko_grep\indexer\extract_office.ps1）のテスト。
# Excel・Word・PowerPoint は使わない。COM の入口の getApp を Mock して、同じ呼び方ができる偽のオブジェクトを返す。
# 偽のオブジェクトは、呼ばれたメソッドと引数を $log に記録し、保存（SaveAs）では本物と同じ形式のファイルを書く。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\office\office_reader.ps1"
. "${scriptsDir}\shared\office\office_app.ps1"
. "${scriptsDir}\tebunko_grep\indexer\extract_office.ps1"

# Worksheets は、foreach で列挙でき、引数なしの Add() で一時シートを足せる。
# クラスのメソッドからはテストの変数・関数が見えないため、必要なものはプロパティに持たせる
class FakeSheets : System.Collections.IEnumerable {
    [System.Collections.Generic.List[object]]$Items = (New-Object System.Collections.Generic.List[object])
    [object]$Temp
    [bool]$Protected
    [System.Collections.ArrayList]$Log
    [object] Add() {
        if ($this.Protected) { throw "ブックの構成は保護されているため変更できません。" }
        [void]$this.Log.Add("AddSheet")
        return $this.Temp
    }
    [System.Collections.IEnumerator] GetEnumerator() { return $this.Items.GetEnumerator() }
}

$log = New-Object System.Collections.ArrayList

function newFake([hashtable]$properties, [hashtable]$methods) {
    $fake = New-Object psobject -Property $properties
    foreach ($name in $methods.Keys) {
        $fake | Add-Member -MemberType ScriptMethod -Name $name -Value $methods[$name]
    }
    return $fake
}

function newRange([int]$row, [int]$column, [int]$rows = 1, [int]$columns = 1) {
    return newFake @{ Row = $row; Column = $column; Rows = @{ Count = $rows }; Columns = @{ Count = $columns } } @{}
}

# 偽のシート。SaveAs では Excel の「Unicode テキスト」と同じ UTF-16（BOM 付き）で $text を書く。
#   used      : 使用範囲（UsedRange）の @(行, 列, 行数, 列数)
#   dataLast  : 値のある最後のセルの @(行, 列)（Cells.Find の結果）
function newSheet([string]$name, [int]$visible, [string]$text, [int[]]$used = @(1, 1, 1, 1), [int[]]$dataLast = @(1, 1)) {
    $cells = newFake @{} @{
        Find = {
            # 引数の 5 番目が検索の向き（1 = 行ごと → 最後の行、2 = 列ごと → 最後の列）
            if ($args[4] -eq 1) { return newRange $this.DataLast[0] 1 }
            return newRange 1 $this.DataLast[1]
        }
        Item = { return newRange $args[0] $args[1] }
    }
    $cells | Add-Member -NotePropertyName DataLast -NotePropertyValue $dataLast
    $sheet = newFake @{
        Name = $name; Visible = $visible; Index = 0; Text = $text
        UsedRange = (newRange $used[0] $used[1] $used[2] $used[3]); Cells = $cells
    } @{
        Range = {
            $range = newRange 1 1
            $range | Add-Member -NotePropertyName SheetName -NotePropertyValue $this.Name
            $range | Add-Member -MemberType ScriptMethod -Name Copy -Value { [void]$log.Add("Copy:$($this.SheetName)") }
            return $range
        }
        Activate = { [void]$log.Add("Activate:$($this.Name)") }
        SaveAs = {
            [void]$log.Add("SaveAs:$($this.Name):$($args[1])")
            [System.IO.File]::WriteAllText($args[0], $this.Text, [System.Text.Encoding]::Unicode)
        }
        Delete = { [void]$log.Add("Delete:$($this.Name)") }
    }
    return $sheet
}

function newExcel([object[]]$sheets, [switch]$protectedStructure) {
    $worksheets = New-Object FakeSheets
    $index = 0
    foreach ($sheet in $sheets) {
        $index++
        $sheet.Index = $index
        $worksheets.Items.Add($sheet)
    }
    $worksheets.Temp = newSheet "一時" -1 "`t`tデータ`r`n"
    $worksheets.Protected = [bool]$protectedStructure
    $worksheets.Log = $log
    $workbook = newFake @{ Worksheets = $worksheets } @{
        Close = { [void]$log.Add("Close:$($args[0])") }
    }
    $workbooks = newFake @{ Book = $workbook } @{
        Open = {
            [void]$log.Add("Open:$([System.IO.Path]::GetFileName($args[0])):ReadOnly=$($args[2]):Password=$($args[4])")
            return $this.Book
        }
    }
    return newFake @{ Workbooks = $workbooks } @{}
}

# 偽の Word・PowerPoint が「新形式で保存」するときに書く、最小限の .docx / .pptx
function writeMinimalDocx([string]$path, [string]$text) {
    $stream = [System.IO.File]::Create($path)
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $writer = New-Object System.IO.StreamWriter($zip.CreateEntry("word/document.xml").Open(), (New-Object System.Text.UTF8Encoding($false)))
        $writer.Write("<w:document xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`"><w:body><w:p><w:r><w:t>$text</w:t></w:r></w:p></w:body></w:document>")
        $writer.Dispose()
    } finally {
        $zip.Dispose()
        $stream.Dispose()
    }
}

function writeMinimalPptx([string]$path, [string]$text) {
    $pNs = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $relNs = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
    $slideRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    $entries = [ordered]@{
        "ppt/presentation.xml"            = "<p:presentation $pNs><p:sldIdLst><p:sldId id=`"256`" r:id=`"rId2`"/></p:sldIdLst></p:presentation>"
        "ppt/_rels/presentation.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId2`" Type=`"$slideRel`" Target=`"slides/slide1.xml`"/></Relationships>"
        "ppt/slides/slide1.xml"           = "<p:sld $pNs><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id=`"2`" name=`"s`"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>$text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"
    }
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

function newWord([string]$text) {
    $doc = newFake @{ Text = $text } @{
        Repaginate = { [void]$log.Add("Repaginate") }
        SaveAs2 = {
            [void]$log.Add("SaveAs2:$([System.IO.Path]::GetFileName($args[0])):$($args[1])")
            writeMinimalDocx $args[0] $this.Text
        }
        Close = { [void]$log.Add("Close:$($args[0])") }
    }
    $documents = newFake @{ Doc = $doc } @{
        Open = {
            [void]$log.Add("Open:$([System.IO.Path]::GetFileName($args[0])):ReadOnly=$($args[2]):Password=$($args[4]):Visible=$($args[11])")
            return $this.Doc
        }
    }
    return newFake @{ Documents = $documents } @{}
}

function newPowerPoint([string]$text) {
    $pres = newFake @{ Text = $text } @{
        SaveAs = {
            [void]$log.Add("SaveAs:$([System.IO.Path]::GetFileName($args[0])):$($args[1])")
            writeMinimalPptx $args[0] $this.Text
        }
        Close = { [void]$log.Add("Close") }
    }
    $presentations = newFake @{ Pres = $pres } @{
        Open = {
            # パスの後ろに ::dummy:: が付くため、GetFileName ではなく最後の \ から後ろを取る
            [void]$log.Add("Open:$([regex]::Match($args[0], '[^\\]+$').Value):$($args[1]),$($args[2]),$($args[3])")
            return $this.Pres
        }
    }
    return newFake @{ Presentations = $presentations } @{}
}

function readTsv([string]$name) {
    return [System.IO.File]::ReadAllText((Join-Path $tmpDir $name))
}

function listTmp {
    return @(Get-ChildItem -LiteralPath $tmpDir -File | ForEach-Object { $_.Name } | Sort-Object)
}

Describe "extractWorkbook（偽の Excel）" -Tag Io {
    $tmpDir = Join-Path $TestDrive "tmp"
    $source = Join-Path $TestDrive "見積[1].xlsx"
    [System.IO.File]::WriteAllText($source, "元のファイル")

    BeforeEach {
        $log.Clear()
        if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force }
        [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
    }

    Mock releaseComObject {}

    It "表示しているシートだけを書き出し、空のシートは TSV にしない" {
        $excel = newExcel @(
            (newSheet "売上" -1 "品名`t金額`r`nりんご`t100`r`n"),
            (newSheet "非表示" 0 "隠し`r`n"),
            (newSheet "完全に非表示" 2 "隠し`r`n"),
            (newSheet "空" -1 "")
        )
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        extractWorkbook $source | Should Be 1
        listTmp | Should Be @("売上.tsv")
        readTsv "売上.tsv" | Should Be "品名`t金額`r`nりんご`t100`r`n"
        @($log | Where-Object { $_ -like "SaveAs:*" }) -join "|" | Should Be "SaveAs:売上:42|SaveAs:空:42"
    }

    It "作業フォルダのコピーを読み取り専用・ダミーのパスワードで開き、保存せずに閉じてコピーを消す" {
        $excel = newExcel @((newSheet "Sheet1" -1 "a`r`n"))
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        [void](extractWorkbook $source)
        # 元と同じファイル名のコピーを開く（CELL("filename") の表示値を変えないため）
        $log[0] | Should Be "Open:見積[1].xlsx:ReadOnly=True:Password=dummy"
        $log[-1] | Should Be "Close:False"
        Test-Path -LiteralPath (Join-Path $tmpDir "見積[1].xlsx") | Should Be $false
        [System.IO.File]::ReadAllText($source) | Should Be "元のファイル"
    }

    It "使用範囲の左上が A1 でないシートは、A1 からの位置に戻す" {
        $excel = newExcel @((newSheet "D5から" -1 "a`tb`r`nc`td`r`n" @(2, 3, 2, 2) @(3, 4)))
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        [void](extractWorkbook $source)
        readTsv "D5から.tsv" | Should Be "`r`n`t`ta`tb`r`n`t`tc`td`r`n"
    }

    It "使用範囲がデータよりずっと広いシートは、データの範囲だけを一時シートにコピーして書き出し、一時シートを消す" {
        # 使用範囲は 1,048,576 行 × 2 列、データは 3 行 × 2 列
        $excel = newExcel @((newSheet "肥大" -1 "元のシート`r`n" @(1, 1, 1048576, 2) @(3, 2)))
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        extractWorkbook $source | Should Be 1
        ($log | Where-Object { $_ -notlike "Open:*" -and $_ -notlike "Close:*" }) -join "|" |
            Should Be "AddSheet|Copy:肥大|Activate:一時|SaveAs:一時:42|Delete:一時"
        readTsv "肥大.tsv" | Should Be "`t`tデータ`r`n"
    }

    It "一時シートを足せない（ブックの構成が保護されている）ときは、元のシートをそのまま書き出す" {
        $excel = newExcel @((newSheet "肥大" -1 "元のシート`r`n" @(1, 1, 1048576, 2) @(3, 2))) -protectedStructure
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        extractWorkbook $source | Should Be 1
        readTsv "肥大.tsv" | Should Be "元のシート`r`n"
    }

    It "使用範囲とデータの差が小さいシートは、一時シートを使わない" {
        $excel = newExcel @((newSheet "普通" -1 "a`r`n" @(1, 1, 100, 10) @(90, 10)))
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        [void](extractWorkbook $source)
        @($log | Where-Object { $_ -eq "AddSheet" }).Count | Should Be 0
    }

    It "新形式（ZIP）のブックは、図形・コメントの文字を別の場所の TSV にする" {
        $zipSource = Join-Path $TestDrive "図形あり.xlsx"
        $xNs = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        $relNs = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
        $officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        $entries = [ordered]@{
            "xl/workbook.xml" = "<workbook $xNs><sheets><sheet name=`"売上`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            "xl/_rels/workbook.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
            "xl/worksheets/sheet1.xml" = "<worksheet $xNs/>"
            "xl/worksheets/_rels/sheet1.xml.rels" = "<Relationships $relNs><Relationship Id=`"rId1`" Type=`"$officeRel/comments`" Target=`"../comments1.xml`"/></Relationships>"
            "xl/comments1.xml" = "<comments $xNs><commentList><comment ref=`"B2`"><text><t>税抜</t></text></comment></commentList></comments>"
        }
        $stream = [System.IO.File]::Create($zipSource)
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($name in $entries.Keys) {
            $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open(), (New-Object System.Text.UTF8Encoding($false)))
            $writer.Write($entries[$name])
            $writer.Dispose()
        }
        $zip.Dispose()
        $stream.Dispose()

        $excel = newExcel @((newSheet "売上" -1 "品名`r`n"))
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        extractWorkbook $zipSource | Should Be 2
        listTmp | Should Be @("売上.tsv", "売上[コメント].tsv")
        readTsv "売上[コメント].tsv" | Should Be "B2`t税抜`r`n"
    }

    It "図形・コメントを読めなくても、セルの値は取り込む" {
        $broken = Join-Path $TestDrive "壊れたZIP.xlsx"
        [System.IO.File]::WriteAllBytes($broken, [byte[]](0x50, 0x4B, 0x03, 0x04, 0x00))
        $excel = newExcel @((newSheet "売上" -1 "品名`r`n"))
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        extractWorkbook $broken | Should Be 1
        listTmp | Should Be @("売上.tsv")
    }

    It "開けない（パスワード付きなど）ときは例外を返す" {
        $workbooks = newFake @{} @{ Open = { throw "パスワードが正しくありません。" } }
        $excel = newFake @{ Workbooks = $workbooks } @{}
        Mock getApp { $excel } -ParameterFilter { $name -eq "Excel" }

        { extractWorkbook $source } | Should Throw "パスワードが正しくありません。"
    }
}

Describe "extractDocument（偽の Word・PowerPoint）" -Tag Io {
    $tmpDir = Join-Path $TestDrive "tmp"

    BeforeEach {
        $log.Clear()
        if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force }
        [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
    }

    Mock releaseComObject {}

    It "新形式（.docx）は Word を使わずに読む" {
        $source = Join-Path $TestDrive "報告書.docx"
        writeMinimalDocx $source "新形式の本文"
        Mock getApp { throw "Word を起動してはいけない" }

        extractDocument $source | Should Be 1
        (readTsv "ページ001.tsv").Trim() | Should Be "新形式の本文"
        Assert-MockCalled getApp -Times 0 -Exactly -Scope It
    }

    It "旧形式（.doc）は Word で読み取り専用・ウィンドウ無しで開き、ページ割りを確定させてから .docx で保存して読む" {
        $source = Join-Path $TestDrive "旧形式.doc"
        [System.IO.File]::WriteAllBytes($source, [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00))
        $word = newWord "旧形式の本文"
        Mock getApp { $word } -ParameterFilter { $name -eq "Word" }

        extractDocument $source | Should Be 1
        (readTsv "ページ001.tsv").Trim() | Should Be "旧形式の本文"
        $log -join "|" | Should Be "Open:source.doc:ReadOnly=True:Password=dummy:Visible=False|Repaginate|SaveAs2:converted.docx:12|Close:0"
        # 作業ファイル（コピーと TSV）は消す
        listTmp | Should Be @("ページ001.tsv")
    }

    It "拡張子と中身が違うファイル（中身が旧形式の .docx）は、旧形式の拡張子を付け直して Word で開く" {
        $source = Join-Path $TestDrive "中身はdoc.docx"
        [System.IO.File]::WriteAllBytes($source, [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00))
        $word = newWord "中身は旧形式"
        Mock getApp { $word } -ParameterFilter { $name -eq "Word" }

        extractDocument $source | Should Be 1
        $log[0] | Should BeLike "Open:source.doc:*"
    }

    It "旧形式（.ppt）は PowerPoint でウィンドウ無し・ダミーのパスワード付きで開き、.pptx で保存して読む" {
        $source = Join-Path $TestDrive "提案.ppt"
        [System.IO.File]::WriteAllBytes($source, [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00))
        $ppt = newPowerPoint "旧形式のスライド"
        Mock getApp { $ppt } -ParameterFilter { $name -eq "PowerPoint" }

        extractDocument $source | Should Be 1
        (readTsv "スライド001.tsv").Trim() | Should Be "旧形式のスライド"
        # ファイル名の後ろの ::dummy:: で、パスワード付きのファイルはダイアログを出さずにエラーになる
        $log -join "|" | Should Be "Open:source.ppt::dummy:::-1,0,0|SaveAs:converted.pptx:24|Close"
        listTmp | Should Be @("スライド001.tsv")
    }

    It "新形式でも旧形式でもない .pptx は、PowerPoint に渡さずに失敗にする（作業ファイルは消す）" {
        $source = Join-Path $TestDrive "壊れたファイル.pptx"
        [System.IO.File]::WriteAllText($source, "中身はテキスト")
        Mock getApp { throw "PowerPoint を起動してはいけない" }

        { extractDocument $source } | Should Throw "PowerPointのファイルではありません"
        Assert-MockCalled getApp -Times 0 -Exactly -Scope It
        (listTmp).Count | Should Be 0
    }

    It "Word が保存に失敗しても、文書を閉じて作業ファイルを消す" {
        $source = Join-Path $TestDrive "保存失敗.doc"
        [System.IO.File]::WriteAllBytes($source, [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00))
        $word = newWord ""
        $word.Documents.Doc | Add-Member -MemberType ScriptMethod -Name SaveAs2 -Value { throw "保存できません。" } -Force
        Mock getApp { $word } -ParameterFilter { $name -eq "Word" }

        { extractDocument $source } | Should Throw "保存できません。"
        $log[-1] | Should Be "Close:0"
        (listTmp).Count | Should Be 0
    }
}

Describe "ingestFile" -Tag Io {
    Mock extractWorkbook { "Excel" }
    Mock extractDocument { "Word・PowerPoint" }

    It "拡張子で Excel とそれ以外に振り分ける" {
        ingestFile "a.xlsx" | Should Be "Excel"
        ingestFile "a.XLS" | Should Be "Excel"
        ingestFile "a.docx" | Should Be "Word・PowerPoint"
        ingestFile "a.ppt" | Should Be "Word・PowerPoint"
    }
}
