# tebunko_grep の画面で使う型（tebunko_grep\ui\types_grep.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\ui\types.ps1"
. "${scriptsDir}\tebunko_grep\ui\types_grep.ps1"

# PropertyChanged で通知されたプロパティ名を集める
function watchChanges($target) {
    $names = New-Object System.Collections.Generic.List[string]
    $target.add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]{ param($sender, $e) $names.Add($e.PropertyName) }.GetNewClosure())
    , $names
}

# 検索結果の 1 行を作る（Create の引数を名前で渡すため）
function newHitRow {
    param (
        [string]$book = "見積.xlsx",
        [string]$location = "[シート] 4月",
        [int]$lineNumber = 3,
        [string]$line = "",
        [string]$word = "",
        [regex]$pattern = $null,
        [string]$relDir = "営業部"
    )
    [HitRow]::Create("営業", "C:\index\営業", "$relDir\$book", $relDir, "$book.tsv", $book, $location, $lineNumber, $line, $word, $pattern)
}

# セグメントを「文字」と「強調の有無」の組にして比べやすくする
function describeSegments($segments) {
    @($segments | ForEach-Object { if ($_.IsHit) { "[$($_.Text)]" } else { $_.Text } }) -join ""
}

Describe "Segment" -Tag Unit {
    It "文字と強調の有無を持つ" {
        $segment = [Segment]::new("見積", $true)
        $segment.Text | Should Be "見積"
        $segment.IsHit | Should Be $true
        [Segment]::new().IsHit | Should Be $false
    }
}

Describe "PreviewColumn" -Tag Unit {
    It "幅を変えると通知する" {
        $column = [PreviewColumn]::new()
        $names = watchChanges $column
        $column.SetWidth(100)
        $column.Width | Should Be 100
        @($names) | Should Be @("Width")
    }

    It "下限より狭くしない" {
        $column = [PreviewColumn]::new()
        $column.SetWidth(3)
        $column.Width | Should Be ([PreviewColumn]::MinWidth)
    }

    It "同じ幅なら通知しない" {
        $column = [PreviewColumn]::new()
        $column.SetWidth(100)
        $names = watchChanges $column
        $column.SetWidth(100)
        $names.Count | Should Be 0
    }
}

Describe "PreviewCell" -Tag Unit {
    It "選択は変わったときだけ通知する" {
        $cell = [PreviewCell]::new()
        $names = watchChanges $cell
        $cell.SetSelected($true)
        $cell.SetSelected($true)
        $cell.IsSelected | Should Be $true
        @($names) | Should Be @("IsSelected")
    }
}

Describe "PreviewTable（範囲の選択とコピー）" -Tag Unit {
    # 3 行 × 3 列の表（A1〜C3 の値を持つ）
    function newTable {
        $row = newHitRow -lineNumber 2 -line "A2`tB2`tC2" -word "B2"
        $row.BuildPreview(@(1, 2, 3), @("A1`tB1`tC1", "A2`tB2`tC2", "A3`tB3`tC3"))
    }

    It "何も選んでいなければ空" {
        $table = newTable
        $table.HasSelection() | Should Be $false
        $table.GetSelectionText() | Should Be ""
        $table.SelectedCount() | Should Be 0
    }

    It "1 つのセルを選ぶとその値をそのまま返す" {
        $table = newTable
        $table.Select($table.Rows[1].Cells[1], $false)
        $table.HasSelection() | Should Be $true
        $table.GetSelectionText() | Should Be "B2"
        $table.SelectedCount() | Should Be 1
        $table.Rows[1].Cells[1].IsSelected | Should Be $true
        $table.Rows[0].Cells[0].IsSelected | Should Be $false
    }

    It "広げて選ぶと、行はCRLF・列はタブでつなぐ" {
        $table = newTable
        $table.Select($table.Rows[2].Cells[2], $false)
        $table.Select($table.Rows[1].Cells[1], $true)
        $table.GetSelectionText() | Should Be "B2`tC2`r`nB3`tC3"
        $table.SelectedCount() | Should Be 4
        @($table.Rows | ForEach-Object { $_.Cells } | Where-Object { $_.IsSelected }).Count | Should Be 4
    }

    It "選んでいないときに広げると、そのセルから始める" {
        $table = newTable
        $table.Select($table.Rows[0].Cells[0], $true)
        $table.GetSelectionText() | Should Be "A1"
    }

    It "広げずに選び直すと、前の選択を外す" {
        $table = newTable
        $table.Select($table.Rows[0].Cells[0], $false)
        $table.Select($table.Rows[2].Cells[2], $false)
        $table.Rows[0].Cells[0].IsSelected | Should Be $false
        $table.GetSelectionText() | Should Be "C3"
    }

    It "行を選ぶと、その行のすべての列" {
        $table = newTable
        $table.SelectRow($table.Rows[0].Cells[1])
        $table.GetSelectionText() | Should Be "A1`tB1`tC1"
        $table.SelectedCount() | Should Be 3
    }

    It "選択を外す" {
        $table = newTable
        $table.SelectRow($table.Rows[0].Cells[1])
        $table.ClearSelection()
        $table.HasSelection() | Should Be $false
        @($table.Rows | ForEach-Object { $_.Cells } | Where-Object { $_.IsSelected }).Count | Should Be 0
    }

    It "セルが無ければ何もしない" {
        $table = newTable
        $table.Select($null, $false)
        $table.SelectRow($null)
        $table.HasSelection() | Should Be $false
    }

    It "列の無い表では行を選べない" {
        $table = [PreviewTable]::new()
        $table.Columns = [System.Collections.Generic.List[PreviewColumn]]::new()
        $table.Rows = [System.Collections.Generic.List[PreviewRow]]::new()
        $table.SelectRow([PreviewCell]@{ RowIndex = 0; ColumnIndex = 0 })
        $table.HasSelection() | Should Be $false
    }

    It "表に無いセルは空として扱う" {
        $table = newTable
        $table.Select([PreviewCell]@{ RowIndex = 5; ColumnIndex = 0 }, $false)
        $table.GetSelectionText() | Should Be ""
        $table.Select([PreviewCell]@{ RowIndex = 0; ColumnIndex = 9 }, $false)
        $table.GetSelectionText() | Should Be ""
    }

    It "複数のセルをコピーするとき、改行・タブ・引用符を含むセルは Excel の形で囲む" {
        $row = newHitRow -lineNumber 1 -line "x" -word "x"
        $table = $row.BuildPreview(@(1), @("A`t`"B`"`"`"`tC"))
        # 1 列目は改行を含むセルにする
        $table.Rows[0].Cells[0].Text = "1行目`n2行目"
        $table.Select($table.Rows[0].Cells[0], $false)
        $table.Select($table.Rows[0].Cells[2], $true)
        $table.GetSelectionText() | Should Be "`"1行目`n2行目`"`t`"B`"`"`"`tC"
    }
}

Describe "PreviewTable.QuoteForExcel" -Tag Unit {
    It "改行・タブ・引用符が無ければそのまま" {
        [PreviewTable]::QuoteForExcel("見積") | Should Be "見積"
    }

    It "改行・タブを含めば引用符で囲む" {
        [PreviewTable]::QuoteForExcel("a`nb") | Should Be "`"a`nb`""
        [PreviewTable]::QuoteForExcel("a`tb") | Should Be "`"a`tb`""
    }

    It "引用符は 2 つにして囲む" {
        [PreviewTable]::QuoteForExcel('a"b') | Should Be '"a""b"'
    }
}

Describe "FileGroup" -Tag Unit {
    It "場所は初めてのときだけ $true" {
        $group = [FileGroup]::new()
        $group.AddLocation("4月") | Should Be $true
        $group.AddLocation("4月") | Should Be $false
        $group.AddLocation("5月") | Should Be $true
    }

    It "同じ表記は 1 つにし、足した順に返す" {
        $group = [FileGroup]::new()
        $group.AddLabel("[シート] 4月") | Should Be $true
        $group.AddLabel("[シート] 5月") | Should Be $true
        $group.AddLabel("[シート] 4月") | Should Be $false
        @($group.GetLocations()) | Should Be @("[シート] 4月", "[シート] 5月")
    }

    It "表記が無ければ空" {
        @([FileGroup]::new().GetLocations()).Count | Should Be 0
    }

    It "見出しの形で、閉じた状態で作る" {
        $group = [FileGroup]::new()
        $group.IsFileHeader | Should Be $true
        $group.IsExpanded | Should Be $false
        $group.Hits.Count | Should Be 0
        $group.Rows.Count | Should Be 0
        $group.ShownRows.Count | Should Be 0
    }

    It "場所の表記は毎回通知する" {
        $group = [FileGroup]::new()
        $names = watchChanges $group
        $group.SetLocationText("[シート] 4月")
        $group.SetLocationText("[シート] 4月")
        $group.LocationText | Should Be "[シート] 4月"
        @($names) | Should Be @("LocationText", "LocationText")
    }

    It "開閉・件数は変わったときだけ通知する" {
        $group = [FileGroup]::new()
        $names = watchChanges $group
        $group.SetExpanded($true)
        $group.SetExpanded($true)
        $group.SetCount(3)
        $group.SetCount(3)
        $group.IsExpanded | Should Be $true
        $group.ShownCount | Should Be 3
        @($names) | Should Be @("IsExpanded", "ShownCount")
    }
}

Describe "HitRow.Create" -Tag Unit {
    It "生データを入れ、Excel かどうか・図形の場所かどうかを決める" {
        $row = newHitRow -book "見積.xlsx" -location "[シート] 4月[図形]" -line "A1`t見積"
        $row.IndexName | Should Be "営業"
        $row.Root | Should Be "C:\index\営業"
        $row.RelPath | Should Be "営業部\見積.xlsx"
        $row.FileName | Should Be "見積.xlsx.tsv"
        $row.IsExcel | Should Be $true
        $row.IsObjectPlace | Should Be $true
        $row.Prepared | Should Be $false
        $row.IsFileHeader | Should Be $false
    }

    It "Excel の拡張子は大文字小文字を区別しない" {
        (newHitRow -book "見積.XLS").IsExcel | Should Be $true
        (newHitRow -book "見積.xlsm").IsExcel | Should Be $true
        (newHitRow -book "議事録.docx").IsExcel | Should Be $false
    }

    It "コメントの場所も図形と同じく扱う" {
        (newHitRow -location "[シート] 4月[コメント]").IsObjectPlace | Should Be $true
        (newHitRow -location "[シート] 4月").IsObjectPlace | Should Be $false
    }

    It "空の値は空文字にする" {
        $row = [HitRow]::Create("営業", "C:\index\営業", "見積.xlsx", $null, "見積.xlsx.tsv", $null, $null, 1, $null, "見積", $null)
        $row.RelDir | Should Be ""
        $row.Line | Should Be ""
        $row.IsExcel | Should Be $false
        $row.IsObjectPlace | Should Be $false
    }
}

Describe "HitRow.Prepare" -Tag Unit {
    It "Excel では最初に一致したセルの番地と、ほかの件数を出す" {
        $row = newHitRow -lineNumber 7 -line "品名`t見積A`t単価`t見積B" -word "見積"
        $row.Prepare()
        $row.Prepared | Should Be $true
        $row.MatchCell | Should Be "B7"
        $row.CellText | Should Be "B7 ほか 1"
        $row.DisplayLine | Should Be "品名 │ 見積A │ 単価 │ 見積B"
        describeSegments $row.Segments | Should Be "品名`t[見積]A`t単価`t[見積]B".Replace("`t", " │ ")
    }

    It "一致が 1 つなら番地だけ" {
        $row = newHitRow -lineNumber 2 -line "品名`t見積" -word "見積"
        $row.Prepare()
        $row.CellText | Should Be "B2"
    }

    It "Excel でなければセル番地を出さない" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "見積を送付" -word "見積"
        $row.Prepare()
        $row.MatchCell | Should Be ""
        $row.CellText | Should Be ""
        describeSegments $row.Segments | Should Be "[見積]を送付"
    }

    It "Excel の図形の行は、先頭のセル番地を出し、文字だけを表示する" {
        $row = newHitRow -location "[シート] 4月[図形]" -line "B3`t`"納期は$([char]0x2028)別途`"" -word "納期"
        $row.Prepare()
        $row.MatchCell | Should Be "B3"
        $row.CellText | Should Be "B3"
        $row.DisplayLine | Should Be "納期は$([char]0x21b5)別途"
    }

    It "図形の行で一致しなければセル番地を出さない" {
        $row = newHitRow -location "[シート] 4月[図形]" -line "B3`t納期" -word "見積"
        $row.Prepare()
        $row.MatchCell | Should Be ""
    }

    It "図形の行でセルが 1 つしかなければ行のまま表示する" {
        $row = newHitRow -location "[シート] 4月[図形]" -line "納期" -word "納期"
        $row.Prepare()
        $row.DisplayLine | Should Be "納期"
    }

    It "2 回目は作り直さない" {
        $row = newHitRow -line "見積" -word "見積"
        $row.Prepare()
        $row.DisplayLine = "変えた"
        $row.Prepare()
        $row.DisplayLine | Should Be "変えた"
    }

    It "表示用の項目を作ったことを通知する" {
        $row = newHitRow -line "見積" -word "見積"
        $names = watchChanges $row
        $row.Prepare()
        @($names) | Should Be @("DisplayLine", "Segments", "CellText", "MatchCell")
    }

    It "正規表現の一致を強調する" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "第1回と第22回" -pattern ([regex]"第\d+回")
        $row.Prepare()
        describeSegments $row.Segments | Should Be "[第1回]と[第22回]"
    }

    It "一致より前が長いときは、先頭を省いて一致の直前だけ残す" {
        $lead = "あ" * 50
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "${lead}見積" -word "見積"
        $row.Prepare()
        $row.Segments[0].Text | Should Be ([string][char]0x2026 + ("あ" * [HitRow]::LeadLength))
        $row.Segments[1].IsHit | Should Be $true
    }

    It "一致の後が長いときは切り詰める" {
        $rest = "い" * ([HitRow]::MaxDisplay + 10)
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "見積$rest" -word "見積"
        $row.Prepare()
        $row.Segments[1].Text | Should Be (("い" * [HitRow]::MaxDisplay) + [char]0x2026)
    }

    It "表示する長さを超えたら、その先の一致は強調しない" {
        $line = (("う" * 100) + "見積") * 10
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line $line -word "見積"
        $row.Prepare()
        @($row.Segments | Where-Object { $_.IsHit }).Count | Should BeLessThan 10
    }

    It "重なった一致は先のものだけ強調する" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "ああああ" -pattern ([regex]"(?=(ああ))あ")
        $row.Prepare()
        describeSegments $row.Segments | Should Be "[あ][あ][あ]あ"
    }
}

Describe "HitRow.Contains" -Tag Unit {
    $row = newHitRow -relDir "営業部" -book "見積.xlsx" -location "[シート] 4月" -lineNumber 12 -line "品名`tABC"
    $row.PlaceText = "[シート] 四月"
    $row.Kind = "セル"

    It "<name>で絞り込める" -TestCases @(
        @{ name = "相対フォルダ"; text = "営業" }
        @{ name = "元ファイル名"; text = "見積.XLSX" }
        @{ name = "場所"; text = "4月" }
        @{ name = "場所の表記"; text = "四月" }
        @{ name = "種別"; text = "セル" }
        @{ name = "行番号"; text = "12" }
        @{ name = "行（大文字小文字を区別しない）"; text = "abc" }
    ) {
        param ($name, $text)
        $row.Contains($text) | Should Be $true
    }

    It "どれにも無ければ $false" {
        $row.Contains("請求") | Should Be $false
    }
}

Describe "HitRow の静的な関数" -Tag Unit {
    It "ToDisplay はタブを区切り線に、セル内改行を記号にする" {
        [HitRow]::ToDisplay("a`tb$([char]0x2028)c") | Should Be "a │ b$([char]0x21b5)c"
        [HitRow]::ToDisplay($null) | Should Be ""
    }

    It "HasMatch は語句を大文字小文字を区別せずに探す" {
        [HitRow]::HasMatch("ABC", "b", $null) | Should Be $true
        [HitRow]::HasMatch("ABC", "x", $null) | Should Be $false
        [HitRow]::HasMatch("", "b", $null) | Should Be $false
        [HitRow]::HasMatch("ABC", "", $null) | Should Be $false
    }

    It "HasMatch は正規表現の長さ 0 の一致を数えない" {
        [HitRow]::HasMatch("ABC", "", [regex]"x*") | Should Be $false
        [HitRow]::HasMatch("ABC", "", [regex]"x*B") | Should Be $true
    }

    It "FindMatches は語句の位置と長さを返す" {
        $found = [HitRow]::FindMatches("見積と見積", "見積", $null)
        $found.Count | Should Be 2
        ($found[0] -join ",") | Should Be "0,2"
        ($found[1] -join ",") | Should Be "3,2"
    }

    It "FindMatches は空の文字・空の語句で何も返さない" {
        [HitRow]::FindMatches("", "見積", $null).Count | Should Be 0
        [HitRow]::FindMatches("見積", "", $null).Count | Should Be 0
    }

    It "SplitCells は Excel 以外はタブで分けるだけ" {
        ([HitRow]::SplitCells("a`t`"b`tc`"", $false) -join "|") | Should Be "a|`"b|c`""
        [HitRow]::SplitCells($null, $false).Count | Should Be 1
    }

    It "SplitCells は Excel の引用符で囲んだセルを 1 つにし、囲みを外す" {
        ([HitRow]::SplitCells("a`t`"b`tc`"`t`"x`"`"y`"`td", $true) -join "|") | Should Be "a|b`tc|x`"y|d"
    }

    It "SplitCells は Excel の空のセルも数える" {
        ([HitRow]::SplitCells("a`t`tb", $true) -join "|") | Should Be "a||b"
    }

    It "ColumnName は列番号を Excel の列名にする" {
        [HitRow]::ColumnName(1) | Should Be "A"
        [HitRow]::ColumnName(13) | Should Be "M"
        [HitRow]::ColumnName(27) | Should Be "AA"
        [HitRow]::ColumnName(703) | Should Be "AAA"
        [HitRow]::ColumnName(0) | Should Be ""
    }

    # [int](n / 26) は四捨五入になるため、切り捨てでないと 26 列目が "AZ"、15 列目が "AO" になっていた
    It "ColumnName は 26 で割った余りが大きい列も正しく出す" {
        [HitRow]::ColumnName(15) | Should Be "O"
        [HitRow]::ColumnName(26) | Should Be "Z"
        [HitRow]::ColumnName(702) | Should Be "ZZ"
        [HitRow]::ColumnName(16384) | Should Be "XFD"
    }

    It "TextWidth は半角を狭く、全角を広く数える" {
        [HitRow]::TextWidth("") | Should Be 14
        [HitRow]::TextWidth("ab") | Should Be 28
        [HitRow]::TextWidth("あ") | Should Be 26
        [HitRow]::TextWidth("ｱ") | Should Be 21
        [HitRow]::TextWidth($null) | Should Be 14
    }

    It "CellWidth はセル内の最も長い行の幅" {
        [HitRow]::CellWidth("ab$([char]0x2028)abcd`nabc") | Should Be ([HitRow]::TextWidth("abcd"))
        [HitRow]::CellWidth($null) | Should Be 14
    }

    It "CellWidth は 1 行を先頭の決まった文字数までで測る" {
        [HitRow]::CellWidth("a" * 500) | Should Be ([HitRow]::TextWidth("a" * [HitRow]::MaxWidthChars))
    }

    It "CellWidth は段落の幅の上限に届いたら残りの行を測らない" {
        $wide = "あ" * [HitRow]::MaxWidthChars
        [HitRow]::CellWidth("$wide`n$wide") | Should Be ([HitRow]::TextWidth($wide))
    }

    It "Shorten は長いときだけ切り詰めて … を付ける" {
        [HitRow]::Shorten("abcdef", 3) | Should Be "abc…"
        [HitRow]::Shorten("abc", 3) | Should Be "abc"
        [HitRow]::Shorten($null, 3) | Should BeNullOrEmpty
    }
}

Describe "HitRow.BuildPreview" -Tag Unit {
    It "Excel は列名を見出しにし、選択行で一致したセルの位置を覚える" {
        $row = newHitRow -lineNumber 5 -line "品名`t見積" -word "見積"
        $table = $row.BuildPreview(@(4, 5, 6), @("品名`t単価", "品名`t見積", "a"))
        @($table.Columns | ForEach-Object { $_.Label }) | Should Be @("A", "B")
        $table.TotalColumns | Should Be 2
        $table.ShownColumns | Should Be 2
        $table.RangeLabel | Should Be "A〜B"
        $table.Rows.Count | Should Be 3
        @($table.Rows | ForEach-Object { $_.Number }) | Should Be @("4", "5", "6")
        @($table.Rows | ForEach-Object { $_.IsHitRow }) | Should Be @($false, $true, $false)
        $table.Rows[1].Cells[1].IsHit | Should Be $true
        $table.Rows[1].Cells[0].IsHit | Should Be $false
        $table.HitOffset | Should Be ([HitRow]::NumberWidth + $table.Columns[0].Width)
        $table.HitWidth | Should Be $table.Columns[1].Width
        # 短い行の足りないセルは空
        $table.Rows[2].Cells[1].Text | Should Be ""
        $table.Rows[2].Cells[1].ToolTip | Should BeNullOrEmpty
    }

    It "列の幅は下限と上限の間に収める" {
        $row = newHitRow -lineNumber 1 -line "a`t$("あ" * 80)" -word "a"
        $table = $row.BuildPreview(@(1), @("a`t$("あ" * 80)"))
        $table.Columns[0].Width | Should Be ([HitRow]::MinCellWidth)
        $table.Columns[1].Width | Should Be ([HitRow]::MaxCellWidth)
    }

    It "Word・PowerPoint は列番号を見出しにし、広い幅まで許す" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -lineNumber 1 -line ("あ" * 80) -word "あ"
        $table = $row.BuildPreview(@(1), @("あ" * 80))
        $table.Columns[0].Label | Should Be "1"
        $table.RangeLabel | Should Be "1〜1"
        $table.Columns[0].Width | Should BeGreaterThan ([HitRow]::MaxCellWidth)
    }

    It "Excel の図形の行は、見出しを「セル」「文字」にする" {
        $row = newHitRow -location "[シート] 4月[図形]" -lineNumber 1 -line "B3`t納期`tx" -word "納期"
        $table = $row.BuildPreview(@(1), @("B3`t納期`tx"))
        @($table.Columns | ForEach-Object { $_.Label }) | Should Be @("セル", "文字", "3")
        $table.RangeLabel | Should Be "セル〜3"
    }

    It "前後の行が無い・数が合わないときは、その行だけで作る" {
        $row = newHitRow -lineNumber 9 -line "品名`t見積" -word "見積"
        foreach ($args2 in @(@($null, $null), @(@(), @()), @(@(1, 2), @("a")))) {
            $table = $row.BuildPreview($args2[0], $args2[1])
            $table.Rows.Count | Should Be 1
            $table.Rows[0].Number | Should Be "9"
            $table.Rows[0].IsHitRow | Should Be $true
        }
    }

    It "一致したセルが無ければ左端を表示位置にする" {
        $row = newHitRow -lineNumber 1 -line "a`tb" -word "見積"
        $table = $row.BuildPreview(@(1), @("a`tb"))
        $table.HitOffset | Should Be 0
        $table.HitWidth | Should Be 0
    }

    It "セル内改行は改行にし、長いセルは表示と説明を切り詰める" {
        $long = "え" * ([HitRow]::MaxToolTipChars + 5)
        $row = newHitRow -lineNumber 1 -line "a$([char]0x2028)b`t$long" -word "a"
        $table = $row.BuildPreview(@(1), @("a$([char]0x2028)b`t$long"))
        $table.Rows[0].Cells[0].Text | Should Be "a`nb"
        $table.Rows[0].Cells[0].ToolTip | Should Be "a`nb"
        $table.Rows[0].Cells[1].Display.Length | Should Be ([HitRow]::MaxCellChars + 1)
        $table.Rows[0].Cells[1].ToolTip.Length | Should Be ([HitRow]::MaxToolTipChars + 1)
        $table.Rows[0].Cells[1].Column | Should Be $table.Columns[1]
    }

    It "列が多すぎるときは、一致した列を中心に上限の数だけ出す" {
        $max = [HitRow]::MaxPreviewColumns
        $cells = @(1..($max + 300) | ForEach-Object { "c$_" })
        $cells[249] = "見積"
        $line = $cells -join "`t"
        $row = newHitRow -lineNumber 1 -line $line -word "見積"
        $table = $row.BuildPreview(@(1), @($line))
        $table.TotalColumns | Should Be ($max + 300)
        $table.ShownColumns | Should Be $max
        $table.Columns.Count | Should Be $max
        # 250 列目が中央に来るよう 150 列目から出す
        $table.Rows[0].Cells[0].Text | Should Be "c150"
        $table.Rows[0].Cells[$max - 1].Text | Should Be "c$(149 + $max)"
        $table.Rows[0].Cells[100].Text | Should Be "見積"
        $table.Rows[0].Cells[100].IsHit | Should Be $true
    }

    It "列が多すぎて一致した列が右端に近いときは、右端までを出す" {
        $max = [HitRow]::MaxPreviewColumns
        $cells = @(1..($max + 10) | ForEach-Object { "c$_" })
        $cells[$max + 9] = "見積"
        $line = $cells -join "`t"
        $row = newHitRow -lineNumber 1 -line $line -word "見積"
        $table = $row.BuildPreview(@(1), @($line))
        $table.Rows[0].Cells[0].Text | Should Be "c11"
        $table.Rows[0].Cells[$max - 1].Text | Should Be "見積"
    }

    It "列が多すぎて一致した列が無いときは、左端から出す" {
        $max = [HitRow]::MaxPreviewColumns
        $line = @(1..($max + 10) | ForEach-Object { "c$_" }) -join "`t"
        $row = newHitRow -lineNumber 1 -line $line -word "見積"
        $table = $row.BuildPreview(@(1), @($line))
        $table.Columns[0].Label | Should Be "A"
        $table.Rows[0].Cells[0].Text | Should Be "c1"
    }

    It "空の行も 1 列として出す" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -lineNumber 1 -line "" -word "x"
        $table = $row.BuildPreview(@(1), [string[]]@($null))
        $table.ShownColumns | Should Be 1
        $table.RangeLabel | Should Be "1〜1"
        $table.Rows[0].Cells[0].Text | Should Be ""
    }
}

Describe "FolderItem" -Tag Unit {
    It "有効・名前・パスは変わったときだけ通知する" {
        $item = [FolderItem]::new()
        $names = watchChanges $item
        $item.SetEnabled($true); $item.SetEnabled($true)
        $item.SetName("営業"); $item.SetName("営業")
        $item.SetPath("C:\共有\営業部"); $item.SetPath("C:\共有\営業部")
        $item.Enabled | Should Be $true
        $item.Name | Should Be "営業"
        $item.Path | Should Be "C:\共有\営業部"
        @($names) | Should Be @("Enabled", "Name", "Path")
    }

    It "状態・件数をまとめて変えて通知する" {
        $item = [FolderItem]::new()
        $names = watchChanges $item
        $item.SetStatus("作成済み", "Green")
        $item.SetStats("12 件", "Excel 10 件", "2026/09/01")
        $item.StatusText | Should Be "作成済み"
        $item.StatusBrush | Should Be "Green"
        $item.FileCountText | Should Be "12 件"
        $item.FileCountToolTip | Should Be "Excel 10 件"
        $item.LastIngestedText | Should Be "2026/09/01"
        @($names) | Should Be @("StatusText", "StatusBrush", "FileCountText", "FileCountToolTip", "LastIngestedText")
    }
}

Describe "IndexNode（静的な関数）" -Tag Unit {
    It "LongPath は長いパスの形にする" {
        [IndexNode]::LongPath("C:\index") | Should Be "\\?\C:\index"
        [IndexNode]::LongPath("C:") | Should Be "\\?\C:\"
        [IndexNode]::LongPath("\\server\share") | Should Be "\\?\UNC\server\share"
        [IndexNode]::LongPath("\\?\C:\index") | Should Be "\\?\C:\index"
    }

    It "IsBookDir は Office ファイルのインデックスのフォルダ名を見分ける" -TestCases @(
        @{ name = "見積.xlsx"; expected = $true }
        @{ name = "見積.XLS"; expected = $true }
        @{ name = "議事録.docx"; expected = $true }
        @{ name = "資料.pptm"; expected = $true }
        @{ name = "営業部"; expected = $false }
        @{ name = "memo.txt"; expected = $false }
        @{ name = "a.xlsxx"; expected = $false }
        @{ name = "a.xl"; expected = $false }
    ) {
        param ($name, $expected)
        [IndexNode]::IsBookDir($name) | Should Be $expected
    }

    It "IsBookDirPath は、名前が .xlsx などで終わる本物のフォルダ（まとめファイル・サブフォルダがある）を見分ける" {
        $dir = "$TestDrive\bookdir_path"
        newTsv "$dir\資料.xlsx\本文.docx.tsv" @("x")
        [void][System.IO.Directory]::CreateDirectory("$dir\親.xlsx\子")
        [void][System.IO.Directory]::CreateDirectory("$dir\空.xlsx")
        newTsv "$dir\B.xlsx\S.tsv" @("x")
        [IndexNode]::IsBookDirPath("$dir\資料.xlsx") | Should Be $false
        [IndexNode]::IsBookDirPath("$dir\親.xlsx") | Should Be $false
        [IndexNode]::IsBookDirPath("$dir\空.xlsx") | Should Be $true
        [IndexNode]::IsBookDirPath("$dir\B.xlsx") | Should Be $true
        [IndexNode]::IsBookDirPath("$dir\営業部") | Should Be $false
        # 本物のフォルダはツリーに出し、元のファイルごとのフォルダは出さない
        [IndexNode]::HasSubfolders($dir) | Should Be $true
    }
}

Describe "IndexNode（チェック）" -Tag Unit {
    # 根 ─ 営業部 ─ 東京・大阪、総務部（フォルダは読まない。子を手で足す）
    function newTree {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $sales = [IndexNode]::new($root, "営業部", $root.Root, "営業部", $false)
        $tokyo = [IndexNode]::new($sales, "東京", $root.Root, "営業部\東京", $false)
        $osaka = [IndexNode]::new($sales, "大阪", $root.Root, "営業部\大阪", $false)
        $general = [IndexNode]::new($root, "総務部", $root.Root, "総務部", $false)
        $sales.Children.Add($tokyo); $sales.Children.Add($osaka); $sales.Children.Add([IndexNode]::NewPlaceholder($sales))
        $root.Children.Add($sales); $root.Children.Add($general)
        @{ Root = $root; Sales = $sales; Tokyo = $tokyo; Osaka = $osaka; General = $general }
    }

    It "作ったときは親のチェックを受け継ぐ" {
        $parent = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        [IndexNode]::new($parent, "a", "C:\index\営業", "a", $false).IsChecked | Should Be $true
        $parent.IsChecked = $false
        [IndexNode]::new($parent, "b", "C:\index\営業", "b", $false).IsChecked | Should Be $false
        $parent.IsChecked = $null
        [IndexNode]::new($parent, "c", "C:\index\営業", "c", $false).IsChecked | Should Be $true
    }

    It "外すと子も外れ、親は一部だけの状態（null）になる" {
        $t = newTree
        $t.Sales.SetChecked($false)
        $t.Tokyo.IsChecked | Should Be $false
        $t.Osaka.IsChecked | Should Be $false
        $t.Root.IsChecked | Should Be $null
        $t.General.IsChecked | Should Be $true
    }

    It "子の 1 つを外すと、親と根は一部だけになる" {
        $t = newTree
        $t.Tokyo.Toggle()
        $t.Tokyo.IsChecked | Should Be $false
        $t.Sales.IsChecked | Should Be $null
        $t.Root.IsChecked | Should Be $null
    }

    It "子をすべて外すと親も外れる" {
        $t = newTree
        $t.Tokyo.SetChecked($false)
        $t.Osaka.SetChecked($false)
        $t.Sales.IsChecked | Should Be $false
        $t.General.SetChecked($false)
        $t.Root.IsChecked | Should Be $false
    }

    It "一部だけの状態で切り替えるとすべて付く" {
        $t = newTree
        $t.Tokyo.SetChecked($false)
        $t.Sales.Toggle()
        $t.Sales.IsChecked | Should Be $true
        $t.Tokyo.IsChecked | Should Be $true
        $t.Root.IsChecked | Should Be $true
    }

    It "チェックの変化を通知し、変わらなければ通知しない" {
        $t = newTree
        $names = watchChanges $t.Tokyo
        $t.Tokyo.SetChecked($true)
        $names.Count | Should Be 0
        $t.Tokyo.SetChecked($false)
        @($names) | Should Be @("IsChecked")
    }

    It "検索対象は、すべて付いた枝だけを 1 件にまとめる" {
        $t = newTree
        $t.Tokyo.SetChecked($false)
        $targets = [System.Collections.Generic.List[SearchTarget]]::new()
        $t.Root.AddTargets($targets)
        @($targets | ForEach-Object { "$($_.RelPath)|$($_.Recurse)" }) | Should Be @("営業部\大阪|True", "総務部|True")
    }

    It "すべて付いていれば根だけを対象にする" {
        $t = newTree
        $targets = [System.Collections.Generic.List[SearchTarget]]::new()
        $t.Root.AddTargets($targets)
        $targets.Count | Should Be 1
        $targets[0].Root | Should Be "C:\index\営業"
        $targets[0].RelPath | Should Be ""
    }

    It "直下のファイルの項目は、サブフォルダを含めない対象にする" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $files = [IndexNode]::new($root, "（このフォルダ直下のファイル）", $root.Root, "", $true)
        $sub = [IndexNode]::new($root, "営業部", $root.Root, "営業部", $false)
        $root.Children.Add($files); $root.Children.Add($sub)
        $sub.SetChecked($false)
        $targets = [System.Collections.Generic.List[SearchTarget]]::new()
        $root.AddTargets($targets)
        $targets.Count | Should Be 1
        $targets[0].Recurse | Should Be $false
    }

    It "外したものは、外した枝ごとに記録する" {
        $t = newTree
        $t.Tokyo.SetChecked($false)
        $t.General.SetChecked($false)
        $excludes = [System.Collections.Generic.List[SearchExclude]]::new()
        $t.Root.AddExcludes($excludes)
        @($excludes | ForEach-Object { "$($_.Path)|$($_.Subfolders)" }) | Should Be @("C:\index\営業\営業部\東京|True", "C:\index\営業\総務部|True")
    }

    It "すべて付いていれば外したものは無い" {
        $t = newTree
        $excludes = [System.Collections.Generic.List[SearchExclude]]::new()
        $t.Root.AddExcludes($excludes)
        $excludes.Count | Should Be 0
    }

    It "展開しているフォルダのパスを集める" {
        $t = newTree
        $t.Root.IsExpanded = $true
        $t.Sales.IsExpanded = $true
        $paths = [System.Collections.Generic.List[string]]::new()
        $t.Root.AddExpanded($paths)
        @($paths) | Should Be @("C:\index\営業", "C:\index\営業\営業部")
    }

    It "FullPath は根ならそのまま、下なら相対パスをつなぐ" {
        $t = newTree
        $t.Root.FullPath() | Should Be "C:\index\営業"
        $t.Tokyo.FullPath() | Should Be "C:\index\営業\営業部\東京"
    }
}

Describe "IndexNode（フォルダの読み込み）" -Tag Io {
    # インデックスのフォルダ:
    #   営業\直下.xlsx\Sheet1.tsv          （根の直下のファイル）
    #   営業\営業部\見積.xlsx\4月.tsv
    #   営業\営業部\東京\
    #   営業\総務部\
    BeforeEach {
        $script:indexRoot = "$TestDrive\index\営業"
        Remove-Item -LiteralPath "$TestDrive\index" -Recurse -Force -ErrorAction SilentlyContinue
        newTsv "$script:indexRoot\直下.xlsx\Sheet1.tsv" @("1`ta")
        newTsv "$script:indexRoot\営業部\見積.xlsx\4月.tsv" @("1`ta")
        [System.IO.Directory]::CreateDirectory("$script:indexRoot\営業部\東京") | Out-Null
        [System.IO.Directory]::CreateDirectory("$script:indexRoot\総務部") | Out-Null
    }

    It "根を作ると、サブフォルダがあれば読み込み中の子を置く" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", "C:\共有\営業")
        $root.Exists | Should Be $true
        $root.SourcePath | Should Be "C:\共有\営業"
        $root.ToolTip | Should Be "元のフォルダ：C:\共有\営業`nインデックス：$script:indexRoot"
        $root.Children.Count | Should Be 1
        $root.Children[0].IsPlaceholder | Should Be $true
    }

    # 引数 [string]$sourcePath は $null を "" にするため、$null だけを見ると「元のフォルダ：」が空のまま出ていた
    It "元のフォルダが分からなければ、ツールチップはインデックスの場所" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.ToolTip | Should Be $script:indexRoot
    }

    It "フォルダが無ければ、その旨をツールチップに出し、子を読まない" {
        $root = [IndexNode]::CreateRoot("$TestDrive\index\無い", "無い", "", $null)
        $root.Exists | Should Be $false
        $root.ToolTip | Should Match "フォルダが見つかりません"
        $root.Children.Count | Should Be 0
        $root.LoadChildren()
        $root.Children.Count | Should Be 0
    }

    It "展開すると、直下のファイルの項目とサブフォルダを名前順に読む（Office のフォルダは除く）" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", "C:\共有\営業")
        $names = watchChanges $root
        $root.SetExpanded($true)
        $root.IsExpanded | Should Be $true
        @($names) | Should Be @("IsExpanded")
        @($root.Children | ForEach-Object { $_.Name }) | Should Be @("（このフォルダ直下のファイル）", "営業部", "総務部")
        $root.Children[0].IsFiles | Should Be $true
        $root.Children[0].ToolTip | Should Be "サブフォルダを除く、C:\共有\営業 の直下のファイル"
        $sales = $root.Children[1]
        $sales.RelPath | Should Be "営業部"
        $sales.SourcePath | Should Be "C:\共有\営業\営業部"
        $sales.ToolTip | Should Be "元のフォルダ：C:\共有\営業\営業部"
        # 営業部には東京があるので読み込み中の子を置く。総務部には置かない
        $sales.Children.Count | Should Be 1
        $sales.Children[0].IsPlaceholder | Should Be $true
        $root.Children[2].Children.Count | Should Be 0
    }

    It "2 回目は読み直さない" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.LoadChildren()
        [System.IO.Directory]::CreateDirectory("$script:indexRoot\経理部") | Out-Null
        $root.SetExpanded($false)
        $root.SetExpanded($true)
        @($root.Children | ForEach-Object { $_.Name }) -contains "経理部" | Should Be $false
    }

    It "元のフォルダが分からなければ、子のツールチップはインデックスの場所" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.LoadChildren()
        $root.Children[1].SourcePath | Should BeNullOrEmpty
        $root.Children[0].ToolTip | Should Be "サブフォルダを除く、$script:indexRoot の直下のファイル"
        $root.Children[1].ToolTip | Should Be "$script:indexRoot\営業部"
    }

    It "サブフォルダがあっても直下にファイルが無ければ、直下のファイルの項目を作らない" {
        Remove-Item -LiteralPath "$script:indexRoot\直下.xlsx" -Recurse -Force
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.LoadChildren()
        @($root.Children | ForEach-Object { $_.Name }) | Should Be @("営業部", "総務部")
    }

    It "直下のまとめファイルもファイルとして数える（ほかの .tsv は数えない）" {
        newTsv "$script:indexRoot\人事部\a.tsv" @("x")
        [IndexNode]::HasFiles("$script:indexRoot\人事部") | Should Be $false
        newTsv "$script:indexRoot\総務部\本文.xlsx.tsv" @("x")
        [IndexNode]::HasFiles("$script:indexRoot\総務部") | Should Be $true
        [IndexNode]::HasFiles("$script:indexRoot\営業部\東京") | Should Be $false
        [IndexNode]::HasFiles("$TestDrive\無い") | Should Be $false
        [IndexNode]::HasSubfolders("$TestDrive\無い") | Should Be $false
    }

    It "直下のファイルの項目・読み込み中の子は読み込まない" {
        $files = [IndexNode]::new($null, "直下", $script:indexRoot, "", $true)
        $files.LoadChildren()
        $files.Children.Count | Should Be 0
    }

    It "Find はパスでたどり、必要なところだけ読み込む" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $found = $root.Find("$script:indexRoot\営業部\東京\")
        $found.RelPath | Should Be "営業部\東京"
        $root.Find($script:indexRoot) | Should Be $root
        $root.Find("$script:indexRoot\無い") | Should Be $null
        $root.Find("$TestDrive\ほか") | Should Be $null
    }

    It "ApplyExclude でサブフォルダごと外す" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.ApplyExclude("$script:indexRoot\営業部", $true)
        $root.Find("$script:indexRoot\営業部").IsChecked | Should Be $false
        $root.IsChecked | Should Be $null
    }

    It "ApplyExclude でサブフォルダを除く（直下のファイルだけ外す）" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.ApplyExclude($script:indexRoot, $false)
        $root.Children[0].IsFiles | Should Be $true
        $root.Children[0].IsChecked | Should Be $false
        $root.Children[1].IsChecked | Should Be $true
        $root.IsChecked | Should Be $null
    }

    It "ApplyExclude でサブフォルダを除くとき、子が無いフォルダはそのものを外す" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.ApplyExclude("$script:indexRoot\総務部", $false)
        $root.Find("$script:indexRoot\総務部").IsChecked | Should Be $false
    }

    It "ApplyExclude でサブフォルダを除くとき、直下のファイルの項目が無くサブフォルダがあれば何もしない" {
        [System.IO.Directory]::CreateDirectory("$script:indexRoot\総務部\人事") | Out-Null
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.ApplyExclude("$script:indexRoot\総務部", $false)
        $root.Find("$script:indexRoot\総務部").IsChecked | Should Be $true
        $root.IsChecked | Should Be $true
    }

    It "ApplyExclude でサブフォルダを除くとき、直下のファイルの項目があればそれだけ外す" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.ApplyExclude("$script:indexRoot\営業部", $false)
        $sales = $root.Find("$script:indexRoot\営業部")
        $sales.Children[0].IsFiles | Should Be $true
        $sales.Children[0].IsChecked | Should Be $false
        $sales.IsChecked | Should Be $null
    }

    It "ApplyExclude で見つからないパスは無視する" {
        $root = [IndexNode]::CreateRoot($script:indexRoot, "営業", "", $null)
        $root.ApplyExclude("$TestDrive\ほか", $true)
        $root.IsChecked | Should Be $true
    }
}

Describe "HitRow（境界値）" -Tag Unit {
    It "空の行は、表示も強調も空" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "" -word "見積"
        $row.Prepare()
        $row.DisplayLine | Should Be ""
        $row.Segments.Count | Should Be 0
    }

    It "一致しない行は、行全体を強調なしの 1 つにする" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "請求書" -word "見積"
        $row.Prepare()
        $row.Segments.Count | Should Be 1
        $row.Segments[0].IsHit | Should Be $false
    }

    It "語句の検索では、正規表現の記号をそのままの文字として探す" {
        [HitRow]::HasMatch("125", "1.5", $null) | Should Be $false
        [HitRow]::HasMatch("1.5 倍", "1.5", $null) | Should Be $true
        [HitRow]::HasMatch("[シート] 4月", "[シート]", $null) | Should Be $true
    }

    It "語句の検索では、重なる一致を二重に数えない" {
        $found = [HitRow]::FindMatches("あああ", "ああ", $null)
        $found.Count | Should Be 1
        ($found[0] -join ",") | Should Be "0,2"
    }

    It "大文字小文字を区別する正規表現は、そのとおりに探す" {
        [HitRow]::HasMatch("abc", "", [regex]"ABC") | Should Be $false
        [HitRow]::HasMatch("abc", "", [regex]::new("ABC", "IgnoreCase")) | Should Be $true
    }

    It "Excel の引用符で囲んだセルの後ろに続く文字も同じセルにする" {
        ([HitRow]::SplitCells("`"a`"b`tc", $true) -join "|") | Should Be "ab|c"
    }

    It "Excel の空の行は空のセル 1 つ" {
        [HitRow]::SplitCells("", $true).Count | Should Be 1
    }

    It "空の文字で絞り込むとすべて合う" {
        (newHitRow -line "見積").Contains("") | Should Be $true
    }

    It "正規表現で Excel のセル番地を決める" {
        $row = newHitRow -lineNumber 4 -line "品名`t第1回`t第2回" -pattern ([regex]"第\d回")
        $row.Prepare()
        $row.CellText | Should Be "B4 ほか 1"
    }

    It "27 列目以降の一致は 2 文字の列名で出す" {
        $cells = @(1..30 | ForEach-Object { "c$_" })
        $cells[26] = "見積"
        $row = newHitRow -lineNumber 8 -line ($cells -join "`t") -word "見積"
        $row.Prepare()
        $row.MatchCell | Should Be "AA8"
    }

    It "選択行が前後の行に無ければ、強調する行を作らず左端を表示位置にする" {
        $row = newHitRow -lineNumber 9 -line "見積" -word "見積"
        $table = $row.BuildPreview(@(1, 2), @("見積", "x"))
        @($table.Rows | Where-Object { $_.IsHitRow }).Count | Should Be 0
        $table.HitOffset | Should Be 0
        # 一致したセルの印は付ける
        $table.Rows[0].Cells[0].IsHit | Should Be $true
    }

    It "引用符だけのセルは Excel の形で囲む" {
        [PreviewTable]::QuoteForExcel('"') | Should Be '""""'
        [PreviewTable]::QuoteForExcel("") | Should Be ""
    }
}

Describe "IndexNode（境界値）" -Tag Unit {
    It "読み込み中の子を変えても親は変わらない" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $placeholder = [IndexNode]::NewPlaceholder($root)
        $root.Children.Add($placeholder)
        $placeholder.SetChecked($false)
        $root.IsChecked | Should Be $true
    }

    It "読み込み中の子は検索対象・除外・展開に入れない" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $placeholder = [IndexNode]::NewPlaceholder($root)
        $placeholder.IsChecked = $null
        $placeholder.IsExpanded = $true
        $targets = [System.Collections.Generic.List[SearchTarget]]::new()
        $excludes = [System.Collections.Generic.List[SearchExclude]]::new()
        $paths = [System.Collections.Generic.List[string]]::new()
        $placeholder.AddTargets($targets)
        $placeholder.AddExcludes($excludes)
        $placeholder.AddExpanded($paths)
        $targets.Count | Should Be 0
        $excludes.Count | Should Be 0
        $paths.Count | Should Be 0
        $placeholder.Find("C:\index\営業") | Should Be $null
    }

    It "直下のファイルの項目は、外すとサブフォルダを含めない除外にし、展開には入れない" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $files = [IndexNode]::new($root, "（このフォルダ直下のファイル）", $root.Root, "", $true)
        $root.Children.Add($files)
        $files.IsExpanded = $true
        $files.SetChecked($false)
        $excludes = [System.Collections.Generic.List[SearchExclude]]::new()
        $files.AddExcludes($excludes)
        $excludes[0].Path | Should Be "C:\index\営業"
        $excludes[0].Subfolders | Should Be $false
        $paths = [System.Collections.Generic.List[string]]::new()
        $files.AddExpanded($paths)
        $paths.Count | Should Be 0
        $files.Find("C:\index\営業") | Should Be $null
    }

    It "Find はパスの大文字小文字・末尾の \ を区別しない" {
        $root = [IndexNode]::new($null, "営業", "C:\Index\Sales", "", $false)
        $root.Find("c:\index\sales\") | Should Be $root
    }

    It "Find は名前の先頭が同じだけの別のフォルダを選ばない" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $root.Find("C:\index\営業部") | Should Be $null
    }

    It "根の末尾に \ があっても、子のパスは \ を重ねない" {
        [IndexNode]::new($null, "a", "C:\index\", "営業部", $false).FullPath() | Should Be "C:\index\営業部"
    }

    It "閉じるときは通知し、閉じたままなら通知しない" {
        $node = [IndexNode]::new($null, "営業", "C:\index\無い", "", $true)
        $names = watchChanges $node
        $node.SetExpanded($false)
        $names.Count | Should Be 0
        $node.SetExpanded($true)
        $node.SetExpanded($false)
        @($names) | Should Be @("IsExpanded", "IsExpanded")
    }
}