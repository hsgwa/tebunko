# 検索結果とプレビューの型（tebunko\ui\search\result_types.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_types.ps1"

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
            [string]$location = "[シート]4月",
            [int]$lineNumber = 3,
            [string]$line = "",
            [string]$word = "",
            [regex]$pattern = $null,
            [string]$relDir = "営業部"
        )
        $row = [HitRow]::Create("営業", "C:\index\営業", "$relDir\$book", $relDir, "$book.tsv", $book, $location, $lineNumber, $line, $word, $pattern)
        # 行を作る側（result_list.ps1）と同じ式で IsExcel・IsCellPrefixed を入れる
        $row.IsExcel = ((getContentIndexFileKind $book) -eq "Excel")
        $row.IsCellPrefixed = testCellPrefixedPlace $book (splitObjectPlace $location).Kind
        $row
    }

    # 表の全部の列・全部の行のセルを作る（画面は見えている範囲だけ作る。テストでは全部を見えている範囲にして確かめる）
    function showAll($table) {
        [void]$table.SetWindow(0, $table.Columns.Count - 1, 0, 0)
        $table.EnsureCells(0, $table.Rows.Count - 1)
        return $table
    }

    # セグメントを「文字」と「強調の有無」の組にして比べやすくする
    function describeSegments($segments) {
        @($segments | ForEach-Object { if ($_.IsHit) { "[$($_.Text)]" } else { $_.Text } }) -join ""
    }
}

Describe "PreviewColumn" -Tag Unit {
    It "幅を変えると通知する" {
        $column = [PreviewColumn]::new()
        $names = watchChanges $column
        $column.SetWidth(100)
        $column.Width | Should -Be 100
        @($names) | Should -Be @("Width")
    }

    It "下限より狭くしない" {
        $column = [PreviewColumn]::new()
        $column.SetWidth(3)
        $column.Width | Should -Be ([PreviewColumn]::MinWidth)
    }

    It "同じ幅なら通知しない" {
        $column = [PreviewColumn]::new()
        $column.SetWidth(100)
        $names = watchChanges $column
        $column.SetWidth(100)
        $names.Count | Should -Be 0
    }
}

Describe "PreviewCell" -Tag Unit {
    It "選択は変わったときだけ通知する" {
        $cell = [PreviewCell]::new()
        $names = watchChanges $cell
        $cell.SetSelected($true)
        $cell.SetSelected($true)
        $cell.IsSelected | Should -Be $true
        @($names) | Should -Be @("IsSelected")
    }
}

Describe "PreviewTable（範囲の選択とコピー）" -Tag Unit {
    BeforeAll {
        # 3 行 × 3 列の表（A1〜C3 の値を持つ）
        function newTable {
            $row = newHitRow -lineNumber 2 -line "A2`tB2`tC2" -word "B2"
            showAll ($row.BuildPreview(@(1, 2, 3), @("A1`tB1`tC1", "A2`tB2`tC2", "A3`tB3`tC3")))
        }
    }

    It "何も選んでいなければ空" {
        $table = newTable
        $table.HasSelection() | Should -Be $false
        $table.GetSelectionText() | Should -Be ""
        $table.SelectedCount() | Should -Be 0
    }

    It "1 つのセルを選ぶとその値をそのまま返す" {
        $table = newTable
        $table.Select($table.Rows[1].Cells[1], $false)
        $table.HasSelection() | Should -Be $true
        $table.GetSelectionText() | Should -Be "B2"
        $table.SelectedCount() | Should -Be 1
        $table.Rows[1].Cells[1].IsSelected | Should -Be $true
        $table.Rows[0].Cells[0].IsSelected | Should -Be $false
    }

    It "広げて選ぶと、行はCRLF・列はタブでつなぐ" {
        $table = newTable
        $table.Select($table.Rows[2].Cells[2], $false)
        $table.Select($table.Rows[1].Cells[1], $true)
        $table.GetSelectionText() | Should -Be "B2`tC2`r`nB3`tC3"
        $table.SelectedCount() | Should -Be 4
        @($table.Rows | ForEach-Object { $_.Cells } | Where-Object { $_.IsSelected }).Count | Should -Be 4
    }

    It "選んでいないときに広げると、そのセルから始める" {
        $table = newTable
        $table.Select($table.Rows[0].Cells[0], $true)
        $table.GetSelectionText() | Should -Be "A1"
    }

    It "広げずに選び直すと、前の選択を外す" {
        $table = newTable
        $table.Select($table.Rows[0].Cells[0], $false)
        $table.Select($table.Rows[2].Cells[2], $false)
        $table.Rows[0].Cells[0].IsSelected | Should -Be $false
        $table.GetSelectionText() | Should -Be "C3"
    }

    It "行を選ぶと、その行のすべての列" {
        $table = newTable
        $table.SelectRow($table.Rows[0].Cells[1])
        $table.GetSelectionText() | Should -Be "A1`tB1`tC1"
        $table.SelectedCount() | Should -Be 3
    }

    It "選択を外す" {
        $table = newTable
        $table.SelectRow($table.Rows[0].Cells[1])
        $table.ClearSelection()
        $table.HasSelection() | Should -Be $false
        @($table.Rows | ForEach-Object { $_.Cells } | Where-Object { $_.IsSelected }).Count | Should -Be 0
    }

    It "セルが無ければ何もしない" {
        $table = newTable
        $table.Select($null, $false)
        $table.SelectRow($null)
        $table.HasSelection() | Should -Be $false
    }

    It "列の無い表では行を選べない" {
        $table = [PreviewTable]::new()
        $table.Columns = [System.Collections.Generic.List[PreviewColumn]]::new()
        $table.Rows = [System.Collections.ObjectModel.ObservableCollection[PreviewRow]]::new()
        $table.SelectRow([PreviewCell]@{ LineNumber = 1; ColumnIndex = 0 })
        $table.HasSelection() | Should -Be $false
    }

    It "表に無いセルは空として扱う" {
        $table = newTable
        $table.Select([PreviewCell]@{ LineNumber = 5; ColumnIndex = 0 }, $false)
        $table.GetSelectionText() | Should -Be ""
        $table.Select([PreviewCell]@{ LineNumber = 1; ColumnIndex = 9 }, $false)
        $table.GetSelectionText() | Should -Be ""
    }

    It "複数のセルをコピーするとき、改行・タブ・引用符を含むセルは Excel の形で囲む" {
        $row = newHitRow -lineNumber 1 -line "x" -word "x"
        $table = showAll ($row.BuildPreview(@(1), @("A`t`"B`"`"`"`tC")))
        # 1 列目は改行を含むセルにする
        $table.Rows[0].Texts[0] = "1行目`n2行目"
        $table.Select($table.Rows[0].Cells[0], $false)
        $table.Select($table.Rows[0].Cells[2], $true)
        $table.GetSelectionText() | Should -Be "`"1行目`n2行目`"`t`"B`"`"`"`tC"
    }
}

Describe "PreviewTable（行を前後に足す・捨てる・メッセージの行）" -Tag Unit {
    BeforeAll {
        # 行の番号 4〜6（選択行は 5）の 2 列の表
        function newReadTable {
            $row = newHitRow -lineNumber 5 -line "A5`tB5" -word "B5"
            $table = showAll ($row.BuildPreview(@(4, 5, 6), @("A4`tB4", "A5`tB5", "A6`tB6")))
            return @{ Row = $row; Table = $table }
        }
        function lineNumbers {
            param ($table)
            return (@($table.Rows | ForEach-Object { if ($_.IsMessage) { "[$($_.Message)]" } else { $_.LineNumber } }) -join ",")
        }
    }

    It "最初の行の番号・最後の行の番号・本物の行の数を返す。行の番号が 1 からなら、上にはもう読む行が無い" {
        $made = newReadTable
        $made.Table.FirstLine() | Should -Be 4
        $made.Table.LastLine() | Should -Be 6
        $made.Table.DataCount() | Should -Be 3
        $made.Table.TopDone | Should -Be $false
        ($made.Row.BuildPreview(@(1, 2), @("a", "b"))).TopDone | Should -Be $true
        $empty = [PreviewTable]::new()
        $empty.Rows = [System.Collections.ObjectModel.ObservableCollection[PreviewRow]]::new()
        $empty.FirstLine() | Should -Be 0
        $empty.LastLine() | Should -Be 0
    }

    It "読み足した行を、上と下に足す（メッセージの行の内側）。セルは表の列に合わせ、行の番号を持つ" {
        $made = newReadTable
        $table = $made.Table
        $table.ShowMessage("up", "読み込み中…")
        $table.ShowMessage("down", "読み込み中…")
        lineNumbers $table | Should -Be "[読み込み中…],4,5,6,[読み込み中…]"
        $table.LeadCount() | Should -Be 1
        $table.DataCount() | Should -Be 3
        $above = $made.Row.BuildPreviewRows($table, @(2, 3), @("A2`tB2", "A3`tB3"))
        $prefixBefore = $table.Prefix()
        $below = $made.Row.BuildPreviewRows($table, @(7), @("A7`tB7`tC7"))
        $table.InsertRows("up", $above)
        $table.InsertRows("down", $below)
        lineNumbers $table | Should -Be "[読み込み中…],2,3,4,5,6,7,[読み込み中…]"
        $table.FirstLine() | Should -Be 2
        $table.LastLine() | Should -Be 7
        # 表の列（2 列）より多い列がある行を足すと、列を右に足す。前からある列の幅は変えない
        $table.Columns.Count | Should -Be 3
        $table.Columns[2].Label | Should -Be "C"
        $prefixAfter = $table.Prefix()
        $prefixAfter.Length | Should -Be 4 -Because "列が増えたので、累積も作り直す"
        $prefixAfter[1] | Should -Be $prefixBefore[1]
        $prefixAfter[2] | Should -Be $prefixBefore[2]
        $prefixAfter[3] | Should -Be ($prefixAfter[2] + $table.Columns[2].Width)
        # セルは行の番号と列の番号を持つ（足した行のセルは、見えている範囲を決め直したときに作る）
        $table.Rows[6].Cells.Count | Should -Be 0
        [void]$table.SetWindow(0, 2, 0, 0)
        $table.EnsureCells(0, $table.Rows.Count - 1)
        $table.Rows[6].Cells.Count | Should -Be 3
        $table.Rows[6].Cells[1].LineNumber | Should -Be 7
        $table.Rows[6].Cells[1].Column | Should -Be $table.Columns[1]
        $table.Rows[1].Number | Should -Be "2"
        $table.Rows[1].IsHitRow | Should -Be $false
        $table.Rows[4].IsHitRow | Should -Be $true
        $table.HideMessage("up")
        $table.HideMessage("down")
        lineNumbers $table | Should -Be "2,3,4,5,6,7"
        $table.HideMessage("up")
        $table.LeadCount() | Should -Be 0
    }

    It "メッセージの行は、選択にもコピーにも入らない。メッセージを出し直すと置き換える" {
        $made = newReadTable
        $table = $made.Table
        $table.ShowMessage("down", "読み込み中…")
        $table.ShowMessage("down", "読めませんでした")
        lineNumbers $table | Should -Be "4,5,6,[読めませんでした]"
        $table.Select($table.Rows[0].Cells[0], $false)
        $table.Select($table.Rows[2].Cells[1], $true)
        $table.GetSelectionText() | Should -Be "A4`tB4`r`nA5`tB5`r`nA6`tB6"
        @($table.Rows[3].Cells).Count | Should -Be 0
        $table.SelectedCount() | Should -Be 6
    }

    It "上や下を足しても、選んだ範囲は行の番号のままで、同じセルを指す" {
        $made = newReadTable
        $table = $made.Table
        $table.Select($table.Rows[1].Cells[0], $false)
        $table.Select($table.Rows[2].Cells[1], $true)
        $table.InsertRows("up", $made.Row.BuildPreviewRows($table, @(2, 3), @("A2`tB2", "A3`tB3")))
        $table.GetSelectionText() | Should -Be "A5`tB5`r`nA6`tB6"
        $table.Rows[2].Cells[0].IsSelected | Should -Be $false
        $table.Rows[3].Cells[0].IsSelected | Should -Be $true
    }

    It "<Name>" -TestCases @(
        @{ Name = "上から 2 行捨てる"; Side = "top"; Count = 2; Expected = "6" }
        @{ Name = "下から 2 行捨てる"; Side = "bottom"; Count = 2; Expected = "4" }
        @{ Name = "行の数より多く頼んでも、ある分だけ捨てる"; Side = "top"; Count = 9; Expected = "" }
        @{ Name = "0 行なら何もしない"; Side = "top"; Count = 0; Expected = "4,5,6" }
    ) {
        param ($Name, $Side, $Count, $Expected)
        $made = newReadTable
        $made.Table.TrimRows($Side, $Count)
        lineNumbers $made.Table | Should -Be $Expected
    }

    It "捨てた行を含む範囲を選んでいたら選択を外し、含まなければ残す。メッセージの行は捨てない" {
        $made = newReadTable
        $table = $made.Table
        $table.ShowMessage("up", "読み込み中…")
        $table.Select($table.Rows[2].Cells[0], $false)
        $table.Select($table.Rows[3].Cells[1], $true)
        $table.TrimRows("bottom", 1)
        lineNumbers $table | Should -Be "[読み込み中…],4,5"
        $table.HasSelection() | Should -Be $false
        $table.Rows[2].Cells[0].IsSelected | Should -Be $false
        $kept = newReadTable
        $kept.Table.Select($kept.Table.Rows[0].Cells[0], $false)
        $kept.Table.TrimRows("bottom", 2)
        $kept.Table.HasSelection() | Should -Be $true
        $kept.Table.GetSelectionText() | Should -Be "A4"
    }
}

Describe "PreviewTable.QuoteForExcel" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "改行・タブ・引用符が無ければそのまま"; text = "見積"; expected = "見積" }
        @{ name = "空のセルはそのまま"; text = ""; expected = "" }
        @{ name = "改行を含めば引用符で囲む"; text = "a`nb"; expected = "`"a`nb`"" }
        @{ name = "タブを含めば引用符で囲む"; text = "a`tb"; expected = "`"a`tb`"" }
        @{ name = "引用符は 2 つにして囲む"; text = 'a"b'; expected = '"a""b"' }
        @{ name = "引用符だけのセルは Excel の形で囲む"; text = '"'; expected = '""""' }
    ) {
        param ($name, $text, $expected)
        [PreviewTable]::QuoteForExcel($text) | Should -Be $expected
    }
}

Describe "FileGroup" -Tag Unit {
    It "場所は初めてのときだけ $true" {
        $group = [FileGroup]::new()
        $group.AddLocation("4月") | Should -Be $true
        $group.AddLocation("4月") | Should -Be $false
        $group.AddLocation("5月") | Should -Be $true
    }

    It "同じ表記は 1 つにし、足した順に返す" {
        $group = [FileGroup]::new()
        $group.AddLabel("[シート]4月") | Should -Be $true
        $group.AddLabel("[シート]5月") | Should -Be $true
        $group.AddLabel("[シート]4月") | Should -Be $false
        @($group.GetLocations()) | Should -Be @("[シート]4月", "[シート]5月")
    }

    It "表記が無ければ空" {
        @([FileGroup]::new().GetLocations()).Count | Should -Be 0
    }

    It "開閉・件数は変わったときだけ通知する" {
        $group = [FileGroup]::new()
        $names = watchChanges $group
        $group.SetExpanded($true)
        $group.SetExpanded($true)
        $group.SetCount(3)
        $group.SetCount(3)
        $group.IsExpanded | Should -Be $true
        $group.ShownCount | Should -Be 3
        @($names) | Should -Be @("IsExpanded", "ShownCount")
    }
}

Describe "HitRow.Create" -Tag Unit {
    It "生データを入れ、Excel かどうか・図形の場所かどうかを決める" {
        $row = newHitRow -book "見積.xlsx" -location "[シート] 4月[図形]" -line "A1`t見積"
        $row.IndexName | Should -Be "営業"
        $row.Root | Should -Be "C:\index\営業"
        $row.RelPath | Should -Be "営業部\見積.xlsx"
        $row.FileName | Should -Be "見積.xlsx.tsv"
        $row.IsExcel | Should -Be $true
        $row.IsObjectPlace | Should -Be $true
        $row.Prepared | Should -Be $false
        $row.IsFileHeader | Should -Be $false
    }

    It "Excel かどうかは、行を作る側が入れる（Create は決めない）" {
        $created = [HitRow]::Create("営業", "C:\index\営業", "見積.xlsx", "営業", "見積.xlsx.tsv", "見積.xlsx", "4月", 1, "A1`t見積", "見積", $null)
        $created.IsExcel | Should -Be $false
    }

    It "コメントの場所も図形と同じく扱う" {
        (newHitRow -location "[シート]4月[コメント]").IsObjectPlace | Should -Be $true
        (newHitRow -location "[シート]4月").IsObjectPlace | Should -Be $false
    }

    It "埋め込みの場所（番号つき）は図形の場所の一種で、番号なしはふつうの場所" {
        (newHitRow -book "議事録.docx" -location "ページ001[埋め込み1]").IsObjectPlace | Should -Be $true
        (newHitRow -book "議事録.docx" -location "ページ001[埋め込み12]").IsObjectPlace | Should -Be $true
        (newHitRow -book "議事録.docx" -location "ページ001[埋め込み]").IsObjectPlace | Should -Be $false
    }

    It "図形の場所の判定の正規表現は、種類の表から組み立てたものと同じ" {
        [HitRow]::ObjectPlaceRegex.ToString() | Should -Be (newObjectPlaceKindPattern)
    }

    It "Excel のヘッダー・フッターの場所は、図形の場所の一種で、行の先頭がセル番地ではない（IsCellPrefixed は入れた値のまま。既定は偽）" {
        $row = newHitRow -location "[シート]4月[ヘッダー・フッター]"
        $row.IsObjectPlace | Should -Be $true
        $row.IsCellPrefixed | Should -Be $false
        (newHitRow -location "[シート]4月[図形]").IsCellPrefixed | Should -Be $true
        (newHitRow -location "[シート]4月[コメント]").IsCellPrefixed | Should -Be $true
        (newHitRow -location "[シート]4月").IsCellPrefixed | Should -Be $false
        $created = [HitRow]::Create("営業", "C:\index\営業", "見積.xlsx", "営業", "見積.xlsx.tsv", "見積.xlsx", "[シート]4月[図形]", 1, "A1`t見積", "見積", $null)
        $created.IsCellPrefixed | Should -Be $false
        # Word・PowerPoint の場所の名前 ヘッダー・フッター は、種類ではない
        (newHitRow -book "議事録.docx" -location "ヘッダー・フッター").IsObjectPlace | Should -Be $false
    }

    It "空の値は空文字にする" {
        $row = [HitRow]::Create("営業", "C:\index\営業", "見積.xlsx", $null, "見積.xlsx.tsv", $null, $null, 1, $null, "見積", $null)
        $row.RelDir | Should -Be ""
        $row.Line | Should -Be ""
        $row.IsObjectPlace | Should -Be $false
    }
}

Describe "HitRow.Prepare" -Tag Unit {
    It "Excel では最初に一致したセルの番地と、一致したセルの数を求める" {
        $row = newHitRow -lineNumber 7 -line "品名`t見積A`t単価`t見積B" -word "見積"
        $row.Prepare()
        $row.Prepared | Should -Be $true
        $row.MatchCell | Should -Be "B7"
        $row.MatchCount | Should -Be 2
        $row.DisplayLine | Should -Be "品名 │ 見積A │ 単価 │ 見積B"
        describeSegments $row.Segments | Should -Be "品名`t[見積]A`t単価`t[見積]B".Replace("`t", " │ ")
    }

    It "一致が 1 つなら数は 1" {
        $row = newHitRow -lineNumber 2 -line "品名`t見積" -word "見積"
        $row.Prepare()
        $row.MatchCell | Should -Be "B2"
        $row.MatchCount | Should -Be 1
    }

    It "Excel でなければセル番地を出さない" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "見積を送付" -word "見積"
        $row.Prepare()
        $row.MatchCell | Should -Be ""
        $row.MatchCount | Should -Be 0
        describeSegments $row.Segments | Should -Be "[見積]を送付"
    }

    It "Excel の図形の行は、先頭のセル番地を出し、文字だけを表示する" {
        $row = newHitRow -location "[シート] 4月[図形]" -line "B3`t`"納期は$([char]0x2028)別途`"" -word "納期"
        $row.Prepare()
        $row.MatchCell | Should -Be "B3"
        $row.MatchCount | Should -Be 1
        $row.DisplayLine | Should -Be "納期は$([char]0x21b5)別途"
        describeSegments $row.Segments | Should -Be "[納期]は$([char]0x21b5)別途"
    }

    It "図形の行で一致しなければセル番地を出さない" {
        $row = newHitRow -location "[シート] 4月[図形]" -line "B3`t納期" -word "見積"
        $row.Prepare()
        $row.MatchCell | Should -Be ""
    }

    It "図形の行でセルが 1 つしかなければ行のまま表示する" {
        $row = newHitRow -location "[シート] 4月[図形]" -line "納期" -word "納期"
        $row.Prepare()
        $row.DisplayLine | Should -Be "納期"
    }

    It "Excel のヘッダー・フッターの行は、セル番地を出さず（一致は 1 つ）、囲みの引用符を外した文字を表示する" {
        $row = newHitRow -location "[シート]4月[ヘッダー・フッター]" -line "`"`"`"至急`"`" と `"`"確認`"`"`"" -word "至急"
        $row.Prepare()
        $row.MatchCell | Should -Be ""
        $row.MatchCount | Should -Be 1
        $row.DisplayLine | Should -Be "`"至急`" と `"確認`""
    }

    It "Excel のヘッダー・フッターの行の一致しない語は、一致 0 で番地も空" {
        $row = newHitRow -location "[シート]4月[ヘッダー・フッター]" -line "社外秘" -word "見積"
        $row.Prepare()
        $row.MatchCell | Should -Be ""
        $row.MatchCount | Should -Be 0
    }

    It "2 回目は作り直さない" {
        $row = newHitRow -line "見積" -word "見積"
        $row.Prepare()
        $row.DisplayLine = "変えた"
        $row.Prepare()
        $row.DisplayLine | Should -Be "変えた"
    }

    It "表示用の項目を作ったことを通知する" {
        $row = newHitRow -line "見積" -word "見積"
        $names = watchChanges $row
        $row.Prepare()
        # 通知の順は決まりではないため、名前の集まりで比べる
        @($names | Sort-Object) -join "," | Should -Be "DisplayLine,MatchCell,MatchCount,Segments"
    }

    It "正規表現の一致を強調する" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "第1回と第22回" -pattern ([regex]"第\d+回")
        $row.Prepare()
        describeSegments $row.Segments | Should -Be "[第1回]と[第22回]"
    }

    It "一致より前が長いときは、先頭を省いて一致の直前だけ残す" {
        $lead = "あ" * 50
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "${lead}見積" -word "見積"
        $row.Prepare()
        $row.Segments[0].Text | Should -Be ([string][char]0x2026 + ("あ" * [HitRow]::LeadLength))
        $row.Segments[1].IsHit | Should -Be $true
    }

    It "一致の後が長いときは切り詰める" {
        $rest = "い" * ([HitRow]::MaxDisplay + 10)
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "見積$rest" -word "見積"
        $row.Prepare()
        $row.Segments[1].Text | Should -Be (("い" * [HitRow]::MaxDisplay) + [char]0x2026)
    }

    It "表示する長さを超えたら、その先の一致は強調しない" {
        $line = (("う" * 100) + "見積") * 10
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line $line -word "見積"
        $row.Prepare()
        @($row.Segments | Where-Object { $_.IsHit }).Count | Should -BeLessThan 10
    }

    It "重なった一致は先のものだけ強調する" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "ああああ" -pattern ([regex]"(?=(ああ))あ")
        $row.Prepare()
        describeSegments $row.Segments | Should -Be "[あ][あ][あ]あ"
    }
}

Describe "HitRow.Contains" -Tag Unit {
    BeforeAll {
        $row = newHitRow -relDir "営業部" -book "見積.xlsx" -location "4月" -lineNumber 12 -line "品名`tABC"
        $row.PlaceText = "[シート]四月"
        $row.Kind = "セル"
    }

    It "<name>で絞り込める" -TestCases @(
        @{ name = "相対フォルダ"; text = "営業" }
        @{ name = "元ファイル名"; text = "見積.XLSX" }
        @{ name = "場所"; text = "4月" }
        @{ name = "場所の表記"; text = "四月" }
        @{ name = "種別"; text = "セル" }
        @{ name = "行（大文字小文字を区別しない）"; text = "abc" }
    ) {
        param ($name, $text)
        $row.Contains($text) | Should -Be $true
    }

    It "どれにも無ければ $false" {
        $row.Contains("請求") | Should -Be $false
    }

    It "列に出さない行番号では絞り込めない" {
        $row.Contains("12") | Should -Be $false
    }
}

Describe "HitRow.SetPlaceDisplay" -Tag Unit {
    It "「場所」の列の表記を入れて通知する。同じ値なら通知しない" {
        $row = newHitRow -line "見積" -word "見積"
        $names = watchChanges $row
        $row.SetPlaceDisplay("[シート]4月!B2")
        $row.SetPlaceDisplay("[シート]4月!B2")
        $row.PlaceDisplay | Should -Be "[シート]4月!B2"
        @($names) -join "," | Should -Be "PlaceDisplay"
    }
}

Describe "HitRow の静的な関数" -Tag Unit {
    It "ToDisplay はタブを区切り線に、セル内改行を記号にする" {
        [HitRow]::ToDisplay("a`tb$([char]0x2028)c") | Should -Be "a │ b$([char]0x21b5)c"
        [HitRow]::ToDisplay($null) | Should -Be ""
    }

    It "HasMatch は<name>" -TestCases @(
        @{ name = "語句を大文字小文字を区別せずに探す"; text = "ABC"; word = "b"; pattern = $null; expected = $true }
        @{ name = "語句が無ければ一致しない"; text = "ABC"; word = "x"; pattern = $null; expected = $false }
        @{ name = "空の文字では一致しない"; text = ""; word = "b"; pattern = $null; expected = $false }
        @{ name = "語句も正規表現も無ければ一致しない"; text = "ABC"; word = ""; pattern = $null; expected = $false }
        @{ name = "正規表現の長さ 0 の一致を数えない"; text = "ABC"; word = ""; pattern = [regex]"x*"; expected = $false }
        @{ name = "正規表現の長さ 1 以上の一致は数える"; text = "ABC"; word = ""; pattern = [regex]"x*B"; expected = $true }
        @{ name = "語句の正規表現の記号をそのままの文字として探す（1.5 は 125 に当たらない）"; text = "125"; word = "1.5"; pattern = $null; expected = $false }
        @{ name = "語句の . はそのままの文字に当たる"; text = "1.5 倍"; word = "1.5"; pattern = $null; expected = $true }
        @{ name = "語句の [ ] はそのままの文字に当たる"; text = "[シート]4月"; word = "[シート]"; pattern = $null; expected = $true }
    ) {
        param ($name, $text, $word, $pattern, $expected)
        [HitRow]::HasMatch($text, $word, $pattern) | Should -Be $expected
    }

    It "FindMatches は語句の位置と長さを返す" {
        $found = [HitRow]::FindMatches("見積と見積", "見積", $null)
        $found.Count | Should -Be 2
        ($found[0] -join ",") | Should -Be "0,2"
        ($found[1] -join ",") | Should -Be "3,2"
    }

    It "FindMatches は空の文字・空の語句で何も返さない" {
        [HitRow]::FindMatches("", "見積", $null).Count | Should -Be 0
        [HitRow]::FindMatches("見積", "", $null).Count | Should -Be 0
    }

    It "SplitCells は<name>" -TestCases @(
        @{ name = "Excel 以外はタブで分けるだけ"; line = "a`t`"b`tc`""; isExcel = $false; expected = "a|`"b|c`""; count = 3 }
        @{ name = "Excel 以外で値が無ければ空のセル 1 つ"; line = $null; isExcel = $false; expected = ""; count = 1 }
        @{ name = "Excel の引用符で囲んだセルを 1 つにし、囲みを外す"; line = "a`t`"b`tc`"`t`"x`"`"y`"`td"; isExcel = $true; expected = "a|b`tc|x`"y|d"; count = 4 }
        @{ name = "Excel の空のセルも数える"; line = "a`t`tb"; isExcel = $true; expected = "a||b"; count = 3 }
        @{ name = "Excel の引用符で囲んだセルの後ろに続く文字も同じセルにする"; line = "`"a`"b`tc"; isExcel = $true; expected = "ab|c"; count = 2 }
        @{ name = "Excel の空の行は空のセル 1 つ"; line = ""; isExcel = $true; expected = ""; count = 1 }
    ) {
        param ($name, $line, $isExcel, $expected, $count)
        $cells = [HitRow]::SplitCells($line, $isExcel)
        ($cells -join "|") | Should -Be $expected
        $cells.Count | Should -Be $count
    }

    It "ColumnName は列番号を Excel の列名にする" {
        [HitRow]::ColumnName(1) | Should -Be "A"
        [HitRow]::ColumnName(13) | Should -Be "M"
        [HitRow]::ColumnName(27) | Should -Be "AA"
        [HitRow]::ColumnName(703) | Should -Be "AAA"
        [HitRow]::ColumnName(0) | Should -Be ""
    }

    # [int](n / 26) は四捨五入になるため、切り捨てでないと 26 列目が "AZ"、15 列目が "AO" になっていた
    It "ColumnName は 26 で割った余りが大きい列も正しく出す" {
        [HitRow]::ColumnName(15) | Should -Be "O"
        [HitRow]::ColumnName(26) | Should -Be "Z"
        [HitRow]::ColumnName(702) | Should -Be "ZZ"
        [HitRow]::ColumnName(16384) | Should -Be "XFD"
    }

    It "TextWidth は半角を狭く、全角を広く数える" {
        [HitRow]::TextWidth("") | Should -Be 14
        [HitRow]::TextWidth("ab") | Should -Be 28
        [HitRow]::TextWidth("あ") | Should -Be 26
        [HitRow]::TextWidth("ｱ") | Should -Be 21
        [HitRow]::TextWidth($null) | Should -Be 14
    }

    It "CellWidth はセル内の最も長い行の幅" {
        [HitRow]::CellWidth("ab$([char]0x2028)abcd`nabc") | Should -Be ([HitRow]::TextWidth("abcd"))
        [HitRow]::CellWidth($null) | Should -Be 14
    }

    It "CellWidth は 1 行を先頭の決まった文字数までで測る" {
        [HitRow]::CellWidth("a" * 500) | Should -Be ([HitRow]::TextWidth("a" * [HitRow]::MaxWidthChars))
    }

    It "Shorten は長いときだけ切り詰めて … を付ける" {
        [HitRow]::Shorten("abcdef", 3) | Should -Be "abc…"
        [HitRow]::Shorten("abc", 3) | Should -Be "abc"
        [HitRow]::Shorten($null, 3) | Should -BeNullOrEmpty
    }
}

Describe "HitRow.BuildPreview" -Tag Unit {
    It "Excel は列名を見出しにし、選択行で一致したセルの位置を覚える" {
        $row = newHitRow -lineNumber 5 -line "品名`t見積" -word "見積"
        $table = showAll ($row.BuildPreview(@(4, 5, 6), @("品名`t単価", "品名`t見積", "a")))
        @($table.Columns | ForEach-Object { $_.Label }) | Should -Be @("A", "B")
        $table.Rows.Count | Should -Be 3
        @($table.Rows | ForEach-Object { $_.Number }) | Should -Be @("4", "5", "6")
        @($table.Rows | ForEach-Object { $_.IsHitRow }) | Should -Be @($false, $true, $false)
        $table.Rows[1].Cells[1].IsHit | Should -Be $true
        $table.Rows[1].Cells[0].IsHit | Should -Be $false
        $table.HitOffset | Should -Be ([HitRow]::NumberWidth + $table.Columns[0].Width)
        $table.HitWidth | Should -Be $table.Columns[1].Width
        # 短い行の足りないセルは空
        $table.Rows[2].Cells[1].Text | Should -Be ""
        $table.Rows[2].Cells[1].ToolTip | Should -BeNullOrEmpty
    }

    It "Excel のヘッダー・フッターは、1 行が文字だけの 1 列なので、見出しを「文字」にする" {
        $row = newHitRow -location "[シート]4月[ヘッダー・フッター]" -lineNumber 2 -line "社外秘" -word "社外秘"
        $table = showAll ($row.BuildPreview(@(1, 2), @("月次報告", "社外秘")))
        @($table.Columns | ForEach-Object { $_.Label }) | Should -Be @("文字")
        $table.Rows[1].Cells[0].IsHit | Should -Be $true
    }

    It "列の幅は下限と上限の間に収める" {
        $row = newHitRow -lineNumber 1 -line "a`t$("あ" * 80)" -word "a"
        $table = $row.BuildPreview(@(1), @("a`t$("あ" * 80)"))
        $table.Columns[0].Width | Should -Be ([HitRow]::MinCellWidth)
        $table.Columns[1].Width | Should -Be ([HitRow]::MaxCellWidth)
    }

    It "Word・PowerPoint は列番号を見出しにし、広い幅まで許す" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -lineNumber 1 -line ("あ" * 80) -word "あ"
        $table = $row.BuildPreview(@(1), @("あ" * 80))
        $table.Columns[0].Label | Should -Be "1"
        $table.Columns[0].Width | Should -BeGreaterThan ([HitRow]::MaxCellWidth)
    }

    It "Excel の図形の行は、見出しを「セル」「文字」にする" {
        $row = newHitRow -location "[シート] 4月[図形]" -lineNumber 1 -line "B3`t納期`tx" -word "納期"
        $table = $row.BuildPreview(@(1), @("B3`t納期`tx"))
        @($table.Columns | ForEach-Object { $_.Label }) | Should -Be @("セル", "文字", "3")
    }

    It "前後の行が無い・数が合わないときは、その行だけで作る" {
        $row = newHitRow -lineNumber 9 -line "品名`t見積" -word "見積"
        foreach ($args2 in @(@($null, $null), @(@(), @()), @(@(1, 2), @("a")))) {
            $table = $row.BuildPreview($args2[0], $args2[1])
            $table.Rows.Count | Should -Be 1
            $table.Rows[0].Number | Should -Be "9"
            $table.Rows[0].IsHitRow | Should -Be $true
        }
    }

    It "一致したセルが無ければ左端を表示位置にする" {
        $row = newHitRow -lineNumber 1 -line "a`tb" -word "見積"
        $table = $row.BuildPreview(@(1), @("a`tb"))
        $table.HitOffset | Should -Be 0
        $table.HitWidth | Should -Be 0
    }

    It "セル内改行は改行にし、長いセルは表示と説明を切り詰める" {
        $long = "え" * ([HitRow]::MaxToolTipChars + 5)
        $row = newHitRow -lineNumber 1 -line "a$([char]0x2028)b`t$long" -word "a"
        $table = showAll ($row.BuildPreview(@(1), @("a$([char]0x2028)b`t$long")))
        $table.Rows[0].Cells[0].Text | Should -Be "a`nb"
        $table.Rows[0].Cells[0].ToolTip | Should -Be "a`nb"
        $table.Rows[0].Cells[1].Display.Length | Should -Be ([HitRow]::MaxCellChars + 1)
        $table.Rows[0].Cells[1].ToolTip.Length | Should -Be ([HitRow]::MaxToolTipChars + 1)
        $table.Rows[0].Cells[1].Column | Should -Be $table.Columns[1]
    }

    Context "列が多いとき" {
        BeforeAll {
            # 2000 列の行（1500 列目が「見積」）を、前後の行つきで作る
            function newWideTable {
                $cells = @(1..2000 | ForEach-Object { "c$_" })
                $cells[1499] = "見積"
                $line = $cells -join "`t"
                $row = newHitRow -lineNumber 2 -line $line -word "見積"
                return $row.BuildPreview(@(1, 2, 3), @($line, $line, $line))
            }
        }

        It "全部の列を持つ（列を減らさない）。一致したセルの位置は、全部の列の中での位置" {
            $table = newWideTable
            $table.Columns.Count | Should -Be 2000
            $table.Columns[1999].Label | Should -Be "BXX"
            $expected = [HitRow]::NumberWidth
            for ($c = 0; $c -lt 1499; $c++) { $expected += $table.Columns[$c].Width }
            $table.HitOffset | Should -Be $expected
            $table.HitWidth | Should -Be $table.Columns[1499].Width
            # 範囲を決めるまでは、行の右の空きに全部の列の幅を入れて、横にスクロールできる広さにする
            $table.Window.RightSpace | Should -Be $table.Prefix()[$table.Columns.Count]
        }

        It "列の幅の累積は、幅が変わるまで使い回し、SetWidth で変わったら作り直す" {
            $table = newWideTable
            $first = $table.Prefix()
            $first.Length | Should -Be ($table.Columns.Count + 1)
            $first[0] | Should -Be 0
            $first[2] | Should -Be ($table.Columns[0].Width + $table.Columns[1].Width)
            [object]::ReferenceEquals($table.Prefix(), $first) | Should -Be $true
            $table.Columns[0].SetWidth($table.Columns[0].Width + 50)
            $second = $table.Prefix()
            [object]::ReferenceEquals($second, $first) | Should -Be $false
            $second[1] | Should -Be $table.Columns[0].Width
            $second[$table.Columns.Count] | Should -Be ($first[$table.Columns.Count] + 50)
        }

        It "セルは、見えている行の、見えている列の範囲の分だけ作る" {
            $table = newWideTable
            @($table.Rows | ForEach-Object { $_.Cells.Count } | Measure-Object -Sum).Sum | Should -Be 0
            [void]$table.SetWindow(1495, 1499, 100, 200)
            $table.EnsureCells(1, 1)
            @($table.Rows | ForEach-Object { $_.Cells.Count } | Measure-Object -Sum).Sum | Should -Be 5
            @($table.Rows[1].Cells | ForEach-Object { $_.ColumnIndex }) | Should -Be @(1495, 1496, 1497, 1498, 1499)
            $table.Rows[1].Cells[4].Text | Should -Be "見積"
            $table.Rows[1].Cells[4].IsHit | Should -Be $true
            $table.Rows[1].Cells[4].Column | Should -Be $table.Columns[1499]
            $table.Rows[0].Cells.Count | Should -Be 0
            $table.Window.LeftSpace | Should -Be 100
            $table.Window.RightSpace | Should -Be 200
            @($table.WindowColumns() | ForEach-Object { $_.Label }).Count | Should -Be 5
        }

        It "範囲が同じなら作り直さない。範囲が変わったら、作ってあるセルを捨てて作り直す" {
            $table = newWideTable
            [void]$table.SetWindow(0, 3, 0, 0)
            $table.EnsureCells(0, 2)
            $first = $table.Rows[0].Cells[0]
            @($table.Rows | ForEach-Object { $_.Cells.Count } | Measure-Object -Sum).Sum | Should -Be 12
            $table.SetWindow(0, 3, 10, 20) | Should -Be $false
            $table.EnsureCells(0, 2)
            [object]::ReferenceEquals($table.Rows[0].Cells[0], $first) | Should -Be $true -Because "範囲が同じなら、作ってあるセルをそのまま使う"
            $table.SetWindow(4, 7, 10, 20) | Should -Be $true
            $table.Rows[0].Cells.Count | Should -Be 0
            $table.EnsureCells(0, 2)
            [object]::ReferenceEquals($table.Rows[0].Cells[0], $first) | Should -Be $false
            @($table.Rows | ForEach-Object { $_.Cells.Count } | Measure-Object -Sum).Sum | Should -Be 12
            $table.Rows[2].Cells[0].ColumnIndex | Should -Be 4
        }

        It "見えなくなった行のセルは捨てる（作ってある数が、見えている範囲の分に収まる）" {
            $table = newWideTable
            [void]$table.SetWindow(0, 1, 0, 0)
            $table.EnsureCells(0, 0)
            $table.EnsureCells(2, 2)
            $table.Rows[0].Cells.Count | Should -Be 0
            $table.Rows[1].Cells.Count | Should -Be 0
            $table.Rows[2].Cells.Count | Should -Be 2
        }

        It "メッセージの行にはセルを作らない。範囲が列の数を超えていても、ある列までにする" {
            $table = newWideTable
            $table.ShowMessage("up", "読み込み中…")
            [void]$table.SetWindow(1998, 2005, 0, 0)
            $table.EnsureCells(0, $table.Rows.Count - 1)
            $table.Rows[0].Cells.Count | Should -Be 0
            $table.Rows[1].Cells.Count | Should -Be 2
            $table.Rows[1].Cells[1].ColumnIndex | Should -Be 1999
        }

        It "見えていない列にまたがる範囲を選んでコピーすると、全部のセルをタブでつなぐ" {
            $table = newWideTable
            [void]$table.SetWindow(0, 4, 0, 0)
            $table.EnsureCells(0, 2)
            $table.Select($table.Rows[1].Cells[0], $false)
            $table.Select([PreviewCell]@{ LineNumber = 2; ColumnIndex = 1999 }, $true)
            $copied = $table.GetSelectionText() -split "`t"
            $copied.Count | Should -Be 2000
            $copied[0] | Should -Be "c1"
            $copied[1499] | Should -Be "見積"
            $copied[1999] | Should -Be "c2000"
            $table.SelectedCount() | Should -Be 2000
            # 見えている列のセルには、選択の印が付く
            $table.Rows[1].Cells[4].IsSelected | Should -Be $true
            $table.Rows[0].Cells[4].IsSelected | Should -Be $false
        }

        It "選択は行の番号と列の番号で持つので、あとから作ったセルにも選択の印が付く" {
            $table = newWideTable
            [void]$table.SetWindow(0, 2, 0, 0)
            $table.EnsureCells(0, 2)
            $table.Select($table.Rows[0].Cells[0], $false)
            $table.Select($table.Rows[1].Cells[1], $true)
            [void]$table.SetWindow(1, 3, 0, 0)
            $table.EnsureCells(0, 2)
            @($table.Rows | ForEach-Object { @($_.Cells | Where-Object { $_.IsSelected }).Count }) | Should -Be @(1, 1, 0)
            $table.Rows[0].Cells[0].ColumnIndex | Should -Be 1
        }

        It "セルを作るとき、行の値が短くても列の数だけ空のセルを作る" {
            $table = newWideTable
            $row = $table.Owner.BuildPreviewRows($table, @(4), @("a"))
            $table.InsertRows("down", $row)
            [void]$table.SetWindow(0, 2, 0, 0)
            $table.EnsureCells(3, 3)
            @($table.Rows[3].Cells | ForEach-Object { $_.Text }) | Should -Be @("a", "", "")
        }
    }

    It "空の行も 1 列として出す" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -lineNumber 1 -line "" -word "x"
        $table = showAll ($row.BuildPreview(@(1), [string[]]@($null)))
        $table.Columns.Count | Should -Be 1
        $table.Rows[0].Cells[0].Text | Should -Be ""
    }
}

Describe "HitRow（境界値）" -Tag Unit {
    It "空の行は、表示も強調も空" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "" -word "見積"
        $row.Prepare()
        $row.DisplayLine | Should -Be ""
        $row.Segments.Count | Should -Be 0
    }

    It "一致しない行は、行全体を強調なしの 1 つにする" {
        $row = newHitRow -book "議事録.docx" -location "[本文]" -line "請求書" -word "見積"
        $row.Prepare()
        $row.Segments.Count | Should -Be 1
        $row.Segments[0].IsHit | Should -Be $false
    }

    It "語句の検索では、重なる一致を二重に数えない" {
        $found = [HitRow]::FindMatches("あああ", "ああ", $null)
        $found.Count | Should -Be 1
        ($found[0] -join ",") | Should -Be "0,2"
    }

    It "空の文字で絞り込むとすべて合う" {
        (newHitRow -line "見積").Contains("") | Should -Be $true
    }

    It "正規表現で Excel のセル番地を決める" {
        $row = newHitRow -lineNumber 4 -line "品名`t第1回`t第2回" -pattern ([regex]"第\d回")
        $row.Prepare()
        $row.MatchCount | Should -Be 2
    }

    It "27 列目以降の一致は 2 文字の列名で出す" {
        $cells = @(1..30 | ForEach-Object { "c$_" })
        $cells[26] = "見積"
        $row = newHitRow -lineNumber 8 -line ($cells -join "`t") -word "見積"
        $row.Prepare()
        $row.MatchCell | Should -Be "AA8"
    }

    It "選択行が前後の行に無ければ、強調する行を作らず左端を表示位置にする" {
        $row = newHitRow -lineNumber 9 -line "見積" -word "見積"
        $table = showAll ($row.BuildPreview(@(1, 2), @("見積", "x")))
        @($table.Rows | Where-Object { $_.IsHitRow }).Count | Should -Be 0
        $table.HitOffset | Should -Be 0
        # 一致したセルの印は付ける
        $table.Rows[0].Cells[0].IsHit | Should -Be $true
    }
}
