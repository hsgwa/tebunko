# 制限モードの検索結果のブックの XML（tebunko_grep\restricted\result_book_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\restricted\result_book_view.ps1"

function newNs {
    param ($doc, $map)
    $manager = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    foreach ($key in $map.Keys) { $manager.AddNamespace($key, $map[$key]) }
    return , $manager
}

Describe "escapeXlsxText" -Tag Unit {
    It "& < > `" をエンティティにする（タブ・改行はそのまま）" {
        escapeXlsxText "a&b<c>d`"e`tf`ng" | Should Be "a&amp;b&lt;c&gt;d&quot;e`tf`ng"
    }

    It "XML に書けない制御文字は _xHHHH_ にし、元からある _xHHHH_ は _x005F_ で守る" {
        escapeXlsxText "a$([char]1)b$([char]0x1F)c" | Should Be "a_x0001_b_x001F_c"
        escapeXlsxText "_x0041_" | Should Be "_x005F_x0041_"
        escapeXlsxText "a$([char]0xFFFF)" | Should Be "a_xFFFF_"
    }

    It "32767 文字を超える文字は切り詰める" {
        (escapeXlsxText ("あ" * 40000)).Length | Should Be 32767
    }
}

Describe "getXlsxTextXml" -Tag Unit {
    It "一致が無ければ 1 つの <t> にする" {
        getXlsxTextXml "a<b" | Should Be '<t xml:space="preserve">a&lt;b</t>'
        getXlsxTextXml "abc" ([regex]"x") | Should Be '<t xml:space="preserve">abc</t>'
    }

    It "一致した部分だけを色付きの区切りにする（複数の一致・先頭・末尾）" {
        $xml = getXlsxTextXml "りんごとりんご飴" ([regex]"りんご")
        $xml | Should Be ('<r>' + ${resultHighlightRunProperties} + '<t xml:space="preserve">りんご</t></r>' +
            '<r><t xml:space="preserve">と</t></r>' +
            '<r>' + ${resultHighlightRunProperties} + '<t xml:space="preserve">りんご</t></r>' +
            '<r><t xml:space="preserve">飴</t></r>')
    }

    It "長さ 0 の一致は色を付けない" {
        getXlsxTextXml "abc" ([regex]"x*") | Should Be '<t xml:space="preserve">abc</t>'
    }
}

Describe "toXlsxLinkTarget / toXlsxLocation" -Tag Unit {
    It "file:/// に続けてパスを書き、% # 空白だけを %XX にする" {
        toXlsxLinkTarget "C:\元 データ #1\a b%.xlsx" | Should Be "file:///C:\元%20データ%20%231\a%20b%25.xlsx"
        toXlsxLinkTarget "\\server\share\見積.docx" | Should Be "file:///\\server\share\見積.docx"
    }

    It "シート名を ' で囲み、中の ' は重ねる" {
        toXlsxLocation "売上 '1'" "C12" | Should Be "'売上 ''1'''!C12"
    }
}

Describe "getResultSheetXml" -Tag Unit {
    $rows = @(
        @{ Style = ${resultStyleHeader}; Level = 0; Texts = @("ファイル", "場所", "種別", "行", "A") },
        @{ Style = ${resultStyleGroup}; Level = 0; Texts = @("見積\a.xlsx", "フォルダを開く", "1 件"); Styles = @{ 0 = ${resultStyleGroupLink}; 1 = ${resultStyleGroupLink} }; Links = @{ 0 = @{ Target = "C:\data\a.xlsx" }; 1 = @{ Target = "C:\data" } } },
        @{ Style = ${resultStyleNormal}; Level = 1; Collapsed = $true; HighlightFrom = 4; Texts = @("a.xlsx", "[シート] 売上", "セル", "2", "りんご & =SUM(A1)"); Styles = @{ 0 = ${resultStyleLink} }; Links = @{ 0 = @{ Target = "C:\data\a.xlsx"; Location = "'売上'!A2" } } },
        @{ Style = ${resultStyleContext}; Level = 2; Hidden = $true; Texts = @("", "", "前の行", "1", "品名<x>") }
    )
    $result = getResultSheetXml $rows 5 @(40, 24) ([regex]"りんご")
    [xml]$sheet = $result.Sheet
    [xml]$rels = $result.Rels
    $ns = @{ m = "http://schemas.openxmlformats.org/spreadsheetml/2006/main" }

    It "XML として読める" {
        $sheet.worksheet | Should Not BeNullOrEmpty
        $rels.Relationships | Should Not BeNullOrEmpty
    }

    It "行の深さ・畳み込み・隠す行を書く" {
        $r = @($sheet.worksheet.sheetData.row)
        $r.Count | Should Be 4
        $r[2].outlineLevel | Should Be "1"
        $r[2].collapsed | Should Be "1"
        $r[3].outlineLevel | Should Be "2"
        $r[3].hidden | Should Be "1"
        $sheet.worksheet.sheetFormatPr.outlineLevelRow | Should Be "2"
        $sheet.worksheet.sheetPr.outlinePr.summaryBelow | Should Be "0"
    }

    It "見出しを固定し、オートフィルターを付ける" {
        $sheet.worksheet.sheetViews.sheetView.pane.state | Should Be "frozen"
        $sheet.worksheet.autoFilter.ref | Should Be "A1:E4"
    }

    It "セルはすべて文字列（inlineStr）で、=SUM も文字のまま" {
        @($sheet.SelectNodes("//m:c", (newNs $sheet $ns)) | Where-Object { $_.t -ne "inlineStr" }).Count | Should Be 0
        $sheet.OuterXml.Contains("<f>") | Should Be $false
        $result.Sheet.Contains("=SUM(A1)") | Should Be $true
    }

    It "一致した部分に色を付けるのは HighlightFrom より右のセルだけ" {
        $hitRow = @($sheet.worksheet.sheetData.row)[2]
        @($hitRow.c)[4].is.r.Count | Should Be 2
        @($hitRow.c)[0].is.t.'#text' | Should Be "a.xlsx"
    }

    It "リンク: 同じ行き先は 1 つのリレーションにまとめ、シート・セルは location に書く" {
        @($rels.Relationships.Relationship).Count | Should Be 2
        @($rels.Relationships.Relationship | ForEach-Object { $_.Target }) -contains "file:///C:\data\a.xlsx" | Should Be $true
        $links = @($sheet.worksheet.hyperlinks.hyperlink)
        $links.Count | Should Be 3
        ($links | Where-Object { $_.ref -eq "A3" }).location | Should Be "'売上'!A2"
        $result.LinkCount | Should Be 3
    }

    It "書式・リンク・色付けの無い行は、セルの番地を省いてまとめて書く（空のセルも書く）" {
        $context = @($sheet.worksheet.sheetData.row)[3]
        @($context.c).Count | Should Be 5
        @($context.c)[0].r | Should BeNullOrEmpty
        @($context.c)[4].is.t.'#text' | Should Be "品名<x>"
    }

    It "列の幅は指定どおり（足りない列は 20）" {
        $cols = @($sheet.worksheet.cols.col)
        $cols.Count | Should Be 5
        $cols[0].width | Should Be "40"
        $cols[4].width | Should Be "20"
    }
}

Describe "getResultSheetXml（書けない文字・長い文字・行が無い）" -Tag Unit {
    It "書式の無い行でも、XML に書けない文字と長すぎる文字は直す" {
        $rows = @(@{ Style = 0; Level = 0; Texts = @("a$([char]1)", ("x" * 40000)) })
        [xml]$sheet = (getResultSheetXml $rows 2).Sheet
        @($sheet.worksheet.sheetData.row.c)[0].is.t.'#text' | Should Be "a_x0001_"
        @($sheet.worksheet.sheetData.row.c)[1].is.t.'#text'.Length | Should Be 32767
    }

    It "行が無くても、列の数が 0 でも XML になる" {
        [xml]$sheet = (getResultSheetXml @() 0).Sheet
        $sheet.worksheet.dimension.ref | Should Be "A1:A1"
        [xml]$empty = (getResultSheetXml @(@{ Style = 0; Level = 0; Texts = @() }) 1).Sheet
        @($empty.worksheet.sheetData.row).Count | Should Be 1
    }

    It "列の数より多いセルは書かない" {
        [xml]$sheet = (getResultSheetXml @(@{ Style = 0; Level = 0; Texts = @("a", "b", "c") }) 2).Sheet
        @($sheet.worksheet.sheetData.row.c).Count | Should Be 2
    }
}

Describe "getResultInfoSheetXml / getResultBookParts" -Tag Unit {
    It "条件のシートと、ブックの部品がどれも XML として読める" {
        $sheet = getResultSheetXml @(@{ Style = 1; Level = 0; Texts = @("ファイル") }) 1
        $parts = getResultBookParts $sheet (getResultInfoSheetXml @(, @("検索ワード", "a<b"))) 1 1
        @($parts.Keys) -contains "[Content_Types].xml" | Should Be $true
        @($parts.Keys) -contains "xl/worksheets/_rels/sheet1.xml.rels" | Should Be $true
        foreach ($name in $parts.Keys) {
            { [xml]$parts[$name] } | Should Not Throw
        }
        ([xml]$parts["xl/workbook.xml"]).workbook.sheets.sheet[1].name | Should Be "条件"
        [xml]$info = $parts["xl/worksheets/sheet2.xml"]
        $infoRows = $info.SelectNodes("//m:row", (newNs $info @{ m = "http://schemas.openxmlformats.org/spreadsheetml/2006/main" }))
        $infoRows[1].InnerText | Should Be "検索ワードa<b"
    }
}

Describe "getXlsxTextXml / getResultSheetXml（長い文字・書けない文字のヒット）" -Tag Unit {
    It "一致を探す前に 32767 文字に切り詰める" {
        $xml = getXlsxTextXml (("x" * 40000) + "りんご") ([regex]"りんご")
        $xml.Contains("りんご") | Should Be $false
    }

    It "色を付ける列でも、一致の無いセルの書けない文字は直す" {
        $rows = @(@{ Style = 0; Level = 1; HighlightFrom = 1; Texts = @("a", "b$([char]1)") })
        [xml]$sheet = (getResultSheetXml $rows 2 @() ([regex]"りんご")).Sheet
        @($sheet.worksheet.sheetData.row.c)[1].is.t.'#text' | Should Be "b_x0001_"
    }
}
