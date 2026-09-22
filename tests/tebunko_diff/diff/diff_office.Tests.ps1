# 種類ごとの比べ方（tebunko_diff\diff\diff_office.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"

function getRowText {
    # 行の種類と行番号を短い文字にする（"same:1:1 change:3:3 ..."）
    param ([object[]]$rows)

    return (@($rows | ForEach-Object { "$($_.Kind):$($_.LeftNo):$($_.RightNo)" }) -join " ")
}

Describe "Excel" -Tag Unit {
    $left = [ordered]@{
        "明細" = [string[]]@("御見積書", "", "見積先`t`t`t`tQ-2024", "品名`t数量", "サーバー構築`t10`t120,000", "初期設定`t1", "送料`t1`t500")
        "表紙" = [string[]]@("x")
    }
    $right = [ordered]@{
        "明細" = [string[]]@("御見積書", "", "見積先`t`t`t`tQ-2025", "品名`t数量", "サーバー構築`t12`t144,000", "保守費用（年間）`t1`t240,000", "初期設定`t1", "送料`t1`t500")
        "条件" = [string[]]@("y")
    }
    $result = compareOfficeUnits "Excel" $left $right
    $sheet = $result.Places[0]

    It "シート名で対応づけ、片方にしか無いシートは追加・削除にする" {
        (@($result.Places | ForEach-Object { "$($_.Name)=$($_.Status)" }) -join ",") | Should Be "明細=change,表紙=delete,条件=insert"
    }

    It "行の挿入を見つけ、下の行の対応がずれない（空の行は差分に出さない）" {
        getRowText $sheet.Rows | Should Be "same:1:1 change:3:3 same:4:4 change:5:5 insert::6 same:6:7 same:7:8"
    }

    It "変更の行は、違うセルの番地と値を出し、そのセルに印を付ける" {
        $row = $sheet.Rows[3]
        $row.Detail | Should Be "B5：10 → 12　C5：120,000 → 144,000"
        $row.RightCell | Should Be "B5"
        (@($row.RightCells | ForEach-Object { $_.Changed }) -join ",") | Should Be "False,True,True"
        (@($row.LeftCells | ForEach-Object { $_.Text }) -join "|") | Should Be "サーバー構築|10|120,000"
    }

    It "行番号が変わった行の違うセルは、左右の番地を出す" {
        $l = [ordered]@{ S = [string[]]@("a", "b`t1") }
        $r = [ordered]@{ S = [string[]]@("新", "a", "b`t2") }
        $place = (compareOfficeUnits "Excel" $l $r).Places[0]
        $place.Rows[2].Detail | Should Be "B2→B3：1 → 2"
    }

    It "セルが 1 つだけの行も、セルに分けて出す" {
        $l = [ordered]@{ S = [string[]]@("見出し") }
        $r = [ordered]@{ S = [string[]]@("見出し2") }
        $place = (compareOfficeUnits "Excel" $l $r).Places[0]
        $place.Rows[0].LeftCells[0].Text | Should Be "見出し"
    }

    It "先頭のセルが空の行も、似ていれば変更に組む（セルの一致の割合が 0 でも、文字が似ていれば組む）" {
        $l = [ordered]@{ S = [string[]]@("", "`t御見積書", "`tA社 御中") }
        $r = [ordered]@{ S = [string[]]@("", "`t御見積書（2025年度）", "`tA社 御中") }
        $place = (compareOfficeUnits "Excel" $l $r).Places[0]
        getRowText $place.Rows | Should Be "change:2:2 same:3:3"
    }

    It "空の行を足した・消しただけでは差分にしない" {
        $l = [ordered]@{ S = [string[]]@("見出し", "a`t1", "b`t2") }
        $r = [ordered]@{ S = [string[]]@("見出し", "", "`t`t", "a`t1", "", "b`t2", " ") }
        $place = (compareOfficeUnits "Excel" $l $r).Places[0]
        $place.Status | Should Be "same"
        getRowText $place.Rows | Should Be "same:1:1 same:2:4 same:3:6"
    }

    It "見えない文字だけの行（空文字の引用・ゼロ幅スペース・BOM・セル内の改行だけ）も差分にしない" {
        $l = [ordered]@{ S = [string[]]@("TC01-005`t田中 健", "TC01-006`tSmith John") }
        $r = [ordered]@{ S = [string[]]@("TC01-005`t田中 健", "TC01-006`tSmith John", '""', "$([char]0x200B)`t$([char]0xFEFF)", "`"$([char]0x2028)`"") }
        $place = (compareOfficeUnits "Excel" $l $r).Places[0]
        $place.Status | Should Be "same"
        getRowText $place.Rows | Should Be "same:1:1 same:2:2"
    }

    It "見た目が空のセルどうしは、値の違うセルにしない" {
        $l = [ordered]@{ S = [string[]]@("a`t`t1") }
        $r = [ordered]@{ S = [string[]]@("a`t$([char]0x200B)`t2") }
        $row = (compareOfficeUnits "Excel" $l $r).Places[0].Rows[0]
        $row.Detail | Should Be "C1：1 → 2"
    }

    $people = [string[]]@("ID`t氏名`t部署`t内線", "1`t山田`t営業部`t101", "2`t佐藤`t総務部`t102", "3`t鈴木`t開発部`t103", "4`t高橋`t開発部`t104")

    It "列を挿入しても、ほかの行を変更にせず、列の追加として出す" {
        $r = [string[]]@("ID`t氏名`t役職`t部署`t内線", "1`t山田`t課長`t営業部`t101", "2`t佐藤`t`t総務部`t102", "3`t鈴木`t`t開発部`t103", "4`t高橋`t主任`t開発部`t104")
        $result = compareOfficeUnits "Excel" ([ordered]@{ S = $people }) ([ordered]@{ S = $r })
        $place = $result.Places[0]
        getRowText $place.Rows | Should Be "same:1:1 same:2:2 same:3:3 same:4:4 same:5:5"
        $place.Status | Should Be "change"
        $place.ColumnNote | Should Be "列の追加 C（役職）"
        $result.ColumnInserts | Should Be 1
        (@($place.RightColumnNames) -join ",") | Should Be "A,B,C,D,E"
        (@($place.LeftColumnNames) -join ",") | Should Be "A,B,,C,D"
        (@($place.ColumnKinds) -join ",") | Should Be ",,insert,,"
        # 追加した列の値のセルに印を付け、比較元の側は空きにする
        $place.Rows[1].RightCells[2].Kind | Should Be "insert"
        $place.Rows[1].LeftCells[2].Kind | Should Be "empty"
        $place.Rows[2].RightCells[2].Kind | Should Be ""
        # 追加した列に値のある行（1・2・5 行目）はたたまずに見せる
        (@($place.FoldedRows | Where-Object { $_.Type -eq "Line" -and $_.HasColumnChange } | ForEach-Object { $_.RightNo }) -join ",") | Should Be "1,2,5"
    }

    It "列を削除しても、ほかの行を変更にせず、列の削除として出す" {
        $r = [string[]]@("ID`t氏名`t内線", "1`t山田`t101", "2`t佐藤`t102", "3`t鈴木`t103", "4`t高橋`t104")
        $result = compareOfficeUnits "Excel" ([ordered]@{ S = $people }) ([ordered]@{ S = $r })
        getRowText $result.Places[0].Rows | Should Be "same:1:1 same:2:2 same:3:3 same:4:4 same:5:5"
        $result.Places[0].ColumnNote | Should Be "列の削除 C（部署）"
        $result.ColumnDeletes | Should Be 1
    }

    It "空の列を挿入しただけでは差分にしない" {
        $r = [string[]]@($people | ForEach-Object { $parts = $_ -split "`t"; ($parts[0], "", $parts[1], $parts[2], $parts[3]) -join "`t" })
        $place = (compareOfficeUnits "Excel" ([ordered]@{ S = $people }) ([ordered]@{ S = $r })).Places[0]
        $place.Status | Should Be "same"
        (@($place.RightColumnNames) -join ",") | Should Be "A,C,D,E"
    }

    It "列を動かしたら、列の移動として出す" {
        $r = [string[]]@($people | ForEach-Object { $parts = $_ -split "`t"; ($parts[0], $parts[3], $parts[1], $parts[2]) -join "`t" })
        $result = compareOfficeUnits "Excel" ([ordered]@{ S = $people }) ([ordered]@{ S = $r })
        $result.Places[0].ColumnNote | Should Be "列の移動 D → B（内線）"
        $result.ColumnMoves | Should Be 1
        $result.ColumnInserts + $result.ColumnDeletes | Should Be 0
    }

    It "行の挿入・列の削除・セルの修正を一度にしても、直したところだけを出す" {
        $r = [string[]]@("ID`t氏名`t内線", "1`t山田 次郎`t101", "2`t佐藤`t102", "5`t中村`t105", "3`t鈴木`t103", "4`t高橋`t104")
        $place = (compareOfficeUnits "Excel" ([ordered]@{ S = $people }) ([ordered]@{ S = $r })).Places[0]
        getRowText $place.Rows | Should Be "same:1:1 change:2:2 same:3:3 insert::4 same:4:5 same:5:6"
        $place.Rows[1].Detail | Should Be "B2：山田 → 山田 次郎"
        $place.ColumnNote | Should Be "列の削除 C（部署）"
    }

    It "列の番地が左右で違う変更は、左右の番地を出す" {
        $l = [ordered]@{ S = [string[]]@("a`tb`tc", "1`t2`t3", "4`t5`t6") }
        $r = [ordered]@{ S = [string[]]@("x`ta`tb`tc", "y`t1`t2`t3", "z`t4`t5`t9") }
        $place = (compareOfficeUnits "Excel" $l $r).Places[0]
        $place.Rows[2].Detail | Should Be "C3→D3：6 → 9"
    }

    It "シート名を変えたシートは、中身が似ていれば同じシートとして比べる" {
        $result = compareOfficeUnits "Excel" ([ordered]@{ 売上 = $people; 表紙 = [string[]]@("x") }) ([ordered]@{ 売上実績 = $people; 表紙 = [string[]]@("x") })
        (@($result.Places | ForEach-Object { "$($_.LeftName)>$($_.RightName)=$($_.Status)" }) -join ",") | Should Be "売上>売上実績=change,表紙>表紙=same"
        $result.Places[0].Note | Should Be "シート名を変えました（売上 → 売上実績）。"
        $result.Inserts + $result.Deletes | Should Be 0
    }

    It "中身の違うシートは、名前が違えば追加と削除のまま" {
        $result = compareOfficeUnits "Excel" ([ordered]@{ A = $people }) ([ordered]@{ B = [string[]]@("別の", "中身") })
        (@($result.Places | ForEach-Object { $_.Status }) -join ",") | Should Be "delete,insert"
    }

    It "行を挿入して番号を振り直しても、挿入した行を別の行の変更と取り違えない" {
        $l = [ordered]@{ S = [string[]]@("No`t品名`t数量`t単価`t金額", "1`tノートPC`t10`t150,000`t1,500,000", "2`tモニター`t10`t40,000`t400,000", "3`t保守サービス`t1`t200,000`t200,000") }
        $r = [ordered]@{ S = [string[]]@("No`t品名`t数量`t単価`t金額", "1`tノートPC`t10`t150,000`t1,500,000", "2`tモニター`t10`t40,000`t400,000", "3`tキーボード`t10`t5,000`t50,000", "4`t保守サービス`t1`t200,000`t200,000") }
        $place = (compareOfficeUnits "Excel" $l $r).Places[0]
        getRowText $place.Rows | Should Be "same:1:1 same:2:2 same:3:3 insert::4 change:4:5"
        $place.Rows[4].Detail | Should Be "A4→A5：3 → 4"
    }

    It "動かした行は、移動の元と先を説明に書く" {
        $r = [string[]]@($people[0], $people[2], $people[3], $people[1], $people[4])
        $result = compareOfficeUnits "Excel" ([ordered]@{ S = $people }) ([ordered]@{ S = $r })
        $place = $result.Places[0]
        (@($place.Rows | Where-Object { $_.Detail } | ForEach-Object { "$($_.Kind):$($_.Detail)" }) -join ",") | Should Be "delete:行 4 へ移動,insert:行 2 から移動"
        # 追加・削除ではなく、移動として 1 回数える
        "$($result.Inserts),$($result.Deletes),$($result.Moves)" | Should Be "0,0,1"
        $place.Status | Should Be "change"
    }

    It "右端の空のセルの数が違うだけの行は、同じ行にする" {
        $l = [ordered]@{ S = [string[]]@("a`tb", "c`t1") }
        $r = [ordered]@{ S = [string[]]@("a`tb`t`t", "c`t2`t") }
        $place = (compareOfficeUnits "Excel" $l $r).Places[0]
        getRowText $place.Rows | Should Be "same:1:1 change:2:2"
        $place.Rows[1].Detail | Should Be "B2：1 → 2"
    }

    It "値の違うセルが無い行は、変更にしない（空白の違いを無視するとき）" {
        $l = [ordered]@{ S = [string[]]@("見積`t10", "x") }
        $r = [ordered]@{ S = [string[]]@("見積 `t10", "x") }
        $place = (compareOfficeUnits "Excel" $l $r @{ IgnoreWhitespace = $true }).Places[0]
        $place.Status | Should Be "same"
    }

    It "シート名の大文字・小文字の違いは同じシートとする" {
        $r = compareOfficeUnits "Excel" ([ordered]@{ Sheet1 = [string[]]@("a") }) ([ordered]@{ SHEET1 = [string[]]@("a") })
        @($r.Places).Count | Should Be 1
        $r.Places[0].Status | Should Be "same"
    }

    It "引用符で囲まれたセル（改行を含むセル）は、引用符を外して改行を ↵ で出す" {
        $cells = splitExcelLines ([string[]]@("`"税抜の$([char]0x2028)金額`"`tb"))
        $cells[0][0] | Should Be "税抜の ↵ 金額"
    }

    It "図形・コメントは、セル番地を行番号の欄に出し、開くときは元のシートにする" {
        $l = [ordered]@{ "見積[コメント]" = [string[]]@("C2`t税抜の金額") }
        $r = [ordered]@{ "見積[コメント]" = [string[]]@("C2`t税込の金額") }
        $row = (compareOfficeUnits "Excel" $l $r).Places[0].Rows[0]
        $row.Kind | Should Be "change"
        $row.LeftNo | Should Be "C2"
        $row.RightPlace | Should Be "見積"
        $row.RightText | Should Be "税込の金額"
    }

    It "［図形も比較］［コメントも比較］をオフにすると、その場所を比べない" {
        $l = [ordered]@{ S = [string[]]@("a"); "S[図形]" = [string[]]@("A1`tx"); "S[コメント]" = [string[]]@("A1`ty") }
        $r = [ordered]@{ S = [string[]]@("a"); "S[図形]" = [string[]]@("A1`tX"); "S[コメント]" = [string[]]@("A1`tY") }
        $result = compareOfficeUnits "Excel" $l $r @{ IncludeShapes = $false; IncludeComments = $false }
        (@($result.Places | ForEach-Object { $_.Name }) -join ",") | Should Be "S"
    }

    It "列の幅は、左右で長い方の値に合わせる（全角は 2 文字分）" {
        $widths = getGridColumnWidths (splitExcelLines ([string[]]@("a`t株式会社"))) (splitExcelLines ([string[]]@("abcdefghij`tb"))) 2
        $widths[0] | Should Be (10 * 7 + 14)
        $widths[1] | Should Be (8 * 7 + 14)
    }

    It "列の名前" {
        getColumnName 1 | Should Be "A"
        getColumnName 26 | Should Be "Z"
        getColumnName 27 | Should Be "AA"
        getColumnName 703 | Should Be "AAA"
    }
}

Describe "Word" -Tag Unit {
    It "ページをつないだ段落の並びで比べ、ページをまたいで動いた段落を差分にしない" {
        $l = [ordered]@{ "ページ001" = [string[]]@("第1条", "第2条 支払は30日以内"); "ページ002" = [string[]]@("第3条", "第4条") }
        $r = [ordered]@{ "ページ001" = [string[]]@("第1条", "追加の条文"); "ページ002" = [string[]]@("第2条 支払は45日以内", "第3条", "第4条") }
        $place = (compareOfficeUnits "Word" $l $r).Places[0]
        $place.Name | Should Be "本文"
        getRowText $place.Rows | Should Be "same:p.1 ¶1:p.1 ¶1 insert::p.1 ¶2 change:p.1 ¶2:p.2 ¶1 same:p.2 ¶1:p.2 ¶2 same:p.2 ¶2:p.2 ¶3"
        (@($place.Rows[2].RightSegs | Where-Object { $_.Changed } | ForEach-Object { $_.Text }) -join "") | Should Be "45"
    }

    It "本文・脚注・コメント・図形・ヘッダー・フッターを分けて比べる" {
        $units = [ordered]@{
            "ページ001" = [string[]]@("a"); "ページ001[図形]" = [string[]]@("図"); "ページ002[コメント]" = [string[]]@("c")
            "文書[コメント]" = [string[]]@("d"); "脚注" = [string[]]@("f"); "ヘッダー・フッター" = [string[]]@("h")
        }
        $groups = getWordGroups $units
        (@($groups.Keys) -join ",") | Should Be "本文,脚注,コメント,図形,ヘッダー・フッター"
        (@($groups["コメント"].Lines) -join ",") | Should Be "c,d"
    }
}

Describe "空の段落" -Tag Unit {
    It "Word の空の段落を足した・消しただけでは差分にしない" {
        $l = [ordered]@{ "ページ1" = [string[]]@("はじめに", "本文") }
        $r = [ordered]@{ "ページ1" = [string[]]@("はじめに", "", "　", "本文", "") }
        $result = compareOfficeUnits "Word" $l $r
        $result.Places[0].Status | Should Be "same"
        getRowText $result.Places[0].Rows | Should Be "same:p.1 ¶1:p.1 ¶1 same:p.1 ¶2:p.1 ¶4"
    }

    It "Word で動かした段落は、移動の元と先を説明に書く" {
        $l = [ordered]@{ "ページ1" = [string[]]@("第1章", "本文A", "本文B", "本文C") }
        $r = [ordered]@{ "ページ1" = [string[]]@("第1章", "本文B", "本文C", "本文A") }
        $rows = (compareOfficeUnits "Word" $l $r).Places[0].Rows
        (@($rows | Where-Object { $_.Detail } | ForEach-Object { "$($_.Kind):$($_.Detail)" }) -join ",") | Should Be "delete:p.1 ¶4 へ移動,insert:p.1 ¶2 から移動"
    }

    It "PowerPoint で並べ替えたスライドは、見出しに移動の元と先を書く" {
        $l = [ordered]@{ "スライド1" = [string[]]@("表紙"); "スライド2" = [string[]]@("目次"); "スライド3" = [string[]]@("本文") }
        $r = [ordered]@{ "スライド1" = [string[]]@("目次"); "スライド2" = [string[]]@("本文"); "スライド3" = [string[]]@("表紙") }
        $result = compareOfficeUnits "PowerPoint" $l $r
        $rows = $result.Places[0].Rows
        (@($rows | Where-Object { $_.Type -eq "Header" -and $_.Kind -ne "same" } | ForEach-Object { "$($_.Kind):$($_.Detail)" }) -join ",") | Should Be "delete:スライド 3 へ移動,insert:スライド 1 から移動"
        "$($result.SlideInserts),$($result.SlideDeletes),$($result.SlideMoves),$($result.Inserts),$($result.Deletes)" | Should Be "0,0,1,0,0"
        (getFileSummaryText $result).Counts | Should Be "スライド移動 1"
    }

    It "PowerPoint の空の段落だけが違うスライドは、同じスライドとして対応づける" {
        $l = [ordered]@{ "スライド1" = [string[]]@("表題", "本文") }
        $r = [ordered]@{ "スライド1" = [string[]]@("表題", "", "本文") }
        $result = compareOfficeUnits "PowerPoint" $l $r
        $result.Places[0].Status | Should Be "same"
    }
}

Describe "PowerPoint" -Tag Unit {
    $l = [ordered]@{
        "スライド001" = [string[]]@("表紙")
        "スライド002" = [string[]]@("概要", "導入 2025年10月"); "スライド002_ノート" = [string[]]@("税抜で説明する")
        "スライド003" = [string[]]@("費用", "初期費用 1,200 万円（税抜）")
        "スライド004" = [string[]]@("旧体制図")
        "スライド005" = [string[]]@("補足資料")
    }
    $r = [ordered]@{
        "スライド001" = [string[]]@("表紙")
        "スライド002" = [string[]]@("概要", "導入 2026年4月"); "スライド002_ノート" = [string[]]@("税込で説明する")
        "スライド003" = [string[]]@("導入スケジュール")
        "スライド004" = [string[]]@("費用", "初期費用 1,350 万円（税抜）")
        "スライド005（非表示）" = [string[]]@("補足資料")
    }
    $result = compareOfficeUnits "PowerPoint" $l $r
    $headers = @($result.Places[0].Rows | Where-Object { $_.Type -eq "Header" })

    It "スライドを中身で対応づけ、挿入・削除・番号のずれを見つける" {
        (@($headers | ForEach-Object { "$($_.Kind):$($_.LeftNo):$($_.RightNo)" }) -join " ") |
            Should Be "same:スライド 1:スライド 1 change:スライド 2:スライド 2 insert::スライド 3 change:スライド 3:スライド 4 delete:スライド 4: change:スライド 5:スライド 5"
        $headers[3].Detail | Should Be "比較元はスライド 3"
    }

    It "非表示の切り替えを変更として出す" {
        $headers[5].Detail | Should Be "非表示にした"
    }

    It "ノートは小見出しの下に比べる" {
        $rows = $result.Places[0].Rows
        $index = [Array]::IndexOf($rows, $headers[1])
        $rows[$index + 3].Type | Should Be "Section"
        $rows[$index + 3].LeftText | Should Be "ノート"
        $rows[$index + 4].Kind | Should Be "change"
    }

    It "枚数と、追加・削除したスライドの数" {
        $result.LeftSlides | Should Be 5
        $result.RightSlides | Should Be 5
        $result.SlideInserts | Should Be 1
        $result.SlideDeletes | Should Be 1
    }

    It "変わらないスライドをたたむ" {
        $folded = $result.Places[0].FoldedRows
        $folded[0].Type | Should Be "Fold"
        $folded[0].LeftText | Should Be "変わらないスライド 1 枚（スライド 1）"
    }

    It "スライドに属さない場所（ヘッダー・フッター）は別に比べる" {
        $x = compareOfficeUnits "PowerPoint" ([ordered]@{ "スライド001" = [string[]]@("a"); "ヘッダー・フッター" = [string[]]@("社外秘") }) ([ordered]@{ "スライド001" = [string[]]@("a"); "ヘッダー・フッター" = [string[]]@("社内限り") })
        (@($x.Places | ForEach-Object { "$($_.Name)=$($_.Status)" }) -join ",") | Should Be "スライドとノート=same,ヘッダー・フッター=change"
    }
}

Describe "mergePlaceOrder" -Tag Unit {
    It "比較先の順に並べ、比較元にしか無いものは元の前の場所の後ろに入れる" {
        $pairs = mergePlaceOrder @("A", "削除1", "B", "C") @("A", "B", "追加", "C")
        (@($pairs | ForEach-Object { "$($_.Left)/$($_.Right)" }) -join ",") | Should Be "A/A,削除1/,B/B,/追加,C/C"
    }

    It "比較元の先頭にしか無いものは先頭に入れる" {
        $pairs = mergePlaceOrder @("X", "A") @("A")
        (@($pairs | ForEach-Object { "$($_.Left)/$($_.Right)" }) -join ",") | Should Be "X/,A/A"
    }
}

Describe "大きすぎる場所" -Tag Unit {
    It "上限を超える行数の場所は、行ごとに比べず、同じかどうかだけを出す" {
        $place = newTooLargePlace "S" "S" "S" ([string[]]@("a", "b")) ([string[]]@("a", "c"))
        $place.Status | Should Be "change"
        $place.Note | Should Match "中身が違います"
    }
}


Describe "種類ごとの比べ方（細かい場合）" -Tag Unit {
    It "比べられない種類は例外にする" {
        { compareOfficeUnits "PDF" ([ordered]@{}) ([ordered]@{}) } | Should Throw "比べられない種類です"
    }

    It "片方の場所が無くても比べられる（空の辞書・null）" {
        $result = compareOfficeUnits "Word" $null ([ordered]@{ "ページ001" = [string[]]@("a") })
        $result.Places[0].Status | Should Be "insert"
        (getUnitLines ([ordered]@{ a = $null }) "a").Count | Should Be 0
    }

    It "違うセルが 6 つ以上なら、5 つまで出して残りの数を出す" {
        $l = [ordered]@{ S = [string[]]@("k`t10001`t10002`t10003`t10004`t10005`t10006`t10007") }
        $r = [ordered]@{ S = [string[]]@("k`t10011`t10012`t10013`t10014`t10015`t10016`t10017") }
        $row = (compareOfficeUnits "Excel" $l $r).Places[0].Rows[0]
        $row.Kind | Should Be "change"
        $row.Detail | Should Match "ほか 2 セル$"
    }

    It "上限を超える行数の場所は、行ごとに比べない（Excel・Word）" {
        $saved = ${diffMaxLines}
        try {
            ${script:diffMaxLines} = 2
            $big = [string[]]@("a", "b", "c")
            $excel = compareOfficeUnits "Excel" ([ordered]@{ S = $big }) ([ordered]@{ S = $big })
            $excel.Places[0].Note | Should Match "中身は同じです"
            $word = compareOfficeUnits "Word" ([ordered]@{ "ページ001" = $big }) ([ordered]@{ "ページ001" = [string[]]@("x", "y", "z") })
            $word.Places[0].Status | Should Be "change"
        } finally {
            ${script:diffMaxLines} = $saved
        }
    }

    It "PowerPoint の図形・コメントはスライドの中の小見出しの下に比べる" {
        $l = [ordered]@{ "スライド001" = [string[]]@("表紙"); "スライド001[図形]" = [string[]]@("図1"); "スライド001[コメント]" = [string[]]@("確認") }
        $r = [ordered]@{ "スライド001" = [string[]]@("表紙"); "スライド001[図形]" = [string[]]@("図2"); "スライド001[コメント]" = [string[]]@("確認") }
        $rows = (compareOfficeUnits "PowerPoint" $l $r).Places[0].Rows
        (@($rows | Where-Object { $_.Type -eq "Section" } | ForEach-Object { $_.LeftText }) -join ",") | Should Be "図形,コメント"
    }

    It "非表示かどうかだけが変わったスライドも、場所を変更にする" {
        $l = [ordered]@{ "スライド001" = [string[]]@("補足") }
        $r = [ordered]@{ "スライド001（非表示）" = [string[]]@("補足") }
        $place = (compareOfficeUnits "PowerPoint" $l $r).Places[0]
        $place.Status | Should Be "change"
        $place.Rows[0].Detail | Should Be "非表示にした"
    }

    It "片方だけの非表示のスライドには「非表示」を出す" {
        $inserted = compareOfficeUnits "PowerPoint" ([ordered]@{ "スライド001" = [string[]]@("a") }) ([ordered]@{ "スライド001" = [string[]]@("a"); "スライド002（非表示）" = [string[]]@("隠し") })
        (@($inserted.Places[0].Rows | Where-Object { $_.Type -eq "Header" })[1]).Detail | Should Be "非表示"
        $deleted = compareOfficeUnits "PowerPoint" ([ordered]@{ "スライド001" = [string[]]@("a"); "スライド002（非表示）" = [string[]]@("隠し") }) ([ordered]@{ "スライド001" = [string[]]@("a") })
        (@($deleted.Places[0].Rows | Where-Object { $_.Type -eq "Header" })[1]).Detail | Should Be "非表示"
    }

    It "スライドの名前は本文の最初の段落（空なら空、長ければ切る）" {
        getSlideTitle @{ Body = [string[]]@("", " ") } | Should Be ""
        getSlideTitle @{ Body = [string[]]@("あ" * 50) } | Should Be (("あ" * 40) + "…")
    }
}
