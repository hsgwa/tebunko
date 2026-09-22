# Excel の表示形式（shared\office\office_numfmt.ps1）のテスト。
#
# 期待する文字は、テストデータの Excel（tests\testdata\office\Excel\セル内容.xlsx・日付.xlsx）を
# いつものインデクサ（Excel のテキスト保存）で取り込んだ結果に合わせてある。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\office\office_numfmt.ps1"

Describe "formatExcelCellText（数値）" -Tag Unit {
    It "標準は Excel と同じけた数にし、大きい数は指数にする" {
        formatExcelCellText "1234567" "number" "General" | Should Be "1234567"
        formatExcelCellText "0" "number" "General" | Should Be "0"
        formatExcelCellText "-12.5" "number" "General" | Should Be "-12.5"
        formatExcelCellText "1.23e20" "number" "General" | Should Be "1.23E+20"
        formatExcelCellText "12345678901234567890" "number" "General" | Should Be "1.23457E+19"
        formatExcelCellText "4901234567894" "number" "General" | Should Be "4.90123E+12"
    }

    It "桁区切り・小数・百分率・通貨の文字を当てる" {
        formatExcelCellText "1234567" "number" "#,##0" | Should Be "1,234,567"
        formatExcelCellText "1234567" "number" '#,##0"円"' | Should Be "1,234,567円"
        formatExcelCellText "3.14159" "number" "0.00" | Should Be "3.14"
        formatExcelCellText "0.125" "number" "0.0%" | Should Be "12.5%"
        formatExcelCellText "1000001" "number" "000-0000" | Should Be "100-0001"
    }

    It "節（正の数・負の数・ゼロ）を使い分ける" {
        formatExcelCellText "-1500" "number" '#,##0;"▲"#,##0' | Should Be "▲1,500"
        formatExcelCellText "1500" "number" '#,##0;\-#,##0;' | Should Be "1,500"
        formatExcelCellText "-1500" "number" '#,##0;\-#,##0;' | Should Be "-1,500"
        formatExcelCellText "0" "number" '#,##0;\-#,##0;' | Should Be ""
    }

    It "分数は、? のけた数で表せるいちばん近い分数にする" {
        formatExcelCellText "0.75" "number" "# ?/?" | Should Be " 3/4"
        formatExcelCellText "2.25" "number" "# ?/?" | Should Be "2 1/4"
        formatExcelCellText "0.3" "number" "# ??/??" | Should Be " 3/10"
        formatExcelCellText "5" "number" "# ?/?" | Should Be "5"
    }

    It "色・条件・ロケールの指定は出さず、引用符とエスケープの文字はそのまま出す" {
        formatExcelCellText "1500" "number" '[Red]#,##0' | Should Be "1,500"
        formatExcelCellText "1500" "number" '[$-411]#,##0' | Should Be "1,500"
        formatExcelCellText "1500" "number" '#,##0\円' | Should Be "1,500円"
        formatExcelCellText "1500" "number" '_(#,##0_)' | Should Be " 1,500 "
    }

    It "数値として読めない値・文字列・真偽値・エラーはそのまま扱う" {
        formatExcelCellText "abc" "number" "#,##0" | Should Be "abc"
        formatExcelCellText "テキスト" "text" "@" | Should Be "テキスト"
        formatExcelCellText "1" "boolean" "General" | Should Be "TRUE"
        formatExcelCellText "0" "boolean" "General" | Should Be "FALSE"
        formatExcelCellText "#DIV/0!" "error" "General" | Should Be "#DIV/0!"
    }
}

Describe "formatExcelCellText（日付・時刻）" -Tag Unit {
    It "日付・時刻・日時の書式を当てる" {
        formatExcelCellText "45383" "number" "yyyy/mm/dd" | Should Be "2024/04/01"
        formatExcelCellText "45383" "number" "yyyy/m/d" | Should Be "2024/4/1"
        formatExcelCellText "0.524259259259259" "number" "hh:mm:ss" | Should Be "12:34:56"
        formatExcelCellText "45383.5240740741" "number" "yyyy/mm/dd hh:mm" | Should Be "2024/04/01 12:34"
        formatExcelCellText "45383.7569444444" "number" "yyyy/mm/dd hh:mm:ss" | Should Be "2024/04/01 18:10:00"
        formatExcelCellText "0.75" "number" "h:mm AM/PM" | Should Be "6:00 PM"
        formatExcelCellText "0.25" "number" "h:mm A/P" | Should Be "6:00 A"
    }

    It "曜日（aaa）と和暦（ggge）を出す" {
        formatExcelCellText "45383" "number" "[$-ja-JP]yyyy/m/d(aaa)" | Should Be "2024/4/1(月)"
        formatExcelCellText "45383" "number" "yyyy/m/d(aaaa)" | Should Be "2024/4/1(月曜日)"
        formatExcelCellText "45383" "number" '[$-411]ggge"年"m"月"d"日"' | Should Be "令和6年4月1日"
        formatExcelCellText "32874" "number" '[$-411]ggge"年"m"月"d"日"' | Should Be "平成2年1月1日"
        formatExcelCellText "45383" "number" '[$-411]gge"年"' | Should Be "R6年"
        # 明治より前は元号にできないため、書式の文字をそのまま出す
        formatExcelCellText "1" "number" '[$-411]ggge"年"' | Should Match "年$"
    }

    It "経過時間（[h] [mm] [ss]）は 24 時間を超えても足し続ける" {
        formatExcelCellText "1.5" "number" "[h]:mm" | Should Be "36:00"
        formatExcelCellText "1.5" "number" "[h]:mm:ss" | Should Be "36:00:00"
        formatExcelCellText "0.0006944444" "number" "mm:ss" | Should Be "01:00"
    }

    It "1900 年方式・1904 年方式のシリアル値を、Excel と同じ日付にする" {
        formatExcelCellText "1" "number" "yyyy/mm/dd" | Should Be "1900/01/01"
        formatExcelCellText "59" "number" "yyyy/mm/dd" | Should Be "1900/02/28"
        # Excel にしかない 1900/02/29
        formatExcelCellText "60" "number" "yyyy/mm/dd" | Should Be "1900/02/29"
        formatExcelCellText "61" "number" "yyyy/mm/dd" | Should Be "1900/03/01"
        formatExcelCellText "45351" "number" "yyyy/mm/dd" | Should Be "2024/02/29"
        formatExcelCellText "2958465" "number" "yyyy/mm/dd" | Should Be "9999/12/31"
        formatExcelCellText "45383" "number" "yyyy/mm/dd" $true | Should Be "2028/04/02"
        # 日付にできない値（負・上限より後ろ）は、数として出す
        formatExcelCellText "-0.5" "number" "yyyy/mm/dd" | Should Be "-0.5"
        formatExcelCellText "2958466" "number" "yyyy/mm/dd" | Should Be "2958466"
    }
}

Describe "getNumberFormatCode / splitNumberFormatSections" -Tag Unit {
    It "ブックに書かれた書式を優先し、無ければ組み込みの書式にする" {
        getNumberFormatCode "179" @{ "179" = '#,##0"円"' } | Should Be '#,##0"円"'
        getNumberFormatCode "3" @{} | Should Be "#,##0"
        getNumberFormatCode "49" @{} | Should Be "@"
        # 日本語版の Excel が使う組み込みの和暦
        getNumberFormatCode "58" @{} | Should Be '[$-411]ggge"年"m"月"d"日"'
        getNumberFormatCode "999" @{} | Should Be "General"
        getNumberFormatCode "" @{} | Should Be "General"
    }

    It "; で節に分ける（引用符の中とエスケープした ; では分けない）" {
        (splitNumberFormatSections '#,##0;-#,##0;0;@') -join "|" | Should Be '#,##0|-#,##0|0|@'
        (splitNumberFormatSections '"a;b";0') -join "|" | Should Be '"a;b"|0'
        (splitNumberFormatSections '0\;0') -join "|" | Should Be '0\;0'
        (splitNumberFormatSections "0") -join "|" | Should Be "0"
    }
}

Describe "getExcelDateTime" -Tag Unit {
    It "シリアル値を日時にする（読めない値は $null）" {
        (getExcelDateTime 45383).ToString("yyyy-MM-dd") | Should Be "2024-04-01"
        (getExcelDateTime 45383 $true).ToString("yyyy-MM-dd") | Should Be "2028-04-02"
        # 秒より細かい端数は秒に丸める
        (getExcelDateTime 45383.7569444444).ToString("HH:mm:ss") | Should Be "18:10:00"
        getExcelDateTime -1 | Should BeNullOrEmpty
        getExcelDateTime 2958466 | Should BeNullOrEmpty
    }

    It "Excel にしかない 1900/02/29 を見分ける" {
        testExcelFakeLeapDay 60 | Should Be $true
        testExcelFakeLeapDay 60 $true | Should Be $false
        testExcelFakeLeapDay 61 | Should Be $false
    }
}

Describe "formatExcelCellText（あまり使わない書き方）" -Tag Unit {
    It "埋め文字（*）・幅合わせ（_）・閉じていない引用符・行末の \ を読み飛ばす" {
        formatExcelCellText "1500" "number" '*-#,##0' | Should Be "1,500"
        formatExcelCellText "1500" "number" '#,##0_' | Should Be "1,500 "
        formatExcelCellText "1500" "number" '#,##0"円' | Should Be "1,500円"
        formatExcelCellText "1500" "number" '#,##0\' | Should Be "1,500"
    }

    It "書式が文字だけなら、その文字を出す" {
        formatExcelCellText "5" "number" '"（未定）"' | Should Be "（未定）"
    }

    It "文字列の節（4 つ目）があれば、文字列にも当てる" {
        formatExcelCellText "メモ" "text" '0;0;0;"[" @ "]"' | Should Be "[ メモ ]"
        # 節が 3 つまでなら、文字列はそのまま
        formatExcelCellText "メモ" "text" '0;0;0' | Should Be "メモ"
    }

    It "月と曜日の名前（mmm・mmmm・ddd・dddd）を出す" {
        formatExcelCellText "45383" "number" "yyyy-mmm-dd" | Should Match "^2024-"
        formatExcelCellText "45383" "number" "mmmm" | Should Not BeNullOrEmpty
        formatExcelCellText "45383" "number" "ddd" | Should Be "月"
        formatExcelCellText "45383" "number" "dddd" | Should Be "月曜日"
    }

    It "経過時間は分（[m]）・秒（[s]）でも出せる" {
        formatExcelCellText "1.5" "number" "[m]" | Should Be "2160"
        formatExcelCellText "0.5" "number" "[s]" | Should Be "43200"
    }

    It "分母の決まった分数（# ?/8）と、繰り上がる分数を出す" {
        formatExcelCellText "0.375" "number" "# ?/8" | Should Be " 3/8"
        formatExcelCellText "2.5" "number" "# ?/8" | Should Be "2 4/8"
        # 分数にすると整数になる値
        formatExcelCellText "0.99" "number" "# ?/?" | Should Be "1"
        # 整数の部分を出さない書式
        formatExcelCellText "2.25" "number" "?/?" | Should Be "9/4"
        formatExcelCellText "-0.75" "number" "# ?/?" | Should Be "- 3/4"
        # 分数として読めない書き方は、数のまま出す
        formatExcelCellText "0.75" "number" "# ?/?/?" | Should Be "0.75"
    }

    It "標準の書式で、けた数が多い数は Excel と同じところで丸める" {
        formatExcelCellText "1234567.891234" "number" "General" | Should Be "1234567.8912"
        formatExcelCellText "0.000000000001" "number" "General" | Should Be "1E-12"
    }

    It "1904 年方式でも、日付にできない値は数のまま出す" {
        formatExcelCellText "-1" "number" "yyyy/mm/dd" $true | Should Be "-1"
        getExcelDateTime 2958466 $true | Should BeNullOrEmpty
    }

    It "けた数を埋める（padNumberText）" {
        padNumberText "5" 3 | Should Be "005"
        padNumberText "-5" 3 | Should Be "-005"
        padNumberText "1234" 2 | Should Be "1234"
    }
}
