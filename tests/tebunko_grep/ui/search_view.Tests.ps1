# ［2 検索］の判断（tebunko_grep\ui\search_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\ui\search_view.ps1"

Describe "describeSearchOption" -Tag Unit {
    It "既定のままなら空" {
        describeSearchOption @{ CaseSensitive = $false; FileFilter = "" } | Should Be ""
    }

    It "大文字と小文字の区別を出す" {
        describeSearchOption @{ CaseSensitive = $true; FileFilter = "" } | Should Be "大文字と小文字を区別"
    }

    It "対象ファイルを出す" {
        describeSearchOption @{ CaseSensitive = $false; FileFilter = "*.xlsx" } | Should Be "対象ファイル：*.xlsx"
    }

    It "両方あれば中黒でつなぐ" {
        describeSearchOption @{ CaseSensitive = $true; FileFilter = "*.xlsx" } | Should Be "大文字と小文字を区別・対象ファイル：*.xlsx"
    }
}

Describe "getWordNotice" -Tag Unit {
    It "正規表現でなければ出さない" {
        getWordNotice "(" $false | Should Be ""
    }

    It "正規表現として正しければ出さない" {
        getWordNotice "見積.*確定" $true | Should Be ""
    }

    It "空のワードでは出さない" {
        getWordNotice "" $true | Should Be ""
    }

    It "正規表現として不正なら、文字どおり検索すると伝える" {
        getWordNotice "(" $true | Should Be "正規表現として不正なため、文字どおり検索します。"
    }
}

Describe "newSearchButtonState" -Tag Unit {
    It "検索中は［中止］にする" {
        $state = newSearchButtonState $true $false "見積" $true 1
        $state.Content | Should Be "中止"
        $state.Enabled | Should Be $true
    }

    It "中止を頼んだ後は押せない" {
        (newSearchButtonState $true $true "見積" $true 1).Enabled | Should Be $false
    }

    It "ワード・インデックス・検索対象がそろえば押せる" {
        $state = newSearchButtonState $false $false "見積" $true 2
        $state.Content | Should Be "検索"
        $state.Enabled | Should Be $true
    }

    It "ワードが空なら押せない" {
        (newSearchButtonState $false $false "" $true 2).Enabled | Should Be $false
    }

    It "インデックスが無ければ押せない" {
        (newSearchButtonState $false $false "見積" $false 2).Enabled | Should Be $false
    }

    It "検索対象が選ばれていなければ押せない" {
        (newSearchButtonState $false $false "見積" $true 0).Enabled | Should Be $false
    }
}

Describe "getAppKind" -Tag Unit {
    It "拡張子からアプリの種類を返す（大文字・小文字は問わない）" {
        getAppKind "見積.xlsx" | Should Be "Excel"
        getAppKind "古い見積.XLS" | Should Be "Excel"
        getAppKind "マクロ.xlsm" | Should Be "Excel"
        getAppKind "報告書.docx" | Should Be "Word"
        getAppKind "報告書.doc" | Should Be "Word"
        getAppKind "提案.pptx" | Should Be "PowerPoint"
    }

    It "Office のファイルでなければ空" {
        getAppKind "メモ.txt" | Should Be ""
        getAppKind "" | Should Be ""
    }
}

Describe "formatLocationLabel" -Tag Unit {
    It "Excel はシート名に「シート」を付ける" {
        formatLocationLabel "見積.xlsx" "4月" | Should Be "シート 4月"
    }

    It "Word のページは「N ページ」にする" {
        formatLocationLabel "報告書.docx" "ページ003" | Should Be "3 ページ"
    }

    It "PowerPoint のスライドは「スライド N」にし、非表示・ノートの印を残す" {
        formatLocationLabel "提案.pptx" "スライド007" | Should Be "スライド 7"
        formatLocationLabel "提案.pptx" "スライド003（非表示）" | Should Be "スライド 3（非表示）"
        formatLocationLabel "提案.pptx" "スライド003_ノート" | Should Be "スライド 3 ノート"
    }

    It "ページ・スライドでない場所はそのまま" {
        formatLocationLabel "報告書.docx" "ヘッダー・フッター" | Should Be "ヘッダー・フッター"
    }
}

Describe "describeFileLocations" -Tag Unit {
    It "場所が無ければ空" {
        describeFileLocations @() | Should Be ""
    }

    It "1 か所ならその場所" {
        describeFileLocations @("シート 4月") | Should Be "シート 4月"
    }

    It "2 か所以上なら先頭と、ほかの数" {
        describeFileLocations @("シート 4月", "シート 5月", "シート 6月") | Should Be "シート 4月 ほか 2 か所"
    }
}