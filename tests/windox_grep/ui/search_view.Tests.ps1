# ［2 検索］の判断（windox_grep\ui\search_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\windox_grep\ui\search_view.ps1"

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