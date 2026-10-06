# ナビ（左の欄）の画面の並びと切り替え（tebunko\ui\shell\nav_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\shell\nav_view.ps1"
}

Describe "getScreenOrder" -Tag Unit {
    It "検索・インデックス管理・設定・Office の停止の順に並ぶ" {
        @(getScreenOrder) | Should -Be @("SearchTab", "IndexTab", "SettingsTab", "KillTab")
    }
}

Describe "isScreenName" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "画面の名前は true"; value = "SettingsTab"; expected = $true }
        @{ name = "知らない名前は false"; value = "Tabs"; expected = $false }
        @{ name = "空は false"; value = ""; expected = $false }
    ) {
        (isScreenName $value) | Should -Be $expected
    }
}

Describe "getNextScreen" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "進むと次の画面"; current = "SearchTab"; step = 1; expected = "IndexTab" }
        @{ name = "戻ると前の画面"; current = "SettingsTab"; step = -1; expected = "IndexTab" }
        @{ name = "末尾から進むと先頭に回る"; current = "KillTab"; step = 1; expected = "SearchTab" }
        @{ name = "先頭から戻ると末尾に回る"; current = "SearchTab"; step = -1; expected = "KillTab" }
        @{ name = "まだ選んでいないとき、進むなら先頭"; current = ""; step = 1; expected = "SearchTab" }
        @{ name = "まだ選んでいないとき、戻るなら末尾"; current = ""; step = -1; expected = "KillTab" }
    ) {
        (getNextScreen $current $step) | Should -Be $expected
    }
}

Describe "getShortcutAction" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "Ctrl+F は検索の画面で検索ワード欄"; key = "F"; ctrl = $true; shift = $false; current = "KillTab"; action = "FocusSearchWord"; screen = "SearchTab" }
        @{ name = "Ctrl+Shift+F は検索の画面で絞り込み欄"; key = "F"; ctrl = $true; shift = $true; current = "IndexTab"; action = "FocusFilter"; screen = "SearchTab" }
        @{ name = "Ctrl+Tab は次の画面（検索からインデックス管理）"; key = "Tab"; ctrl = $true; shift = $false; current = "SearchTab"; action = "SwitchScreen"; screen = "IndexTab" }
        @{ name = "Ctrl+Tab は設定から Office の停止"; key = "Tab"; ctrl = $true; shift = $false; current = "SettingsTab"; action = "SwitchScreen"; screen = "KillTab" }
        @{ name = "Ctrl+Tab は末尾から先頭に回る"; key = "Tab"; ctrl = $true; shift = $false; current = "KillTab"; action = "SwitchScreen"; screen = "SearchTab" }
        @{ name = "Ctrl+Shift+Tab は前の画面"; key = "Tab"; ctrl = $true; shift = $true; current = "SettingsTab"; action = "SwitchScreen"; screen = "IndexTab" }
        @{ name = "Ctrl+Shift+Tab は先頭から末尾に回る"; key = "Tab"; ctrl = $true; shift = $true; current = "SearchTab"; action = "SwitchScreen"; screen = "KillTab" }
        @{ name = "F5 は今の画面のまま読み直す"; key = "F5"; ctrl = $false; shift = $false; current = "KillTab"; action = "Refresh"; screen = "KillTab" }
        @{ name = "Escape は検索の取り消し"; key = "Escape"; ctrl = $false; shift = $false; current = "SearchTab"; action = "CancelSearch"; screen = "SearchTab" }
        @{ name = "Ctrl なしの F は扱わない"; key = "F"; ctrl = $false; shift = $false; current = "SearchTab"; action = "None"; screen = "SearchTab" }
        @{ name = "Ctrl 付きの F5 も読み直す（修飾キーは見ない）"; key = "F5"; ctrl = $true; shift = $false; current = "KillTab"; action = "Refresh"; screen = "KillTab" }
        @{ name = "Shift 付きの Escape も検索の取り消し（修飾キーは見ない）"; key = "Escape"; ctrl = $false; shift = $true; current = "SearchTab"; action = "CancelSearch"; screen = "SearchTab" }
        @{ name = "Ctrl なしの Tab は扱わない"; key = "Tab"; ctrl = $false; shift = $false; current = "SearchTab"; action = "None"; screen = "SearchTab" }
    ) {
        $result = getShortcutAction $key $ctrl $shift $current
        $result.Action | Should -Be $action
        $result.Screen | Should -Be $screen
    }
}
