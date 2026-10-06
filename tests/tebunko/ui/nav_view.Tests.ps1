# ナビ（左の欄）の画面の並びと切り替え（tebunko\ui\nav_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\nav_view.ps1"
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
