# ステータスバーの更新中の 1 行（tebunko\ui\shell\status_bar_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\shell\status_bar_view.ps1"
}

Describe "getIndexingStatusLine" -Tag Unit {
    It "名前・件数・失敗・残り時間をつなげる" {
        (getIndexingStatusLine "営業部2025" 205 455 1 "残り約 3 分") | Should -Be "更新中: 営業部2025　205 / 455 件（失敗 1 件）・残り約 3 分"
    }

    It "<case>" -TestCases @(
        @{ case = "失敗が無ければ失敗を出さない"; name = "営業部"; processed = 10; total = 20; failed = 0; eta = "残り約 2 分"; expected = "更新中: 営業部　10 / 20 件・残り約 2 分" }
        @{ case = "残り時間がまだ無ければ出さない"; name = "営業部"; processed = 10; total = 20; failed = 2; eta = ""; expected = "更新中: 営業部　10 / 20 件（失敗 2 件）" }
        @{ case = "1,000 以上は桁区切りにする"; name = "営業部"; processed = 1200; total = 2530; failed = 0; eta = ""; expected = "更新中: 営業部　1,200 / 2,530 件" }
        @{ case = "名前が取れないときは名前を出さない"; name = ""; processed = 5; total = 9; failed = 0; eta = ""; expected = "更新中　5 / 9 件" }
        @{ case = "1 件も終わっていない間は件数を出さない"; name = "営業部"; processed = 0; total = 20; failed = 0; eta = ""; expected = "更新中: 営業部　準備しています" }
        @{ case = "全体の件数が分からない間は件数を出さない"; name = ""; processed = 0; total = 0; failed = 0; eta = ""; expected = "更新中　準備しています" }
    ) {
        (getIndexingStatusLine $name $processed $total $failed $eta) | Should -Be $expected
    }
}

Describe "getStatusBarView" -Tag Unit {
    It "<case>" -TestCases @(
        @{ case = "更新中は更新中の行が、いつもの文に代わる"; indexing = $true; line = "更新中: 営業部　1 / 2 件"; status = "名前を変えました"; source = "Indexing"; text = "更新中: 営業部　1 / 2 件" }
        @{ case = "更新中で、いつもの文が無くても更新中の行"; indexing = $true; line = "更新中　準備しています"; status = ""; source = "Indexing"; text = "更新中　準備しています" }
        @{ case = "更新中でなければ、直前の操作の結果"; indexing = $false; line = "更新中: 営業部　1 / 2 件"; status = "名前を変えました"; source = "Status"; text = "名前を変えました" }
        @{ case = "出すものが無ければ空"; indexing = $false; line = ""; status = ""; source = "Empty"; text = "" }
    ) {
        param ($case, $indexing, $line, $status, $source, $text)
        $view = getStatusBarView $indexing $line $status
        $view.Source | Should -Be $source
        $view.Text | Should -Be $text
    }
}
