# 知らせの帯（tebunko\ui\shell\notice_banner_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\shell\notice_banner_view.ps1"
}

Describe "getNoticeBannerView" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "結果を保存できなかった: 橙・閉じられる・自動では消えない"; kind = "ResultSaveFailed"; message = "search_results.txt に書き込めません。"; level = "warn"; closable = $true; action = "" }
        @{ label = "まとめての失敗: 橙・閉じられる・自動では消えない"; kind = "IndexOpFailed"; message = "1 件のインデックスを削除できませんでした。"; level = "warn"; closable = $true; action = "" }
        @{ label = "ワークスペースが使えない: 赤・閉じられない・［設定を開く］"; kind = "WorkspaceBlock"; message = "空のフォルダではありません。"; level = "ng"; closable = $false; action = "設定を開く" }
    ) {
        param ($label, $kind, $message, $level, $closable, $action)
        $view = getNoticeBannerView $kind $message
        $view.Show | Should -BeTrue
        $view.Text | Should -Be $message
        $view.Level | Should -Be $level
        $view.Closable | Should -Be $closable
        $view.AutoCloseSeconds | Should -Be 0
        $view.ActionText | Should -Be $action
    }

    It "<label>: 帯を出さない" -TestCases @(
        @{ label = "文が空"; kind = "IndexOpFailed"; message = "" }
        @{ label = "空白だけ"; kind = "ResultSaveFailed"; message = " `n  " }
        @{ label = "知らない種類"; kind = "Unknown"; message = "文" }
    ) {
        param ($label, $kind, $message)
        $view = getNoticeBannerView $kind $message
        $view.Show | Should -BeFalse
        $view.Text | Should -Be ""
    }

    It "複数行の文は 1 行につなぎ、全文をツールヒントに出す" {
        $view = getNoticeBannerView "IndexOpFailed" "2 件を削除できませんでした。`n`n「営業」: 使用中です`n「総務」: 使用中です"
        $view.Text | Should -Be "2 件を削除できませんでした。  「営業」: 使用中です  「総務」: 使用中です"
        $view.ToolTip | Should -Be "2 件を削除できませんでした。`n「営業」: 使用中です`n「総務」: 使用中です"
    }
}
