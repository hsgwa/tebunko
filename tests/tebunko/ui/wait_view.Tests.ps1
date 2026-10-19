# 待つ間のダイアログに出す文言・秒読み・棒の値（tebunko\ui\wait_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\wait_view.ps1"
}

Describe "getWaitDialogView" -Tag Unit {
    It "<kind> の見出し・説明・ボタン" -ForEach @(
        @{ kind = "Stopping"; heading = "中止しています…"; description = "取り込みを止めて、取り込めた分をまとめています。しばらくお待ちください。"; button = "" }
        @{ kind = "CloseWait"; heading = "中止しています…"; description = "取り込めた分をまとめ終わると、画面を閉じます。"; button = "閉じない" }
        @{ kind = "Moving"; heading = "ワークスペースを移しています…"; description = "中身の量によっては時間がかかります。終わるまでお待ちください。"; button = "" }
        @{ kind = "MovingClose"; heading = "ワークスペースを移しています…"; description = "移し終わると、画面を閉じます。"; button = "" }
    ) {
        $view = getWaitDialogView $kind
        $view.Heading | Should -Be $heading
        $view.Description | Should -Be $description
        $view.ButtonText | Should -Be $button
    }

    It "知らない場面は例外にする" {
        { getWaitDialogView "Unknown" } | Should -Throw
    }
}

Describe "getForceCloseCountdownText" -Tag Unit {
    It "残り <left> 秒は「<text>」" -ForEach @(
        @{ left = 60.0; text = "あと 60 秒で強制終了します" }
        @{ left = 14.2; text = "あと 15 秒で強制終了します" }
        @{ left = 0.4; text = "あと 1 秒で強制終了します" }
        @{ left = 0.0; text = "あと 0 秒で強制終了します" }
        @{ left = -3.5; text = "あと 0 秒で強制終了します" }
    ) {
        getForceCloseCountdownText $left | Should -Be $text
    }
}

Describe "getWaitProgressValue" -Tag Unit {
    It "処理 <processed> / 総数 <total> は 不確定=<indeterminate>・値 <value>" -ForEach @(
        @{ processed = 0; total = 0; indeterminate = $true; value = 0.0 }
        @{ processed = 5; total = 0; indeterminate = $true; value = 0.0 }
        @{ processed = 0; total = -1; indeterminate = $true; value = 0.0 }
        @{ processed = 0; total = 10; indeterminate = $false; value = 0.0 }
        @{ processed = 5; total = 10; indeterminate = $false; value = 0.5 }
        @{ processed = 10; total = 10; indeterminate = $false; value = 1.0 }
        @{ processed = 12; total = 10; indeterminate = $false; value = 1.0 }
        @{ processed = -2; total = 10; indeterminate = $false; value = 0.0 }
    ) {
        $bar = getWaitProgressValue $processed $total
        $bar.Indeterminate | Should -Be $indeterminate
        $bar.Value | Should -Be $value
    }
}
