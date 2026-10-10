# 元のファイルを開くときの文言の判断（tebunko\ui\search\open_source_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\search\open_source_view.ps1"
}

Describe "getSourceCheckingStatus / getSourceNotFoundStatus / getSourceUnreachableStatus" -Tag Unit {
    It "確かめている間の文言にパスを入れる" {
        getSourceCheckingStatus "\\server\share\見積\a.xlsx" | Should -Be "元のファイルを確かめています…：\\server\share\見積\a.xlsx（共有フォルダに接続できないときは、しばらくかかります）"
    }

    It "見つからないときの文言にパスを入れる" {
        getSourceNotFoundStatus "C:\data\見積\a.xlsx" | Should -Be "元のファイルが見つかりません：C:\data\見積\a.xlsx"
    }

    It "接続できないときの文言に元のフォルダを入れる" {
        getSourceUnreachableStatus "\\server\share\見積" | Should -Be "元のフォルダに接続できません：\\server\share\見積"
    }
}

Describe "getSourceConnectFailureStatus" -Tag Unit {
    It "接続できないときは、接続できないステータスにする" {
        getSourceConnectFailureStatus "Unreachable" "\\server\share\見積" "" | Should -Be "元のフォルダに接続できません：\\server\share\見積"
    }

    It "その他のときは、例外の文面を入れる" {
        getSourceConnectFailureStatus "Other" "\\server\share\見積" "アクセスが拒否されました。" | Should -Be "元のファイルを確かめられませんでした：アクセスが拒否されました。"
    }
}

Describe "getSourceConnectFailureDialog" -Tag Unit {
    It "接続できないときは、フォルダと接続を確かめる案内を出す" {
        $dialog = getSourceConnectFailureDialog "見積.xlsx" "Unreachable" "\\server\share\見積" ""
        $dialog.Heading | Should -Be "見積.xlsx を開けません"
        $dialog.Title | Should -Be "元のフォルダに接続できません"
        $dialog.Detail | Should -Be "\\server\share\見積"
        $dialog.Hint | Should -Be "ネットワーク・VPN の接続を確かめてから、もう一度開いてください"
    }

    It "その他のときは、例外の文面と権限を確かめる案内を出す" {
        $dialog = getSourceConnectFailureDialog "見積.xlsx" "Other" "\\server\share\見積" "アクセスが拒否されました。"
        $dialog.Heading | Should -Be "見積.xlsx を開けません"
        $dialog.Title | Should -Be "元のファイルを確かめられませんでした"
        $dialog.Detail | Should -Be "アクセスが拒否されました。"
        $dialog.Hint | Should -Be "アクセスの権限・サインインを確かめてください"
    }
}

Describe "getSourceLookingStatus / getSourceLookupFailedStatus" -Tag Unit {
    It "調べている間は、接続できないときに時間がかかることを知らせる" {
        getSourceLookingStatus | Should -Be "パスを調べています…（共有フォルダに接続できないときは、しばらくかかります）"
    }

    It "<name>" -TestCases @(
        @{ name = "調べられなかったときは、例外の文面を添える"; message = "届きません"; expected = "元のファイルの場所を調べられませんでした：届きません" }
        @{ name = "文面が無ければ、知らせだけ"; message = ""; expected = "元のファイルの場所を調べられませんでした" }
    ) {
        getSourceLookupFailedStatus $message | Should -Be $expected
    }
}
