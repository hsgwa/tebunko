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

Describe "testSourceNeedsConfirm" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "設定に無い名前は確かめる"; name = "受領"; known = $true; confirmed = @("見積"); expected = $true }
        @{ label = "確かめ済みの名前は確かめない"; name = "見積"; known = $true; confirmed = @("見積"); expected = $false }
        @{ label = "大文字・小文字は区別しない"; name = "ABC"; known = $true; confirmed = @("abc"); expected = $false }
        @{ label = "元のフォルダが分からなければ確かめない"; name = "受領"; known = $false; confirmed = @(); expected = $false }
        @{ label = "名前が空なら確かめ済みと言えないので毎回確かめる"; name = ""; known = $true; confirmed = @(); expected = $true }
        @{ label = "名前が空でも元のフォルダが分からなければ確かめない"; name = ""; known = $false; confirmed = @(); expected = $false }
        @{ label = "確かめ済みが無ければ確かめる"; name = "受領"; known = $true; confirmed = @(); expected = $true }
        @{ label = "見えない文字（ソフトハイフン）が混じった名前は別の名前として確かめる"; name = ("見" + [string][char]0xAD + "積"); known = $true; confirmed = @("見積"); expected = $true }
        @{ label = "幅ゼロの文字が混じった名前は別の名前として確かめる"; name = ("見" + [string][char]0x200B + "積"); known = $true; confirmed = @("見積"); expected = $true }
        @{ label = "全角の英字は半角の英字と別の名前として確かめる"; name = "ＡＢＣ"; known = $true; confirmed = @("abc"); expected = $true }
    ) {
        param ($name, $known, $confirmed, $expected)
        testSourceNeedsConfirm $name $known $confirmed | Should -Be $expected
    }
}

Describe "testNameInList / testSourceReceived" -Tag Unit {
    It "testNameInList: <label>" -TestCases @(
        @{ label = "同じ名前"; name = "営業"; list = @("見積", "営業"); expected = $true }
        @{ label = "大文字・小文字は区別しない"; name = "ABC"; list = @("abc"); expected = $true }
        @{ label = "見えない文字が混じれば別の名前"; name = ("営" + [string][char]0xAD + "業"); list = @("営業"); expected = $false }
        @{ label = "一覧が空"; name = "営業"; list = @(); expected = $false }
    ) {
        param ($name, $list, $expected)
        testNameInList $name $list | Should -Be $expected
    }

    It "testSourceReceived: <label>" -TestCases @(
        @{ label = "自分で作った名前はもらったものではない"; name = "営業"; crawled = @("営業"); expected = $false }
        @{ label = "一覧に無い名前はもらったもの"; name = "受取"; crawled = @("営業"); expected = $true }
        @{ label = "見えない文字が混じった名前はもらったもの"; name = ("営" + [string][char]0xAD + "業"); crawled = @("営業"); expected = $true }
        @{ label = "名前が分からないときはもらったもの"; name = ""; crawled = @("営業"); expected = $true }
    ) {
        param ($name, $crawled, $expected)
        testSourceReceived $name $crawled | Should -Be $expected
    }
}

Describe "getSourceConfirmDialog / getSourceConfirmCanceledStatus" -Tag Unit {
    It "名前と元のフォルダを入れる" {
        $dialog = getSourceConfirmDialog "a.xlsx" "受領" "C:\共有\営業部" $false
        $dialog.Heading | Should -BeLike "a.xlsx *"
        $dialog.Title | Should -BeLike "*[[]受領]*"
        $dialog.Detail | Should -Be "C:\共有\営業部"
        $dialog.UseText | Should -Be "このフォルダを使う"
        $dialog.PickText | Should -Be "フォルダを選ぶ"
    }

    It "ローカルのドライブと分からないときだけ、サインイン情報の注意を足す" -TestCases @(
        @{ mayConnect = $true; expected = $true }
        @{ mayConnect = $false; expected = $false }
    ) {
        param ($mayConnect, $expected)
        $dialog = getSourceConfirmDialog "a.xlsx" "受領" "\\server\share" $mayConnect
        ($dialog.Hint -like "*サインイン情報*") | Should -Be $expected
    }

    It "記録できる名前は「次からは聞きません」、記録できない名前は「開くたびに確かめます」と伝える" -TestCases @(
        @{ recordable = $true; ask = "*次からはこのインデックスについて聞きません"; notAsk = "*開くたびに確かめます" }
        @{ recordable = $false; ask = "*開くたびに確かめます"; notAsk = "*次からはこのインデックスについて聞きません" }
    ) {
        param ($recordable, $ask, $notAsk)
        $dialog = getSourceConfirmDialog "a.xlsx" "受領" "C:\x" $false $recordable
        $dialog.Hint | Should -BeLike $ask
        $dialog.Hint | Should -Not -BeLike $notAsk
    }

    It "キャンセルの文言" {
        getSourceConfirmCanceledStatus | Should -Be "開くのをやめました"
    }
}

Describe "getSourceOpenMode / getSourceReadOnlyFailedStatus" -Tag Unit {
    It "もらったインデックスのマクロを持てる形式 <book> は、<mode> でも読み取り専用にする" -TestCases @(
        @{ book = "a.xlsm"; mode = "normal" }
        @{ book = "a.XLSM"; mode = "normal" }
        @{ book = "a.xlsb"; mode = "new" }
        @{ book = "a.xls"; mode = "normal" }
        @{ book = "a.docm"; mode = "normal" }
        @{ book = "a.doc"; mode = "new" }
        @{ book = "a.pptm"; mode = "normal" }
        @{ book = "a.ppt"; mode = "normal" }
        @{ book = "evil.xlsm."; mode = "normal" }
        @{ book = "evil.xlsm "; mode = "normal" }
        @{ book = "evil.docm. ."; mode = "normal" }
    ) {
        param ($book, $mode)
        $result = getSourceOpenMode $book $mode $true
        $result.Mode | Should -Be "readOnly"
        $result.Strict | Should -Be $true
        $result.Notice | Should -Not -BeNullOrEmpty
    }

    It "もう読み取り専用なら知らせは空" {
        $result = getSourceOpenMode "a.xlsm" "readOnly" $true
        $result.Mode | Should -Be "readOnly"
        $result.Strict | Should -Be $true
        $result.Notice | Should -Be ""
    }

    It "<label> はそのまま開く" -TestCases @(
        @{ label = "もらった .xlsx"; book = "a.xlsx"; received = $true }
        @{ label = "もらった .docx"; book = "a.docx"; received = $true }
        @{ label = "もらった .pptx"; book = "a.pptx"; received = $true }
        @{ label = "自分で取り込んだ .xlsm"; book = "a.xlsm"; received = $false }
    ) {
        param ($book, $received)
        $result = getSourceOpenMode $book "normal" $received
        $result.Mode | Should -Be "normal"
        $result.Strict | Should -Be $false
        $result.Notice | Should -Be ""
    }

    It "開けなかったときの文言にパスを入れる" {
        getSourceReadOnlyFailedStatus "C:\x\a.docm" | Should -BeLike "*C:\x\a.docm"
    }
}
