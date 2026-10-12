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
        $dialog = getSourceConfirmDialog "a.xlsx" "受領" "C:\共有\営業部" $false $true
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
        $dialog = getSourceConfirmDialog "a.xlsx" "受領" "\\server\share" $mayConnect $true
        ($dialog.Hint -like "*サインイン情報*") | Should -Be $expected
    }

    It "記録できる名前は「次からは聞きません」、記録できない名前は「開くたびに確かめます」と伝える" -TestCases @(
        @{ recordable = $true; ask = "*次からはこのインデックスについて聞きません"; notAsk = "*開くたびに確かめます" }
        @{ recordable = $false; ask = "*開くたびに確かめます"; notAsk = "*次からはこのインデックスについて聞きません" }
        @{ recordable = $false; ask = "*開くたびに確かめます"; notAsk = "*次からはこのインデックスについて聞きません"; mayConnect = $true }
    ) {
        param ($recordable, $ask, $notAsk, $mayConnect = $false)
        $dialog = getSourceConfirmDialog "a.xlsx" "受領" "C:\x" $mayConnect $recordable
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

Describe "getResultSaveFailedText" -Tag Unit {
    It "<kind> の文" -TestCases @(
        @{ kind = "Locked"; expected = "search_results.txt に書き込めません。開いているアプリを閉じてから、もう一度出力してください。" }
        @{ kind = "Denied"; expected = "search_results.txt に書き込む権限がありません（読み取り専用など）。C:\w\search_results.txt を確かめてから、もう一度出力してください。" }
    ) {
        param ($kind, $expected)
        getResultSaveFailedText $kind "C:\w\search_results.txt" | Should -Be $expected
    }
}

Describe "Office の種類の表（officeOpenKinds）" -Tag Unit {
    It "取り込みの対象の拡張子（officeExtensions）と同じ並びである" {
        @(${officeOpenKinds}.Keys | Sort-Object) | Should -Be @(${officeExtensions} | Sort-Object)
    }
}

Describe "getOfficeOpenKind・testReadOnlySupported・getOpenNewMethod" -Tag Unit {
    # 拡張子ごとの結果（Office の形式だけ読み取り専用で開け、新規は Excel = COM・Word/PowerPoint = 動詞・それ以外 = コピー）
    It "<book> は 種類 <kind>・読み取り専用 <readOnly>・新規 <method>" -TestCases @(
        @{ book = "見積.xlsx"; kind = "Excel"; readOnly = $true; method = "Excel" }
        @{ book = "見積.XLSM"; kind = "Excel"; readOnly = $true; method = "Excel" }
        @{ book = "見積.xls"; kind = "Excel"; readOnly = $true; method = "Excel" }
        @{ book = "見積.xlsb"; kind = "Excel"; readOnly = $true; method = "Excel" }
        @{ book = "報告.docx"; kind = "Office"; readOnly = $true; method = "Verb" }
        @{ book = "報告.DOC"; kind = "Office"; readOnly = $true; method = "Verb" }
        @{ book = "報告.docm"; kind = "Office"; readOnly = $true; method = "Verb" }
        @{ book = "発表.pptx"; kind = "Office"; readOnly = $true; method = "Verb" }
        @{ book = "発表.pptm"; kind = "Office"; readOnly = $true; method = "Verb" }
        @{ book = "発表.ppt"; kind = "Office"; readOnly = $true; method = "Verb" }
        @{ book = "evil.xlsm. "; kind = "Excel"; readOnly = $true; method = "Excel" }
        @{ book = "ひな形.dot"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "ひな形.dotx"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "ひな形.xltx"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "docs.txt"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "議事メモ.txt"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "仕様書.pdf"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "起動.bat"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "拡張子なし"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = ""; kind = ""; readOnly = $false; method = "Copy" }
    ) {
        param ($book, $kind, $readOnly, $method)
        getOfficeOpenKind $book | Should -Be $kind
        testReadOnlySupported $book | Should -Be $readOnly
        getOpenNewMethod $book | Should -Be $method
    }
}

Describe "getOpenCopyFailedStatus" -Tag Unit {
    It "パスを入れ、［開く］を使うよう伝える" {
        getOpenCopyFailedStatus "C:\data\仕様書.pdf" | Should -Be "新規で開くためのコピーを作れませんでした。元のファイルを開くには［開く］を使ってください：C:\data\仕様書.pdf"
    }

    It "理由（<blocker>）があれば、理由を添える" -TestCases @(
        @{ blocker = "TooLong"; reason = "ファイル名かフォルダが長すぎるためです。" }
        @{ blocker = "Place"; reason = "コピーの置き場所を確かめられなかったためです。" }
        @{ blocker = "Failed"; reason = "コピーを書き込めなかったためです。" }
    ) {
        param ($blocker, $reason)
        getOpenCopyFailedStatus "C:\data\a.pdf" $blocker | Should -Be "新規で開くためのコピーを作れませんでした。${reason}元のファイルを開くには［開く］を使ってください：C:\data\a.pdf"
    }
}

Describe "getOpenCopyOpenedStatus" -Tag Unit {
    It "コピーのパスと、元のファイルが変わらないことを出す（<app>）" -TestCases @(
        @{ app = ""; expected = "コピーを新規で開きました（コピー：C:\work\tmp\key\open_copy\abc\a（コピー 20261012-1530）.pdf。保存してもこのコピーに書かれ、元のファイルは変わりません）：C:\data\a.pdf" }
        @{ app = "メモ帳"; expected = "コピーをメモ帳で新規で開きました（コピー：C:\work\tmp\key\open_copy\abc\a（コピー 20261012-1530）.pdf。保存してもこのコピーに書かれ、元のファイルは変わりません）：C:\data\a.pdf" }
    ) {
        param ($app, $expected)
        getOpenCopyOpenedStatus "C:\data\a.pdf" "C:\work\tmp\key\open_copy\abc\a（コピー 20261012-1530）.pdf" $app | Should -Be $expected
    }
}

