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

Describe "getOfficeOpenKind と取り込みの対象の拡張子（officeExtensions）" -Tag Unit {
    It "取り込みの対象の拡張子は、すべて Excel か Office に分かれる" {
        foreach ($extension in ${officeExtensions}) {
            @("Excel", "Office") -contains (getOfficeOpenKind "a$extension") | Should -Be $true
        }
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
        @{ book = "a.xlsxyz"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "a.xlsa"; kind = ""; readOnly = $false; method = "Copy" }
        @{ book = "a.xlsx.bak"; kind = ""; readOnly = $false; method = "Copy" }
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

Describe "表示する名前・パスの制御文字（getSourceBookRejectedStatus ほか）" -Tag Unit {
    BeforeAll {
        $rlo = [string][char]0x202E
        $repl = [string][char]0xFFFD
    }

    It "開かない理由に、ファイル名を入れる" {
        getSourceBookRejectedStatus "evil.hta" | Should -Be "取り込み対象のファイル名ではないため、開きません：evil.hta"
    }

    It "ファイル名・パスの制御文字（RLO・C0）は U+FFFD にして出す" {
        getSourceBookRejectedStatus ("a" + $rlo + "txt.hta") | Should -Be "取り込み対象のファイル名ではないため、開きません：a${repl}txt.hta"
        getSourceNotFoundStatus ("C:\data\x" + [string][char]1 + ".xlsx") | Should -Be "元のファイルが見つかりません：C:\data\x${repl}.xlsx"
    }

    It "確認の画面の見出しと中身の制御文字も U+FFFD にする" {
        $dialog = getSourceConfirmDialog ("a" + $rlo + ".xlsx") "受領" ("C:\共有" + $rlo + "x") $false $true
        $dialog.Heading | Should -BeLike "a${repl}.xlsx *"
        $dialog.Detail | Should -Be "C:\共有${repl}x"
    }
}

Describe "結果の出力の文言" -Tag Unit {
    It "出力に失敗した文（<kind>）" -ForEach @(
        @{ kind = "Locked"; expected = "結果.xlsx に書き込めません。開いているアプリを閉じてから、もう一度出力してください。" }
        @{ kind = "Denied"; expected = "結果.xlsx に書き込む権限がありません（読み取り専用など）。C:\w\結果.xlsx を確かめてから、もう一度出力してください。" }
    ) {
        getResultSaveFailedText $kind "C:\w\結果.xlsx" | Should -Be $expected
    }

    It "出力ボタンの文字: 待ち・書き込み中は「出力を中止」" -ForEach @(
        @{ Busy = $false; Expected = "出力" }
        @{ Busy = $true; Expected = "出力を中止" }
    ) {
        getExportButtonText $Busy | Should -Be $Expected
    }

    It "XAML の ExportButtonText の初期値は、ふだんの文字（getExportButtonText $false）と同じ" {
        $xaml = Get-Content -LiteralPath "${scriptsDir}\tebunko\xaml\search\result_list.xaml" -Raw -Encoding UTF8
        $xaml | Should -Match ([regex]::Escape('x:Name="ExportButtonText" Text="' + (getExportButtonText $false) + '"'))
    }

    It "<Name>" -ForEach @(
        @{ Name = "表がすべて・絞り込みなし"; Params = @("a.xlsx", "All", 1234, $false); Expected = "a.xlsx に 1,234 件を出力しました（出力したフォルダを開きます）" }
        @{ Name = "表がすべて・絞り込み"; Params = @("a.csv", "All", 5, $true); Expected = "絞り込み後の 5 件を a.csv に出力しました" }
        @{ Name = "溢れも含む"; Params = @("a.tsv", "Overflow", 100000, $false); Expected = "表に入らなかった分も含めて、全 100,000 件を a.tsv に出力しました" }
        @{ Name = "溢れも含む・絞り込み"; Params = @("a.tsv", "Overflow", 70, $true); Expected = "表に入らなかった分も含めて、絞り込み後の 70 件を a.tsv に出力しました" }
        @{ Name = "溢れも含む・並べ替え中"; Params = @("a.tsv", "Overflow", 9, $false, "", $true); Expected = "表に入らなかった分も含めて、全 9 件を a.tsv に出力しました（件数が多いため、並べ替えは当てずに検索の順で出力しました）" }
        @{ Name = "先頭だけ（総数の時間切れ）"; Params = @("a.csv", "HeadOnly", 50, $false, "CountTimeout"); Expected = "表に出ている 50 件だけを a.csv に出力しました（総数を数える時間が切れたため、残りは出力していません）" }
        @{ Name = "先頭だけ（溢れを書けなかった）"; Params = @("a.csv", "HeadOnly", 50, $false, "OverflowFailed"); Expected = "表に出ている 50 件だけを a.csv に出力しました（表に入らない分を一時ファイルに書けなかったため、残りは出力していません）" }
        @{ Name = "先頭だけ（検索を止めた）"; Params = @("a.csv", "HeadOnly", 50, $true, "Cancelled"); Expected = "表に出ている 50 件だけを a.csv に出力しました（検索を途中で止めたため、残りは出力していません）" }
        @{ Name = "先頭だけ（理由が不明）"; Params = @("a.csv", "HeadOnly", 1, $false, ""); Expected = "表に出ている 1 件だけを a.csv に出力しました（残りは出力していません）" }
        @{ Name = "Excel の行の上限で書かなかった分"; Params = @("a.xlsx", "All", 1048575, $false, "", $false, 1500, 0); Expected = "a.xlsx に 1,048,575 件を出力しました（出力したフォルダを開きます）（1,500 件は Excel の行の上限を超えたため書いていません。全件は CSV か TSV で出力してください）" }
        @{ Name = "長すぎて切ったセル"; Params = @("a.xlsx", "All", 3, $false, "", $false, 0, 2); Expected = "a.xlsx に 3 件を出力しました（出力したフォルダを開きます）（2 個のセルは長すぎるため、途中で切りました）" }
    ) {
        getExportDoneText @Params | Should -Be $Expected
    }

    It "待ち・書き込み中・中止・失敗・確認などの文言" {
        getExportWaitingText | Should -Be "検索が終わってから出力します…"
        getExportWritingText 12345 | Should -Be "検索結果を出力中…（12,345 件）"
        getExportCancelledText | Should -Be "出力を中止しました。出力先のファイルは変えていません。"
        getExportStoppedText | Should -Be "出力をやめました。"
        getExportExistsPrompt "a.xlsx" | Should -Be "a.xlsx は既にあります。上書きしますか？"
        getExportForbiddenText | Should -BeLike "クロール対象のフォルダ*には出力できません。*"
        getExportFailedText "ディスクが一杯です" | Should -Be "検索結果を出力できませんでした：ディスクが一杯です"
        getExportFailedText "" | Should -Be "検索結果を出力できませんでした"
        getExportFolderOpenFailedText "a.xlsx" 1234 "C:\w\a.xlsx" | Should -Be "a.xlsx に 1,234 件を出力しました。フォルダを開けなかったため、C:\w\a.xlsx を開いてください"
        getSavedFileName "" | Should -Be ""
        getSavedFileName "C:\w\x\" | Should -Be "x"
    }
}

Describe "出力の種類（メニュー・ダイアログ・吹き出し）" -Tag Unit {
    It "メニューは 3 つで、最後に選んだ種類（<last>）に印を付ける。分からなければ Excel" -ForEach @(
        @{ last = "xlsx"; checked = "xlsx" }
        @{ last = "csv"; checked = "csv" }
        @{ last = "tsv"; checked = "tsv" }
        @{ last = ""; checked = "xlsx" }
        @{ last = "txt"; checked = "xlsx" }
    ) {
        $items = @(getExportMenuItems $last)
        $items.Count | Should -Be 3
        ($items | ForEach-Object { $_.Header }) -join "|" | Should -Be "Excel ブック (.xlsx)|CSV (.csv)|TSV (.tsv)"
        @($items | Where-Object { $_.Checked }).Count | Should -Be 1
        ($items | Where-Object { $_.Checked }).Format | Should -Be $checked
    }

    It "種類 <format> の FilterIndex は <index>" -ForEach @(
        @{ format = "xlsx"; index = 1 }
        @{ format = "csv"; index = 2 }
        @{ format = "tsv"; index = 3 }
        @{ format = "x"; index = 1 }
    ) {
        getExportFilterIndex $format | Should -Be $index
        getExportFilter | Should -Be "Excel ブック (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv"
    }

    It "最後の種類は、書けたときだけ替わる（現在 <current>・書いた <written> → <expected>）" -ForEach @(
        @{ current = "xlsx"; written = "csv"; expected = "csv" }
        @{ current = "csv"; written = ""; expected = "csv" }
        @{ current = ""; written = ""; expected = "xlsx" }
        @{ current = "tsv"; written = "x"; expected = "tsv" }
    ) {
        getLastExportFormat $current $written | Should -Be $expected
    }

    It "吹き出し（待ち中 <busy>）" -ForEach @(
        @{ busy = $false; expected = "出力の形式（Excel・CSV・TSV）を選びます" }
        @{ busy = $true; expected = "出力を中止します" }
    ) {
        getExportButtonToolTip $busy | Should -Be $expected
    }
}

Describe "resolveExportTarget・getDefaultExportName・testExportForbiddenPath" -Tag Unit {
    It "<path>（FilterIndex <index>）→ <expected>・付け足し <added>" -ForEach @(
        @{ path = "C:\w\a.xlsx"; index = 2; expected = "C:\w\a.xlsx"; format = "xlsx"; added = $false }
        @{ path = "C:\w\a.CSV"; index = 1; expected = "C:\w\a.CSV"; format = "csv"; added = $false }
        @{ path = "C:\w\a.tsv"; index = 1; expected = "C:\w\a.tsv"; format = "tsv"; added = $false }
        @{ path = "C:\w\a"; index = 1; expected = "C:\w\a.xlsx"; format = "xlsx"; added = $true }
        @{ path = "C:\w\a"; index = 2; expected = "C:\w\a.csv"; format = "csv"; added = $true }
        @{ path = "C:\w\a"; index = 3; expected = "C:\w\a.tsv"; format = "tsv"; added = $true }
        @{ path = "C:\w\a.txt"; index = 2; expected = "C:\w\a.txt.csv"; format = "csv"; added = $true }
        @{ path = "C:\w\a."; index = 1; expected = "C:\w\a.xlsx"; format = "xlsx"; added = $true }
    ) {
        $target = resolveExportTarget $path $index
        $target.Path | Should -Be $expected
        $target.Format | Should -Be $format
        $target.Added | Should -Be $added
    }

    It "空の名前は null" {
        resolveExportTarget "" 1 | Should -Be $null
        resolveExportTarget "  " 1 | Should -Be $null
    }

    It "既定の名前: <word> → <expected>" -ForEach @(
        @{ word = "りんご"; expected = "検索結果_りんご_20261022_103005" }
        @{ word = 'a/b:c*"d'; expected = "検索結果_a_b_c__d_20261022_103005" }
        @{ word = ""; expected = "検索結果_20261022_103005" }
        @{ word = ("あ" * 40); expected = "検索結果_" + ("あ" * 30) + "_20261022_103005" }
        @{ word = ("あ" * 29 + [char]::ConvertFromUtf32(0x20BB7) + "い"); expected = "検索結果_" + ("あ" * 29) + "_20261022_103005" }
    ) {
        getDefaultExportName $word ([datetime]"2026-10-22 10:30:05") | Should -Be $expected
    }

    It "出力先が禁止の場所か: <name>" -ForEach @(
        @{ name = "対象フォルダの中"; path = "C:\共有\営業\a.xlsx"; expected = $true }
        @{ name = "対象フォルダそのもの"; path = "C:\共有\営業"; expected = $true }
        @{ name = "大文字小文字違い・末尾の区切り"; path = "c:\共有\営業\sub\a.csv"; expected = $true }
        @{ name = "名前が似ているだけの別のフォルダ"; path = "C:\共有\営業部\a.xlsx"; expected = $false }
        @{ name = "別の場所"; path = "C:\Users\test\a.xlsx"; expected = $false }
        @{ name = "長いパスの形"; path = "\\?\C:\共有\営業\a.xlsx"; expected = $true }
        @{ name = "UNC の長い形"; path = "\\?\UNC\srv\share\x\a.csv"; expected = $true }
        @{ name = "インデックスのフォルダ"; path = "C:\w\store\content\a.tsv"; expected = $true }
        @{ name = "不正な形"; path = "C:\a|b"; expected = $false }
        @{ name = "空"; path = ""; expected = $false }
    ) {
        testExportForbiddenPath $path @("C:\共有\営業\", "\\srv\share\x", "C:\w\store\content", "", "C:\a|b") | Should -Be $expected
    }
}
