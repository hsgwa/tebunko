# ［2 検索］の判断（tebunko\ui\search_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\search_view.ps1"

    # 結果の表の項目のテストに使う、HitRow・FileGroup と同じ項目を持つもの
    function newTestRow {
        param ([int]$order, [string]$line, [int]$lineNumber = 1, [string]$location = "")

        $row = [pscustomobject]@{ Order = $order; Line = $line; LineNumber = $lineNumber; Location = $location }
        $row | Add-Member -MemberType ScriptMethod -Name Contains -Value { param ($text) $this.Line.Contains($text) }
        return $row
    }

    function newTestGroup {
        param ([int]$order, [object[]]$rows, [bool]$expanded = $false)

        $group = [pscustomobject]@{
            Order = $order; IsExpanded = $expanded; ShownCount = $rows.Count
            Rows = New-Object 'System.Collections.Generic.List[object]'
            ShownRows = New-Object 'System.Collections.Generic.List[object]'
        }
        $group.Rows.AddRange($rows)
        $group.ShownRows.AddRange($rows)
        return $group
    }
}

Describe "describeSearchOption" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "既定のままなら空"; option = @{ CaseSensitive = $false; FileFilter = "" }; expected = "" }
        @{ name = "大文字と小文字の区別を出す"; option = @{ CaseSensitive = $true; FileFilter = "" }; expected = "大文字と小文字を区別" }
        @{ name = "対象ファイルを出す"; option = @{ CaseSensitive = $false; FileFilter = "*.xlsx" }; expected = "対象ファイル：*.xlsx" }
        @{ name = "両方あれば中黒でつなぐ"; option = @{ CaseSensitive = $true; FileFilter = "*.xlsx" }; expected = "大文字と小文字を区別・対象ファイル：*.xlsx" }
        @{ name = "図形・コメントを含めるなら出さない"; option = @{ CaseSensitive = $false; FileFilter = ""; IncludeShapes = $true; IncludeComments = $true }; expected = "" }
        @{ name = "図形・コメントを外したら出す"; option = @{ CaseSensitive = $false; FileFilter = ""; IncludeShapes = $false; IncludeComments = $false }; expected = "図形を除く・コメントを除く" }
        @{ name = "コメントだけ外したらコメントだけ出す"; option = @{ CaseSensitive = $false; FileFilter = ""; IncludeComments = $false }; expected = "コメントを除く" }
    ) {
        param ($name, $option, $expected)
        describeSearchOption $option | Should -Be $expected
    }
}

Describe "getFastSearchView" -Tag Unit {
    BeforeAll {
        function script:newFastStatus {
            param ([string]$reason, $progress = $null)
            return @{ Reason = $reason; Progress = $progress }
        }
    }

    It "<name>" -TestCases @(
        @{ name = "まだ確かめていない（null）は使用可"; status = $null; useRegex = $false; word = "見積"; text = "高速検索：使用可"; usable = $true }
        @{ name = "Windows Search が使え、反映待ちが無ければ使用可"; status = @{ Reason = "Ok"; Progress = @{ Folders = 150; Waiting = 0; ContentIndexed = $false } }; useRegex = $false; word = "見積"; text = "高速検索：使用可"; usable = $true }
        @{ name = "進み具合がまだ無ければ使用可"; status = @{ Reason = "Ok" }; useRegex = $false; word = "見積"; text = "高速検索：使用可"; usable = $true }
        @{ name = "反映待ちがあれば、反映の百分率を添える"; status = @{ Reason = "Ok"; Progress = @{ Folders = 150; Waiting = 30; ContentIndexed = $false } }; useRegex = $false; word = "見積"; text = "高速検索：使用可（反映 80%）"; usable = $true }
        @{ name = "百分率は切り捨て（99.3% は 99%）"; status = @{ Reason = "Ok"; Progress = @{ Folders = 150; Waiting = 1; ContentIndexed = $false } }; useRegex = $false; word = "見積"; text = "高速検索：使用可（反映 99%）"; usable = $true }
        @{ name = "正規表現がオン"; status = @{ Reason = "Ok" }; useRegex = $true; word = "見積"; text = "高速検索：使用不可（正規表現）"; usable = $false }
        @{ name = "2 文字以上の部分が無い"; status = @{ Reason = "Ok" }; useRegex = $false; word = "見"; text = "高速検索：使用不可（1 文字）"; usable = $false }
        @{ name = "NoFolder"; status = @{ Reason = "NoFolder" }; useRegex = $false; word = "見積"; text = "高速検索：使用不可（インデックスがありません）"; usable = $false }
        @{ name = "NoConnection"; status = @{ Reason = "NoConnection" }; useRegex = $false; word = "見積"; text = "高速検索：使用不可（Windows Search に接続できません）"; usable = $false }
        @{ name = "NotInScope"; status = @{ Reason = "NotInScope" }; useRegex = $false; word = "見積"; text = "高速検索：使用不可（Windows Search の対象外）"; usable = $false }
        @{ name = "NotYet"; status = @{ Reason = "NotYet"; Progress = @{ Folders = 150; Waiting = 150; ContentIndexed = $false } }; useRegex = $false; word = "見積"; text = "高速検索：使用不可（Windows Search の準備中）"; usable = $false }
        @{ name = "ワードの理由（正規表現）は Windows Search の理由より先に出す"; status = @{ Reason = "NotInScope" }; useRegex = $true; word = "見積"; text = "高速検索：使用不可（正規表現）"; usable = $false }
        @{ name = "ワードの理由（1 文字）は Windows Search の理由より先に出す"; status = @{ Reason = "NoConnection" }; useRegex = $false; word = "見"; text = "高速検索：使用不可（1 文字）"; usable = $false }
        @{ name = "ワードが空のときは、1 文字とは言わず、Windows Search の理由を出す"; status = @{ Reason = "NotInScope" }; useRegex = $false; word = ""; text = "高速検索：使用不可（Windows Search の対象外）"; usable = $false }
        @{ name = "ワードが空でも、Windows Search が使えれば使用可と出す（検索はワードを入れてから）"; status = $null; useRegex = $false; word = ""; text = "高速検索：使用可"; usable = $false }
        @{ name = "知らない理由は、理由なしの使用不可"; status = @{ Reason = "Other" }; useRegex = $false; word = "見積"; text = "高速検索：使用不可"; usable = $false }
    ) {
        param ($name, $status, $useRegex, $word, $text, $usable)
        $view = getFastSearchView $status $useRegex $word
        $view.Text | Should -Be $text
        $view.Usable | Should -Be $usable
    }

    It "確かめている間は、表示だけを「確認中…」にする（使えるかは変えない）" {
        $view = getFastSearchView (newFastStatus "Ok") $false "見積" $true
        $view.Text | Should -Be "高速検索：確認中…"
        $view.Usable | Should -Be $true
        (getFastSearchView (newFastStatus "NotInScope") $false "見積" $true).Usable | Should -Be $false
    }
}

Describe "testFastSearchPreparing" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "まだ確かめていない"; status = $null; expected = $false }
        @{ name = "NotYet は準備中"; status = @{ Reason = "NotYet"; Progress = $null }; expected = $true }
        @{ name = "反映待ちがあれば準備中"; status = @{ Reason = "Ok"; Progress = @{ Folders = 10; Waiting = 1 } }; expected = $true }
        @{ name = "反映待ちが無ければ準備は終わり"; status = @{ Reason = "Ok"; Progress = @{ Folders = 10; Waiting = 0 } }; expected = $false }
        @{ name = "進み具合が無ければ準備中ではない"; status = @{ Reason = "Ok"; Progress = $null }; expected = $false }
        @{ name = "対象外は、待っても変わらないため準備中ではない"; status = @{ Reason = "NotInScope"; Progress = $null }; expected = $false }
    ) {
        param ($name, $status, $expected)
        testFastSearchPreparing $status | Should -Be $expected
    }
}

Describe "getFastSearchDetail" -Tag Unit {
    BeforeAll {
        function script:newFastStatus {
            param ([string]$reason, $progress = $null, $checkedAt = $null)
            return @{ Reason = $reason; Progress = $progress; CheckedAt = $checkedAt }
        }
    }

    It "<reason>: 今の状態と直し方を含む" -TestCases @(
        @{ reason = "NoFolder"; expected = @("インデックス（高速検索用）がありません", "［1 インデックス管理］でインデックスを作成") }
        @{ reason = "NoConnection"; expected = @("接続できません", "Windows Search のサービス（WSearch）") }
        @{ reason = "NotInScope"; expected = @("索引の対象外", "［インデックスのオプション］", "system_index フォルダを索引の対象に加えて", "PC の管理者に頼んでください") }
        @{ reason = "NotYet"; expected = @("まだ索引していません", "索引し終えるのを待って") }
    ) {
        param ($reason, $expected)
        $detail = getFastSearchDetail (newFastStatus $reason) $false "見積"
        $detail.Title | Should -Not -BeNullOrEmpty
        foreach ($text in $expected) {
            $detail.Message | Should -BeLike "*$text*"
        }
        $detail.Message | Should -BeLike "*検索の結果は同じです*"
    }

    It "使えるときは、反映の進み具合と、待てば使えることを書く" {
        $detail = getFastSearchDetail (newFastStatus "Ok" @{ Folders = 150; Waiting = 30; ContentIndexed = $false }) $false "見積"
        $detail.Message | Should -BeLike "*Windows Search：使えます*"
        $detail.Message | Should -BeLike "*反映済み 120 / 150 フォルダ（反映待ち 30）*"
        $detail.Message | Should -BeLike "*待てば使えるようになります*"
    }

    It "反映待ちが無ければ、待つ案内は出さない" {
        $detail = getFastSearchDetail (newFastStatus "Ok" @{ Folders = 150; Waiting = 0; ContentIndexed = $false }) $false "見積"
        $detail.Message | Should -BeLike "*反映済み 150 / 150 フォルダ（反映待ち 0）*"
        $detail.Message | Should -Not -BeLike "*待てば*"
    }

    It "反映が終わっていない間、本文の索引が対象のままのときだけ、対象から外す案内（content_index）を出す" -TestCases @(
        @{ name = "反映待ちがあり、対象のまま"; reason = "Ok"; progress = @{ Folders = 150; Waiting = 30; ContentIndexed = $true }; shown = $true }
        @{ name = "準備中（NotYet）で、対象のまま"; reason = "NotYet"; progress = @{ Folders = 150; Waiting = 150; ContentIndexed = $true }; shown = $true }
        @{ name = "対象から外している"; reason = "Ok"; progress = @{ Folders = 150; Waiting = 30; ContentIndexed = $false }; shown = $false }
        @{ name = "対象のままでも、反映が終わっていれば出さない"; reason = "Ok"; progress = @{ Folders = 150; Waiting = 0; ContentIndexed = $true }; shown = $false }
        @{ name = "進み具合を数えていない（NoFolder）"; reason = "NoFolder"; progress = $null; shown = $false }
    ) {
        param ($name, $reason, $progress, $shown)
        $message = (getFastSearchDetail (newFastStatus $reason $progress) $false "見積").Message
        if ($shown) {
            $message | Should -BeLike "*content_index フォルダを対象から外すと*"
        } else {
            $message | Should -Not -BeLike "*content_index*"
        }
    }

    It "正規表現・1 文字のときは、今の検索ですべてを検索することも書く" {
        (getFastSearchDetail (newFastStatus "Ok") $true "見積").Message | Should -BeLike "*［正規表現を使う］がオン*"
        (getFastSearchDetail (newFastStatus "Ok") $false "見").Message | Should -BeLike "*2 文字以上の部分が無い*"
        (getFastSearchDetail (newFastStatus "Ok") $false "見積").Message | Should -Not -BeLike "*今の検索ではすべてを検索*"
    }

    It "確かめた時刻があれば書く" {
        $detail = getFastSearchDetail (newFastStatus "Ok" $null ([datetime]"2026-09-29 00:10:05")) $false "見積"
        $detail.Message | Should -BeLike "*確かめた時刻：2026/09/29 00:10:05*"
        (getFastSearchDetail (newFastStatus "Ok") $false "見積").Message | Should -Not -BeLike "*確かめた時刻*"
    }

    It "まだ確かめていなくても（null）、こわれずに書く" {
        (getFastSearchDetail $null $false "見積").Message | Should -BeLike "*Windows Search：使えます*"
    }

    It "画面の文言に、内部の言葉を書かない" -TestCases @(
        @{ reason = "Ok"; progress = @{ Folders = 150; Waiting = 30; ContentIndexed = $true } }
        @{ reason = "NoFolder"; progress = $null }
        @{ reason = "NoConnection"; progress = $null }
        @{ reason = "NotInScope"; progress = $null }
        @{ reason = "NotYet"; progress = @{ Folders = 150; Waiting = 150; ContentIndexed = $true } }
    ) {
        param ($reason, $progress)
        $status = newFastStatus $reason $progress ([datetime]"2026-09-29 00:10:05")
        $detail = getFastSearchDetail $status $false "見積"
        $shown = @($detail.Title, $detail.Message, (getFastSearchView $status $false "見積").Text)
        foreach ($word in @("システムインデックス", "集約ファイル", "本文インデックス")) {
            foreach ($text in $shown) {
                $text | Should -Not -BeLike "*$word*"
            }
        }
    }
}

Describe "getSearchProgressText / getSearchSummaryText" -Tag Unit {
    It "検索中は該当件数を出す" {
        getSearchProgressText 1234 | Should -Be "検索中…　該当 1,234 件"
    }

    It "終わったら該当件数・ファイル数・秒数を出す" {
        getSearchSummaryText 1234 5 1.25 | Should -Match "^該当 1,234 件（5 ファイル） ・ 1\.[23] 秒$"
    }
}

Describe "getWordNotice" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "正規表現でなければ出さない"; word = "("; useRegex = $false; expected = "" }
        @{ name = "正規表現として正しければ出さない"; word = "見積.*確定"; useRegex = $true; expected = "" }
        @{ name = "空のワードでは出さない"; word = ""; useRegex = $true; expected = "" }
        @{ name = "正規表現として不正なら、文字どおり検索すると伝える"; word = "("; useRegex = $true; expected = "正規表現として不正なため、文字どおり検索します。" }
    ) {
        param ($name, $word, $useRegex, $expected)
        getWordNotice $word $useRegex | Should -Be $expected
    }
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

Describe "newSearchButtonState" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "検索中は［中止］にする"; searching = $true; stopping = $false; word = "見積"; hasIndex = $true; targetCount = 1; content = "中止"; enabled = $true }
        @{ name = "中止を頼んだ後は押せない"; searching = $true; stopping = $true; word = "見積"; hasIndex = $true; targetCount = 1; content = "中止"; enabled = $false }
        @{ name = "ワード・インデックス・検索対象がそろえば押せる"; searching = $false; stopping = $false; word = "見積"; hasIndex = $true; targetCount = 2; content = "検索"; enabled = $true }
        @{ name = "ワードが空なら押せない"; searching = $false; stopping = $false; word = ""; hasIndex = $true; targetCount = 2; content = "検索"; enabled = $false }
        @{ name = "インデックスが無ければ押せない"; searching = $false; stopping = $false; word = "見積"; hasIndex = $false; targetCount = 2; content = "検索"; enabled = $false }
        @{ name = "検索対象が選ばれていなければ押せない"; searching = $false; stopping = $false; word = "見積"; hasIndex = $true; targetCount = 0; content = "検索"; enabled = $false }
    ) {
        param ($name, $searching, $stopping, $word, $hasIndex, $targetCount, $content, $enabled)
        $state = newSearchButtonState $searching $stopping $word $hasIndex $targetCount
        $state.Content | Should -Be $content
        $state.Enabled | Should -Be $enabled
    }
}

Describe "getAppKind" -Tag Unit {
    It "拡張子からアプリの種類を返す（大文字・小文字は問わない）" {
        getAppKind "見積.xlsx" | Should -Be "Excel"
        getAppKind "古い見積.XLS" | Should -Be "Excel"
        getAppKind "マクロ.xlsm" | Should -Be "Excel"
        getAppKind "報告書.docx" | Should -Be "Word"
        getAppKind "報告書.doc" | Should -Be "Word"
        getAppKind "提案.pptx" | Should -Be "PowerPoint"
    }

    It "Office のファイルでなければ空" {
        getAppKind "メモ.txt" | Should -Be ""
        getAppKind "" | Should -Be ""
    }
}

Describe "describeFileLocations" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "場所が無ければ空"; locations = @(); expected = "" }
        @{ name = "1 か所ならその場所"; locations = @("[シート]4月"); expected = "[シート]4月" }
        @{ name = "2 か所以上なら先頭と、ほかの数"; locations = @("[シート]4月", "[シート]5月", "[シート]6月"); expected = "[シート]4月 ほか 2 か所" }
    ) {
        param ($name, $locations, $expected)
        describeFileLocations $locations | Should -Be $expected
    }
}

Describe "selectShownRows" -Tag Unit {
    It "絞り込みが空ならすべて、あれば合う行だけを元の順で返す" {
        $rows = @((newTestRow 1 "見積 A"), (newTestRow 2 "請求 B"), (newTestRow 3 "見積 C"))
        (selectShownRows $rows "").Count | Should -Be 3
        $shown = selectShownRows $rows "見積"
        $shown.Count | Should -Be 2
        $shown[0].Order | Should -Be 1
        $shown[1].Order | Should -Be 3
    }

    It "合う行が無ければ空の一覧" {
        (selectShownRows @((newTestRow 1 "見積")) "請求").Count | Should -Be 0
    }
}

Describe "getResultItems" -Tag Unit {
    It "閉じているファイルは見出しだけ、開いているファイルは見出しと行" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"), (newTestRow 2 "a2"))
        $b = newTestGroup 2 @((newTestRow 3 "b1")) $true
        $items = getResultItems @($a, $b)
        $items.Count | Should -Be 3
        [object]::ReferenceEquals($items[0], $a) | Should -Be $true
        [object]::ReferenceEquals($items[1], $b) | Should -Be $true
        $items[2].Line | Should -Be "b1"
    }

    It "絞り込みで行が残らないファイルは見出しも出さない" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"))
        $a.ShownRows.Clear()
        $a.ShownCount = 0
        $b = newTestGroup 2 @((newTestRow 2 "b1"))
        $items = getResultItems @($a, $b)
        $items.Count | Should -Be 1
        [object]::ReferenceEquals($items[0], $b) | Should -Be $true
    }
}

Describe "getShownHitRows" -Tag Unit {
    It "閉じているファイルの行も含め、表の順に返す" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"), (newTestRow 2 "a2"))
        $b = newTestGroup 2 @((newTestRow 3 "b1")) $true
        $rows = getShownHitRows @($a, $b)
        @($rows | ForEach-Object { $_.Line }) -join "," | Should -Be "a1,a2,b1"
    }
}

Describe "sortFileGroups" -Tag Unit {
    It "ファイルの中の行を並べ替え、ファイルは先頭の行の順にする" {
        $a = newTestGroup 1 @((newTestRow 1 "a" 5), (newTestRow 2 "a" 9))
        $b = newTestGroup 2 @((newTestRow 3 "b" 7), (newTestRow 4 "b" 1))
        $sorted = sortFileGroups @($a, $b) "LineNumber" $false
        [object]::ReferenceEquals($sorted[0], $b) | Should -Be $true
        @($b.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "1,7"
        @($a.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "5,9"
    }

    It "逆順にもできる" {
        $a = newTestGroup 1 @((newTestRow 1 "a" 5), (newTestRow 2 "a" 9))
        $b = newTestGroup 2 @((newTestRow 3 "b" 7), (newTestRow 4 "b" 1))
        $sorted = sortFileGroups @($a, $b) "LineNumber" $true
        [object]::ReferenceEquals($sorted[0], $a) | Should -Be $true
        @($a.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "9,5"
    }

    It "「場所」（Location）では、同じ場所の中を行番号の順にする（見つかった順と食い違っていても）" {
        $a = newTestGroup 1 @((newTestRow 0 "a" 10 "売上"), (newTestRow 1 "a" 9 "売上"), (newTestRow 2 "a" 1 "仕入"))
        $sorted = sortFileGroups @($a) "Location" $false
        @($a.Rows | ForEach-Object { "$($_.Location)$($_.LineNumber)" }) -join "," | Should -Be "仕入1,売上9,売上10"
        $sorted = sortFileGroups @($a) "Location" $true
        @($a.Rows | ForEach-Object { "$($_.Location)$($_.LineNumber)" }) -join "," | Should -Be "売上10,売上9,仕入1"
    }

    It "同じ値のときは見つかった順" {
        $a = newTestGroup 1 @((newTestRow 2 "x" 1), (newTestRow 1 "x" 1))
        $b = newTestGroup 2 @((newTestRow 3 "x" 1))
        $sorted = sortFileGroups @($b, $a) "LineNumber" $false
        [object]::ReferenceEquals($sorted[0], $a) | Should -Be $true
        @($a.Rows | ForEach-Object { $_.Order }) -join "," | Should -Be "1,2"
    }
}