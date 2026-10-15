# ［インデックス管理］の追加・改名・削除と、その仕事の間の排他・閉じるときの判断（tebunko\ui\index\index_edit_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit_view.ps1"

    function newItem {
        param ([string]$name, [string]$path)
        return [pscustomobject]@{ Name = $name; Path = $path }
    }
}

Describe "getIndexJobBlocker" -Tag Unit {
    It "動いているものの名前を返す（優先順位は インデックス作成 > 削除 > エクスポート・インポート）" -TestCases @(
        @{ isIndexing = $false; indexBusy = $false; archiveBusy = $false; expected = "" }
        @{ isIndexing = $true;  indexBusy = $false; archiveBusy = $false; expected = "インデックス作成中" }
        @{ isIndexing = $false; indexBusy = $true;  archiveBusy = $false; expected = "削除・名前の変更中" }
        @{ isIndexing = $false; indexBusy = $false; archiveBusy = $true;  expected = "エクスポート・インポート中" }
        @{ isIndexing = $true;  indexBusy = $true;  archiveBusy = $true;  expected = "インデックス作成中" }
    ) {
        param ($isIndexing, $indexBusy, $archiveBusy, $expected)
        getIndexJobBlocker $isIndexing $indexBusy $archiveBusy | Should -Be $expected
    }

    It "インデックス名を決めている間（namingBusy）も止める。作成中の次、削除・名前の変更中より先" -TestCases @(
        @{ isIndexing = $false; indexBusy = $false; archiveBusy = $false; expected = "インデックス名の決定中" }
        @{ isIndexing = $false; indexBusy = $true;  archiveBusy = $true;  expected = "インデックス名の決定中" }
        @{ isIndexing = $true;  indexBusy = $false; archiveBusy = $false; expected = "インデックス作成中" }
    ) {
        param ($isIndexing, $indexBusy, $archiveBusy, $expected)
        getIndexJobBlocker $isIndexing $indexBusy $archiveBusy $true | Should -Be $expected
    }

    It "ワークスペースを移している間（workspaceMoving）は、ほかの何が動いていても最優先で止める" -TestCases @(
        @{ isIndexing = $false; indexBusy = $false; archiveBusy = $false; naming = $false }
        @{ isIndexing = $true;  indexBusy = $true;  archiveBusy = $true;  naming = $true }
    ) {
        param ($isIndexing, $indexBusy, $archiveBusy, $naming)
        getIndexJobBlocker $isIndexing $indexBusy $archiveBusy $naming $true | Should -Be "ワークスペースの移動中"
    }
}

Describe "testIndexReadAllowed" -Tag Unit {
    It "ワークスペースを移している間だけ、インデックスを読む操作を止める" -TestCases @(
        @{ moving = $false; expected = $true }
        @{ moving = $true; expected = $false }
    ) {
        param ($moving, $expected)
        testIndexReadAllowed $moving | Should -Be $expected
    }

    It "引数が無ければ読んでよい" {
        testIndexReadAllowed | Should -BeTrue
    }
}

Describe "testMainContentOperable" -Tag Unit {
    It "ワークスペースを移している間だけ、ナビと中身の操作を受けない" -TestCases @(
        @{ moving = $false; expected = $true }
        @{ moving = $true; expected = $false }
    ) {
        param ($moving, $expected)
        testMainContentOperable $moving | Should -Be $expected
    }

    It "引数が無ければ操作を受ける" {
        testMainContentOperable | Should -BeTrue
    }
}

Describe "getIndexStoreSourceFolderError" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "戻り値が空なら空"; output = @(); expected = "" }
        @{ label = "書けたなら空"; output = @(@{ SourceFolderError = "" }); expected = "" }
        @{ label = "書けなかった理由を取り出す"; output = @(@{ SourceFolderError = "届きません" }); expected = "届きません" }
        @{ label = "ほかの出力が混ざっていても取り出す"; output = @("x", @{ SourceFolderError = "届きません" }); expected = "届きません" }
    ) {
        param ($label, $output, $expected)
        getIndexStoreSourceFolderError $output | Should -Be $expected
    }
}

Describe "getIndexJobCloseAction" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "仕事が無ければふつうの閉じ方に進む"; busy = $false; asked = $false; expected = "proceed" }
        @{ label = "仕事が終わっていれば、待った印が残っていても進む"; busy = $false; asked = $true; expected = "proceed" }
        @{ label = "仕事の途中の 1 回目は待つ"; busy = $true; asked = $false; expected = "wait" }
        @{ label = "待っている間にもう一度閉じようとしたら確かめる"; busy = $true; asked = $true; expected = "confirm" }
    ) {
        param ($label, $busy, $asked, $expected)
        getIndexJobCloseAction $busy $asked | Should -Be $expected
    }

    It "ワークスペースを移している間は、ほかの状態に関わらず移し終わるまで待つ" -TestCases @(
        @{ busy = $false; asked = $false }
        @{ busy = $true; asked = $false }
        @{ busy = $true; asked = $true }
    ) {
        param ($busy, $asked)
        getIndexJobCloseAction $busy $asked $true | Should -Be "waitMove"
    }

    It "移し終わるまで閉じない旨の表示を返す" {
        getWorkspaceMoveClosingStatus | Should -Be "ワークスペースを移し終わってから閉じます…"
    }

    It "確認の文言を返す" {
        $confirm = getIndexJobCloseConfirm
        $confirm.Title | Should -Be "変更の途中です"
        $confirm.Heading | Should -BeLike "*閉じますか*"
        $confirm.Hint | Should -BeLike "*食い違う*"
        $confirm.CloseText | Should -Be "待たずに閉じる"
        $confirm.CancelText | Should -Be "終わるまで待つ"
        $confirm.CloseValue | Should -Not -BeNullOrEmpty
    }

    It "更新中に閉じるときの確認の文言を返す" {
        $confirm = getIndexingCloseConfirm
        $confirm.Title | Should -Be "更新中です"
        $confirm.Heading | Should -BeLike "*更新中*中止して閉じますか*"
        $confirm.Hint | Should -BeLike "*続きから再開*"
        $confirm.StopText | Should -Be "中止して閉じる"
        $confirm.StopValue | Should -Be "stop"
        $confirm.CancelText | Should -Be "閉じない"
    }

    It "<label>" -TestCases @(
        @{ label = "［待たずに閉じる］なら、まだ仕事の途中でも閉じる"; answer = "close"; busy = $true; expected = "close" }
        @{ label = "［終わるまで待つ］で、まだ仕事の途中なら閉じない"; answer = ""; busy = $true; expected = "stay" }
        @{ label = "聞いている間に仕事が終わったなら、［終わるまで待つ］でも閉じる"; answer = ""; busy = $false; expected = "close" }
        @{ label = "聞いている間に仕事が終わったなら、［待たずに閉じる］でも閉じる"; answer = "close"; busy = $false; expected = "close" }
    ) {
        param ($label, $answer, $busy, $expected)
        getIndexJobCloseAfterConfirm $answer $busy | Should -Be $expected
    }

    It "複数削除の失敗の件数を知らせる" {
        getBulkDeleteFailedStatus 2 | Should -Be "2 件のインデックスは削除できませんでした"
    }

    It "待つときの文言を返す" {
        $status = getIndexJobClosingStatus
        $status | Should -BeLike "*終わってから閉じます*"
        $status | Should -BeLike "*もう一度閉じる操作*"
    }
}

Describe "getIndexJobBlockedMessage" -Tag Unit {
    It "blocker が空なら空文字列" {
        getIndexJobBlockedMessage "" "エクスポート" | Should -Be ""
    }

    It "<blocker> のとき、<operation> の案内を返す" -TestCases @(
        @{ blocker = "インデックス作成中"; operation = "エクスポート"; expected = "更新中はインデックスをエクスポートできません。更新が終わるまでお待ちください（［中止］で止められます）。" }
        @{ blocker = "インデックス作成中"; operation = "追加"; expected = "更新中はインデックスを追加できません。更新が終わるまでお待ちください（［中止］で止められます）。" }
        @{ blocker = "削除・名前の変更中"; operation = "インポート"; expected = "削除・名前の変更中はインポートできません。終わるまでお待ちください。" }
        @{ blocker = "エクスポート・インポート中"; operation = "削除"; expected = "エクスポート・インポート中は削除できません。終わるまでお待ちください。" }
        @{ blocker = "インデックス作成中"; operation = "ワークスペースの変更"; expected = "更新中はワークスペースを変えられません。更新が終わるまでお待ちください（［中止］で止められます）。" }
        @{ blocker = "削除・名前の変更中"; operation = "ワークスペースの変更"; expected = "前のインデックスの削除・名前の変更が終わるまでお待ちください。" }
        @{ blocker = "エクスポート・インポート中"; operation = "ワークスペースの変更"; expected = "エクスポート・インポートが終わるまでお待ちください。" }
        @{ blocker = "インデックス名の決定中"; operation = "ワークスペースの変更"; expected = "インデックスの一覧を読み込んでいます。終わるまでお待ちください。" }
        @{ blocker = "インデックス名の決定中"; operation = "削除"; expected = "インデックス名の決定中は削除できません。終わるまでお待ちください。" }
        @{ blocker = "ワークスペースの移動中"; operation = "ワークスペースの変更"; expected = "ワークスペースを移している間は、もう一度変えられません。終わるまでお待ちください。" }
        @{ blocker = "ワークスペースの移動中"; operation = "削除"; expected = "ワークスペースの移動中は削除できません。終わるまでお待ちください。" }
    ) {
        param ($blocker, $operation, $expected)
        getIndexJobBlockedMessage $blocker $operation | Should -Be $expected
    }
}

Describe "getIndexDeleteConfirmMessage" -Tag Unit {
    It "<count> 件: <heading>" -TestCases @(
        @{ count = 1; names = @("A"); heading = "「A」のインデックスを削除しますか？"; body = "元のファイルは削除されません。" }
        @{ count = 2; names = @("A", "B"); heading = "選んだ 2 件のインデックスを削除しますか？"; body = "「A」「B」を削除します。元のファイルは削除されません。" }
        @{ count = 3; names = @("A", "B", "C"); heading = "選んだ 3 件のインデックスを削除しますか？"; body = "「A」「B」「C」を削除します。元のファイルは削除されません。" }
        @{ count = 4; names = @("A", "B", "C", "D"); heading = "選んだ 4 件のインデックスを削除しますか？"; body = "「A」「B」「C」ほか 1 件を削除します。元のファイルは削除されません。" }
        @{ count = 10; names = @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J"); heading = "選んだ 10 件のインデックスを削除しますか？"; body = "「A」「B」「C」ほか 7 件を削除します。元のファイルは削除されません。" }
    ) {
        param ($count, $names, $heading, $body)
        $view = getIndexDeleteConfirmMessage $names
        $view.Heading | Should -Be $heading
        $view.Body | Should -Be $body
    }

    It "空の名前は数えない" {
        (getIndexDeleteConfirmMessage @("営業", "", "経理")).Heading | Should -Be "選んだ 2 件のインデックスを削除しますか？"
    }
}

Describe "getIndexBulkResultView" -Tag Unit {
    It "全部成功: ステータスの 1 行だけ" -TestCases @(
        @{ operation = "エクスポート"; expected = "2 件のインデックスをエクスポートしました" }
        @{ operation = "削除"; expected = "2 件のインデックスを削除しました" }
    ) {
        param ($operation, $expected)
        $view = getIndexBulkResultView $operation @(@{ Name = "営業"; Ok = $true; Reason = "" }, @{ Name = "経理"; Ok = $true; Reason = "" })
        $view.Status | Should -Be $expected
        $view.HasFailure | Should -BeFalse
    }

    It "一部失敗: 成功・失敗の件数と、失敗した名前・理由の一覧" {
        $view = getIndexBulkResultView "削除" @(@{ Name = "営業"; Ok = $true; Reason = "" }, @{ Name = "経理"; Ok = $false; Reason = "使用中です" })
        $view.Status | Should -Be "削除: 1 件成功・1 件失敗"
        $view.HasFailure | Should -BeTrue
        $view.FailureHeading | Should -Match "1 件のインデックスを削除できませんでした"
        $view.FailureDetail | Should -Be "「経理」: 使用中です"
    }
}

Describe "getBulkDeleteErrorStatus" -Tag Unit {
    It "理由を添える" {
        getBulkDeleteErrorStatus "届きません" | Should -Be "削除に失敗しました：届きません"
    }
}

Describe "testIndexEditInput" -Tag Unit {
    BeforeAll {
        $items = @((newItem "売上" "C:\data\売上"), (newItem "見積" "C:\data\見積"))
    }

    It "フォルダが空なら、指定するよう伝える" {
        testIndexEditInput "" "新しい名前" $items | Should -Be "元のフォルダを指定してください。"
    }

    It "同じフォルダのインデックスがあれば断る" {
        testIndexEditInput "C:\data\売上" "別名" $items | Should -Match "インデックス \[売上\] が既にあります"
    }

    It "既にあるインデックスの中のフォルダは断る" {
        testIndexEditInput "C:\data\売上\2024" "別名" $items | Should -Match "の中のフォルダです"
    }

    It "既にあるインデックスを含むフォルダは断る" {
        testIndexEditInput "C:\data" "別名" $items | Should -Match "があります"
    }

    It "編集中の行は重複の判定から外す" {
        testIndexEditInput "C:\data\売上" "売上" $items $items[0] | Should -Be ""
    }

    It "名前が重複していれば断る" {
        testIndexEditInput "C:\data\新規" "見積" $items | Should -Be "「見積」は、ほかのインデックスが使っています。別の名前を付けてください。"
    }

    It "問題が無ければ空文字列" {
        testIndexEditInput "C:\data\新規" "新規" $items | Should -Be ""
    }
}
