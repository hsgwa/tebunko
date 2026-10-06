# ［1 インデックス管理］の判断（tebunko\ui\index_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index_view.ps1"

    function newItem {
        param ([string]$name, [string]$path)
        return [pscustomobject]@{ Name = $name; Path = $path }
    }

    function newFastProgress {
        param ([hashtable]$byIndex, [bool]$contentIndexed = $false)
        $dict = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($key in $byIndex.Keys) { $dict[$key] = $byIndex[$key] }
        return @{ Folders = 0; Waiting = 0; ContentIndexed = $contentIndexed; ByIndex = $dict }
    }
}

Describe "getUsedIndexNames" -Tag Unit {
    BeforeAll {
        $items = @((newItem "売上" "C:\data\売上"), (newItem "見積" "C:\data\見積"), (newItem "" "C:\data\新規"))
    }

    It "名前のある行の名前を集める" {
        $used = getUsedIndexNames $items
        @($used).Count | Should -Be 2
        $used.Contains("売上") | Should -Be $true
    }

    It "大文字と小文字を区別しない" {
        $used = getUsedIndexNames @((newItem "Sales" "C:\data\a"))
        $used.Contains("sales") | Should -Be $true
    }

    It "except に渡した行は含めない" {
        $used = getUsedIndexNames $items $items[0]
        $used.Contains("売上") | Should -Be $false
        $used.Contains("見積") | Should -Be $true
    }
}

Describe "getIndexAddedStatus / getFailedFileCheckingStatus / getFailedFileUnreachableStatus / getFailedFileOtherStatus" -Tag Unit {
    It "追加のときの文言は、フォルダの有無によらず同じにする" {
        getIndexAddedStatus "見積" | Should -Be "インデックス [見積] を追加しました。［すべて更新］を押すと中身を更新します"
    }

    It "確かめている間の文言" {
        getFailedFileCheckingStatus "\\server\share\a.xlsx" | Should -Be "元のファイルを確かめています…：\\server\share\a.xlsx"
    }

    It "接続できないときの文言（見つからないときとは別）" {
        getFailedFileUnreachableStatus "\\server\share" | Should -Be "元のフォルダに接続できません：\\server\share"
    }

    It "その他（アクセス拒否など）のときは、例外の文面を出す" {
        getFailedFileOtherStatus "アクセスが拒否されました。" | Should -Be "元のファイルを確かめられませんでした：アクセスが拒否されました。"
    }
}

Describe "getIndexJobBlocker" -Tag Unit {
    It "動いているものの名前を返す（優先順位は インデックス作成 > 削除 > エクスポート・インポート）" -TestCases @(
        @{ isIndexing = $false; indexBusy = $false; archiveBusy = $false; expected = "" }
        @{ isIndexing = $true;  indexBusy = $false; archiveBusy = $false; expected = "インデックス作成中" }
        @{ isIndexing = $false; indexBusy = $true;  archiveBusy = $false; expected = "削除中" }
        @{ isIndexing = $false; indexBusy = $false; archiveBusy = $true;  expected = "エクスポート・インポート中" }
        @{ isIndexing = $true;  indexBusy = $true;  archiveBusy = $true;  expected = "インデックス作成中" }
    ) {
        param ($isIndexing, $indexBusy, $archiveBusy, $expected)
        getIndexJobBlocker $isIndexing $indexBusy $archiveBusy | Should -Be $expected
    }
}

Describe "getIndexJobBlockedMessage" -Tag Unit {
    It "blocker が空なら空文字列" {
        getIndexJobBlockedMessage "" "エクスポート" | Should -Be ""
    }

    It "<blocker> のとき、<operation> の案内を返す" -TestCases @(
        @{ blocker = "インデックス作成中"; operation = "エクスポート"; expected = "更新中はインデックスをエクスポートできません。更新が終わるまでお待ちください（［中止］で止められます）。" }
        @{ blocker = "インデックス作成中"; operation = "追加"; expected = "更新中はインデックスを追加できません。更新が終わるまでお待ちください（［中止］で止められます）。" }
        @{ blocker = "削除中"; operation = "インポート"; expected = "削除中はインポートできません。終わるまでお待ちください。" }
        @{ blocker = "エクスポート・インポート中"; operation = "削除"; expected = "エクスポート・インポート中は削除できません。終わるまでお待ちください。" }
        @{ blocker = "インデックス作成中"; operation = "ワークスペースの変更"; expected = "更新中はワークスペースを変えられません。更新が終わるまでお待ちください（［中止］で止められます）。" }
        @{ blocker = "削除中"; operation = "ワークスペースの変更"; expected = "前のインデックスの削除が終わるまでお待ちください。" }
        @{ blocker = "エクスポート・インポート中"; operation = "ワークスペースの変更"; expected = "エクスポート・インポートが終わるまでお待ちください。" }
    ) {
        param ($blocker, $operation, $expected)
        getIndexJobBlockedMessage $blocker $operation | Should -Be $expected
    }
}

Describe "getIndexTabButtonsEnabled" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "何も動いておらず、選んでいる: すべて有効"
           blocker = ""; hasSelection = $true
           new = $true; edit = $true; remove = $true; export = $true; import = $true }
        @{ label = "何も動いていないが、選んでいない: 追加・インポートだけ有効"
           blocker = ""; hasSelection = $false
           new = $true; edit = $false; remove = $false; export = $false; import = $true }
        @{ label = "インデックス作成中: すべて無効（エクスポート・インポートを含む）"
           blocker = "インデックス作成中"; hasSelection = $true
           new = $false; edit = $false; remove = $false; export = $false; import = $false }
        @{ label = "削除中: すべて無効（エクスポート・インポートを含む）"
           blocker = "削除中"; hasSelection = $true
           new = $false; edit = $false; remove = $false; export = $false; import = $false }
        @{ label = "エクスポート・インポート中: すべて無効（追加・編集・削除を含む）"
           blocker = "エクスポート・インポート中"; hasSelection = $true
           new = $false; edit = $false; remove = $false; export = $false; import = $false }
    ) {
        param ($label, $blocker, $hasSelection, $new, $edit, $remove, $export, $import)
        $result = getIndexTabButtonsEnabled $blocker $hasSelection
        $result.New | Should -Be $new
        $result.Edit | Should -Be $edit
        $result.Remove | Should -Be $remove
        $result.Export | Should -Be $export
        $result.Import | Should -Be $import
    }
}

Describe "getImportResultStatus" -Tag Unit {
    It "件数と、高速検索が次の更新の後に効くことを伝える" {
        $status = getImportResultStatus @{ Name = "営業"; Files = 12; Warnings = @() }
        $status | Should -Be "インデックス [営業] をインポートしました（12 ファイル）。高速検索は次の更新の後に効きます。"
    }

    It "Warnings があれば添える" {
        $status = getImportResultStatus @{ Name = "営業"; Files = 12; Warnings = @("元のフォルダ「C:\共有」は、インデックス [既存] としても登録されています。") }
        $status | Should -Match "既存"
    }
}

Describe "getImportSuggestedName / testImportNameCollision" -Tag Unit {
    # continueImportIndex は getUsedIndexNames の結果を変数に受けて渡す。0・1・2 件以上のどれでも、重なりを見落とさない
    BeforeAll {
        function newUsed([string[]]$names) {
            $items = @($names | ForEach-Object { newItem $_ "C:\data\$_" })
            $used = getUsedIndexNames $items
            return , $used
        }
    }

    It "<label>: 提案する名前は <expected>、確認が要るかは <collision>" -TestCases @(
        @{ label = "0 件"; names = @(); name = "営業"; expected = "営業"; collision = $false }
        @{ label = "1 件で重ならない"; names = @("経理"); name = "営業"; expected = "営業"; collision = $false }
        @{ label = "1 件で重なる"; names = @("営業"); name = "営業"; expected = "営業(2)"; collision = $true }
        @{ label = "2 件以上で重なる"; names = @("経理", "営業", "総務"); name = "営業"; expected = "営業(2)"; collision = $true }
        @{ label = "2 件以上で重ならない"; names = @("経理", "総務"); name = "営業"; expected = "営業"; collision = $false }
        @{ label = "(2) も使われている"; names = @("営業", "営業(2)"); name = "営業"; expected = "営業(3)"; collision = $true }
        @{ label = "大文字・小文字は区別しない"; names = @("Sales", "経理"); name = "sales"; expected = "sales(2)"; collision = $true }
    ) {
        param ($label, $names, $name, $expected, $collision)
        $used = newUsed $names
        getImportSuggestedName $name $used | Should -Be $expected
        testImportNameCollision $name $used | Should -Be $collision
    }

    It "集合が 1 要素の配列に入っていても（@(...) で受けた場合）、同じ答えを返す" {
        $used = @(newUsed @("経理", "営業"))
        testImportNameCollision "営業" $used | Should -BeTrue
        getImportSuggestedName "営業" $used | Should -Be "営業(2)"
    }
}


Describe "testIndexExportInput" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "フォルダが空なら、指定するよう伝える"; folder = ""; expected = "書き出し先のフォルダを指定してください。" }
        @{ label = "空白だけでも同じ"; folder = "  "; expected = "書き出し先のフォルダを指定してください。" }
        @{ label = "フォルダがあれば問題なし"; folder = "C:\共有\書き出し"; expected = "" }
    ) {
        param ($label, $folder, $expected)
        testIndexExportInput $folder | Should -Be $expected
    }
}

Describe "getIndexExportNotice" -Tag Unit {
    It "インデックスの名前を含める" {
        getIndexExportNotice "営業" | Should -Match "\[営業\]"
    }
}

Describe "testIndexImportInput" -Tag Unit {
    It "フォルダが空なら、指定するよう伝える" {
        testIndexImportInput "" "営業" | Should -Be "元のフォルダを指定してください。"
    }

    It "インデックス名として使えなければ理由を返す" {
        testIndexImportInput "C:\共有\営業部" "" | Should -Match "入力してください"
    }

    It "既にある名前でも断らない（上書きの確認に回すため）" {
        testIndexImportInput "C:\共有\営業部" "既存" | Should -Be ""
    }
}

Describe "getIndexImportNotice" -Tag Unit {
    It "ファイル数と大きさを含める" {
        getIndexImportNotice @{ Files = 12; Bytes = 5MB } | Should -Match "12 ファイル・約 5 MB"
    }
}

Describe "getIndexImportOverwriteConfirmMessage" -Tag Unit {
    It "名前を含める" {
        getIndexImportOverwriteConfirmMessage "営業" | Should -Match "「営業」"
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

Describe "getIndexRowView" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "更新中・チェックあり: 更新中（Pending があっても優先）"
           stat = @{ Total = 10; Done = 3; Pending = 7; Failed = 0 }; indexing = $true; enabled = $true
           text = "更新中"; sub = ""; level = "Run" }
        @{ label = "未集計（null）: 未作成"
           stat = $null; indexing = $false; enabled = $true
           text = "未作成"; sub = ""; level = "None" }
        @{ label = "Total が 0: 未作成"
           stat = @{ Total = 0; Done = 0; Pending = 0; Failed = 0 }; indexing = $false; enabled = $true
           text = "未作成"; sub = ""; level = "None" }
        @{ label = "Pending・Failed とも 1 以上: 要更新（Pending が優先。残りの件数を添える）"
           stat = @{ Total = 10; Done = 5; Pending = 2; Failed = 3 }; indexing = $false; enabled = $true
           text = "要更新"; sub = "残り 2 件"; level = "Wait" }
        @{ label = "Failed のみ 1 以上: エラー"
           stat = @{ Total = 10; Done = 9; Pending = 0; Failed = 1 }; indexing = $false; enabled = $true
           text = "エラー"; sub = ""; level = "Ng" }
        @{ label = "それ以外: 最新"
           stat = @{ Total = 10; Done = 10; Pending = 0; Failed = 0 }; indexing = $false; enabled = $true
           text = "最新"; sub = ""; level = "Ok" }
    ) {
        param ($label, $stat, $indexing, $enabled, $text, $sub, $level)
        $view = getIndexRowView $stat $indexing $enabled
        $view.Text | Should -Be $text
        $view.Sub | Should -Be $sub
        $view.Level | Should -Be $level
    }

    It "更新中の行は、進みの割合（0〜1）から「更新中 N%」と棒の長さ（Percent）を出す。割合が分からなければ割合なし" -TestCases @(
        @{ ratio = 0.456; text = "更新中 45%"; percent = 45 }
        @{ ratio = 0.0; text = "更新中 0%"; percent = 0 }
        @{ ratio = 1.0; text = "更新中 100%"; percent = 100 }
        @{ ratio = 1.7; text = "更新中 100%"; percent = 100 }
        @{ ratio = -1.0; text = "更新中"; percent = 0 }
    ) {
        param ($ratio, $text, $percent)
        $view = getIndexRowView @{ Total = 10; Done = 3; Pending = 7; Failed = 0 } $true $true $ratio
        $view.Text | Should -Be $text
        $view.Percent | Should -Be $percent
        $view.Level | Should -Be "Run"
    }

    It "チェックが外れた行は、更新中でも更新中にしない" {
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $true $false 0.5).Text | Should -Be "最新"
    }

    It "enabled が偽なら、作成中以外のツールヒントにチェックの案内を足す（作成中の行には足さない）" {
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $false $false).ToolTip | Should -Match "チェックが外れている"
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $true $true).ToolTip | Should -Not -Match "チェックが外れている"
    }

    It "ツールヒントに前の版・内部の言葉を含めない" {
        $views = @(
            (getIndexRowView $null $false $true)
            (getIndexRowView @{ Total = 10; Done = 3; Pending = 7; Failed = 0 } $true $true)
            (getIndexRowView @{ Total = 10; Done = 5; Pending = 2; Failed = 3 } $false $true)
            (getIndexRowView @{ Total = 10; Done = 9; Pending = 0; Failed = 1 } $false $true)
            (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $false $true)
        )
        foreach ($view in $views) {
            $view.Text | Should -Not -Match "システムインデックス|集約ファイル|本文インデックス|使用不可|インデックス済|不明"
            $view.ToolTip | Should -Not -Match "システムインデックス|集約ファイル|本文インデックス|使用不可|インデックス済|不明"
        }
    }
}

Describe "getIndexDetailRowPlan" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "初めて・空: 畳む"; was = $null; empty = $true; expected = "Hide" }
        @{ name = "初めて・空でない: 出す"; was = $null; empty = $false; expected = "Show" }
        @{ name = "空でないまま（別の行を選んだ）: 動かさない"; was = $false; empty = $false; expected = "Keep" }
        @{ name = "空のまま: 動かさない"; was = $true; empty = $true; expected = "Keep" }
        @{ name = "空でなくなった: 出す"; was = $true; empty = $false; expected = "Show" }
        @{ name = "空になった: 畳む"; was = $false; empty = $true; expected = "Hide" }
    ) {
        getIndexDetailRowPlan $was $empty | Should -Be $expected
    }
}

Describe "getIndexFooterView" -Tag Unit {
    It "フォルダ数とファイルの合計を、3 桁ごとの区切りで出す" -TestCases @(
        @{ folders = 0; files = 0; f = "登録済み: 0 フォルダ"; t = "合計: 0 ファイル" }
        @{ folders = 3; files = 1234; f = "登録済み: 3 フォルダ"; t = "合計: 1,234 ファイル" }
    ) {
        param ($folders, $files, $f, $t)
        $view = getIndexFooterView $folders $files
        $view.Folders | Should -Be $f
        $view.Files | Should -Be $t
    }
}

Describe "getFastSearchRowView" -Tag Unit {
    # progress は newFastProgress（BeforeAll で定義）で、Run のとき（It の中）に byIndex から組み立てる。
    # -TestCases は Discovery の時点で評価されるため、ここでは組み立てずハッシュテーブルのデータだけを置く
    It "<label>" -TestCases @(
        @{ label = "確かめる前: 確認中…"
           reason = $null; byIndex = $null; name = "営業"; hasContent = $true
           text = "確認中…"; level = "None" }
        @{ label = "NoFolder で本文も無い: －"
           reason = "NoFolder"; byIndex = $null; name = "営業"; hasContent = $false
           text = "－"; level = "None" }
        @{ label = "NoFolder だが本文がある（インポートなど）: 不可"
           reason = "NoFolder"; byIndex = $null; name = "営業"; hasContent = $true
           text = "不可"; level = "Ng" }
        @{ label = "NoFolder だが本文がある・インデックス作成中: 不可にせず－（作成中は前の結果のままにしない）"
           reason = "NoFolder"; byIndex = $null; name = "営業"; hasContent = $true; indexing = $true
           text = "－"; level = "None" }
        @{ label = "NoConnection: 不可"
           reason = "NoConnection"; byIndex = $null; name = "営業"; hasContent = $false
           text = "不可"; level = "Ng" }
        @{ label = "NotInScope: 不可"
           reason = "NotInScope"; byIndex = $null; name = "営業"; hasContent = $false
           text = "不可"; level = "Ng" }
        @{ label = "Ok で進み具合を確かめられない: －"
           reason = "Ok"; byIndex = $null; name = "営業"; hasContent = $true
           text = "－"; level = "None" }
        @{ label = "Ok で ByIndex に無く本文がある（インポートなど）: 不可"
           reason = "Ok"; byIndex = @{}; name = "営業"; hasContent = $true
           text = "不可"; level = "Ng" }
        @{ label = "Ok で ByIndex に無く本文がある・インデックス作成中: 不可にせず－"
           reason = "Ok"; byIndex = @{}; name = "営業"; hasContent = $true; indexing = $true
           text = "－"; level = "None" }
        @{ label = "Ok で ByIndex に無く本文も無い: －"
           reason = "Ok"; byIndex = @{}; name = "営業"; hasContent = $false
           text = "－"; level = "None" }
        @{ label = "NotYet: 反映待ち"
           reason = "NotYet"; byIndex = @{ "営業" = @{ Folders = 4; Waiting = 4 } }; name = "営業"; hasContent = $true
           text = "反映待ち"; level = "Wait" }
        @{ label = "Ok で反映済みが 0: 反映待ち（反映中 0% にしない）"
           reason = "Ok"; byIndex = @{ "営業" = @{ Folders = 4; Waiting = 4 } }; name = "営業"; hasContent = $true
           text = "反映待ち"; level = "Wait" }
        @{ label = "Ok で反映待ちが 1 以上: 反映中 N%"
           reason = "Ok"; byIndex = @{ "営業" = @{ Folders = 4; Waiting = 2 } }; name = "営業"; hasContent = $true
           text = "反映中 50%"; level = "Wait" }
        @{ label = "Ok で反映待ちが 0: 可"
           reason = "Ok"; byIndex = @{ "営業" = @{ Folders = 4; Waiting = 0 } }; name = "営業"; hasContent = $true
           text = "可"; level = "Ok" }
    ) {
        param ($label, $reason, $byIndex, $name, $hasContent, $text, $level, $indexing = $false)
        $progress = if ($null -eq $byIndex) { $null } else { newFastProgress $byIndex }
        $view = getFastSearchRowView $reason $progress $name $hasContent $null $indexing
        $view.Text | Should -Be $text
        $view.Level | Should -Be $level
    }

    It "Text は 可・反映中 N%・反映待ち・不可・－・確認中… のどれかだけ" {
        $texts = @(
            (getFastSearchRowView $null $null "営業" $false $null).Text
            (getFastSearchRowView "NoFolder" $null "営業" $false $null).Text
            (getFastSearchRowView "NoConnection" $null "営業" $false $null).Text
            (getFastSearchRowView "NotInScope" $null "営業" $false $null).Text
            (getFastSearchRowView "NotYet" (newFastProgress @{ "営業" = @{ Folders = 2; Waiting = 1 } }) "営業" $true $null).Text
            (getFastSearchRowView "Ok" (newFastProgress @{ "営業" = @{ Folders = 2; Waiting = 1 } }) "営業" $true $null).Text
            (getFastSearchRowView "Ok" (newFastProgress @{ "営業" = @{ Folders = 2; Waiting = 0 } }) "営業" $true $null).Text
        )
        foreach ($text in $texts) {
            $text | Should -BeIn @("可", "反映中 50%", "反映待ち", "不可", "－", "確認中…")
        }
    }

    It "<label>" -TestCases @(
        @{ label = "F=3 W=1 → 66%"; folders = 3; waiting = 1; expected = "反映中 66%" }
        @{ label = "F=200 W=199 → 1%（切り捨てで 0 のとき 1）"; folders = 200; waiting = 199; expected = "反映中 1%" }
        @{ label = "F=1000 W=1 → 99%（100% にしない）"; folders = 1000; waiting = 1; expected = "反映中 99%" }
        @{ label = "F=5 W=5 → 反映待ち（反映中 0% にしない）"; folders = 5; waiting = 5; expected = "反映待ち" }
        @{ label = "F=5 W=0 → 可"; folders = 5; waiting = 0; expected = "可" }
    ) {
        param ($label, $folders, $waiting, $expected)
        $progress = newFastProgress @{ "営業" = @{ Folders = $folders; Waiting = $waiting } }
        (getFastSearchRowView "Ok" $progress "営業" $true $null).Text | Should -Be $expected
    }

    It "NotInScope のツールヒントに「管理者」を含む" {
        (getFastSearchRowView "NotInScope" $null "営業" $false $null).ToolTip | Should -Match "管理者"
    }

    It "ContentIndexed が真のときだけ「自動で外れなかった」を含む" {
        $progressOn = newFastProgress @{ "営業" = @{ Folders = 4; Waiting = 2 } } $true
        $progressOff = newFastProgress @{ "営業" = @{ Folders = 4; Waiting = 2 } } $false
        (getFastSearchRowView "Ok" $progressOn "営業" $true $null).ToolTip | Should -Match "自動で外れなかった"
        (getFastSearchRowView "Ok" $progressOff "営業" $true $null).ToolTip | Should -Not -Match "自動で外れなかった"
    }

    It "checkedAt があれば、確認中…以外のツールヒントに最終確認を含む" {
        $checkedAt = [DateTime]::new(2026, 10, 1, 9, 30, 0)
        (getFastSearchRowView $null $null "営業" $false $checkedAt).ToolTip | Should -Not -Match "最終確認"
        (getFastSearchRowView "NoFolder" $null "営業" $false $checkedAt).ToolTip | Should -Match "最終確認 09:30"
        (getFastSearchRowView "NoConnection" $null "営業" $false $checkedAt).ToolTip | Should -Match "最終確認 09:30"
        (getFastSearchRowView "Ok" $null "営業" $true $checkedAt).ToolTip | Should -Match "最終確認 09:30"
        $progress = newFastProgress @{ "営業" = @{ Folders = 4; Waiting = 0 } }
        (getFastSearchRowView "Ok" $progress "営業" $true $checkedAt).ToolTip | Should -Match "最終確認 09:30"
    }

    It "本文はあるが高速検索用のデータが無い行は、理由の文だけを出す" {
        $progress = newFastProgress @{}
        $toolTip = (getFastSearchRowView "Ok" $progress "営業" $true $null).ToolTip
        $toolTip | Should -Match "このインデックスには高速検索用のデータがありません。"
        # 直し方の案内（［インデックスのオプション］など）や、使っても結果は同じという注記は出さない仕様（理由の文だけを出す）
        $toolTip | Should -Not -Match "インデックスのオプション|system_index|時間だけが違う"
    }

    It "進み具合を確かめられないときのツールヒントに「確かめられなかった」を含む" {
        (getFastSearchRowView "Ok" $null "営業" $true $null).ToolTip | Should -Match "確かめられなかった"
    }

    It "ツールヒントに前の版・内部の言葉を含めない" {
        $progress = newFastProgress @{ "営業" = @{ Folders = 4; Waiting = 2 } } $true
        $views = @(
            (getFastSearchRowView $null $null "営業" $false $null)
            (getFastSearchRowView "NoFolder" $null "営業" $false $null)
            (getFastSearchRowView "NoConnection" $null "営業" $false $null)
            (getFastSearchRowView "NotInScope" $null "営業" $false $null)
            (getFastSearchRowView "NotYet" $progress "営業" $true $null)
            (getFastSearchRowView "Ok" $progress "営業" $true $null)
            (getFastSearchRowView "Ok" (newFastProgress @{ "営業" = @{ Folders = 4; Waiting = 0 } }) "営業" $true $null)
            (getFastSearchRowView "Ok" (newFastProgress @{}) "営業" $true $null)
        )
        foreach ($view in $views) {
            $view.Text | Should -Not -Match "システムインデックス|集約ファイル|本文インデックス|使用不可|インデックス済|不明"
            $view.ToolTip | Should -Not -Match "システムインデックス|集約ファイル|本文インデックス|使用不可|インデックス済|不明|反映中 0%"
        }
    }
}

Describe "testIndexImportInput（元のフォルダの重なり）" -Tag Unit {
    BeforeAll {
        $items = @((newItem "売上" "C:\data\売上"), (newItem "見積" "C:\data\見積"))
    }

    It "<label>: 別のインデックスと重なれば断る（追加・編集と同じ決まり）" -TestCases @(
        @{ label = "同じフォルダ"; folder = "C:\data\売上"; expected = "インデックス \[売上\] が既にあります" }
        @{ label = "中のフォルダ"; folder = "C:\data\売上\2024"; expected = "の中のフォルダです" }
        @{ label = "含むフォルダ"; folder = "C:\data"; expected = "があります" }
    ) {
        param ($label, $folder, $expected)
        testIndexImportInput $folder "新しい名前" $items | Should -Match $expected
    }

    It "同じ名前の行の元のフォルダは、上書きで置き換わるため断らない" {
        testIndexImportInput "C:\data\売上" "売上" $items | Should -Be ""
    }
}

Describe "getIndexDetailView（インデックスの詳細）" -Tag Unit {
    BeforeAll {
        function newDetailItem {
            param ($name = "営業", $enabled = $true, $fileCount = "1,234", $last = "09/30 10:00", $folderStatus = "フォルダがあります",
                $fastLevel = "Ok", $fastTip = "")
            return @{
                Name = $name; Path = "C:\共有\営業部"; Enabled = $enabled; FolderStatus = $folderStatus
                IndexText = "最新"; IndexLevel = "Ok"; IndexSub = ""; FileCountText = $fileCount; LastIngestedText = $last
                FastText = "可"; FastLevel = $fastLevel; FastToolTip = $fastTip; FastCheckedText = "最終確認 10:05"
            }
        }
    }

    It "<label>: 中身を出さず、見出しだけを出す" -TestCases @(
        @{ label = "何も選んでいない（空）"; items = @() }
        @{ label = "何も選んでいない（null）"; items = @($null) }
        @{ label = "複数を選んでいる"; items = @(@{ Name = "営業"; Path = "C:\共有\営業部" }, @{ Name = "経理"; Path = "C:\共有\経理部" }) }
    ) {
        param ($label, $items)
        $view = getIndexDetailView $items $null
        $view.Title | Should -Be "インデックスの状態"
        $view.Selected | Should -BeFalse
        @($view.Rows).Count | Should -Be 0
        $view.Fast.Shown | Should -BeFalse
    }

    It "1 つ選ぶと、題に名前を入れ、左の基本設定と右の情報の表を作る" {
        $view = getIndexDetailView @(newDetailItem) $null
        $view.Title | Should -Be "営業 - 詳細"
        $view.Selected | Should -BeTrue
        $view.Name | Should -Be "営業"
        $view.Path | Should -Be "C:\共有\営業部"
        $view.Badge.Text | Should -Be "最新"
        $view.Badge.Level | Should -Be "Ok"
        $view.Updated | Should -Be "最終更新 09/30 10:00"
        $view.Count | Should -Be "1,234 ファイル"
        ($view.Rows | ForEach-Object { $_.Label }) -join "," | Should -Be "対象ファイル数,最終更新"
        $view.Rows[0].Value | Should -Be "1,234 ファイル"
        $view.Fast.State | Should -Be "可"
        $view.Fast.Checked | Should -Be "最終確認 10:05"
    }

    It "<label>: 出せない値は空にするか、言い換える" -TestCases @(
        @{ label = "まだ更新していない（件数）"; args1 = @{ fileCount = "－"; last = "" }; field = "count"; expected = "まだ更新していません" }
        @{ label = "まだ更新していない（日時）"; args1 = @{ fileCount = "－"; last = "" }; field = "updated"; expected = "" }
    ) {
        param ($label, $args1, $field, $expected)
        $view = getIndexDetailView @(newDetailItem @args1) $null
        $actual = switch ($field) {
            "count" { $view.Count }
            "updated" { $view.Updated }
        }
        $actual | Should -Be $expected
    }

    It "<label>: 高速検索の反映の進み具合" -TestCases @(
        @{ label = "途中"; entry = @{ Folders = 4; Waiting = 1 }; visible = $true; value = 0.75; text = "反映済み 3 / 4 フォルダ"; note = $true }
        @{ label = "すべて反映済み"; entry = @{ Folders = 2; Waiting = 0 }; visible = $true; value = 1.0; text = "反映済み 2 / 2 フォルダ"; note = $false }
        @{ label = "すべて反映待ち"; entry = @{ Folders = 3; Waiting = 3 }; visible = $true; value = 0.0; text = "反映済み 0 / 3 フォルダ"; note = $true }
        @{ label = "待ちが多すぎる値でも 0 未満にしない"; entry = @{ Folders = 3; Waiting = 5 }; visible = $true; value = 0.0; text = "反映済み 0 / 3 フォルダ"; note = $true }
        @{ label = "フォルダが無い"; entry = @{ Folders = 0; Waiting = 0 }; visible = $false; value = 0.0; text = ""; note = $false }
        @{ label = "値が無い"; entry = $null; visible = $false; value = 0.0; text = ""; note = $false }
    ) {
        param ($label, $entry, $visible, $value, $text, $note)
        $view = getIndexDetailView @(newDetailItem) $entry
        $view.Fast.Shown | Should -Be $visible
        $view.Fast.Value | Should -Be $value
        $view.Fast.Text | Should -Be $text
        # 反映待ちのフォルダがあるときだけ、通常の検索で調べる旨の注記を出す
        ($view.Fast.Note -ne "") | Should -Be $note
    }

    It "高速検索が不可のときは、理由（ツールヒントの 1 行目）を出す" {
        $view = getIndexDetailView @(newDetailItem -fastLevel "Ng" -fastTip "Windows Search に接続できない。`n最終確認 10:05") $null
        $view.Fast.Reason | Should -Be "Windows Search に接続できない。"
        (getIndexDetailView @(newDetailItem) $null).Fast.Reason | Should -Be ""
    }

    It "<label>: 行の鍵（RowsKey）" -TestCases @(
        @{ label = "値が変わると変わる"; changed = @{ fileCount = "99" } }
        @{ label = "日時が変わると変わる"; changed = @{ last = "" } }
        @{ label = "フォルダの状態が変わると変わる"; changed = @{ folderStatus = "" } }
        @{ label = "高速検索の状態が変わると変わる"; changed = @{ fastLevel = "Ng"; fastTip = "理由" } }
    ) {
        param ($label, $changed)
        $base = (getIndexDetailView @(newDetailItem) $null).RowsKey
        (getIndexDetailView @(newDetailItem @changed) $null).RowsKey | Should -Not -Be $base
    }

    It "高速検索の進み具合だけが変わっても、行の鍵は同じ" {
        $a = getIndexDetailView @(newDetailItem) @{ Folders = 4; Waiting = 3 }
        $b = getIndexDetailView @(newDetailItem) @{ Folders = 4; Waiting = 1 }
        $a.RowsKey | Should -Be $b.RowsKey
        $a.RowsKey | Should -Not -BeNullOrEmpty
    }

    It "何も選んでいないときの行の鍵は空" {
        (getIndexDetailView @() $null).RowsKey | Should -Be ""
    }
}

Describe "getIndexNavBadge" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "更新中は割合を出す"; indexing = $true; ratio = 0.66; pending = 0; failed = 0; text = "● 66%"; kind = "Run" }
        @{ label = "割合が 100% を超えても 100% で止める"; indexing = $true; ratio = 1.2; pending = 0; failed = 0; text = "● 100%"; kind = "Run" }
        @{ label = "割合が分からないうちは丸だけ"; indexing = $true; ratio = -1.0; pending = 0; failed = 0; text = "●"; kind = "Run" }
        @{ label = "更新中は、残りや失敗より更新中を優先する"; indexing = $true; ratio = 0.0; pending = 3; failed = 2; text = "● 0%"; kind = "Run" }
        @{ label = "中断中"; indexing = $false; ratio = -1.0; pending = 155; failed = 0; text = "中断"; kind = "Warn" }
        @{ label = "失敗があるだけなら警告の印"; indexing = $false; ratio = -1.0; pending = 0; failed = 1; text = "失敗"; kind = "Warn" }
        @{ label = "何も無ければ出さない"; indexing = $false; ratio = -1.0; pending = 0; failed = 0; text = ""; kind = "" }
    ) {
        param ($label, $indexing, $ratio, $pending, $failed, $text, $kind)
        $badge = getIndexNavBadge $indexing $ratio $pending $failed
        $badge.Text | Should -Be $text
        $badge.Kind | Should -Be $kind
    }
}
