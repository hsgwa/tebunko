# ［インデックス管理］の一覧・選択・ボタンの可否・状態の文言（tebunko\ui\index\index_list_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit_view.ps1"

    function newItem {
        param ([string]$name, [string]$path)
        return [pscustomobject]@{ Name = $name; Path = $path }
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

Describe "getUsedIndexNames（設定から外れて残してあるインデックス）" -Tag Unit {
    BeforeDiscovery {
        $cases = @(
            @{ name = "違うフォルダには付けられない"; folder = "C:\data\別"; query = "売上"; expected = $true }
            @{ name = "前回と同じフォルダには付けられる"; folder = "C:\data\売上"; query = "売上"; expected = $false }
            @{ name = "末尾の区切り・大文字小文字の違いは同じフォルダ"; folder = "c:\DATA\売上\"; query = "売上"; expected = $false }
            @{ name = "大文字・小文字違いの名前も付けられない"; folder = "C:\data\別"; query = "SALES"; expected = $true }
            @{ name = "大文字・小文字違いの名前も、同じフォルダなら付けられる"; folder = "C:\data\売上"; query = "SALES"; expected = $false }
            @{ name = "ドライブの割り当てだけが違うパスは拒む"; folder = "\\srv\share\売上"; query = "売上"; expected = $true }
            @{ name = "フォルダが決まっていないときは付けられない"; folder = ""; query = "売上"; expected = $true }
            @{ name = "置き場にだけあり Path の無い名前はどのフォルダにも付けられない"; folder = "C:\data\孤立"; query = "孤立"; expected = $true }
            @{ name = "残っていない名前は付けられる"; folder = "C:\data\別"; query = "見積"; expected = $false }
        )
    }
    BeforeAll {
        $items = @((newItem "現在" "C:\data\現在"))
        $leftovers = @(
            [pscustomobject]@{ Name = "売上"; Path = "C:\data\売上" },
            [pscustomobject]@{ Name = "Sales"; Path = "C:\data\売上" },
            [pscustomobject]@{ Name = "孤立"; Path = "" }
        )
    }

    It "<name>" -TestCases $cases {
        param ($name, $folder, $query, $expected)
        (getUsedIndexNames $items $null $leftovers $folder).Contains($query) | Should -Be $expected
    }

    It "残りを渡さないときは、一覧の名前だけ（今と同じ）" {
        $used = getUsedIndexNames $items
        @($used).Count | Should -Be 1
        $used.Contains("売上") | Should -Be $false
    }

    It "残っている名前が今の一覧にあれば、重複扱いにしない（一覧の名前として 1 つだけ）" {
        $listed = @((newItem "売上" "C:\data\新しい場所"))
        $used = getUsedIndexNames $listed $null $leftovers "C:\data\新しい場所"
        $used.Contains("売上") | Should -Be $true
        @($used | Where-Object { $_ -ieq "売上" }).Count | Should -Be 1
        (getBlockedIndexNames $listed $leftovers "C:\data\新しい場所").Contains("売上") | Should -Be $false
    }

    It "残っている名前が Keys でも、使えない名前の集合に入る（Hashtable の .Keys が値を返さない）" {
        $left = @(
            [pscustomobject]@{ Name = "Keys"; Path = "C:\data\Keys" }
            [pscustomobject]@{ Name = "旧"; Path = "C:\data\旧" }
        )
        $used = getUsedIndexNames $items $null $left "C:\data\別"
        $used.Contains("Keys") | Should -Be $true
        $used.Contains("旧") | Should -Be $true
        $used.Contains("detached") | Should -Be $false
    }

    It "今の一覧の自分の名前は except に渡せば使える（残りに同じ名前が無いとき）" {
        $used = getUsedIndexNames $items $items[0] $leftovers "C:\data\現在"
        $used.Contains("現在") | Should -Be $false
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
        @{ label = "削除・名前の変更中: すべて無効（エクスポート・インポートを含む）"
           blocker = "削除・名前の変更中"; hasSelection = $true
           new = $false; edit = $false; remove = $false; export = $false; import = $false }
        @{ label = "エクスポート・インポート中: すべて無効（追加・編集・削除を含む）"
           blocker = "エクスポート・インポート中"; hasSelection = $true
           new = $false; edit = $false; remove = $false; export = $false; import = $false }
    ) {
        param ($label, $blocker, $hasSelection, $new, $edit, $remove, $export, $import)
        $result = getIndexTabButtonsEnabled $blocker $hasSelection
        $result.New | Should -Be $new
        $result.Edit | Should -Be $edit
        # 詳細のフォルダパスの［変更…］は、詳細の名前・フォルダパスの欄と同じ決まり
        $result.ChangeFolder | Should -Be $edit
        $result.Remove | Should -Be $remove
        $result.Export | Should -Be $export
        $result.Import | Should -Be $import
    }
}

Describe "getIndexCheckedItems（チェックを付けている行）" -Tag Unit {
    It "Checked の行だけを返す（チェックは保存しないその場の選び。設定の enabled とは別）" {
        $items = @(@{ Name = "営業"; Checked = $true; Enabled = $false }, @{ Name = "経理"; Checked = $false; Enabled = $true }, $null, @{ Name = "人事"; Checked = $true })
        @(getIndexCheckedItems $items | ForEach-Object { $_.Name }) | Should -Be @("営業", "人事")
        @(getIndexCheckedItems @()).Count | Should -Be 0
    }
}

Describe "getIndexStatusHelpText（見出しの説明）" -Tag Unit {
    It "ステータス列の説明は、行のバッジと同じ 6 つの状態を です・ます で書く" {
        $lines = @(getIndexStatusHelpText)
        $lines.Count | Should -Be 7
        $lines[0] | Should -Be "検索の準備の状態です。"
        (($lines | Select-Object -Skip 1) -join "`n") | Should -Match "(?s)最新.*要更新.*更新中.*警告.*エラー.*未作成"
        # 行のバッジの名前（getIndexRowView の Text）と同じ言葉だけを使う
        foreach ($name in "最新", "要更新", "更新中", "警告", "エラー", "未作成") {
            @($lines | Where-Object { $_ -like "・${name}:*" }).Count | Should -Be 1
        }
    }
}

Describe "getIndexSelectionView（全選択と選択中の件数）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "1 件も無い"; total = 0; checked = 0; text = ""; all = $false }
        @{ label = "1 件も選んでいない: 件数は出さない"; total = 4; checked = 0; text = ""; all = $false }
        @{ label = "一部だけ: 横棒"; total = 4; checked = 2; text = "2 / 4 件を選択中"; all = $null }
        @{ label = "すべて"; total = 4; checked = 4; text = "4 / 4 件を選択中"; all = $true }
        @{ label = "千件を超える"; total = 1500; checked = 1200; text = "1,200 / 1,500 件を選択中"; all = $null }
    ) {
        param ($label, $total, $checked, $text, $all)
        $view = getIndexSelectionView $total $checked
        $view.CountText | Should -Be $text
        $view.AllChecked | Should -Be $all
    }
}

Describe "getIndexActionsEnabled（［アクション ▾］のメニューの可否）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "チェックなし: インポートだけ"; blocker = ""; checked = 0; update = $false; export = $false; import = $true; delete = $false }
        @{ label = "1 件チェック: すべて"; blocker = ""; checked = 1; update = $true; export = $true; import = $true; delete = $true }
        @{ label = "2 件チェック: すべて（エクスポート・削除はまとめて行う）"; blocker = ""; checked = 2; update = $true; export = $true; import = $true; delete = $true }
        @{ label = "更新中: すべて無効"; blocker = "インデックス作成中"; checked = 1; update = $false; export = $false; import = $false; delete = $false }
        @{ label = "削除・名前の変更中: すべて無効"; blocker = "削除・名前の変更中"; checked = 2; update = $false; export = $false; import = $false; delete = $false }
    ) {
        param ($label, $blocker, $checked, $update, $export, $import, $delete)
        $result = getIndexActionsEnabled $blocker $checked
        $result.Update | Should -Be $update
        $result.Export | Should -Be $export
        $result.Import | Should -Be $import
        $result.Delete | Should -Be $delete
        $result.ContainsKey("Edit") | Should -BeFalse -Because "編集のダイアログは無い（名前とフォルダパスは詳細の欄で直す）"
    }
}

Describe "getIndexRowMenuEnabled（行の右クリックのメニューの可否）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "ふだんの行: すべて使える"; blocker = ""; hasRow = $true; level = "Ok"; hasPath = $true; update = $true; open = $true; export = $true; delete = $true }
        @{ label = "まだ取り込んでいない行（待ち）: すべて使える"; blocker = ""; hasRow = $true; level = "Wait"; hasPath = $true; update = $true; open = $true; export = $true; delete = $true }
        @{ label = "エラーの行: 更新だけ使えない（行の右端に［更新］が出ない行と同じ）"; blocker = ""; hasRow = $true; level = "Ng"; hasPath = $true; update = $false; open = $true; export = $true; delete = $true }
        @{ label = "警告の行（失敗のあるファイルだけ）: 更新も使える"; blocker = ""; hasRow = $true; level = "Warn"; hasPath = $true; update = $true; open = $true; export = $true; delete = $true }
        @{ label = "元のフォルダの記録が無い行: 元のフォルダを開くだけ使えない"; blocker = ""; hasRow = $true; level = "Ok"; hasPath = $false; update = $true; open = $false; export = $true; delete = $true }
        @{ label = "更新中の行: 元のフォルダを開く以外は使えない"; blocker = "インデックス作成中"; hasRow = $true; level = "Run"; hasPath = $true; update = $false; open = $true; export = $false; delete = $false }
        @{ label = "ほかの行の更新中: 元のフォルダを開く以外は使えない"; blocker = "インデックス作成中"; hasRow = $true; level = "Ok"; hasPath = $true; update = $false; open = $true; export = $false; delete = $false }
        @{ label = "削除・エクスポート・インポート中: 元のフォルダを開く以外は使えない"; blocker = "削除・名前の変更中"; hasRow = $true; level = "Ok"; hasPath = $true; update = $false; open = $true; export = $false; delete = $false }
        @{ label = "行が無い: すべて使えない"; blocker = ""; hasRow = $false; level = "None"; hasPath = $false; update = $false; open = $false; export = $false; delete = $false }
    ) {
        param ($label, $blocker, $hasRow, $level, $hasPath, $update, $open, $export, $delete)
        $result = getIndexRowMenuEnabled $blocker $hasRow $level $hasPath
        $result.Update | Should -Be $update
        $result.OpenFolder | Should -Be $open
        $result.Export | Should -Be $export
        $result.Delete | Should -Be $delete
        $result.ContainsKey("Edit") | Should -BeFalse
    }
}

Describe "getIndexSourceFolderStatus" -Tag Unit {
    It "<kind>" -TestCases @(
        @{ kind = "Empty"; expected = "元のフォルダの場所が分かりません（元のフォルダの記録がありません）" }
        @{ kind = "Checking"; expected = "元のフォルダを確かめています…：C:\共有\営業部" }
        @{ kind = "Missing"; expected = "元のフォルダが見つかりません（移動・削除された可能性があります）：C:\共有\営業部" }
        @{ kind = "Unreachable"; expected = "元のフォルダに接続できません：C:\共有\営業部" }
        @{ kind = "Other"; expected = "元のフォルダを確かめられませんでした：アクセスが拒否されました" }
    ) {
        param ($kind, $expected)
        getIndexSourceFolderStatus $kind "C:\共有\営業部" "アクセスが拒否されました" | Should -Be $expected
    }
}

Describe "getIndexSourceFolderKind" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "フォルダが見つかった"; state = @{ State = "Found"; IsDirectory = $true; Message = "" }; expected = "Open" }
        @{ name = "ファイルだった"; state = @{ State = "Found"; IsDirectory = $false; Message = "" }; expected = "Missing" }
        @{ name = "見つからない"; state = @{ State = "Missing"; IsDirectory = $false; Message = "" }; expected = "Missing" }
        @{ name = "つながらない"; state = @{ State = "Unreachable"; IsDirectory = $false; Message = "" }; expected = "Unreachable" }
        @{ name = "ほかの失敗"; state = @{ State = "Other"; IsDirectory = $false; Message = "拒否" }; expected = "Other" }
    ) {
        param ($state, $expected)
        getIndexSourceFolderKind $state | Should -Be $expected
    }
}

Describe "selectFailedRowsForIndex（失敗の一覧をインデックスごとに）" -Tag Unit {
    BeforeAll {
        $rows = @(
            [pscustomobject]@{ 相対パス = "営業\a.xlsx" }
            [pscustomobject]@{ 相対パス = "営業部\b.xlsx" }
            [pscustomobject]@{ 相対パス = "Sales\sub\c.xlsx" }
            [pscustomobject]@{ 相対パス = "経理\d.xlsx" }
        )
    }

    It "<label>" -TestCases @(
        @{ label = "選んだ名前だけ"; name = "経理"; expected = @("経理\d.xlsx") }
        @{ label = "別の名前の前方一致を拾わない"; name = "営業"; expected = @("営業\a.xlsx") }
        @{ label = "大文字小文字の違いは同じ名前"; name = "sales"; expected = @("Sales\sub\c.xlsx") }
        @{ label = "選んでいない（空）と全部"; name = ""; expected = @("営業\a.xlsx", "営業部\b.xlsx", "Sales\sub\c.xlsx", "経理\d.xlsx") }
        @{ label = "失敗の無いインデックスは 0 件"; name = "人事"; expected = @() }
    ) {
        param ($label, $name, $expected)
        @(selectFailedRowsForIndex $rows $name | ForEach-Object { $_.相対パス }) | Should -Be $expected
    }

    It "行が無くても空を返す" {
        @(selectFailedRowsForIndex @() "営業").Count | Should -Be 0
        @(selectFailedRowsForIndex $null "").Count | Should -Be 0
    }
}

Describe "getIndexRowActions（行の右端のボタン）" -Tag Unit {
    It "<level>・blocker=<blocker>: <action>" -TestCases @(
        @{ level = "Ok"; blocker = ""; action = "Update"; enabled = $true }
        @{ level = "Wait"; blocker = ""; action = "Update"; enabled = $true }
        @{ level = "None"; blocker = ""; action = "Update"; enabled = $true }
        @{ level = "Ok"; blocker = "インデックス作成中"; action = "Update"; enabled = $false }
        @{ level = "None"; blocker = "削除・名前の変更中"; action = "Update"; enabled = $false }
        @{ level = "Run"; blocker = "インデックス作成中"; action = "Stop"; enabled = $false }
        @{ level = "Warn"; blocker = ""; action = "Update"; enabled = $true }
        @{ level = "Warn"; blocker = "インデックス作成中"; action = "Update"; enabled = $false }
        @{ level = "Ng"; blocker = ""; action = "None"; enabled = $true }
    ) {
        param ($level, $blocker, $action, $enabled)
        $result = getIndexRowActions $level $blocker
        $result.Action | Should -Be $action
        $result.UpdateEnabled | Should -Be $enabled
    }

    It "エクスポート中の行は、止められないので［中止］も［更新］も出さない" {
        $result = getIndexRowActions "Run" "" $true
        $result.Action | Should -Be "None"
        $result.UpdateEnabled | Should -BeFalse
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
        @{ label = "Failed のみ 1 以上: 警告（エラーではない）"
           stat = @{ Total = 10; Done = 9; Pending = 0; Failed = 1 }; indexing = $false; enabled = $true
           text = "警告"; sub = ""; level = "Warn" }
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
        @{ ratio = 0.456; sub = "45%"; percent = 45 }
        @{ ratio = 0.0; sub = "0%"; percent = 0 }
        @{ ratio = 1.0; sub = "100%"; percent = 100 }
        @{ ratio = 1.7; sub = "100%"; percent = 100 }
        @{ ratio = -1.0; sub = ""; percent = 0 }
    ) {
        param ($ratio, $sub, $percent)
        $view = getIndexRowView @{ Total = 10; Done = 3; Pending = 7; Failed = 0 } $true $true $ratio
        $view.Text | Should -Be "更新中"
        $view.Text | Should -Not -Match "[0-9]"
        $view.Sub | Should -Be $sub
        $view.Percent | Should -Be $percent
        $view.Level | Should -Be "Run"
    }

    It "enabled が偽の行は、更新中でも更新中にしない（［すべて更新］の対象外）" {
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $true $false 0.5).Text | Should -Be "最新"
    }

    It "今の回に入っていない行（選んだものだけの回の、選ばなかった行）は、更新中にしない" {
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $true $true 0.5 $false).Text | Should -Be "最新"
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $true $true 0.5 $true).Text | Should -Be "更新中"
    }

    It "enabled が偽なら、更新中以外のツールヒントに［すべて更新］の対象外の案内を足す" {
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $false $false).ToolTip | Should -Match "対象から外れています"
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $true $true).ToolTip | Should -Not -Match "対象から外れています"
    }

    It "エクスポート中の行は、状態によらず「エクスポート中」を更新中と同じ Run の色で出し、棒と補足は出さない" {
        foreach ($stat in @($null, @{ Total = 10; Done = 10; Pending = 0; Failed = 0 }, @{ Total = 10; Done = 3; Pending = 7; Failed = 0 })) {
            $view = getIndexRowView $stat $false $true -1.0 $true $true
            $view.Text | Should -Be "エクスポート中"
            $view.Level | Should -Be "Run"
            $view.Sub | Should -Be ""
            $view.Percent | Should -Be 0
            $view.ToolTip | Should -Match "書き出しています"
        }
        (getIndexRowView @{ Total = 10; Done = 10; Pending = 0; Failed = 0 } $false $true -1.0 $true $false).Text | Should -Be "最新"
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

Describe "getIndexLocationView（インデックスの保存先の表示）" -Tag Unit {
    It "パス <name> はそのまま返り、注意は常にあり、警告はネットワークのときだけある（ネットワーク=<net>）" -TestCases @(
        @{ name = "ふつう"; path = "C:\Users\test\Documents\tebunko_ws"; net = $false }
        @{ name = "ふつう"; path = "C:\Users\test\Documents\tebunko_ws"; net = $true }
        @{ name = "UNC"; path = "\\server\share\ws"; net = $true }
        @{ name = "日本語と角括弧の長いパス"; path = "D:\共有\営業部\[2026]\とても長い名前のフォルダ\さらに深い階層\tebunko_ws"; net = $false }
    ) {
        $view = getIndexLocationView $path $net
        $view.Path | Should -BeExactly $path
        $view.Heading | Should -Be "インデックスの保存先"
        $view.Note | Should -BeLike "*本文がそのまま入り*アクセス権は引き継がれません*［設定］*"
        if ($net) {
            $view.Warning | Should -BeLike "注意：保存先は共有フォルダにあります*"
        } else {
            $view.Warning | Should -BeExactly ""
        }
    }

    It "内部の言葉を画面の文言に書かない" {
        $view = getIndexLocationView "C:\ws" $true
        ($view.Note + $view.Warning) | Should -Not -Match "content_index|system_index|TSV"
    }
}

Describe "getIndexLocationOpenStatus（保存先を開くときの文言）" -Tag Unit {
    It "<kind> の文言" -TestCases @(
        @{ kind = "Checking"; message = ""; expected = "保存先のフォルダを確かめています…：C:\ws" }
        @{ kind = "Missing"; message = ""; expected = "保存先のフォルダはまだありません（インデックスを作ると作られます）：C:\ws" }
        @{ kind = "Unreachable"; message = ""; expected = "保存先のフォルダに接続できません：C:\ws" }
        @{ kind = "Other"; message = "アクセスが拒否されました"; expected = "保存先のフォルダを開けません：C:\ws（アクセスが拒否されました）" }
    ) {
        getIndexLocationOpenStatus $kind "C:\ws" $message | Should -Be $expected
    }
}