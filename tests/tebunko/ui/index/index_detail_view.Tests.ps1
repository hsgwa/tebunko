# ［インデックス管理］の詳細欄の判断（tebunko\ui\index\index_detail_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\text_ext_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_detail_view.ps1"

    function newItem {
        param ([string]$name, [string]$path)
        return [pscustomobject]@{ Name = $name; Path = $path }
    }
}

Describe "getIndexDetailItem（詳細に出す行・右クリックのメニューなどの対象）" -Tag Unit {
    BeforeAll {
        $script:a = [pscustomobject]@{ Name = "A" }
        $script:b = [pscustomobject]@{ Name = "B" }
    }
    It "<label>" -TestCases @(
        @{ label = "押した行があればそれ（チェックより先）"; sel = "a"; checked = @("b"); expected = "a" }
        @{ label = "押した行が無く、チェックが 1 件だけならその行"; sel = ""; checked = @("b"); expected = "b" }
        @{ label = "押した行もチェックも無ければ無し"; sel = ""; checked = @(); expected = "" }
        @{ label = "押した行が無く、チェックが 2 件なら無し（詳細は件数の表示）"; sel = ""; checked = @("a", "b"); expected = "" }
        @{ label = "押した行があり、チェックが 2 件でも押した行（［アクション ▾］の可否は getIndexActionsEnabled が別に決める）"; sel = "a"; checked = @("a", "b"); expected = "a" }
    ) {
        param ($label, $sel, $checked, $expected)
        $pick = { param ($k) if ($k -eq "") { $null } else { Get-Variable -Scope Script -Name $k -ValueOnly } }
        $item = getIndexDetailItem (& $pick $sel) @($checked | ForEach-Object { & $pick $_ })
        if ($expected -eq "") { $item | Should -BeNullOrEmpty } else { $item.Name | Should -Be (& $pick $expected).Name }
    }
}

Describe "getIndexDetailMultiCount" -Tag Unit {
    It "2 件以上チェックしているときだけ、その件数（0・1 件は 0 で、詳細欄は押した行のまま）" -TestCases @(
        @{ checked = 0; expected = 0 }, @{ checked = 1; expected = 0 }, @{ checked = 2; expected = 2 }, @{ checked = 4; expected = 4 }
    ) {
        param ($checked, $expected)
        getIndexDetailMultiCount $checked | Should -Be $expected
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

Describe "getIndexDetailView（インデックスの詳細）" -Tag Unit {
    BeforeAll {
        function newDetailItem {
            param ($name = "営業", $enabled = $true, $fileCount = "1,234", $last = "09/30 10:00", $folderStatus = "フォルダがあります", $failed = 2, $pending = 3)
            return @{
                Name = $name; Path = "C:\共有\営業部"; Enabled = $enabled; FolderStatus = $folderStatus
                IndexText = "最新"; IndexLevel = "Ok"; IndexSub = ""; FileCountText = $fileCount; LastIngestedText = $last; FailedCount = $failed; PendingCount = $pending
            }
        }
    }

    It "選んだ行の Subfolders（<label>）を返す。何も選んでいない・複数のときは既定の true" -TestCases @(
        @{ label = "false"; value = $false; expected = $false }
        @{ label = "true"; value = $true; expected = $true }
        @{ label = "未設定は含める"; value = $null; expected = $true }
    ) {
        param ($label, $value, $expected)
        $item = newDetailItem
        if ($null -ne $value) { $item.Subfolders = $value }
        (getIndexDetailView @($item)).Subfolders | Should -Be $expected
        (getIndexDetailView @()).Subfolders | Should -BeTrue
        (getIndexDetailView @($item, $item) 2).Subfolders | Should -BeTrue
    }

    It "<label>: 中身を出さず、見出しだけを出す" -TestCases @(
        @{ label = "何も選んでいない（空）"; items = @() }
        @{ label = "何も選んでいない（null）"; items = @($null) }
        @{ label = "複数を選んでいる"; items = @(@{ Name = "営業"; Path = "C:\共有\営業部" }, @{ Name = "経理"; Path = "C:\共有\経理部" }) }
    ) {
        param ($label, $items)
        $view = getIndexDetailView $items
        $view.Title | Should -Be "インデックスの状態"
        $view.Selected | Should -BeFalse
        @($view.Rows).Count | Should -Be 0
    }

    It "元のフォルダが無いとき（FolderMissing）は、バッジを「エラー」にして理由と直し方を出す。更新中は出さない" -TestCases @(
        @{ missing = $true; level = "Ok"; badge = "エラー"; badgeLevel = "Ng"; error = "フォルダが見つかりません。" }
        @{ missing = $false; level = "Ok"; badge = "最新"; badgeLevel = "Ok"; error = "" }
        @{ missing = $true; level = "Run"; badge = "最新"; badgeLevel = "Run"; error = "" }
    ) {
        param ($missing, $level, $badge, $badgeLevel, $error)
        $item = newDetailItem
        $item.FolderMissing = $missing
        $item.IndexLevel = $level
        $view = getIndexDetailView @($item)
        $view.Badge.Text | Should -Be $badge
        $view.Badge.Level | Should -Be $badgeLevel
        $view.Error | Should -Be $error
        if ($error) {
            $view.ErrorHint | Should -Be "フォルダが移動・削除されたか、ネットワークにつながっていない可能性があります。`nフォルダパスの［変更…］でフォルダを選び直してください。"
        } else {
            $view.ErrorHint | Should -Be ""
        }
    }

    It "フォルダの状態の文の色: <name> は <brush>" -TestCases @(
        @{ name = "ある"; missing = $false; checked = $true; brush = "Ok" }
        @{ name = "無い"; missing = $true; checked = $true; brush = "Danger.Text" }
        @{ name = "確認中"; missing = $false; checked = $false; brush = "Ink.Muted" }
        @{ name = "更新中でも無ければ赤"; missing = $true; checked = $true; brush = "Danger.Text"; level = "Run" }
    ) {
        param ($missing, $checked, $brush, $level = "Ok")
        $item = newDetailItem
        $item.FolderMissing = $missing
        $item.FolderChecked = $checked
        $item.IndexLevel = $level
        (getIndexDetailView @($item)).FolderBrush | Should -Be $brush
    }

    It "フォルダの状態のアイコン（FolderIcon）: <name> は <icon>" -TestCases @(
        @{ name = "ある"; missing = $false; checked = $true; status = "フォルダがあります"; icon = "Check" }
        @{ name = "無い"; missing = $true; checked = $true; status = "フォルダが見つかりません"; icon = "Close" }
        @{ name = "確認中"; missing = $false; checked = $false; status = "… フォルダを確認しています"; icon = "" }
        @{ name = "文が空（出す文が無い）"; missing = $false; checked = $true; status = ""; icon = "" }
        @{ name = "更新中でも無ければ Close"; missing = $true; checked = $true; status = "フォルダが見つかりません"; icon = "Close"; level = "Run" }
    ) {
        param ($missing, $checked, $status, $icon, $level = "Ok")
        $item = newDetailItem -folderStatus $status
        $item.FolderMissing = $missing
        $item.FolderChecked = $checked
        $item.IndexLevel = $level
        (getIndexDetailView @($item)).FolderIcon | Should -Be $icon
    }

    It "何も選んでいないとき・複数を選んでいるときは、フォルダの状態のアイコンを出さない。調べ終えたかを省くと調べ終えたものとして Check にする" {
        (getIndexDetailView @()).FolderIcon | Should -Be ""
        (getIndexDetailView @() 2).FolderIcon | Should -Be ""
        (getIndexDetailView @(newDetailItem)).FolderIcon | Should -Be "Check"
    }

    It "何も選んでいないときのフォルダの状態の色は灰、調べ終えたかを省くと調べ終えたものとして扱う" {
        (getIndexDetailView @()).FolderBrush | Should -Be "Ink.Muted"
        (getIndexDetailView @(newDetailItem)).FolderBrush | Should -Be "Ok"
    }

    It "2 件以上を選んでいる（multiCount）と、題に件数を出し、詳細の代わりに案内を出す" {
        $view = getIndexDetailView @() 2
        $view.Title | Should -Be "2 件を選択中"
        $view.Multi | Should -BeTrue
        $view.Selected | Should -BeFalse
        $view.Hint | Should -Be "1 件だけ選ぶと、ここに詳細を表示します。"
        (getIndexDetailView @(newDetailItem) 1).Multi | Should -BeFalse
        (getIndexDetailView @(newDetailItem) 0).Hint | Should -Be ""
    }

    It "更新中の行は、棒の長さ・件数と残り時間・読み取り中のファイルを返す" {
        $item = newDetailItem
        $item.IndexLevel = "Run"; $item.IndexText = "更新中"
        $running = @{ Ratio = 0.45; Processed = 205; Total = 455; Failed = 1; Eta = "残り約 3 分"; Current = "見積もり\A社\A社_見積書_改訂.xlsx" }
        $view = getIndexDetailView @($item) 0 $running
        $view.Run.Shown | Should -BeTrue
        $view.Run.Value | Should -Be 0.45
        $view.Run.CountText | Should -Be "205 / 455 件（失敗 1 件）・残り約 3 分"
        $view.Run.FileText | Should -Be "読み取り中のファイル：見積もり\A社\A社_見積書_改訂.xlsx"
    }

    It "更新中の行の件数は、無いものを出さない（失敗 0 件・残り時間なし・まだ件数が分からない）" -TestCases @(
        @{ running = @{ Ratio = 0.5; Processed = 10; Total = 20; Failed = 0; Eta = ""; Current = "" }; count = "10 / 20 件"; file = "" }
        @{ running = @{ Ratio = 1.5; Processed = 0; Total = 0; Failed = 0; Eta = "残り 1 分未満"; Current = "a.xlsx" }; count = "残り 1 分未満"; file = "読み取り中のファイル：a.xlsx" }
        @{ running = $null; count = ""; file = "" }
    ) {
        param ($running, $count, $file)
        $item = newDetailItem
        $item.IndexLevel = "Run"
        $view = getIndexDetailView @($item) 0 $running
        $view.Run.Shown | Should -BeTrue
        $view.Run.CountText | Should -Be $count
        $view.Run.FileText | Should -Be $file
        $view.Run.Value | Should -BeLessOrEqual 1.0
    }

    It "更新中でない行には、更新中の表示を出さない" {
        (getIndexDetailView @(newDetailItem) 0 @{ Ratio = 0.5; Processed = 1; Total = 2; Failed = 0; Eta = ""; Current = "" }).Run.Shown | Should -BeFalse
    }

    It "1 つ選ぶと、題に名前を入れ、左の基本設定と右の情報の表を作る" {
        $view = getIndexDetailView @(newDetailItem)
        $view.Title | Should -Be "営業 - 詳細"
        $view.Selected | Should -BeTrue
        $view.Name | Should -Be "営業"
        $view.Path | Should -Be "C:\共有\営業部"
        $view.Badge.Text | Should -Be "最新"
        $view.Badge.Level | Should -Be "Ok"
        $view.Updated | Should -Be "最終更新 09/30 10:00"
        $view.Count | Should -Be "1,234 ファイル"
        ($view.Rows | ForEach-Object { $_.Label }) -join "," | Should -Be "対象ファイル数,最終更新,うち失敗,うち未取り込み"
        $view.Rows[0].Value | Should -Be "1,234 ファイル"
        $view.Rows[2].Value | Should -Be "2 件"
        $view.Rows[3].Value | Should -Be "3 件"
    }

    It "<label>: 出せない値は空にするか、言い換える" -TestCases @(
        @{ label = "まだ更新していない（件数）"; args1 = @{ fileCount = "－"; last = "" }; field = "count"; expected = "まだ更新していません" }
        @{ label = "まだ更新していない（日時）"; args1 = @{ fileCount = "－"; last = "" }; field = "updated"; expected = "" }
    ) {
        param ($label, $args1, $field, $expected)
        $view = getIndexDetailView @(newDetailItem @args1)
        $actual = switch ($field) {
            "count" { $view.Count }
            "updated" { $view.Updated }
        }
        $actual | Should -Be $expected
    }

    It "<label>: 行の鍵（RowsKey）" -TestCases @(
        @{ label = "値が変わると変わる"; changed = @{ fileCount = "99" } }
        @{ label = "日時が変わると変わる"; changed = @{ last = "" } }
        @{ label = "フォルダの状態が変わると変わる"; changed = @{ folderStatus = "" } }
        @{ label = "失敗の件数が変わると変わる"; changed = @{ failed = 5 } }
        @{ label = "未取り込みの件数が変わると変わる"; changed = @{ pending = 9 } }
    ) {
        param ($label, $changed)
        $base = (getIndexDetailView @(newDetailItem)).RowsKey
        (getIndexDetailView @(newDetailItem @changed)).RowsKey | Should -Not -Be $base
    }

    It "何も選んでいないときの行の鍵は空" {
        (getIndexDetailView @()).RowsKey | Should -Be ""
    }
}

Describe "getIndexDetailView（取り込む拡張子の行）" -Tag Unit {
    It "何も選んでいなければ空で、［変更…］は押せない" {
        $view = getIndexDetailView @() 0 $null
        $view.TextExt.Summary | Should -Be ""
        $view.TextExt.Stale | Should -BeFalse
        $view.TextExt.ChangeEnabled | Should -BeFalse
    }

    It "<name>" -TestCases @(
        @{ name = "既定（全部取り込む・控え無し）"; excludes = @(); ingested = @(); entry = $null; editable = $true
           summary = "更新すると、件数が出ます"; caption = ""; stale = $false; enabled = $true; source = "none" }
        @{ name = "前回の控えあり"; excludes = @(); ingested = @(); entry = @{ Time = "2024/10/14 15:30"; Counts = @{ ".txt" = 3 } }; editable = $true
           summary = "すべて取り込み（3 件）"; caption = "前回の更新（2024/10/14 15:30）の件数です"; stale = $false; enabled = $true; source = "previous" }
        @{ name = "外して、前回は全部取り込んだ（キー無し）→ 差あり"; excludes = @(".log"); ingested = $null; entry = $null; editable = $true
           summary = "更新すると、件数が出ます"; caption = ""; stale = $true; enabled = $true; source = "none" }
        @{ name = "前回と同じ外し方なら差なし"; excludes = @(".log"); ingested = @(".LOG"); entry = $null; editable = $true
           summary = "更新すると、件数が出ます"; caption = ""; stale = $false; enabled = $true; source = "none" }
        @{ name = "更新中は［変更…］を押せない"; excludes = @(); ingested = @(); entry = $null; editable = $false
           summary = "更新すると、件数が出ます"; caption = ""; stale = $false; enabled = $false; source = "none" }
    ) {
        param ($name, $excludes, $ingested, $entry, $editable, $summary, $caption, $stale, $enabled, $source)
        $item = @{ Name = "営業"; Path = "C:\共有\営業部"; TextExcludes = $excludes; TextExcludesIngested = $ingested; TextCountsEntry = $entry; TextExtEditable = $editable }
        $view = getIndexDetailView @($item) 0 $null
        $view.TextExt.Summary | Should -Be $summary
        $view.TextExt.Caption | Should -Be $caption
        $view.TextExt.Stale | Should -Be $stale
        $view.TextExt.ChangeEnabled | Should -Be $enabled
        $view.TextExt.CountsSource | Should -Be $source
    }
}

Describe "getIndexDetailEditDecision（詳細の名前・フォルダパスの欄の確定）" -Tag Unit {
    BeforeAll {
        $script:items = @((newItem "売上" "C:\data\売上"), (newItem "見積" "C:\data\見積"))
        $script:item = $script:items[0]
    }

    It "<label>" -TestCases @(
        @{ label = "変わっていなければ何もしない"; name = "売上"; path = "C:\data\売上"; action = "None"; newName = "売上"; newPath = "C:\data\売上"; message = "" }
        @{ label = "前後の空白と末尾の \ だけの違いは変わっていない扱い"; name = " 売上 "; path = "C:\data\売上\"; action = "None"; newName = "売上"; newPath = "C:\data\売上"; message = "" }
        @{ label = "名前を変えると反映する"; name = "資料"; path = "C:\data\売上"; action = "Apply"; newName = "資料"; newPath = "C:\data\売上"; message = "" }
        @{ label = "大文字・小文字だけの名前の違いも反映する"; name = "ABC"; path = "C:\data\売上"; action = "Apply"; newName = "ABC"; newPath = "C:\data\売上"; message = ""; item = "abc" }
        @{ label = "フォルダを変えると反映する"; name = "売上"; path = "C:\data\新規"; action = "Apply"; newName = "売上"; newPath = "C:\data\新規"; message = "" }
        @{ label = "名前が空なら断る"; name = "  "; path = "C:\data\売上"; action = "Invalid"; newName = "売上"; newPath = "C:\data\売上"; message = "?*" }
        @{ label = "使えない文字の名前は断る"; name = "a\b"; path = "C:\data\売上"; action = "Invalid"; newName = "売上"; newPath = "C:\data\売上"; message = "?*" }
        @{ label = "ほかのインデックスと同じ名前は断る"; name = "見積"; path = "C:\data\売上"; action = "Invalid"; newName = "売上"; newPath = "C:\data\売上"; message = "「見積」は、ほかのインデックスが使っています。*" }
        @{ label = "フォルダが空なら断る"; name = "売上"; path = ""; action = "Invalid"; newName = "売上"; newPath = "C:\data\売上"; message = "元のフォルダを指定してください。" }
        @{ label = "ほかのインデックスのフォルダの中は断る"; name = "売上"; path = "C:\data\見積\2024"; action = "Invalid"; newName = "売上"; newPath = "C:\data\売上"; message = "*の中のフォルダです*" }
    ) {
        param ($label, $name, $path, $action, $newName, $newPath, $message, $item = $null)
        $target = $script:item
        $list = $script:items
        if ($item) {
            $target = newItem $item "C:\data\売上"
            $list = @($target, $script:items[1])
        }
        $result = getIndexDetailEditDecision $name $path $target $list
        $result.Action | Should -Be $action
        $result.Name | Should -Be $newName
        $result.Path | Should -Be $newPath
        $result.Message | Should -BeLike $message
    }

    It "<label>" -TestCases @(
        @{ label = "売上 → 売上2 に直したあと、同じフォルダの「売上」に戻せる"; folder = "C:\data\売上"; action = "Apply"; message = "" }
        @{ label = "別のフォルダの「売上」には付けられない（設定から外したインデックスが残っている）"; folder = "C:\data\別"; action = "Invalid"; message = "*設定から外したインデックスが残っているため*" }
    ) {
        param ($label, $folder, $action, $message)
        $renamed = newItem "売上2" "C:\data\売上"
        $list = @($renamed, (newItem "見積" "C:\data\見積"))
        $leftovers = @([pscustomobject]@{ Name = "売上"; Path = "C:\data\売上" })
        $target = $renamed
        if ($folder -ne "C:\data\売上") {
            $target = newItem "売上2" $folder
            $list = @($target, $list[1])
        }
        $result = getIndexDetailEditDecision "売上" $folder $target $list $leftovers
        $result.Action | Should -Be $action
        $result.Message | Should -BeLike $message
    }
}

Describe "getIndexFileStateLabels / getIndexFilesHeading（取り込んだファイルの一覧）" -Tag Unit {
    It "状態の文言は 3 つの種類にそろっている" {
        $labels = getIndexFileStateLabels
        $labels.Done | Should -Be "取り込み済"
        $labels.Failed | Should -Be "失敗"
        $labels.Pending | Should -Be "未取り込み"
    }

    It "<label>" -TestCases @(
        @{ label = "0 件"; total = 0; failed = 0; pending = 0; expected = "まだ取り込んだファイルはありません" }
        @{ label = "失敗・未取り込みなし"; total = 1234; failed = 0; pending = 0; expected = "ファイル 1,234 件" }
        @{ label = "失敗あり"; total = 1234; failed = 5; pending = 0; expected = "ファイル 1,234 件（うち失敗 5 件）" }
        @{ label = "未取り込みあり"; total = 1234; failed = 0; pending = 7; expected = "ファイル 1,234 件（うち未取り込み 7 件）" }
        @{ label = "失敗も未取り込みもあり"; total = 1234; failed = 5; pending = 7; expected = "ファイル 1,234 件（うち失敗 5 件・未取り込み 7 件）" }
    ) {
        param ($label, $total, $failed, $pending, $expected)
        getIndexFilesHeading $total $failed $pending | Should -Be $expected
    }
}

Describe "getIndexDetailInfoRows（インデックス情報の行）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "取り込んだ件数があれば、失敗・未取り込みの行も出す"; has = $true; labels = "対象ファイル数,最終更新,うち失敗,うち未取り込み" }
        @{ label = "まだ取り込んでいなければ、2 行だけ"; has = $false; labels = "対象ファイル数,最終更新" }
    ) {
        param ($label, $has, $labels)
        $item = @{ FileCountText = "10"; LastIngestedText = ""; FailedCount = 1; PendingCount = 0 }
        $rows = getIndexDetailInfoRows $item $has
        ($rows | ForEach-Object { $_.Label }) -join "," | Should -Be $labels
        $rows[1].Value | Should -Be "－"
    }
}

Describe "selectIndexFileRows / getIndexFailedEntryView（失敗の絞り込みと入口）" -Tag Unit {
    BeforeAll {
        $script:rows = @(
            [pscustomobject]@{ RelPath = "a.docx"; Kind = "Failed" }
            [pscustomobject]@{ RelPath = "b.xlsx"; Kind = "Done" }
            [pscustomobject]@{ RelPath = "c.pptx"; Kind = "Pending" }
            [pscustomobject]@{ RelPath = "d.ppt"; Kind = "Failed" }
        )
    }

    It "<label>" -TestCases @(
        @{ label = "すべては全部を元の順で返す"; filter = "All"; expected = "a.docx,b.xlsx,c.pptx,d.ppt" }
        @{ label = "失敗のみは失敗の行だけを元の順で返す"; filter = "Failed"; expected = "a.docx,d.ppt" }
        @{ label = "分からない名前はすべて扱い"; filter = ""; expected = "a.docx,b.xlsx,c.pptx,d.ppt" }
    ) {
        param ($label, $filter, $expected)
        (@(selectIndexFileRows $rows $filter | ForEach-Object { $_.RelPath })) -join "," | Should -Be $expected
    }

    It "<label>" -TestCases @(
        @{ label = "失敗のみで 0 件なら案内を出す"; filter = "Failed"; count = 0; expected = "失敗したファイルはありません" }
        @{ label = "失敗のみで 1 件以上なら出さない"; filter = "Failed"; count = 2; expected = "" }
        @{ label = "すべてで 0 件でも出さない（見出しの要約が「まだ取り込んだファイルはありません」を言う）"; filter = "All"; count = 0; expected = "" }
    ) {
        param ($label, $filter, $count, $expected)
        getIndexFilesEmptyNotice $filter $count | Should -Be $expected
    }

    It "行が無い・null でも空を返す" {
        @(selectIndexFileRows @() "Failed").Count | Should -Be 0
        @(selectIndexFileRows $null "All").Count | Should -Be 0
    }

    It "<label>" -TestCases @(
        @{ label = "失敗が 0 件なら入口を出さない"; count = 0; shown = $false; text = "" }
        @{ label = "失敗があれば件数つきで出す"; count = 2; shown = $true; text = "更新に失敗したファイル 2 件" }
        @{ label = "千の位は区切る"; count = 1234; shown = $true; text = "更新に失敗したファイル 1,234 件" }
    ) {
        param ($label, $count, $shown, $text)
        $view = getIndexFailedEntryView $count
        $view.Shown | Should -Be $shown
        $view.Text | Should -Be $text
    }
}
