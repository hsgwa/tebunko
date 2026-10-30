# インデックス作成の確認に出す文言（tebunko\ui\indexing_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\indexing_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_list_view.ps1"  # getIndexTopOnlyTagText（札の文言）

    function newPlanItem {
        param ([string]$kind, [int]$files = 0, [int]$targets = 0, [int]$new = 0, [int]$updated = 0, [int]$failed = 0,
            [int]$cloud = 0, [int]$cloudFailed = 0, [long]$cloudBytes = 0, [long]$cloudFailedBytes = 0, [string]$name = "売上")
        return [pscustomobject]@{
            インデックス名 = $name; 元のフォルダ = "C:\data\売上"; 区分 = $kind
            ファイル数 = $files; 取り込み対象 = $targets; 新規 = $new; 更新あり = $updated
            前回未完了 = 0; インデックスなし = 0; 前回失敗 = $failed
            クラウド = $cloud; クラウド失敗 = $cloudFailed; クラウド容量 = $cloudBytes; クラウド失敗容量 = $cloudFailedBytes
        }
    }
}

Describe "newPlanViewRows" -Tag Unit {
    It "更新するものがあれば「要更新」と、対象ファイル数・内訳を出す" {
        $row = (newPlanViewRows (newPlanItem ${planKindIngest} 120 12 10 2))[0]
        $row.TotalText | Should -Be "120"
        $row.StatusText | Should -Be "要更新"
        $row.Level | Should -Be "Wait"
        $row.DetailText | Should -Be "更新するファイル 12 件（新規 10 件 / 更新あり 2 件）"
    }

    It "更新対象が無ければ「最新」" {
        $row = (newPlanViewRows (newPlanItem ${planKindIngest} 120 0))[0]
        $row.StatusText | Should -Be "最新"
        $row.Level | Should -Be "Ok"
        $row.DetailText | Should -Be "すべて最新です"
    }

    It "チェックが外れていれば数えず「対象外」" {
        $row = (newPlanViewRows (newPlanItem ${planKindUnchecked}))[0]
        $row.StatusText | Should -Be "対象外"
        $row.Level | Should -Be "None"
        $row.TotalText | Should -Be "－"
    }

    It "元のフォルダが無ければ「フォルダなし」と出す" {
        $row = (newPlanViewRows (newPlanItem ${planKindMissing}))[0]
        $row.StatusText | Should -Be "フォルダなし"
        $row.Level | Should -Be "Ng"
    }

    It "件数は 3 桁ごとに区切る" {
        $row = (newPlanViewRows (newPlanItem ${planKindIngest} 12345 1234 1234))[0]
        $row.TotalText | Should -Be "12,345"
        $row.DetailText | Should -Be "更新するファイル 1,234 件（新規 1,234 件）"
    }

    It "行が無ければ空の配列" {
        $views = newPlanViewRows @()
        @($views).Count | Should -Be 0
    }
}

Describe "newPlanViewRows（直下のみの札）" -Tag Unit {
    It "subfolders が <label> の名前の行の TopOnlyText は「<text>」" -TestCases @(
        @{ label = "false"; subfolders = @{ "売上" = $false }; text = "直下のみ" }
        @{ label = "true"; subfolders = @{ "売上" = $true }; text = "" }
        @{ label = "表に無い"; subfolders = @{}; text = "" }
    ) {
        param ($label, $subfolders, $text)
        $row = (newPlanViewRows (newPlanItem ${planKindIngest} 120 12 10 2) @() $subfolders)[0]
        $row.TopOnlyText | Should -Be $text
    }
}

Describe "newPlanViewRows（サブフォルダ側を外すだけの行）" -Tag Unit {
    It "取り込むファイルが無くても、外す件数（SubfolderRemovedCount）が <removed> 件なら、ステータスは「<status>」" -TestCases @(
        @{ removed = 3; status = "要更新"; level = "Wait"; detail = "サブフォルダ側の 3 件をインデックスから外します" }
        @{ removed = 0; status = "最新"; level = "Ok"; detail = "すべて最新です" }
    ) {
        param ($removed, $status, $level, $detail)
        $item = newPlanItem ${planKindIngest} 10 0
        $item | Add-Member -NotePropertyName SubfolderRemovedCount -NotePropertyValue $removed
        $row = (newPlanViewRows @($item))[0]
        $row.StatusText | Should -Be $status
        $row.Level | Should -Be $level
        $row.DetailText | Should -Be $detail
    }
}

Describe "getIndexingSubfolderRemovalNote・getIndexingSubfolderRemovalCount" -Tag Unit {
    It "件数 <count> の文言は「<text>」" -TestCases @(
        @{ count = 0; text = "" }
        @{ count = 3; text = "サブフォルダ側の 3 件をインデックスから外します" }
        @{ count = 1234; text = "サブフォルダ側の 1,234 件をインデックスから外します" }
    ) {
        param ($count, $text)
        getIndexingSubfolderRemovalNote $count | Should -Be $text
    }

    It "確認に出す件数は、取り込む行の SubfolderRemovedCount の合計（選んだものだけの回は選んだ名前だけ。値が無い行は 0）" {
        $a = newPlanItem ${planKindIngest} 10 1 1 0 0 0 0 0 0 "売上"
        $a | Add-Member -NotePropertyName SubfolderRemovedCount -NotePropertyValue 2
        $b = newPlanItem ${planKindIngest} 10 1 1 0 0 0 0 0 0 "経理"
        $b | Add-Member -NotePropertyName SubfolderRemovedCount -NotePropertyValue 5
        $c = newPlanItem ${planKindIngest} 10 1 1 0 0 0 0 0 0 "総務"
        getIndexingSubfolderRemovalCount @($a, $b, $c) | Should -Be 7
        getIndexingSubfolderRemovalCount @($a, $b, $c) @("売上") | Should -Be 2
    }

    It "サブフォルダ側を外す件数があれば、取り込むファイルが無くても『何も無い』にせず、［更新を開始］と取りやめを選べる（閉じるだけにすると、見せたまま消してしまう）" {
        isIndexingConfirmNothing 0 0 0 0 0 | Should -Be $true
        isIndexingConfirmNothing 0 0 0 0 3 | Should -Be $false
        $text = getIndexingConfirmText 0 0 $false 1 0 0 $false 3
        $text.Button | Should -Be "更新を開始"
        $text.Total | Should -Be 0
        $a = newPlanItem ${planKindIngest} 10 0 0 0 0 0 0 0 0 "売上"
        $a | Add-Member -NotePropertyName SubfolderRemovedCount -NotePropertyValue 3
        $view = getIndexingConfirmView @($a)
        $view.Nothing | Should -Be $false
        $view.Confirm.Button | Should -Be "更新を開始"
    }

    It "getIndexingConfirmView は、件数があれば SubfolderRemovalNote に文言を、無ければ空を返す" -TestCases @(
        @{ removed = 4; note = "サブフォルダ側の 4 件をインデックスから外します" }
        @{ removed = 0; note = "" }
    ) {
        param ($removed, $note)
        $a = newPlanItem ${planKindIngest} 10 1 1 0 0 0 0 0 0 "売上"
        $a | Add-Member -NotePropertyName SubfolderRemovedCount -NotePropertyValue $removed
        (getIndexingConfirmView @($a)).SubfolderRemovalNote | Should -Be $note
    }
}

Describe "newPlanViewRows（選んだものだけの回）" -Tag Unit {
    It "onlyNames に無いインデックスは出さない（空なら全部）" {
        $plan = @(
            [pscustomobject]@{ インデックス名 = "売上"; 元のフォルダ = "C:\data\売上"; 区分 = ${planKindUnchecked}; ファイル数 = 0; 取り込み対象 = 0; 新規 = 0; 更新あり = 0; 前回未完了 = 0; インデックスなし = 0; 前回失敗 = 0 }
            [pscustomobject]@{ インデックス名 = "見積"; 元のフォルダ = "C:\data\見積"; 区分 = ${planKindUnchecked}; ファイル数 = 0; 取り込み対象 = 0; 新規 = 0; 更新あり = 0; 前回未完了 = 0; インデックスなし = 0; 前回失敗 = 0 }
        )
        (newPlanViewRows $plan @("見積")).Count | Should -Be 1
        (newPlanViewRows $plan @("見積"))[0].Name | Should -Be "見積"
        (newPlanViewRows $plan @()).Count | Should -Be 2
    }
}

Describe "newPlanViewRows（削除予定）" -Tag Unit {
    It "設定から外れたインデックスは「削除予定」（赤）で、削除されることを ToolTip に出す" {
        $row = (newPlanViewRows (newPlanItem ${planKindDropped}))[0]
        $row.StatusText | Should -Be "削除予定"
        $row.Level | Should -Be "Ng"
        $row.TotalText | Should -Be "－"
        $row.DetailText | Should -Be "設定に無いため、［更新を開始］でこのインデックスを削除します"
    }

    It "選んだものだけの回でも、onlyNames に無い名前の削除予定は必ず出す" {
        $plan = @(
            [pscustomobject]@{ インデックス名 = "見積"; 元のフォルダ = "C:\data\見積"; 区分 = ${planKindUnchecked}; ファイル数 = 0; 取り込み対象 = 0; 新規 = 0; 更新あり = 0; 前回未完了 = 0; インデックスなし = 0; 前回失敗 = 0 }
            (newPlanItem ${planKindDropped})
        )
        $rows = newPlanViewRows $plan @("見積")
        $rows.Count | Should -Be 2
        @($rows | Where-Object { $_.StatusText -eq "削除予定" }).Count | Should -Be 1
    }
}

Describe "getIndexingDroppedCount・削除予定のある確認の文言" -Tag Unit {
    It "削除予定の行だけを数える" {
        getIndexingDroppedCount @((newPlanItem ${planKindDropped}), (newPlanItem ${planKindIngest} 10 2), $null) | Should -Be 1
        getIndexingDroppedCount @() | Should -Be 0
    }

    It "確認の文言: 更新 <targets>・失敗 <failed>・再取り込み <retry>・削除予定 <dropped> のとき、ボタンは <button>" -TestCases @(
        @{ targets = 0; failed = 0; retry = $false; dropped = 2; button = "更新を開始"; text = "設定に無いインデックス 2 件を削除します" }
        @{ targets = 0; failed = 1; retry = $false; dropped = 1; button = "更新を開始"; text = "設定に無いインデックス 1 件を削除します" }
        @{ targets = 0; failed = 1; retry = $true; dropped = 1; button = "更新を開始"; text = "1 ファイル.*設定に無いインデックス 1 件を削除します" }
        @{ targets = 12; failed = 0; retry = $false; dropped = 1; button = "更新を開始"; text = "12 ファイル.*設定に無いインデックス 1 件を削除します" }
        @{ targets = 0; failed = 0; retry = $false; dropped = 0; button = "閉じる"; text = "すべて最新" }
    ) {
        param ($targets, $failed, $retry, $dropped, $button, $text)
        $view = getIndexingConfirmText $targets $failed $retry 1 $dropped
        $view.Button | Should -Be $button
        $view.Text | Should -Match $text
    }

    It "閉じるだけにするか: 更新 <targets>・失敗 <failed>・削除予定 <dropped> のとき <expected>" -TestCases @(
        @{ targets = 0; failed = 0; dropped = 0; expected = $true }
        @{ targets = 0; failed = 0; dropped = 1; expected = $false }
        @{ targets = 0; failed = 1; dropped = 0; expected = $false }
        @{ targets = 3; failed = 0; dropped = 0; expected = $false }
    ) {
        param ($targets, $failed, $dropped, $expected)
        isIndexingConfirmNothing $targets $failed $dropped | Should -Be $expected
    }

    It "確認の説明文: 削除予定 <dropped> 件のとき <expected>" -TestCases @(
        @{ dropped = 0; expected = "" }
        @{ dropped = 1; expected = "［更新を開始］を押すと、設定に無いインデックスも削除します。" }
        @{ dropped = 3; expected = "［更新を開始］を押すと、設定に無いインデックスも削除します。" }
    ) {
        param ($dropped, $expected)
        getIndexingConfirmIntro $dropped | Should -Be $expected
    }

    It "帯の消え方: <level> のとき 閉じるボタン <closable>・自動で消えるまで <seconds> 秒" -TestCases @(
        @{ level = "run"; closable = $true; seconds = 0 }
        @{ level = "resume"; closable = $true; seconds = 0 }
        @{ level = "ok"; closable = $true; seconds = 8 }
        @{ level = "info"; closable = $true; seconds = 0 }
        @{ level = "warn"; closable = $true; seconds = 0 }
        @{ level = "ng"; closable = $true; seconds = 0 }
        @{ level = ""; closable = $true; seconds = 0 }
        @{ level = "ok"; pending = 3; closable = $true; seconds = 0 }
        @{ level = "warn"; pending = 1; closable = $true; seconds = 0 }
        @{ level = "ng"; pending = 2; closable = $true; seconds = 0 }
        @{ level = "run"; pending = 5; closable = $true; seconds = 0 }
        @{ level = "ok"; pending = 0; closable = $true; seconds = 8 }
    ) {
        param ($level, $closable, $seconds, $pending = 0)
        $behavior = getIndexingBannerBehavior $level $pending
        $behavior.Closable | Should -Be $closable
        $behavior.AutoCloseSeconds | Should -Be $seconds
    }

    It "集計を読み直したときの帯: <name> は <action>・閉じた印 <dismissed>" -TestCases @(
        @{ name = "残りがあり帯が無い"; pending = 3; indexing = $false; visible = $false; kind = ""; was = $false; action = "Show"; dismissed = $false }
        @{ name = "閉じたあとの読み直しでは出し直さない"; pending = 3; indexing = $false; visible = $false; kind = ""; was = $true; action = "None"; dismissed = $true }
        @{ name = "更新を始めると印が戻る"; pending = 3; indexing = $true; visible = $false; kind = ""; was = $true; action = "None"; dismissed = $false }
        @{ name = "残りが 0 になると印が戻る"; pending = 0; indexing = $false; visible = $false; kind = ""; was = $true; action = "None"; dismissed = $false }
        @{ name = "印が戻ったあとは出る"; pending = 2; indexing = $false; visible = $false; kind = ""; was = $false; action = "Show"; dismissed = $false }
        @{ name = "出ている中断の帯は文だけ直す"; pending = 2; indexing = $false; visible = $true; kind = "interrupted"; was = $false; action = "Text"; dismissed = $false }
        @{ name = "残りが無くなった中断の帯は消す"; pending = 0; indexing = $false; visible = $true; kind = "interrupted"; was = $false; action = "Hide"; dismissed = $false }
        @{ name = "更新中に残っている中断の帯は消す"; pending = 2; indexing = $true; visible = $true; kind = "interrupted"; was = $false; action = "Hide"; dismissed = $false }
        @{ name = "更新中の帯には何もしない"; pending = 2; indexing = $true; visible = $true; kind = "running"; was = $false; action = "None"; dismissed = $false }
    ) {
        param ($pending, $indexing, $visible, $kind, $was, $action, $dismissed)
        $decision = getIndexingResumeBannerDecision $pending $indexing $visible $kind $was
        $decision.Action | Should -Be $action
        $decision.Dismissed | Should -Be $dismissed
        $decision.CanResume | Should -Be (($pending -gt 0) -and !$indexing)
    }

    It "終わりの帯は、残りの有無で秒数が変わったときだけ時計を決め直す: 残り <pending>・いま <current> 秒は <action>" -TestCases @(
        @{ pending = 0; current = 8; action = "None" }
        @{ pending = 0; current = 0; action = "Retime" }
        @{ pending = 3; current = 8; action = "Retime" }
        @{ pending = 3; current = 0; action = "None" }
    ) {
        param ($pending, $current, $action)
        $decision = getIndexingResumeBannerDecision $pending $false $true "done" $false "ok" $current
        $decision.Action | Should -Be $action
    }
}

Describe "getIndexingCurrentName・getIndexingSkippedView" -Tag Unit {
    It "取り込み中のファイル <current> のインデックス名は <expected>" -TestCases @(
        @{ current = "営業\2025\a.xlsx"; expected = "営業" }
        @{ current = "a.xlsx"; expected = "" }
        @{ current = ""; expected = "" }
    ) {
        param ($current, $expected)
        getIndexingCurrentName $current | Should -Be $expected
    }

    It "更新できなかった名前が無ければ null、あれば件数と名前・理由" {
        getIndexingSkippedView @() | Should -BeNullOrEmpty
        $view = getIndexingSkippedView @(@{ Name = "営業"; Reason = "設定にありません" })
        $view.Heading | Should -Be "1 件のインデックスは更新できませんでした。"
        $view.Detail | Should -Be "「営業」: 設定にありません"
    }
}

Describe "getIndexingBannerLevel（更新が終わった帯の色の種類）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "失敗なく完了: ok（緑）"; exitCode = 0; failed = 0; expected = "ok" }
        @{ label = "失敗したファイルがある完了: warn（橙）"; exitCode = 0; failed = 3; expected = "warn" }
        @{ label = "更新そのものができなかった: warn（橙。赤にしない）"; exitCode = 1; failed = 0; expected = "warn" }
        @{ label = "中止・取りやめ: warn（橙）"; exitCode = 2; failed = 0; expected = "warn" }
    ) {
        param ($label, $exitCode, $failed, $expected)
        getIndexingBannerLevel $exitCode $failed | Should -Be $expected
    }
}

Describe "getIndexingBannerView（帯に何を出すか）" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "更新中の帯は run（色の level は使わない）。閉じられ、自動では消えない"; kind = "running"; level = "info"; pending = 0; expectedLevel = "run"; closable = $true; seconds = 0 }
        @{ name = "残りの無い成功の帯は ok。閉じられ、8 秒で消える"; kind = "done"; level = "ok"; pending = 0; expectedLevel = "ok"; closable = $true; seconds = 8 }
        @{ name = "残りがある成功の帯は、再開のために自動では消えない"; kind = "done"; level = "ok"; pending = 3; expectedLevel = "ok"; closable = $true; seconds = 0 }
        @{ name = "注意の帯（warn）は自動では消えない"; kind = "done"; level = "warn"; pending = 0; expectedLevel = "warn"; closable = $true; seconds = 0 }
        @{ name = "中断の帯は warn のまま、閉じられ、自動では消えない"; kind = "interrupted"; level = "warn"; pending = 5; expectedLevel = "warn"; closable = $true; seconds = 0 }
        @{ name = "更新中の帯は、残りがあっても run のまま消えない"; kind = "running"; level = "ok"; pending = 4; indexing = $true; expectedLevel = "run"; closable = $true; seconds = 0 }
        @{ name = "更新中は残りを 0 とみる（終わりの帯の ok でも、残りがあるからと自動で消えない扱いにしない）"; kind = "done"; level = "ok"; pending = 4; indexing = $true; expectedLevel = "ok"; closable = $true; seconds = 8 }
        @{ name = "更新中でなければ残りを数える（同じ帯で、自動では消えない）"; kind = "done"; level = "ok"; pending = 4; indexing = $false; expectedLevel = "ok"; closable = $true; seconds = 0 }
    ) {
        param ($name, $kind, $level, $pending, $indexing = $false, $expectedLevel, $closable, $seconds)
        $view = getIndexingBannerView $kind $level $pending $indexing
        $view.Level | Should -Be $expectedLevel
        $view.Closable | Should -Be $closable
        $view.AutoCloseSeconds | Should -Be $seconds
    }
}

Describe "getIndexingEndText" -Tag Unit {
    It "取り込んだ（成功＋失敗）が1件以上なら、成功・失敗の見出し。利用者のPowerPointのために取り込めなかったファイルも失敗の件数に入っていて、「更新せずに残した」「残り」の文言は出ない" -TestCases @(
        @{ Success = 10; Failed = 0; Text = "更新が終わりました（成功 10 件 / 失敗 0 件）"; Detail = "" }
        @{ Success = 8; Failed = 2; Text = "更新が終わりました（成功 8 件 / 失敗 2 件）"; Detail = "失敗したファイルと原因は「ファイルの一覧」（失敗のみ）で確認できます。" }
    ) {
        param ($Success, $Failed, $Text, $Detail)
        $view = getIndexingEndText $Success $Failed
        $view.Text | Should -Be $Text
        $view.Detail | Should -Be $Detail
        $view.Text | Should -Not -Match "更新せずに残した|残り"
    }

    It "どちらも0件なら「更新が必要なファイルはありませんでした」" {
        $view = getIndexingEndText 0 0
        $view.Text | Should -Be "更新が必要なファイルはありませんでした"
        $view.Detail | Should -Be ""
    }
}

Describe "getIndexingStateText" -Tag Unit {
    It "残りがあり、インデックス作成中でなければ出す" -TestCases @(
        @{ Pending = 3; Indexing = $false; Expected = "更新を中断しました（残り 3 件）" }
        @{ Pending = 0; Indexing = $false; Expected = "" }
        @{ Pending = 3; Indexing = $true; Expected = "" }
    ) {
        param ($Pending, $Indexing, $Expected)
        getIndexingStateText $Pending $Indexing | Should -Be $Expected
    }
}

Describe "getIndexingConfirmFolderCount" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "取り込み対象があるフォルダだけ数える"; retry = $false; expected = 2 }
        @{ name = "失敗分を含めると、失敗だけのフォルダも数える"; retry = $true; expected = 3 }
    ) {
        param ($name, $retry, $expected)
        $plan = @(
            (newPlanItem ${planKindIngest} 10 5)
            (newPlanItem ${planKindIngest} 10 3)
            (newPlanItem ${planKindIngest} 10 0 0 0 4)
            (newPlanItem ${planKindIngest} 10 0)
        )
        getIndexingConfirmFolderCount $plan $retry | Should -Be $expected
    }

    It "取り込み対象でない区分（最新・対象外）と空の予定は数えない" {
        getIndexingConfirmFolderCount @((newPlanItem ${planKindUnchecked} 10 5), $null) $true | Should -Be 0
        getIndexingConfirmFolderCount @() $false | Should -Be 0
    }
}

Describe "getIndexingConfirmText" -Tag Unit {
    It "取り込み対象があれば件数と［更新を開始］" {
        $view = getIndexingConfirmText 12 3 $false 2
        $view.Total | Should -Be 12
        $view.Text | Should -Be "更新対象: 2 フォルダ / 12 ファイル（最新のフォルダは更新しません）"
        $view.Button | Should -Be "更新を開始"
    }

    It "失敗分も再取り込みするなら足す" {
        (getIndexingConfirmText 12 3 $true).Total | Should -Be 15
    }

    It "0 件なら［閉じる］にする" {
        $view = getIndexingConfirmText 0 0 $false
        $view.Text | Should -Be "更新が必要なファイルはありません（すべて最新です）。"
        $view.Button | Should -Be "閉じる"
    }

    It "失敗分だけがあるときは、再取り込みのチェックで開始に変わる" {
        (getIndexingConfirmText 0 5 $false).Button | Should -Be "閉じる"
        (getIndexingConfirmText 0 5 $true).Button | Should -Be "更新を開始"
    }
}

Describe "getReingestConfirm" -Tag Unit {
    It "しるしがあり content_index が空のときだけ確かめの文言を返す" -TestCases @(
        @{ hasLegacyIndex = $true;  contentEmpty = $true;  expectConfirm = $true }
        @{ hasLegacyIndex = $true;  contentEmpty = $false; expectConfirm = $false }
        @{ hasLegacyIndex = $false; contentEmpty = $true;  expectConfirm = $false }
        @{ hasLegacyIndex = $false; contentEmpty = $false; expectConfirm = $false }
    ) {
        param ($hasLegacyIndex, $contentEmpty, $expectConfirm)
        $text = getReingestConfirm $hasLegacyIndex $contentEmpty
        if ($expectConfirm) {
            $text | Should -Not -BeNullOrEmpty
        } else {
            $text | Should -Be ""
        }
    }
}

Describe "クラウドにだけあるファイルの確認の文言" -Tag Unit {
    It "newPlanViewRows は、内訳に「クラウドにだけある N 件」を出す（取り込み対象には数えない）" {
        $rows = newPlanViewRows @((newPlanItem ${planKindIngest} 10 2 2 0 0 3 1 3000 500))
        $rows[0].DetailText | Should -Be "更新するファイル 2 件（新規 2 件 / クラウドにだけある 4 件）"
        $rows = newPlanViewRows @((newPlanItem ${planKindIngest} 10 0 0 0 0 3))
        # 取り込み対象が 0 件でも、クラウドにだけあるファイルがあれば「最新」にせず、注意の色で出す
        $rows[0].StatusText | Should -Be "要確認"
        $rows[0].Level | Should -Be "Wait"
        $rows[0].DetailText | Should -Be "クラウドにだけある 3 件"
        $rows = newPlanViewRows @((newPlanItem ${planKindIngest} 10 0 0 0 0 0 0 0 0))
        $rows[0].StatusText | Should -Be "最新"
        $rows[0].Level | Should -Be "Ok"
        # 前回失敗だけがクラウドにだけある場合は、手元の前回失敗と同じく「最新」のまま、内訳に出す（更新し直す指定のときだけ取り込む）
        $rows = newPlanViewRows @((newPlanItem ${planKindIngest} 10 0 0 0 0 0 1 0 500))
        $rows[0].StatusText | Should -Be "最新"
        $rows[0].Level | Should -Be "Ok"
        $rows[0].DetailText | Should -Be "クラウドにだけある 1 件"
    }

    It "formatCloudSize は KB・MB・GB のどれかで、小数点以下 1 桁にする" -TestCases @(
        @{ bytes = 0; expected = "1 KB 未満" }
        @{ bytes = 1023; expected = "1 KB 未満" }
        @{ bytes = 1024; expected = "約 1.0 KB" }
        @{ bytes = 1536; expected = "約 1.5 KB" }
        # 小数点以下 1 桁に丸めて 1024 になるときは、次の単位で書く
        @{ bytes = 1048575; expected = "約 1.0 MB" }
        @{ bytes = 1048000; expected = "約 1023.4 KB" }
        @{ bytes = 1073741823; expected = "約 1.0 GB" }
        @{ bytes = 1048576; expected = "約 1.0 MB" }
        @{ bytes = 13002342; expected = "約 12.4 MB" }
        @{ bytes = 1073741824; expected = "約 1.0 GB" }
        @{ bytes = 5368709120; expected = "約 5.0 GB" }
    ) {
        param ($bytes, $expected)
        formatCloudSize $bytes | Should -Be $expected
    }

    It "getIndexingCloudView は件数と合計サイズを足す。前回失敗は「失敗分も更新し直す」のときだけ足し、選んだものだけの回は選んだインデックスだけ" {
        $plan = @(
            (newPlanItem ${planKindIngest} 10 2 2 0 0 3 1 3000 500 "営業")
            (newPlanItem ${planKindIngest} 10 0 0 0 0 2 0 200 0 "経理")
            (newPlanItem ${planKindUnchecked} 0 0 0 0 0 9 9 9 9 "対象外")
        )
        $all = getIndexingCloudView $plan $false
        $all.Count | Should -Be 5
        $all.Bytes | Should -Be 3200
        $retry = getIndexingCloudView $plan $true
        $retry.Count | Should -Be 6
        $retry.Bytes | Should -Be 3700
        $only = getIndexingCloudView $plan $false @("経理")
        $only.Count | Should -Be 2
        $only.Bytes | Should -Be 200
        (getIndexingCloudView @() $true).Count | Should -Be 0
    }

    It "getIndexingConfirmFolderCount は、ダウンロードして取り込むときだけ、クラウドだけのフォルダも数える" {
        $plan = @((newPlanItem ${planKindIngest} 10 5), (newPlanItem ${planKindIngest} 10 0 0 0 0 2), (newPlanItem ${planKindIngest} 10 0 0 0 0 0 3))
        getIndexingConfirmFolderCount $plan $false | Should -Be 1
        getIndexingConfirmFolderCount $plan $false $true | Should -Be 2
        getIndexingConfirmFolderCount $plan $true $true | Should -Be 3
        # 選んだものだけの回は、選んだインデックスのフォルダだけ数える
        $named = @((newPlanItem ${planKindIngest} 10 5 5 0 0 0 0 0 0 "営業"), (newPlanItem ${planKindIngest} 10 5 5 0 0 0 0 0 0 "経理"))
        getIndexingConfirmFolderCount $named $false $false @("経理") | Should -Be 1
    }

    It "isIndexingConfirmNothing は、クラウドにだけあるファイルが 1 件でもあれば「何も無い」にしない" {
        isIndexingConfirmNothing 0 0 0 | Should -Be $true
        isIndexingConfirmNothing 0 0 0 0 | Should -Be $true
        isIndexingConfirmNothing 0 0 0 1 | Should -Be $false
    }

    It "getIndexingConfirmText は、クラウドにだけあるファイルの合計・ボタンをチェックに合わせる" -TestCases @(
        @{ name = "対象があり、ダウンロードしない"; targets = 12; cloud = 3; include = $false; total = 12; button = "更新を開始"
           text = "更新対象: 2 フォルダ / 12 ファイル（最新のフォルダは更新しません） クラウドにだけあるファイル 3 件は、ダウンロードせずに残します。" }
        @{ name = "対象があり、ダウンロードする"; targets = 12; cloud = 3; include = $true; total = 15; button = "更新を開始"
           text = "更新対象: 2 フォルダ / 15 ファイル（最新のフォルダは更新しません）" }
        @{ name = "対象が 0 件でクラウドだけ、ダウンロードしない（［閉じる］にせず［更新を開始］のまま）"; targets = 0; cloud = 3; include = $false; total = 0; button = "更新を開始"
           text = "更新が必要なファイルはありません。クラウドにだけあるファイル 3 件は、ダウンロードせずに残します。" }
        @{ name = "対象が 0 件でクラウドだけ、ダウンロードする"; targets = 0; cloud = 3; include = $true; total = 3; button = "更新を開始"
           text = "更新対象: 2 フォルダ / 3 ファイル（最新のフォルダは更新しません）" }
        @{ name = "クラウドが無ければ今までどおり"; targets = 12; cloud = 0; include = $false; total = 12; button = "更新を開始"
           text = "更新対象: 2 フォルダ / 12 ファイル（最新のフォルダは更新しません）" }
    ) {
        param ($name, $targets, $cloud, $include, $total, $button, $text)
        $view = getIndexingConfirmText $targets 0 $false 2 0 $cloud $include
        $view.Total | Should -Be $total
        $view.Text | Should -Be $text
        $view.Button | Should -Be $button
    }

    It "getIndexingConfirmText は、削除予定とクラウドの文を並べる" {
        $view = getIndexingConfirmText 0 0 $false 0 1 3 $false
        $view.Button | Should -Be "更新を開始"
        $view.Text | Should -Be "更新が必要なファイルはありません。設定に無いインデックス 1 件を削除します。 クラウドにだけあるファイル 3 件は、ダウンロードせずに残します。"
    }

    It "注意書き・チェックの文言に件数と合計サイズを入れる" {
        getIndexingCloudWarning 3 13002342 | Should -Be ("クラウドにだけあるファイル 3 件（合計 約 12.4 MB）は、既定ではダウンロードせずに残します。" +
            "取り込むと、そのすべてをこの PC にダウンロードします。登録するフォルダを、要るものだけに絞ることをおすすめします。")
        getIndexingCloudCheckText 1234 | Should -Be "クラウドにだけあるファイル 1,234 件もダウンロードして取り込む"
        ${indexingCloudRememberText} | Should -Be "次からこの確認を出さずにダウンロードする"
    }

    It "2 度目の確認は、見出しに件数と合計サイズを入れ、通信量・ディスク・残り続けること・制限時間を知らせる" {
        $view = getIndexingCloudConfirmView 3 13002342 10
        $view.Heading | Should -Be "3 件（合計 約 12.4 MB）をダウンロードして取り込みますか？"
        $view.Facts.Count | Should -Be 3
        $view.Facts[0] | Should -Match "帯域.*ディスクの空き"
        $view.Facts[1] | Should -Match "空き容量を増やす"
        $view.Facts[2] | Should -Match "10 分以内"
        $view.Facts[2] | Should -Match "手元に落としてから取り込み直"
        # 制限時間は呼び出し側が渡した値を使う
        (getIndexingCloudConfirmView 3 13002342 15).Facts[2] | Should -Match "15 分以内"
        $view.Choices[0].Text | Should -Be "ダウンロードして取り込む"
        $view.CancelText | Should -Be "キャンセル"
    }

    It "testIndexingCloudNeedsConfirm は、チェックが付いていて、設定がまだ download でないときだけ 2 度目の確認を出す" -TestCases @(
        @{ checked = $true; saved = "ask"; expected = $true }
        @{ checked = $true; saved = "download"; expected = $false }
        @{ checked = $false; saved = "ask"; expected = $false }
        @{ checked = $false; saved = "download"; expected = $false }
    ) {
        param ($checked, $saved, $expected)
        testIndexingCloudNeedsConfirm $checked $saved | Should -Be $expected
    }

    It "getIndexingCloudSavedMode は、両方のチェックが付いているときだけ download を保存する" -TestCases @(
        @{ cloud = $true; remember = $true; expected = "download" }
        @{ cloud = $true; remember = $false; expected = "ask" }
        @{ cloud = $false; remember = $true; expected = "ask" }
        @{ cloud = $false; remember = $false; expected = "ask" }
    ) {
        param ($cloud, $remember, $expected)
        getIndexingCloudSavedMode $cloud $remember | Should -Be $expected
    }
}

Describe "getIndexingEndText（クラウドにだけあるファイル）" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "取り込みが 0 件でクラウドを残したなら、その一言だけ（「更新が必要なファイルはありませんでした」は出さない）"; success = 0; failed = 0; cloud = 2
           text = "クラウドにだけあるファイル 2 件は、ダウンロードせずに残しました（次のインデックス作成でも確かめます）。"; detail = "" }
        @{ name = "取り込んだものがあり、クラウドを残したなら、説明に足す"; success = 5; failed = 0; cloud = 2
           text = "更新が終わりました（成功 5 件 / 失敗 0 件）"; detail = "クラウドにだけあるファイル 2 件は、ダウンロードせずに残しました（次のインデックス作成でも確かめます）。" }
        @{ name = "取り込んだものがあり、クラウドを残していなければ今までどおり"; success = 5; failed = 0; cloud = 0
           text = "更新が終わりました（成功 5 件 / 失敗 0 件）"; detail = "" }
        @{ name = "どちらも 0 件なら今までどおり"; success = 0; failed = 0; cloud = 0
           text = "更新が必要なファイルはありませんでした"; detail = "" }
    ) {
        param ($name, $success, $failed, $cloud, $text, $detail)
        $view = getIndexingEndText $success $failed $cloud
        $view.Text | Should -Be $text
        $view.Detail | Should -Be $detail
    }
}

Describe "getIndexingConfirmView（チェックの値で変わる表示）" -Tag Unit {
    # 1 つのインデックス（ファイル 20 件）。targets・failed（手元にある前回失敗）・cloud・cloudFailed（クラウドにだけある前回失敗）で組む
    It "<name>" -TestCases @(
        @{ name = "クラウドにだけある前回失敗だけ・更新し直さない: 更新し直すのチェックを出し、クラウドの欄は出さない（ダウンロードのチェックが付いたままでも取り込まない）"
           targets = 0; failed = 0; cloud = 0; cloudFailed = 1; retry = $false; checked = $true
           retryVisible = $true; retryText = "前回更新に失敗し、その後変わっていないファイル 1 件も更新し直す（パスワード付きなど）。うちクラウドにだけある 1 件は、ダウンロードして取り込むを選んだときだけ"
           cloudVisible = $false; include = $false; nothing = $false; total = 0; button = "閉じる"
           text = "更新が必要なファイルはありません（すべて最新です）。" }
        @{ name = "クラウドにだけある前回失敗だけ・更新し直す: クラウドの欄を 1 件で出し、チェックなしなら残す"
           targets = 0; failed = 0; cloud = 0; cloudFailed = 1; retry = $true; checked = $false
           retryVisible = $true; retryText = "前回更新に失敗し、その後変わっていないファイル 1 件も更新し直す（パスワード付きなど）。うちクラウドにだけある 1 件は、ダウンロードして取り込むを選んだときだけ"
           cloudVisible = $true; include = $false; nothing = $false; total = 0; button = "更新を開始"
           text = "更新が必要なファイルはありません。クラウドにだけあるファイル 1 件は、ダウンロードせずに残します。" }
        @{ name = "クラウドにだけある前回失敗だけ・更新し直す・ダウンロードする: 1 件を取り込む"
           targets = 0; failed = 0; cloud = 0; cloudFailed = 1; retry = $true; checked = $true
           retryVisible = $true; retryText = "前回更新に失敗し、その後変わっていないファイル 1 件も更新し直す（パスワード付きなど）。うちクラウドにだけある 1 件は、ダウンロードして取り込むを選んだときだけ"
           cloudVisible = $true; include = $true; nothing = $false; total = 1; button = "更新を開始"
           text = "更新対象: 1 フォルダ / 1 ファイル（最新のフォルダは更新しません）" }
        @{ name = "手元にある前回失敗 2 件とクラウド 3 件・更新し直す・ダウンロードしない"
           targets = 12; failed = 2; cloud = 3; cloudFailed = 0; retry = $true; checked = $false
           retryVisible = $true; retryText = "前回更新に失敗し、その後変わっていないファイル 2 件も更新し直す（パスワード付きなど）"
           cloudVisible = $true; include = $false; nothing = $false; total = 14; button = "更新を開始"
           text = "更新対象: 1 フォルダ / 14 ファイル（最新のフォルダは更新しません） クラウドにだけあるファイル 3 件は、ダウンロードせずに残します。" }
        @{ name = "手元にある前回失敗 2 件とクラウド 3 件・更新し直す・ダウンロードする"
           targets = 12; failed = 2; cloud = 3; cloudFailed = 0; retry = $true; checked = $true
           retryVisible = $true; retryText = "前回更新に失敗し、その後変わっていないファイル 2 件も更新し直す（パスワード付きなど）"
           cloudVisible = $true; include = $true; nothing = $false; total = 17; button = "更新を開始"
           text = "更新対象: 1 フォルダ / 17 ファイル（最新のフォルダは更新しません）" }
        @{ name = "手元にある前回失敗 2 件とクラウド 3 件・更新し直さない・ダウンロードする"
           targets = 12; failed = 2; cloud = 3; cloudFailed = 0; retry = $false; checked = $true
           retryVisible = $true; retryText = "前回更新に失敗し、その後変わっていないファイル 2 件も更新し直す（パスワード付きなど）"
           cloudVisible = $true; include = $true; nothing = $false; total = 15; button = "更新を開始"
           text = "更新対象: 1 フォルダ / 15 ファイル（最新のフォルダは更新しません）" }
        @{ name = "何も無ければ「何も無い」"
           targets = 0; failed = 0; cloud = 0; cloudFailed = 0; retry = $false; checked = $false
           retryVisible = $false; retryText = $null
           cloudVisible = $false; include = $false; nothing = $true; total = 0; button = "閉じる"
           text = "更新が必要なファイルはありません（すべて最新です）。" }
    ) {
        param ($name, $targets, $failed, $cloud, $cloudFailed, $retry, $checked, $retryVisible, $retryText, $cloudVisible, $include, $nothing, $total, $button, $text)
        $plan = @((newPlanItem ${planKindIngest} 20 $targets $targets 0 $failed $cloud $cloudFailed ($cloud * 1000) ($cloudFailed * 500)))
        $view = getIndexingConfirmView $plan @() $retry $checked
        $view.RetryVisible | Should -Be $retryVisible
        if ($retryVisible) { $view.RetryText | Should -Be $retryText }
        $view.CloudVisible | Should -Be $cloudVisible
        $view.Include | Should -Be $include
        $view.Nothing | Should -Be $nothing
        $view.Confirm.Total | Should -Be $total
        $view.Confirm.Button | Should -Be $button
        $view.Confirm.Text | Should -Be $text
        if ($cloudVisible) {
            $view.CloudWarn | Should -Match "^クラウドにだけあるファイル "
            $view.CloudCheckText | Should -Be (getIndexingCloudCheckText $view.Cloud.Count)
        } else {
            $view.CloudWarn | Should -Be ""
            $view.CloudCheckText | Should -Be ""
        }
    }

    It "選んだものだけの回は、選んだインデックスの件数だけを使う。削除予定は選ばなくても数える" {
        $plan = @(
            (newPlanItem ${planKindIngest} 10 5 5 0 1 2 0 200 0 "営業")
            (newPlanItem ${planKindIngest} 10 0 0 0 0 3 0 300 0 "経理")
            (newPlanItem ${planKindDropped})
        )
        $view = getIndexingConfirmView $plan @("経理") $false $false
        $view.Targets | Should -Be 0
        $view.Failed | Should -Be 0
        $view.RetryVisible | Should -Be $false
        $view.Cloud.Count | Should -Be 3
        $view.Dropped | Should -Be 1
        $all = getIndexingConfirmView $plan @() $false $false
        $all.Targets | Should -Be 5
        $all.Failed | Should -Be 1
        $all.Cloud.Count | Should -Be 5
    }

    It "getIndexingCloudModeToSave は、クラウドの欄を出したときだけ、今の設定と違う値を返す（無ければ `$null）" -TestCases @(
        @{ visible = $false; checked = $true; remember = $true; saved = "ask"; expected = $null }
        @{ visible = $true; checked = $true; remember = $true; saved = "ask"; expected = "download" }
        @{ visible = $true; checked = $true; remember = $true; saved = "download"; expected = $null }
        @{ visible = $true; checked = $true; remember = $false; saved = "download"; expected = "ask" }
        @{ visible = $true; checked = $false; remember = $true; saved = "download"; expected = "ask" }
        @{ visible = $true; checked = $false; remember = $false; saved = "ask"; expected = $null }
        @{ visible = $true; checked = $true; remember = $false; saved = "ask"; expected = $null }
    ) {
        param ($visible, $checked, $remember, $saved, $expected)
        $mode = getIndexingCloudModeToSave $visible $checked $remember $saved
        if ($null -eq $expected) { $mode | Should -BeNullOrEmpty } else { $mode | Should -Be $expected }
    }

    It "testIndexingPlanSelected は、選んだものだけの回では選んだインデックスの行だけ「選ばれた」とする" -TestCases @(
        @{ names = @(); expected = $true }
        @{ names = @("売上"); expected = $true }
        @{ names = @("経理"); expected = $false }
    ) {
        param ($names, $expected)
        testIndexingPlanSelected (newPlanItem ${planKindIngest}) $names | Should -Be $expected
    }

    It "testIndexingPlanSelected は `$null の行を選ばれたとしない" {
        testIndexingPlanSelected $null @() | Should -Be $false
    }
}

Describe "getIndexingConfirmText（失敗分の更新し直し × クラウド）" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "更新し直す・ダウンロードしない: 手元の失敗分だけ足し、クラウドは残す"; retry = $true; include = $false; total = 14
           text = "更新対象: 2 フォルダ / 14 ファイル（最新のフォルダは更新しません） クラウドにだけあるファイル 3 件は、ダウンロードせずに残します。" }
        @{ name = "更新し直す・ダウンロードする: 手元の失敗分もクラウドも足す"; retry = $true; include = $true; total = 17
           text = "更新対象: 2 フォルダ / 17 ファイル（最新のフォルダは更新しません）" }
        @{ name = "更新し直さない・ダウンロードしない: 取り込み対象だけ"; retry = $false; include = $false; total = 12
           text = "更新対象: 2 フォルダ / 12 ファイル（最新のフォルダは更新しません） クラウドにだけあるファイル 3 件は、ダウンロードせずに残します。" }
    ) {
        param ($name, $retry, $include, $total, $text)
        $view = getIndexingConfirmText 12 2 $retry 2 0 3 $include
        $view.Total | Should -Be $total
        $view.Text | Should -Be $text
        $view.Button | Should -Be "更新を開始"
    }
}
