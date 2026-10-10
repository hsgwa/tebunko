# インデックス作成の確認に出す文言（tebunko\ui\indexing_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\indexing_view.ps1"

    function newPlanItem {
        param ([string]$kind, [int]$files = 0, [int]$targets = 0, [int]$new = 0, [int]$updated = 0, [int]$failed = 0)
        return [pscustomobject]@{
            インデックス名 = "売上"; 元のフォルダ = "C:\data\売上"; 区分 = $kind
            ファイル数 = $files; 取り込み対象 = $targets; 新規 = $new; 更新あり = $updated
            前回未完了 = 0; インデックスなし = 0; 前回失敗 = $failed
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

Describe "getIndexingEndText" -Tag Unit {
    It "取り込んだ（成功＋失敗）が1件以上なら、成功・失敗（・残り）の見出し" -TestCases @(
        @{ Success = 10; Failed = 0; Postponed = 0; Notice = ""; Text = "更新が終わりました（成功 10 件 / 失敗 0 件）"; Detail = "" }
        @{ Success = 8; Failed = 2; Postponed = 0; Notice = ""; Text = "更新が終わりました（成功 8 件 / 失敗 2 件）"; Detail = "失敗したファイルと原因は「更新に失敗したファイル」の一覧で確認できます。" }
        @{ Success = 9; Failed = 0; Postponed = 3; Notice = "案内"; Text = "更新が終わりました（成功 9 件 / 失敗 0 件 / 残り 3 件）"; Detail = "案内" }
        @{ Success = 7; Failed = 1; Postponed = 2; Notice = "案内"; Text = "更新が終わりました（成功 7 件 / 失敗 1 件 / 残り 2 件）"; Detail = "失敗したファイルと原因は「更新に失敗したファイル」の一覧で確認できます。 案内" }
    ) {
        param ($Success, $Failed, $Postponed, $Notice, $Text, $Detail)
        $view = getIndexingEndText $Success $Failed $Postponed $Notice
        $view.Text | Should -Be $Text
        $view.Detail | Should -Be $Detail
    }

    It "取り込んだのが0件で後回しがあれば、後回しの見出しと終わりの一言" {
        $view = getIndexingEndText 0 0 5 "PowerPoint が起動していたため、5 件を更新せずに残しました。"
        $view.Text | Should -Be "更新が終わりました（更新せずに残したファイル 5 件）"
        $view.Detail | Should -Be "PowerPoint が起動していたため、5 件を更新せずに残しました。"
    }

    It "どちらも0件なら「更新が必要なファイルはありませんでした」" {
        $view = getIndexingEndText 0 0 0 ""
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