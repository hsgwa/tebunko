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
        $view = getIndexingEndText 0 0 5 "PowerPoint が起動していたため、5 件を取り込まずに残しました。"
        $view.Text | Should -Be "更新が終わりました（更新せずに残したファイル 5 件）"
        $view.Detail | Should -Be "PowerPoint が起動していたため、5 件を取り込まずに残しました。"
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