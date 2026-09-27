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
    It "取り込むものがあれば件数と内訳を出す" {
        $row = (newPlanViewRows (newPlanItem ${planKindIngest} 120 12 10 2))[0]
        $row.TotalText | Should -Be "120 件"
        $row.TargetText | Should -Be "12 件"
        $row.Tone | Should -Be "info"
        $row.DetailText | Should -Be "新規 10 件 / 更新あり 2 件"
    }

    It "取り込み対象が無ければ「更新不要」" {
        $row = (newPlanViewRows (newPlanItem ${planKindIngest} 120 0))[0]
        $row.TargetText | Should -Be "更新不要"
        $row.Tone | Should -Be "ok"
        $row.DetailText | Should -Be "すべて取り込み済みです"
    }

    It "チェックが外れていれば数えない" {
        $row = (newPlanViewRows (newPlanItem ${planKindUnchecked}))[0]
        $row.TargetText | Should -Be "取り込みません"
        $row.Tone | Should -Be "gray"
        $row.TotalText | Should -Be "－"
    }

    It "元のフォルダが無ければ取り込めないと出す" {
        $row = (newPlanViewRows (newPlanItem ${planKindMissing}))[0]
        $row.TargetText | Should -Be "取り込めません"
        $row.Tone | Should -Be "ng"
    }

    It "件数は3桁ごとに区切る" {
        $row = (newPlanViewRows (newPlanItem ${planKindIngest} 12345 1234 1234))[0]
        $row.TotalText | Should -Be "12,345 件"
        $row.TargetText | Should -Be "1,234 件"
    }

    It "行が無ければ空の配列" {
        $views = newPlanViewRows @()
        @($views).Count | Should -Be 0
    }
}

Describe "getIndexingEndText" -Tag Unit {
    It "取り込んだ（成功＋失敗）が1件以上なら、成功・失敗（・残り）の見出し" -TestCases @(
        @{ Success = 10; Failed = 0; Postponed = 0; Notice = ""; Text = "インデックス作成が終わりました（成功 10 件 / 失敗 0 件）"; Detail = "" }
        @{ Success = 8; Failed = 2; Postponed = 0; Notice = ""; Text = "インデックス作成が終わりました（成功 8 件 / 失敗 2 件）"; Detail = "失敗したファイルと原因は「取り込みに失敗したファイル」の一覧で確認できます。" }
        @{ Success = 9; Failed = 0; Postponed = 3; Notice = "案内"; Text = "インデックス作成が終わりました（成功 9 件 / 失敗 0 件 / 残り 3 件）"; Detail = "案内" }
        @{ Success = 7; Failed = 1; Postponed = 2; Notice = "案内"; Text = "インデックス作成が終わりました（成功 7 件 / 失敗 1 件 / 残り 2 件）"; Detail = "失敗したファイルと原因は「取り込みに失敗したファイル」の一覧で確認できます。 案内" }
    ) {
        param ($Success, $Failed, $Postponed, $Notice, $Text, $Detail)
        $view = getIndexingEndText $Success $Failed $Postponed $Notice
        $view.Text | Should -Be $Text
        $view.Detail | Should -Be $Detail
    }

    It "取り込んだのが0件で後回しがあれば、後回しの見出しと終わりの一言" {
        $view = getIndexingEndText 0 0 5 "PowerPoint が起動していたため、5 件を取り込まずに残しました。"
        $view.Text | Should -Be "インデックス作成が終わりました（取り込まずに残したファイル 5 件）"
        $view.Detail | Should -Be "PowerPoint が起動していたため、5 件を取り込まずに残しました。"
    }

    It "どちらも0件なら「取り込みが必要なファイルはありませんでした」" {
        $view = getIndexingEndText 0 0 0 ""
        $view.Text | Should -Be "取り込みが必要なファイルはありませんでした"
        $view.Detail | Should -Be ""
    }
}

Describe "getIndexingStateText" -Tag Unit {
    It "残りがあり、インデックス作成中でなければ出す" -TestCases @(
        @{ Pending = 3; Indexing = $false; Expected = "⏸ まだ取り込んでいないファイルがあります（残り 3 件）" }
        @{ Pending = 0; Indexing = $false; Expected = "" }
        @{ Pending = 3; Indexing = $true; Expected = "" }
    ) {
        param ($Pending, $Indexing, $Expected)
        getIndexingStateText $Pending $Indexing | Should -Be $Expected
    }
}

Describe "getIndexingConfirmText" -Tag Unit {
    It "取り込み対象があれば件数と［インデックス作成を開始］" {
        $view = getIndexingConfirmText 12 3 $false
        $view.Total | Should -Be 12
        $view.Text | Should -Be "合計 12 件を取り込みます。"
        $view.Button | Should -Be "インデックス作成を開始"
    }

    It "失敗分も再取り込みするなら足す" {
        (getIndexingConfirmText 12 3 $true).Total | Should -Be 15
    }

    It "0 件なら［閉じる］にする" {
        $view = getIndexingConfirmText 0 0 $false
        $view.Text | Should -Be "更新が必要なファイルはありません（すべて取り込み済みです）。"
        $view.Button | Should -Be "閉じる"
    }

    It "失敗分だけがあるときは、再取り込みのチェックで開始に変わる" {
        (getIndexingConfirmText 0 5 $false).Button | Should -Be "閉じる"
        (getIndexingConfirmText 0 5 $true).Button | Should -Be "インデックス作成を開始"
    }
}