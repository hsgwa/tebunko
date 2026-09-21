# 変換の確認に出す文言（win_grep\ui\convert_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\win_grep\ui\convert_view.ps1"

function newPlanItem {
    param ([string]$kind, [int]$files = 0, [int]$targets = 0, [int]$new = 0, [int]$updated = 0, [int]$failed = 0)
    return [pscustomobject]@{
        インデックス名 = "売上"; 元のフォルダ = "C:\data\売上"; 区分 = $kind
        ファイル数 = $files; 変換対象 = $targets; 新規 = $new; 更新あり = $updated
        前回未完了 = 0; 変換結果なし = 0; 前回失敗 = $failed
    }
}

Describe "newPlanViewRows" -Tag Unit {
    It "変換するものがあれば件数と内訳を出す" {
        $row = (newPlanViewRows (newPlanItem ${planKindConvert} 120 12 10 2))[0]
        $row.TotalText | Should Be "120 件"
        $row.TargetText | Should Be "12 件"
        $row.Tone | Should Be "info"
        $row.DetailText | Should Be "新規 10 件 / 更新あり 2 件"
    }

    It "変換対象が無ければ「更新不要」" {
        $row = (newPlanViewRows (newPlanItem ${planKindConvert} 120 0))[0]
        $row.TargetText | Should Be "更新不要"
        $row.Tone | Should Be "ok"
        $row.DetailText | Should Be "すべて変換済みです"
    }

    It "チェックが外れていれば数えない" {
        $row = (newPlanViewRows (newPlanItem ${planKindUnchecked}))[0]
        $row.TargetText | Should Be "変換しません"
        $row.Tone | Should Be "gray"
        $row.TotalText | Should Be "－"
    }

    It "元のフォルダが無ければ変換できないと出す" {
        $row = (newPlanViewRows (newPlanItem ${planKindMissing}))[0]
        $row.TargetText | Should Be "変換できません"
        $row.Tone | Should Be "ng"
    }

    It "件数は3桁ごとに区切る" {
        $row = (newPlanViewRows (newPlanItem ${planKindConvert} 12345 1234 1234))[0]
        $row.TotalText | Should Be "12,345 件"
        $row.TargetText | Should Be "1,234 件"
    }

    It "行が無ければ空の配列" {
        $views = newPlanViewRows @()
        @($views).Count | Should Be 0
    }
}

Describe "getConvertConfirmText" -Tag Unit {
    It "変換対象があれば件数と［変換を開始］" {
        $view = getConvertConfirmText 12 3 $false
        $view.Total | Should Be 12
        $view.Text | Should Be "合計 12 件を変換します。"
        $view.Button | Should Be "変換を開始"
    }

    It "失敗分も再変換するなら足す" {
        (getConvertConfirmText 12 3 $true).Total | Should Be 15
    }

    It "0 件なら［閉じる］にする" {
        $view = getConvertConfirmText 0 0 $false
        $view.Text | Should Be "更新が必要なファイルはありません（すべて変換済みです）。"
        $view.Button | Should Be "閉じる"
    }

    It "失敗分だけがあるときは、再変換のチェックで開始に変わる" {
        (getConvertConfirmText 0 5 $false).Button | Should Be "閉じる"
        (getConvertConfirmText 0 5 $true).Button | Should Be "変換を開始"
    }
}