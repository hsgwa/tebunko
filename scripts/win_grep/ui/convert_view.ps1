# 変換の確認ダイアログ・進み具合に出す文言の決定。
# 画面に触らないため、そのままテストできる（tests\win_grep\ui\convert_view.Tests.ps1）。
#
# 色は「意味」（Tone）で返し、実際の色は画面側（convert_tab.ps1）で対応表から引く。
#   info = これから変換する / ok = 変換の必要なし / warn = 注意 / ng = 変換できない / gray = 対象外

function newPlanViewRows {
    # 変換予定（変換予定.tsv の行）を、確認のダイアログに出す形にする
    param (
        $plan  # readConvertPlan の結果
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($plan)) {
        # 空の配列を渡すと @($plan) に $null が 1 つ入るため、ここで外す
        if ($null -eq $item) { continue }
        $row = @{ Name = $item.インデックス名; Path = $item.元のフォルダ; Tone = "info" }
        if ($item.区分 -eq ${planKindUnchecked}) {
            $row.TargetText = "変換しません"
            $row.Tone = "gray"
            $row.DetailText = "［変換］のチェックが外れています（インデックスはそのまま残します）"
            $row.TotalText = "－"
        } elseif ($item.区分 -eq ${planKindMissing}) {
            $row.TargetText = "変換できません"
            $row.Tone = "ng"
            $row.DetailText = "元のフォルダが見つかりません（［編集…］で場所を変えられます）"
            $row.TotalText = "－"
        } else {
            $row.TotalText = "{0:#,0} 件" -f $item.ファイル数
            # 0 件の内訳は出さない（ふだんは「新規」「更新あり」だけになる）
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($pair in @(
                    @("新規", $item.新規),
                    @("更新あり", $item.更新あり),
                    @("前回未完了", $item.前回未完了),
                    @("変換結果が無い・壊れている", $item.変換結果なし),
                    @("前回失敗", $item.前回失敗))) {
                if ($pair[1] -gt 0) {
                    $parts.Add("$($pair[0]) $('{0:#,0}' -f $pair[1]) 件")
                }
            }
            if ($item.変換対象 -gt 0) {
                $row.TargetText = "{0:#,0} 件" -f $item.変換対象
                $row.Tone = "info"
            } else {
                $row.TargetText = "更新不要"
                $row.Tone = "ok"
            }
            $row.DetailText = if ($parts.Count -gt 0) { $parts -join " / " } else { "すべて変換済みです" }
        }
        $rows.Add($row)
    }
    return , $rows.ToArray()
}

function getConvertConfirmText {
    # 「失敗分も再変換する」のチェックに合わせた、合計の文言と主ボタンの文言
    param (
        [int]$targets,      # 変換対象の件数
        [int]$failed,       # 前回失敗した件数
        [bool]$retryFailed  # 失敗分も再変換するか
    )

    $total = $targets
    if ($retryFailed) {
        $total += $failed
    }
    if ($total -gt 0) {
        return @{ Total = $total; Text = "合計 {0:#,0} 件を変換します。" -f $total; Button = "変換を開始" }
    }
    return @{ Total = 0; Text = "更新が必要なファイルはありません（すべて変換済みです）。"; Button = "閉じる" }
}