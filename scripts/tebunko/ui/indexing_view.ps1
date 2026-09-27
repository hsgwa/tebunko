# インデックス作成の確認ダイアログ・進み具合に出す文言の決定。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\indexing_view.Tests.ps1）。
#
# 色は「意味」（Tone）で返し、実際の色は画面側（indexing_tab.ps1）で対応表から引く。
#   info = これから取り込む / ok = 取り込みの必要なし / warn = 注意 / ng = 取り込めない / gray = 対象外

function newPlanViewRows {
    # 取り込み予定（取り込み予定.tsv の行）を、確認のダイアログに出す形にする
    param (
        $plan  # readIngestPlan の結果
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($plan)) {
        # 空の配列を渡すと @($plan) に $null が 1 つ入るため、ここで外す
        if ($null -eq $item) { continue }
        $row = @{ Name = $item.インデックス名; Path = $item.元のフォルダ; Tone = "info" }
        if ($item.区分 -eq ${planKindUnchecked}) {
            $row.TargetText = "取り込みません"
            $row.Tone = "gray"
            $row.DetailText = "［作成］のチェックが外れています（インデックスはそのまま残します）"
            $row.TotalText = "－"
        } elseif ($item.区分 -eq ${planKindMissing}) {
            $row.TargetText = "取り込めません"
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
                    @("インデックスが無い・壊れている", $item.インデックスなし),
                    @("前回失敗", $item.前回失敗))) {
                if ($pair[1] -gt 0) {
                    $parts.Add("$($pair[0]) $('{0:#,0}' -f $pair[1]) 件")
                }
            }
            if ($item.取り込み対象 -gt 0) {
                $row.TargetText = "{0:#,0} 件" -f $item.取り込み対象
                $row.Tone = "info"
            } else {
                $row.TargetText = "更新不要"
                $row.Tone = "ok"
            }
            $row.DetailText = if ($parts.Count -gt 0) { $parts -join " / " } else { "すべて取り込み済みです" }
        }
        $rows.Add($row)
    }
    return , $rows.ToArray()
}

function getIndexingEndText {
    # インデックス作成が完了した（終了コード 0）ときの、進み具合の見出しと説明（@{ Text; Detail }）を返す。
    # 後回し（利用者のPowerPointが起動していて取り込まなかったファイル）がある場合の文言もここで決める
    param (
        [int]$success,
        [int]$failed,
        [int]$postponed,
        [string]$notice = ""
    )

    $processed = $success + $failed
    if ($processed -gt 0) {
        $counts = "成功 ${success} 件 / 失敗 ${failed} 件"
        if ($postponed -gt 0) {
            $counts += " / 残り ${postponed} 件"
        }
        $parts = New-Object System.Collections.Generic.List[string]
        if ($failed -gt 0) {
            $parts.Add("失敗したファイルと原因は「取り込みに失敗したファイル」の一覧で確認できます。")
        }
        if ($notice) {
            $parts.Add($notice)
        }
        return @{ Text = "インデックス作成が終わりました（${counts}）"; Detail = ($parts -join " ") }
    }
    if ($postponed -gt 0) {
        return @{ Text = "インデックス作成が終わりました（取り込まずに残したファイル ${postponed} 件）"; Detail = $notice }
    }
    return @{ Text = "取り込みが必要なファイルはありませんでした"; Detail = "" }
}

function getIndexingStateText {
    # ボタンの上の一言（取り込み一覧の「未取り込み」の残り件数から。中断でも後回しでも同じ状態のため見分けない）。
    # インデックス作成中は出さない（そのときは別の一言をボタンの下に出す）
    param (
        [int]$pending,
        [bool]$indexing
    )

    if ($pending -gt 0 -and !$indexing) {
        return "⏸ まだ取り込んでいないファイルがあります（残り ${pending} 件）"
    }
    return ""
}

function getIndexingConfirmText {
    # 「失敗分も再取り込みする」のチェックに合わせた、合計の文言と主ボタンの文言
    param (
        [int]$targets,      # 取り込み対象の件数
        [int]$failed,       # 前回失敗した件数
        [bool]$retryFailed  # 失敗分も再取り込みするか
    )

    $total = $targets
    if ($retryFailed) {
        $total += $failed
    }
    if ($total -gt 0) {
        return @{ Total = $total; Text = "合計 {0:#,0} 件を取り込みます。" -f $total; Button = "インデックス作成を開始" }
    }
    return @{ Total = 0; Text = "更新が必要なファイルはありません（すべて取り込み済みです）。"; Button = "閉じる" }
}