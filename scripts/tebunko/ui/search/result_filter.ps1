# 結果の一覧の絞り込み（件数の行の「絞り込み」）と、空の状態の［インデックス管理へ］ボタンの画面層。

# ---- 結果の絞り込み ----

function applyFilter {
    $script:filterText = $ui.FilterBox.Text.Trim()
    applyResultFilter $script:filterText
    if ($script:lastSearch -and !$script:search -and $script:hitCount -gt 0) {
        $shown = getShownHitCount
        if ($script:filterText -eq "") {
            finishSummaryText
        } else {
            $ui.SummaryText.Text = getFilteredSummaryText $script:hitCount $shown
        }
    }
}

function finishSummaryText {
    # 上限を超えて総数まで数え終えた検索は「全 N 件中…」に戻す（秒は付けない）
    $shared = $script:lastSearch.Shared
    $ui.SummaryText.Text = getSearchFinishSummaryText $script:hitCount $script:fileGroups.Count -1 ([bool]$shared.Truncated) ([int]$shared.HitTotal)
}

$script:filterTimer = newTimer 300 {
    $script:filterTimer.Stop()
    safe { applyFilter }
}

# ---- イベント ----

$ui.GoIndexTabButton.Add_Click({ selectScreen "IndexTab" })
$ui.FilterBox.Add_TextChanged({
    $ui.FilterPlaceholder.Visibility = if ($ui.FilterBox.Text -eq "") { "Visible" } else { "Collapsed" }
    $script:filterTimer.Stop()
    $script:filterTimer.Start()
})
