# 結果の一覧の絞り込み（件数の行の「絞り込み」）と、空の状態の［インデックス管理へ］の画面層。

# ---- 結果の絞り込み ----

function applyFilter {
    $script:filterText = $ui.FilterBox.Text.Trim()
    applyResultFilter $script:filterText
    if ($script:lastSearch -and !$script:search -and $script:hitCount -gt 0) {
        $shown = getShownHitCount
        if ($script:filterText -eq "") {
            finishSummaryText
        } else {
            $ui.SummaryText.Text = "$($script:hitCount.ToString('N0')) 件中 $($shown.ToString('N0')) 件を表示"
        }
    }
}

function finishSummaryText {
    $ui.SummaryText.Text = "$($script:hitCount.ToString('N0')) 件（$($script:fileGroups.Count.ToString('N0')) ファイル）"
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
