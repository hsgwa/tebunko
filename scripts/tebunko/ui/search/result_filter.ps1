# 結果の一覧の絞り込み（件数の行の「絞り込み」）と、空の状態の［インデックス管理へ］ボタンの画面層。

# ---- 結果の絞り込み ----

function applyFilter {
    # 絞り込みを当てる。表は見出しの数だけ回して作り直すので（applyResultFilter。予算ぶんずつ回す）、要約欄の文は回し終えてから出す
    $script:filterText = $ui.FilterBox.Text.Trim()
    applyResultFilter $script:filterText { showFilterSummary }
}

function showFilterSummary {
    # 検索が終わっていれば、要約欄を絞り込みに合う件数の文にする（絞り込みを外したら、検索の終わりの文に戻す）
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
