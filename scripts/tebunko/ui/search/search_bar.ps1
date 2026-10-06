# 検索バー（検索ワード・種類のチップ・探す範囲・検索条件・検索ボタン・高速検索の状態）の画面層。
# 文言と可否の判断は search_bar_view.ps1（テストあり）。検索の実行は search_session.ps1。

# 種類のチップ（${fileKindNames} の値 → ToggleButton）
$script:kindChips = [ordered]@{
    excel = $ui.KindChipExcel; word = $ui.KindChipWord; powerpoint = $ui.KindChipPowerPoint; text = $ui.KindChipText
}

# 検索ワード（前後の空白を除く）
function getWordText {
    return $ui.WordBox.Text.Trim()
}

function getFileKindsFromUi {
    # チップで選ばれている種類（${fileKindNames} の順）
    return @(${fileKindNames} | Where-Object { $script:kindChips[$_].IsChecked })
}

function setFileKindsToUi {
    param (
        [object[]]$kinds
    )

    foreach ($kind in ${fileKindNames}) {
        $script:kindChips[$kind].IsChecked = (@($kinds) -contains $kind)
    }
}

function getSearchOptionFromUi {
    # 画面の検索条件を readSearchOption と同じ形（と、選んだ種類 FileKinds）で返す
    return @{
        UseRegex        = [bool]$ui.RegexCheck.IsChecked
        CaseSensitive   = [bool]$ui.CaseCheck.IsChecked
        FileKinds       = @(getFileKindsFromUi)
        IncludeShapes   = [bool]$ui.ShapeCheck.IsChecked
        IncludeComments = [bool]$ui.CommentCheck.IsChecked
    }
}

function setSearchOptionToUi {
    param (
        [hashtable]$option
    )

    $ui.RegexCheck.IsChecked = [bool]$option.UseRegex
    $ui.CaseCheck.IsChecked = [bool]$option.CaseSensitive
    $ui.ShapeCheck.IsChecked = [bool]$option.IncludeShapes
    $ui.CommentCheck.IsChecked = [bool]$option.IncludeComments
    setFileKindsToUi (readFileKinds)
}

function updateFastSearchView {
    # 高速検索の使用可否（ワード・［正規表現を使う］を変えたらすぐ、Windows Search が使えるかは確かめたときに変わる）
    $ui.FastSearchText.Text = (getFastSearchView $script:fastAvailable ([bool]$ui.RegexCheck.IsChecked) (getWordText)).Text
}

function checkFastSearchAvailable {
    # Windows Search が使えるか（system_index が索引の対象か）を別スレッドで確かめる（画面を固めないように）
    startJob {
        param ($systemRoot)
        @{ Root = $systemRoot; Available = (testWindowsSearch $systemRoot) }
    } @($workspace.SystemIndexDir) {
        param ($output, $errorText)
        $result = if (!$errorText -and $output.Count -gt 0) { $output[0] } else { $null }
        if ($result -and $result.Root -ne $workspace.SystemIndexDir) {
            # 確かめている間にワークスペースを変えた（切り替えたときに確かめ直している）
            return
        }
        $script:fastAvailable = if ($result) { [bool]$result.Available } else { $false }
        updateFastSearchView
    }
}

function updateWordNotice {
    updateFastSearchView
    $notice = getWordNotice (getWordText) ([bool]$ui.RegexCheck.IsChecked)
    if ($notice -ne "") {
        $ui.WordNotice.Text = $notice
        $ui.WordNotice.Visibility = "Visible"
    } else {
        $ui.WordNotice.Visibility = "Collapsed"
    }
    updateSearchButton
}

function updateSearchButton {
    $noIndex = $script:indexSummary -and $script:indexSummary["Count"] -eq 0
    $state = newSearchButtonState ([bool]$script:search) ([bool]($script:search -and $script:search.Shared.Stop)) `
        (getWordText) (!$noIndex) @(getSearchTargets).Count
    $ui.SearchButton.Content = $state.Content
    $ui.SearchButton.IsEnabled = $state.Enabled
}

function updateSearchTarget {
    $targets = @(getSearchTargets)
    $summary = $script:indexSummary
    if ($summary -and $summary["Count"] -eq 0) {
        $ui.SearchTargetText.Text = (getNoIndexTargetText)
    } elseif ($targets.Count -eq 0) {
        $ui.SearchTargetText.Text = "検索対象：なし（左の一覧で、検索するインデックス・フォルダにチェックを付けてください）"
    } elseif (!(isAllIndexChecked)) {
        $ui.SearchTargetText.Text = "検索対象：$(describeSearchTargets $targets)"
    } elseif ($null -eq $summary) {
        $ui.SearchTargetText.Text = "検索対象：すべて（確認中…）"
    } else {
        $ui.SearchTargetText.Text = "検索対象：すべて（集約ファイル $($summary['Count'].ToString('N0')) 件 ・ 最終取り込み $(formatTime $summary['LastWrite'])）"
    }
    $ui.SearchTargetText.ToolTip = $ui.SearchTargetText.Text
    # インデックスが無いときは、結果の表の代わりに空の状態（［インデックス管理へ］）を出す
    $ui.ResultEmptyState.Visibility = if ($summary -and $summary["Count"] -eq 0) { "Visible" } else { "Collapsed" }
    updateSearchButton
}

# ---- イベント ----

$ui.WordBox.Add_TextChanged({ safe { updateWordNotice } })
$ui.WordBox.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe {
            if (!$script:search) {
                startSearch
            }
        }
        $e.Handled = $true
    }
})
$ui.SearchButton.Add_Click({ safe { startSearch } })
$ui.RegexCheck.Add_Click({
    safe {
        writeSearchOption @{ UseRegex = [bool]$ui.RegexCheck.IsChecked }
        updateWordNotice
    }
})
$ui.CaseCheck.Add_Click({ safe { writeSearchOption @{ CaseSensitive = [bool]$ui.CaseCheck.IsChecked } } })
$ui.ShapeCheck.Add_Click({ safe { writeSearchOption @{ IncludeShapes = [bool]$ui.ShapeCheck.IsChecked } } })
$ui.CommentCheck.Add_Click({ safe { writeSearchOption @{ IncludeComments = [bool]$ui.CommentCheck.IsChecked } } })

# ［探す範囲 ▾］は、押したらボタンの下にメニューを開く
$ui.ScopeButton.Add_Click({
    $menu = $ui.ScopeButton.ContextMenu
    $menu.PlacementTarget = $ui.ScopeButton
    $menu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
    $menu.IsOpen = $true
})

# 種類のチップ。押すと ToggleButton が先に入れ替わるため、押す前の状態に戻して、判断（toggleFileKind）の結果を反映する
foreach ($kind in ${fileKindNames}) {
    $chip = $script:kindChips[$kind]
    $chip.Add_Click({
        param ($sender, $e)
        safe {
            $clicked = [string]$sender.Tag
            $before = @(${fileKindNames} | Where-Object { ($script:kindChips[$_].IsChecked) -xor ($_ -eq $clicked) })
            $after = @(toggleFileKind $before $clicked)
            setFileKindsToUi $after
            writeFileKinds $after
        }
    })
}
