# 検索バー（検索ワード・種類のチップ・ファイル内の対象・検索条件・検索ボタン・高速検索の状態）の画面層。
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
    updateScopeButton
    setFileKindsToUi (readFileKinds)
}

function updateFastSearchView {
    # 高速検索の使用可否（ワード・［正規表現］を変えたらすぐ、Windows Search が使えるかは確かめたときに変わる）。
    # 使えるときは緑、使えないときは灰色。ⓘ は常に出し、ツールチップで印の意味（使えないときは理由と直し方）を知らせる。
    # 印の幅は XAML で固定してあり、ここでは変えない
    $view = getFastSearchView $script:fastAvailable ([bool]$ui.RegexCheck.IsChecked) (getWordText)
    $ui.FastSearchText.Text = $view.Text
    $ui.FastBadge.Tag = if ($view.Usable) { "ok" } else { "off" }
    $ui.FastBadge.ToolTip = $view.Tip
}

function setSearchKindBanner {
    # 種類が 1 つも選ばれていないことを知らせる帯（getSearchKindBannerText の結果。$null なら隠す）
    param (
        [string]$text
    )

    $ui.SearchKindBanner.Visibility = if ($text) { "Visible" } else { "Collapsed" }
    $ui.SearchKindBannerText.Text = $text
}

function updateScopeButton {
    # ［ファイル内の対象］の文言と、既定から変えているときの線の色
    $view = getScopeButtonText ([bool]$ui.ShapeCheck.IsChecked) ([bool]$ui.CommentCheck.IsChecked)
    $ui.ScopeButton.Content = $view.Text
    $ui.ScopeButton.Tag = if ($view.Changed) { "changed" } else { $null }
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
    $ui.WordPlaceholder.Visibility = if ($ui.WordBox.Text -eq "") { "Visible" } else { "Collapsed" }
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
    # 左の欄の見出し（選んだ数 / 全部の数）、インデックスが無いときの結果欄の案内。
    # 検索対象の詳しい中身（先頭の数件・集約ファイルの数・最終更新）は、見出しのツールチップに出す
    $targets = @(getSearchTargets)
    $summary = $script:indexSummary
    $total = @($script:indexRoots).Count
    $checked = @($script:indexRoots | Where-Object { $_.IsChecked -eq $true }).Count
    $ui.TargetCountText.Text = getTargetCountText $checked $total
    $ui.TargetCountText.ToolTip = if ($total -eq 0) {
        "［インデックス管理］で作ったインデックスの一覧です"
    } elseif ($targets.Count -eq 0) {
        "検索対象：なし"
    } elseif (!(isAllIndexChecked)) {
        "検索対象：$(describeSearchTargets $targets)"
    } elseif ($null -eq $summary) {
        "検索対象：すべて（確認中…）"
    } else {
        "検索対象：すべて（集約ファイル $($summary['Count'].ToString('N0')) 件 ・ 最終更新 $(formatTime $summary['LastWrite'])）"
    }
    # インデックスが無いときは、結果の表の代わりに空の状態（［インデックス管理を開く］）を出す
    $noIndex = [bool]($summary -and $summary["Count"] -eq 0)
    $ui.ResultEmptyState.Visibility = if ($noIndex) { "Visible" } else { "Collapsed" }
    # 検索できるインデックスが無いときは、高速検索の状態も出さない
    $ui.FastBadge.Visibility = if ($noIndex) { "Collapsed" } else { "Visible" }
    $ui.FastSearchSlot.Visibility = $ui.FastBadge.Visibility
    updateConditionFlow
    updateSearchButton
}

function updateConditionFlow {
    # 検索条件の行の折り返し。判断（getConditionFlow）の結果を、「伸びる空き」の幅に反映する。
    # 並びの幅が変わったとき・高速検索の印を出す／隠したときに呼ぶ
    $panel = $ui.ConditionsPanel
    if ($panel.ActualWidth -le 0) { return }
    $items = @($panel.Children | Where-Object { $_ -ne $ui.ConditionsSpacer })
    # 幅が決まっている項目（Width 指定）は指定の値で数える。出した直後で ActualWidth がまだ 0 でも、正しい幅で計算できる
    $widths = @($items | ForEach-Object {
        if ($_.Visibility -eq "Collapsed") { 0.0 }
        else {
            $width = if ([double]::IsNaN($_.Width)) { $_.ActualWidth } else { $_.Width }
            [double]($width + $_.Margin.Left + $_.Margin.Right)
        }
    })
    # 空きは、［ファイル内の対象］の次（検索条件の前）に置く
    $spacerIndex = $items.IndexOf($ui.CaseCheck)
    $flow = getConditionFlow $widths $spacerIndex $panel.ActualWidth
    if ([Math]::Abs($ui.ConditionsSpacer.Width - $flow.SpacerWidth) -gt 0.1) {
        $ui.ConditionsSpacer.Width = $flow.SpacerWidth
    }
}

# ---- イベント ----
$ui.ConditionsPanel.Add_SizeChanged({ safe { updateConditionFlow } })


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
$ui.ShapeCheck.Add_Click({ safe { writeSearchOption @{ IncludeShapes = [bool]$ui.ShapeCheck.IsChecked }; updateScopeButton } })
$ui.CommentCheck.Add_Click({ safe { writeSearchOption @{ IncludeComments = [bool]$ui.CommentCheck.IsChecked }; updateScopeButton } })

# ［ファイル内の対象 ▾］は、押したらボタンの下にメニューを開く
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
            # 帯を出している間は、選び直したらすぐ消す
            if ($ui.SearchKindBanner.Visibility -eq "Visible") {
                setSearchKindBanner (getSearchKindBannerText $after)
            }
        }
    })
}
