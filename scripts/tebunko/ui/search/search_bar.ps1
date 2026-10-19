# 検索バー（検索ワード・種類のチップ・ファイル内の対象・検索条件・検索ボタン）の画面層。
# 文言と可否の判断は search_bar_view.ps1（テストあり）。検索の実行は search_session.ps1。

# 種類のチップ（${fileKindNames} の値 → ToggleButton）
$script:kindChips = [ordered]@{
    excel = $ui.KindChipExcel; word = $ui.KindChipWord; powerpoint = $ui.KindChipPowerPoint; text = $ui.KindChipText; pdf = $ui.KindChipPdf
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

function placeSearchBalloon {
    # 吹き出しを、基準の項目の下（5 下）に、左端をそろえて置く。右端がはみ出すときは左へ寄せる（判断は getBalloonLeft）。
    # 吹き出しは高さ 0 の BalloonLayer（Canvas）に置いてあり、検索バーの高さは変わらない
    param (
        $balloon,
        $anchor
    )

    if ($balloon.Visibility -ne "Visible") { return }
    $layer = $ui.BalloonLayer
    try {
        $ui.ConditionsPanel.UpdateLayout()
        $balloon.Measure([System.Windows.Size]::new([double]::PositiveInfinity, [double]::PositiveInfinity))
        $origin = $anchor.TransformToVisual($layer).Transform([System.Windows.Point]::new(0, 0))
    } catch {
        return
    }
    $left = getBalloonLeft $origin.X $balloon.DesiredSize.Width $layer.ActualWidth
    [System.Windows.Controls.Canvas]::SetLeft($balloon, $left)
    [System.Windows.Controls.Canvas]::SetTop($balloon, $origin.Y + $anchor.ActualHeight + 5)
}

function placeSearchBalloons {
    placeSearchBalloon $ui.SearchKindBalloon $ui.KindChipExcel
    placeSearchBalloon $ui.RegexBalloon $ui.RegexCheck
}

function setSearchBalloon {
    # 吹き出しを出す（text が空なら隠す）。新しく出したとき、画面の読み上げに知らせる
    param (
        $balloon,
        $textBlock,
        [string]$text
    )

    if (!$text) {
        $balloon.Visibility = "Collapsed"
        return
    }
    $wasVisible = ($balloon.Visibility -eq "Visible") -and ($textBlock.Text -eq $text)
    # 画面の読み上げに、出たときに読み上げるよう頼む（.NET 4.7.1 以降。無い環境では何もしない）
    try { [System.Windows.Automation.AutomationProperties]::SetLiveSetting($textBlock, [System.Windows.Automation.AutomationLiveSetting]::Assertive) } catch { }
    $textBlock.Text = $text
    $balloon.Visibility = "Visible"
    placeSearchBalloons
    if (!$wasVisible) {
        try {
            $peer = [System.Windows.Automation.Peers.UIElementAutomationPeer]::CreatePeerForElement($textBlock)
            if ($peer) { $peer.RaiseAutomationEvent([System.Windows.Automation.Peers.AutomationEvents]::LiveRegionChanged) }
        } catch { }
    }
}

function setSearchKindBalloon {
    # 種類が 1 つも選ばれていないことを知らせる吹き出し（getSearchKindBalloonText の結果。$null なら隠す）
    param (
        [string]$text
    )

    setSearchBalloon $ui.SearchKindBalloon $ui.SearchKindBalloonText $text
}

function updateScopeButton {
    # ［ファイル内の対象］の文言と、既定から変えているときの線の色
    $view = getScopeButtonText ([bool]$ui.ShapeCheck.IsChecked) ([bool]$ui.CommentCheck.IsChecked)
    $ui.ScopeButton.Content = $view.Text
    $ui.ScopeButton.Tag = if ($view.Changed) { "changed" } else { $null }
}

function getCurrentWordNotice {
    # いまの検索ワード・［正規表現］の状態での、正規表現の注意（正しくなければ文言、正しければ ""）
    getWordNotice (getWordText) ([bool]$ui.RegexCheck.IsChecked)
}

function updateWordNotice {
    $ui.WordPlaceholder.Visibility = if ($ui.WordBox.Text -eq "") { "Visible" } else { "Collapsed" }
    # 正規表現が正しくないときは、入力欄の枠を赤くし、［正規表現］の下に吹き出しを出す
    $notice = getCurrentWordNotice
    $ui.WordErrorRing.Visibility = if ($notice -ne "") { "Visible" } else { "Collapsed" }
    setSearchBalloon $ui.RegexBalloon $ui.RegexBalloonText $notice
    updateSearchButton
}

function updateSearchButton {
    $notice = getCurrentWordNotice
    $noIndex = $script:indexSummary -and $script:indexSummary["Count"] -eq 0
    $state = newSearchButtonState ([bool]$script:search) ([bool]($script:search -and $script:search.Shared.Stop)) `
        (getWordText) (!$noIndex) @(getSearchTargets).Count ($notice -ne "") ([bool]$script:indexTreeLoading)
    $ui.SearchButton.Content = $state.Content
    $ui.SearchButton.IsEnabled = $state.Enabled
}

function updateSearchTarget {
    # 左の欄の見出し（選んだ数 / 全部の数）、インデックスが無いときの結果欄の案内。
    # 検索対象の詳しい中身（先頭の数件・本文インデックスのファイルの数・最終更新）は、見出しのツールチップに出す
    $targets = @(getSearchTargets)
    $summary = $script:indexSummary
    $total = @($script:indexRoots).Count
    $checked = @($script:indexRoots | Where-Object { $_.IsChecked -eq $true }).Count
    $ui.TargetCountText.Text = getTargetCountText $checked $total
    $targetsText = if ($targets.Count -gt 0) { describeSearchTargets $targets } else { "" }
    $ui.TargetCountText.ToolTip = getSearchTargetText ([bool]$script:indexTreeLoading) ([string]$script:indexTreeConnectError) `
        $total $targets.Count (isAllIndexChecked) $targetsText $summary ([int]$script:indexTreeOldCount)
    # インデックスが無いときは、結果の表の代わりに空の状態（［インデックス管理を開く］）を出す。
    # 読み込み中・接続できないときは、「インデックスが無い」とは違うため出さない
    $noIndex = [bool]($summary -and $summary["Count"] -eq 0 -and !$script:indexTreeLoading -and !$script:indexTreeConnectError)
    $ui.ResultEmptyState.Visibility = if ($noIndex) { "Visible" } else { "Collapsed" }
    updateConditionFlow
    updateSearchButton
}

function updateConditionFlow {
    # 検索条件の行の折り返し。判断（getConditionFlow）の結果を、「伸びる空き」の幅に反映する。
    # 並びの幅が変わったときに呼ぶ
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
    # 空きの幅が変わると、並びの大きさが同じでも［正規表現］などが横に動く。吹き出しもここで置き直す
    placeSearchBalloons
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

# ［ファイル内の対象 ▼］は、押したらボタンの下にメニューを開く
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
            # 吹き出しを出している間は、選び直したらすぐ消す
            if ($ui.SearchKindBalloon.Visibility -eq "Visible") {
                setSearchKindBalloon (getSearchKindBalloonText $after)
            }
        }
    })
}
