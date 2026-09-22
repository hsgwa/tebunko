# ［1 検索］（2 段目は［インデックス管理］［検索］）の、ウィンドウ全体にかかわる処理。
# 画面の起動口（..\..\tebunko\gui.ps1）が、$ui と ${grepDir} を作った後に読み込み、ウィンドウのイベントから呼ぶ。

${searchLimit} = 10000
# 選択行のプレビューに出す行数は、プレビューの高さ（ドラッグで変わる）に収まるだけ出す（getPreviewContextLines）
${previewRowHeight}     = 22   # プレビューの 1 行の高さの目安。高さから出せる行数を求めるのに使う
${previewScrollBarSize} = 18   # 横スクロールバーの高さの目安（ViewportHeight が取れないときに引く）
${maxPreviewRows}       = 101  # プレビューに出す行数の上限（選択行＋前後 50 行）
${libPath}  = "${grepDir}\lib.ps1"  # 別スレッドで読み込む（startJob に渡す）
# インデクサ（ウィンドウを出さずに別プロセスで起動する。indexing_tab.ps1）
${indexerScriptPath} = "${grepDir}\indexer.ps1"

function selectGrepTab {
    # ［1 検索］の中のタブ（$ui.IndexTab / $ui.SearchTab）を選ぶ。1 段目も［1 検索］にする
    param (
        $tab
    )

    $ui.Tabs.SelectedItem = $ui.SearchPage
    $ui.SearchTabs.SelectedItem = $tab
}

function onGrepPageShown {
    # ［1 検索］に切り替えたとき・2 段目を切り替えたとき
    if ($ui.SearchTabs.SelectedItem -eq $ui.IndexTab) {
        refreshIndexingState
    }
}

$ui.SearchTabs.Add_SelectionChanged({
    param ($sender, $e)
    # 中の表・一覧の選択変更も伝わってくるため、タブの切り替えだけを扱う
    if ($e.OriginalSource -ne $ui.SearchTabs) {
        return
    }
    safe { onGrepPageShown }
})

function onGrepActivated {
    # ウィンドウが前面に来たとき
    # クロール対象フォルダがほかの画面で変更されていれば読み直す
    if ((getTargetsKey @(getTargetFolders)) -ne $script:savedTargets) {
        loadTargets
        setStatus "インデックス一覧がほかで変更されたため、読み直しました"
    }
    # フォルダの有無は別スレッドで調べる（届かないネットワークのフォルダで画面が固まらないように）
    refreshFolderStatus
    if (!(isIndexing)) {
        refreshIndexingState
    }
    updateSearchTarget
}

function onGrepKeyDown {
    # ［1 検索］のキー操作。扱ったら $true を返す
    param (
        $e,
        $modifiers
    )

    if ($e.Key -eq "F" -and $modifiers -eq "Control") {
        # どのタブからでも［検索］へ移る
        selectGrepTab $ui.SearchTab
        $ui.WordBox.Focus() | Out-Null
        $ui.WordBox.SelectAll()
        return $true
    }
    if ($e.Key -eq "F" -and $modifiers -eq ([System.Windows.Input.ModifierKeys]::Control -bor [System.Windows.Input.ModifierKeys]::Shift)) {
        selectGrepTab $ui.SearchTab
        $ui.FilterBox.Focus() | Out-Null
        return $true
    }
    if ($ui.Tabs.SelectedItem -ne $ui.SearchPage) {
        return $false
    }
    if ($e.Key -eq "F5") {
        safe {
            refreshIndexingState
            refreshIndexSummary
            loadIndexTree
        }
        return $true
    }
    if ($e.Key -eq "Escape" -and $script:search) {
        cancelSearch
        return $true
    }
    return $false
}

function confirmGrepClosing {
    # ウィンドウを閉じる前。閉じてよければ $true
    # インデックス作成はウィンドウを出さずに動いているため、閉じる前にどうするか聞く
    if (isIndexing) {
        $answer = showConfirm `
            -heading "まだインデックス作成の途中です。どうしますか？" `
            -choices @(
                @{ Text = "インデックス作成を続けたまま閉じる"; Detail = "インデックス作成は裏で続きます。もう一度開くと進み具合が出ます"; Value = "keep" },
                @{ Text = "インデックス作成を止めてから閉じる"; Detail = "いま取り込んでいるファイルが終わったところで止まります（次に開いたとき続きから再開できます）"; Value = "stop" }
            ) `
            -cancelText "閉じない"
        if ($null -eq $answer) {
            return $false
        }
        if ($answer -eq "stop") {
            [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
        }
    }
    if ($script:search) {
        $script:search.Shared.Stop = $true
    }
    return $true
}

function closeGrepPage {
    # ウィンドウを閉じた後の後始末
    if ($script:search) {
        $script:search.PS.Stop()
    }
}

function startGrepPage {
    # 起動時。設定を読み、インデックス作成が続いていれば進み具合を表示し、2 段目のタブを決める
    loadTargets
    setSearchOptionToUi (readSearchOption)
    setOpenMode (readOpenMode)
    updateOpenMenu
    refreshIndexingState
    loadIndexTree
    updateWordNotice
    refreshIndexSummary

    # 前回の画面で起動したインデックス作成が続いていれば、進み具合を表示する
    $runningIndexing = findRunningIndexer
    if ($runningIndexing) {
        adoptIndexing $runningIndexing
    }

    # 2 段目：インデックス作成中・中断中、またはインデックスが無ければ［インデックス管理］、それ以外は［検索］
    $openIndexTab = $runningIndexing -or ($script:indexingState -and $script:indexingState.Pending -gt 0) -or !(testIndexExists)
    $ui.SearchTabs.SelectedItem = if ($openIndexTab) { $ui.IndexTab } else { $ui.SearchTab }
}
