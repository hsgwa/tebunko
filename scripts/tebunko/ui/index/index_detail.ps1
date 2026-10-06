# インデックス管理の画面の詳細（インデックス作成の開始ボタン・状態・失敗の一覧・状態の合計）。

# 詳細の左右の部品。名前の一覧（gui_main.ps1）に足さず、読み込んだ中身から取る
foreach ($detailName in @(
        "IndexDetailBody", "IndexDetailName", "IndexDetailPath", "IndexDetailFolderStatus",
        "IndexDetailBadge", "IndexDetailBadgeText", "IndexDetailBadgeSub", "IndexDetailUpdated", "IndexDetailCount",
        "IndexDetailFastBadge", "IndexDetailFastBadgeText", "IndexDetailFastReason", "IndexDetailFastChecked", "IndexDetailFastNote")) {
    $ui[$detailName] = $ui.IndexDetailHost.Content.FindName($detailName)
}

function setIndexBadge {
    # 詳細の状態のバッジの文言と色（一覧の LevelBadge と同じ組み合わせ。level は Ok / Wait / Ng / Run / None）
    param ($border, $textBlock, [string]$text, [string]$level)

    $kind = if ($level -in @("Ok", "Wait", "Ng", "Run")) { $level } else { "None" }
    $border.Background = themeBrush "Badge.${kind}.Bg"
    $border.BorderBrush = themeBrush "Badge.${kind}.Bg"
    $textBlock.Foreground = themeBrush "Badge.${kind}.Text"
    $textBlock.Text = $text
    $border.Visibility = if ($text) { "Visible" } else { "Collapsed" }
}

function setIndexingBanner {
    # 画面の上の帯の色と種類。level は info（更新中）/ warn（中断・中止）/ ok（終わった）/ ng（できなかった）。
    # kind は帯の中身の持ち主（running・interrupted・done）。中断の帯だけは、状態が変わったとき自分で消す
    param ([string]$level, [string]$kind)

    $colors = @{
        info = @("Accent.Soft", "Accent.Ring"); warn = @("Warn.Note", "Warn.Line")
        ok = @("Badge.Ok.Bg", "Border.Soft"); ng = @("Badge.Ng.Bg", "Border.Soft")
    }
    $pair = $colors[$level]
    if ($null -eq $pair) { $pair = $colors.info }
    $ui.IndexingProgressPanel.Background = themeBrush $pair[0]
    $ui.IndexingProgressPanel.BorderBrush = themeBrush $pair[1]
    $script:indexingBannerKind = $kind
}

function updateIndexingResume {
    # 中断した更新（残りがあり、いま更新していない）の帯。［続きから再開］を出す。
    # 中断の帯は、残りが無くなったら消す（更新が終わった・中止した直後の帯は消さない）
    $pending = if ($script:indexingState) { [int]$script:indexingState.Pending } else { 0 }
    $canResume = ($pending -gt 0) -and !(isIndexing)
    $ui.IndexingResumeButton.Visibility = if ($canResume) { "Visible" } else { "Collapsed" }
    if ($canResume -and $ui.IndexingProgressPanel.Visibility -ne "Visible") {
        setIndexingBanner "warn" "interrupted"
        $ui.IndexingProgressText.Text = getIndexingStateText $pending $false
        $ui.IndexingProgressEta.Text = ""
        $ui.IndexingProgressDetail.Text = ""
        $ui.IndexingProgress.Visibility = "Collapsed"
        $ui.IndexingStopButton.Visibility = "Collapsed"
        $ui.IndexingLogButton.Visibility = "Collapsed"
        $ui.IndexingProgressPanel.Visibility = "Visible"
    } elseif ($canResume -and $script:indexingBannerKind -eq "interrupted") {
        $ui.IndexingProgressText.Text = getIndexingStateText $pending $false
    } elseif (!$canResume -and $script:indexingBannerKind -eq "interrupted") {
        $ui.IndexingProgressPanel.Visibility = "Collapsed"
        $script:indexingBannerKind = ""
    }
}

function updateIndexingButton {
    $ready = $false
    foreach ($item in $script:targetItems) {
        # フォルダの有無は refreshFolderStatus が別スレッドで調べた結果を使う（調べ終えるまではあるものとする）
        if ($item.Enabled -and (!$item.StatusChecked -or $item.FolderExists)) {
            $ready = $true
            break
        }
    }

    $state = $script:indexingState
    if ($script:indexBusy -or $script:archiveBusy) {
        # インデックスの削除中・エクスポート・インポート中（別スレッド）は、インデックス作成もインデックスの操作も始めない
        $ready = $false
    }
    if (isIndexing) {
        $ui.IndexingButton.Content = "更新中…"
        $ui.IndexingButton.IsEnabled = $false
    } else {
        $ui.IndexingButton.Content = "すべて更新"
        $ui.IndexingButton.IsEnabled = $ready
    }

    # ボタンの下の一言は、押せないときの理由と、前回の失敗があるときだけ出す。押すと何が起きるかは、ボタンのツールヒントに書く
    $ui.IndexingButton.ToolTip = if (isIndexing) {
        "更新中も検索できます。やめるときは［中止］を押してください（次に［すべて更新］を押すと続きから再開します）。"
    } else {
        "押すと、何件更新するかを確認してから、新しいファイル・変わったファイルだけを更新します。"
    }
    $hint = ""
    if (!$ready -and !(isIndexing)) {
        $hint = if ($script:targetItems.Count -eq 0) { "フォルダを追加すると更新できます" } else { "更新するインデックスにチェックを付けてください。" }
    } elseif ($state -and $state.Failed -gt 0 -and !(isIndexing)) {
        $hint = "前回うまく更新できなかったファイルがあります（押したあとで、もう一度ためすか選べます）。"
    }
    $ui.IndexingHint.Text = $hint
    $ui.IndexingHint.Visibility = if ($hint) { "Visible" } else { "Collapsed" }

    # インデックスの追加・編集・削除・エクスポート・インポートと、［設定］のワークスペースの［変更…］は互いに排他
    # （getIndexJobBlocker・getIndexTabButtonsEnabled。settings\settings.ps1 の testWorkspaceChangeable も同じ排他を見る）
    $selected = $null -ne $ui.IndexGrid.SelectedItem
    $blocker = getIndexJobBlocker (isIndexing) $script:indexBusy $script:archiveBusy
    $buttons = getIndexTabButtonsEnabled $blocker $selected
    $ui.NewIndexButton.IsEnabled = $buttons.New
    $ui.EditIndexButton.IsEnabled = $buttons.Edit
    $ui.RemoveIndexButton.IsEnabled = $buttons.Remove
    $ui.ExportIndexButton.IsEnabled = $buttons.Export
    $ui.ImportIndexButton.IsEnabled = $buttons.Import
}

function refreshIndexingState {
    # 取り込み一覧の集計は、ファイルが数万行になると数秒〜十数秒かかる。
    # 画面のスレッドで行うと、起動時・タブの切り替え時に画面が固まる（応答なしになる）ため別スレッドで数える
    if ($script:stateRunning) {
        $script:stateAgain = $true
        return
    }
    $script:stateRunning = $true
    $script:stateAgain = $false
    $script:stateJobPath = $workspace.StatusFile
    startJob {
        param ($path)
        getIndexingState -path $path
    } @($script:stateJobPath) {
        param ($output, $errorText)
        $script:stateRunning = $false
        if ($script:stateJobPath -ne $workspace.StatusFile) {
            # 集計している間にワークスペースを変えた。前のワークスペースの結果は出さず、読み直す
            $script:stateAgain = $true
        } elseif ($output -and $output.Count -gt 0 -and $output[0]) {
            # インデクサが書き込んでいる瞬間などは、次の機会に読み直す
            applyIndexingState $output[0]
        }
        if ($script:stateAgain) {
            refreshIndexingState
        }
    }
}

function getIndexingRatio {
    # 更新全体の進み（0〜1）。更新中でない・割合がまだ分からないときは負の値
    if ((isIndexing) -and !$ui.IndexingProgress.IsIndeterminate) {
        return [double]$ui.IndexingProgress.Value
    }
    return -1.0
}

function updateIndexTabBadge {
    # ナビの［インデックス管理］の横の印（更新中の割合・中断・失敗。文言は getIndexNavBadge）
    $state = $script:indexingState
    $pending = if ($state) { [int]$state.Pending } else { 0 }
    $failed = if ($state) { [int]$state.Failed } else { 0 }
    $ratio = getIndexingRatio
    $badge = getIndexNavBadge (isIndexing) $ratio $pending $failed
    $ui.IndexTabBadge.Text = $badge.Text
    $ui.IndexTabBadge.ToolTip = $badge.ToolTip
    $ui.IndexTabBadge.Foreground = if ($badge.Kind -eq "Run") { ${infoBrush} } else { ${warnBrush} }
    $ui.IndexTabBadge.Visibility = if ($badge.Text) { "Visible" } else { "Collapsed" }
}

function applyIndexingState {
    # 集計（別スレッド）の結果を画面に反映する
    param (
        $state  # getIndexingState の結果
    )

    $script:indexingState = $state

    # 失敗したファイルは下の一覧に原因とともに表示する
    applyIndexStats $state.IndexStats (isIndexing)
    updateFailedList $state
    updateIndexSummaryText
    updateIndexingButton
    updateIndexingResume
    updateIndexTabBadge
}

function updateFailedList {
    # 取り込みに失敗したファイルと原因（取り込み一覧のエラー列）を一覧に表示する
    param (
        $state  # getIndexingState の結果
    )

    $folderPaths = @{}  # インデックス名 → クロール対象フォルダ（大文字・小文字を区別しない）
    foreach ($folder in $state.Folders) {
        if ($folder.Name) {
            $folderPaths[$folder.Name] = $folder.Path
        }
    }

    $rows = New-Object System.Collections.ArrayList
    foreach ($status in $state.FailedRows) {
        $row = New-Object FailRow
        $row.RelPath = $status.相対パス
        $row.Reason = if ($status.エラー) { $status.エラー } else { "（原因は記録されていません）" }
        $ingested = [datetime]::MinValue
        if ([datetime]::TryParseExact([string]$status.取り込み日時, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ingested)) {
            $row.IngestedText = formatTime $ingested
        }
        $parts = splitIndexRelPath $status.相対パス
        if ($folderPaths.ContainsKey($parts.Name)) {
            $row.SourcePath = Join-Path $folderPaths[$parts.Name] $parts.Rest
        }
        [void]$rows.Add($row)
    }

    $ui.FailedGrid.ItemsSource = $rows
    $ui.FailedHeading.Text = "更新に失敗したファイル $($rows.Count) 件"
    $ui.FailedPanel.Visibility = if ($rows.Count -gt 0) { "Visible" } else { "Collapsed" }
}

$script:failedFileRequest = [ref]0  # 失敗したファイルを確かめる依頼の番号（find SourceFile の $script:openSourceRequest と同じ理由で [ref] のまま閉じ込める）

function openFailedFileFolder {
    # 失敗したファイルの場所をエクスプローラーで開く（ファイルを選択した状態）。
    # ネットワークにあれば裏のスレッドで確かめ、画面のスレッドは待たない
    $row = $ui.FailedGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    if (!$row.SourcePath) {
        setStatus "元のファイルの場所が分かりません（元のフォルダの記録がありません）：$($row.RelPath)"
        return
    }
    $path = $row.SourcePath
    $requestBox = $script:failedFileRequest
    $requestBox.Value++
    $requestId = $requestBox.Value
    $applyState = ${function:applyFailedFileState}
    $apply = {
        param ($state, $dirState)
        if ($requestId -ne $requestBox.Value) {
            # 待っている間に別の行を選んだ。前の依頼は捨てる
            return
        }
        & $applyState $state $path $dirState
    }.GetNewClosure()

    if (!(testNetworkPath $path)) {
        $state = getPathState $path
        $dirState = if ($state.State -eq ${pathStateMissing}) { getPathState (Split-Path $path -Parent) } else { $null }
        & $apply $state $dirState
        return
    }
    setStatus (getFailedFileCheckingStatus $path)
    $otherState = ${pathStateOther}
    startJob {
        param ($path)
        # フォルダの有無も、ファイルが無い（Missing）ときだけ、ここ（裏のスレッド）で調べて返す
        # （画面のスレッドで Test-Path すると、届かない共有・一覧に無いネットワークのエラーで止まるおそれがあるため）
        $state = getPathState $path
        $state
        if ($state.State -eq ${pathStateMissing}) {
            getPathState (Split-Path $path -Parent)
        }
    } @($path) {
        param ($output, $errorText)
        if ($errorText) {
            & $apply @{ State = $otherState; Message = $errorText } $null
        } else {
            & $apply $output[0] $output[1]
        }
    }.GetNewClosure() "network"
}

function applyFailedFileState {
    # openFailedFileFolder の続き（getPathState の結果を画面に反映する）。
    #   dirState: ファイルが無い（Missing）ときだけ渡す、フォルダの getPathState の結果。それ以外は $null
    param (
        $state,
        [string]$path,
        $dirState = $null
    )

    if ($state.State -eq ${pathStateFound} -and !$state.IsDirectory) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"${path}`""
        return
    }
    if ($state.State -eq ${pathStateUnreachable}) {
        setStatus (getFailedFileUnreachableStatus $path)
        return
    }
    if ($state.State -eq ${pathStateMissing}) {
        # フォルダが見つかったときだけ、フォルダを開く（有無は裏のスレッドで調べてある）
        if ($dirState -and $dirState.State -eq ${pathStateFound}) {
            Start-Process -FilePath "explorer.exe" -ArgumentList "`"$(Split-Path $path -Parent)`""
            setStatus "ファイルが見つからないため、フォルダを開きました（移動・削除された可能性があります）：${path}"
            return
        }
        setStatus "ファイルが見つかりません（移動・削除された可能性があります）：${path}"
        return
    }
    # その他（Other。アクセス拒否・一覧に無いネットワークのエラーなど）は、フォルダをたどらず文言だけ出す
    setStatus (getFailedFileOtherStatus $state.Message)
}

# 詳細の値の行を最後に置いたときの getIndexDetailView の RowsKey（同じ中身なら行を置き直さない）
$script:detailRowsKey = $null

function updateIndexDetailPanel {
    # 選んだインデックスの値の行と、高速検索の反映の進み具合を出す（文言の組み立ては getIndexDetailView）。
    # 選びなおしたとき・取り込みの集計や高速検索の確かめが届いたときに呼ぶ
    $selected = @()
    $entry = $null
    $item = $ui.IndexGrid.SelectedItem
    if ($null -ne $item) {
        $checkedText = if ($null -ne $script:fastSearchCheckedAt) { "最終確認 $($script:fastSearchCheckedAt.ToString('HH:mm'))" } else { "" }
        $selected = @(@{
            Name = $item.Name; Path = $item.Path; Enabled = $item.Enabled; FolderStatus = $item.StatusText
            IndexText = $item.IndexText; IndexLevel = $item.IndexLevel; IndexSub = $item.IndexSub; FileCountText = $item.FileCountText
            LastIngestedText = $item.LastIngestedText; FastText = $item.FastText; FastLevel = $item.FastLevel
            FastToolTip = $item.FastToolTip; FastCheckedText = $checkedText
        })
        $progress = $script:fastSearchProgress
        if ($progress -and $progress.ByIndex -and $progress.ByIndex.ContainsKey($item.Name)) {
            $entry = $progress.ByIndex[$item.Name]
        }
    }
    $view = getIndexDetailView $selected $entry

    $ui.IndexDetailTitle.Text = $view.Title
    $ui.IndexDetailBody.Visibility = if ($view.Selected) { "Visible" } else { "Collapsed" }
    $ui.IndexDetailName.Text = $view.Name
    $ui.IndexDetailPath.Text = $view.Path
    $ui.IndexDetailFolderStatus.Text = $view.FolderStatus
    setIndexBadge $ui.IndexDetailBadge $ui.IndexDetailBadgeText $view.Badge.Text $view.Badge.Level
    $ui.IndexDetailBadgeSub.Text = if ($view.Selected) { [string]$view.Badge.Sub } else { "" }
    $ui.IndexDetailUpdated.Text = $view.Updated
    $ui.IndexDetailCount.Text = $view.Count
    if ($script:detailRowsKey -ne $view.RowsKey -or $null -eq $ui.IndexDetailRows.ItemsSource) {
        $script:detailRowsKey = $view.RowsKey
        $rows = New-Object 'System.Collections.ObjectModel.ObservableCollection[DetailRow]'
        foreach ($row in $view.Rows) {
            $detailRow = New-Object DetailRow
            $detailRow.Label = $row.Label
            $detailRow.Value = $row.Value
            $rows.Add($detailRow)
        }
        $ui.IndexDetailRows.ItemsSource = $rows
    }
    setIndexBadge $ui.IndexDetailFastBadge $ui.IndexDetailFastBadgeText $view.Fast.State $view.Fast.Level
    $ui.IndexDetailFastReason.Text = $view.Fast.Reason
    $ui.IndexDetailFastReason.Visibility = if ($view.Fast.Reason) { "Visible" } else { "Collapsed" }
    $ui.IndexDetailFastPanel.Visibility = if ($view.Fast.Shown) { "Visible" } else { "Collapsed" }
    $ui.IndexDetailFastText.Text = $view.Fast.Text
    $ui.IndexDetailFastBar.Value = $view.Fast.Value
    $ui.IndexDetailFastChecked.Text = $view.Fast.Checked
    $ui.IndexDetailFastNote.Text = $view.Fast.Note
    $ui.IndexDetailFastNote.Visibility = if ($view.Fast.Note) { "Visible" } else { "Collapsed" }
}

function updateIndexSummaryText {
    $summary = $script:indexSummary
    if ($null -eq $summary) {
        $ui.IndexSummaryText.Text = "確認中…"
        return
    }
    if ($summary["Count"] -eq 0) {
        $ui.IndexSummaryText.Text = "まだインデックスがありません。"
        return
    }
    $text = "集約ファイル $($summary['Count'].ToString('N0')) 件 ・ 最終更新 $(formatTime $summary['LastWrite'])"
    $state = $script:indexingState
    if ($state -and $state.Done -gt 0) {
        $text = "更新済み $($state.Done.ToString('N0')) ファイル（$text）"
    }
    $ui.IndexSummaryText.Text = $text
}

function refreshIndexSummary {
    # 集約ファイルの件数は数えるのに時間がかかることがあるため、別スレッドで数える
    if ($script:summaryRunning) {
        $script:summaryAgain = $true
        return
    }
    $script:summaryRunning = $true
    $script:summaryAgain = $false
    $script:summaryJobDir = $workspace.IndexDir
    startJob {
        param ($folders)
        getIndexSummary $folders
    } @(, @($script:summaryJobDir)) {
        param ($output, $errorText)
        $script:summaryRunning = $false
        if ($script:summaryJobDir -ne $workspace.IndexDir) {
            # 数えている間にワークスペースを変えた。前のワークスペースの件数は出さず、数え直す
            $script:summaryAgain = $true
        } elseif ($output -and $output.Count -gt 0) {
            $script:indexSummary = $output[0]
        }
        updateIndexSummaryText
        updateSearchTarget
        if ($script:summaryAgain) {
            refreshIndexSummary
        }
    }
}
