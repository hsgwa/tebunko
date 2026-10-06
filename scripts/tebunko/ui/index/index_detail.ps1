# インデックス管理の画面の詳細（インデックス作成の開始ボタン・状態・失敗の一覧・状態の合計）。

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
        $ui.IndexingButton.Content = "インデックス作成中…"
        $ui.IndexingButton.IsEnabled = $false
    } else {
        $ui.IndexingButton.Content = if ($state -and $state.Pending -gt 0) { "続きから再開（残り $($state.Pending) 件）" } else { "インデックス作成を開始" }
        $ui.IndexingButton.IsEnabled = $ready
    }

    # ボタンの下の一言。押す前は「押すと何が起きるか」、インデックス作成中は「やめるとどうなるか」を書く
    $hint = if (isIndexing) {
        "インデックス作成中も検索できます。やめるときは［中止］を押してください（次に［インデックス作成を開始］を押すと続きから再開します）。"
    } else {
        "押すと、何件取り込むかを確認してから、新しいファイル・変わったファイルだけを取り込みます。"
    }
    if (!$ready -and !(isIndexing)) {
        $hint = "まずインデックスを追加して、［作成］にチェックを付けてください。"
    } elseif ($state -and $state.Failed -gt 0 -and !(isIndexing)) {
        $hint = "前回うまく取り込めなかったファイルがあります（押したあとで、もう一度ためすか選べます）。" + $hint
    }
    $ui.IndexingHint.Text = $hint

    # インデックスの追加・編集・削除・エクスポート・インポートと、［8 設定］のワークスペースの［変更…］は互いに排他
    # （getIndexJobBlocker・getIndexTabButtonsEnabled。settings_tab.ps1 の testWorkspaceChangeable も同じ排他を見る）
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

function applyIndexingState {
    # 集計（別スレッド）の結果を画面に反映する
    param (
        $state  # getIndexingState の結果
    )

    $script:indexingState = $state

    # 失敗したファイルは下の一覧に原因とともに表示する
    $ui.IndexingStateText.Text = getIndexingStateText $state.Pending (isIndexing)
    $ui.IndexTabBadge.Visibility = if ($state.Failed -gt 0) { "Visible" } else { "Collapsed" }
    applyIndexStats $state.IndexStats (isIndexing)
    updateFailedList $state
    updateIndexSummaryText
    updateIndexingButton
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
    $ui.FailedHeading.Text = "⚠ 取り込みに失敗したファイル $($rows.Count) 件"
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
        setStatus "元のファイルの場所が分かりません（取り込み一覧にクロール対象フォルダの記録がありません）：$($row.RelPath)"
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
        $selected = @(@{
            Name = $item.Name; Path = $item.Path; Enabled = $item.Enabled; FolderStatus = $item.StatusText
            IndexText = $item.IndexText; FileCountText = $item.FileCountText
            LastIngestedText = $item.LastIngestedText; FastText = $item.FastText
        })
        $progress = $script:fastSearchProgress
        if ($progress -and $progress.ByIndex -and $progress.ByIndex.ContainsKey($item.Name)) {
            $entry = $progress.ByIndex[$item.Name]
        }
    }
    $view = getIndexDetailView $selected $entry

    $ui.IndexDetailTitle.Text = $view.Title
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
        $ui.IndexDetailRows.Visibility = if ($rows.Count -gt 0) { "Visible" } else { "Collapsed" }
    }
    $ui.IndexDetailFastPanel.Visibility = if ($view.Fast.Shown) { "Visible" } else { "Collapsed" }
    $ui.IndexDetailFastText.Text = $view.Fast.Text
    $ui.IndexDetailFastBar.Value = $view.Fast.Value
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
    $text = "集約ファイル $($summary['Count'].ToString('N0')) 件 ・ 最終取り込み $(formatTime $summary['LastWrite'])"
    $state = $script:indexingState
    if ($state -and $state.Done -gt 0) {
        $text = "取り込み済み $($state.Done.ToString('N0')) ファイル（$text）"
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
