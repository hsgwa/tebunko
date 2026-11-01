# インデックス管理の画面の詳細（インデックス作成の開始ボタン・状態・失敗の一覧・状態の合計）。

# 詳細の左右の部品。名前の一覧（shell\regions.ps1）に足さず、読み込んだ中身から取る
foreach ($detailName in @(
        "IndexDetailBody", "IndexDetailName", "IndexDetailPath", "IndexDetailPathButton", "IndexDetailSubfolders", "IndexDetailTextExtSummary", "ChangeIndexTextExtensionsButton", "IndexDetailTextExtStaleBadge", "IndexDetailTextExtStaleBadgeText", "IndexDetailTextExtCaption", "IndexDetailEditError", "IndexDetailFolderStatus", "IndexDetailFolderIcon",
        "IndexDetailBadge", "IndexDetailBadgeText", "IndexDetailBadgeSub", "IndexErrorText", "IndexErrorHintText", "IndexDetailUpdated", "IndexDetailCount",
        "IndexingDetailPanel", "IndexingProgressBar", "IndexingCountText", "IndexingFileText", "MultiSelectHintText",
        "IndexFilesButton")) {
    $ui[$detailName] = $ui.IndexDetailHost.Content.FindName($detailName)
}
# 詳細の「取り込む拡張子」の［変更…］（拡張子を選ぶダイアログを開き、［OK］で選びが変わっていればその場で設定に書く）
$ui.ChangeIndexTextExtensionsButton.Add_Click({ safe { changeIndexTextExtensions } })
# 読み取り中のファイルは、「…」で切れているときだけ全文をツールチップで見せる（名前とフォルダパスは入力欄で、欄の中で読める）
addTrimmedToolTip $ui.IndexingFileText

function setIndexBadge {
    # 詳細の状態のバッジの文言と色（一覧の LevelBadge と同じ組み合わせ。level は Ok / Wait / Warn / Ng / Run / None）
    param ($border, $textBlock, [string]$text, [string]$level)

    $kind = if ($level -in @("Ok", "Wait", "Warn", "Ng", "Run")) { $level } else { "None" }
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

    # 地・左端の線（アイコンと［検索する］［続きから再開］［中止］の枠と文字も同じ色）・文字・アイコンの形。見本の 3 状態に合わせ、ng は同じ規則で赤にする
    $looks = @{
        info = @("Badge.Run.Bg", "Accent", "Banner.Info.Text", "Icon.Info")
        warn = @("Warn.Soft", "Warn", "Banner.Warn.Text", "Icon.CircleAlert")
        ok = @("Ok.Soft", "Ok", "Banner.Ok.Text", "Icon.CircleCheck")
        ng = @("Danger.Soft", "Danger.Text", "Banner.Ng.Text", "Icon.CircleAlert")
    }
    $look = $looks[$level]
    if ($null -eq $look) { $look = $looks.info }
    $line = themeBrush $look[1]
    $text = themeBrush $look[2]
    $ui.IndexingProgressPanel.Background = themeBrush $look[0]
    $ui.IndexingProgressPanel.BorderBrush = $line
    $ui.IndexingProgressText.Foreground = $text
    $ui.IndexingProgressDetail.Foreground = $text
    $ui.IndexingBannerIcon.Data = themeBrush $look[3]
    $ui.IndexingBannerIcon.Stroke = $line
    foreach ($button in @($ui.IndexingSearchButton, $ui.IndexingResumeButton, $ui.IndexingStopButton)) {
        $button.BorderBrush = $line
        $button.Foreground = $line
    }
    $ui.IndexingBannerClose.Foreground = $text
    $script:indexingBannerKind = $kind
    $script:indexingBannerDismissed = $false
    setIndexingBannerBehavior $kind $level
}

function setIndexingBannerBehavior {
    # 帯の種類・閉じるボタン・自動で消える時機を、getIndexingBannerView の結果 1 回分で当てる。
    # 渡さなければ、出ている帯の種類と level（引き直し）。判断（更新中は残りを数えない など）は判断層
    param (
        [string]$kind = $script:indexingBannerKind,
        [string]$level = $script:indexingBannerLevel
    )

    $pending = if ($script:indexingState) { [int]$script:indexingState.Pending } else { 0 }
    $view = getIndexingBannerView $kind $level $pending (isIndexing)
    $script:indexingBannerLevel = $view.Level
    $ui.IndexingBannerClose.Visibility = if ($view.Closable) { "Visible" } else { "Collapsed" }
    $script:indexingBannerAutoClose = $view.AutoCloseSeconds
    $script:indexingBannerTimer.Stop()
    if ($view.AutoCloseSeconds -gt 0) {
        $script:indexingBannerTimer.Interval = [TimeSpan]::FromSeconds($view.AutoCloseSeconds)
        $script:indexingBannerTimer.Start()
    }
}

function closeIndexingBanner {
    # 帯を閉じる（×と、自動で消える時機）。残りがあるときは、閉じたあとも［続きから再開］の入口（［すべて更新］・行の［更新］・ナビの印）が残る。
    # 帯は、更新を始める・終わる・起動し直すまで出し直さない
    $script:indexingBannerTimer.Stop()
    $ui.IndexingProgressPanel.Visibility = "Collapsed"
    $script:indexingBannerKind = ""
    $script:indexingBannerDismissed = $true
}

function resetIndexingBanner {
    # 帯を初期に戻す（ワークスペースを切り替えたとき）。自動で消える時計も止め、前のワークスペースで閉じた印を持ち越さない
    $script:indexingBannerTimer.Stop()
    $ui.IndexingProgressPanel.Visibility = "Collapsed"
    $script:indexingBannerKind = ""
    $script:indexingBannerDismissed = $false
}

function updateIndexingResume {
    # 中断した更新（残りがあり、いま更新していない）の帯。［続きから再開］を出す。
    # 判断は getIndexingResumeBannerDecision（判断層）。ここは結果を画面に当てるだけ
    $pending = if ($script:indexingState) { [int]$script:indexingState.Pending } else { 0 }
    $decision = getIndexingResumeBannerDecision $pending (isIndexing) ($ui.IndexingProgressPanel.Visibility -eq "Visible") $script:indexingBannerKind $script:indexingBannerDismissed $script:indexingBannerLevel $script:indexingBannerAutoClose
    $script:indexingBannerDismissed = $decision.Dismissed
    $ui.IndexingResumeButton.Visibility = if ($decision.CanResume) { "Visible" } else { "Collapsed" }
    switch ($decision.Action) {
        "Show" {
            setIndexingBanner "warn" "interrupted"
            $ui.IndexingProgressText.Text = getIndexingStateText $pending $false
            $ui.IndexingProgressEta.Text = ""
            $ui.IndexingProgressDetail.Text = ""
            $ui.IndexingProgress.Visibility = "Collapsed"
            $ui.IndexingStopButton.Visibility = "Collapsed"
            $ui.IndexingSearchButton.Visibility = "Collapsed"
            $ui.IndexingProgressPanel.Visibility = "Visible"
        }
        "Text" { $ui.IndexingProgressText.Text = getIndexingStateText $pending $false }
        "Retime" { setIndexingBannerBehavior }
        "Hide" {
            $ui.IndexingProgressPanel.Visibility = "Collapsed"
            $script:indexingBannerKind = ""
        }
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
    if ($script:indexBusy -or $script:archiveBusy -or $script:indexingPreparing -or $script:targetsNaming -or $script:workspaceMoving) {
        # インデックスの削除・名前の変更中・エクスポート・インポート中（別スレッド）・ワークスペースを確かめている間は、インデックス作成もインデックスの操作も始めない
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
        $hint = if ($script:targetItems.Count -eq 0) { "フォルダを追加すると更新できます" } else { "更新できるインデックスがありません（行の［更新］で更新できます）。" }
    } elseif ($state -and $state.Failed -gt 0 -and !(isIndexing)) {
        $hint = "前回うまく更新できなかったファイルがあります（詳細の［ファイルの一覧］を「失敗のみ」で開くと確かめられます）。"
    }
    $ui.IndexingHint.Text = $hint
    $ui.IndexingHint.Visibility = if ($hint) { "Visible" } else { "Collapsed" }

    # インデックスの追加・編集・削除・エクスポート・インポートと、［設定］のワークスペースの［変更…］は互いに排他
    # （getIndexJobBlocker・getIndexTabButtonsEnabled。settings\settings.ps1 の testWorkspaceChangeable も同じ排他を見る）
    $selected = $null -ne (getIndexTargetItem)
    $blocker = getCurrentIndexJobBlocker
    $buttons = getIndexTabButtonsEnabled $blocker $selected
    $ui.NewIndexButton.IsEnabled = $buttons.New
    $ui.IndexDetailPathButton.IsEnabled = $buttons.ChangeFolder
    # 名前とフォルダパスの欄は、動いている処理が無く、1 件選んでいるときだけ直せる
    $ui.IndexDetailName.IsEnabled = $buttons.Edit
    $ui.IndexDetailPath.IsEnabled = $buttons.Edit
    $ui.IndexDetailSubfolders.IsEnabled = $buttons.Edit   # 更新中・ほかの処理の間は押せない
    $ui.ChangeIndexTextExtensionsButton.IsEnabled = $buttons.Edit   # 同じ（取り込む拡張子の［変更…］）
    # 行の右クリックのメニュー（対象は押した行。可否は getIndexRowMenuEnabled）
    $pressed = $ui.IndexGrid.SelectedItem
    $menu = getIndexRowMenuEnabled $blocker ($null -ne $pressed) $(if ($pressed) { $pressed.IndexLevel } else { "None" }) ($null -ne $pressed -and [bool]$pressed.Path)
    $ui.RowMenuUpdate.IsEnabled = $menu.Update
    $ui.RowMenuOpenFolder.IsEnabled = $menu.OpenFolder
    $ui.ExportIndexButton.IsEnabled = $menu.Export
    $ui.RemoveIndexButton.IsEnabled = $menu.Delete
    $actions = getIndexActionsEnabled $blocker @(getIndexCheckedItems @($script:targetItems)).Count
    $ui.ActionsButton.IsEnabled = $true
    $ui.ActionUpdate.IsEnabled = $actions.Update
    $ui.ActionExport.IsEnabled = $actions.Export
    $ui.ActionImport.IsEnabled = $actions.Import
    $ui.ActionDelete.IsEnabled = $actions.Delete
    # 行の［更新］も同じ排他（更新中はほかの行の［更新］も押せない）。行ごとのボタンを置き直す
    foreach ($item in $script:targetItems) {
        setIndexRowActions $item
    }
}

function refreshIndexingState {
    # 取り込み一覧の集計は、ファイルが数万行になると数秒～十数秒かかる。
    # 画面のスレッドで行うと、起動時・タブの切り替え時に画面が固まる（応答なしになる）ため別スレッドで数える
    if ($script:stateRunning) {
        $script:stateAgain = $true
        return
    }
    $script:stateRunning = $true
    $script:stateAgain = $false
    $script:stateJobPath = $workspace.StatusFile
    startJob {
        param ($path, $workspaceDir)
        $state = getIndexingState -path $path
        # 格納形式の状態（古い形式・作り直しの途中など）も同じ裏の仕事で調べる（置き場を数えるため）
        $state.Format = getWorkspaceIndexFormat ([Workspace]::new($workspaceDir))
        $state
    } @($script:stateJobPath, [string]$workspace.Dir) {
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
    } (getWorkspaceJobQueue $script:stateJobPath)
}

function getIndexingRatio {
    # 更新全体の進み（0～1）。更新中でない・割合がまだ分からないときは負の値
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

    # 失敗したファイルは、詳細の［ファイルの一覧］が「失敗のみ」で開いて見せる（applyIndexStats が詳細も更新する）
    applyIndexStats $state.IndexStats (isIndexing) $state.Format
    updateIndexingButton
    updateIndexingResume
    updateIndexTabBadge
}

function openIndexFilesDialog {
    # ［ファイルの一覧］: 詳細に出しているインデックスの、取り込んだファイルの一覧をダイアログで開く（index_files_dialog.ps1）。
    # 失敗が 1 件以上あれば「失敗のみ」、無ければ「すべて」で開く（getIndexFilesOpenFilter）
    $item = getIndexTargetItem
    if ($null -eq $item) {
        return
    }
    showIndexFilesDialog -Owner $window -IndexName ([string]$item.Name) -Filter (getIndexFilesOpenFilter ([int]$item.FailedCount))
}

$script:failedFileRequest = [ref]0  # 失敗したファイルを確かめる依頼の番号（findSourceFile の $script:openSourceRequest と同じ理由で [ref] のまま閉じ込める）

function openListedFileFolder {
    # 一覧の行（RelPath・SourcePath を持つ）のファイルの場所をエクスプローラーで開く（ファイルを選択した状態）。
    # ネットワークにあれば裏のスレッドで確かめ、画面のスレッドは待たない
    param (
        $row
    )

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
    }.GetNewClosure() (getWorkspaceJobQueue $path)
}

function applyFailedFileState {
    # openListedFileFolder の続き（getPathState の結果を画面に反映する）。
    #   dirState: ファイルが無い（Missing）ときだけ渡す、フォルダの getPathState の結果。それ以外は $null
    param (
        $state,
        [string]$path,
        $dirState = $null
    )

    if ($state.State -eq ${pathStateFound} -and !$state.IsDirectory) {
        [void](showInExplorer $path)
        return
    }
    if ($state.State -eq ${pathStateUnreachable}) {
        setStatus (getFailedFileUnreachableStatus $path)
        return
    }
    if ($state.State -eq ${pathStateMissing}) {
        # フォルダが見つかったときだけ、フォルダを開く（有無は裏のスレッドで調べてある）
        if ($dirState -and $dirState.State -eq ${pathStateFound} -and $dirState.IsDirectory) {
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
$script:indexingView = $null  # 詳細の「インデックス」の箱に出す、全体の進み具合（updateIndexingProgress が置く。更新中でなければ $null）

function getIndexTargetItem {
    # 詳細に出している行（名前・フォルダパスの欄と［変更…］、右クリックの［削除］［エクスポート…］の対象・可否もこの行。判定は getIndexDetailItem。［アクション ▼］は別で、可否・対象は getIndexActionsEnabled と、チェックの行が決める）
    return getIndexDetailItem $ui.IndexGrid.SelectedItem @(getIndexCheckedItems @($script:targetItems))
}

function setIndexDetailSubfolders {
    # 詳細の『サブフォルダを含める』に行の値を入れる。入れ直している間は、チェックの出来事で設定を書かない（changeIndexSubfolders が見る）
    param (
        [bool]$value
    )

    $script:indexDetailApplying = $true
    try { $ui.IndexDetailSubfolders.IsChecked = $value } finally { $script:indexDetailApplying = $false }
}

function getIndexTextExcludesOf {
    # 控えた設定（$script:textExtState）の、そのインデックスの値。無ければ空の配列
    param (
        $table,
        [string]$name
    )

    if ($null -eq $table -or [string]::IsNullOrEmpty($name) -or -not $table.ContainsKey($name)) {
        return @()
    }
    return , @($table[$name])
}

function getIndexTextExtItemValues {
    # getIndexDetailView に渡す、そのインデックスの取り込む拡張子の値（今の設定・前回取り込んだときの設定・件数の控えの値）
    param (
        [string]$name
    )

    $entry = $(if ($name -and $script:textExtState.Counts.ContainsKey($name)) { $script:textExtState.Counts[$name] } else { $null })
    return @{
        TextExcludes = (getIndexTextExcludesOf $script:textExtState.Excludes $name)
        TextExcludesIngested = (getIndexTextExcludesOf $script:textExtState.Ingested $name)
        TextCountsEntry = $entry
        TextExtEditable = (getIndexTabButtonsEnabled (getCurrentIndexJobBlocker) $true).Edit   # 更新中・ほかの処理の間は押せない
    }
}

function changeIndexTextExtensions {
    # 詳細の「取り込む拡張子」の［変更…］。ダイアログを開き、［OK］で選びが変わっていれば、設定（indexExtensionExcludes）にすぐ書いて一覧・詳細を直す。
    # ［要更新］は印を持たず、前回取り込んだときの設定との差から求めるので、描き直すだけでよい（testIndexTextExcludesChanged）。更新中・ほかの処理の間は何もしない
    if (!$ui.ChangeIndexTextExtensionsButton.IsEnabled -or (isIndexingOrPreparing)) {
        return
    }
    $item = getIndexTargetItem
    if ($null -eq $item -or [string]::IsNullOrEmpty($item.Name)) {
        return
    }
    syncIndexTextExtensions
    $current = getIndexTextExcludesOf $script:textExtState.Excludes $item.Name
    $entry = $(if ($script:textExtState.Counts.ContainsKey($item.Name)) { $script:textExtState.Counts[$item.Name] } else { $null })
    $row = getIndexTextExtensionDetailRow $current $entry
    $result = showIndexTextExtensionsDialog -Owner $window -IndexName $item.Name -Excluded $current -Counts $row.Counts -CountsSource $row.Source -CountsTime $row.Time
    if ($null -ne $result) {
        $result = @(addOutOfListExtensions $result $current @(getIndexSelectableExtensions))
    }
    if ($null -eq $result -or -not (testIndexTextExcludesChanged $result $current)) {
        return
    }
    try {
        writeIndexTextExcludes $item.Name $result
    } catch {
        showMessage "設定を保存できませんでした。もう一度お試しください。" "OK" "Warning" | Out-Null
        return
    }
    if ($null -ne $script:indexingState) {
        applyIndexStats $script:indexingState.IndexStats (isIndexing) $script:indexingState.Format
    } else {
        syncIndexTextExtensions
        updateIndexDetailPanel
    }
}

function setIndexDetailTextExt {
    # 詳細の「取り込む拡張子」の行に、getIndexDetailView の TextExt（要約・件数の見出し・［要更新］の札・［変更…］の可否）を入れる
    param (
        $textExt
    )

    $ui.IndexDetailTextExtSummary.Text = $textExt.Summary
    $ui.IndexDetailTextExtCaption.Text = $textExt.Caption
    setIndexBadge $ui.IndexDetailTextExtStaleBadge $ui.IndexDetailTextExtStaleBadgeText $(if ($textExt.Stale) { "要更新" } else { "" }) "Wait"
    $ui.ChangeIndexTextExtensionsButton.IsEnabled = $textExt.ChangeEnabled
}

function updateIndexDetailPanel {
    # 選んだインデックスの値の行を出す（文言の組み立ては getIndexDetailView）。
    # 選びなおしたとき・取り込みの集計が届いたときに呼ぶ
    $selected = @()
    $checkedItems = @(getIndexCheckedItems @($script:targetItems))
    $item = getIndexDetailItem $ui.IndexGrid.SelectedItem $checkedItems
    if ($null -ne $item) {
        $selected = @(@{
            Name = $item.Name; Path = $item.Path; Enabled = $item.Enabled; Subfolders = $item.Subfolders; FolderStatus = $item.StatusText
            FolderMissing = ($item.StatusChecked -and !$item.FolderExists); FolderChecked = [bool]$item.StatusChecked
            IndexText = $item.IndexText; IndexLevel = $item.IndexLevel; IndexSub = $item.IndexSub; FileCountText = $item.FileCountText
            LastIngestedText = $item.LastIngestedText; FailedCount = $item.FailedCount; PendingCount = $item.PendingCount
        } + (getIndexTextExtItemValues $item.Name))
    }
    $multiCount = getIndexDetailMultiCount $checkedItems.Count
    $view = getIndexDetailView $selected $multiCount $script:indexingView

    $ui.IndexDetailTitle.Text = $view.Title
    $ui.IndexDetailBody.Visibility = if ($view.Selected) { "Visible" } else { "Collapsed" }
    $ui.MultiSelectHintText.Text = $view.Hint
    $ui.MultiSelectHintText.Visibility = if ($view.Hint) { "Visible" } else { "Collapsed" }
    # 更新中の行では、最終更新・件数の代わりに進み具合（棒・件数と残り時間・ファイル名）を出す
    $ui.IndexingDetailPanel.Visibility = if ($view.Run.Shown) { "Visible" } else { "Collapsed" }
    $ui.IndexingProgressBar.Value = $view.Run.Value
    $ui.IndexingCountText.Text = $view.Run.CountText
    $ui.IndexingFileText.Text = $view.Run.FileText
    # 名前・フォルダパスの欄は、直している最中（同じ行で、欄に入力の途中）なら、定期の更新で入力を消さない
    $editing = ($null -ne $item -and [object]::ReferenceEquals($item, $script:detailEditItem) -and
        ($ui.IndexDetailName.IsFocused -or $ui.IndexDetailPath.IsFocused))
    if (!$editing) {
        if (![object]::ReferenceEquals($item, $script:detailEditItem)) {
            setIndexDetailEditError ""
        }
        $script:detailEditItem = $item
        $ui.IndexDetailName.Text = $view.Name
        $ui.IndexDetailPath.Text = $view.Path
    }
    setIndexDetailSubfolders $view.Subfolders
    setIndexDetailTextExt $view.TextExt
    $ui.IndexDetailFolderStatus.Text = $view.FolderStatus
    $ui.IndexDetailFolderStatus.Foreground = themeBrush $view.FolderBrush
    # 文の左のアイコン（Lucide の check・x）。空なら隠す
    if ($view.FolderIcon) {
        $ui.IndexDetailFolderIcon.Data = (loadTheme)["Icon.$($view.FolderIcon)"]
        $ui.IndexDetailFolderIcon.Stroke = themeBrush $view.FolderBrush
        $ui.IndexDetailFolderIcon.Visibility = "Visible"
    } else {
        $ui.IndexDetailFolderIcon.Visibility = "Collapsed"
    }
    setIndexBadge $ui.IndexDetailBadge $ui.IndexDetailBadgeText $view.Badge.Text $view.Badge.Level
    $ui.IndexDetailBadgeSub.Text = if ($view.Selected) { [string]$view.Badge.Sub } else { "" }
    $ui.IndexErrorText.Text = $view.Error
    $ui.IndexErrorText.Visibility = if ($view.Error) { "Visible" } else { "Collapsed" }
    $ui.IndexErrorHintText.Text = $view.ErrorHint
    $ui.IndexErrorHintText.Visibility = if ($view.ErrorHint) { "Visible" } else { "Collapsed" }
    $ui.IndexDetailUpdated.Text = if ($view.Run.Shown) { "" } else { $view.Updated }
    $ui.IndexDetailCount.Text = if ($view.Run.Shown) { "" } else { $view.Count }
    $ui.IndexDetailUpdated.Visibility = if ($ui.IndexDetailUpdated.Text) { "Visible" } else { "Collapsed" }
    $ui.IndexDetailCount.Visibility = if ($ui.IndexDetailCount.Text) { "Visible" } else { "Collapsed" }
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
}

function refreshIndexSummary {
    # 本文インデックスのファイルの件数は数えるのに時間がかかることがあるため、別スレッドで数える
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
        updateSearchTarget
        if ($script:summaryAgain) {
            refreshIndexSummary
        }
    } (getWorkspaceJobQueue $script:summaryJobDir)
}
