# インデックス管理の画面の一覧（状態の保持・フォルダと高速検索の状態の調べ・一覧の読み込みと更新）。

function getTargetsKey {
    # クロール対象フォルダの一覧（@{ Path; Enabled } の配列）を比べるための文字列
    param (
        [object[]]$folders
    )

    return (@($folders | Where-Object { $_ } | ForEach-Object { "$($_.Enabled)`t$($_.Path)" }) -join "`n")
}

$script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexGrid.ItemsSource = $script:targetItems
# 1 件も無いときの［＋ フォルダを追加］。名前の一覧（gui_main.ps1）に足さず、読み込んだ中身から取る（件数の帯は窓の下のステータスバーにある）
foreach ($footerName in @("IndexEmptyAddButton")) {
    $ui[$footerName] = $ui.IndexListHost.Content.FindName($footerName)
}
$script:loadingTargets = $false
$script:indexBusy = $false  # 前のインデックスの削除中（別スレッド）。getIndexJobBlocker に [bool] で渡すため、$null のままにしない
# 行のチェック（一時の選択。保存しない）の状態が変わったら、行の値と画面の表示を合わせる（Checked・Unchecked。
# ToggleButton の状態が変わったときに出る、バブルするイベント）。マウスの Click だけでなく、UI オートメーションの
# TogglePattern（キーボード操作も同様）でも状態が変わったときに出るため、どの操作でも拾える。
# UI オートメーションの Toggle は IsChecked（表示）を変えるが、TwoWay バインドの先（Checked。PS class の
# プレーンなプロパティで PropertyChanged を出さない）へは反映されないことがあるため、ここで明示的に合わせる。
# 行の初期化としての Checked・Unchecked（DataGrid が行の見た目を作る・仮想化）でも出るが、値を合わせるだけで害は無い
$onIndexGridToggled = {
    param ($s, $e)
    safe {
        $cb = $e.OriginalSource
        if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.DataContext -is [FolderItem] -and !$script:loadingTargets) {
            $cb.DataContext.Checked = [bool]$cb.IsChecked
            updateIndexSelectionView
            updateIndexingButton
            updateIndexDetailPanel
        }
    }
}
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::CheckedEvent, [System.Windows.RoutedEventHandler]$onIndexGridToggled)
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::UncheckedEvent, [System.Windows.RoutedEventHandler]$onIndexGridToggled)
$script:indexListEmpty = $null  # 前回の updateIndexListView で一覧が空だったか（詳細の行の高さを動かすかの判断に使う）
$script:indexDetailRowRestore = $null  # 空で畳む前の詳細の行の高さ（境目で変えた高さも含む）
$script:savedTargets = $null  # 最後に読み込み・保存したインデックス一覧（getTargetsKey）。ほかでの変更の検出に使う
$script:editDialog = $null    # 追加・編集のダイアログ（開いている間だけ）
$script:indexingSession = $null  # 実行中のインデックス作成（IndexingSession。終わって片づけたら $null）
$script:indexingOnlyNames = @()  # 今の回が選んだものだけのときの、インデックス名（空なら全部）
$script:indexingStart = $null
$script:ingestFailed = 0  # インデックス作成中に一覧へ反映済みの失敗件数
$script:indexingState = $null
$script:indexSummary = $null
$script:archiveBusy = $false  # エクスポート・インポート中（別スレッド）。settings\settings.ps1 の testWorkspaceChangeable も見る
$script:archiveJobOperation = ""
$script:archiveJobOnSuccess = $null

# 一覧の「高速検索」列（別スレッドで確かめ、届いた結果を全行に置き直す）
$script:fastSearchChecking = $false
$script:fastSearchAgain = $false
$script:fastSearchJobDir = $null
$script:fastSearchReason = $null      # 最後に届いた getWindowsSearchState の値（まだなら $null）
$script:fastSearchProgress = $null    # 最後に届いた getSystemIndexProgress の値
$script:fastSearchCheckedAt = $null   # 最後に確かめ終えた時刻（DateTime）
$script:fastSearchResultDir = $null   # 上の 3 つを確かめたときのワークスペース（$workspace.Dir）。ワークスペースを切り替えた後、
                                       # 新しい確かめが終わるまでは古いワークスペースの結果なので、更新のたびに確かめて捨てる
$script:fastSearchGeneration = 0      # インデックス作成が終わるたびに finishIndexing が 1 増やす世代番号。
                                       # 作成前から走っていた確かめジョブ（fastSearchJobGeneration が古いまま）が
                                       # 作成後に結果を届けても、世代が合わなければ古い結果として捨て、確かめ直す
$script:fastSearchJobGeneration = $null  # 今走っている確かめジョブを始めたときの世代番号（fastSearchJobDir と同じ使い方）
# 反映待ち・反映中の行がある間だけ動かす（無くなったら止める。画面を閉じれば、ほかのタイマーと同じく動かなくなる）
$script:fastSearchTimer = newTimer (5 * 60 * 1000) { safe { refreshFastSearchStatus } }

function isIndexing {
    return ($null -ne $script:indexingSession) -and $script:indexingSession.IsRunning()
}

function shouldRefreshFastSearchStatus {
    # 高速検索の列は、前の確かめから 60 秒たっていなければ飛ばす（ウィンドウを前に出すたびに問い合わせない）
    return ($null -eq $script:fastSearchCheckedAt) -or ((Get-Date) - $script:fastSearchCheckedAt).TotalSeconds -ge 60
}

$script:folderCheckRunning = $false
$script:folderCheckAgain = $false

function updateFolderItemStatus {
    # フォルダの有無を調べ直す（別スレッド。refreshFolderStatus）。調べ終えるまでは前の表示のまま（初回は「確認中」）
    param (
        $item
    )

    if (!$item.StatusChecked) {
        $item.SetStatus("… フォルダを確認しています", ${grayBrush})
    }
    updateIndexDetailPanel
    refreshFolderStatus
}

function refreshFolderStatus {
    # 一覧のすべてのフォルダの有無を別スレッドで調べ、表示と［すべて更新］の可否に反映する。
    # 届かないネットワークのフォルダ（VPN の切断・サーバーの停止）では Test-Path が十数秒戻らないため、
    # 画面のスレッドで調べると起動時・画面を前に出すたびに「応答なし」になる（実測 約 17 秒）。
    # 調べている間に呼ばれたら、終わってからもう一度だけ調べる
    if ($script:folderCheckRunning) {
        $script:folderCheckAgain = $true
        return
    }
    $paths = @($script:targetItems | ForEach-Object { [string]$_.Path } | Where-Object { $_ } | Select-Object -Unique)
    if ($paths.Count -eq 0) {
        return
    }
    $script:folderCheckRunning = $true
    $script:folderCheckAgain = $false
    # 届かないネットワークのフォルダが 1 つでもあれば、専用の列（network）を使う。
    # プレビュー等の列（既定。2 スレッド）は、届かない共有の Test-Path で塞がれても待たされないようにする
    $queue = if (testAnyNetworkPath $paths) { "network" } else { "default" }
    startJob {
        param ($paths)
        $result = @{}
        foreach ($path in $paths) {
            # 届かないネットワークのフォルダでは Test-Path が例外（ネットワーク パスが見つかりません）になるため、無いものとする
            $found = $false
            try {
                $found = [bool](Test-Path -LiteralPath $path -PathType Container -ErrorAction Stop)
            } catch {
            }
            $result[$path] = $found
        }
        $result
    } @(, [string[]]$paths) {
        param ($output, $errorText)
        $script:folderCheckRunning = $false
        if ($output -and $output.Count -gt 0) {
            applyFolderStatus $output[0]
        }
        if ($script:folderCheckAgain) {
            refreshFolderStatus
        }
    } $queue
}

function applyFolderStatus {
    # refreshFolderStatus の結果（パス → 有無）を一覧に反映する（調べている間にパスが変わった行はそのまま）
    param (
        $exists
    )

    foreach ($item in $script:targetItems) {
        if (!$exists.ContainsKey([string]$item.Path)) {
            continue
        }
        $item.StatusChecked = $true
        $item.FolderExists = [bool]$exists[[string]$item.Path]
        if ($item.FolderExists) {
            $item.SetStatus("✓ フォルダがあります", ${okBrush})
        } else {
            $item.SetStatus("✗ フォルダが見つかりません", ${ngBrush})
        }
    }
    updateIndexingButton
    updateIndexDetailPanel
}

function refreshFastSearchStatus {
    # 一覧の「高速検索」列を、別スレッドで Windows Search に 1 回問い合わせて確かめ直す（getWindowsSearchState・
    # Reason が Ok・NotYet のときだけ getSystemIndexProgress を同じ接続で）。確かめている間に呼ばれたら、
    # 終わってからもう一度だけ確かめる。確かめている間にワークスペースが変わったら結果を捨てる。
    # インデックス作成中（isIndexing）は、状態ファイルと txt が書き換わっている途中のため確かめない
    # （インデックス作成が終わったとき（finishIndexing）に、別途呼ぶ）
    if (isIndexing) {
        return
    }
    if ($script:fastSearchChecking) {
        $script:fastSearchAgain = $true
        return
    }
    $script:fastSearchChecking = $true
    $script:fastSearchAgain = $false
    $script:fastSearchJobDir = $workspace.Dir
    $script:fastSearchJobGeneration = $script:fastSearchGeneration
    # 届かないネットワークのワークスペースでは、フォルダの有無と同じ専用の列（network）を使う
    $queue = if (testAnyNetworkPath @($workspace.Dir)) { "network" } else { "default" }
    startJob {
        param ($systemRoot, $indexRoot, $statePath)
        $connection = openWindowsSearch
        try {
            if ($null -eq $connection) {
                return @{ Reason = "NoConnection"; Progress = $null }
            }
            $reason = getWindowsSearchState $systemRoot "" $connection
            $progress = $null
            if ($reason -eq "Ok" -or $reason -eq "NotYet") {
                $progress = getSystemIndexProgress $null $indexRoot $systemRoot $statePath $connection
            }
            @{ Reason = $reason; Progress = $progress }
        } finally {
            if ($connection) {
                $connection.Dispose()
            }
        }
    } @($workspace.SystemIndexDir, $workspace.IndexDir, $workspace.SystemIndexStateFile) {
        param ($output, $errorText)
        $script:fastSearchChecking = $false
        if ($script:fastSearchJobDir -ne $workspace.Dir -or $script:fastSearchJobGeneration -ne $script:fastSearchGeneration) {
            # 確かめている間にワークスペースを変えた、またはインデックス作成が終わって世代が変わった。
            # 古い結果は出さず、確かめ直す
            $script:fastSearchAgain = $true
        } elseif ($output -and $output.Count -gt 0) {
            applyFastSearchStatus $output[0]
        } elseif ($errorText) {
            # 確かめの途中で例外が起きたときも、接続できなかったときと同じ扱いにする（バッジが固まったままにしない）
            applyFastSearchStatus @{ Reason = "NoConnection"; Progress = $null }
        }
        if ($script:fastSearchAgain) {
            refreshFastSearchStatus
        }
    } $queue
}

function applyFastSearchStatus {
    # refreshFastSearchStatus の結果（@{ Reason; Progress }）を覚えて、一覧の「高速検索」列に置き直す
    param (
        $result
    )

    $script:fastSearchReason = $result.Reason
    $script:fastSearchProgress = $result.Progress
    $script:fastSearchCheckedAt = Get-Date
    $script:fastSearchResultDir = $workspace.Dir
    updateFastSearchRows
}

function updateFastSearchRows {
    # 一覧の各行の「高速検索」列を getFastSearchRowView で置き直す（確かめの結果が届いたときと、
    # applyIndexStats で本文の集計が届いたときの両方から呼ぶ。ここでは I/O をせず判断層を呼ぶだけ）。
    # $hasContent は、そのときの getIndexStats の値（Done が 1 以上か）から渡す。
    # isIndexing も渡し、作成中に hasContent が真へ変わっても（高速検索の確かめは作成中は走らない
    # ため）前の確かめの結果のまま「不可」にしない（getFastSearchRowView の indexing 引数）。
    # 覚えている結果が今のワークスペースのものでなければ（切り替えた直後）、古い可否を出さないよう捨てて「確認中…」に戻す
    if ($script:fastSearchResultDir -ne $workspace.Dir) {
        $script:fastSearchReason = $null
        $script:fastSearchProgress = $null
        $script:fastSearchCheckedAt = $null
    }
    $stats = if ($script:indexingState) { $script:indexingState.IndexStats } else { $null }
    $waiting = $false
    foreach ($item in $script:targetItems) {
        $hasContent = $false
        if ($item.Name -and $null -ne $stats -and $stats.ContainsKey($item.Name)) {
            $hasContent = ($stats[$item.Name].Done -ge 1)
        }
        $row = getFastSearchRowView $script:fastSearchReason $script:fastSearchProgress $item.Name $hasContent $script:fastSearchCheckedAt (isIndexing)
        $item.SetFast($row.Text, $row.ToolTip, $row.Level)
        if ($row.Level -eq "Wait") {
            $waiting = $true
        }
    }
    updateIndexDetailPanel
    if ($waiting) {
        if (!$script:fastSearchTimer.IsEnabled) {
            $script:fastSearchTimer.Start()
        }
    } elseif ($script:fastSearchTimer.IsEnabled) {
        $script:fastSearchTimer.Stop()
    }
}

function newFolderItem {
    param (
        [string]$path,
        [bool]$enabled,
        [string]$name = ""
    )

    $item = New-Object FolderItem
    $item.Name = $name
    $item.Path = $path
    $item.Enabled = $enabled
    $item.FileCountText = "－"
    $item.LastIngestedText = ""
    # フォルダの有無は一覧に加えた後にまとめて調べる（refreshFolderStatus）
    $item.SetStatus("… フォルダを確認しています", ${grayBrush})
    # 行のチェックの保存は、一覧のチェックボックスの Checked・Unchecked（IndexGrid.AddHandler）で行う。
    # PS class のプレーンなプロパティは TwoWay セットで PropertyChanged を出さないため、購読では拾えない。
    return $item
}

function loadTargets {
    $script:loadingTargets = $true
    try {
        # 読み直しても、チェック（一時の選択）は名前で引き継ぐ
        $checkedNames = @(getIndexCheckedItems @($script:targetItems) | ForEach-Object { $_.Name })
        $script:targetItems.Clear()
        $folders = @(getTargetFolders)
        # 名前の決まっていないインデックス（設定ファイルを直接書き換えた場合など）には、ここで名前を割り当てて確定する。
        # 一覧・編集・削除はインデックス名で扱うため、画面に出す時点で名前があるようにする（インデクサと同じ assignIndexNames を使う）。
        # 取り込み一覧の読み込み（ネットワーク上のこともある）は排他の外で済ませ、設定の読み直しと書き込みだけを saveAssignedIndexNames が排他の中で行う
        if (@($folders | Where-Object { $_ -and !$_.Name }).Count -gt 0) {
            saveAssignedIndexNames @(assignIndexNames $folders (readStatusFile).Folders)
            $folders = @(getTargetFolders)
        }
        foreach ($folder in $folders) {
            $newItem = newFolderItem $folder.Path $folder.Enabled $folder.Name
            $newItem.Checked = ($folder.Name -and $checkedNames -contains $folder.Name)
            $script:targetItems.Add($newItem)
        }
        $script:savedTargets = getTargetsKey @(getTargetFolders)
    } finally {
        $script:loadingTargets = $false
    }
    updateIndexListView
    refreshFolderStatus
    refreshFastSearchStatus
}

function saveTargets {
    writeTargetFolders @($script:targetItems | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Enabled = $_.Enabled } })
    $script:savedTargets = getTargetsKey @(getTargetFolders)
    setStatus "インデックス一覧を保存しました（$(Get-Date -Format 'H:mm')）"
}

function updateIndexSourceFile {
    # インデックスのフォルダの source_folder.txt を今の一覧に合わせて書き直す。
    # 次のインデックス作成を待たずに、検索結果から元のファイルを開けるようにする（インデックスが無ければ何もしない）
    if (!(Test-Path -LiteralPath $workspace.IndexDir -PathType Container)) {
        return
    }
    writeSourceFolderFile @($script:targetItems | Where-Object { $_.Name } | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path } })
}

function refreshIndexViews {
    # インデックスを追加・編集・削除した後、検索タブ（検索対象のツリー・件数）も読み直す。
    # 一覧の「高速検索」列も確かめ直す（編集・削除で中身が変わっているため）
    $script:sourceFolderMaps = @{}
    $script:indexSummary = $null
    loadIndexTree
    refreshIndexSummary
    refreshIndexingState
    refreshFastSearchStatus
    updateIndexDetailPanel
}

function applyIndexStats {
    # 取り込み一覧の集計（getIndexStats）を一覧の各行のファイル数・最終更新・「ステータス」列に反映し、
    # 「高速検索」列も（getIndexStats の Done を使って）置き直す。
    #   indexing: インデックス作成中か（getIndexRowView にそのまま渡す）
    param (
        $stats,
        [bool]$indexing = $false
    )

    foreach ($item in $script:targetItems) {
        $stat = $null
        if ($item.Name -and $null -ne $stats -and $stats.ContainsKey($item.Name)) {
            $stat = $stats[$item.Name]
        }
        if ($null -eq $stat) {
            $item.SetStats("－", "まだ更新していません", "")
        } else {
            $ingested = [datetime]::MinValue
            $lastText = if ($stat.LastIngested -and [datetime]::TryParseExact($stat.LastIngested, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ingested)) {
                formatTime $ingested
            } else {
                ""
            }
            $item.SetStats(("{0:#,0}" -f $stat.Total), ("更新済み {0:#,0} 件 ・ 未更新 {1:#,0} 件 ・ 失敗 {2:#,0} 件" -f $stat.Done, $stat.Pending, $stat.Failed), $lastText)
        }
        $row = getIndexRowView $stat $indexing $item.Enabled (getIndexingRatio) (testIndexInRound $item)
        $item.SetIndexState($row.Text, $row.ToolTip, $row.Level, $row.Sub, [double]$row.Percent)
        setIndexRowActions $item
    }
    $script:indexFileTotal = 0
    if ($null -ne $stats) {
        foreach ($item in $script:targetItems) {
            if ($item.Name -and $stats.ContainsKey($item.Name)) {
                $script:indexFileTotal += [int]$stats[$item.Name].Total
            }
        }
    }
    updateIndexFooter
    updateFastSearchRows
}

$script:indexFileTotal = 0

function testIndexInRound {
    # 今の回で更新する行か（選んだものだけの回では、選んだ名前の行だけ）
    param ($item)

    return (@($script:indexingOnlyNames).Count -eq 0 -or @($script:indexingOnlyNames) -contains $item.Name)
}

function setIndexRowActions {
    # 行の右端のボタン（［更新］［中止］・何も出さない）を、行の状態（IndexLevel）と動いている処理に合わせる。判断は getIndexRowActions
    param ($item)

    $blocker = getIndexJobBlocker (isIndexing) $script:indexBusy $script:archiveBusy
    $actions = getIndexRowActions $item.IndexLevel $blocker
    $item.SetRowActions($actions.Action, $actions.UpdateEnabled)
}

function updateSelectedIndexes {
    # 選んだインデックス（名前の配列）だけを更新する。行の［更新］と［アクション ▾］の［更新］の共通の入口。
    # 選んだ名前の行を、設定の enabled を true にして取り込み、その回だけの更新を始める（startIndexing）。
    # 動いている処理があるときは始めない（ボタンも使えなくしてある）
    param (
        [string[]]$names
    )

    $names = @($names | Where-Object { $_ })
    if ($names.Count -eq 0 -or (getIndexJobBlocker (isIndexing) $script:indexBusy $script:archiveBusy) -ne "") {
        return
    }
    startIndexing $names
}

function updateIndexSelectionView {
    # 見出しの全選択（一部だけなら横棒）と、見出しの横の「N / M 件を選択中」。判断は getIndexSelectionView
    $view = getIndexSelectionView $script:targetItems.Count @(getIndexCheckedItems @($script:targetItems)).Count
    $ui.SelectionCountText.Text = $view.CountText
    $ui.SelectionCountText.Visibility = if ($view.CountText) { "Visible" } else { "Collapsed" }
    $ui.SelectAllCheckBox.IsChecked = $view.AllChecked
    $ui.SelectAllCheckBox.IsEnabled = ($script:targetItems.Count -gt 0)
}

function updateIndexRowsProgress {
    # 更新中の行のバッジ（「更新中 45%」）と棒を、進み具合に合わせて置き直す（進み具合の更新のたびに呼ぶ）
    $ratio = getIndexingRatio
    foreach ($item in $script:targetItems) {
        if ($item.IndexLevel -eq "Run") {
            $row = getIndexRowView $null $true $item.Enabled $ratio (testIndexInRound $item)
            $item.SetIndexState($row.Text, $row.ToolTip, $row.Level, $row.Sub, [double]$row.Percent)
            setIndexRowActions $item
        }
    }
    updateIndexDetailPanel
}

function updateIndexFooter {
    # 一覧の下の帯。文言は判断層（getIndexFooterView）
    $view = getIndexFooterView $script:targetItems.Count $script:indexFileTotal
    $ui.IndexFooterFolders.Text = $view.Folders
    $ui.IndexFooterFiles.Text = $view.Files
}

function updateIndexListView {
    $empty = ($script:targetItems.Count -eq 0)
    $ui.IndexGridPlaceholder.Visibility = if ($empty) { "Visible" } else { "Collapsed" }
    # 1 件も無いときは、詳細と境目を出さず、案内を画面の中央に出す
    $ui.IndexDetailHost.Visibility = if ($empty) { "Collapsed" } else { "Visible" }
    $ui.IndexSplitter.Visibility = if ($empty) { "Collapsed" } else { "Visible" }
    # 詳細の行の高さは、空かどうかが変わったときだけ動かす（選択を変えるたびに入れ直すと、境目で変えた高さが戻る）。
    # 戻す高さの初めの値は index.xaml の RowDefinition（ここに数を持たない）
    $plan = getIndexDetailRowPlan $script:indexListEmpty $empty
    if ($plan -eq "Hide") {
        $script:indexDetailRowRestore = @{ Height = $ui.IndexDetailRow.Height; MinHeight = $ui.IndexDetailRow.MinHeight }
        $ui.IndexDetailRow.MinHeight = 0
        $ui.IndexDetailRow.Height = [System.Windows.GridLength]::new(0)
    }
    elseif ($plan -eq "Show") {
        if ($null -ne $script:indexDetailRowRestore) {
            $ui.IndexDetailRow.MinHeight = $script:indexDetailRowRestore.MinHeight
            $ui.IndexDetailRow.Height = $script:indexDetailRowRestore.Height
        }
    }
    $script:indexListEmpty = $empty
    updateIndexFooter
    updateIndexSelectionView
    updateIndexingButton
    updateIndexDetailPanel
}

function testIndexOperable {
    # インデックス作成・削除・エクスポート・インポートが動いている間はインデックスの追加・編集・削除・エクスポート・インポートをしない
    # （インデックスのフォルダ・取り込み一覧を使っているため）
    param (
        [string]$operation
    )

    $blocker = getIndexJobBlocker (isIndexing) $script:indexBusy $script:archiveBusy
    if ($blocker -ne "") {
        showMessage (getIndexJobBlockedMessage $blocker $operation) "OK" "Warning" | Out-Null
        return $false
    }
    return $true
}
