# ［1 インデックス管理］タブ（インデックスの一覧・追加・編集・削除）。

function getTargetsKey {
    # クロール対象フォルダの一覧（@{ Path; Enabled } の配列）を比べるための文字列
    param (
        [object[]]$folders
    )

    return (@($folders | Where-Object { $_ } | ForEach-Object { "$($_.Enabled)`t$($_.Path)" }) -join "`n")
}

# ============================================================================
# ［1 インデックス管理］（インデックスの追加・編集・削除と、インデックス作成の実行）
# ============================================================================

$script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexGrid.ItemsSource = $script:targetItems
$script:loadingTargets = $false
$script:indexBusy = $false  # 前のインデックスの削除中（別スレッド）。getIndexJobBlocker に [bool] で渡すため、$null のままにしない
# ［作成］チェックの状態が変わったら保存する（Checked・Unchecked。ToggleButton の状態が変わったときに出る、
# バブルするイベント）。マウスの Click だけでなく、UI オートメーションの TogglePattern（キーボード操作も同様）でも
# 状態が変わったときに出るため、どの操作でも保存できる。
# UI オートメーションの Toggle は IsChecked（表示）を変えるが、TwoWay バインドの先（Enabled。PS class の
# プレーンなプロパティで PropertyChanged を出さない）へは反映されないことがあるため、ここで明示的に合わせる。
# 読み込み時（loadTargets が Enabled をセットする間）は $script:loadingTargets を立てて、保存が走らないようにするが、
# DataGrid が行の見た目を作る（描画・仮想化）のは loadTargets の完了後で、その時点では $script:loadingTargets は
# 既に false に戻っている。行の初期化としての Checked・Unchecked（チェックの付いた行が表示される・［OK］で
# 追加した行がすぐ表示されるなど）でも出るため、一覧の中身が保存済みの内容と同じときは書き直さない
$onIndexGridToggled = {
    param ($s, $e)
    safe {
        $cb = $e.OriginalSource
        if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.DataContext -is [FolderItem] -and !$script:loadingTargets) {
            $cb.DataContext.Enabled = [bool]$cb.IsChecked
            if ((getTargetsKey @($script:targetItems)) -ne $script:savedTargets) {
                saveTargets
            }
            updateIndexingButton
        }
    }
}
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::CheckedEvent, [System.Windows.RoutedEventHandler]$onIndexGridToggled)
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::UncheckedEvent, [System.Windows.RoutedEventHandler]$onIndexGridToggled)
$script:savedTargets = $null  # 最後に読み込み・保存したインデックス一覧（getTargetsKey）。ほかでの変更の検出に使う
$script:editDialog = $null    # 追加・編集のダイアログ（開いている間だけ）
$script:indexingSession = $null  # 実行中のインデックス作成（IndexingSession。終わって片づけたら $null）
$script:indexingStart = $null
$script:ingestFailed = 0  # インデックス作成中に一覧へ反映済みの失敗件数
$script:indexingState = $null
$script:indexSummary = $null
$script:archiveBusy = $false  # エクスポート・インポート中（別スレッド）。settings_tab.ps1 の testWorkspaceChangeable も見る
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
    refreshFolderStatus
}

function refreshFolderStatus {
    # 一覧のすべてのフォルダの有無を別スレッドで調べ、表示と［インデックス作成を開始］の可否に反映する。
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
    # ［作成］チェックの保存は、一覧のチェックボックスの Checked・Unchecked（IndexGrid.AddHandler）で行う。
    # PS class のプレーンなプロパティは TwoWay セットで PropertyChanged を出さないため、購読では拾えない。
    return $item
}

function loadTargets {
    $script:loadingTargets = $true
    try {
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
            $script:targetItems.Add((newFolderItem $folder.Path $folder.Enabled $folder.Name))
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
}

function applyIndexStats {
    # 取り込み一覧の集計（getIndexStats）を一覧の各行のファイル数・最終取り込み・「ステータス」列に反映し、
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
            $item.SetStats("－", "まだ取り込んでいません", "")
        } else {
            $ingested = [datetime]::MinValue
            $lastText = if ($stat.LastIngested -and [datetime]::TryParseExact($stat.LastIngested, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ingested)) {
                formatTime $ingested
            } else {
                ""
            }
            $item.SetStats(("{0:#,0}" -f $stat.Total), ("済 {0:#,0} 件 ・ 未取り込み {1:#,0} 件 ・ 失敗 {2:#,0} 件" -f $stat.Done, $stat.Pending, $stat.Failed), $lastText)
        }
        $row = getIndexRowView $stat $indexing $item.Enabled
        $item.SetIndexState($row.Text, $row.ToolTip, $row.Level)
    }
    updateFastSearchRows
}

function updateIndexListView {
    $ui.IndexGridPlaceholder.Visibility = if ($script:targetItems.Count -eq 0) { "Visible" } else { "Collapsed" }
    updateIndexingButton
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

function showIndexEditDialog {
    # インデックスの追加・編集のダイアログ。決めた内容 @{ Path; Name } を返す（キャンセルは $null）。
    #   item: 編集するインデックス（$null なら追加）
    param (
        $item = $null
    )

    $dialog = loadWindow "${xamlDir}\dialog_index_edit.xaml"
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("OkButton", "BrowseButton", "FolderBox", "NameBox", "IntroText", "NoticeText", "ErrorText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $script:editDialog = @{ Window = $dialog; Ctrl = $ctrl; Item = $item; Suggested = "" }

    if ($null -eq $item) {
        $dialog.Title = "インデックスの追加"
        $ctrl.IntroText.Text = "Office ファイル（Excel・Word・PowerPoint）の入っているフォルダを 1 つ選んでください。" +
            "ここでは一覧に加えるだけです。中のファイルを読むのは［インデックス作成を開始］を押してからです。"
    } else {
        $dialog.Title = "インデックスの編集"
        $ctrl.IntroText.Text = "名前と、元のフォルダの場所を変えられます。"
        $ctrl.FolderBox.Text = $item.Path
        $ctrl.NameBox.Text = $item.Name
        $ctrl.NoticeText.Visibility = "Visible"
        $ctrl.NoticeText.Text = "変えるのは名前と場所だけです。インデックスはそのまま使います（作り直しません）。" +
            "フォルダを別のドライブや共有フォルダへ移したときは、ここで新しい場所を指定してください。"
    }

    $ctrl.FolderBox.Add_TextChanged({
        safe {
            # 追加のときは、フォルダ名からインデックス名を自動で入れる（利用者が名前を変えた後は触らない）
            $d = $script:editDialog
            if ($null -ne $d.Item -or ($d.Ctrl.NameBox.Text -ne "" -and $d.Ctrl.NameBox.Text -ne $d.Suggested)) {
                return
            }
            $path = normalizeFolderPath $d.Ctrl.FolderBox.Text
            $d.Suggested = if ($path -eq "") { "" } else { newIndexName $path (getUsedIndexNames $script:targetItems) }
            $d.Ctrl.NameBox.Text = $d.Suggested
        }
    })
    $ctrl.BrowseButton.Add_Click({
        safe {
            $d = $script:editDialog
            $initial = normalizeFolderPath $d.Ctrl.FolderBox.Text
            # ネットワークのパスは画面のスレッドで有無を調べない。編集中のインデックスの元のフォルダで、
            # 場所を書き換えていない（直前の refreshFolderStatus の結果がそのまま使える）ときだけ、調べずに開始フォルダにする
            $knownExisting = ($null -ne $d.Item) -and ($d.Item.Path -eq $initial) -and $d.Item.StatusChecked -and $d.Item.FolderExists
            $path = selectFolder "インデックスにする、Office ファイルのあるフォルダを選んでください" $initial $d.Window $knownExisting
            if ($path) {
                $d.Ctrl.FolderBox.Text = $path
            }
        }
    })
    $ctrl.FolderBox.Add_PreviewDragOver({ onFolderDragOver @args })
    $ctrl.FolderBox.Add_PreviewDrop({
        param ($sender, $e)
        safe {
            $folders = @(getDroppedFolders $e)
            if ($folders.Count -gt 0) {
                $script:editDialog.Ctrl.FolderBox.Text = $folders[0]
            }
        }
        $e.Handled = $true
    })
    $ctrl.OkButton.Add_Click({
        safe {
            $d = $script:editDialog
            $message = checkIndexEditInput
            if ($message -ne "") {
                $d.Ctrl.ErrorText.Text = $message
                $d.Ctrl.ErrorText.Visibility = "Visible"
                return
            }
            $d.Window.DialogResult = $true
        }
    })

    $result = $null
    if ($dialog.ShowDialog()) {
        $result = @{ Path = (normalizeFolderPath $ctrl.FolderBox.Text); Name = $ctrl.NameBox.Text.Trim() }
    }
    $script:editDialog = $null
    return $result
}

function checkIndexEditInput {
    # 追加・編集のダイアログの入力を調べ、直してほしい内容を返す（問題なければ空文字列）
    $d = $script:editDialog
    return (testIndexEditInput $d.Ctrl.FolderBox.Text $d.Ctrl.NameBox.Text $script:targetItems $d.Item)
}

function addIndexItem {
    # インデックスを一覧に加えて保存する
    param (
        [string]$path,
        [string]$name
    )

    $item = newFolderItem $path $true $name
    $script:targetItems.Add($item)
    refreshFolderStatus
    $ui.IndexGrid.SelectedItem = $item
    $ui.IndexGrid.ScrollIntoView($item)
    saveTargets
    updateIndexSourceFile
    updateIndexListView
    refreshIndexingState
    refreshFastSearchStatus
    # フォルダの有無は refreshFolderStatus（別スレッド）が調べるので、ここでは Test-Path を呼ばない
    # （届かないネットワークのフォルダで画面のスレッドが止まらないようにする）。有無は一覧の列で分かる
    setStatus (getIndexAddedStatus $name)
}

function newIndex {
    # ［追加…］。フォルダとインデックス名を決めて一覧に加える（インデックス作成はしない）
    if (!(testIndexOperable "追加")) {
        return
    }
    $result = showIndexEditDialog $null
    if ($null -eq $result) {
        return
    }
    addIndexItem $result.Path $result.Name
}

function addIndexForFolder {
    # 一覧へのドラッグ＆ドロップでインデックスを追加する（名前はフォルダ名から自動で決める）
    param (
        [string]$path
    )

    if (!(testIndexOperable "追加")) {
        return
    }
    $path = normalizeFolderPath $path
    if ($path -eq "") {
        return
    }
    foreach ($item in $script:targetItems) {
        # 書き方が違うだけで同じフォルダ（ネットワークドライブと UNC パスなど）も、すでにあるとみなす
        if (testSameFolder $item.Path $path) {
            $ui.IndexGrid.SelectedItem = $item
            setStatus "「$($item.Path)」のインデックス [$($item.Name)] は既にあります"
            return
        }
    }
    addIndexItem $path (newIndexName $path (getUsedIndexNames $script:targetItems))
}

function editIndex {
    # ［編集…］。インデックス名と元のフォルダの場所を変える。インデックスは作り直さない
    $item = $ui.IndexGrid.SelectedItem
    if ($null -eq $item -or !(testIndexOperable "編集")) {
        return
    }
    $result = showIndexEditDialog $item
    if ($null -eq $result) {
        return
    }

    $changes = New-Object System.Collections.Generic.List[string]
    # 大文字・小文字だけの変更も改名する（-ne は大文字・小文字を区別しないため -cne で比べる）
    if ($result.Name -cne $item.Name) {
        # インデックスのフォルダ（work\index\<名前>）と取り込み一覧の記録も名前を変える（中身は作り直さない）
        renameIndex $item.Name $result.Name
        $changes.Add("名前 [$($item.Name)] → [$($result.Name)]")
        $item.SetName($result.Name)
    }
    if ($result.Path -ne $item.Path) {
        $changes.Add("場所 $($item.Path) → $($result.Path)")
        $item.SetPath($result.Path)
        $item.StatusChecked = $false
        updateFolderItemStatus $item
    }
    if ($changes.Count -eq 0) {
        return
    }

    saveTargets
    updateIndexSourceFile
    refreshIndexViews
    setStatus ("インデックスを変更しました（" + ($changes -join " / ") + "）")
}

function startIndexRemoveJob {
    # インデックス（work\index\<名前>）と取り込み一覧の記録の削除を別スレッドで行う。
    # 数万フォルダの削除は数十秒かかることがあり、画面のスレッドで行うと「応答なし」になるため。
    # 終わるまでインデックスの操作・インデックス作成の開始はできないようにし、何をしているかをステータスに出す
    param (
        [string]$name,
        [string]$operation,   # "削除"（表示に使う）
        [scriptblock]$onDone  # 削除が終わった後に画面のスレッドで行うこと（$script:indexJobName で名前を参照できる）
    )

    $script:indexBusy = $true
    $script:indexJobName = $name
    $script:indexJobOnDone = $onDone
    $script:indexJobOperation = $operation
    updateIndexingButton
    setStatus "インデックス [${name}] を削除しています…（件数によっては少し時間がかかります）"
    # 裏のスレッドは lib.ps1 を読み込んだときのワークスペースを覚えているため、場所は渡す
    startJob {
        param ($name, $dir, $statusPath, $settingsPath)
        removeIndex $name $dir $statusPath $settingsPath
    } @($name, $workspace.IndexDir, $workspace.StatusFile, ${settingsFile}) {
        param ($output, $errorText)
        $script:indexBusy = $false
        updateIndexingButton
        $name = $script:indexJobName
        if ($errorText) {
            setStatus "インデックス [${name}] の $($script:indexJobOperation)に失敗しました：${errorText}"
            refreshIndexViews
            return
        }
        refreshIndexViews
        if ($script:indexJobOnDone) {
            & $script:indexJobOnDone
        }
    }
}

function deleteIndex {
    # ［削除］。一覧から削除し、インデックス（work\index\<名前>）と取り込み一覧の記録も削除する
    $item = $ui.IndexGrid.SelectedItem
    if ($null -eq $item -or !(testIndexOperable "削除")) {
        return
    }

    $answer = showConfirm `
        -heading "インデックス「$($item.Name)」を一覧から削除しますか？" `
        -facts @(
            (factGone "tebunko が作ったインデックスが消えます" "このフォルダは検索できなくなります（もう一度［インデックス作成を開始］すれば作り直せます）"),
            (factKept "元のフォルダと、その中のファイルはそのままです" $item.Path)
        ) `
        -hint "しばらく検索しないだけなら、削除せずに［作成］のチェックを外してください。インデックスは残ったままです。" `
        -choices @(@{ Text = "削除する"; Value = "delete"; Danger = $true })
    if ($answer -ne "delete") {
        return
    }

    # 一覧からはすぐ消し、インデックスの削除（時間がかかることがある）は別スレッドで行う
    $script:targetItems.Remove($item)
    saveTargets
    updateIndexSourceFile
    updateIndexListView
    startIndexRemoveJob $item.Name "削除" {
        setStatus "インデックス [$($script:indexJobName)] を削除しました"
    }
}

function startIndexArchiveJob {
    # エクスポート・インポートを別スレッドで行う。終わるまでインデックスの操作・インデックス作成の開始・
    # もう一方のエクスポート・インポート・ワークスペースの変更はできないようにする（$script:archiveBusy）。
    # scriptBlock（2 つ目の引数）はこの関数が startJob に渡す、裏で動く仕事。
    # tests/meta/ui_thread_io.Tests.ps1 の getStartJobExcludedRanges が、この 2 つ目の引数を startJob の 1 つ目と同じ扱いで除く
    param (
        [string]$operation,          # "エクスポート" / "インポート"（表示に使う）
        [scriptblock]$scriptBlock,
        [object[]]$arguments,
        [scriptblock]$onSuccess      # 成功したときに画面のスレッドで行うこと { param($output) }
    )

    $script:archiveBusy = $true
    $script:archiveJobOperation = $operation
    $script:archiveJobOnSuccess = $onSuccess
    updateIndexingButton
    setStatus "${operation}しています…"
    startJob $scriptBlock $arguments {
        param ($output, $errorText)
        $script:archiveBusy = $false
        if ($errorText) {
            setStatus "$($script:archiveJobOperation)に失敗しました：${errorText}"
            updateIndexingButton
            return
        }
        if ($output -and $output.Count -gt 0 -and $script:archiveJobOnSuccess) {
            & $script:archiveJobOnSuccess $output[0]
        }
        updateIndexingButton
    }
}

function newExportIndex {
    # ［エクスポート…］。選んだインデックスを 1 つの zip に書き出す
    $item = $ui.IndexGrid.SelectedItem
    if ($null -eq $item -or !(testIndexOperable "エクスポート")) {
        return
    }
    $folder = selectFolder "エクスポート先のフォルダを選んでください" $workspace.Dir
    if ($null -eq $folder) {
        return
    }
    $name = $item.Name

    # 保存先フォルダの中身を数えるのも別スレッドで行う（届かないネットワークのフォルダで画面が止まらないように）
    startIndexArchiveJob "エクスポート" {
        param ($name, $folder, $dir, $statusPath, $settingsPath)
        $used = @(Get-ChildItem -LiteralPath (toLongPath $folder) -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
        $destPath = Join-Path $folder (getExportFileName $name $used)
        $ws = [Workspace]::new($dir)
        exportIndex $name $destPath $ws $settingsPath
    } @($name, $folder, $workspace.Dir, $workspace.StatusFile, ${settingsFile}) {
        param ($result)
        setStatus "インデックス [$($result.Name)] を「$($result.Path)」に書き出しました（$($result.Files) ファイル）"
    }
}

function showIndexImportDialog {
    # インポートのダイアログ（追加・編集と同じ見た目の XAML を使う）。決めた内容 @{ Path; Name } を返す（キャンセルは $null）
    param (
        [string]$suggestedName,
        $info   # readIndexArchiveInfo の結果
    )

    $dialog = loadWindow "${xamlDir}\dialog_index_edit.xaml"
    $dialog.Owner = $window
    $dialog.Title = "インデックスのインポート"
    $ctrl = @{}
    foreach ($name in @("OkButton", "BrowseButton", "FolderBox", "NameBox", "IntroText", "NoticeText", "ErrorText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $ctrl.IntroText.Text = getIndexImportNotice $info
    $ctrl.FolderBox.Text = $info.SourceFolder
    $ctrl.NameBox.Text = $suggestedName
    $ctrl.NoticeText.Visibility = "Collapsed"

    # 追加・編集のダイアログ（editIndex の $script:editDialog）と同じく、ボタンの Click からは
    # script スコープの入れ物を参照する。GetNewClosure() でこの関数のローカル変数（$ctrl・$dialog）を
    # 取り込むと、tebunko.bat の起動（powershell -Command "...; & gui.ps1"）のように呼び出しが
    # 入れ子になっている実機では、閉じ込めたスクリプトブロックから名前で関数を解決できなくなるため
    # （continueImportIndex で見つかった不具合と同じ原因。docs\design\gui\responsiveness.md の
    # 「画面を固まらせない待たせ方」の注意（PowerShell 5.1）を参照）、ここでは使わない
    $script:importDialog = @{ Window = $dialog; Ctrl = $ctrl }

    $ctrl.BrowseButton.Add_Click({
        safe {
            $d = $script:importDialog
            $initial = normalizeFolderPath $d.Ctrl.FolderBox.Text
            $path = selectFolder "インポートしたインデックスの、今の元のフォルダを選んでください" $initial $d.Window $false
            if ($path) {
                $d.Ctrl.FolderBox.Text = $path
            }
        }
    })
    $ctrl.OkButton.Add_Click({
        safe {
            $d = $script:importDialog
            $message = testIndexImportInput $d.Ctrl.FolderBox.Text $d.Ctrl.NameBox.Text $script:targetItems
            if ($message -ne "") {
                $d.Ctrl.ErrorText.Text = $message
                $d.Ctrl.ErrorText.Visibility = "Visible"
                return
            }
            $d.Window.DialogResult = $true
        }
    })

    $result = $null
    if ($dialog.ShowDialog()) {
        $result = @{ Path = (normalizeFolderPath $ctrl.FolderBox.Text); Name = $ctrl.NameBox.Text.Trim() }
    }
    return $result
}

function newImportIndex {
    # ［インポート…］。zip を選ぶ。目録を読むのは別スレッド（届かないネットワークの zip で画面が止まらないように）
    if (!(testIndexOperable "インポート")) {
        return
    }
    $zipPath = selectZipFile "インポートする zip を選んでください"
    if ($null -eq $zipPath) {
        return
    }

    # onSuccess は別のスレッドの完了を受けて、あとから（別の呼び出しの流れで）実行される。
    # $zipPath（この関数のローカル変数）を参照するため、GetNewClosure() で今の値を取り込む。
    # continueImportIndex は名前のまま呼ばず、openFailedFileFolder と同じく ${function:...} で先に
    # 関数の実体を変数に取り込んでから & で呼ぶ（GetNewClosure() したスクリプトブロックは、
    # tebunko.bat の起動（powershell -Command "...; & gui.ps1"）のように呼び出しが入れ子になっていると、
    # 名前による関数の解決ができないため）
    $continueImport = ${function:continueImportIndex}
    startIndexArchiveJob "インポート" {
        param ($zipPath)
        readIndexArchiveInfo $zipPath
    } @($zipPath) {
        param ($info)
        & $continueImport $zipPath $info
    }.GetNewClosure()
}

function continueImportIndex {
    # newImportIndex の続き（目録を読み終えたところから）。名前・元のフォルダを決めて一覧に加える
    param (
        [string]$zipPath,
        $info   # readIndexArchiveInfo の結果
    )

    # @(getUsedIndexNames ...) と直接書くと集合が 1 要素の配列に入るだけで、名前の重なりを見落とす。変数に受けてから使う
    $usedNames = getUsedIndexNames $script:targetItems
    $suggestedName = getImportSuggestedName $info.IndexName $usedNames
    $result = showIndexImportDialog $suggestedName $info
    if ($null -eq $result) {
        setStatus "インポートを取りやめました"
        return
    }

    $collisionMode = ${importCollisionRename}
    if (testImportNameCollision $result.Name $usedNames) {
        $answer = showConfirm `
            -heading (getIndexImportOverwriteConfirmMessage $result.Name) `
            -choices @(
                @{ Text = "上書きする"; Value = "overwrite"; Danger = $true },
                @{ Text = "別名で入れる"; Value = "rename" }
            )
        if ($null -eq $answer) {
            setStatus "インポートを取りやめました"
            return
        }
        $collisionMode = if ($answer -eq "overwrite") { ${importCollisionOverwrite} } else { ${importCollisionRename} }
    }

    startIndexArchiveJob "インポート" {
        param ($zipPath, $collisionMode, $name, $sourceFolder, $dir, $settingsPath)
        $ws = [Workspace]::new($dir)
        importIndex $zipPath $collisionMode $name $sourceFolder $ws $settingsPath
    } @($zipPath, $collisionMode, $result.Name, $result.Path, $workspace.Dir, ${settingsFile}) {
        param ($result)
        if ($null -eq $result) {
            return  # Cancel で置き換えなかった
        }
        # インポートは別スレッドで設定・取り込み一覧を直接書き換えるため、一覧を読み直す（$script:targetItems には無い変更）
        loadTargets
        refreshIndexViews
        setStatus (getImportResultStatus $result)
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

# ---- イベント ----

$ui.NewIndexButton.Add_Click({ safe { newIndex } })
$ui.EditIndexButton.Add_Click({ safe { editIndex } })
$ui.RemoveIndexButton.Add_Click({ safe { deleteIndex } })
$ui.ExportIndexButton.Add_Click({ safe { newExportIndex } })
$ui.ImportIndexButton.Add_Click({ safe { newImportIndex } })
$ui.IndexGrid.Add_SelectionChanged({ safe { updateIndexListView } })
$ui.IndexGrid.Add_MouseDoubleClick({ safe { editIndex } })
$ui.IndexGrid.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Delete") {
        safe { deleteIndex }
    }
})
$ui.IndexGrid.Add_PreviewDragOver({ onFolderDragOver @args })
$ui.IndexGrid.Add_PreviewDrop({
    param ($sender, $e)
    safe {
        foreach ($folder in (getDroppedFolders $e)) {
            addIndexForFolder $folder
        }
    }
    $e.Handled = $true
})
