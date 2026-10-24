# インデックス管理の画面の一覧（状態の保持・フォルダの状態の調べ・一覧の読み込みと更新）。

function getTargetsKey {
    # クロール対象フォルダの一覧（@{ Path; Enabled } の配列）を比べるための文字列
    param (
        [object[]]$folders
    )

    return (@($folders | Where-Object { $_ } | ForEach-Object { "$($_.Enabled)`t$($_.Path)" }) -join "`n")
}

$script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexGrid.ItemsSource = $script:targetItems
# 1 件も無いときの［＋ フォルダを追加］。名前の一覧（shell\regions.ps1）に足さず、読み込んだ中身から取る（件数の帯は窓の下のステータスバーにある）
foreach ($footerName in @("IndexEmptyAddButton")) {
    $ui[$footerName] = $ui.IndexListHost.Content.FindName($footerName)
}
$script:loadingTargets = $false
$script:targetsLoadRequest = @{ Value = 0 }  # 一覧の読み込みの依頼番号（新しい依頼が出たら、前の依頼の結果は捨てる）
$script:leftoverIndexes = @()  # ワークスペースに残っているインデックスの「名前と前回のフォルダの組」（getWorkspaceIndexFolders）。名前を付ける入口が使う
$script:leftoverJobDir = $null  # 上を最後に読み始めたときのワークスペース（$workspace.Dir）
$script:indexingPreparing = $false  # インデックス作成を始める前に、ネットワークのワークスペースを裏で確かめている間
$script:indexBusy = $false  # 前のインデックスの削除・名前の変更中（別スレッド）。getIndexJobBlocker に [bool] で渡すため、$null のままにしない
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
$script:detailEditItem = $null   # 詳細の名前・フォルダパスの欄で直している行（updateIndexDetailPanel が置く）
$script:indexingSession = $null  # 実行中のインデックス作成（IndexingSession。終わって片づけたら $null）
$script:indexingOnlyNames = @()  # 今の回が選んだものだけのときの、インデックス名（空なら全部）
$script:indexingStart = $null
$script:ingestFailed = 0  # インデックス作成中に一覧へ反映済みの失敗件数
$script:indexingState = $null
$script:indexSummary = $null
$script:targetsNaming = $false  # ネットワークのワークスペースで、一覧のインデックス名を裏で決めている間（loadTargets）。この間は一覧が空で、追加・削除・インデックス作成の開始・ワークスペースの変更をしない
$script:archiveBusy = $false  # エクスポート・インポート中（別スレッド）。settings\settings.ps1 の testWorkspaceChangeable も見る
$script:workspaceMoving = $false  # ワークスペースの中身を移している間（別スレッド。settings\settings.ps1 の startWorkspaceMove）。indexBusy とは分ける。閉じるのは移し終わるまで待つ
$script:closeAfterWorkspaceMove = $false  # 移している間に窓を閉じようとした（移し終わったら閉じる）
$script:archiveJobOperation = ""
$script:archiveJobOnSuccess = $null

function isIndexing {
    return ($null -ne $script:indexingSession) -and $script:indexingSession.IsRunning()
}

function isIndexingOrPreparing {
    # インデックス作成が動いている、または始める前にワークスペースを確かめている間（startIndexing）。
    # どちらの間も、インデックスの操作・ワークスペースの変更・次のインデックス作成の開始はしない
    return (isIndexing) -or [bool]$script:indexingPreparing
}

function getCurrentIndexJobBlocker {
    # いま動いている処理の名前（getIndexJobBlocker の結果。無ければ空文字列）。画面の状態から集める。
    #   extraIndexing: ほかに「インデックス作成中」とみなすもの（画面を使わずに起動したインデクサなど）
    #   indexerUnchecked: ほかのインデックス作成が動いていないかを確かめられなかった（testIndexerRunning の Error）
    param (
        [bool]$extraIndexing = $false,
        [bool]$indexerUnchecked = $false
    )

    return (getIndexJobBlocker ((isIndexingOrPreparing) -or $extraIndexing) ([bool]$script:indexBusy) ([bool]$script:archiveBusy) ([bool]$script:targetsNaming) ([bool]$script:workspaceMoving) $indexerUnchecked)
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
    $queue = getWorkspaceJobQueue $paths
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
            $item.SetStatus("フォルダがあります", ${okBrush})
        } else {
            $item.SetStatus("フォルダが見つかりません", ${ngBrush})
        }
    }
    updateIndexingButton
    updateIndexDetailPanel
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
    # 設定のインデックス一覧を画面の一覧に読み込む。
    # 名前の決まっていないインデックス（設定ファイルを直接書き換えた場合など）には、ここで名前を割り当てて確定する。
    # 一覧・編集・削除はインデックス名で扱うため、画面に出す時点で名前があるようにする（インデクサと同じ assignIndexNames を使う）。
    # 名前を決めるには取り込み一覧（ワークスペースの中）を読み、保存する。ローカルでもネットワークでも裏の列で行い、
    # 終わるまで一覧は空のまま（数に比例する読み書きや、届かない場所で画面のスレッドが止まらないように）
    $requestBox = $script:targetsLoadRequest
    $requestBox.Value++
    $script:targetsNaming = $false
    # 読み直しても、チェック（一時の選択）は名前で引き継ぐ
    $checkedNames = @(getIndexCheckedItems @($script:targetItems) | ForEach-Object { $_.Name })
    $folders = @(getTargetFolders)
    if (@($folders | Where-Object { $_ -and !$_.Name }).Count -gt 0) {
        $script:targetItems.Clear()
        # 名前が決まって一覧を読み込み直すまで、ほかの操作を止める（空の一覧に保存して、設定のほかのインデックスを消さないため）
        $script:targetsNaming = $true
        try {
            updateIndexListView
            updateIndexingButton
            setStatus (getIndexNamingStatus)
            $requestId = $requestBox.Value
            $finish = ${function:finishTargetsNaming}   # 終わったときの処理は、関数を変数に取って呼ぶ（クロージャからは関数の名前を引けないため）
            startJob {
                param ($folders, $workspaceDir)
                # インスタンスは裏のスレッドへ渡せないため、場所から作り直す
                [void]((newIndexCatalog $workspaceDir).AssignAndSaveNames($folders))
            } @(,$folders + @($workspace.Dir)) {
                param ($output, $errorText)
                if ($requestId -ne $requestBox.Value) {
                    # 待っている間にもう一度読み込んだ・ワークスペースを切り替えた。この結果は捨てる
                    return
                }
                & $finish $errorText $checkedNames
            }.GetNewClosure() (getWorkspaceJobQueue $workspace.Dir)
        } catch {
            # 裏の仕事を出せなかったときは、印を戻す（残ると、ほかの操作が止まったままになる）
            $script:targetsNaming = $false
            throw
        }
        return
    }
    fillTargetItems $folders $checkedNames
}

function refreshLeftoverIndexes {
    # 設定から外れたまま残してあるインデックス（名前と前回のフォルダの組）を、別スレッドで読み直して控える。
    # 追加・名前の変更・インポートが、その名前を別のフォルダに付けないために使う（付けると次の更新で前の中身が消える）。
    # 入力のたびにワークスペースを読むと、ネットワークの作業フォルダで画面が止まるため、一覧を読み直すところで 1 回だけ読む。
    # 届かないときは空にして画面は止めない。読んでいる間にワークスペースを変えたときは、結果を捨てる
    # （仕事は順に動くので、続けて呼んだときは最後に読んだ結果が残る）
    $dir = [string]$workspace.Dir
    if ($script:leftoverJobDir -ne $dir) {
        # ワークスペースを替えたときは、前の場所の控えを使わないよう空に戻す
        $script:leftoverIndexes = @()
    }
    $script:leftoverJobDir = $dir
    startJob {
        param ($workspaceDir)
        $ws = [Workspace]::new($workspaceDir)
        foreach ($folder in @(getWorkspaceIndexFolders $ws)) {
            @{ Name = [string]$folder.Name; Path = [string]$folder.Path }
        }
    } @($dir) {
        param ($output, $errorText)
        if ($script:leftoverJobDir -ne [string]$workspace.Dir) {
            return
        }
        $script:leftoverIndexes = if ($errorText) { @() } else { @($output | Where-Object { $_ } | ForEach-Object { [pscustomobject]@{ Name = [string]$_.Name; Path = [string]$_.Path } }) }
    } (getWorkspaceJobQueue $dir)
}

function finishTargetsNaming {
    # ネットワークのワークスペースで、裏の名前の割り当てが終わったときの処理。保存し直された一覧を読み込む（失敗したら理由を出し、設定のまま）
    param (
        [string]$errorText,
        [string[]]$checkedNames
    )

    $script:targetsNaming = $false
    if ($errorText) {
        setStatus (getIndexNamingFailedStatus $errorText)
    }
    fillTargetItems @(getTargetFolders) $checkedNames
    updateIndexingButton
}

function fillTargetItems {
    # 画面の一覧を folders（@{ Name; Path; Enabled }）の内容に置き換える。checkedNames の名前の行はチェックを付ける
    param (
        [object[]]$folders,
        [string[]]$checkedNames
    )

    $script:loadingTargets = $true
    try {
        $script:targetItems.Clear()
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
    refreshLeftoverIndexes
}

function saveTargets {
    # 一覧を設定に書く。保存した旨のステータスは出す。-Quiet のときは出さない（更新を始めるときの保存は、更新を始めるか確認で決まるため、文を出さない）
    param ([switch]$Quiet)

    (newIndexCatalog $workspace.Dir).SaveTargets(@($script:targetItems | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Enabled = $_.Enabled } }))
    $script:savedTargets = getTargetsKey @(getTargetFolders)
    if (!$Quiet) {
        setStatus "インデックス一覧を保存しました（$(Get-Date -Format 'H:mm')）"
    }
}

function updateIndexSourceFile {
    # インデックスのフォルダの source_folder.txt を今の一覧に合わせて書き直す。
    # 次のインデックス作成を待たずに、検索結果から元のファイルを開けるようにする（インデックスが無いフォルダには書かない）。
    # 書くのは別スレッドで行う（ワークスペースが届かない共有フォルダにあっても、画面のスレッドが止まらないように）。
    # 書けなかったときはステータスに知らせる
    $folders = @($script:targetItems | Where-Object { $_.Name } | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path } })
    $dir = [string]$workspace.IndexDir
    startJob {
        param ($folders, $workspaceDir)
        (newIndexCatalog $workspaceDir).WriteSourceFile($folders)
    } @(,$folders + @([string]$workspace.Dir)) {
        param ($output, $errorText)
        if ($errorText) {
            setStatus (getSourceFolderFileFailedStatus $errorText)
        }
    } (getWorkspaceJobQueue $dir)
}

function refreshIndexViews {
    # インデックスを追加・編集・削除した後、検索タブ（検索対象のツリー・件数）も読み直す。
    $script:sourceFolderMaps = @{}
    $script:indexSummary = $null
    loadIndexTree
    refreshIndexSummary
    refreshIndexingState
    refreshLeftoverIndexes
    updateIndexDetailPanel
}

function refreshIndexViewsKeepingStatus {
    # 失敗の知らせ（$notice）を出したまま、画面を読み直す。検索対象のツリーは裏で読み直され、
    # 届かないときはその結果がステータスを「接続できません」で上書きするため、その 1 回はステータスを書かせない
    # （接続できない旨は検索対象の欄に出る）。読み直しの開始でもステータスが変わりうるため、開始のあとにも書く
    param (
        [string]$notice
    )

    $script:indexTreeKeepStatus = $true
    refreshIndexViews
    setStatus $notice
}

function applyIndexStats {
    # 取り込み一覧の集計（getIndexStats）を一覧の各行のファイル数・最終更新・「ステータス」列に反映する。
    #   indexing: インデックス作成中か（getIndexRowView にそのまま渡す）
    param (
        $stats,
        [bool]$indexing = $false,
        $formatState = $null    # getWorkspaceIndexFormat の結果（インデックスの格納形式の状態）。無ければ $null
    )

    foreach ($item in $script:targetItems) {
        $stat = $null
        $formatEntry = $null
        if ($item.Name -and $null -ne $formatState -and $formatState.Indexes.Contains($item.Name)) {
            $formatEntry = $formatState.Indexes[$item.Name]
        }
        if ($item.Name -and $null -ne $stats -and $stats.ContainsKey($item.Name)) {
            $stat = $stats[$item.Name]
        }
        if ($null -eq $stat) {
            $item.SetStats("－", "まだ更新していません", "")
            $item.FailedCount = 0
            $item.PendingCount = 0
        } else {
            $ingested = [datetime]::MinValue
            $lastText = if ($stat.LastIngested -and [datetime]::TryParseExact($stat.LastIngested, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ingested)) {
                formatTime $ingested
            } else {
                ""
            }
            $item.FailedCount = [int]$stat.Failed
            $item.PendingCount = [int]$stat.Pending
            $item.SetStats(("{0:#,0}" -f $stat.Total), ("更新済み {0:#,0} 件 ・ 未更新 {1:#,0} 件 ・ 失敗 {2:#,0} 件" -f $stat.Done, $stat.Pending, $stat.Failed), $lastText)
        }
        $row = getIndexRowView $stat $indexing $item.Enabled (getIndexingRatio) (testIndexInRound $item) (testIndexExporting $item) $formatEntry
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
    updateIndexDetailPanel
}

$script:indexFileTotal = 0

function testIndexInRound {
    # 今の回で更新する行か（選んだものだけの回では、選んだ名前の行だけ）
    param ($item)

    return (@($script:indexingOnlyNames).Count -eq 0 -or @($script:indexingOnlyNames) -contains $item.Name)
}

$script:exportingNames = @()  # エクスポートで書き出している最中のインデックスの名前（行のバッジを「エクスポート中」にする）

function testIndexExporting {
    # 行がエクスポートで書き出している最中か
    param ($item)

    return (@($script:exportingNames) -contains [string]$item.Name)
}

function setExportingRows {
    # エクスポートで書き出している最中の名前を置き（空なら終わり）、一覧の行のバッジを置き直す
    param ([string[]]$names)

    $script:exportingNames = @($names | Where-Object { $_ })
    if ($null -ne $script:indexingState) {
        applyIndexStats $script:indexingState.IndexStats (isIndexing) $script:indexingState.Format
    }
}

function setIndexRowActions {
    # 行の右端のボタン（［更新］［中止］・何も出さない）を、行の状態（IndexLevel）と動いている処理に合わせる。判断は getIndexRowActions
    param ($item)

    $blocker = getCurrentIndexJobBlocker
    $actions = getIndexRowActions $item.IndexLevel $blocker (testIndexExporting $item)
    $item.SetRowActions($actions.Action, $actions.UpdateEnabled)
}

function updateSelectedIndexes {
    # 選んだインデックス（名前の配列）だけを更新する。行の［更新］と［アクション ▼］の［更新］の共通の入口。
    # 選んだ名前の行を、設定の enabled を true にして取り込み、その回だけの更新を始める（startIndexing）。
    # 動いている処理があるときは始めない（ボタンも使えなくしてある）
    param (
        [string[]]$names
    )

    $names = @($names | Where-Object { $_ })
    if ($names.Count -eq 0 -or (getCurrentIndexJobBlocker) -ne "") {
        return
    }
    startIndexing $names
}

$script:indexSourceRequest = [ref]0  # ［元のフォルダを開く］の確かめの依頼の番号（openListedFileFolder と同じ理由で [ref] のまま閉じ込める）

function applyIndexSourceFolderState {
    # openIndexSourceFolder の続き（getPathState の結果を画面に反映する）。フォルダが見つかれば、エクスプローラーで開く
    param (
        $state,
        [string]$path
    )

    $kind = getIndexSourceFolderKind $state
    if ($kind -eq "Open") {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"${path}`""
        return
    }
    setStatus (getIndexSourceFolderStatus $kind $path $state.Message)
}

function openIndexSourceFolder {
    # 行の右クリックの［元のフォルダを開く］。押した行の元のフォルダをエクスプローラーで開く。
    # フォルダが無い・つながらないときは、押したときにステータスバーで知らせる（検索結果の［フォルダを開く］と同じ。無効にはしない）。
    # ネットワークにあれば裏のスレッドで確かめ、画面のスレッドは待たない
    $item = getIndexTargetItem
    if ($null -eq $item) {
        return
    }
    $path = [string]$item.Path
    if (!$path) {
        setStatus (getIndexSourceFolderStatus "Empty" "")
        return
    }
    $requestBox = $script:indexSourceRequest
    $requestBox.Value++
    $requestId = $requestBox.Value
    $applyState = ${function:applyIndexSourceFolderState}
    $apply = {
        param ($state)
        if ($requestId -ne $requestBox.Value) {
            # 待っている間に別の行を開いた。前の依頼は捨てる
            return
        }
        & $applyState $state $path
    }.GetNewClosure()

    if (!(testNetworkPath $path)) {
        & $apply (getPathState $path)
        return
    }
    setStatus (getIndexSourceFolderStatus "Checking" $path)
    $otherState = ${pathStateOther}
    startJob {
        param ($path)
        getPathState $path
    } @($path) {
        param ($output, $errorText)
        if ($errorText) {
            & $apply @{ State = $otherState; IsDirectory = $false; Message = $errorText }
        } else {
            & $apply $output[0]
        }
    }.GetNewClosure() (getWorkspaceJobQueue $path)
}

$script:indexLocationRequest = [ref]0  # 保存先の［フォルダを開く］の確かめの依頼の番号（openIndexSourceFolder と同じ理由で [ref] のまま閉じ込める）

function updateIndexLocationView {
    # ［インデックス管理］の保存先の行（ワークスペースのフォルダ。共有のときは警告も）を、今のワークスペースに合わせる。
    # 文言は判断層（getIndexLocationView）。ネットワークかどうかは共有に接続せずに見分ける（testNetworkPath）
    $dir = [string]$workspace.Dir
    $view = getIndexLocationView $dir (testNetworkPath $dir)
    $ui.IndexLocationLabel.Text = $view.Heading
    $ui.IndexLocationText.Text = $view.Path
    # パスと注意書きは、情報の印とパスのツールチップで読める（一覧の邪魔をしない）
    $tip = "$($view.Path)`n$($view.Note)"
    $ui.IndexLocationText.ToolTip = $tip
    $ui.IndexLocationInfo.ToolTip = $tip
    $ui.IndexLocationWarning.Text = $view.Warning
    $ui.IndexLocationWarning.Visibility = if ($view.Warning) { "Visible" } else { "Collapsed" }
}

function applyIndexLocationState {
    # openIndexLocation の続き（getPathState の結果を画面に反映する）。フォルダが見つかれば、エクスプローラーで開く。
    # 開けないときの文言は showMessage に渡す（タブはステータス、確認のダイアログはダイアログの中）
    param (
        $state,
        [string]$path,
        [scriptblock]$showMessage
    )

    $kind = getIndexSourceFolderKind $state
    if ($kind -eq "Open") {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"${path}`""
        & $showMessage ""
        return
    }
    & $showMessage (getIndexLocationOpenStatus $kind $path $state.Message)
}

function openIndexLocation {
    # 保存先（ワークスペースのフォルダ）をエクスプローラーで開く。タブと確認のダイアログの［フォルダを開く］の共通の入口。
    # 無い・つながらない・開けないときの文言は showMessage（文字を受け取る処理）に渡す。
    # ネットワークにあれば裏のスレッドで確かめ、画面のスレッドは待たない
    param (
        [scriptblock]$showMessage
    )

    $path = [string]$workspace.Dir
    $requestBox = $script:indexLocationRequest
    $requestBox.Value++
    $requestId = $requestBox.Value
    $applyState = ${function:applyIndexLocationState}
    $apply = {
        param ($state)
        if ($requestId -ne $requestBox.Value) {
            # 待っている間にもう一度押した。前の依頼は捨てる
            return
        }
        & $applyState $state $path $showMessage
    }.GetNewClosure()

    if (!(testNetworkPath $path)) {
        & $apply (getPathState $path)
        return
    }
    & $showMessage (getIndexLocationOpenStatus "Checking" $path)
    $otherState = ${pathStateOther}
    startJob {
        param ($path)
        getPathState $path
    } @($path) {
        param ($output, $errorText)
        if ($errorText) {
            & $apply @{ State = $otherState; IsDirectory = $false; Message = $errorText }
        } else {
            & $apply $output[0]
        }
    }.GetNewClosure() (getWorkspaceJobQueue $path)
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
    # 更新中の行のバッジ（「更新中」と、右の補足「45%」）と棒を、進み具合に合わせて置き直す（進み具合の更新のたびに呼ぶ）
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

    $blocker = getCurrentIndexJobBlocker
    if ($blocker -ne "") {
        showMessage (getIndexJobBlockedMessage $blocker $operation) "OK" "Warning" | Out-Null
        return $false
    }
    return $true
}
