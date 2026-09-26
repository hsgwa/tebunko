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
# ［作成］チェックのクリックで保存する（TwoWay バインドで Enabled は更新済み。PS class のプレーンな
# プロパティは PropertyChanged を出さないため、購読ではなくここで保存する）
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler] {
    param ($s, $e)
    safe {
        $cb = $e.OriginalSource
        if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.DataContext -is [FolderItem] -and !$script:loadingTargets) {
            saveTargets
            updateIndexingButton
        }
    }
})
$script:savedTargets = $null  # 最後に読み込み・保存したインデックス一覧（getTargetsKey）。ほかでの変更の検出に使う
$script:editDialog = $null    # 追加・編集のダイアログ（開いている間だけ）
$script:indexingSession = $null  # 実行中のインデックス作成（IndexingSession。終わって片づけたら $null）
$script:indexingStart = $null
$script:ingestFailed = 0  # インデックス作成中に一覧へ反映済みの失敗件数
$script:indexingState = $null
$script:indexSummary = $null

function isIndexing {
    return ($null -ne $script:indexingSession) -and $script:indexingSession.IsRunning()
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
    }
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
    # ［作成］チェックの保存は、一覧のチェックボックスの Click（IndexGrid.AddHandler）で行う。
    # PS class のプレーンなプロパティは TwoWay セットで PropertyChanged を出さないため、購読では拾えない。
    return $item
}

function loadTargets {
    $script:loadingTargets = $true
    try {
        $script:targetItems.Clear()
        $folders = @(getTargetFolders)
        # 名前の決まっていないインデックス（以前の版の設定から移した直後など）には、ここで名前を割り当てて確定する。
        # 一覧・編集・削除はインデックス名で扱うため、画面に出す時点で名前があるようにする（インデクサと同じ assignIndexNames を使う）
        if (@($folders | Where-Object { $_ -and !$_.Name }).Count -gt 0) {
            $folders = @(assignIndexNames $folders (readStatusFile).Folders)
            writeTargetFolders $folders
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
}

function saveTargets {
    writeTargetFolders @($script:targetItems | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Enabled = $_.Enabled } })
    $script:savedTargets = getTargetsKey @(getTargetFolders)
    setStatus "インデックス一覧を保存しました（$(Get-Date -Format 'H:mm')）"
}

function updateIndexSourceFile {
    # インデックスのフォルダの 元のフォルダ.txt を今の一覧に合わせて書き直す。
    # 次のインデックス作成を待たずに、検索結果から元のファイルを開けるようにする（インデックスが無ければ何もしない）
    if (!(Test-Path -LiteralPath $workspace.IndexDir -PathType Container)) {
        return
    }
    writeSourceFolderFile @($script:targetItems | Where-Object { $_.Name } | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path } })
}

function refreshIndexViews {
    # インデックスを追加・編集・削除した後、検索タブ（検索対象のツリー・件数）も読み直す
    $script:sourceFolderMaps = @{}
    $script:indexSummary = $null
    loadIndexTree
    refreshIndexSummary
    refreshIndexingState
}

function applyIndexStats {
    # 取り込み一覧の集計（getIndexStats）を一覧の各行のファイル数・最終取り込みに反映する
    param (
        $stats
    )

    foreach ($item in $script:targetItems) {
        $stat = $null
        if ($item.Name -and $null -ne $stats -and $stats.ContainsKey($item.Name)) {
            $stat = $stats[$item.Name]
        }
        if ($null -eq $stat) {
            $item.SetStats("－", "まだ取り込んでいません", "")
            continue
        }
        $ingested = [datetime]::MinValue
        $lastText = if ($stat.LastIngested -and [datetime]::TryParseExact($stat.LastIngested, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ingested)) {
            formatTime $ingested
        } else {
            ""
        }
        $item.SetStats(("{0:#,0}" -f $stat.Total), ("済 {0:#,0} 件 ・ 未取り込み {1:#,0} 件 ・ 失敗 {2:#,0} 件" -f $stat.Done, $stat.Pending, $stat.Failed), $lastText)
    }
}

function updateIndexListView {
    $ui.IndexGridPlaceholder.Visibility = if ($script:targetItems.Count -eq 0) { "Visible" } else { "Collapsed" }
    updateIndexingButton
}

function testIndexOperable {
    # インデックス作成中はインデックスの追加・編集・削除をしない（インデックスのフォルダ・取り込み一覧をインデクサが使っているため）
    param (
        [string]$operation
    )

    if (isIndexing) {
        showMessage "インデックス作成中はインデックスを${operation}できません。インデックス作成が終わるまでお待ちください（［中止］で止められます）。" "OK" "Warning" | Out-Null
        return $false
    }
    if ($script:indexBusy) {
        # 前のインデックス（集約ファイル）を削除している最中（別スレッド）
        showMessage "前のインデックスの削除が終わるまでお待ちください。" "OK" "Warning" | Out-Null
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
            $path = selectFolder "インデックスにする、Office ファイルのあるフォルダを選んでください" $initial $d.Window
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
    if (Test-Path -LiteralPath $path -PathType Container) {
        setStatus "インデックス [${name}] を追加しました。［インデックス作成を開始］を押すと中身を取り込みます"
    } else {
        setStatus "インデックス [${name}] を追加しましたが、フォルダが見つかりません：${path}"
    }
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
        param ($name, $dir, $statusPath)
        removeIndex $name $dir $statusPath
    } @($name, $workspace.IndexDir, $workspace.StatusFile) {
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
    if ($script:indexBusy) {
        # インデックスの削除中（別スレッド）は、インデックス作成もインデックスの操作も始めない
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

    # インデックスの追加・編集・削除は、選んでいるかどうかとインデックス作成中かどうかで切り替える
    # （インデックス作成中はインデックスのフォルダ・取り込み一覧をインデクサが使っているため触らない）
    $selected = $null -ne $ui.IndexGrid.SelectedItem
    $editable = !(isIndexing) -and !$script:indexBusy
    $ui.NewIndexButton.IsEnabled = $editable
    $ui.EditIndexButton.IsEnabled = $selected -and $editable
    $ui.RemoveIndexButton.IsEnabled = $selected -and $editable
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
    startJob {
        param ($path)
        getIndexingState -path $path
    } @($workspace.StatusFile) {
        param ($output, $errorText)
        $script:stateRunning = $false
        # インデクサが書き込んでいる瞬間などは、次の機会に読み直す
        if ($output -and $output.Count -gt 0 -and $output[0]) {
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
    $ui.IndexingStateText.Text = if ($state.Pending -gt 0 -and !(isIndexing)) { "⏸ 前回のインデックス作成が中断しています（残り $($state.Pending) 件）" } else { "" }
    $ui.IndexTabHeader.Text = if ($state.Failed -gt 0) { "⚠ 1 インデックス管理" } else { "1 インデックス管理" }
    applyIndexStats $state.IndexStats
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

function openFailedFileFolder {
    # 失敗したファイルの場所をエクスプローラーで開く（ファイルを選択した状態）
    $row = $ui.FailedGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    if (!$row.SourcePath) {
        setStatus "元のファイルの場所が分かりません（取り込み一覧にクロール対象フォルダの記録がありません）：$($row.RelPath)"
        return
    }
    if (Test-Path -LiteralPath $row.SourcePath -PathType Leaf) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$($row.SourcePath)`""
        return
    }
    $dir = Split-Path $row.SourcePath -Parent
    if (Test-Path -LiteralPath $dir -PathType Container) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"${dir}`""
        setStatus "ファイルが見つからないため、フォルダを開きました（移動・削除された可能性があります）：$($row.SourcePath)"
        return
    }
    setStatus "ファイルが見つかりません（移動・削除された可能性があります）：$($row.SourcePath)"
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
    $folders = @($workspace.IndexDir)
    startJob {
        param ($folders)
        getIndexSummary $folders
    } @(, $folders) {
        param ($output, $errorText)
        $script:summaryRunning = $false
        if ($output -and $output.Count -gt 0) {
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
