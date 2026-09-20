# ［1 インデックス管理］タブ（インデックスの一覧・新規作成・編集・削除）。

function getTargetsKey {
    # 変換対象フォルダの一覧（@{ Path; Enabled } の配列）を比べるための文字列
    param (
        [object[]]$folders
    )

    return (@($folders | Where-Object { $_ } | ForEach-Object { "$($_.Enabled)`t$($_.Path)" }) -join "`n")
}

# ============================================================================
# ［1 インデックス管理］（インデックスの作成・編集・削除と、変換の実行）
# ============================================================================

$script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexGrid.ItemsSource = $script:targetItems
$script:loadingTargets = $false
# ［変換］チェックのクリックで保存する（TwoWay バインドで Enabled は更新済み。PS class のプレーンな
# プロパティは PropertyChanged を出さないため、購読ではなくここで保存する）
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler] {
    param ($s, $e)
    safe {
        $cb = $e.OriginalSource
        if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.DataContext -is [FolderItem] -and !$script:loadingTargets) {
            saveTargets
            updateConvertButton
        }
    }
})
$script:savedTargets = $null  # 最後に読み込み・保存したインデックス一覧（getTargetsKey）。ほかでの変更の検出に使う
$script:editDialog = $null    # 新規作成・編集のダイアログ（開いている間だけ）
$script:convertProcess = $null
$script:convertStart = $null
$script:convertFailed = 0  # 変換中に一覧へ反映済みの失敗件数
$script:conversionState = $null
$script:indexSummary = $null

function isConverting {
    return ($null -ne $script:convertProcess) -and !$script:convertProcess.HasExited
}

function updateFolderItemStatus {
    param (
        $item
    )

    if (Test-Path -LiteralPath $item.Path -PathType Container) {
        $item.SetStatus("✓ フォルダがあります", ${okBrush})
    } else {
        $item.SetStatus("✗ フォルダが見つかりません", ${ngBrush})
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
    $item.LastConvertedText = ""
    updateFolderItemStatus $item
    # ［変換］チェックの保存は、一覧のチェックボックスの Click（IndexGrid.AddHandler）で行う。
    # PS class のプレーンなプロパティは TwoWay セットで PropertyChanged を出さないため、購読では拾えない。
    return $item
}

function loadTargets {
    $script:loadingTargets = $true
    try {
        $script:targetItems.Clear()
        $folders = @(getTargetFolders)
        # 名前の決まっていないインデックス（以前の版の設定から移した直後など）には、ここで名前を割り当てて確定する。
        # 一覧・編集・削除はインデックス名で扱うため、画面に出す時点で名前があるようにする（変換側と同じ assignIndexNames を使う）
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
}

function saveTargets {
    writeTargetFolders @($script:targetItems | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Enabled = $_.Enabled } })
    $script:savedTargets = getTargetsKey @(getTargetFolders)
    setStatus "インデックス一覧を保存しました（$(Get-Date -Format 'H:mm')）"
}

function updateIndexSourceFile {
    # インデックスのフォルダの 元のフォルダ.txt を今の一覧に合わせて書き直す。
    # 次の変換を待たずに、検索結果から元のファイルを開けるようにする（インデックスが無ければ何もしない）
    if (!(Test-Path -LiteralPath ${indexDir} -PathType Container)) {
        return
    }
    writeSourceFolderFile @($script:targetItems | Where-Object { $_.Name } | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path } })
}

function refreshIndexViews {
    # インデックスを作成・編集・削除した後、検索タブ（検索対象のツリー・件数）も読み直す
    $script:sourceFolderMaps = @{}
    $script:indexSummary = $null
    loadIndexTree
    refreshIndexSummary
    refreshConversionState
}

function applyIndexStats {
    # 変換一覧の集計（getIndexStats）を一覧の各行のファイル数・最終変換に反映する
    param (
        $stats
    )

    foreach ($item in $script:targetItems) {
        $stat = $null
        if ($item.Name -and $null -ne $stats -and $stats.ContainsKey($item.Name)) {
            $stat = $stats[$item.Name]
        }
        if ($null -eq $stat) {
            $item.SetStats("－", "まだ変換していません", "")
            continue
        }
        $converted = [datetime]::MinValue
        $lastText = if ($stat.LastConverted -and [datetime]::TryParseExact($stat.LastConverted, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$converted)) {
            formatTime $converted
        } else {
            ""
        }
        $item.SetStats(("{0:#,0}" -f $stat.Total), ("済 {0:#,0} 件 ・ 未変換 {1:#,0} 件 ・ 失敗 {2:#,0} 件" -f $stat.Done, $stat.Pending, $stat.Failed), $lastText)
    }
}

function updateIndexListView {
    $ui.IndexGridPlaceholder.Visibility = if ($script:targetItems.Count -eq 0) { "Visible" } else { "Collapsed" }
    updateConvertButton
}

function testIndexOperable {
    # 変換中はインデックスの作成・編集・削除をしない（インデックスのフォルダ・変換一覧を変換側が使っているため）
    param (
        [string]$operation
    )

    if (isConverting) {
        showMessage "変換中はインデックスを${operation}できません。変換が終わるまでお待ちください（［中止］で止められます）。" "OK" "Warning" | Out-Null
        return $false
    }
    if ($script:indexBusy) {
        # 前のインデックスの TSV を削除している最中（別スレッド）
        showMessage "前のインデックスの削除が終わるまでお待ちください。" "OK" "Warning" | Out-Null
        return $false
    }
    return $true
}

function showIndexEditDialog {
    # インデックスの新規作成・編集のダイアログ。決めた内容 @{ Path; Name } を返す（キャンセルは $null）。
    #   item: 編集するインデックス（$null なら新規作成）
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
        $dialog.Title = "インデックスの新規作成"
        $ctrl.IntroText.Text = "Office ファイル（Excel・Word・PowerPoint）の入っているフォルダを 1 つ選んでください。" +
            "ここでは一覧に加えるだけです。中のファイルを読むのは［変換を開始］を押してからです。"
    } else {
        $dialog.Title = "インデックスの編集"
        $ctrl.IntroText.Text = "名前と、元のフォルダの場所を変えられます。"
        $ctrl.FolderBox.Text = $item.Path
        $ctrl.NameBox.Text = $item.Name
        $ctrl.NoticeText.Visibility = "Visible"
        $ctrl.NoticeText.Text = "変えるのは名前と場所だけです。変換したデータはそのまま使います（作り直しません）。" +
            "フォルダを別のドライブや共有フォルダへ移したときは、ここで新しい場所を指定してください。"
    }

    $ctrl.FolderBox.Add_TextChanged({
        safe {
            # 新規作成のときは、フォルダ名からインデックス名を自動で入れる（利用者が名前を変えた後は触らない）
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
    # 新規作成・編集のダイアログの入力を調べ、直してほしい内容を返す（問題なければ空文字列）
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
    $ui.IndexGrid.SelectedItem = $item
    $ui.IndexGrid.ScrollIntoView($item)
    saveTargets
    updateIndexSourceFile
    updateIndexListView
    refreshConversionState
    if (Test-Path -LiteralPath $path -PathType Container) {
        setStatus "インデックス [${name}] を作成しました。［変換を開始］を押すと中身を変換します"
    } else {
        setStatus "インデックス [${name}] を作成しましたが、フォルダが見つかりません：${path}"
    }
}

function newIndex {
    # ［新規作成…］。フォルダとインデックス名を決めて一覧に加える（変換はしない）
    if (!(testIndexOperable "作成")) {
        return
    }
    $result = showIndexEditDialog $null
    if ($null -eq $result) {
        return
    }
    addIndexItem $result.Path $result.Name
}

function addIndexForFolder {
    # 一覧へのドラッグ＆ドロップでインデックスを作る（名前はフォルダ名から自動で決める）
    param (
        [string]$path
    )

    if (!(testIndexOperable "作成")) {
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
    if ($result.Name -ne $item.Name) {
        # インデックスのフォルダ（work\index\<名前>）と変換一覧の記録も名前を変える（中身は作り直さない）
        renameIndex $item.Name $result.Name
        $changes.Add("名前 [$($item.Name)] → [$($result.Name)]")
        $item.SetName($result.Name)
    }
    if ($result.Path -ne $item.Path) {
        $changes.Add("場所 $($item.Path) → $($result.Path)")
        $item.SetPath($result.Path)
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
    # インデックス（work\index\<名前>）と変換一覧の記録の削除を別スレッドで行う。
    # 数万フォルダの削除は数十秒かかることがあり、画面のスレッドで行うと「応答なし」になるため。
    # 終わるまでインデックスの操作・変換の開始はできないようにし、何をしているかをステータスに出す
    param (
        [string]$name,
        [string]$operation,   # "削除"（表示に使う）
        [scriptblock]$onDone  # 削除が終わった後に画面のスレッドで行うこと（$script:indexJobName で名前を参照できる）
    )

    $script:indexBusy = $true
    $script:indexJobName = $name
    $script:indexJobOnDone = $onDone
    $script:indexJobOperation = $operation
    updateConvertButton
    setStatus "インデックス [${name}] の TSV を削除しています…（件数によっては少し時間がかかります）"
    startJob {
        param ($libPath, $name)
        . $libPath
        removeIndex $name
    } @(${libPath}, $name) {
        param ($output, $errorText)
        $script:indexBusy = $false
        updateConvertButton
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
    # ［削除］。一覧から削除し、変換した TSV（work\index\<名前>）と変換一覧の記録も削除する
    $item = $ui.IndexGrid.SelectedItem
    if ($null -eq $item -or !(testIndexOperable "削除")) {
        return
    }

    $answer = showConfirm `
        -heading "インデックス「$($item.Name)」を一覧から削除しますか？" `
        -facts @(
            (factGone "win_grep が作った検索用のデータが消えます" "このフォルダは検索できなくなります（もう一度［変換を開始］すれば作り直せます）"),
            (factKept "元のフォルダと、その中のファイルはそのままです" $item.Path)
        ) `
        -hint "しばらく検索しないだけなら、削除せずに［変換］のチェックを外してください。検索用のデータは残ったままです。" `
        -choices @(@{ Text = "削除する"; Value = "delete"; Danger = $true })
    if ($answer -ne "delete") {
        return
    }

    # 一覧からはすぐ消し、TSV の削除（時間がかかることがある）は別スレッドで行う
    $script:targetItems.Remove($item)
    saveTargets
    updateIndexSourceFile
    updateIndexListView
    startIndexRemoveJob $item.Name "削除" {
        setStatus "インデックス [$($script:indexJobName)] を削除しました"
    }
}

function updateConvertButton {
    $ready = $false
    foreach ($item in $script:targetItems) {
        if ($item.Enabled -and (Test-Path -LiteralPath $item.Path -PathType Container)) {
            $ready = $true
            break
        }
    }

    $state = $script:conversionState
    if ($script:indexBusy) {
        # インデックスの削除中（別スレッド）は、変換もインデックスの操作も始めない
        $ready = $false
    }
    if (isConverting) {
        $ui.ConvertButton.Content = "変換中…"
        $ui.ConvertButton.IsEnabled = $false
    } else {
        $ui.ConvertButton.Content = if ($state -and $state.Pending -gt 0) { "続きから再開（残り $($state.Pending) 件）" } else { "変換を開始" }
        $ui.ConvertButton.IsEnabled = $ready
    }

    # ボタンの下の一言。押す前は「押すと何が起きるか」、変換中は「やめるとどうなるか」を書く
    $hint = if (isConverting) {
        "変換中も検索できます。やめるときは［中止］を押してください（次に［変換を開始］を押すと続きから再開します）。"
    } else {
        "押すと、何件変換するかを確認してから、新しいファイル・変わったファイルだけを変換します。"
    }
    if (!$ready -and !(isConverting)) {
        $hint = "まずインデックスを作って、［変換］にチェックを付けてください。"
    } elseif ($state -and $state.Failed -gt 0 -and !(isConverting)) {
        $hint = "前回うまく変換できなかったファイルがあります（押したあとで、もう一度ためすか選べます）。" + $hint
    }
    $ui.ConvertHint.Text = $hint

    # インデックスの作成・編集・削除は、選んでいるかどうかと変換中かどうかで切り替える
    # （変換中はインデックスのフォルダ・変換一覧を変換側が使っているため触らない）
    $selected = $null -ne $ui.IndexGrid.SelectedItem
    $editable = !(isConverting) -and !$script:indexBusy
    $ui.NewIndexButton.IsEnabled = $editable
    $ui.EditIndexButton.IsEnabled = $selected -and $editable
    $ui.RemoveIndexButton.IsEnabled = $selected -and $editable
}

function refreshConversionState {
    # 変換一覧の集計は、ファイルが数万行になると数秒〜十数秒かかる。
    # 画面のスレッドで行うと、起動時・タブの切り替え時に画面が固まる（応答なしになる）ため別スレッドで数える
    if ($script:stateRunning) {
        $script:stateAgain = $true
        return
    }
    $script:stateRunning = $true
    $script:stateAgain = $false
    startJob {
        param ($libPath)
        . $libPath
        getConversionState
    } @(${libPath}) {
        param ($output, $errorText)
        $script:stateRunning = $false
        # 変換側が書き込んでいる瞬間などは、次の機会に読み直す
        if ($output -and $output.Count -gt 0 -and $output[0]) {
            applyConversionState $output[0]
        }
        if ($script:stateAgain) {
            refreshConversionState
        }
    }
}

function applyConversionState {
    # 集計（別スレッド）の結果を画面に反映する
    param (
        $state  # getConversionState の結果
    )

    $script:conversionState = $state

    # 失敗したファイルは下の一覧に原因とともに表示する
    $ui.ConversionStateText.Text = if ($state.Pending -gt 0 -and !(isConverting)) { "⏸ 前回の変換が中断しています（残り $($state.Pending) 件）" } else { "" }
    $ui.IndexTabHeader.Text = if ($state.Failed -gt 0) { "⚠ 1 インデックス管理" } else { "1 インデックス管理" }
    applyIndexStats $state.IndexStats
    updateFailedList $state
    updateIndexSummaryText
    updateConvertButton
}

function updateFailedList {
    # 変換に失敗したファイルと原因（変換一覧のエラー列）を一覧に表示する
    param (
        $state  # getConversionState の結果
    )

    $folderPaths = @{}  # インデックス名 → 変換対象フォルダ（大文字・小文字を区別しない）
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
        $converted = [datetime]::MinValue
        if ([datetime]::TryParseExact([string]$status.変換日時, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$converted)) {
            $row.ConvertedText = formatTime $converted
        }
        $parts = splitIndexRelPath $status.相対パス
        if ($folderPaths.ContainsKey($parts.Name)) {
            $row.SourcePath = Join-Path $folderPaths[$parts.Name] $parts.Rest
        }
        [void]$rows.Add($row)
    }

    $ui.FailedGrid.ItemsSource = $rows
    $ui.FailedHeading.Text = "⚠ 変換に失敗したファイル $($rows.Count) 件"
    $ui.FailedPanel.Visibility = if ($rows.Count -gt 0) { "Visible" } else { "Collapsed" }
}

function openFailedFileFolder {
    # 失敗したファイルの場所をエクスプローラーで開く（ファイルを選択した状態）
    $row = $ui.FailedGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    if (!$row.SourcePath) {
        setStatus "元のファイルの場所が分かりません（変換一覧に変換対象フォルダの記録がありません）：$($row.RelPath)"
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
    $text = "TSV $($summary['Count'].ToString('N0')) 件 ・ 最終変換 $(formatTime $summary['LastWrite'])"
    $state = $script:conversionState
    if ($state -and $state.Done -gt 0) {
        $text = "変換済み $($state.Done.ToString('N0')) ファイル（$text）"
    }
    $ui.IndexSummaryText.Text = $text
}

function refreshIndexSummary {
    # TSV の件数は数えるのに時間がかかることがあるため、別スレッドで数える
    if ($script:summaryRunning) {
        $script:summaryAgain = $true
        return
    }
    $script:summaryRunning = $true
    $script:summaryAgain = $false
    $folders = @(${indexDir})
    startJob {
        param ($libPath, $folders)
        . $libPath
        getIndexSummary $folders
    } @(${libPath}, $folders) {
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
