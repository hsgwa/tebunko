# インデックス管理の画面の、インデックスの追加・編集・削除。

function showIndexEditDialog {
    # インデックスの追加・編集のダイアログ。決めた内容 @{ Path; Name } を返す（キャンセルは $null）。
    #   item: 編集するインデックス（$null なら追加）
    param (
        $item = $null
    )

    $dialog = loadWindow "${xamlDir}\dialog_index_edit.xaml" ${fontsDir}
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("OkButton", "DialogHeadingText", "BrowseButton", "FolderBox", "NameBox", "IntroText", "NoticeText", "ErrorText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $script:editDialog = @{ Window = $dialog; Ctrl = $ctrl; Item = $item; Suggested = "" }

    if ($null -eq $item) {
        $dialog.Title = "インデックスの追加"
        $ctrl.DialogHeadingText.Text = "インデックスを追加する"
        $ctrl.IntroText.Text = "Office ファイル（Excel・Word・PowerPoint）の入っているフォルダを 1 つ選んでください。" +
            "ここでは一覧に加えるだけです。中のファイルを読むのは［すべて更新］を押してからです。"
    } else {
        $dialog.Title = "インデックスの編集"
        $ctrl.DialogHeadingText.Text = "インデックスを編集する"
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
    if (showOwnedDialog $dialog) {
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
    # ［＋ フォルダを追加］。フォルダとインデックス名を決めて一覧に加える（インデックス作成はしない）
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

function deleteIndexes {
    # ［削除…］で 2 件以上チェックしているとき。確認は 1 回だけ。1 つ失敗しても残りを続け、終わりに失敗した名前と理由を出す
    param (
        [string[]]$names
    )

    if (!(testIndexOperable "削除")) {
        return
    }
    $confirm = getIndexBulkDeleteConfirm $names
    $answer = showConfirm `
        -title "インデックスの削除" `
        -heading $confirm.Heading `
        -hint $confirm.Hint `
        -detail $confirm.Detail `
        -choices @(@{ Text = "削除する"; Value = "delete"; Danger = $true })
    if ($answer -ne "delete") {
        return
    }

    $script:indexBusy = $true
    $script:indexJobName = ""
    $script:indexJobOnDone = $null
    updateIndexingButton
    setStatus "$($names.Count) 件のインデックスを削除しています…（件数によっては少し時間がかかります）"
    startJob {
        param ($names, $dir, $statusPath, $settingsPath)
        $results = removeIndexes -names $names -dir $dir -statusPath $statusPath -settingsPath $settingsPath
        , $results
    } @(,$names + @($workspace.IndexDir, $workspace.StatusFile, ${settingsFile})) {
        param ($output, $errorText)
        $script:indexBusy = $false
        updateIndexingButton
        if ($errorText) {
            setStatus "削除に失敗しました：${errorText}"
            loadTargets
            refreshIndexViews
            return
        }
        # 消せたものを一覧から外す（失敗したものは残す）
        $results = @($output | ForEach-Object { $_ })
        foreach ($result in $results | Where-Object { $_.Ok }) {
            $item = @($script:targetItems | Where-Object { $_.Name -eq $result.Name })[0]
            if ($item) { $script:targetItems.Remove($item) | Out-Null }
        }
        saveTargets
        updateIndexSourceFile
        updateIndexListView
        refreshIndexViews
        showBulkIndexResult "削除" $results
    }
}

function deleteIndex {
    # ［削除］。一覧から削除し、インデックス（work\index\<名前>）と取り込み一覧の記録も削除する
    $item = $ui.IndexGrid.SelectedItem
    if ($null -eq $item -or !(testIndexOperable "削除")) {
        return
    }

    $answer = showConfirm `
        -title "インデックスの削除" `
        -heading "「$($item.Name)」のインデックスを削除しますか？" `
        -hint "元のファイルは削除されません。" `
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
