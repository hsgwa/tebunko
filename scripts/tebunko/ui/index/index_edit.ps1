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
    $item = getIndexTargetItem
    if ($null -eq $item -or !(testIndexOperable "編集")) {
        return
    }
    $result = showIndexEditDialog $item
    if ($null -eq $result) {
        return
    }

    applyIndexEdit $item $result
}

function applyIndexEdit {
    # 編集の結果（@{ Path; Name }）を行に反映して保存する。［編集…］のダイアログと、詳細のフォルダパスの［...］が同じ道で使う。
    # 変わったところが無ければ何もしない。
    # 名前を変えるときは、インデックスのフォルダ（work\index\<名前>）と取り込み一覧の記録の改名を別スレッドで行い、
    # 終わってから行に反映する（ワークスペースが共有フォルダにあると改名に時間がかかる・届かないことがあるため）
    param (
        $item,
        $result
    )

    # 大文字・小文字だけの変更も改名する（-ne は大文字・小文字を区別しないため -cne で比べる）
    $renamed = $result.Name -cne $item.Name
    if (!$renamed -and $result.Path -eq $item.Path) {
        return
    }
    if (!$renamed) {
        completeIndexEdit $item $result
        return
    }

    # 改名と同じ仕事の中で、元のフォルダの記録（source_folder.txt）も新しい名前・場所で書き直す
    $folders = @($script:targetItems | Where-Object { $_.Name } | ForEach-Object {
        if ($_ -eq $item) {
            [pscustomobject]@{ Name = $result.Name; Path = $result.Path }
        } else {
            [pscustomobject]@{ Name = $_.Name; Path = $_.Path }
        }
    })
    $complete = ${function:completeIndexEdit}   # 終わったときの処理は、関数を変数に取って呼ぶ（クロージャからは関数の名前を引けないため）
    startIndexStoreJob "rename" $item.Name $result.Name $folders {
        & $complete $item $result
    }.GetNewClosure()
}

function completeIndexEdit {
    # 編集の結果を行に反映して保存する（名前の変更は、インデックスのフォルダの改名が終わってから呼ぶ）
    param (
        $item,
        $result
    )

    $changes = New-Object System.Collections.Generic.List[string]
    $renamed = $result.Name -cne $item.Name
    if ($renamed) {
        $changes.Add("名前 [$($item.Name)] → [$($result.Name)]")
        $item.SetName($result.Name)
    }
    if ($result.Path -ne $item.Path) {
        $changes.Add("場所 $($item.Path) → $($result.Path)")
        $item.SetPath($result.Path)
        $item.StatusChecked = $false
        updateFolderItemStatus $item
    }

    saveTargets
    if (!$renamed) {
        # 名前を変えたときは、改名の仕事の中で書き直し済み
        updateIndexSourceFile
    }
    refreshIndexViews
    setStatus ("インデックスを変更しました（" + ($changes -join " / ") + "）")
}

function changeIndexFolder {
    # 詳細のフォルダパスの［...］。フォルダを選ぶ画面を出し、選んだ場所を［編集…］で変えて［OK］したときと同じ検査・同じ反映で変える。
    # 取り消したとき・同じフォルダを選んだときは何も変えない
    $item = getIndexTargetItem
    if ($null -eq $item -or !(testIndexOperable "編集")) {
        return
    }
    $initial = normalizeFolderPath $item.Path
    # 開始フォルダの有無は、直前の refreshFolderStatus の結果がそのまま使えるときだけ調べずに使う（［編集…］と同じ）
    $knownExisting = $item.StatusChecked -and $item.FolderExists
    $path = selectFolder "インデックスにする、Office ファイルのあるフォルダを選んでください" $initial $window $knownExisting
    if (!$path) {
        return
    }
    $path = normalizeFolderPath $path
    if ($path -eq $item.Path) {
        return
    }
    $message = testIndexEditInput $path $item.Name $script:targetItems $item
    if ($message -ne "") {
        showMessage $message "OK" "Warning" | Out-Null
        return
    }
    applyIndexEdit $item @{ Path = $path; Name = $item.Name }
}

function startIndexStoreJob {
    # インデックス（work\index\<名前>）と取り込み一覧の記録の削除・名前の変更を別スレッドで行う。
    # 数万フォルダの削除は数十秒かかることがあり、共有フォルダでは改名にも時間がかかる・届かないことがある。
    # 画面のスレッドで行うと「応答なし」になるため。
    # 終わるまでインデックスの操作・インデックス作成の開始・ワークスペースの変更はできないようにし、何をしているかをステータスに出す
    param (
        [string]$kind,        # "delete"（削除）か "rename"（名前の変更。元のフォルダの記録の書き直しまで行う）
        [string]$name,
        [string]$newName,     # 名前の変更のときの新しい名前
        [object[]]$folders,   # 名前の変更のときに書き直す元のフォルダの記録（@{ Name; Path }）
        [scriptblock]$onDone  # 終わった後に画面のスレッドで行うこと（$script:indexJobName で名前を参照できる）
    )

    $script:indexBusy = $true
    $script:indexJobName = $name
    $script:indexJobOnDone = $onDone
    $script:indexJobKind = $kind
    updateIndexingButton
    setStatus (getIndexStoreJobStatus $kind $name $newName)
    # 裏のスレッドは lib.ps1 を読み込んだときのワークスペースを覚えているため、場所は渡す。
    # 届かない共有フォルダで他の仕事を巻き込んで待たせないよう、ワークスペースの場所で裏の列を選ぶ
    startJob {
        param ($kind, $name, $newName, $folders, $dir, $statusPath, $settingsPath)
        # 届かない場所では、何もせずに成功にしない（設定だけが変わってしまう）。失敗として返す
        assertWorkspaceReachable $dir
        if ($kind -eq "rename") {
            renameIndex $name $newName $dir $statusPath $settingsPath
            # 元のフォルダの記録の書き直しに失敗しても、改名は済んでいる。改名の失敗にはせず、書けなかった理由を別に返す
            $sourceFolderError = ""
            try {
                writeSourceFolderFile $folders $dir
            } catch {
                $sourceFolderError = $_.Exception.Message
            }
            @{ SourceFolderError = $sourceFolderError }
        } else {
            removeIndex $name $dir $statusPath $settingsPath
        }
    } (@($kind, $name, $newName) + @(,@($folders)) + @($workspace.IndexDir, $workspace.StatusFile, ${settingsFile})) {
        param ($output, $errorText)
        $script:indexBusy = $false
        updateIndexingButton
        $name = $script:indexJobName
        if ($errorText) {
            try {
                if ($script:indexJobKind -eq "rename") {
                    # 行にはまだ反映していない。一覧を保存済みの内容に戻す
                    loadTargets
                }
                # 読み直しの結果（届かないときの「接続できません」）で、失敗の知らせが消えないようにする
                refreshIndexViewsKeepingStatus (getIndexStoreJobFailedStatus $script:indexJobKind $name $errorText)
            } finally {
                finishIndexJobClose
            }
            return
        }
        try {
            if ($script:indexJobOnDone) {
                & $script:indexJobOnDone
            }
            if ($script:indexJobKind -ne "rename") {
                # 削除: 一覧からの除去と保存は onDone で済ませてから、画面を読み直す
                refreshIndexViews
            }
            if ($script:indexJobKind -eq "rename") {
                # 改名は済んでいるので、行への反映（上の onDone）は行い、元のフォルダの記録だけ書けなかったことを知らせる
                $sourceFolderError = getIndexStoreSourceFolderError $output
                if ($sourceFolderError -ne "") {
                    setStatus (getSourceFolderFileFailedStatus $sourceFolderError)
                }
            }
        } finally {
            # 後始末が例外になっても、待っていた閉じる操作は必ず進める
            finishIndexJobClose
        }
    } (getWorkspaceJobQueue $workspace.IndexDir)
}

function finishIndexJobClose {
    # 改名・削除の途中で閉じようとして待っていたなら、終わったので閉じる（設定への反映は済んでいる）。
    # 同じ周期で動く残りの仕事の結果が、閉じた窓に届かないよう、閉じるのは画面のスレッドの次の順番にする
    $script:closeAskedDuringIndexJob = $false
    if ($script:closeConfirming) {
        # 「待たずに閉じるか」の確認を出している間（Closing の途中）。ここで Close を呼ぶと窓が壊れるので、確認から戻ったときに閉じる
        return
    }
    if ($script:closeAfterIndexJob) {
        $script:closeAfterIndexJob = $false
        $window.Dispatcher.BeginInvoke([action]{ $window.Close() }) | Out-Null
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
    $confirm = getIndexDeleteConfirmMessage $names
    $answer = showConfirm `
        -title "インデックスの削除" `
        -heading $confirm.Heading `
        -hint $confirm.Body `
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
        # 届かない場所では、何もせずに成功にしない（設定だけが変わってしまう）。失敗として返す
        assertWorkspaceReachable $dir
        $results = removeIndexes -names $names -dir $dir -statusPath $statusPath -settingsPath $settingsPath
        , $results
    } @(,$names + @($workspace.IndexDir, $workspace.StatusFile, ${settingsFile})) {
        param ($output, $errorText)
        $script:indexBusy = $false
        updateIndexingButton
        try {
            if ($errorText) {
                loadTargets
                refreshIndexViewsKeepingStatus "削除に失敗しました：${errorText}"
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
            if ($script:closeAfterIndexJob) {
                # 閉じるのを待っているときは、結果のダイアログ（閉じるまで進めない）を出さずに、失敗があればステータスに残して閉じる
                $failed = @($results | Where-Object { !$_.Ok })
                if ($failed.Count -gt 0) {
                    setStatus (getBulkDeleteFailedStatus $failed.Count)
                }
            } else {
                showBulkIndexResult "削除" $results
            }
        } finally {
            finishIndexJobClose
        }
    } (getWorkspaceJobQueue $workspace.IndexDir)
}

function deleteIndex {
    # ［削除］。一覧から削除し、インデックス（work\index\<名前>）と取り込み一覧の記録も削除する
    $item = getIndexTargetItem
    if ($null -eq $item -or !(testIndexOperable "削除")) {
        return
    }

    $confirm = getIndexDeleteConfirmMessage @($item.Name)
    $answer = showConfirm `
        -title "インデックスの削除" `
        -heading $confirm.Heading `
        -hint $confirm.Body `
        -choices @(@{ Text = "削除する"; Value = "delete"; Danger = $true })
    if ($answer -ne "delete") {
        return
    }

    # インデックスの削除（時間がかかることがある）は別スレッドで行う。一覧からの除去と設定の保存は、削除できてから行う
    # （ワークスペースに届かないときは、一覧も設定も変えずに失敗を知らせる）
    # （終わるまでの間に一覧が読み直されると、行は別のオブジェクトになる。名前で探して外す）
    startIndexStoreJob "delete" $item.Name "" @() {
        $removed = @($script:targetItems | Where-Object { $_.Name -eq $script:indexJobName })[0]
        if ($removed) { $script:targetItems.Remove($removed) | Out-Null }
        saveTargets
        updateIndexSourceFile
        updateIndexListView
        setStatus "インデックス [$($script:indexJobName)] を削除しました"
    }
}
