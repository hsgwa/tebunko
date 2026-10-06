# インデックス管理の画面の、インデックスのエクスポート・インポート。

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

    $dialog = loadWindow "${xamlDir}\dialog_index_edit.xaml" ${fontsDir}
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
