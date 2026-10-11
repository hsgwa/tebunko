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
        [scriptblock]$onSuccess,     # 成功したときに画面のスレッドで行うこと { param($output) }
        [string[]]$paths = @(),      # 仕事が触る、ワークスペースのほかの場所（書き出し先・取り込む zip）。ワークスペースと合わせて、裏の列を選ぶ
        [string]$busyText = ""       # 動いている間のステータス。空なら「<operation>しています…」
    )

    $script:archiveBusy = $true
    $script:archiveJobOperation = $operation
    $script:archiveJobOnSuccess = $onSuccess
    updateIndexingButton
    setStatus $(if ($busyText) { $busyText } else { "${operation}しています…" })
    startJob $scriptBlock $arguments {
        param ($output, $errorText)
        $script:archiveBusy = $false
        setExportingRows @()
        if ($errorText) {
            setStatus "$($script:archiveJobOperation)に失敗しました：${errorText}"
            updateIndexingButton
            return
        }
        if ($output -and $output.Count -gt 0 -and $script:archiveJobOnSuccess) {
            & $script:archiveJobOnSuccess $output[0]
        }
        updateIndexingButton
    } (getWorkspaceJobQueue (@($workspace.Dir) + @($paths)))
}

function showIndexExportDialog {
    # エクスポートのダイアログ。書き出し先のフォルダを返す（キャンセルは $null）
    param (
        [object[]]$items,  # 書き出すインデックス（Name・FileCountText を持つ行）
        [long]$bytes = -1  # 数え終えた zip に入る大きさ（getIndexExportSize）。負なら大きさは出さない
    )

    $dialog = loadWindow "${xamlDir}\dialog_export.xaml" ${fontsDir}
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("ExportButton", "BrowseButton", "ExportPathBox", "IntroText", "ErrorText", "ExportTitleText", "ExportSizeText", "ExportFileText", "CautionText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $header = getIndexExportHeader $items (Get-Date) $bytes
    $ctrl.ExportTitleText.Text = $header.Title
    $ctrl.ExportSizeText.Text = $header.Size
    $ctrl.ExportSizeText.Visibility = if ($header.Size) { "Visible" } else { "Collapsed" }
    $ctrl.ExportFileText.Text = $header.FileText
    $ctrl.CautionText.Text = $header.Caution
    $ctrl.IntroText.Text = getIndexExportNotice
    $ctrl.ExportPathBox.Text = $workspace.Dir

    # ボタンの Click からは script スコープの入れ物を参照する（showIndexImportDialog と同じ理由）
    $script:exportDialog = @{ Window = $dialog; Ctrl = $ctrl }

    $ctrl.BrowseButton.Add_Click({
        safe {
            $d = $script:exportDialog
            $initial = normalizeFolderPath $d.Ctrl.ExportPathBox.Text
            $path = selectFolder "エクスポート先のフォルダを選んでください" $initial $d.Window $false
            if ($path) {
                $d.Ctrl.ExportPathBox.Text = $path
            }
        }
    })
    $ctrl.ExportPathBox.Add_PreviewDragOver({ onFolderDragOver @args })
    $ctrl.ExportPathBox.Add_PreviewDrop({
        param ($sender, $e)
        safe {
            $folders = @(getDroppedFolders $e)
            if ($folders.Count -gt 0) {
                $script:exportDialog.Ctrl.ExportPathBox.Text = $folders[0]
            }
        }
        $e.Handled = $true
    })
    $ctrl.ExportButton.Add_Click({
        safe {
            $d = $script:exportDialog
            $message = testIndexExportInput $d.Ctrl.ExportPathBox.Text
            if ($message -ne "") {
                $d.Ctrl.ErrorText.Text = $message
                $d.Ctrl.ErrorText.Visibility = "Visible"
                return
            }
            $d.Window.DialogResult = $true
        }
    })

    $result = $null
    if ((showOwnedDialog $dialog)) {
        $result = normalizeFolderPath $ctrl.ExportPathBox.Text
    }
    return $result
}

function newExportIndex {
    # ［エクスポート…］。選んだインデックスを 1 つの zip に書き出す
    $item = getIndexTargetItem
    if ($null -eq $item -or !(testIndexOperable "エクスポート")) {
        return
    }
    startExportSizeCount @($item)
}

function startExportSizeCount {
    # エクスポートのダイアログを開く前に、zip に入れる大きさを別スレッドで数える（数える間は archiveBusy で他の操作を止める）。
    # 数え終えたら continueExport がダイアログを開く。対象は script スコープの入れ物に置いて渡す
    # （GetNewClosure() したスクリプトブロックから名前で関数を呼べない環境があるため、閉じ込めない）
    param (
        [object[]]$items   # 書き出すインデックス（Name・FileCountText を持つ行）
    )

    $script:exportPlan = @{ Items = @($items) }
    startIndexArchiveJob "エクスポート" {
        param ($names, $dir)
        $ws = [Workspace]::new($dir)
        getIndexExportSize $names $ws
    } @(,@($items | ForEach-Object { $_.Name }) + @($workspace.Dir)) {
        param ($size)
        continueExport $size
    } @() "エクスポートの大きさを数えています…"
}

function continueExport {
    # startExportSizeCount の続き（大きさを数え終えたところから）。ダイアログで書き出し先を決め、書き出す
    param (
        $size   # getIndexExportSize の結果
    )

    # 大きさを数えている間の文は、ここから先で要らない（ダイアログを開いている間も残さない）
    setStatus ""
    if (@($size.Missing).Count -gt 0) {
        setStatus (getIndexExportMissingText)
        return
    }
    $items = @($script:exportPlan.Items)
    $folder = showIndexExportDialog $items $size.Bytes
    if ($null -eq $folder) {
        setStatus (getIndexArchiveCancelledStatus "エクスポート")
        return
    }
    if ($items.Count -eq 1) {
        $name = $items[0].Name
        setExportingRows @($name)
        # 保存先フォルダの中身を数えるのも別スレッドで行う（届かないネットワークのフォルダで画面が止まらないように）
        startIndexArchiveJob "エクスポート" {
            param ($name, $folder, $dir, $statusPath, $settingsPath)
            $ws = [Workspace]::new($dir)
            exportIndexToFolder $name $folder $ws $settingsPath
        } @($name, $folder, $workspace.Dir, $workspace.StatusFile, ${settingsFile}) {
            param ($result)
            setStatus "インデックス [$($result.Name)] を「$($result.Path)」に書き出しました（$($result.Files) ファイル）"
        }
        return
    }
    $names = @($items | ForEach-Object { $_.Name })
    setExportingRows $names
    startIndexArchiveJob "エクスポート" {
        param ($names, $folder, $dir, $settingsPath)
        $ws = [Workspace]::new($dir)
        $results = exportIndexes -names $names -destination $folder -ws $ws -settingsPath $settingsPath
        , $results
    } @(,$names + @($folder, $workspace.Dir, ${settingsFile})) {
        param ($results)
        showBulkIndexResult "エクスポート" $results
    } @($folder)
}

function showBulkIndexResult {
    # まとめてのエクスポート・削除の終わりに、ステータスバーの 1 行を出し、失敗があれば名前と理由を出す
    param (
        [string]$operation,
        [object[]]$results
    )

    $view = getIndexBulkResultView $operation $results
    setStatus $view.Status
    if ($view.HasFailure) {
        showNoticeBanner "IndexNotice" "IndexOpFailed" "$($view.FailureHeading)`n$($view.FailureDetail)"
    } else {
        hideNoticeBanner "IndexNotice"
    }
}

function newBulkExportIndexes {
    # ［エクスポート…］で 2 件以上チェックしているとき。書き出し先は 1 回だけ選び、1 つ失敗しても残りを続ける
    param (
        [string[]]$names
    )

    if (!(testIndexOperable "エクスポート")) {
        return
    }
    startExportSizeCount @($script:targetItems | Where-Object { $names -contains $_.Name })
}

function showIndexImportDialog {
    # インポートのダイアログ。決めた内容 @{ Path; Name; CollisionMode } を返す（キャンセルは $null）
    param (
        [string]$suggestedName,
        $info   # readIndexArchiveInfo の結果
    )

    $dialog = loadWindow "${xamlDir}\dialog_import.xaml" ${fontsDir}
    $dialog.Owner = $window
    $dialog.Title = "インデックスのインポート"
    $ctrl = @{}
    foreach ($name in @("ImportButton", "BrowseButton", "FolderBox", "NameBox", "IntroText", "ErrorText", "CollisionPanel", "CollisionHeading", "CollisionNotice", "RenameRadio", "OverwriteRadio",
            "ImportTitleText", "ImportExportedText", "ImportSizeText", "SourceCheckText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $header = getIndexImportHeader $info
    $ctrl.ImportTitleText.Text = $header.Name
    $ctrl.ImportExportedText.Text = $header.Exported
    $ctrl.ImportExportedText.Visibility = if ($header.Exported) { "Visible" } else { "Collapsed" }
    $ctrl.ImportSizeText.Text = $header.Size
    $ctrl.IntroText.Text = getIndexImportNotice
    $ctrl.FolderBox.Text = $info.SourceFolder
    $ctrl.NameBox.Text = $suggestedName

    # ほかのダイアログと同じく、ボタンの Click からは
    # script スコープの入れ物を参照する。GetNewClosure() でこの関数のローカル変数（$ctrl・$dialog）を
    # 取り込むと、tebunko.bat の起動（powershell -Command "...; & gui.ps1"）のように呼び出しが
    # 入れ子になっている実機では、閉じ込めたスクリプトブロックから名前で関数を解決できなくなるため
    # （continueImportIndex で見つかった不具合と同じ原因。docs\design\gui\responsiveness.md の
    # 「画面を固まらせない待たせ方」の注意（PowerShell 5.1）を参照）、ここでは使わない
    # Exists は、フォルダの有無が分かっているパス（正規化して小文字にしたもの → 有無）。画面のスレッドでは調べないので、
    # zip の元のフォルダ（目録を読む別スレッドで調べた info.SourceExists）と、［参照…］で選んだフォルダだけが入る
    $exists = @{}
    if ($null -ne $info.SourceExists) {
        $exists[(normalizeFolderPath $info.SourceFolder).ToLowerInvariant()] = [bool]$info.SourceExists
    }
    # ZipName は zip の中の名前、Mode は同じ名前のときの選び方の今の値（getImportCollisionChoiceView の Mode）
    $script:importDialog = @{ Window = $dialog; Ctrl = $ctrl; Exists = $exists; ZipName = [string]$info.IndexName; Mode = ${importCollisionRename} }

    $ctrl.FolderBox.Add_TextChanged({ safe { updateImportDialogChecks } })
    $ctrl.NameBox.Add_TextChanged({ safe { updateImportDialogChecks } })
    $ctrl.RenameRadio.Add_Checked({ safe { selectImportCollisionMode ${importCollisionRename} } })
    $ctrl.OverwriteRadio.Add_Checked({ safe { selectImportCollisionMode ${importCollisionOverwrite} } })
    $ctrl.BrowseButton.Add_Click({
        safe {
            $d = $script:importDialog
            $initial = normalizeFolderPath $d.Ctrl.FolderBox.Text
            $path = selectFolder "インポートしたインデックスの、今の元のフォルダを選んでください" $initial $d.Window $false
            if ($path) {
                $d.Exists[(normalizeFolderPath $path).ToLowerInvariant()] = $true
                $d.Ctrl.FolderBox.Text = $path
            }
        }
    })
    $ctrl.ImportButton.Add_Click({
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

    updateImportDialogChecks

    $result = $null
    if ((showOwnedDialog $dialog)) {
        $choice = getImportCollisionChoiceView $ctrl.NameBox.Text $script:importDialog.ZipName (getUsedIndexNames $script:targetItems) $script:importDialog.Mode
        $result = @{ Path = (normalizeFolderPath $ctrl.FolderBox.Text); Name = $ctrl.NameBox.Text.Trim(); CollisionMode = $choice.Mode }
    }
    return $result
}

function updateImportDialogChecks {
    # インポートのダイアログの、元のフォルダの下の 1 行（見つかったか・ほかのインデックスと重なるか）と、名前が重なるときの案内を、
    # 今の入力で出し直す。文言と色の区分は getImportSourceCheckText・getImportCollisionChoiceView
    $d = $script:importDialog
    $ctrl = $d.Ctrl
    $name = $ctrl.NameBox.Text.Trim()
    $key = (normalizeFolderPath $ctrl.FolderBox.Text).ToLowerInvariant()
    $known = if ($d.Exists.ContainsKey($key)) { $d.Exists[$key] } else { $null }
    $check = getImportSourceCheckText $ctrl.FolderBox.Text $known @($script:targetItems | Where-Object { $_.Name -ine $name })
    $brush = switch ($check.Kind) { "Ok" { "Badge.Ok.Text" } "Warn" { "Banner.Warn.Text" } "Ng" { "Badge.Ng.Text" } default { "Ink.Muted" } }
    $ctrl.SourceCheckText.Text = $check.Text
    $ctrl.SourceCheckText.Foreground = themeBrush $brush
    $ctrl.SourceCheckText.Visibility = if ($check.Text) { "Visible" } else { "Collapsed" }
    $choice = getImportCollisionChoiceView $ctrl.NameBox.Text $d.ZipName (getUsedIndexNames $script:targetItems) $d.Mode
    $ctrl.CollisionHeading.Text = $choice.Heading
    $ctrl.CollisionNotice.Text = $choice.Notice
    $ctrl.CollisionNotice.Visibility = if ($choice.Notice) { "Visible" } else { "Collapsed" }
    $ctrl.RenameRadio.Content = $choice.RenameText
    $ctrl.OverwriteRadio.Content = $choice.OverwriteText
    $ctrl.CollisionPanel.Visibility = if ($choice.Visible) { "Visible" } else { "Collapsed" }
    # 名前の欄を手で替えて置き換えられなくなったときは、選び方も別の名前にするに戻す
    # （Mode を先に揃えるので、下の選び直しで selectImportCollisionMode は何もしない）
    $d.Mode = $choice.Mode
    $ctrl.RenameRadio.IsChecked = ($choice.Mode -eq ${importCollisionRename})
    $ctrl.OverwriteRadio.IsChecked = ($choice.Mode -eq ${importCollisionOverwrite})
    $ctrl.ImportButton.Style = $d.Window.FindResource($(if ($choice.Visible -and $choice.Danger) { "Danger.Filled" } else { "Primary" }))
}

function selectImportCollisionMode {
    # 同じ名前のときの選び方を利用者が選んだとき。名前の欄を、別の名前にするなら候補、置き換えるなら置き換える相手の名前にする
    # （updateImportDialogChecks が選び方を揃えたときは、Mode が同じなので何もしない）
    param (
        [string]$mode
    )

    $d = $script:importDialog
    if ($d.Mode -eq $mode) {
        return
    }
    $d.Mode = $mode
    $choice = getImportCollisionChoiceView $d.Ctrl.NameBox.Text $d.ZipName (getUsedIndexNames $script:targetItems) $mode
    $d.Ctrl.NameBox.Text = if ($mode -eq ${importCollisionOverwrite}) { $choice.OverwriteName } else { $choice.Candidate }
    updateImportDialogChecks
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
        # 元のフォルダがこの PC にあるか（画面のスレッドで調べないよう、ここで調べる。ネットワークの場所へは接続しない）
        getImportArchiveInfo $zipPath
    } @($zipPath) {
        param ($info)
        & $continueImport $zipPath $info
    }.GetNewClosure() @($zipPath)
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
        setStatus (getIndexArchiveCancelledStatus "インポート")
        return
    }

    $collisionMode = $result.CollisionMode

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
    } @($zipPath, $result.Path)
}
