# 設定の画面（ワークスペース・設定ファイルの場所。xaml\settings\settings.xaml）。文言と可否の判定は settings_view.ps1。

${workspaceCountLimit} = 1000      # 選んだフォルダの中身を数える上限（大きなフォルダで待たせない）

function updateSettingsView {
    # ワークスペースと設定ファイルの場所を表示する
    $view = getWorkspaceView $workspace.Dir (getDefaultWorkDir)
    $ui.WorkspaceText.Text = $view.Path
    $ui.WorkspaceText.ToolTip = $view.Path  # 長いパスは省略して出すため、全体はツールチップで見せる
    $ui.WorkspaceNote.Text = $view.Note
    $ui.ResetWorkspaceButton.Visibility = if ($view.CanReset) { "Visible" } else { "Collapsed" }

    $file = getSettingsFileView ${settingsFile} ${rootDir}
    $ui.SettingsFileText.Text = $file.Path
    $ui.SettingsFileText.ToolTip = $file.Path
    $ui.SettingsFileNote.Text = $file.Note
}

function testWorkspaceChangeable {
    # インデックス作成中（インデクサが今のワークスペースに書いている。画面を使わずに起動したものも含む）・
    # 前のインデックスの削除中・エクスポート・インポート中（別スレッド。今のワークスペースの content_index・取り込み一覧・設定を使っている）は、
    # ワークスペースを変えない。可否と文言は、インデックス管理の操作と同じ判断層（getIndexJobBlocker・getIndexJobBlockedMessage）で決める
    $blocker = getIndexJobBlocker ((isIndexing) -or (testIndexerRunning)) $script:indexBusy $script:archiveBusy
    if ($blocker -ne "") {
        showMessage (getIndexJobBlockedMessage $blocker "ワークスペースの変更") "OK" "Warning" | Out-Null
        return $false
    }
    return $true
}

function chooseWorkspace {
    # ［変更…］。空のフォルダを選んで、ワークスペースにする（空でなければ警告する）
    if (!(testWorkspaceChangeable)) {
        return
    }
    $folder = selectFolder "ワークスペースにする空のフォルダを選んでください。インデックス・ログをここに置きます。" $workspace.Dir
    if ($null -eq $folder) {
        return
    }
    applyWorkspace $folder $true
}

function resetWorkspace {
    # ［既定に戻す］。既定の場所（ドキュメントの tebunko_ws）に戻す（無ければ作る）。ほかのファイルが置いてあれば戻さない
    if (!(testWorkspaceChangeable)) {
        return
    }
    $folder = getDefaultWorkDir
    $check = testDefaultWorkspace $folder
    if (!$check.Usable) {
        showMessage $check.Message "OK" "Warning" | Out-Null
        return
    }
    try {
        [System.IO.Directory]::CreateDirectory($folder) | Out-Null
    } catch {
        # 作れなければ、書き込めないフォルダとして下の確認で伝える
    }
    applyWorkspace $folder $false
}

function getFolderEntrySample {
    # フォルダの中のファイル・フォルダを上限まで数え、先頭の名前を返す: @{ Count; Names; Capped }
    param (
        [string]$folder,
        [bool]$isDefaultWorkspace = $false   # 既定のワークスペースなら、設定ファイルと付いてできるファイルを数えない（getCountedEntryPaths）
    )

    $entries = getCountedEntryPaths (selectFirstEntries ([System.IO.Directory]::EnumerateFileSystemEntries((toLongPath $folder))) ${workspaceCountLimit}) $isDefaultWorkspace
    $count = $entries.Count
    $names = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt [Math]::Min($count, ${workspaceSampleCount}); $i++) {
        $names.Add([System.IO.Path]::GetFileName($entries[$i]))
    }
    return @{ Count = $count; Names = $names.ToArray(); Capped = ($count -ge ${workspaceCountLimit}) }
}

function applyWorkspace {
    # 確かめてから、今のワークスペースの中身を移してワークスペースを保存し、画面をそのワークスペースに切り替える。
    # 選んだフォルダにインデックスなどがあれば、それを使う（中身は移さない）か、消して最初からやり直す（消してから移す）かを選ばせる
    param (
        [string]$folder,
        [bool]$requireEmpty   # 空のフォルダを求める（［変更…］）。空でなければ警告する
    )

    $check = testWorkspaceChoice $folder $workspace.Dir (testWritableFolder $folder)
    if ($check.Kind -eq "same") {
        setStatus "今と同じワークスペースです"
        return
    }
    if ($check.Kind -eq "error") {
        showMessage $check.Message "OK" "Warning" | Out-Null
        return
    }

    $folder = normalizeFolderPath $folder
    $entries = if ($requireEmpty) { getFolderEntrySample $folder (testSameFolder $folder (getDefaultWorkDir)) } else { @{ Count = 0; Names = @(); Capped = $false } }
    $sub = Join-Path $folder ${workspaceSubFolderName}
    $canMakeSub = -not (Test-Path -LiteralPath $sub) -or
        ((Test-Path -LiteralPath $sub -PathType Container) -and (getFolderEntrySample $sub).Count -eq 0)
    $workspaceNames = @(getWorkspaceEntries $folder | ForEach-Object { [System.IO.Path]::GetFileName($_) })
    $confirm = newWorkspaceConfirm $folder $workspace.Dir $entries.Count $entries.Names $entries.Capped $workspaceNames $canMakeSub (!$requireEmpty)
    # switch の中の $_ は switch の値になるため、行を変数に受けてから使う
    $facts = @($confirm.Facts | ForEach-Object {
        $fact = $_
        switch ($fact.Kind) {
            "kept"  { factKept $fact.Title $fact.Detail }
            "warn"  { factWarn $fact.Title $fact.Detail }
            default { factNext $fact.Title $fact.Detail }
        }
    })
    $answer = showConfirm -title $confirm.Title -heading $confirm.Heading -facts $facts -hint $confirm.Hint -choices $confirm.Choices
    if ($null -eq $answer) {
        return
    }
    if ($answer -eq "sub") {
        [System.IO.Directory]::CreateDirectory($sub) | Out-Null
        if (-not (testWritableFolder $sub)) {
            showMessage "「${sub}」にはファイルを作れません。書き込めるフォルダを選んでください。" "OK" "Warning" | Out-Null
            return
        }
        $folder = $sub
    }

    # 集約ファイルを読んでいる検索があると移せないため、先に止める
    clearSearchView
    $previous = $workspace.Dir
    if ($answer -eq "use") {
        # 共有されたワークスペースなどを使う。今の中身は移さず、インデックスの一覧をこのワークスペースのものにする
        $targets = useWorkspaceTargets $folder
        writeWorkspaceFolder $folder
        switchWorkspace
        setStatus "ワークスペースを「${folder}」に変え、そこにあるインデックスを使います（インデックス $targets 件。前のワークスペースの中身は「${previous}」に残しています）"
        return
    }
    if ($answer -eq "reset") {
        try {
            [void](removeWorkspaceEntries $folder)
        } catch {
            showMessage "「${folder}」のインデックスを削除できませんでした（$($_.Exception.Message)）。ファイルを開いているアプリを閉じてから、もう一度変えてください。" "OK" "Warning" | Out-Null
            return
        }
    }
    try {
        $count = moveWorkspace $previous $folder
    } catch {
        showMessage $_.Exception.Message "OK" "Warning" | Out-Null
        return
    }
    # 検索対象ツリーでチェックを外したフォルダも、移した先のインデックスに付け替える
    [void](moveSearchExcludes $previous $folder)
    writeWorkspaceFolder $folder
    switchWorkspace
    if ($count -gt 0) {
        setStatus "ワークスペースを「${folder}」に変え、中身を移しました（移す前の場所：${previous}）"
    }
}

function switchWorkspace {
    # 設定のワークスペースに切り替える。画面は開き直さない。
    # 関数は既定値で $workspace の場所を使い、裏のスレッドには場所を渡しているため、$workspace を差し替えれば新しい場所を使う
    # （docs/design/structure/data.md「ワークスペースの中の場所（Workspace）」）。インデックス作成中・削除中は testWorkspaceChangeable が止めている
    clearSearchView
    $script:workspace = [Workspace]::new((getWorkDir))
    $script:workspaceBlock = getWorkspaceBlockMessage
    $script:indexingState = $null
    $script:indexSummary = $null
    $ui.IndexingProgressPanel.Visibility = "Collapsed"
    updateSettingsView
    loadWorkspaceViews
    # 前の版のインデックスの知らせ。ネットワークの場所は裏で調べるため、分かったときにステータスへ出す
    refreshLegacyIndexMessage
    if ($script:legacyIndexMessage) {
        setStatus $script:legacyIndexMessage
    } else {
        setStatus "ワークスペースを「$($workspace.Dir)」に切り替えました"
    }
}

# ---- イベント ----

$ui.ChangeWorkspaceButton.Add_Click({ safe { chooseWorkspace } })
$ui.ResetWorkspaceButton.Add_Click({ safe { resetWorkspace } })
