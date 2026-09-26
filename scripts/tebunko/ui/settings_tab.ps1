# ［8 設定］タブ（ワークスペース・設定ファイルの場所）。文言と可否の判定は settings_view.ps1。

${workspaceCountLimit} = 1000      # 選んだフォルダの中身を数える上限（大きなフォルダで待たせない）

function updateSettingsView {
    # ワークスペースと設定ファイルの場所を表示する
    $view = getWorkspaceView $workspace.Dir (getDefaultWorkDir)
    $ui.WorkspaceText.Text = $view.Path
    $ui.WorkspaceNote.Text = $view.Note
    $ui.ResetWorkspaceButton.Visibility = if ($view.CanReset) { "Visible" } else { "Collapsed" }

    $file = getSettingsFileView ${settingsFile} ${rootDir}
    $ui.SettingsFileText.Text = $file.Path
    $ui.SettingsFileNote.Text = $file.Note
}

function testWorkspaceChangeable {
    # インデックス作成中はワークスペースを変えない（インデクサが今のワークスペースに書いている。画面を使わずに起動したものも含む）
    if ((isIndexing) -or (testIndexerRunning)) {
        showMessage "インデックス作成中はワークスペースを変えられません。インデックス作成が終わるまでお待ちください（［中止］で止められます）。" "OK" "Warning" | Out-Null
        return $false
    }
    if ($script:indexBusy) {
        showMessage "前のインデックスの削除が終わるまでお待ちください。" "OK" "Warning" | Out-Null
        return $false
    }
    return $true
}

function chooseWorkspace {
    # ［変更…］。空のフォルダを選んで、ワークスペースにする（空でなければ警告する）
    if (!(testWorkspaceChangeable)) {
        return
    }
    $folder = selectFolder "ワークスペースにする空のフォルダを選んでください。インデックス・取り込み一覧・ログをここに置きます。" $workspace.Dir
    if ($null -eq $folder) {
        return
    }
    applyWorkspace $folder $true
}

function resetWorkspace {
    # ［既定に戻す］。既定の場所（ドキュメントの tebunko）に戻す（無ければ作る）。ほかのファイルが置いてあれば戻さない
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
        [string]$folder
    )

    $count = 0
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($entry in [System.IO.Directory]::EnumerateFileSystemEntries((toLongPath $folder))) {
        $count++
        if ($names.Count -lt ${workspaceSampleCount}) {
            $names.Add([System.IO.Path]::GetFileName($entry))
        }
        if ($count -ge ${workspaceCountLimit}) {
            break
        }
    }
    return @{ Count = $count; Names = $names.ToArray(); Capped = ($count -ge ${workspaceCountLimit}) }
}

function applyWorkspace {
    # 確かめてからワークスペースを保存し、画面をそのワークスペースに切り替える
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
    $entries = if ($requireEmpty) { getFolderEntrySample $folder } else { @{ Count = 0; Names = @(); Capped = $false } }
    $sub = Join-Path $folder ${workspaceSubFolderName}
    $canMakeSub = -not (Test-Path -LiteralPath $sub) -or
        ((Test-Path -LiteralPath $sub -PathType Container) -and (getFolderEntrySample $sub).Count -eq 0)
    $confirm = newWorkspaceConfirm $folder $workspace.Dir $entries.Count $entries.Names $entries.Capped `
        (Test-Path -LiteralPath (Join-Path $folder "index") -PathType Container) $canMakeSub
    # switch の中の $_ は switch の値になるため、行を変数に受けてから使う
    $facts = @($confirm.Facts | ForEach-Object {
        $fact = $_
        switch ($fact.Kind) {
            "kept"  { factKept $fact.Title $fact.Detail }
            "warn"  { factWarn $fact.Title $fact.Detail }
            default { factNext $fact.Title $fact.Detail }
        }
    })
    $answer = showConfirm -heading $confirm.Heading -facts $facts -hint $confirm.Hint -choices $confirm.Choices
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

    writeWorkspaceFolder $folder
    switchWorkspace
}

function switchWorkspace {
    # 設定のワークスペースに切り替える。画面は開き直さない。
    # 関数は既定値で $workspace の場所を使い、裏のスレッドには場所を渡しているため、$workspace を差し替えれば新しい場所を使う
    # （docs/00_共通_2_共通モジュール.md 5.1.1）。インデックス作成中・削除中は testWorkspaceChangeable が止めている
    clearSearchView
    $script:workspace = [Workspace]::new((getWorkDir))
    $script:workspaceBlock = getWorkspaceBlockMessage
    $script:indexingState = $null
    $script:indexSummary = $null
    $ui.IndexingProgressPanel.Visibility = "Collapsed"
    loadWorkspaceViews
    setStatus "ワークスペースを「$($workspace.Dir)」に切り替えました"
}

# ---- イベント ----

$ui.ChangeWorkspaceButton.Add_Click({ safe { chooseWorkspace } })
$ui.ResetWorkspaceButton.Add_Click({ safe { resetWorkspace } })
