# 設定の画面（ワークスペース・設定ファイルの場所。xaml\settings\settings.xaml）。文言と可否の判定は settings_view.ps1。

${workspaceCountLimit} = 1000      # 選んだフォルダの中身を数える上限（大きなフォルダで待たせない）

function updateSettingsView {
    # ワークスペースと設定ファイルの場所を表示する
    $view = getWorkspaceView $workspace.Dir (getDefaultWorkDir)
    $ui.WorkspaceBox.Text = $view.Path
    $ui.WorkspaceBox.ToolTip = $view.Path  # 長いパスは欄に入りきらないため、全体はツールチップで見せる
    setWorkspaceInputError ""
    $ui.WorkspaceNote.Text = $view.Note
    $caution = getWorkspaceCautionTexts
    $ui.WorkspaceCaution.Text = $caution.Caution
    $ui.WorkspaceHowTo.Text = $caution.HowTo
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
    $blocker = getCurrentIndexJobBlocker (testIndexerRunning)
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
    $ui.WorkspaceBox.Text = $workspace.Dir
}

function setWorkspaceInputError {
    # ワークスペースの欄の下に、直してほしい内容を出す（空なら隠す）
    param ([string]$message)
    $ui.WorkspaceInputError.Text = $message
    $ui.WorkspaceInputError.Visibility = if ($message) { "Visible" } else { "Collapsed" }
}

function commitWorkspaceInput {
    # 欄に入力したパスを確定する。確かめ・確認・切り替えは、［変更…］で選んだときと同じ applyWorkspace を通す。
    # 直してほしいときは欄の下に理由を出し、欄の文字は直せるように残す（$keepText。Enter のとき）。欄から出たときは今の場所に戻す
    param ([bool]$keepText)
    if ($script:workspaceCommitting) {
        return   # 確認の画面を出している間に、欄から出た知らせが重なって、もう一度確認しない
    }
    $script:workspaceCommitting = $true
    try {
        commitWorkspaceInputCore $keepText
    } finally {
        $script:workspaceCommitting = $false
    }
}

function commitWorkspaceInputCore {
    param ([bool]$keepText)
    $current = $workspace.Dir
    $text = $ui.WorkspaceBox.Text
    if ($text -ceq $current) {
        setWorkspaceInputError ""
        return
    }
    # 更新中などは反映しない（入力の途中の欄から出たときに割り込まないよう、メッセージは出さず今の場所に戻す）
    if ((getCurrentIndexJobBlocker (testIndexerRunning)) -ne "") {
        $ui.WorkspaceBox.Text = $current
        setWorkspaceInputError ""
        return
    }
    $decision = getWorkspaceInputDecision $text $current
    if ($decision.Action -eq "None") {
        $ui.WorkspaceBox.Text = $current
        setWorkspaceInputError ""
        return
    }
    if ($decision.Action -eq "Invalid") {
        if (!$keepText) {
            $ui.WorkspaceBox.Text = $current
        }
        setWorkspaceInputError $decision.Message
        return
    }
    setWorkspaceInputError ""
    applyWorkspace $decision.Path $true
    # 取りやめ・使えないときは、欄を今の場所に戻す（切り替えたときは updateSettingsView が新しい場所を入れる）
    $ui.WorkspaceBox.Text = $workspace.Dir
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
    $mover = [WorkspaceMover]::new(${settingsFile})
    $workspaceNames = @($mover.ListEntries($folder) | ForEach-Object { [System.IO.Path]::GetFileName($_) })
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

    # 本文インデックスのファイルを読んでいる検索があると移せないため、先に止める
    clearSearchView
    $previous = $workspace.Dir
    # 共有されたワークスペースなどを使う（use）・消して最初から（reset）・中身を移す（move）。保存まで WorkspaceMover が行う
    $mode = if ($answer -eq "use" -or $answer -eq "reset") { $answer } else { "move" }
    $change = $mover.ChangeWorkspace($mode, $previous, $folder)
    if (!$change.Ok) {
        if ($change.Stage -eq "remove") {
            showMessage "「${folder}」のインデックスを削除できませんでした（$($change.Message)）。ファイルを開いているアプリを閉じてから、もう一度変えてください。" "OK" "Warning" | Out-Null
        } else {
            showMessage $change.Message "OK" "Warning" | Out-Null
        }
        return
    }
    switchWorkspace
    if ($mode -eq "use") {
        setStatus "ワークスペースを「${folder}」に変え、そこにあるインデックスを使います（インデックス $($change.Targets) 件。前のワークスペースの中身は「${previous}」に残しています）"
        return
    }
    $count = $change.Count
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
    updateWorkspaceBanner
    hideNoticeBanner "SearchNotice"
    hideNoticeBanner "IndexNotice"
    $script:indexingState = $null
    $script:indexSummary = $null
    resetIndexingBanner
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

$ui.WorkspaceBox.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe { commitWorkspaceInput $true }
        $e.Handled = $true
    } elseif ($e.Key -eq "Escape") {
        safe { $ui.WorkspaceBox.Text = $workspace.Dir; setWorkspaceInputError "" }
        $e.Handled = $true
    }
})
# 欄から出た知らせの中で確認の画面を出すと、フォーカスを動かした側（操作を送ったもの）が待たされるため、知らせが終わってから確定する
$ui.WorkspaceBox.Add_LostFocus({
    $window.Dispatcher.BeginInvoke([action]{ safe { commitWorkspaceInput $false } }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
})
$ui.ChangeWorkspaceButton.Add_Click({ safe { chooseWorkspace } })
$ui.ResetWorkspaceButton.Add_Click({ safe { resetWorkspace } })
