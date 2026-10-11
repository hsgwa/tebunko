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

function invokeWorkspaceChange {
    # ワークスペースを変える 3 つの入口（［変更…］・［既定に戻す］・欄の入力）を、同時に 1 つしか走らせない。
    # 確認の画面を出している間に、欄から出た知らせやボタンの押下が重なって、確認が 2 枚出たり、古い状態の確かめのまま進んだりしない
    param ([scriptblock]$change)
    if ($script:workspaceCommitting) {
        return
    }
    $script:workspaceCommitting = $true
    try {
        & $change
    } finally {
        $script:workspaceCommitting = $false
    }
}

function chooseWorkspace {
    # ［変更…］。空のフォルダを選んで、ワークスペースにする（空でなければ警告する）
    invokeWorkspaceChange {
        try {
            if (!(testWorkspaceChangeable)) {
                return
            }
            $folder = selectFolder "ワークスペースにする空のフォルダを選んでください。インデックス・ログをここに置きます。" $workspace.Dir
            if ($null -eq $folder) {
                return
            }
            applyWorkspace $folder $true
        } finally {
            # どの終わり方でも、入力途中の文字と知らせを消し、欄を今の場所に戻す（切り替えたときは新しい場所）
            $ui.WorkspaceBox.Text = $workspace.Dir
            setWorkspaceInputError ""
        }
    }
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
    invokeWorkspaceChange { commitWorkspaceInputCore $keepText }
}

function commitWorkspaceInputCore {
    param ([bool]$keepText)
    $current = $workspace.Dir
    $text = $ui.WorkspaceBox.Text
    if ($text -ceq $current) {
        setWorkspaceInputError ""
        return
    }
    # 更新中などは反映しない（判断層が Blocked で返す。理由は［変更…］と同じ文言）。
    # 欄から出たときは、割り込まないよう、メッセージを出さず今の場所に戻す。Enter のときは、理由を欄の下に出し、入力した文字は残す
    $blocker = getCurrentIndexJobBlocker (testIndexerRunning)
    $blockedMessage = if ($blocker -ne "") { getIndexJobBlockedMessage $blocker "ワークスペースの変更" } else { "" }
    $decision = getWorkspaceInputDecision $text $current $null $null $blockedMessage
    if ($decision.Action -eq "None") {
        $ui.WorkspaceBox.Text = $current
        setWorkspaceInputError ""
        return
    }
    if ($decision.Action -eq "Invalid" -or $decision.Action -eq "Blocked") {
        if ($keepText) {
            setWorkspaceInputError $decision.Message
        } else {
            $ui.WorkspaceBox.Text = $current
            # 直してほしい入力の理由は、欄から出たあとも出しておく。更新中（Blocked）は割り込まないよう、出さない
            setWorkspaceInputError $(if ($decision.Action -eq "Invalid") { $decision.Message } else { "" })
        }
        return
    }
    setWorkspaceInputError ""
    try {
        applyWorkspace $decision.Path $true
    } finally {
        # どの終わり方でも、欄を今の場所に戻す（切り替えたときは updateSettingsView が新しい場所を入れる）
        $ui.WorkspaceBox.Text = $workspace.Dir
    }
}

function resetWorkspace {
    # ［既定に戻す］。既定の場所（ドキュメントの tebunko_ws）に戻す（無ければ作る）。ほかのファイルが置いてあれば戻さない
    invokeWorkspaceChange {
        try {
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
        } finally {
            # どの終わり方でも、入力途中の文字と知らせを消し、欄を今の場所に戻す（切り替えたときは新しい場所）
            $ui.WorkspaceBox.Text = $workspace.Dir
            setWorkspaceInputError ""
        }
    }
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
    startWorkspaceMove $mode $previous $folder
}

function updateMainContentOperable {
    # 移している間は、ナビと中身を操作できなくする（マウス・Tab・Space など、中身の入口をまとめて止める。窓のキーは gui_main.ps1 の PreviewKeyDown）。
    # 終わったら必ず戻す（始められなかったとき・失敗のときも）
    $operable = testMainContentOperable ([bool]$script:workspaceMoving)
    $ui.NavHost.IsEnabled = $operable
    $ui.ContentHost.IsEnabled = $operable
}

function startWorkspaceMove {
    # 中身を移す・消す処理は、量によっては数十秒かかるため、裏のスレッドで行う（画面のスレッドで行うと「応答なし」になる）。
    # 動かしている間は待つ間のダイアログを出し、ほかの操作は止める（$script:workspaceMoving。getIndexJobBlocker が見る）。
    # 終わりの処理は finishWorkspaceMove（画面のスレッド）。閉じるのは、移し終わるまで待つ（gui_main.ps1 の Closing）
    param (
        [string]$mode,       # "use"・"reset"・"move"（WorkspaceMover.ChangeWorkspace の mode）
        [string]$previous,
        [string]$folder
    )

    $script:workspaceMoving = $true
    $script:closeAfterWorkspaceMove = $false
    try {
        updateMainContentOperable
        updateIndexingButton
        setStatus "ワークスペースを「${folder}」へ移しています…"
        # ダイアログが出せなくても移動は続ける
        try { showIndexingWait "Moving" } catch { }
        # 裏のスレッドは lib.ps1 を読み込んだときの場所を覚えているため、場所は渡す。
        # 移動元と移動先のどちらかがネットワークの場所なら、届かない共有でほかの仕事を待たせない専用の列を使う
        # 終わったときの処理は、あとからタイマーで呼ばれる。この関数のローカル変数（$mode・$previous・$folder）は
        # GetNewClosure() で今の値を取り込み、関数は変数に取って & で呼ぶ（クロージャからは関数の名前を引けないため）
        $finish = ${function:finishWorkspaceMove}
        startJob {
            param ($mode, $previous, $folder, $settingsPath)
            # 届かない場所では、何もせずに成功にしない。失敗として返す（移動元は、使わない "use" のときは調べない）
            if ($mode -ne "use") {
                assertWorkspaceReachable $previous
            }
            assertWorkspaceReachable $folder
            $mover = [WorkspaceMover]::new($settingsPath)
            $mover.ChangeWorkspace($mode, $previous, $folder)
        } @($mode, $previous, $folder, ${settingsFile}) {
            param ($output, $errorText)
            & $finish $mode $previous $folder $output $errorText
        }.GetNewClosure() (getWorkspaceJobQueue @($previous, $folder))
    } catch {
        # 始められなかった。印を戻し、ダイアログを閉じて、元の画面に戻す
        $script:workspaceMoving = $false
        try { updateMainContentOperable } catch { }
        try { closeIndexingWait } catch { }
        try { updateIndexingButton } catch { }
        throw
    }
}

function finishWorkspaceMove {
    # 裏の移動が終わったときの、画面のスレッドでの後始末（成功・失敗とも）。どの終わり方でも、印を下ろしてダイアログを閉じる。
    # 成功のときだけ画面を切り替える。移している間に窓を閉じようとしていたなら、最後に閉じる
    param (
        [string]$mode,
        [string]$previous,
        [string]$folder,
        $output,
        [string]$errorText
    )

    $closeRequested = [bool]$script:closeAfterWorkspaceMove
    $script:workspaceMoving = $false
    $script:closeAfterWorkspaceMove = $false
    try { updateMainContentOperable } catch { }
    $outcome = @{ Ok = $false; Warning = "ワークスペースを移した結果を確かめられませんでした。"; Status = ""; Close = $false }
    try {
        $outcome = getWorkspaceMoveOutcome (getWorkspaceMoveChange $output) $errorText $closeRequested $mode $folder $previous
        try { closeIndexingWait } catch { }
        if ($outcome.Ok) {
            if (!$outcome.Close) {
                switchWorkspace
            }
            if ($outcome.Status -ne "") {
                setStatus $outcome.Status
            }
            return
        }
        $ui.WorkspaceBox.Text = $workspace.Dir
        updateIndexingButton
        # 失敗のときは閉じない（閉じるのを待っていても、失敗を知らせるために画面を残す）
        showMessage $outcome.Warning "OK" "Warning" | Out-Null
    } finally {
        try { closeIndexingWait } catch { }
        if ($outcome.Close) {
            $window.Dispatcher.BeginInvoke([action]{ $window.Close() }) | Out-Null
        }
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
