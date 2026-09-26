# ［1 インデックス管理］タブのうち、インデックス作成の開始・中止と進み具合の表示。

# ---- インデックス作成の起動と進み具合 ----

function getIndexingProgress {
    # インデックス作成の進み具合を返す（インデクサが書く インデックス作成進捗.txt の1行を読む）。
    # 取り込み一覧（数万行）を読み直すと1回に数秒かかり、毎秒読むと画面が固まるため、この1行だけを読む
    param (
        [datetime]$since
    )

    $progress = @{ Scanned = $false; Processed = 0; Failed = 0; Remaining = 0; Current = ""; Detail = ""; Finishing = $false; Confirming = $false }
    $current = readIndexingProgress
    if ($null -eq $current) {
        return $progress
    }

    $progress.Scanned = ($current.Phase -ne ${indexingPhaseCrawl})
    $progress.Confirming = ($current.Phase -eq ${indexingPhaseConfirm})
    $progress.Finishing = ($current.Phase -eq ${indexingPhaseFinish})
    $progress.Processed = $current.Processed
    $progress.Remaining = $current.Remaining
    $progress.Failed = $current.Failed
    $progress.Detail = $current.Detail
    if ($progress.Scanned) {
        $progress.Current = $current.Detail  # 取り込み中のファイルの相対パス
    }
    return $progress
}

function findRunningIndexer {
    # このツールのインデックス作成（tebunko\indexer.ps1）が実行中なら、そのプロセスを返す（画面を閉じて開き直した場合など）
    $script = ${indexerScriptPath}
    foreach ($process in @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue)) {
        if ($process.CommandLine -and $process.CommandLine.IndexOf($script, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            try {
                return Get-Process -Id $process.ProcessId -ErrorAction Stop
            } catch {
            }
        }
    }
    return $null
}

function showIndexingPanel {
    $ui.IndexingProgressPanel.Visibility = "Visible"
    $ui.IndexingProgress.IsIndeterminate = $true
    $ui.IndexingProgressText.Text = "インデックス作成の準備をしています…"
    $ui.IndexingProgressEta.Text = ""
    $ui.IndexingProgressDetail.Text = "取り込み対象のファイルを確認しています。"
    $ui.IndexingStopButton.Visibility = "Visible"
    $ui.IndexingStopButton.IsEnabled = $true
    $ui.IndexingLogButton.Visibility = "Collapsed"
    $taskbar.ProgressState = "Indeterminate"
}

function buildPlanRows {
    # 取り込み予定を確認のダイアログの一覧に変える（文言は indexing_view.ps1 が決め、ここで色を付ける）
    param (
        $plan  # readIngestPlan の結果
    )

    $tones = @{ info = ${infoBrush}; ok = ${okBrush}; warn = ${warnBrush}; ng = ${ngBrush}; gray = ${grayBrush} }
    $rows = New-Object System.Collections.Generic.List[PlanRow]
    # , で包んだ戻り値は、そのまま foreach に渡すと空のときも 1 回まわるため、変数に受けてから回す
    $views = newPlanViewRows $plan
    foreach ($view in $views) {
        $row = [PlanRow]::new()
        $row.Name = $view.Name
        $row.Path = $view.Path
        $row.TargetText = $view.TargetText
        $row.TargetBrush = $tones[$view.Tone]
        $row.DetailText = $view.DetailText
        $row.TotalText = $view.TotalText
        $rows.Add($row)
    }
    return , $rows.ToArray()
}
function updateIndexingConfirmTotal {
    # 「失敗分も再取り込みする」のチェックに合わせて、合計と主ボタンの文言を変える
    $d = $script:confirmDialog
    $view = getIndexingConfirmText $d.Targets $d.Failed ([bool]$d.Ctrl.RetryCheck.IsChecked)
    $d.Ctrl.TotalText.Text = $view.Text
    $d.Ctrl.StartButton.Content = $view.Button
}
function showIndexingConfirmDialog {
    # インデックス作成の確認。インデクサが数えた結果（インデックスごとの取り込み対象の件数）を出して、取り込むかどうかを選んでもらう。
    #   取り込む → @{ RetryFailed } ／ 取りやめ → $null
    param (
        $plan  # readIngestPlan の結果
    )

    $targets = 0
    $failed = 0
    foreach ($item in @($plan)) {
        if ($item.区分 -eq ${planKindIngest}) {
            $targets += $item.取り込み対象
            $failed += $item.前回失敗
        }
    }

    $dialog = loadWindow "${xamlDir}\dialog_indexing_confirm.xaml"
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("StartButton", "CancelButton", "RetryCheck", "TotalText", "NoteText", "IntroText", "PlanGrid")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $script:confirmDialog = @{ Window = $dialog; Ctrl = $ctrl; Targets = $targets; Failed = $failed; Answer = $null }

    $ctrl.PlanGrid.ItemsSource = buildPlanRows $plan
    $ctrl.IntroText.Text = "元のファイルの更新日時とサイズを、前回取り込んだときの記録と比べました。" +
        "［インデックス作成を開始］を押すと、取り込み対象のファイルだけを取り込みます。"
    if ($failed -gt 0) {
        $ctrl.RetryCheck.Visibility = "Visible"
        $ctrl.RetryCheck.Content = "前回取り込みに失敗し、その後更新されていないファイル {0:#,0} 件も再取り込みする（パスワード付きなど）" -f $failed
    }
    if ($targets -eq 0 -and $failed -eq 0) {
        # 取り込むものが無いときは、閉じるだけ（［キャンセル］との違いが無い）
        $ctrl.CancelButton.Visibility = "Collapsed"
        $ctrl.NoteText.Visibility = "Visible"
        $ctrl.NoteText.Text = "更新日時が変わらないまま中身が変わったファイルは、取り込み対象になりません。" +
            "そのインデックスを一から作り直すときは、［削除］してから追加し直してください。"
    }
    updateIndexingConfirmTotal

    $ctrl.RetryCheck.Add_Click({ safe { updateIndexingConfirmTotal } })
    $ctrl.StartButton.Add_Click({
        safe {
            $d = $script:confirmDialog
            $d.Answer = @{ RetryFailed = [bool]$d.Ctrl.RetryCheck.IsChecked }
            $d.Window.DialogResult = $true
        }
    })
    $null = $dialog.ShowDialog()

    $answer = $script:confirmDialog.Answer
    if ($null -eq $answer -and $targets -eq 0 -and $failed -eq 0) {
        # 取り込むものが無いときは、どう閉じても同じ（インデクサはそのまま終わる）
        $answer = @{ RetryFailed = $false }
    }
    $script:confirmDialog = $null
    return $answer
}

function confirmIndexingTargets {
    # インデクサが数え終えて確認を待っている間に、確認のダイアログを1回だけ開いて返事を返す。
    # ダイアログを開いている間も進み具合のタイマーは動くため、開く前に「開いた」ことにしておく
    if ($script:indexingConfirmed) {
        return
    }
    $plan = readIngestPlan
    if ($null -eq $plan) {
        return  # 書き込みの途中・まだ読めない。次の機会に読む
    }
    $script:indexingConfirmed = $true

    $ui.IndexingProgress.IsIndeterminate = $true
    $ui.IndexingProgressText.Text = "取り込む内容を確認してください"
    $ui.IndexingProgressDetail.Text = "取り込み対象の一覧を表示しています。"
    $answer = showIndexingConfirmDialog $plan
    if ($null -eq $answer) {
        # 取りやめ。インデクサはインデックス作成中止要求を見て、何も取り込まずに終わる
        $script:indexingCanceledAtConfirm = $true
        $ui.IndexingStopButton.IsEnabled = $false
        $ui.IndexingProgressText.Text = "インデックス作成を取りやめています…"
        $ui.IndexingProgressDetail.Text = ""
        [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
        setStatus "インデックス作成を取りやめました"
        return
    }
    writeIndexingStartRequest $answer.RetryFailed
    $script:indexingRate = $null  # 残り時間の目安は、確認を待っていた時間を含めずに計る
    $ui.IndexingProgressText.Text = "インデックス作成を始めています…"
    setStatus "インデックス作成を開始しました"
}

function startIndexing {
    if (isIndexing) {
        return
    }
    $existing = findRunningIndexer
    if ($existing) {
        adoptIndexing $existing
        setStatus "実行中のインデックス作成があるため、その進み具合を表示します"
        return
    }

    # 既定のワークスペースにほかのファイルが置いてあれば、始めずに［8 設定］で別のフォルダを選んでもらう
    $workspaceBlock = getWorkspaceBlockMessage
    if ($workspaceBlock) {
        showMessage $workspaceBlock "OK" "Warning" | Out-Null
        $ui.Tabs.SelectedItem = $ui.SettingsTab
        return
    }

    saveTargets
    # 何件取り込むかは、元のファイルの更新日時とサイズを見ないと分からない。
    # -ConfirmTargets を付けると、インデクサは数え終えたところで止まって確認（インデックス作成開始要求）を待つ
    $arguments = "-NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File `"${indexerScriptPath}`" -ConfirmTargets"
    $script:indexingStart = Get-Date
    $script:indexingRate = $null
    $script:indexingConfirmed = $false
    $script:indexingCanceledAtConfirm = $false
    $script:indexingProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WorkingDirectory ${rootDir} -WindowStyle Hidden -PassThru
    # PowerShell 5.1 では、起動直後にハンドルを取っておかないと終了コードを取得できないことがある
    $null = $script:indexingProcess.Handle
    $script:indexingAdopted = $false

    showIndexingPanel
    setStatus "クロールしています…"
    updateIndexingButton
    updateKillBadge
    $script:indexingTimer.Start()
}

function adoptIndexing {
    # 画面の外で起動された（または前回の画面で起動した）インデックス作成の進み具合を表示する
    param (
        [System.Diagnostics.Process]$process
    )

    $script:indexingProcess = $process
    try {
        $null = $process.Handle
    } catch {
    }
    $script:indexingStart = $process.StartTime
    $script:indexingRate = $null
    $script:indexingConfirmed = $false
    $script:indexingCanceledAtConfirm = $false
    $script:indexingAdopted = $true
    showIndexingPanel
    updateIndexingButton
    $script:indexingTimer.Start()
}

function stopIndexing {
    if (!(isIndexing)) {
        return
    }
    $answer = showConfirm `
        -heading "インデックス作成を中止しますか？" `
        -facts @(
            (factNext "いま取り込んでいるファイルが終わったところで止まります"),
            (factKept "ここまで取り込んだ分はそのまま残ります" "次に［インデックス作成を開始］を押すと、続きから再開します")
        ) `
        -choices @(@{ Text = "中止する"; Value = "stop"; Careful = $true })
    if ($answer -ne "stop") {
        return
    }
    [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
    $ui.IndexingStopButton.IsEnabled = $false
    $ui.IndexingProgressDetail.Text = "中止しています…（取り込み中のファイルが終わるまでお待ちください）"
    setStatus "インデックス作成の中止を要求しました"
}

function updateIndexingProgress {
    if (!(isIndexing)) {
        finishIndexing
        return
    }

    try {
        $progress = getIndexingProgress $script:indexingStart
    } catch {
        return
    }
    $stopping = !$ui.IndexingStopButton.IsEnabled

    $total = $progress.Processed + $progress.Remaining
    if (!$progress.Scanned) {
        # クロールしている間（大きいフォルダ・ネットワーク越しでは数分かかることがある）。
        # 何を見ているかが分かるよう、インデクサが書いた内容をそのまま出す
        $ui.IndexingProgress.IsIndeterminate = $true
        $ui.IndexingProgressText.Text = "クロールしています…"
        if (!$stopping) {
            $ui.IndexingProgressDetail.Text = [string]$progress.Detail
        }
        $taskbar.ProgressState = "Indeterminate"
        return
    }
    if ($progress.Confirming) {
        # 数え終えて、画面で取り込むかどうかを選ぶのを待っている（インデクサは返事があるまで止まっている）。
        # 表示は confirmIndexingTargets がダイアログを開く直前に 1 回だけ変える（取りやめの表示を上書きしないため）
        $taskbar.ProgressState = "Paused"
        confirmIndexingTargets
        return
    }
    if ($progress.Finishing) {
        # 後片付け（Officeアプリの終了・取り込み一覧の書き直し）。止まって見えないよう、何をしているかを出す
        $ui.IndexingProgress.IsIndeterminate = $true
        $ui.IndexingProgressText.Text = "インデックス作成を終えています…"
        $ui.IndexingProgressDetail.Text = [string]$progress.Detail
        $taskbar.ProgressState = "Indeterminate"
        return
    }
    if ($progress.Processed -eq 0) {
        $ui.IndexingProgress.IsIndeterminate = $true
        $ui.IndexingProgressText.Text = if ($progress.Remaining -gt 0) { "$($progress.Remaining) 件のファイルを取り込みます" } else { "取り込みが必要なファイルを確認しています…" }
        if (!$stopping) {
            $ui.IndexingProgressDetail.Text = if ($progress.Current) { "取り込み中のファイル：$($progress.Current)" } else { "" }
        }
        $taskbar.ProgressState = "Indeterminate"
        return
    }

    $ratio = if ($total -gt 0) { $progress.Processed / $total } else { 1 }
    $ui.IndexingProgress.IsIndeterminate = $false
    $ui.IndexingProgress.Value = $ratio
    $taskbar.ProgressState = if ($progress.Failed -gt 0) { "Paused" } else { "Normal" }
    $taskbar.ProgressValue = $ratio

    $text = "インデックス作成中… $($progress.Processed.ToString('N0')) / $($total.ToString('N0')) 件"
    if ($progress.Failed -gt 0) {
        $text += "（失敗 $($progress.Failed) 件）"
    }
    $ui.IndexingProgressText.Text = $text

    # 残り時間の目安（最初の1件が終わってからの速さで計算する）
    $now = Get-Date
    if ($null -eq $script:indexingRate) {
        $script:indexingRate = @{ Time = $now; Processed = $progress.Processed }
    }
    $done = $progress.Processed - $script:indexingRate.Processed
    if ($progress.Remaining -eq 0) {
        $ui.IndexingProgressEta.Text = ""
    } elseif ($done -gt 0) {
        $seconds = ($now - $script:indexingRate.Time).TotalSeconds / $done * $progress.Remaining
        $ui.IndexingProgressEta.Text = if ($seconds -lt 60) { "残り 1 分未満" } else { "残り約 $([math]::Ceiling($seconds / 60)) 分" }
    }
    if (!$stopping) {
        $ui.IndexingProgressDetail.Text = if ($progress.Current) { "取り込み中のファイル：$($progress.Current)" } else { "" }
    }
}

function finishIndexing {
    $script:indexingTimer.Stop()
    $taskbar.ProgressState = "None"
    # インデックス作成完了の通知。以前はタスクバーのボタンを光らせていたが（FlashWindowEx）、P/Invoke は
    # 実行時コンパイル（csc.exe）を無くすため廃止した。完了は進捗表示・ステータスで分かる。
    $exitCode = $null
    try {
        $script:indexingProcess.WaitForExit()
        $exitCode = $script:indexingProcess.ExitCode
    } catch {
    }
    $progress = $null
    try {
        $progress = getIndexingProgress $script:indexingStart
    } catch {
    }

    $counts = ""
    if ($progress -and $progress.Processed -gt 0) {
        $counts = "成功 $($progress.Processed - $progress.Failed) 件 / 失敗 $($progress.Failed) 件"
        if ($progress.Remaining -gt 0) {
            $counts += " / 残り $($progress.Remaining) 件"
        }
        $ui.IndexingProgress.IsIndeterminate = $false
        $ui.IndexingProgress.Value = $progress.Processed / ($progress.Processed + $progress.Remaining)
    } else {
        $ui.IndexingProgress.IsIndeterminate = $false
        $ui.IndexingProgress.Value = 0
    }

    if ($exitCode -eq 1) {
        # インデックス作成を続けられないエラー（クロール対象フォルダが無い など）
        $message = (readTextShared ${indexingErrorFile}).Trim()
        if ($message -eq "") {
            $message = "詳しくはログを確認してください。"
        }
        $ui.IndexingProgressText.Text = "インデックスを作成できませんでした"
        $ui.IndexingProgressDetail.Text = $message
        setStatus "インデックスを作成できませんでした：$message"
        showMessage "インデックスを作成できませんでした。`n`n$message" "OK" "Error" | Out-Null
    } elseif ($exitCode -eq 2 -and $script:indexingCanceledAtConfirm) {
        # 確認のダイアログで取りやめた（1件も取り込んでいない）
        $ui.IndexingProgressText.Text = "インデックス作成を取りやめました"
        $ui.IndexingProgressDetail.Text = "取り込んだファイルはありません。［インデックス作成を開始］を押すと、もう一度確認できます。"
        setStatus $ui.IndexingProgressText.Text
    } elseif ($exitCode -eq 2) {
        $ui.IndexingProgressText.Text = if ($counts) { "インデックス作成を中止しました（$counts）" } else { "インデックス作成を中止しました" }
        $ui.IndexingProgressDetail.Text = "次回は続きから再開できます。"
        setStatus $ui.IndexingProgressText.Text
    } elseif ($progress -and $progress.Processed -gt 0) {
        $ui.IndexingProgressText.Text = "インデックス作成が終わりました（$counts）"
        $ui.IndexingProgressDetail.Text = if ($progress.Failed -gt 0) { "失敗したファイルと原因は「取り込みに失敗したファイル」の一覧で確認できます。" } else { "" }
        setStatus $ui.IndexingProgressText.Text
    } else {
        $ui.IndexingProgressText.Text = "取り込みが必要なファイルはありませんでした"
        $ui.IndexingProgressDetail.Text = ""
        setStatus $ui.IndexingProgressText.Text
    }
    $ui.IndexingProgressEta.Text = ""
    $ui.IndexingStopButton.Visibility = "Collapsed"
    $ui.IndexingLogButton.Visibility = if (Test-Path -LiteralPath ${indexingLogFile}) { "Visible" } else { "Collapsed" }

    $script:indexingProcess = $null
    $script:sourceFolderMaps = @{}
    refreshIndexingState
    refreshIndexSummary
    loadIndexTree  # 新しいインデックス・フォルダをツリーに出す
    updateKillBadge
}

$script:indexingTimer = newTimer 1000 { safe { updateIndexingProgress } }

$ui.FailedGrid.Add_MouseDoubleClick({
    param ($sender, $e)
    # 行の上でのダブルクリックだけを対象にする（列見出し・スクロールバーは除く）
    $element = $e.OriginalSource
    while ($element -and !($element -is [System.Windows.Controls.DataGridRow])) {
        if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
            return
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    if ($element) {
        safe { openFailedFileFolder }
    }
})
$ui.IndexingButton.Add_Click({ safe { startIndexing } })
$ui.IndexingStopButton.Add_Click({ safe { stopIndexing } })
$ui.IndexingLogButton.Add_Click({
    safe {
        if (Test-Path -LiteralPath ${indexingLogFile}) {
            Invoke-Item -LiteralPath ${indexingLogFile}
        }
    }
})
