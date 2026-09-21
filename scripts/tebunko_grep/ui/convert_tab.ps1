# ［1 インデックス管理］タブのうち、変換の開始・中止と進み具合の表示。

# ---- 変換の起動と進み具合 ----

function getConversionProgress {
    # 変換の進み具合を返す（変換側が書く 変換進捗.txt の1行を読む）。
    # 変換一覧（数万行）を読み直すと1回に数秒かかり、毎秒読むと画面が固まるため、この1行だけを読む
    param (
        [datetime]$since
    )

    $progress = @{ Scanned = $false; Processed = 0; Failed = 0; Remaining = 0; Current = ""; Detail = ""; Finishing = $false; Confirming = $false }
    $current = readConvertProgress
    if ($null -eq $current) {
        return $progress
    }

    $progress.Scanned = ($current.Phase -ne ${convertPhaseScan})
    $progress.Confirming = ($current.Phase -eq ${convertPhaseConfirm})
    $progress.Finishing = ($current.Phase -eq ${convertPhaseFinish})
    $progress.Processed = $current.Processed
    $progress.Remaining = $current.Remaining
    $progress.Failed = $current.Failed
    $progress.Detail = $current.Detail
    if ($progress.Scanned) {
        $progress.Current = $current.Detail  # 変換中のファイルの相対パス
    }
    return $progress
}

function findRunningConversion {
    # このツールの変換（tebunko_grep\convert.ps1）が実行中なら、そのプロセスを返す（画面を閉じて開き直した場合など）
    $script = "${PSScriptRoot}\tebunko_grep\convert.ps1"
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

function showConversionPanel {
    $ui.ConvertProgressPanel.Visibility = "Visible"
    $ui.ConvertProgress.IsIndeterminate = $true
    $ui.ConvertProgressText.Text = "変換の準備をしています…"
    $ui.ConvertProgressEta.Text = ""
    $ui.ConvertProgressDetail.Text = "変換対象のファイルを確認しています。"
    $ui.ConvertStopButton.Visibility = "Visible"
    $ui.ConvertStopButton.IsEnabled = $true
    $ui.ConvertLogButton.Visibility = "Collapsed"
    $taskbar.ProgressState = "Indeterminate"
}

function buildPlanRows {
    # 変換予定を確認のダイアログの一覧に変える（文言は convert_view.ps1 が決め、ここで色を付ける）
    param (
        $plan  # readConvertPlan の結果
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
function updateConvertConfirmTotal {
    # 「失敗分も再変換する」のチェックに合わせて、合計と主ボタンの文言を変える
    $d = $script:confirmDialog
    $view = getConvertConfirmText $d.Targets $d.Failed ([bool]$d.Ctrl.RetryCheck.IsChecked)
    $d.Ctrl.TotalText.Text = $view.Text
    $d.Ctrl.StartButton.Content = $view.Button
}
function showConvertConfirmDialog {
    # 変換の確認。変換側が数えた結果（インデックスごとの変換対象の件数）を出して、変換するかどうかを選んでもらう。
    #   変換する → @{ RetryFailed } ／ 取りやめ → $null
    param (
        $plan  # readConvertPlan の結果
    )

    $targets = 0
    $failed = 0
    foreach ($item in @($plan)) {
        if ($item.区分 -eq ${planKindConvert}) {
            $targets += $item.変換対象
            $failed += $item.前回失敗
        }
    }

    $dialog = loadWindow "${xamlDir}\dialog_convert_confirm.xaml"
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("StartButton", "CancelButton", "RetryCheck", "TotalText", "NoteText", "IntroText", "PlanGrid")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $script:confirmDialog = @{ Window = $dialog; Ctrl = $ctrl; Targets = $targets; Failed = $failed; Answer = $null }

    $ctrl.PlanGrid.ItemsSource = buildPlanRows $plan
    $ctrl.IntroText.Text = "元のファイルの更新日時とサイズを、前回変換したときの記録と比べました。" +
        "［変換を開始］を押すと、変換対象のファイルだけを変換します。"
    if ($failed -gt 0) {
        $ctrl.RetryCheck.Visibility = "Visible"
        $ctrl.RetryCheck.Content = "前回変換に失敗し、その後更新されていないファイル {0:#,0} 件も再変換する（パスワード付きなど）" -f $failed
    }
    if ($targets -eq 0 -and $failed -eq 0) {
        # 変換するものが無いときは、閉じるだけ（［キャンセル］との違いが無い）
        $ctrl.CancelButton.Visibility = "Collapsed"
        $ctrl.NoteText.Visibility = "Visible"
        $ctrl.NoteText.Text = "更新日時が変わらないまま中身が変わったファイルは、変換対象になりません。" +
            "そのインデックスを一から作り直すときは、［削除］してから作成し直してください。"
    }
    updateConvertConfirmTotal

    $ctrl.RetryCheck.Add_Click({ safe { updateConvertConfirmTotal } })
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
        # 変換するものが無いときは、どう閉じても同じ（変換側はそのまま終わる）
        $answer = @{ RetryFailed = $false }
    }
    $script:confirmDialog = $null
    return $answer
}

function confirmConversionTargets {
    # 変換側が数え終えて確認を待っている間に、確認のダイアログを1回だけ開いて返事を返す。
    # ダイアログを開いている間も進み具合のタイマーは動くため、開く前に「開いた」ことにしておく
    if ($script:convertConfirmed) {
        return
    }
    $plan = readConvertPlan
    if ($null -eq $plan) {
        return  # 書き込みの途中・まだ読めない。次の機会に読む
    }
    $script:convertConfirmed = $true

    $ui.ConvertProgress.IsIndeterminate = $true
    $ui.ConvertProgressText.Text = "変換する内容を確認してください"
    $ui.ConvertProgressDetail.Text = "変換対象の一覧を表示しています。"
    $answer = showConvertConfirmDialog $plan
    if ($null -eq $answer) {
        # 取りやめ。変換側は変換中止要求を見て、何も変換せずに終わる
        $script:convertCanceledAtConfirm = $true
        $ui.ConvertStopButton.IsEnabled = $false
        $ui.ConvertProgressText.Text = "変換を取りやめています…"
        $ui.ConvertProgressDetail.Text = ""
        [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
        setStatus "変換を取りやめました"
        return
    }
    writeConvertStartRequest $answer.RetryFailed
    $script:convertRate = $null  # 残り時間の目安は、確認を待っていた時間を含めずに計る
    $ui.ConvertProgressText.Text = "変換を始めています…"
    setStatus "変換を開始しました"
}

function startConversion {
    if (isConverting) {
        return
    }
    $existing = findRunningConversion
    if ($existing) {
        adoptConversion $existing
        setStatus "実行中の変換があるため、その進み具合を表示します"
        return
    }

    saveTargets
    # 何件変換するかは、元のファイルの更新日時とサイズを見ないと分からない。
    # -ConfirmTargets を付けると、変換側は数え終えたところで止まって確認（変換開始要求）を待つ
    $arguments = "-NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File `"${PSScriptRoot}\tebunko_grep\convert.ps1`" -ConfirmTargets"
    $script:convertStart = Get-Date
    $script:convertRate = $null
    $script:convertConfirmed = $false
    $script:convertCanceledAtConfirm = $false
    $script:convertProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WorkingDirectory ${rootDir} -WindowStyle Hidden -PassThru
    # PowerShell 5.1 では、起動直後にハンドルを取っておかないと終了コードを取得できないことがある
    $null = $script:convertProcess.Handle
    $script:convertAdopted = $false

    showConversionPanel
    setStatus "変換対象を確認しています…"
    updateConvertButton
    updateKillBadge
    $script:convertTimer.Start()
}

function adoptConversion {
    # 画面の外で起動された（または前回の画面で起動した）変換の進み具合を表示する
    param (
        [System.Diagnostics.Process]$process
    )

    $script:convertProcess = $process
    try {
        $null = $process.Handle
    } catch {
    }
    $script:convertStart = $process.StartTime
    $script:convertRate = $null
    $script:convertConfirmed = $false
    $script:convertCanceledAtConfirm = $false
    $script:convertAdopted = $true
    showConversionPanel
    updateConvertButton
    $script:convertTimer.Start()
}

function stopConversion {
    if (!(isConverting)) {
        return
    }
    $answer = showConfirm `
        -heading "変換を中止しますか？" `
        -facts @(
            (factNext "いま変換しているファイルが終わったところで止まります"),
            (factKept "ここまで変換した分はそのまま残ります" "次に［変換を開始］を押すと、続きから再開します")
        ) `
        -choices @(@{ Text = "中止する"; Value = "stop"; Careful = $true })
    if ($answer -ne "stop") {
        return
    }
    [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
    $ui.ConvertStopButton.IsEnabled = $false
    $ui.ConvertProgressDetail.Text = "中止しています…（変換中のファイルが終わるまでお待ちください）"
    setStatus "変換の中止を要求しました"
}

function updateConversionProgress {
    if (!(isConverting)) {
        finishConversion
        return
    }

    try {
        $progress = getConversionProgress $script:convertStart
    } catch {
        return
    }
    $stopping = !$ui.ConvertStopButton.IsEnabled

    $total = $progress.Processed + $progress.Remaining
    if (!$progress.Scanned) {
        # 変換対象を探している間（大きいフォルダ・ネットワーク越しでは数分かかることがある）。
        # 何を見ているかが分かるよう、変換側が書いた内容をそのまま出す
        $ui.ConvertProgress.IsIndeterminate = $true
        $ui.ConvertProgressText.Text = "変換対象のファイルを確認しています…"
        if (!$stopping) {
            $ui.ConvertProgressDetail.Text = [string]$progress.Detail
        }
        $taskbar.ProgressState = "Indeterminate"
        return
    }
    if ($progress.Confirming) {
        # 数え終えて、画面で変換するかどうかを選ぶのを待っている（変換側は返事があるまで止まっている）。
        # 表示は confirmConversionTargets がダイアログを開く直前に 1 回だけ変える（取りやめの表示を上書きしないため）
        $taskbar.ProgressState = "Paused"
        confirmConversionTargets
        return
    }
    if ($progress.Finishing) {
        # 後片付け（Officeアプリの終了・変換一覧の書き直し）。止まって見えないよう、何をしているかを出す
        $ui.ConvertProgress.IsIndeterminate = $true
        $ui.ConvertProgressText.Text = "変換を終えています…"
        $ui.ConvertProgressDetail.Text = [string]$progress.Detail
        $taskbar.ProgressState = "Indeterminate"
        return
    }
    if ($progress.Processed -eq 0) {
        $ui.ConvertProgress.IsIndeterminate = $true
        $ui.ConvertProgressText.Text = if ($progress.Remaining -gt 0) { "$($progress.Remaining) 件のファイルを変換します" } else { "変換が必要なファイルを確認しています…" }
        if (!$stopping) {
            $ui.ConvertProgressDetail.Text = if ($progress.Current) { "変換中のファイル：$($progress.Current)" } else { "" }
        }
        $taskbar.ProgressState = "Indeterminate"
        return
    }

    $ratio = if ($total -gt 0) { $progress.Processed / $total } else { 1 }
    $ui.ConvertProgress.IsIndeterminate = $false
    $ui.ConvertProgress.Value = $ratio
    $taskbar.ProgressState = if ($progress.Failed -gt 0) { "Paused" } else { "Normal" }
    $taskbar.ProgressValue = $ratio

    $text = "変換中… $($progress.Processed.ToString('N0')) / $($total.ToString('N0')) 件"
    if ($progress.Failed -gt 0) {
        $text += "（失敗 $($progress.Failed) 件）"
    }
    $ui.ConvertProgressText.Text = $text

    # 残り時間の目安（最初の1件が終わってからの速さで計算する）
    $now = Get-Date
    if ($null -eq $script:convertRate) {
        $script:convertRate = @{ Time = $now; Processed = $progress.Processed }
    }
    $done = $progress.Processed - $script:convertRate.Processed
    if ($progress.Remaining -eq 0) {
        $ui.ConvertProgressEta.Text = ""
    } elseif ($done -gt 0) {
        $seconds = ($now - $script:convertRate.Time).TotalSeconds / $done * $progress.Remaining
        $ui.ConvertProgressEta.Text = if ($seconds -lt 60) { "残り 1 分未満" } else { "残り約 $([math]::Ceiling($seconds / 60)) 分" }
    }
    if (!$stopping) {
        $ui.ConvertProgressDetail.Text = if ($progress.Current) { "変換中のファイル：$($progress.Current)" } else { "" }
    }
}

function finishConversion {
    $script:convertTimer.Stop()
    $taskbar.ProgressState = "None"
    # 変換完了の通知。以前はタスクバーのボタンを光らせていたが（FlashWindowEx）、P/Invoke は
    # 実行時コンパイル（csc.exe）を無くすため廃止した。完了は進捗表示・ステータスで分かる。
    $exitCode = $null
    try {
        $script:convertProcess.WaitForExit()
        $exitCode = $script:convertProcess.ExitCode
    } catch {
    }
    $progress = $null
    try {
        $progress = getConversionProgress $script:convertStart
    } catch {
    }

    $counts = ""
    if ($progress -and $progress.Processed -gt 0) {
        $counts = "成功 $($progress.Processed - $progress.Failed) 件 / 失敗 $($progress.Failed) 件"
        if ($progress.Remaining -gt 0) {
            $counts += " / 残り $($progress.Remaining) 件"
        }
        $ui.ConvertProgress.IsIndeterminate = $false
        $ui.ConvertProgress.Value = $progress.Processed / ($progress.Processed + $progress.Remaining)
    } else {
        $ui.ConvertProgress.IsIndeterminate = $false
        $ui.ConvertProgress.Value = 0
    }

    if ($exitCode -eq 1) {
        # 変換を続けられないエラー（変換対象フォルダが無い など）
        $message = (readTextShared ${convertErrorFile}).Trim()
        if ($message -eq "") {
            $message = "詳しくはログを確認してください。"
        }
        $ui.ConvertProgressText.Text = "変換できませんでした"
        $ui.ConvertProgressDetail.Text = $message
        setStatus "変換できませんでした：$message"
        showMessage "変換できませんでした。`n`n$message" "OK" "Error" | Out-Null
    } elseif ($exitCode -eq 2 -and $script:convertCanceledAtConfirm) {
        # 確認のダイアログで取りやめた（1件も変換していない）
        $ui.ConvertProgressText.Text = "変換を取りやめました"
        $ui.ConvertProgressDetail.Text = "変換したファイルはありません。［変換を開始］を押すと、もう一度確認できます。"
        setStatus $ui.ConvertProgressText.Text
    } elseif ($exitCode -eq 2) {
        $ui.ConvertProgressText.Text = if ($counts) { "変換を中止しました（$counts）" } else { "変換を中止しました" }
        $ui.ConvertProgressDetail.Text = "次回は続きから再開できます。"
        setStatus $ui.ConvertProgressText.Text
    } elseif ($progress -and $progress.Processed -gt 0) {
        $ui.ConvertProgressText.Text = "変換が終わりました（$counts）"
        $ui.ConvertProgressDetail.Text = if ($progress.Failed -gt 0) { "失敗したファイルと原因は「変換に失敗したファイル」の一覧で確認できます。" } else { "" }
        setStatus $ui.ConvertProgressText.Text
    } else {
        $ui.ConvertProgressText.Text = "変換が必要なファイルはありませんでした"
        $ui.ConvertProgressDetail.Text = ""
        setStatus $ui.ConvertProgressText.Text
    }
    $ui.ConvertProgressEta.Text = ""
    $ui.ConvertStopButton.Visibility = "Collapsed"
    $ui.ConvertLogButton.Visibility = if (Test-Path -LiteralPath ${convertLogFile}) { "Visible" } else { "Collapsed" }

    $script:convertProcess = $null
    $script:sourceFolderMaps = @{}
    refreshConversionState
    refreshIndexSummary
    loadIndexTree  # 新しいインデックス・フォルダをツリーに出す
    updateKillBadge
}

$script:convertTimer = newTimer 1000 { safe { updateConversionProgress } }

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
$ui.ConvertButton.Add_Click({ safe { startConversion } })
$ui.ConvertStopButton.Add_Click({ safe { stopConversion } })
$ui.ConvertLogButton.Add_Click({
    safe {
        if (Test-Path -LiteralPath ${convertLogFile}) {
            Invoke-Item -LiteralPath ${convertLogFile}
        }
    }
})
