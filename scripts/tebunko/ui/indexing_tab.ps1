# ［インデックス管理］タブのうち、インデックス作成の開始・中止と進み具合の表示。

# ---- インデックス作成の起動と進み具合 ----

function getIndexingProgress {
    # インデックス作成の進み具合を返す（インデクサが受け渡しの口に入れた進み具合を読む。取り込み一覧は読まない）
    $progress = @{ Scanned = $false; Processed = 0; Failed = 0; Remaining = 0; Current = ""; Detail = ""; Finishing = $false; Confirming = $false }
    $current = $null
    if ($script:indexingSession) {
        $current = readIndexingProgress $script:indexingSession.Channel
    }
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


function showIndexingPanel {
    setIndexingBanner "info" "running"
    $ui.IndexingProgressPanel.Visibility = "Visible"
    $ui.IndexingProgress.Visibility = "Visible"
    $ui.IndexingProgress.IsIndeterminate = $true
    updateIndexTabBadge
    $ui.IndexingResumeButton.Visibility = "Collapsed"
    $ui.IndexingProgressText.Text = "更新の準備をしています…"
    $ui.IndexingProgressEta.Text = ""
    $ui.IndexingProgressDetail.Text = "更新するファイルを確認しています。"
    $ui.IndexingStopButton.Visibility = "Visible"
    $ui.IndexingStopButton.IsEnabled = $true
    $ui.IndexingSearchButton.Visibility = "Collapsed"
    $taskbar.ProgressState = "Indeterminate"
}

function buildPlanRows {
    # 更新の予定を確認のダイアログの一覧に変える（文言と色の種類は indexing_view.ps1 が決める）
    param (
        $plan  # readIngestPlan の結果
    )

    $rows = New-Object System.Collections.Generic.List[PlanRow]
    # , で包んだ戻り値は、そのまま foreach に渡すと空のときも 1 回まわるため、変数に受けてから回す
    $views = newPlanViewRows $plan $script:indexingOnlyNames
    foreach ($view in $views) {
        $row = [PlanRow]::new()
        $row.Name = $view.Name
        $row.Path = $view.Path
        $row.TotalText = $view.TotalText
        $row.StatusText = $view.StatusText
        $row.Level = $view.Level
        $row.DetailText = $view.DetailText
        $rows.Add($row)
    }
    return , $rows.ToArray()
}
function updateIndexingConfirmTotal {
    # 「失敗分も更新し直す」のチェックに合わせて、合計と主ボタンの文言を変える
    $d = $script:confirmDialog
    $retry = [bool]$d.Ctrl.RetryCheck.IsChecked
    $folders = getIndexingConfirmFolderCount $d.Plan $retry
    $view = getIndexingConfirmText $d.Targets $d.Failed $retry $folders (getIndexingDroppedCount $d.Plan)
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
        if (@($script:indexingOnlyNames).Count -gt 0 -and @($script:indexingOnlyNames) -notcontains [string]$item.インデックス名) {
            continue  # 選んだものだけの回では、選ばなかった行は数えない
        }
        if ($item.区分 -eq ${planKindIngest}) {
            $targets += $item.取り込み対象
            $failed += $item.前回失敗
        }
    }

    $dialog = loadWindow "${xamlDir}\dialog_indexing_confirm.xaml" ${fontsDir}
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("StartButton", "CancelButton", "RetryCheck", "TotalText", "NoteText", "IntroText", "PlanGrid")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $script:confirmDialog = @{ Window = $dialog; Ctrl = $ctrl; Targets = $targets; Failed = $failed; Plan = $plan; Answer = $null }

    $ctrl.PlanGrid.ItemsSource = buildPlanRows $plan
    $ctrl.IntroText.Text = "元のファイルの更新日時とサイズを、前回更新したときの記録と比べました。" +
        "［更新を開始］を押すと、更新するファイルだけを更新します。"
    if ($failed -gt 0) {
        $ctrl.RetryCheck.Visibility = "Visible"
        $ctrl.RetryCheck.Content = "前回更新に失敗し、その後変わっていないファイル {0:#,0} 件も更新し直す（パスワード付きなど）" -f $failed
    }
    # 削除予定があるときは、取りやめと［更新を開始］を選べるままにする（閉じるだけにすると、見せたまま消してしまう）
    $nothing = isIndexingConfirmNothing $targets $failed (getIndexingDroppedCount $plan)
    if ($nothing) {
        # 取り込むものが無いときは、閉じるだけ（［キャンセル］との違いが無い）
        $ctrl.CancelButton.Visibility = "Collapsed"
        $ctrl.NoteText.Visibility = "Visible"
        $ctrl.NoteText.Text = "更新日時が変わらないまま中身が変わったファイルは、更新の対象になりません。" +
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
    $null = showOwnedDialog $dialog

    $answer = $script:confirmDialog.Answer
    if ($null -eq $answer -and $nothing) {
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
    $channel = $script:indexingSession.Channel
    $plan = $channel.Plan
    if ($null -eq $plan) {
        return  # まだ入っていない。次の機会に読む
    }
    $script:indexingConfirmed = $true

    $ui.IndexingProgress.IsIndeterminate = $true
    $ui.IndexingProgressText.Text = "更新する内容を確認してください"
    $ui.IndexingProgressDetail.Text = "更新するインデックスの一覧を表示しています。"
    $answer = showIndexingConfirmDialog $plan
    if ($null -eq $answer) {
        # 取りやめ。インデクサは何も取り込まずに終わる
        $script:indexingCanceledAtConfirm = $true
        $ui.IndexingStopButton.IsEnabled = $false
        $ui.IndexingProgressText.Text = "更新を取りやめています…"
        $ui.IndexingProgressDetail.Text = ""
        answerIndexingPlan $channel $null
        setStatus "更新を取りやめました"
        return
    }
    answerIndexingPlan $channel $answer
    $script:indexingRate = $null  # 残り時間の目安は、確認を待っていた時間を含めずに計る
    $ui.IndexingProgressText.Text = "更新を始めています…"
    setStatus "更新を開始しました"
}

function startIndexing {
    # インデックス作成を始める。onlyNames があれば、その名前のインデックスだけの回にする（行の［更新］・メニューの［更新］）。
    # 選んだ行は、設定の enabled を true にして取り込む（確認で取りやめても戻さない）。無ければ今までどおり、enabled の行すべて。
    # ワークスペースがネットワークの場所なら、前の版のインデックスの有無と届くかどうかを裏の列で確かめてから、続きを始める
    # （届かない場所を画面のスレッドで調べると、画面が止まる）。確かめている間は、ほかの操作を止める
    param (
        [string[]]$onlyNames = @()
    )

    if ((isIndexingOrPreparing) -or $script:targetsNaming) {
        return
    }

    # 既定のワークスペースにほかのファイルが置いてあれば、始めずに［設定］で別のフォルダを選んでもらう
    $workspaceBlock = getWorkspaceBlockMessage
    if ($workspaceBlock) {
        showMessage $workspaceBlock "OK" "Warning" | Out-Null
        selectScreen "SettingsTab"
        return
    }

    if ((getWorkspaceJobQueue $workspace.Dir) -ne "network") {
        continueStartIndexing (getLegacyIndexState $workspace.Dir) $onlyNames
        return
    }

    $script:indexingPreparing = $true
    updateIndexingButton
    setStatus (getWorkspaceCheckingStatus)
    $finish = ${function:finishWorkspaceCheck}   # 終わったときの処理は、関数を変数に取って呼ぶ（クロージャからは関数の名前を引けないため）
    startJob {
        param ($dir)
        # 届かない場所は、前の版のインデックスの有無を調べずに知らせる
        if ((getPathState $dir).State -eq ${pathStateUnreachable}) {
            return @{ Unreachable = $true; LegacyState = $null }
        }
        @{ Unreachable = $false; LegacyState = (getLegacyIndexState $dir) }
    } @($workspace.Dir) {
        param ($output, $errorText)
        & $finish $output $errorText $onlyNames
    }.GetNewClosure() (getWorkspaceJobQueue $workspace.Dir)
}

function finishWorkspaceCheck {
    # ネットワークのワークスペースの確かめが終わったときの処理。届いて調べられたら、インデックス作成の続きを始める
    param (
        $output,
        [string]$errorText,
        [string[]]$onlyNames
    )

    $script:indexingPreparing = $false
    updateIndexingButton
    $result = @($output)[0]
    if ($errorText -or $null -eq $result -or $result.Unreachable) {
        setStatus (getWorkspaceUnreachableStatus ([string]$workspace.Dir))
        return
    }
    continueStartIndexing $result.LegacyState $onlyNames
}

function continueStartIndexing {
    # startIndexing の続き。前の版のインデックスがあれば確かめ、インデクサを始める
    param (
        $legacyState,        # getLegacyIndexState の結果
        [string[]]$onlyNames
    )

    # 前の版のインデックス（index\）があり、まだ取り込み直していなければ、始める前に確かめる
    $reingestConfirm = getReingestConfirm $legacyState.HasLegacyIndex $legacyState.ContentEmpty
    if ($reingestConfirm) {
        $answer = showConfirm -heading $reingestConfirm -choices @(@{ Text = "更新し直す"; Value = "start" })
        if ($answer -ne "start") {
            return
        }
    }

    $script:indexingOnlyNames = @($onlyNames | Where-Object { $_ })
    foreach ($item in $script:targetItems) {
        if ($script:indexingOnlyNames -contains $item.Name) {
            $item.SetEnabled($true)
        }
    }
    saveTargets
    updateIndexDetailPanel
    # 何件取り込むかは、元のファイルの更新日時とサイズを見ないと分からない。
    # ConfirmTargets にすると、インデクサは数え終えたところで止まって確認の返事（answerIndexingPlan）を待つ。
    # インデクサは画面のプロセスのスレッドで動く（docs/design/structure/threads.md「プロセス」）
    $script:indexingStart = Get-Date
    $script:indexingRate = $null
    $script:indexingEta = ""
    $script:indexingConfirmed = $false
    $script:indexingCanceledAtConfirm = $false
    # 空なら今までどおり（状態層は enabled かつ onlyNames に入るものだけを取り込む）
    $channel = newIndexerChannel -confirmTargets $true -onlyNames $script:indexingOnlyNames
    $script:indexingSession = newIndexingSession $channel

    showIndexingPanel
    setStatus "更新するファイルを調べています…"
    updateIndexingButton
    $script:indexingTimer.Start()
}

function stopIndexing {
    if (!(isIndexing)) {
        return
    }
    $answer = showConfirm `
        -title "更新の中止" `
        -heading "インデックスの更新を中止しますか？" `
        -hint "更新したところまでは残ります。あとで続きから再開できます。" `
        -choices @(@{ Text = "中止する"; Value = "stop"; Careful = $true })
    if ($answer -ne "stop") {
        return
    }
    requestIndexingStop $script:indexingSession.Channel
    $ui.IndexingStopButton.IsEnabled = $false
    $ui.IndexingProgressDetail.Text = "中止しています…（更新中のファイルが終わるまでお待ちください）"
    setStatus "更新の中止を要求しました"
}

function updateIndexingProgress {
    if (!(isIndexing)) {
        finishIndexing
        return
    }

    try {
        $progress = getIndexingProgress
    } catch {
        return
    }
    $stopping = !$ui.IndexingStopButton.IsEnabled

    $total = $progress.Processed + $progress.Remaining
    if (!$progress.Scanned) {
        # クロールしている間（大きいフォルダ・ネットワーク越しでは数分かかることがある）。
        # 何を見ているかが分かるよう、インデクサが書いた内容をそのまま出す
        $ui.IndexingProgress.IsIndeterminate = $true
        $ui.IndexingProgressText.Text = "更新するファイルを調べています…"
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
        $ui.IndexingProgressText.Text = "更新を終えています…"
        $ui.IndexingProgressDetail.Text = [string]$progress.Detail
        $taskbar.ProgressState = "Indeterminate"
        return
    }
    if ($progress.Processed -eq 0) {
        $ui.IndexingProgress.IsIndeterminate = $true
        $ui.IndexingProgressText.Text = if ($progress.Remaining -gt 0) { "$($progress.Remaining) 件のファイルを更新します" } else { "更新が必要なファイルを確認しています…" }
        if (!$stopping) {
            $ui.IndexingProgressDetail.Text = if ($progress.Current) { "更新中のファイル：$($progress.Current)" } else { "" }
        }
        $taskbar.ProgressState = "Indeterminate"
        return
    }

    $ratio = if ($total -gt 0) { $progress.Processed / $total } else { 1 }
    $ui.IndexingProgress.IsIndeterminate = $false
    $ui.IndexingProgress.Value = $ratio
    $taskbar.ProgressState = if ($progress.Failed -gt 0) { "Paused" } else { "Normal" }
    $taskbar.ProgressValue = $ratio
    updateIndexTabBadge
    updateIndexRowsProgress

    $text = "インデックスを更新しています（$($progress.Processed.ToString('N0')) / $($total.ToString('N0')) 件"
    if ($progress.Failed -gt 0) {
        $text += "・失敗 $($progress.Failed) 件"
    }
    $ui.IndexingProgressText.Text = $text + "）"

    # 残り時間の目安（最初の1件が終わってからの速さで計算する）
    $now = Get-Date
    if ($null -eq $script:indexingRate) {
        $script:indexingRate = @{ Time = $now; Processed = $progress.Processed }
    }
    $done = $progress.Processed - $script:indexingRate.Processed
    if ($progress.Remaining -eq 0) {
        $script:indexingEta = ""
    } elseif ($done -gt 0) {
        $seconds = ($now - $script:indexingRate.Time).TotalSeconds / $done * $progress.Remaining
        $script:indexingEta = if ($seconds -lt 60) { "残り 1 分未満" } else { "残り約 $([math]::Ceiling($seconds / 60)) 分" }
    }
    $ui.IndexingProgressEta.Text = $script:indexingEta
    if (!$stopping) {
        $ui.IndexingProgressDetail.Text = if ($progress.Current) { "更新中のファイル：$($progress.Current)" } else { "" }
    }
    # 更新中の名前・件数・残り時間は、ここに値で 1 か所だけ持つ（帯・詳細の「インデックス」の箱・ステータスバーが読む）。
    # Name は取り込み中のファイル（<名前>\<相対パス>）のインデックス名。件数は回全体（インデックスごとの件数は無い）
    $script:indexingView = @{
        Name = (getIndexingCurrentName ([string]$progress.Current))
        Ratio = $ratio; Processed = $progress.Processed; Total = $total; Failed = $progress.Failed
        Eta = [string]$script:indexingEta; Current = [string]$progress.Current
    }
    updateIndexDetailPanel
}

function finishIndexing {
    $script:indexingTimer.Stop()
    $script:indexingView = $null
    $taskbar.ProgressState = "None"
    # インデックス作成完了の通知。以前はタスクバーのボタンを光らせていたが（FlashWindowEx）、P/Invoke は
    # 実行時コンパイル（csc.exe）を無くすため廃止した。完了は進捗表示・ステータスで分かる。
    $session = $script:indexingSession
    $progress = $null
    try {
        $progress = getIndexingProgress
    } catch {
    }
    $exitCode = $session.GetExitCode()
    $errorText = $session.GetError()
    $onlySkipped = @($session.Channel.OnlySkipped)
    # インデクサのスレッドを片づける（終わっているため待たない）
    $session.Close()
    $script:indexingSession = $null
    $script:indexingOnlyNames = @()
    $script:indexingEta = ""

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
        $message = $errorText.Trim()
        if ($message -eq "") {
            $message = "詳しくはログを確認してください。"
        }
        setIndexingBanner (getIndexingBannerLevel $exitCode 0) "done"
        $ui.IndexingProgressText.Text = "インデックスを更新できませんでした"
        $ui.IndexingProgressDetail.Text = $message
        setStatus "インデックスを更新できませんでした：$message"
        showMessage "インデックスを更新できませんでした。`n`n$message" "OK" "Error" | Out-Null
    } elseif ($exitCode -eq 2 -and $script:indexingCanceledAtConfirm) {
        # 確認のダイアログで取りやめた（1件も取り込んでいない）
        setIndexingBanner "warn" "done"
        $ui.IndexingProgressText.Text = "更新を取りやめました"
        $ui.IndexingProgressDetail.Text = "更新したファイルはありません。［すべて更新］を押すと、もう一度確認できます。"
        setStatus $ui.IndexingProgressText.Text
    } elseif ($exitCode -eq 2) {
        setIndexingBanner "warn" "done"
        $ui.IndexingProgressText.Text = if ($counts) { "更新を中止しました（$counts）" } else { "更新を中止しました" }
        $ui.IndexingProgressDetail.Text = "次回は続きから再開できます。"
        setStatus $ui.IndexingProgressText.Text
    } else {
        # 完了（後回し・失敗の件数に応じた見出しと説明は判断層（indexing_view.ps1）が決める）
        $success = if ($progress) { $progress.Processed - $progress.Failed } else { 0 }
        $failed = if ($progress) { $progress.Failed } else { 0 }
        setIndexingBanner (getIndexingBannerLevel $exitCode $failed) "done"
        $endText = getIndexingEndText $success $failed $session.GetPostponed() $session.GetNotice()
        $ui.IndexingProgressText.Text = $endText.Text
        $ui.IndexingProgressDetail.Text = $endText.Detail
        setStatus $ui.IndexingProgressText.Text
    }
    $ui.IndexingProgressEta.Text = ""
    $ui.IndexingStopButton.Visibility = "Collapsed"
    $ui.IndexingProgress.Visibility = "Collapsed"
    $skippedView = getIndexingSkippedView $onlySkipped
    if ($skippedView -and $exitCode -ne 1) {
        showMessage "$($skippedView.Heading)`n`n$($skippedView.Detail)" "OK" "Warning" | Out-Null
    }
    $ui.IndexingSearchButton.Visibility = if ($exitCode -ne 1) { "Visible" } else { "Collapsed" }

    $script:sourceFolderMaps = @{}
    # 高速検索の列の前の確かめ結果（古い reason）を捨てて「確認中…」に戻す。ここで捨てずに
    # refreshIndexingState を呼ぶと、その集計（本文の取り込みは済んだと分かる）が
    # refreshFastSearchStatus の確かめ直しより先に終わったとき、古い reason のまま
    # updateFastSearchRows が呼ばれ、「不可」が一瞬出てしまう（isIndexing は既に偽になっており、
    # getFastSearchRowView の indexing 引数による作成中ガードが効かないため）。
    # fastSearchResultDir を null にするだけで、updateFastSearchRows 側の「今のワークスペースの結果
    # でなければ捨てる」ガード（index\index_list.ps1・index_detail.ps1）が reason・progress・checkedAt を確認中…に戻してくれる。
    # 世代番号も進め、作成前から走っていた確かめジョブが後から古い結果を届けても捨てて確かめ直すようにする
    $script:fastSearchGeneration++
    $script:fastSearchResultDir = $null
    refreshIndexingState
    refreshIndexSummary
    loadIndexTree  # 新しいインデックス・フォルダをツリーに出す
    refreshFastSearchStatus  # インデックス作成が終わったので、一覧の「高速検索」列を確かめ直す
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
$ui.IndexingResumeButton.Add_Click({ safe { startIndexing } })
$ui.IndexingSearchButton.Add_Click({ safe { selectScreen "SearchTab" } })
