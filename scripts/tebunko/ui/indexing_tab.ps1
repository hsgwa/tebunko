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
        $plan,  # readIngestPlan の結果
        [hashtable]$textExt = @{}  # 取り込む拡張子の欄の材料（newPlanViewRows の textExt）
    )

    $rows = New-Object System.Collections.Generic.List[object]
    # , で包んだ戻り値は、そのまま foreach に渡すと空のときも 1 回まわるため、変数に受けてから回す
    $subfolders = @{}
    foreach ($item in $script:targetItems) {
        if ($item.Name) { $subfolders[[string]$item.Name] = [bool]$item.Subfolders }
    }
    $views = newPlanViewRows $plan $script:indexingOnlyNames $subfolders $textExt
    foreach ($view in $views) {
        # 表の列（拡張子の要約と［変更…］）まで束縛するため、型を決めた PlanRow でなく、列と同じ名前の項目を持つ行にする
        $rows.Add([pscustomobject]@{
                Name = [string]$view.Name; Path = [string]$view.Path; TotalText = [string]$view.TotalText
                StatusText = [string]$view.StatusText; Level = [string]$view.Level; DetailText = [string]$view.DetailText
                TopOnlyText = [string]$view.TopOnlyText
                TextExtShown = [bool]$view.TextExtShown; TextExtSummary = [string]$view.TextExtSummary; TextExtTip = [string]$view.TextExtTip
            })
    }
    return , $rows.ToArray()
}
function updateIndexingConfirmTotal {
    # チェックの値に合わせて、合計・主ボタン・チェックボックスの表示を変える（文言と可否は getIndexingConfirmView が決める）
    $d = $script:confirmDialog
    $view = getIndexingConfirmView $d.Plan @($script:indexingOnlyNames) ([bool]$d.Ctrl.RetryCheck.IsChecked) ([bool]$d.Ctrl.CloudCheck.IsChecked)
    $d.View = $view
    $d.Ctrl.TotalText.Text = $view.Confirm.Text
    # インポートの回は合計の 1 行を出さない（取り込むものが 0 件のとき「すべて最新です」が、行の『要更新』・主ボタンと食い違うため）
    $d.Ctrl.TotalText.Visibility = $(if ($script:indexingConfirmText) { "Collapsed" } else { "Visible" })
    $d.Ctrl.SubfolderRemovalNote.Text = $view.SubfolderRemovalNote
    $d.Ctrl.SubfolderRemovalNote.Visibility = $(if ($view.SubfolderRemovalNote) { "Visible" } else { "Collapsed" })
    $d.Ctrl.TextExcludeRemovalNote.Text = $view.TextExcludeRemovalNote
    $d.Ctrl.TextExcludeRemovalNote.Visibility = $(if ($view.TextExcludeRemovalNote) { "Visible" } else { "Collapsed" })
    # インポートの回は、主ボタンの文言を差し替える（チェックを押すたびに書き直すので、ここで当てる）
    $d.Ctrl.StartButton.Content = $(if ($script:indexingConfirmText) { $script:indexingConfirmText.StartText } else { $view.Confirm.Button })
    $d.Ctrl.RetryCheck.Visibility = $(if ($view.RetryVisible) { "Visible" } else { "Collapsed" })
    $d.Ctrl.RetryCheck.Content = $view.RetryText
    $cloudVisibility = $(if ($view.CloudVisible) { "Visible" } else { "Collapsed" })
    $d.Ctrl.CloudWarnText.Visibility = $cloudVisibility
    $d.Ctrl.CloudCheck.Visibility = $cloudVisibility
    $d.Ctrl.CloudRememberCheck.Visibility = $cloudVisibility
    $d.Ctrl.CloudWarnText.Text = $view.CloudWarn
    $d.Ctrl.CloudCheck.Content = $view.CloudCheckText
    # 「次から確認を出さない」は、ダウンロードして取り込むを選んだときだけ選べる
    $d.Ctrl.CloudRememberCheck.IsEnabled = $view.Include
}
function initIndexingConfirmLocation {
    # 確認のダイアログに、インデックスの保存先と注意（共有のときは警告も）を出し、［フォルダを開く］をつなぐ。タブと同じ文言（getIndexLocationView）。
    # ［フォルダを開く］の結果の文言は、ダイアログの中に出す（モーダルの間は本体のステータスバーが見えにくいため）
    param (
        $ctrl  # ダイアログの部品（IndexLocation… と OpenIndexLocationButton）
    )

    $dir = [string]$workspace.Dir
    $location = getIndexLocationView $dir (testNetworkPath $dir)
    $ctrl.IndexLocationText.Text = $location.Path
    $ctrl.IndexLocationText.ToolTip = $location.Path
    $ctrl.IndexLocationNote.Text = $location.Note
    $ctrl.IndexLocationWarning.Text = $location.Warning
    $ctrl.IndexLocationWarning.Visibility = if ($location.Warning) { "Visible" } else { "Collapsed" }
    $statusText = $ctrl.IndexLocationStatus
    $showMessage = {
        param ($text)
        $statusText.Text = $text
        $statusText.Visibility = if ($text) { "Visible" } else { "Collapsed" }
    }.GetNewClosure()
    # GetNewClosure したブロックの中では関数を名前で探せないことがあるため、関数は先に変数へ入れて & で呼ぶ
    $open = ${function:openIndexLocation}
    $safeCall = ${function:safe}
    $openBlock = { & $open $showMessage }.GetNewClosure()
    $ctrl.OpenIndexLocationButton.Add_Click({ & $safeCall $openBlock }.GetNewClosure())
}

function changeIndexingTextExtensions {
    # 確認のダイアログの［変更…］。拡張子を選ぶダイアログを開き、［OK］で選びが変わっていたら、設定にすぐ書き、
    # 同じ対象で数え直せるよう始め直しの印を立てて、この確認を取りやめとして閉じる（閉じたあとの始め直しは restartIndexingIfRequested）。
    # ［キャンセル］・選びが変わらない［OK］では何もしない。ダイアログの中で拡張子をいくつ替えても、数え直しは［OK］の 1 回だけ
    param (
        [string]$name  # インデックス名
    )

    $d = $script:confirmDialog
    $current = @($d.TextExcludes[$name] | Where-Object { $_ })
    $planRow = @($d.Plan | Where-Object { $null -ne $_ -and $_.区分 -eq ${planKindIngest} -and [string]$_.インデックス名 -eq $name })[0]
    $counts = $(if ($planRow -and $null -ne $planRow.TextExtensionCounts) { $planRow.TextExtensionCounts } else { @{} })
    $selected = showIndexTextExtensionsDialog -Owner $d.Window -IndexName $name -Excluded $current -Counts $counts -CountsSource "now" -CountsTime ""
    if ($null -ne $selected) {
        $selected = @(addOutOfListExtensions $selected $current @(getIndexSelectableExtensions))
    }
    if ($null -eq $selected -or -not (testIndexingTextExcludesChanged $current @($selected))) {
        return
    }
    try {
        writeIndexTextExcludes $name @($selected)
    } catch {
        showMessage "設定を保存できませんでした。もう一度お試しください。" "OK" "Warning" | Out-Null
        return
    }
    # インポートの後の確認なら、始め直しも同じ文言で開く（確認が 2 回にならず、文言も変わらない）
    $script:indexingRestart = @{ OnlyNames = @($script:indexingOnlyNames); ConfirmText = $script:indexingConfirmText }
    $d.Window.DialogResult = $false
}

function initIndexingConfirmTextExt {
    # 確認のダイアログの「取り込む拡張子」の欄をつなぐ。初めて（どの行も拡張子を外していない）ときの案内の一文を出し、表の［変更…］を受ける
    param (
        $d  # $script:confirmDialog
    )

    $rows = @($d.Ctrl.PlanGrid.ItemsSource | Where-Object { $_.TextExtShown })
    $anyExcluded = (@($rows | Where-Object { @($d.TextExcludes[[string]$_.Name] | Where-Object { $_ }).Count -gt 0 }).Count -gt 0)
    $d.Ctrl.TextExtGuideText.Text = getIndexingTextExtGuide $anyExcluded ($rows.Count -gt 0)
    $d.Ctrl.TextExtGuideText.Visibility = $(if ($d.Ctrl.TextExtGuideText.Text) { "Visible" } else { "Collapsed" })
    # 表の中のボタンは、表で Click を受ける（行ごとに作られるため）
    $d.Ctrl.PlanGrid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
            param ($sender, $e)
            $button = $e.OriginalSource
            if ($button -is [System.Windows.Controls.Button] -and $button.Name -eq "ChangeIndexTextExtensionsButton") {
                $name = [string]$button.Tag
                safe { changeIndexingTextExtensions $name }
            }
        })
}

function applyImportConfirmText {
    # インポートの後の確認: 窓の題・見出し・見出しの下の説明・［あとで］・取り込んだ行のステータスを差し替える（文言は getImportIndexingConfirm が決める）。文言が無ければ何もしない
    param (
        $dialog,
        [hashtable]$ctrl,
        $rows,
        $importText
    )

    if (!$importText) {
        return
    }
    $dialog.Title = $importText.Title
    $dialog.FindName("HeadingText").Text = $importText.Heading
    $hint = $dialog.FindName("ImportIndexHint")
    $hint.Text = $importText.Hint
    $hint.Visibility = "Visible"
    $ctrl.CancelButton.Content = $importText.CancelText
    foreach ($row in @($rows)) {
        if ([string]$row.Name -ieq $importText.Name) {
            $row.StatusText = $importText.RowStatus
            $row.Level = "Wait"
        }
    }
}

function showIndexingConfirmDialog {
    # インデックス作成の確認。インデクサが数えた結果（インデックスごとの取り込み対象の件数）を出して、取り込むかどうかを選んでもらう。
    #   取り込む → @{ RetryFailed; IncludeCloud } ／ 取りやめ → $null
    param (
        $plan  # readIngestPlan の結果
    )

    $dialog = loadWindow "${xamlDir}\dialog_indexing_confirm.xaml"
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("StartButton", "CancelButton", "RetryCheck", "CloudWarnText", "CloudCheck", "CloudRememberCheck", "TotalText", "SubfolderRemovalNote", "TextExcludeRemovalNote", "TextExtGuideText", "NoteText", "IntroText", "PlanGrid", "IndexLocationText", "IndexLocationNote", "IndexLocationWarning", "IndexLocationStatus", "OpenIndexLocationButton")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    # 設定の cloudFiles は開くときに 1 回だけ読む（開始するとき、変わったときだけ書き戻す）
    $savedCloud = readCloudFiles
    # 取り込む拡張子の設定も開くときに 1 回だけ読む（要約と、［変更…］で選びが変わったかの比べに使う）
    $textExcludes = (readTextExtensionSettings).IndexExcludes
    $script:confirmDialog = @{ Window = $dialog; Ctrl = $ctrl; Plan = $plan; SavedCloud = $savedCloud; TextExcludes = $textExcludes; View = $null; Answer = $null }

    $rows = buildPlanRows $plan @{ Excludes = $textExcludes; Choices = @(getIndexSelectableExtensions); Executable = @(${textNotepadExtensions}).Count }
    $importText = $script:indexingConfirmText
    applyImportConfirmText $dialog $ctrl $rows $importText
    $ctrl.PlanGrid.ItemsSource = $rows
    $ctrl.IntroText.Text = getIndexingConfirmIntro (getIndexingDroppedCount $plan)
    $ctrl.IntroText.Visibility = if ($ctrl.IntroText.Text) { "Visible" } else { "Collapsed" }
    $ctrl.CloudRememberCheck.Content = ${indexingCloudRememberText}
    if ($savedCloud -eq ${cloudFilesDownload}) {
        # 前回「次から確認を出さずにダウンロードする」を選んだ。チェックを付けておく（欄が出るときに効く）
        $ctrl.CloudCheck.IsChecked = $true
        $ctrl.CloudRememberCheck.IsChecked = $true
    }
    initIndexingConfirmLocation $ctrl
    initIndexingConfirmTextExt $script:confirmDialog
    updateIndexingConfirmTotal
    # 削除予定があるときは、取りやめと［更新を開始］を選べるままにする（閉じるだけにすると、見せたまま消してしまう）
    # インポートの回は、取り込むものが無くても［あとで］を選べる（選べないと 1 回の確認にならない）ので、Nothing として扱わない
    $nothing = [bool]$script:confirmDialog.View.Nothing -and -not $importText
    if ($nothing) {
        # 取り込むものが無いときは、閉じるだけ（［キャンセル］との違いが無い）
        $ctrl.CancelButton.Visibility = "Collapsed"
        $ctrl.NoteText.Visibility = "Visible"
        $ctrl.NoteText.Text = "更新日時が変わらないまま中身が変わったファイルは、更新の対象になりません。" +
            "そのインデックスを一から作り直すときは、［削除］してから追加し直してください。"
    }

    $ctrl.RetryCheck.Add_Click({ safe { updateIndexingConfirmTotal } })
    $ctrl.CloudCheck.Add_Click({ safe { updateIndexingConfirmTotal } })
    $ctrl.StartButton.Add_Click({
        safe {
            $d = $script:confirmDialog
            updateIndexingConfirmTotal
            $view = $d.View
            if (testIndexingCloudNeedsConfirm $view.Include $d.SavedCloud) {
                # ダウンロードして取り込む前に、もう一度確かめる。キャンセルなら、この確認に戻る（閉じない）
                $text = getIndexingCloudConfirmView $view.Cloud.Count $view.Cloud.Bytes ${ingestFileTimeoutMinutes}
                $facts = @($text.Facts | ForEach-Object { factWarn $_ })
                $choice = showConfirm -heading $text.Heading -choices $text.Choices -facts $facts -cancelText $text.CancelText -owner $d.Window
                if ($null -eq $choice) {
                    return
                }
            }
            # 開始するときの選び方を、変わったときだけ設定に残す（次から確認を出さずにダウンロードするか）
            $mode = getIndexingCloudModeToSave $view.CloudVisible $view.Include ([bool]$d.Ctrl.CloudRememberCheck.IsChecked) $d.SavedCloud
            if ($null -ne $mode) {
                writeCloudFiles $mode
            }
            $d.Answer = @{ RetryFailed = $view.Retry; IncludeCloud = $view.Include }
            $d.Window.DialogResult = $true
        }
    })
    $null = showOwnedDialog $dialog

    $answer = $script:confirmDialog.Answer
    if ($null -eq $answer -and $nothing -and -not $script:indexingRestart) {
        # 取り込むものが無いときは、どう閉じても同じ（インデクサはそのまま終わる）。拡張子を変えて始め直すときは、取りやめにする
        $answer = @{ RetryFailed = $false; IncludeCloud = $false }
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
    markTrace "index" "確認の画面を閉じる"
    if ($null -eq $answer) {
        # 取りやめ。インデクサは何も取り込まずに終わる
        $script:indexingCanceledAtConfirm = $true
        $ui.IndexingStopButton.IsEnabled = $false
        $ui.IndexingProgressText.Text = "更新を取りやめています…"
        $ui.IndexingProgressDetail.Text = ""
        answerIndexingPlan $channel $null
        return
    }
    answerIndexingPlan $channel $answer
    # 更新を始めることは新しい操作なので、前の操作の結果の文は消す（取りやめたときは消さない）
    setStatus ""
    $script:indexingRate = $null  # 残り時間の目安は、確認を待っていた時間を含めずに計る
    $ui.IndexingProgressText.Text = "更新を始めています…"
}

function startIndexing {
    # インデックス作成を始める。onlyNames があれば、その名前のインデックスだけの回にする（行の［更新］・メニューの［更新］）。
    # 選んだ行は、設定の enabled を true にして取り込む（確認で取りやめても戻さない）。無ければ今までどおり、enabled の行すべて。
    # ワークスペースがネットワークの場所なら、前の版のインデックスの有無と届くかどうかを裏の列で確かめてから、続きを始める
    # （届かない場所を画面のスレッドで調べると、画面が止まる）。確かめている間は、ほかの操作を止める
    param (
        [string[]]$onlyNames = @(),
        $confirmText = $null   # インポートの後の更新だけが渡す確認の文言（getImportIndexingConfirm の結果）。省けばふだんの確認
    )

    if ((isIndexingOrPreparing) -or $script:targetsNaming) {
        return
    }
    markTrace "index" "開始の押下"

    # 既定のワークスペースにほかのファイルが置いてあれば、始めずに［設定］で別のフォルダを選んでもらう
    $workspaceBlock = getWorkspaceBlockMessage
    if ($workspaceBlock) {
        showMessage $workspaceBlock "OK" "Warning" | Out-Null
        selectScreen "SettingsTab"
        return
    }

    if ((getWorkspaceJobQueue $workspace.Dir) -ne "network") {
        continueStartIndexing (getLegacyIndexState $workspace.Dir) $onlyNames $confirmText
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
        & $finish $output $errorText $onlyNames $confirmText
    }.GetNewClosure() (getWorkspaceJobQueue $workspace.Dir)
}

function finishWorkspaceCheck {
    # ネットワークのワークスペースの確かめが終わったときの処理。届いて調べられたら、インデックス作成の続きを始める
    param (
        $output,
        [string]$errorText,
        [string[]]$onlyNames,
        $confirmText = $null
    )

    $script:indexingPreparing = $false
    updateIndexingButton
    $result = @($output)[0]
    if ($errorText -or $null -eq $result -or $result.Unreachable) {
        setStatus (getWorkspaceUnreachableStatus ([string]$workspace.Dir))
        return
    }
    # 確かめ中の文を消す（以後の確認を断っても、続きを始めても、残さない）
    setStatus ""
    continueStartIndexing $result.LegacyState $onlyNames $confirmText
}

function continueStartIndexing {
    # startIndexing の続き。前の版のインデックスがあれば確かめ、インデクサを始める
    param (
        $legacyState,        # getLegacyIndexState の結果
        [string[]]$onlyNames,
        $confirmText = $null  # startIndexing の confirmText
    )

    # 前の版のインデックス（index\）があり、まだ取り込み直していなければ、始める前に確かめる
    $reingestConfirm = getReingestConfirm $legacyState.HasLegacyIndex $legacyState.ContentEmpty
    if ($reingestConfirm) {
        $answer = showConfirm -heading $reingestConfirm -choices @(@{ Text = "更新し直す"; Value = "start" })
        if ($answer -ne "start") {
            return
        }
    }

    $script:indexingConfirmText = $confirmText
    $script:indexingOnlyNames = @($onlyNames | Where-Object { $_ })
    foreach ($item in $script:targetItems) {
        if ($script:indexingOnlyNames -contains $item.Name) {
            $item.SetEnabled($true)
        }
    }
    saveTargets -Quiet   # 文は出さない（確認で取りやめたとき、前の操作の結果の代わりに残らないように）
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
    $script:indexTraceProgress = $false
    markTrace "index" "インデクサのプロセスを起動"

    showIndexingPanel
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
    $ui.IndexingProgressDetail.Text = "中止しています…（取り込みを止めて、取り込めた分をまとめています）"
    showIndexingWait "Stopping"
}

# 中止を待つ間のダイアログ（shared\ui\shell.ps1 の showWaitDialog）。Stopping（［中止］のあと）と CloseWait（閉じるときの待ち）の 1 つだけを出す
$script:indexingWait = $null
$script:indexingWaitKind = ""

# 確認で取り込む拡張子を変えたときの始め直し: 印（@{ OnlyNames }。取りやめた回が終わったあとに 1 回だけ使って消す）と、始め直した数えの間か
$script:indexingRestart = $null
$script:indexingRecounting = $false

function showIndexingWait {
    param (
        [string]$kind
    )

    if ($null -ne $script:indexingWait -and $script:indexingWaitKind -eq $kind) {
        return
    }
    closeIndexingWait
    $view = getWaitDialogView $kind
    $script:indexingWait = showWaitDialog -heading $view.Heading -description $view.Description -buttonText $view.ButtonText
    $script:indexingWaitKind = $kind
}

function updateIndexingWait {
    # 進み具合の棒を書き換える。secondsLeft（強制終了までの残りの秒）を渡したときは、秒読みの 1 行も書き換える
    param (
        $progress,
        [object]$secondsLeft = $null
    )

    if ($null -eq $script:indexingWait) {
        return
    }
    $content = @{}
    if ($null -ne $progress) {
        $bar = getWaitProgressValue $progress.Processed ($progress.Processed + $progress.Remaining)
        $content.Indeterminate = $bar.Indeterminate
        $content.Value = $bar.Value
    }
    if ($null -ne $secondsLeft) {
        $content.Countdown = getForceCloseCountdownText ([double]$secondsLeft)
    }
    updateWaitDialog $script:indexingWait $content
}

function closeIndexingWait {
    $handle = $script:indexingWait
    $script:indexingWait = $null
    $script:indexingWaitKind = ""
    closeWaitDialog $handle
}

function updateIndexingEta {
    # 残り時間の目安（最初の1件が終わってからの速さで計算する）を求めて、帯に出す
    param (
        $progress
    )

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
    if (-not $script:indexTraceProgress) {
        $script:indexTraceProgress = $true
        markTrace "index" "最初の進み具合"
    }

    $total = $progress.Processed + $progress.Remaining
    updateIndexingWait $progress
    if (!$progress.Scanned) {
        # クロールしている間（大きいフォルダ・ネットワーク越しでは数分かかることがある）。
        # 何を見ているかが分かるよう、インデクサが書いた内容をそのまま出す
        $ui.IndexingProgress.IsIndeterminate = $true
        $ui.IndexingProgressText.Text = getIndexingScanText
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
            $ui.IndexingProgressDetail.Text = if ($progress.Current) { "読み取り中のファイル：$($progress.Current)" } else { "" }
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

    updateIndexingEta $progress
    if (!$stopping) {
        $ui.IndexingProgressDetail.Text = if ($progress.Current) { "読み取り中のファイル：$($progress.Current)" } else { "" }
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

function finishImportConfirm {
    # インポートの後の確認の後始末。確認で［あとで］を選んだ（取りやめ）なら、取りやめの帯を閉じ、ステータスバーで知らせる。どの終わり方でも文言を消し、次の回に持ち越さない
    param (
        [int]$exitCode
    )

    $importText = $script:indexingConfirmText
    $script:indexingConfirmText = $null
    # 拡張子を変えた始め直し（indexingRestart）は、取りやめではない。同じ文言で開き直すので、取りやめの文は出さない
    if ($importText -and $exitCode -eq 2 -and $script:indexingCanceledAtConfirm -and $null -eq $script:indexingRestart) {
        closeIndexingBanner
        setStatus $importText.SkippedStatus
    }
}

function showIndexingFinishCounts {
    # 終わったときの進み具合の棒を決め、件数の文（成功・失敗・残り）を返す。1 件も処理していなければ空
    param (
        $progress
    )

    $ui.IndexingProgress.IsIndeterminate = $false
    if (-not ($progress -and $progress.Processed -gt 0)) {
        $ui.IndexingProgress.Value = 0
        return ""
    }
    $counts = "成功 $($progress.Processed - $progress.Failed) 件 / 失敗 $($progress.Failed) 件"
    if ($progress.Remaining -gt 0) {
        $counts += " / 残り $($progress.Remaining) 件"
    }
    $ui.IndexingProgress.Value = $progress.Processed / ($progress.Processed + $progress.Remaining)
    return $counts
}

function finishIndexing {
    markTrace "index" "終わりを受ける"
    $script:indexingTimer.Stop()
    closeIndexingWait
    $script:indexingView = $null
    $taskbar.ProgressState = "None"
    # 完了の通知は帯で分かる（以前のタスクバーのボタンを光らせる FlashWindowEx は、実行時コンパイルを無くすため廃止した）。
    $session = $script:indexingSession
    $progress = $null
    try {
        $progress = getIndexingProgress
    } catch {
    }
    $exitCode = $session.GetExitCode()
    $errorText = $session.GetError()
    $onlySkipped = @($session.Channel.OnlySkipped); if ($exitCode -eq 0) { rememberTextExtCounts $session.Channel.Plan }
    # インデクサのスレッドを片づける（終わっているため待たない）
    $session.Close()
    $script:indexingSession = $null
    $script:indexingOnlyNames = @()
    $script:indexingEta = ""

    $counts = showIndexingFinishCounts $progress

    if ($exitCode -eq 1) {
        # インデックス作成を続けられないエラー（クロール対象フォルダが無い など）
        $message = $errorText.Trim()
        if ($message -eq "") {
            $message = "詳しくはログを確認してください。"
        }
        setIndexingBanner (getIndexingBannerLevel $exitCode 0) "done"
        $ui.IndexingProgressText.Text = "インデックスを更新できませんでした"
        $ui.IndexingProgressDetail.Text = $message
        showMessage "インデックスを更新できませんでした。`n`n$message" "OK" "Error" | Out-Null
    } elseif ($exitCode -eq 2 -and $script:indexingCanceledAtConfirm) {
        # 確認のダイアログで取りやめた（1件も取り込んでいない）
        setIndexingBanner "warn" "done"
        $ui.IndexingProgressText.Text = "更新を取りやめました"
        $ui.IndexingProgressDetail.Text = "更新したファイルはありません。［すべて更新］を押すと、もう一度確認できます。"
    } elseif ($exitCode -eq 2) {
        setIndexingBanner "warn" "done"
        $ui.IndexingProgressText.Text = if ($counts) { "更新を中止しました（$counts）" } else { "更新を中止しました" }
        $ui.IndexingProgressDetail.Text = "次回は続きから再開できます。"
    } else {
        # 完了（後回し・失敗の件数に応じた見出しと説明は判断層（indexing_view.ps1）が決める）
        $success = if ($progress) { $progress.Processed - $progress.Failed } else { 0 }
        $failed = if ($progress) { $progress.Failed } else { 0 }
        setIndexingBanner (getIndexingBannerLevel $exitCode $failed) "done"
        $endText = getIndexingEndText $success $failed $session.GetCloudSkipped()
        $ui.IndexingProgressText.Text = $endText.Text
        $ui.IndexingProgressDetail.Text = $endText.Detail
    }
    $ui.IndexingProgressEta.Text = ""
    $ui.IndexingStopButton.Visibility = "Collapsed"
    $ui.IndexingProgress.Visibility = "Collapsed"
    # 更新中の帯を閉じていても、終わりの知らせは出す
    $ui.IndexingProgressPanel.Visibility = "Visible"
    finishImportConfirm $exitCode
    $skippedView = getIndexingSkippedView $onlySkipped
    if ($skippedView -and $exitCode -ne 1) {
        # メッセージを読んでいる間に成功の帯が消えないよう、時計を止め、閉じたあとから数え直す
        $script:indexingBannerTimer.Stop()
        showMessage "$($skippedView.Heading)`n`n$($skippedView.Detail)" "OK" "Warning" | Out-Null
        setIndexingBannerBehavior
    }
    $ui.IndexingSearchButton.Visibility = if ($exitCode -ne 1) { "Visible" } else { "Collapsed" }

    $script:sourceFolderMaps = @{}
    refreshIndexingState
    refreshIndexSummary
    loadIndexTree  # 新しいインデックス・フォルダをツリーに出す
    refreshLeftoverIndexes   # 外れたインデックスの削除や、新しい取り込みで、残っているインデックスが変わった
    markTrace "index" "一覧・ツリーの読み直しの終わり"
    writeTrace "index" $workspace.TraceDir
    restartIndexingIfRequested $exitCode
}

function restartIndexingIfRequested {
    # 確認で取り込む拡張子を変えて取りやめた回が終わったあと、同じ対象（onlyNames）で数えから始め直す。1 回だけ呼ぶ。
    # 印は呼ぶ前に必ず消す（エラーで終わった回では始め直さずに消すだけ。startIndexing が断られても、もう呼ばない）
    param (
        [int]$exitCode  # 終わった回の終わり方（1 はエラー）
    )

    $request = $script:indexingRestart
    $script:indexingRestart = $null
    $script:indexingRecounting = $false
    if ($null -eq $request -or $exitCode -eq 1) {
        return
    }
    if ($script:targetsNaming) {
        # ネットワークのワークスペースでは、終わった回のあとの一覧の読み込み（インデックス名の割り当て）が裏で続いていて、
        # 今 startIndexing を呼ぶと黙って断られ、更新の流れが消える。終わるまで待ってから始め直す
        waitAndRestartIndexing @($request.OnlyNames) $request.ConfirmText
        return
    }
    startIndexingForRestart @($request.OnlyNames) $request.ConfirmText
}

function startIndexingForRestart {
    # 始め直しの数えを始める（startIndexing）。数えの間だけ、帯に数えたファイル数を出す（getIndexingScanText）
    param (
        [string[]]$onlyNames,
        $confirmText = $null   # インポートの後の確認の文言（startIndexing の confirmText）
    )

    $script:indexingRecounting = $true
    startIndexing -onlyNames @($onlyNames) -confirmText $confirmText
    if ($null -eq $script:indexingSession -and -not $script:indexingPreparing) {
        $script:indexingRecounting = $false   # 始められなかった（ほかの操作の最中など）
    }
}

function waitAndRestartIndexing {
    # 一覧の読み込みが終わる（targetsNaming が下りる）まで 0.3 秒ごとに見て、終わったら始め直す。最大 45 秒待ち、それでも終わらなければ諦める
    param (
        [string[]]$onlyNames,
        $confirmText = $null
    )

    $state = @{ Tries = 0; Names = @($onlyNames); ConfirmText = $confirmText }
    # GetNewClosure したブロックの中では、関数を名前で探せないことがあり、$script: の変数も読めない（閉じ込めた側の空の変数になる）。
    # 関数は先に変数へ入れて & で呼び、$script: の変数は閉じ込めの外のブロックで読む
    $start = ${function:startIndexingForRestart}
    $decide = ${function:getIndexingRestartWaitAction}
    $giveUpText = getIndexingRestartGiveUpText
    $notify = ${function:setStatus}
    $isNaming = { [bool]$script:targetsNaming }
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(300)
    $timer.Add_Tick({
            $state.Tries++
            $action = & $decide (& $isNaming) $state.Tries 150
            if ($action -eq "wait") { return }
            $timer.Stop()
            if ($action -eq "start") {
                try { & $start $state.Names $state.ConfirmText } catch { Write-Verbose "始め直しを始められませんでした" }
            } else {
                & $notify $giveUpText
            }
        }.GetNewClosure())
    $timer.Start()
}

function getIndexingScanText {
    # 数える間（クロール）の帯の文。確認で拡張子を変えて始め直した数えの間だけ、数えたファイル数つきの文（getIndexingRecountBandText）にする。
    # 受け渡しの口が無いとき（画面なし）は読まない
    $channel = $(if ($script:indexingSession) { $script:indexingSession.Channel } else { $null })
    if ($script:indexingRecounting -and $null -ne $channel) {
        return getIndexingRecountBandText ([int]$channel.CountedFiles)
    }
    return "更新するファイルを調べています…"
}

$script:indexingTimer = newTimer 1000 { safe { updateIndexingProgress } }

$ui.IndexFilesButton.Add_Click({ safe { openIndexFilesDialog } })
$ui.IndexingButton.Add_Click({ safe { startIndexing } })
$ui.IndexingStopButton.Add_Click({ safe { stopIndexing } })
$ui.IndexingResumeButton.Add_Click({ safe { startIndexing } })
$ui.IndexingSearchButton.Add_Click({ safe { selectScreen "SearchTab" } })
