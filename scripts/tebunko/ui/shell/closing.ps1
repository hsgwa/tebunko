# startGui の一部: 閉じるときの処理（closeTimer・enterIndexingStopFlow・Closing）。startGui の中で読み込む。

    # 閉じるときの順番（docs/design/structure/closing.md「閉じるときの順番」）。インデックス作成は画面のプロセスのスレッドで動くため、
    # 止めてから閉じる。止め終わるまで閉じるのを保留し、closeTimer が終わりを待ってから閉じ直す
    $script:closeWaiting = $false   # インデックス作成が止まるのを待っている
    $script:closeDeadline = $null   # これを過ぎたら、インデックス作成が起動した Office を止める
    $script:closeKilled = $false    # Office を止めた
    $script:closeReady = $false     # 待ち終えた（もう聞かずに閉じる）
    $script:closeAskedDuringIndexJob = $false   # インデックスの削除・名前の変更の途中で閉じようとして、終わるのを待った
    $script:closeAfterIndexJob = $false         # 終わったら閉じる（待たずに閉じると決めたら戻す）
    $script:closeConfirming = $false            # 待たずに閉じるかの確認を出している間（この間は仕事が終わっても閉じず、確認から戻ったときに閉じる）
    ${closeWaitSeconds} = 60        # インデックス作成が止まるのを待つ時間。過ぎたら Office を止めて、さらに closeKillWaitSeconds 待つ
    ${closeKillWaitSeconds} = 15

    $script:closeTimer = newTimer 500 {
        safe {
            $now = Get-Date
            $indexing = isIndexing
            if ($indexing) {
                # 待つ間のダイアログの棒と、強制終了までの秒読みを書き換える
                try { updateIndexingWait (getIndexingProgress) ($script:closeDeadline - $now).TotalSeconds } catch { }
            }
            if (!$indexing) {
                $script:closeTimer.Stop()
                $script:closeReady = $true
                try { closeIndexingWait } catch { }   # 待つ間のダイアログ → 暗幕 → 本体の窓の順に閉じる（失敗しても本体は閉じる）
                $window.Close()
            } elseif (!$script:closeKilled -and $now -gt $script:closeDeadline) {
                # 取り込み中の Office が応答しない。インデックス作成が起動した Office だけを PID で止める（COM の呼び出しが戻る）
                $script:closeKilled = $true
                [void]$script:indexingSession.KillOffice()
                $script:closeDeadline = $now.AddSeconds(${closeKillWaitSeconds})
            } elseif ($script:closeKilled -and $now -gt $script:closeDeadline) {
                # それでも止まらない。スレッドは finally で止める
                $script:closeTimer.Stop()
                $script:closeReady = $true
                try { closeIndexingWait } catch { }
                $window.Close()
            }
        }
    }

    function enterIndexingStopFlow {
        param ([switch]$NoWaitDialog)   # 待つ間のダイアログを出さない（確認で取りやめた後始末の間は、すぐ終わるため）
        #ふだんの「インデックス作成を止めてから閉じる」流れに入る（すでに入っていれば、足りない手順だけ補う）。
        # 1 つずつ try で包み、失敗しても残りを進める（finally の IndexingSession.Close の打ち切りに頼らず、
        # 取り込み中のファイルの区切りで止めるため）。closeDeadline は $null のときだけ決め、
        # 途中の失敗で入り直しても延ばし直さない。closeTimer も動いていなければ動かす（M1: ここが抜けていると、
        # 流れの途中の失敗のあとに closeTimer が動かず、外から止める以外に閉じられなくなる）
        writeCloseTrace -Path $script:closeTracePath -Step "enterIndexingStopFlow"
        $script:closeWaiting = $true
        try { $script:indexingTimer.Stop() } catch { }
        try { $script:indexingSession.Stop() } catch { }
        try { $ui.IndexingStopButton.IsEnabled = $false } catch { }
        try { $ui.IndexingProgressText.Text = "更新を止めています…" } catch { }
        try { $ui.IndexingProgressDetail.Text = "読み取り中のファイルが終わると、画面を閉じます。" } catch { }
        try { setStatus "更新を止めてから閉じます…" } catch { }
        try {
            if ($null -eq $script:closeDeadline) {
                $script:closeDeadline = (Get-Date).AddSeconds(${closeWaitSeconds})
            }
            if (!$script:closeTimer.IsEnabled) {
                $script:closeTimer.Start()
            }
        } catch { }
        # 待つ間のダイアログは、失敗しても止める流れの邪魔にしない
        if (!$NoWaitDialog) {
            try {
                showIndexingWait "CloseWait"
                updateIndexingWait (getIndexingProgress) ($script:closeDeadline - (Get-Date)).TotalSeconds
            } catch { }
        }
    }

    $window.Add_Closing({
        param ($sender, $e)
        writeCloseTrace -Path $script:closeTracePath -Step "Closing"
        if ($script:closeReady) {
            return
        }
        if ($script:closeWaiting) {
            # 止まるのを待っている間に、もう一度閉じようとした
            $e.Cancel = $true
            return
        }
        try {
            # インデックスの削除・名前の変更の途中で閉じると、設定とインデックスのフォルダの名前が食い違う。終わるまで待って閉じる
            $indexJobClose = getIndexJobCloseAction ([bool]$script:indexBusy) ([bool]$script:closeAskedDuringIndexJob) ([bool]$script:workspaceMoving)
            if ($indexJobClose -eq "waitMove") {
                # ワークスペースを移している途中で閉じると、中身が移動元と移動先に分かれる。止めずに、移し終わってから閉じる
                # （終わりの処理 finishWorkspaceMove が閉じる）。待つ間のダイアログの文だけ「終わったら閉じます」に変える
                $script:closeAfterWorkspaceMove = $true
                $e.Cancel = $true
                # 表示の失敗で、閉じるのを止めたことを取り消さない（catch が Cancel を戻すため、ここで握る）
                try { setStatus (getWorkspaceMoveClosingStatus) } catch { }
                try { updateWaitDialog $script:indexingWait @{ Description = (getWaitDialogView "MovingClose").Description } } catch { }
                return
            }
            if ($indexJobClose -eq "wait") {
                $script:closeAskedDuringIndexJob = $true
                $script:closeAfterIndexJob = $true
                $e.Cancel = $true
                setStatus (getIndexJobClosingStatus)
                return
            }
            if ($indexJobClose -eq "confirm") {
                # 待っている間にもう一度閉じようとした。ステータスの 1 行は他の表示で消えて見落とされるため、確かめてから閉じる。
                # 確認はこのハンドラの中のモーダル。その間に仕事が終わっても、ここで閉じずに（閉じ途中の窓に Close を呼ばないよう）、戻ってから決める
                $e.Cancel = $true
                $confirm = getIndexJobCloseConfirm
                $script:closeConfirming = $true
                try {
                    $answer = showConfirm `
                        -title $confirm.Title `
                        -heading $confirm.Heading `
                        -hint $confirm.Hint `
                        -choices @(@{ Text = $confirm.CloseText; Value = $confirm.CloseValue; Danger = $true }) `
                        -cancelText $confirm.CancelText
                } finally {
                    $script:closeConfirming = $false
                }
                if ((getIndexJobCloseAfterConfirm ([string]$answer) ([bool]$script:indexBusy)) -eq "close") {
                    $script:closeAfterIndexJob = $false
                    $script:closeReady = $true
                    $window.Dispatcher.BeginInvoke([action]{ $window.Close() }) | Out-Null
                }
                return
            }
            $script:closeAfterIndexJob = $false
            # インデックス作成は画面のプロセスで動いているため、画面を閉じるときは止める
            if (isIndexing) {
                if ((getIndexingCloseAction ([bool]$script:indexingCanceledAtConfirm) ($null -ne $script:indexingRestart)) -eq "stop") {
                    # 確認のダイアログで取りやめた後始末の間。失うものが無いので、聞かず、待つ間のダイアログも出さずに止めてから閉じる
                    $e.Cancel = $true
                    enterIndexingStopFlow -NoWaitDialog
                    return
                }
                $stopConfirm = getIndexingCloseConfirm
                $answer = showConfirm `
                    -title $stopConfirm.Title `
                    -heading $stopConfirm.Heading `
                    -hint $stopConfirm.Hint `
                    -choices @(@{ Text = $stopConfirm.StopText; Value = $stopConfirm.StopValue; Careful = $true }) `
                    -cancelText $stopConfirm.CancelText
                $e.Cancel = $true
                if ($answer -ne $stopConfirm.StopValue -or !(isIndexing)) {
                    if ($answer -eq $stopConfirm.StopValue) {
                        # 聞いている間に終わった
                        $script:closeReady = $true
                        $window.Dispatcher.BeginInvoke([action]{ $window.Close() }) | Out-Null
                    }
                    return
                }
                enterIndexingStopFlow
                return
            }
            if ($script:search) {
                cancelSearch
            }
        } catch {
            # 知らせる処理自体が失敗しても、閉じ方の分岐へは必ず進む（R1）
            try { reportUnexpectedError "画面を閉じる途中" $_ } catch { }
            $stillIndexing = $true
            try {
                $stillIndexing = isIndexing
            } catch {
                # isIndexing 自体が失敗したときは、安全な側（止める流れ）に倒す
                $stillIndexing = $true
            }
            if (!$stillIndexing -and !$script:closeWaiting) {
                # インデックス作成中でなく、まだ止める流れにも入っていなければ、閉じるのを止めない
                $e.Cancel = $false
                return
            }
            # 確認のダイアログは出し直さない。すでに流れの途中（closeWaiting が立っている）で
            # 失敗したときも、抜けている手順（closeTimer が動いていないなど）を補う
            $e.Cancel = $true
            enterIndexingStopFlow
        }
    })
