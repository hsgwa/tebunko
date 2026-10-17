# インデックス作成の本体（invokeIndexer）。画面を使わない起動口（indexer.ps1）と、画面のインデクサのスレッド（IndexingSession）が呼ぶ。
# 設計は docs/design/indexing/parallel.md、docs/design/structure/threads.md「画面とインデクサの受け渡し」。
#
# ・司令のスレッド（invokeIndexer を呼んだスレッド）が、クロール・確認・取り込み一覧・本文インデックスのファイルとシステムインデックスの書き出しを行う。
#   Office の COM には触らない
# ・1 ファイルの取り込み（Office で開いて TSV にする）は、取り込みのワーカー（STA のスレッド。1 種類）が行う。Office のアプリは、
#   初めてそのアプリのファイルを受けたワーカーが 1 つ持ち、そのアプリのファイルはそのワーカーにだけ渡す（IngestDispatcher・selectIngestWorker）。
#   ワーカーの数が 0 のときは、司令のスレッドで取り込む（テストで、途中に割り込むため）
# ・画面とのやり取りは受け渡しの口（newIndexerChannel）で行う。表示内容は indexing_log.txt に書く

# Excel・Word・PowerPoint を利用者に渡したとき・渡すか決められず持ち続けるときのログ（indexing_log.txt）への 1 行。
# office_app.ps1 の handOverApp・keepUndecidedApp が呼ぶ
${officeHandOverNotice} = {
    param ($name, $shown)
    if ($null -eq $shown) {
        # shown が $null: ファイルの一覧を読めず、利用者のファイルがあるか分からない。終了させずに持ち続け、後で読み直す
        writeIndexerLog "${name} のファイルの一覧を読めなかったため、終了させずに残しています（あとで読み直します）" "Yellow"
    } elseif ($shown) {
        writeIndexerLog "開かれたファイルがあるため ${name} を利用者に渡しました"
    } else {
        writeIndexerLog "開かれたファイルがあるため ${name} を利用者に渡しましたが、窓や設定を戻しきれませんでした（終了させずに残しています）" "Yellow"
    }
}

# 取り込みのワーカーのスクリプト。自分の列（tasks）から 1 ファイルずつ取り出して取り込み、結果を results に入れる。
# ワーカーは 1 種類で、どれも STA。どの種類のファイルも受けられ、Office のアプリは受けたファイルに応じて起動する（持ち主の決め方は IngestDispatcher）。
# 列が閉じられたら（CompleteAdding）、Office を終了して終わる
${ingestWorkerScript} = {
    param ($settings, $tasks, $results, $number)
    $ErrorActionPreference = "Stop"
    # indexerLib の部品（Prelude）は IngestDispatcher が、この param ブロックのあとにつなぐ（joinWorkerScript）
    # 置き場所は司令のスレッドと同じにする（読み込み直すと設定から決め直してしまうため）。一時フォルダはスレッドごとに分ける。
    # 部品を読み込んだのと同じスコープ（取り込みのワーカーでは global）に置く
    $own = [Workspace]::new($settings.WorkDir)
    $own.PublishDir = Join-Path $settings.PublishDir "w$number"
    setCallerWorkspace $own
    Set-Variable -Name tmpDir -Value (newWorkerTmpDir $settings.TmpDir $number)
    Set-Variable -Name tmpDirReason -Value $settings.TmpDirReason
    [System.IO.Directory]::CreateDirectory($workspace.PublishDir) | Out-Null
    $script:officePidSink = $settings.OfficePids
    $script:officeRecordDir = $settings.OfficeRecordDir
    $script:officeOwnDir = resolveOwnDir $tmpDir
    $script:onOfficeHandOver = ${officeHandOverNotice}
    # ファイルごとに、Office が要るか（task.Office）で決め直す。始めは Office を使わない
    $script:officeUnavailable = $true
    [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::BelowNormal

    # ワーカーの状態（制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたか）。司令と共有する
    $script:workerState = $settings.WorkerState
    $script:officeWatchdog.State.Abandon = $settings.WorkerState
    if ($settings.UserOfficePids) {
        $script:officeWatchdog.State.Users = $settings.UserOfficePids
    }
    # 中止を求められたら、取り込み中の Office を止めて打ち切る（制限時間を待たない）。
    # Office を持たないワーカーは、止める相手が無いため、制限時間を過ぎても TimedOut を記録するだけ
    $script:officeWatchdog.State.Cancel = $settings.Channel
    $script:officeWatchdog.Start()
    try {
        runIngestWorker $tasks $results $settings.FileTimeoutMinutes $settings.RestartInterval $settings.Channel
    } finally {
        $script:officeWatchdog.Stop()
        # 利用者が使っている PowerPoint を諦めたワーカーは、終了させない（開いているファイルの保存されていない変更を失わせないため）
        if (!$settings.WorkerState.Abandoned) {
            stopAllApps
            # 渡し切れずに持ち続けているアプリは、ここで上限まで仕上げ直す（中止・画面を閉じるときは待たない）
            waitKeptApps ({ [bool]$settings.Channel.Stop }.GetNewClosure())
        }
    }
}

# 司令が並びの先を見る件数（空いているワーカー向けのファイルを先に渡すため）。大きくすると、本文インデックスのファイルの書き出しが遅れるフォルダが増える
${ingestLookAhead} = 200

function runIngestWorker {
    # 取り込みのワーカーの繰り返し（ingestWorkerScript が呼ぶ）。列から 1 件ずつ取り込み、結果を results に入れる。
    # Office が要るファイルを取り込んだあとは、決まった数を取り込むたびに Office を起動し直す
    param (
        $tasks,
        $results,
        [int]$fileTimeoutMinutes,
        [int]$restartInterval,
        $channel = $null   # 受け渡しの口。Stop が立っていたら、列に残っているファイルは取り込まずに Stopped で返す
    )

    $done = 0
    foreach ($task in $tasks.GetConsumingEnumerable()) {
        $result = invokeIngestTask $task $fileTimeoutMinutes $channel
        $results.Add($result)
        if ($result.Postponed -or $result.Stopped) {
            # 後回し（利用者のPowerPointが起動していて使えなかった）は、取り込んだ件数に数えない
            # （件数が変わらないまま後回しが続いたときに、起動し直しを早めないため）。
            # 中止で止めたファイルも数えない（中止のあとは Office を起動し直さない。アプリの片付けは ingestWorkerScript の終わりで行う）
            continue
        }
        $done++
        if (!$script:officeUnavailable -and ($result.TimedOut -or ($done % $restartInterval) -eq 0)) {
            # 制限時間を過ぎて強制終了したアプリは使えないため、すべて終了して次に必要になったときに起動し直す
            stopAllApps
        }
    }
}

function getIngestWorkerCount {
    # Office を使わずに読む仕事の数（取り込みのワーカーの数のうち、Office のアプリ分を除いたもの）を決める。requested が 0 以上ならその数
    # （0 は取り込みのスレッドを使わず、司令のスレッドで取り込む）、負なら設定（ingestThreads）、設定が 0 ならコア数から決める。
    # 取り込むファイルの数より多くはしない
    param (
        [int]$requested,
        [int]$total,
        [int]$configured = 0,
        [int]$processors = [Environment]::ProcessorCount
    )

    $count = if ($requested -ge 0) {
        $requested
    } elseif ($configured -gt 0) {
        [Math]::Min($configured, 4)
    } else {
        getWorkerCount 4 $processors
    }
    return [Math]::Max(0, [Math]::Min($count, $total))
}

function testWorkerAbandoned {
    # このワーカーが、制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたものか（取り込みのワーカーの $script:workerState）
    $state = $script:workerState
    return ($null -ne $state -and [bool]$state.Abandoned)
}

function invokeIngestTask {
    # 1 ファイルを取り込み、結果 @{ RelPath; Ok; Reroute; Postponed; Stopped; TsvCount; Message; TimedOut; ExtractVersion; Log } を返す（例外は投げない）。
    # 取り込みのワーカー、またはワーカーの数が 0 のときは司令のスレッドで動く。表示内容は Log に貯め、司令がログに書く
    param (
        [hashtable]$task,
        [int]$fileTimeoutMinutes,
        $channel = $null   # 受け渡しの口。Stop が立っていれば、取り込まずに Stopped を返す
    )

    # Reroute: Office を使わずに読めなかった（中身が旧形式・パスワード付き）。司令が Word・PowerPoint の持ち主に回し直す
    # Postponed: 利用者のPowerPointが起動していて使えなかった。司令が「未取り込み」のまま次回に回す
    # Stopped: 中止を求められて取り込まなかった・取り込みを打ち切った。失敗にせず、司令が「未取り込み」のまま次回に回す（インデックスは変えない）
    $result = @{ RelPath = $task.RelPath; Ok = $false; Reroute = $false; Postponed = $false; Stopped = $false; TsvCount = 0; Message = ""; TimedOut = $false; ExtractVersion = ""; Log = "" }
    if ($task.ContainsKey("Office")) {
        # 取り込みのワーカーでは、Office が要らないファイルは Office を使わずに読み、読めなければ Office が要ると返す（Reroute）
        $script:officeUnavailable = !$task.Office
    }
    if ($null -ne $channel -and $channel.Stop) {
        $result.Stopped = $true
        return $result
    }
    if (!$tmpDir) {
        # 取り込みの作業フォルダを置けない（selectTmpDir の Brackets・TooLong）ときは、
        # どのファイルも中間 TSV などをこのフォルダに作るため、テキストファイルを含めこのファイルの取り込みをスキップし、取り込みの失敗として記録する
        $result.Message = "$(getTmpDirUnavailableMessage $tmpDirReason)このファイルの取り込みをスキップしました。"
        return $result
    }
    if (testWorkerAbandoned) {
        # 利用者が使っている PowerPoint を諦めたワーカー。取り込まず、後回しにする（司令は結果を受け取らずに後回しにしている）
        $result.Postponed = $true
        return $result
    }
    $log = New-Object System.IO.StringWriter
    $previousLog = $script:indexerLog
    $script:indexerLog = $log
    try {
        clearTmpDir
        $script:officeWatchdog.State.TimedOut = $false
        $script:officeWatchdog.State.Cancelled = $false
        $script:officeWatchdog.State.Deadline = (Get-Date).AddMinutes($fileTimeoutMinutes)
        try {
            $result.TsvCount = ingestFile $task.SourcePath
        } finally {
            $script:officeWatchdog.State.Deadline = [datetime]::MaxValue
        }
        if ($script:officeWatchdog.State.Cancelled) {
            # 中止で Office を止めた。取り込みが終わっていても、書き出さずに次回に回す
            $result.Stopped = $true
            return $result
        }
        if (testWorkerAbandoned) {
            # 戻ってきたときには、司令がこのファイルを後回しにしている。書き出さない
            # （司令が同じフォルダを書き換えるのと重ならないようにし、取り込み一覧・インデックスを変えない）
            $result.Postponed = $true
            return $result
        }
        publishTsv (getBookDir $task.RelPath)
        $result.ExtractVersion = [string](getExtractVersion $task.RelPath)
        $result.Ok = $true
    } catch {
        $base = $_.Exception.GetBaseException()
        if (testWorkerAbandoned) {
            # 諦めたワーカーで、利用者の PowerPoint が閉じられるなどして戻ってきた。失敗にも、アプリの終了にもしない
            $result.Postponed = $true
            return $result
        }
        if ($script:officeWatchdog.State.Cancelled -or ($null -ne $channel -and $channel.Stop)) {
            # 中止で Office を止めたため（または中止を求められたあとの失敗）。失敗にせず、強制終了の回数にも数えず、次回に回す
            $result.Stopped = $true
            return $result
        }
        if (isOfficeRequiredException $_.Exception) {
            $result.Reroute = $true
            return $result
        }
        if (isOfficeAppInUseException $_.Exception) {
            # getApp が投げた例外。$script:apps に入っていないため stopApp は呼ばない
            $result.Postponed = $true
            return $result
        }
        # .docx・.pptx を直接読んでサイズの上限を超えたとき（ZipSizeLimitException）は、原因を調べられるよう
        # 部品名・大きさ・部品ごとか合計かをインデックス作成のログに書く。画面・取り込み一覧の
        # $result.Message は今までどおり簡潔な文言のまま（describeIngestError がそのまま通す）
        writeZipSizeLimitLog $base
        $message = describeIngestError $_.Exception
        if ($script:officeWatchdog.State.TimedOut) {
            $message = "${fileTimeoutMinutes} 分以内に更新が終わらなかったため中止しました（Officeアプリを強制終了しました）"
        }
        $result.Message = $message
        # アプリが不安定になっている可能性があるため終了する（次に必要になったときに起動し直す）。
        # 暗号化の判定でOfficeにまだ触れていない失敗（throwProtectionFailure）は、終了させない
        # （IRM・パスワード付きのファイルが並んでも、そのたびにOfficeを起動し直さないため）
        if (!$script:officeUnavailable -and !$base.Data.Contains("OfficeUntouched")) {
            try { stopApp (getAppName $task.RelPath) } catch {}
        }
    } finally {
        # Excel・Word・PowerPoint で利用者が開いたファイルがあれば、閉じずに利用者に渡す（次のファイルは新しいアプリで取り込む。PowerPoint は後回しになる）
        $appName = getAppName $task.RelPath
        if (!$script:officeUnavailable -and $appName -and !(testWorkerAbandoned)) {
            try { [void](handOverForeignApp $appName) } catch {}
        }
        $result.TimedOut = [bool]$script:officeWatchdog.State.TimedOut
        $script:indexerLog = $previousLog
        $result.Log = $log.ToString()
    }
    return $result
}

function invokeIndexer {
    # インデックスを作成し、終了コード（0 = 完了 / 1 = 続けられないエラー / 2 = 中止（確認で取りやめた場合を含む））を返す。
    # 終了コード・エラーの内容は受け渡しの口（channel）にも入れる（ExitCode は最後に入れる。画面は ExitCode が入ったら終わったとみなす）
    param (
        $channel
    )

    $script:indexerChannel = $channel
    $exitCode = 1
    $mutex = $null
    try {
        # 既定のワークスペースにほかのファイルが置いてあれば、インデックスのファイルと混ざるため作成しない。
        # 空でないフォルダにログを書かないよう、ログを開く前に確かめる（画面は起動する前に同じ確認をする）
        $workspaceBlock = getWorkspaceBlockMessage
        if ($workspaceBlock) {
            if ($script:settingsRecovery) {
                writeIndexerLog (getSettingsRecoveryMessage $script:settingsRecovery) "Yellow"
                $script:settingsRecovery = ""
            }
            writeIndexerLog $workspaceBlock "Red"
            $channel.Error = $workspaceBlock
            $exitCode = 1
        } else {
            # 同じ work に対してインデックス作成を 2 つ動かすと、取り込み一覧・インデックスが食い違うため 1 つだけ動かす。
            # 画面と indexer.ps1 の両方から動かせるため、鍵は work のパスから作る。ログに触る前に確かめる
            $mutex = newAppMutex "indexer" $workspace.Dir ${appId}
            if (!$mutex.Acquired) {
                throw "ほかのインデックス作成が実行中です。インデックス作成が終わってから実行してください。"
            }
            [System.IO.Directory]::CreateDirectory($workspace.Dir) | Out-Null
            # 鍵を取れた後にだけ、前回の記録を残す（取れなかったときはほかのインデックス作成の記録に触らない）
            [void](rotateLogFile $workspace.IndexingLogFile 0)
            $writer = New-Object System.IO.StreamWriter($workspace.IndexingLogFile, $false, ${utf8Bom})
            $writer.AutoFlush = $true
            $script:indexerLog = $writer
            try {
                # 先頭に版と環境を書く（画面なしで動かしたときにコンソールへ出さないよう writeIndexerLog は使わない）。書けなくても続ける
                foreach ($headerLine in @(getToolDiagnosticHeader)) {
                    $writer.WriteLine($headerLine)
                }
            } catch {}
            if ($script:settingsRecovery) {
                # 画面なしで起動したとき、起動口が壊れた設定ファイルを退避している。ログの先頭に知らせる
                writeIndexerLog (getSettingsRecoveryMessage $script:settingsRecovery) "Yellow"
                $script:settingsRecovery = ""
            }
            # 途中の処理が出力した値が混ざらないよう、最後の値（return した終了コード）を使う
            $exitCode = [int]@(invokeIndexerBody $channel)[-1]
        }
    } catch {
        # 続けられないエラー。画面に伝える
        $message = $_.Exception.Message
        writeIndexerLog "＜エラー＞"
        writeIndexerLog $message "Red"
        $channel.Error = $message
        $exitCode = 1
    } finally {
        if ($script:indexerLog) {
            try { $script:indexerLog.Dispose() } catch {}
            $script:indexerLog = $null
        }
        if ($mutex) {
            if ($mutex.Acquired) {
                try { $mutex.Mutex.ReleaseMutex() } catch {}
            }
            $mutex.Mutex.Dispose()
        }
        $script:indexerChannel = $null
        $channel.ExitCode = $exitCode
    }
    return $exitCode
}

function flushPendingPublish {
    # 書き出し待ちのフォルダを本文インデックスへまとめる（書き出しは司令が呼ぶ。PendingPublish は決めて返すだけ）。
    # -All は、取り込み中・まだ渡していない数を見ずに、書き出し待ちのフォルダをすべて書き出す（クロールの前後・インデックス作成の終わり）
    param (
        $pending,
        [switch]$All
    )

    $flush = if ($All) { $pending.TakeAll() } else { $pending.TakeFlushable() }
    if ($flush.Count -eq 0) {
        return
    }
    try {
        [void](publishIndexFolders $flush)
    } catch {
        writeIndexerLog "    インデックスをまとめられませんでした（次のインデックス作成でまとめ直します）: $($_.Exception.Message)" "Yellow"
    }
}

function reportIngestProgress {
    # 取り込みの進み具合を画面に伝える（画面はこの 1 行から進み具合を作る。取り込み一覧は読まない）。残りの数はここだけで数える
    param (
        $run
    )

    $failures = $run.Ledger.Failures.Count
    $run.Reporter.Progress(${indexingPhaseIngest}, ($run.SuccessCount + $failures), ($run.Total - $run.SuccessCount - $failures), $failures, $run.CurrentPath)
}

function completeIngestJob {
    # 取り込みの 1 件の終わり方（成功・失敗・後回し・中止）を、取り込み一覧・書き出し待ち・進み具合・ログに反映する。司令の結果の受け方はここだけ。
    #   received: IngestDispatcher.Receive が返す @{ Kind; Entry; ... }。Kind が Completed（Result を持つ）・Stopped（First・RelPath を持つ）・
    #             Postponed・RerouteAbandoned（Reason を持つ）のどれか。諦めたワーカーの取り込み中だったファイルと、渡さずに残したファイルも
    #             Kind Postponed で渡す（渡さずに残したものは Undispatched を $true にする。Entry は @{ Row; Number; Folder }）
    #   run: 司令がつなぐ値 @{ Ledger; Pending; Reporter; Dispatcher; Previous; Total; SuccessCount; PostponedCount; CurrentPath }
    param (
        $run,
        $received
    )

    $ledger = $run.Ledger
    $pending = $run.Pending
    $entry = $received.Entry
    if ($received.Kind -eq "Stopped") {
        # 中止を求められて、取り込まなかった・打ち切った（回し直しも、中止のあとは回さない）。失敗にはしない。
        # 取り込み一覧は「未取り込み」のまま（AddRow しない）、インデックスは変えず、次回のインデックス作成で取り込む
        $pending.Complete($entry.Folder)
        if ($received.First) {
            writeIndexerLog ""
            writeIndexerLog "中止の要求を受けたため、インデックス作成を中止します。（取り込み中だったファイルと残りは次回取り込みします）" "Yellow"
        }
        writeIndexerLog "    中止のため、取り込みませんでした（次回取り込みします）: $($received.RelPath)" "Yellow"
    } elseif ($received.Kind -ne "Completed") {
        # 後回し: 利用者の PowerPoint が起動していて使えなかった（Postponed）か、回し直し先の持ち主が諦めたワーカーだった（RerouteAbandoned）・
        # 諦めたワーカーで取り込み中だった・諦めたワーカーのアプリのファイルだった。
        # 失敗として理由を見せ、取り込み一覧は「未取り込み」のまま残して次回のインデックス作成で取り込む
        # （インデックスは変えない。$ledger.AddRow・$pending.Add はしない）。
        # 取り込み一覧の行は「未取り込み」のまま、理由を「エラー」列に書く（「失敗」にすると、次の更新で飛ばされるため）。
        # 前の行があれば、TSV 数・取り込み日時・抽出版は前のまま残す（インデックスは変えていない）
        if ($received.Undispatched) {
            $pending.Skip($entry.Folder)
        } else {
            $pending.Complete($entry.Folder)
        }
        $run.PostponedCount++
        $row = $entry.Row
        $relPath = [string]$row.相対パス
        $old = $null
        if ($run.Previous.TryGetValue($relPath, [ref]$old) -and $old) {
            $row.TSV数 = $old.TSV数
            $row.取り込み日時 = $old.取り込み日時
            $row.抽出版 = $old.抽出版
        }
        $row.状態 = ${stateNew}
        $row.エラー = $received.Reason
        $ledger.AddFailure($relPath, $received.Reason)
        writeIndexerLog ("[{0}/{1}] {2}" -f $entry.Number, $run.Total, $relPath)
        writeIndexerLog "    取り込めませんでした（次の更新で取り込みます）: $($received.Reason)" "Yellow"
    } else {
        $result = $received.Result
        $row = $entry.Row
        writeIndexerLog ("[{0}/{1}] {2}" -f $entry.Number, $run.Total, $result.RelPath)
        foreach ($line in ($result.Log -split "\r?\n")) {
            if ($line) {
                writeIndexerLog $line
            }
        }
        if ($result.Ok) {
            $pending.Add($result.RelPath, $false)
            # 高速検索: このフォルダの システムインデックスを作り直すまで、検索ではこのフォルダを必ず照合させる
            if (!(markSystemIndexChanged @([System.IO.Path]::GetDirectoryName($result.RelPath)))) {
                writeIndexerLog "    システムインデックスの状態を書き込めませんでした（インデックス作成の終わりに作り直します）。" "Yellow"
            }
            writeIndexerLog "    TSV $($result.TsvCount) 件を作成しました。"
            $row.状態 = ${stateDone}
            $row.TSV数 = [string]$result.TsvCount
            $row.エラー = ""
            $row.抽出版 = $result.ExtractVersion
            $run.SuccessCount++
        } else {
            writeIndexerLog "    取り込みに失敗しました: $($result.Message)" "Red"
            $row.状態 = ${stateFailed}
            $row.TSV数 = ""
            $row.エラー = $result.Message
            $row.抽出版 = ""
            $ledger.AddFailure($result.RelPath, $result.Message)
        }
        # 中断しても結果が残るよう、1件ごとに取り込み一覧へ追記する（最後に1ファイル1行にまとめ直す）
        $row.取り込み日時 = formatFileTime (Get-Date)
        $ledger.AddRow($row)
        $pending.Complete($entry.Folder)
    }
    $ledger.WriteIngestingFiles($run.Dispatcher.IngestingEntries())
    reportIngestProgress $run
    # 取り込みが揃ったフォルダ（取り込み中が無く、まだ渡していないファイルも無い）を、本文インデックスのファイルに書き出す
    flushPendingPublish $pending
}

function invokeIndexerBody {
    # invokeIndexer の本体（鍵を取り、ログを開いた後）。終了コードを返す。続けられないエラーは例外にする。
    # 取り込み一覧・取り込み中のファイル・今回の失敗と消えた行は StatusLedger（$ledger）、
    # 本文インデックスへの書き出し待ちは PendingPublish（$pending）、進み具合・画面の確認待ちは IndexingReporter（$reporter）、取り込みの振り分けは IngestDispatcher（$dispatcher）にまとめる。
    # 4 つは互いを呼ばず、ここ（司令）がつなぐ（docs/design/structure/classes.md「クラスのつなぎ方」）
    param (
        $channel
    )

    $ledger = [StatusLedger]::new($workspace)
    $pending = [PendingPublish]::new()
    $reporter = [IndexingReporter]::new($channel)

    # この回で最後まで作り終えたインデックス（失敗したファイルの無いもの）の目録の行を「作り終えた」にし、古い形式の分を消す。
    # 中止・続けられないエラーのときは呼ばない（行は「作成中」のまま。古い形式の分も残る）。消せないものは残してログに出す
    function completeRoundIndexes {
        param ([string[]]$names)

        $failedPrefixes = @($ledger.Failures | ForEach-Object { ([string]$_.RelPath).Split("\")[0] })
        $finishedNames = @($names | Where-Object { $failedPrefixes -notcontains $_ })
        if ($finishedNames.Count -eq 0) {
            return
        }
        $completion = completeIndexFormat $finishedNames $workspace
        if (!$completion.Updated) {
            writeIndexerLog "インデックスの目録を書けませんでした（古い形式の分は消していません）。" "Yellow"
        }
        foreach ($failedEntry in $completion.Failed) {
            writeIndexerLog "古い形式のインデックスを消せませんでした（次のインデックス作成の終わりにもう一度消します）: $failedEntry" "Yellow"
        }
    }

    $reporter.Progress(${indexingPhaseCrawl}, 0, 0, 0, "インデックス作成の準備をしています…")

    $restartInterval = 100  # Officeアプリを再起動する間隔（取り込みのスレッドごとのファイル数）。メモリ肥大化対策（再起動 1 回で約 2 秒かかるため、間隔を詰めすぎない）
    $fileTimeoutMinutes = ${ingestFileTimeoutMinutes}  # 1ファイルの取り込みの制限時間（分）。超えたらOfficeアプリを強制終了し、そのファイルは失敗とする
    $approvalTimeoutMinutes = 60  # 画面の返事を待つ制限時間（分）。画面が返事をしない場合に待ち続けないよう打ち切る
    $interruptLimit = 2       # 取り込み中に続けて強制終了した回数がこれに達したファイルは、失敗として以降スキップする
    $failureListLimit = 50    # 終了時に失敗したファイルと原因を表示する最大件数（残りは取り込み一覧で確認する）
    $cloudListLimit = 50      # クラウドにだけあるため残したファイルを、ログに 1 件ずつ書く最大件数（残りは「ほか N 件」にする）

    $targetFolders = @(getTargetFolders)
    if ($targetFolders.Count -eq 0) {
        throw "クロール対象フォルダがありません。画面でクロール対象フォルダを追加してください。"
    }
    if (@($targetFolders | Where-Object { $_.Enabled }).Count -eq 0) {
        throw "チェックの付いたクロール対象フォルダがありません。画面で取り込むフォルダにチェックを付けてください。"
    }

    # 前の版のインデックス（index\）が見つかれば知らせ、取り込み直しを始めるときは前の版のシステムインデックスを片付ける。
    # インデックスのフォルダ（store\）を作る前に行う（片付けに失敗したら作らずに終わる）
    $legacyState = getLegacyIndexState $workspace.Dir
    $legacyMessage = getLegacyIndexMessage $workspace.Dir $legacyState.HasLegacyIndex
    if ($legacyMessage) {
        writeIndexerLog $legacyMessage "Yellow"
    }
    if (testLegacyCleanupNeeded $legacyState) {
        writeIndexerLog "前の版のインデックス（高速検索用）を片付けています…"
        $cleanup = clearLegacySystemIndex $workspace.Dir
        if (!$cleanup.Ok) {
            throw "前の版のインデックス（高速検索用）を片付けられませんでした（$($cleanup.Reason)）。tebunko の画面やエクスプローラーで開いていれば閉じてから、もう一度インデックス作成を行ってください。"
        }
    }

    # 古い形式（作り直す前の形式）のインデックスがあれば知らせる。作り直しは、この回で全部を取り込み直す（古い取り込み一覧は読まない）
    $formatState = getWorkspaceIndexFormat $workspace
    if ($formatState.State -eq ${indexStateNewer}) {
        throw ${indexNewerMessage}
    }
    foreach ($formatName in $formatState.Indexes.Keys) {
        $formatEntry = $formatState.Indexes[$formatName]
        if ($formatEntry.Message) {
            writeIndexerLog "インデックス「${formatName}」: $($formatEntry.Message)" "Yellow"
        }
    }

    [System.IO.Directory]::CreateDirectory($workspace.IndexDir) | Out-Null
    # 前の版から続けて使うワークスペース・別のドライブへ写したワークスペース・取り込んだワークスペースもここで付く
    # （継ぐかどうかに頼らず、ここで一度だけインデックスのフォルダ全体に付ける。入れた直後にはそれぞれの場所でも付ける）。
    # 根フォルダに未だ付いていない回（主に初回）は下をすべてたどるため時間がかかるので、先に 1 行残しておく
    writeIndexerLog "インデックスのフォルダを Windows Search の対象から外しています…"
    $nci = setWindowsSearchExcluded $workspace.StoreDir -Recurse
    if (!$nci.Ok) {
        writeIndexerLog "インデックスのフォルダを Windows Search の対象から外せませんでした（$($nci.Reason)）。高速検索が効くまで時間がかかることがあります。" "Yellow"
    }
    $selected = initTmpDir
    $script:tmpDir = $selected.Dir
    $script:tmpDirReason = $selected.Reason
    [System.IO.Directory]::CreateDirectory($workspace.PublishDir) | Out-Null

    # クロール対象フォルダごとにインデックス名（インデックスのフォルダ直下のフォルダ名）を決める。前回と同じフォルダは同じ名前を使う
    $status = $ledger.ReadStatus()
    if ($status.Rejected -gt 0) {
        writeIndexerLog "取り込み一覧に読めない行が $($status.Rejected) 行ありました（無視しました）。" "Yellow"
    }
    $planner = [IngestPlanner]::new(@{
        Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed}
    })
    $catalog = [IndexCatalog]::new($workspace, ${settingsFile}, $ledger)
    $planner.AssignNames($targetFolders, $status)
    $folders = $planner.Folders
    $previous = $planner.Previous
    # 設定から外れたインデックスは、ここでは消さない。確認ダイアログに「削除予定」として出し、［更新を開始］が押されたときだけ消す
    $dropped = $planner.Dropped
    # 前回のインデックス作成が途中で止まり、本文インデックスのファイルに入れていない TSV（元のファイルごとのフォルダ）が残っていれば、先に入れる
    # （中止を求められていれば、この段は飛ばして次回に回す。段の途中では止めない）
    $leftover = if ($channel.Stop) { @() } else { findIndexFoldersWithBooks $workspace.IndexDir }
    if ($leftover.Count -gt 0) {
        writeIndexerLog "集約ファイルに入れていないインデックス（$($leftover.Count) フォルダ）をまとめています…"
        $reporter.Progress(${indexingPhaseCrawl}, 0, 0, 0, "集約ファイルに入れていないインデックスをまとめています…")
        foreach ($folder in $leftover) {
            $pending.MarkFolder($folder)
        }
        flushPendingPublish $pending -All
    }

    # 新しく割り当てたインデックス名を設定に保存する
    $planner.SaveNames($targetFolders, { param($assigned) $catalog.StoreAssignedNames($assigned) }.GetNewClosure())
    $folders = $planner.Folders

    writeIndexerLog "出力先フォルダ: $($workspace.IndexDir)"
    writeIndexerLog ""
    writeIndexerLog "クロールしています..."
    # 取り込み一覧の「済」に対してインデックス（TSV）が残っているかを調べるため、今あるTSVの数を数えておく
    # （利用者が work\content_index を直接削除した場合に、「済」のまま検索できなくなるのを防ぐ）
    # （中止を求められていれば確認を飛ばす。次のクロールがすぐ止まるため、件数は使わない）
    $indexCounts = $null
    if (!$channel.Stop) {
        $reporter.Progress(${indexingPhaseCrawl}, 0, 0, 0, "更新済みのインデックスを確認しています…")
        $indexCounts = getIndexTsvCounts
    }
    if ($null -eq $indexCounts -and !$channel.Stop) {
        writeIndexerLog "  インデックスのフォルダを調べられないため、インデックスが残っているかの確認は行いません。" "Yellow"
    }
    # 選んだインデックスだけを更新するとき（OnlyNames）。選ばなかったインデックスは、この回では一切触らない
    # （取り込み一覧の行は前回のまま・確認の表に出さない・失敗の一覧に加えない）
    if (@($channel.OnlyNames).Count -gt 0) {
        $channel.OnlySkipped = @($planner.SelectOnly([string[]]@($channel.OnlyNames)))
        if ($planner.Only.Selected.Count -eq 0) {
            throw "選んだインデックスに、更新できるものがありません。"
        }
    }
    # クロールして、取り込む順番・画面の確認に出す件数・取り込み一覧に書く行を組み立てる
    $crawlPhase = ${indexingPhaseCrawl}
    $onCrawlFolder = { param ($target) $reporter.Progress($crawlPhase, 0, 0, 0, "[$($target.Name)] の対象ファイルを探しています… $($target.Path)") }.GetNewClosure()
    # フォルダを読むごと・ファイルを調べる途中で中止を見る（中止を求められたら読んだところまでで止める）
    $planner.Crawl($indexCounts, $onCrawlFolder, { [bool]$channel.Stop }.GetNewClosure())
    foreach ($removedPath in $planner.RemovedPaths) {
        $pending.Add($removedPath, $true)
    }
    $rows = $planner.Rows
    $targets = $planner.Targets
    $failed = $planner.Failed
    $plan = $planner.Plan
    # クラウドにだけあるファイル（取り込むかは確認の後に決める）
    $cloudAll = $planner.Cloud
    # 消さなかったインデックス（取りやめ・画面なしの実行）は、取り込み一覧に前回のまま残す（次の回にまた見つけて確認に出すため）
    $keptDropped = $dropped

    # 画面から始めた場合は、数えた件数を画面に出して、取り込むかどうかの返事を待つ。
    # 「更新不要」かどうかも、この件数を見て画面が知らせる（取り込み対象が 0 件でも、前回失敗の再取り込みを選べる）
    $retryTargets = [bool]$channel.RetryFailed
    $includeCloud = [bool]$channel.IncludeCloud
    # 取りやめ。取り込み対象にした行は前回の記録のまま（一覧に無かったファイルは記録しない）にする。
    # 取り込み一覧自体は書き直す（無くなったファイルの削除を反映する必要があるため）
    function cancelBeforeIngest {
        param ([string]$message)

        $keep = $planner.KeepForCancel()
        writeIndexerLog ""
        writeIndexerLog $message "Yellow"
        writeStatusKeepingDropped $ledger $folders $keep $keptDropped $previous
        writeSourceFolderFile $folders
        $reporter.Progress(${indexingPhaseFinish}, 0, 0, 0, "インデックス作成を取りやめました")
        removeTmpDir
        return 2
    }
    if ($planner.Stopped -or $channel.Stop) {
        # 一覧化の途中で中止を求められた。確認は出さずに取りやめる
        return (cancelBeforeIngest "中止の要求を受けたため、取り込み対象の一覧化を中止しました。（読み終えたフォルダで無くなっていたファイルは取り込み一覧から外し、ほかは前回のままです）")
    }
    if ($channel.ConfirmTargets) {
        $answer = $reporter.WaitForApproval(${indexingPhaseConfirm}, $plan.ToArray(), $targets.Count, $failed.Count, $approvalTimeoutMinutes)
        if ($null -eq $answer) {
            return (cancelBeforeIngest "画面で取りやめたため、取り込みません。（取り込み一覧は前回のままです）")
        }
        $retryTargets = $answer.RetryFailed
        $includeCloud = [bool]$answer.IncludeCloud
        removeDroppedFolders $dropped
        $keptDropped = @()
    } else {
        foreach ($item in $dropped) {
            writeIndexerLog "設定に無いインデックス（$($item.Name): $($item.Path)）は削除せずに残しました。画面の［すべて更新］で確認して削除できます。" "Yellow"
        }
    }

    # クラウドにだけあるファイル: ダウンロードして取り込むと選ばれたものだけ取り込み、あとは前回のまま残す
    # （取り込み一覧の行は createTargetList が前回の行にしてある。次のインデックス作成で、また確かめる）
    $cloudSplit = splitCloudItems $cloudAll $retryTargets $includeCloud
    if ($cloudSplit.Included.Count -gt 0) {
        $includedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($item in $cloudSplit.Included) {
            [void]$includedPaths.Add($item.RelPath)
        }
        $rebuilt = New-Object System.Collections.Generic.List[object]
        foreach ($row in $rows) {
            if (!$includedPaths.Contains($row.相対パス)) {
                $rebuilt.Add($row)
            }
        }
        $rows = $rebuilt
        writeIndexerLog ""
        writeIndexerLog "クラウドにだけあるファイル $($cloudSplit.Included.Count) 件は、ダウンロードして取り込みます。"
        foreach ($item in $cloudSplit.Included) {
            $rows.Add($item.Row)
            $targets.Add($item.Row)
        }
    }
    # 前回失敗したファイルを再取り込みしない（チェックなし）なら、クラウドにあるかどうかに関係なく残すため、クラウドのせいで残した件数には数えない
    $cloudSkippedItems = @($cloudSplit.Skipped | Where-Object { $retryTargets -or -not $_.Failed })
    $cloudSkipped = $cloudSkippedItems.Count
    $channel.CloudSkipped = $cloudSkipped
    if ($cloudSkipped -gt 0) {
        writeIndexerLog ""
        writeIndexerLog "クラウドにだけあるファイル $cloudSkipped 件は、ダウンロードせずに残します。（次のインデックス作成でも確かめます）" "Yellow"
        foreach ($group in ($cloudSkippedItems | Group-Object { ([string]$_.RelPath).Split("\")[0] })) {
            writeIndexerLog "  [$($group.Name)] $($group.Count) 件"
        }
        $listed = 0
        foreach ($item in $cloudSkippedItems) {
            # 大きいフォルダでは毎回この件数ぶん出るため、先頭の $cloudListLimit 件だけ書き、残りは件数にする
            if ($listed -lt $cloudListLimit) {
                writeIndexerLog "    $($item.RelPath)"
                $listed++
            }
            if ($item.PendingOld) {
                # 前回未完了のまま残っている取り込み途中のインデックスは、使えないため消す（一覧にも行を残さない）
                removeBookDir (getBookDir $item.RelPath)
                $pending.Add($item.RelPath, $true)
            }
        }
        if ($cloudSkipped -gt $listed) {
            writeIndexerLog "    ほか $($cloudSkipped - $listed) 件"
        }
    }

    $planner.ApplyFailed($retryTargets)
    # 前回、取り込み中に強制終了したファイルは最後に回す（続けて $interruptLimit 回強制終了したファイルは失敗とする）。
    # 最後に回したファイルは、取り込みを始めるまで記録を残す（その前にまた止まっても、回数が分かるように）
    $planner.ApplyInterrupted(@($ledger.ReadIngestingFiles()), $interruptLimit)
    $carried = $planner.Carried
    $ledger.WriteIngestingFiles(@($carried.Keys | ForEach-Object { @{ RelPath = $_; Count = $carried[$_] } }))
    $reporter.Progress(${indexingPhaseCrawl}, 0, $targets.Count, 0, "更新の記録を書き出しています…")
    writeStatusKeepingDropped $ledger $folders $rows.ToArray() $keptDropped $previous
    # インデックスのフォルダごと別の場所・PCへコピーしても元のファイルの場所が分かるよう、インデックス名とクロール対象フォルダの対応を置く
    writeSourceFolderFile $folders

    # インデックス名 → クロール対象フォルダ
    $folderByName = $planner.FolderByName()

    # この回で作る（取り込み対象にした）インデックスの名前。目録に行が無ければ「作成中」で足し、最後まで終えたら「作り終えた」にする
    $roundNames = @($plan | Where-Object { $_.区分 -eq ${planKindIngest} } | ForEach-Object { [string]$_.インデックス名 })
    if (!(markIndexFormatBuilding $roundNames $workspace)) {
        writeIndexerLog "インデックスの目録を書けませんでした（新しい版の tebunko で作られた可能性があります）。" "Yellow"
    }

    if ($targets.Count -eq 0) {
        writeIndexerLog ""
        if ($cloudSkipped -eq 0) {
            writeIndexerLog "取り込みが必要なファイルはありません。（一覧: $(Split-Path $workspace.StatusFile -Leaf)）" "Green"
        } else {
            writeIndexerLog "取り込めるファイルはありません。（一覧: $(Split-Path $workspace.StatusFile -Leaf)）" "Green"
        }
        # 元のファイルが無くなったフォルダは、本文インデックスのファイルから外す
        flushPendingPublish $pending -All
        # 取り込むファイルが無くても、システムインデックスがまだ無いフォルダ（この版に上げた直後など）は作る
        $reporter.Progress(${indexingPhaseFinish}, 0, 0, 0, "システムインデックス（高速検索用）を確かめています…")
        try {
            [void](updateSystemIndexes -shouldStop { $channel.Stop })
        } catch {
            writeIndexerLog "システムインデックスを作れませんでした（次のインデックス作成で作り直します）: $($_.Exception.Message)" "Yellow"
        }
        completeRoundIndexes $roundNames
        $reporter.Progress(${indexingPhaseFinish}, 0, 0, 0, $(if ($cloudSkipped -eq 0) { "更新が必要なファイルはありませんでした" } else { "クラウドにだけあるファイルは、ダウンロードせずに残しました" }))
        removeTmpDir
        return 0
    }

    $total = $targets.Count
    # 取り込みのワーカーの数 = Office を使わずに読む仕事の数（今までの読み取りのスレッドの数の決め方）＋この回に要る Office のアプリの数（最大 3）
    $apps = New-Object string[] $total
    $bookFolders = New-Object string[] $total
    $appNames = @{}
    for ($i = 0; $i -lt $total; $i++) {
        $apps[$i] = getIngestApp $targets[$i].相対パス
        if ($apps[$i]) {
            $appNames[$apps[$i]] = $true
        }
        $bookFolders[$i] = [System.IO.Path]::GetDirectoryName((getBookDir $targets[$i].相対パス))
        $pending.AddPending($bookFolders[$i])
    }
    $readers = getIngestWorkerCount $channel.Workers $total ([int](readSettings).ingestThreads)
    $workerCount = if ($readers -gt 0) { $readers + $appNames.Count } else { 0 }
    writeIndexerLog ""
    if ($workerCount -gt 0) {
        writeIndexerLog "$total 件のファイルを取り込みます。（Office を使わずに読むファイルは $readers 個のスレッド、Excel・Word・PowerPoint のファイルはそれぞれ 1 つのスレッド（Office も 1 つ）で並べて取り込みます。1ファイルの取り込みに ${fileTimeoutMinutes} 分以上かかった場合は、そのファイルを失敗として次のファイルへ進みます）"
    } else {
        writeIndexerLog "$total 件のファイルを取り込みます。（1ファイルの取り込みに ${fileTimeoutMinutes} 分以上かかった場合は、そのファイルを失敗として次のファイルへ進みます）"
    }

    # 取り込んだ件数・利用者のPowerPointのために取り込めなかった件数（失敗の件数に含む。一覧は「未取り込み」のまま次回に回す。終わりの案内で、再取り込みの選択が要る失敗と分ける）・
    # 最後に取り込みのワーカーに渡したファイル（画面に「取り込み中のファイル」として出す）は、結果の受け方（completeIngestJob）と共有する $run に置く。
    # 今回失敗したファイル（$ledger.Failures。@{ RelPath; Message } の並び）・
    # 取り込みの直前に元のファイルが無くなっていた相対パス（$ledger.DroppedRows。取り込み一覧から除く）は StatusLedger が持つ
    # フォルダごとの取り込み中の数と、まだ渡していない数（どちらも 0 になったら本文インデックスのファイルに書き出す）は $pending が持つ
    # 取り込み中のファイル・ワーカーごとの渡してある数と持っている Office のアプリ・並びの先読み・取り込みのワーカー・中止の状態は IngestDispatcher（$dispatcher）が持つ
    # 起動した Office の PID の記録を置く場所（開始時のワークスペースから決め、取り込みのワーカーへ文字列で渡す。途中でワークスペースが変わっても、書く所と消す所が食い違わない）
    $officeRecordDir = getOfficePidDir $workspace
    $dispatcher = [IngestDispatcher]::new($targets, $apps, $bookFolders, $folderByName, $carried, @{
        LookAhead = ${ingestLookAhead}; WorkerScript = ${ingestWorkerScript}
        PostponedReasonInUse = ${postponedReasonInUse}; PostponedReasonRunning = ${postponedReasonRunning}
    })
    $run = @{
        Ledger = $ledger; Pending = $pending; Reporter = $reporter; Dispatcher = $dispatcher; Previous = $previous; Total = $total
        SuccessCount = 0; PostponedCount = 0; CurrentPath = ""
    }

    try {
        if ($workerCount -gt 0) {
            $dispatcher.Start($workerCount, @{
                Lib = (getPartLoad indexerLib)
                WorkDir = $workspace.Dir; TmpDir = ${tmpDir}; TmpDirReason = ${tmpDirReason}; PublishDir = $workspace.PublishDir
                FileTimeoutMinutes = $fileTimeoutMinutes; RestartInterval = $restartInterval; OfficePids = $channel.OfficePids
                UserOfficePids = $channel.UserOfficePids
                OfficeRecordDir = $officeRecordDir; Channel = $channel
            })
        } else {
            $script:officePidSink = $channel.OfficePids
            $script:officeRecordDir = $officeRecordDir
            $script:officeOwnDir = resolveOwnDir ${tmpDir}
            $script:onOfficeHandOver = ${officeHandOverNotice}
            # 司令のスレッドで取り込むときは、利用者が使っている PowerPoint を止めないだけ（諦めるワーカーが無いため Abandon は渡さない）
            $script:officeWatchdog.State.Users = $channel.UserOfficePids
            # 中止を求められたら、取り込み中の Office を止めて打ち切る（制限時間を待たない）
            $script:officeWatchdog.State.Cancel = $channel
            $script:officeWatchdog.Start()
        }

        while ($true) {
            # 制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたワーカーは、取り込み中だったファイルを後回しにする
            # （固まった呼び出しは戻らないことがあるため、結果を待たない。戻ってきた結果は捨てる）
            while ($null -ne ($abandoned = $dispatcher.TakeAbandonedWorker())) {
                foreach ($lost in $abandoned.Entries) {
                    completeIngestJob $run @{ Kind = "Postponed"; Entry = $lost; Reason = ${postponedReasonInUse} }
                }
            }
            # 画面から中止を求められたら、次のファイルを渡さずに終える（残りは「未取り込み」のまま、次回取り込みする）
            if ($dispatcher.StopIfRequested([bool]$channel.Stop)) {
                writeIndexerLog ""
                writeIndexerLog "中止の要求を受けたため、インデックス作成を中止します。（残り $($dispatcher.Remaining) 件は次回取り込みします）" "Yellow"
            }
            # 空いているワーカーに、次のファイルを渡す（並びの先の見方は IngestDispatcher.BeginPass）
            $dispatcher.BeginPass()
            while ($null -ne ($step = $dispatcher.DispatchNext())) {
                if ($step.Kind -eq "Abandoned") {
                    # 諦めたワーカーのアプリのファイルは、渡さずに「未取り込み」のまま残す
                    completeIngestJob $run @{
                        Kind = "Postponed"; Entry = @{ Row = $step.Row; Number = $step.Number; Folder = $step.BookFolder }
                        Reason = ${postponedReasonInUse}; Undispatched = $true
                    }
                } elseif ($step.Kind -eq "Vanished") {
                    # 取り込み対象を調べてから取り込むまでの間に、元のファイルが移動・削除されたもの。
                    # 「失敗」として記録すると、再取り込みを選ぶまで残ってしまうため、一覧・インデックスから除く
                    $pending.Skip($step.BookFolder)
                    writeIndexerLog ("[{0}/{1}] {2}" -f $step.Number, $total, $step.RelPath)
                    writeIndexerLog "    元のファイルが無くなったため、取り込まずに一覧から除きます。（移動・削除・名前変更された）" "Yellow"
                    removeBookDir (getBookDir $step.RelPath)
                    $pending.Add($step.RelPath, $true)
                    $ledger.AddDropped($step.RelPath)
                } else {
                    # 強制終了したときに、どのファイルの取り込み中に止まったか次回分かるよう記録する
                    $pending.Dispatch($step.BookFolder)
                    $ledger.WriteIngestingFiles($dispatcher.IngestingEntries())
                    $run.CurrentPath = $step.RelPath
                    reportIngestProgress $run
                    if ($dispatcher.UsesThreads) {
                        $dispatcher.Submit($step)
                    } else {
                        $dispatcher.SetInlineResult((invokeIngestTask $step.Task $fileTimeoutMinutes $channel))
                    }
                }
            }
            if ($dispatcher.InflightCount() -eq 0) {
                if ($dispatcher.Stopped -or $dispatcher.FolderLost -or $dispatcher.Undispatched -eq 0) {
                    break
                }
                continue
            }

            # 取り込みの結果を 1 つ受け取って記録する
            $received = $dispatcher.Receive([bool]$channel.Stop)
            if ($null -eq $received) {
                # 諦めたワーカーがある・回し直した。ループの先頭で続ける
                continue
            }
            completeIngestJob $run $received
            if ($received.Kind -eq "Completed" -and !$dispatcher.UsesThreads -and ($received.Result.TimedOut -or (($run.SuccessCount + $ledger.Failures.Count) % $restartInterval) -eq 0)) {
                # 制限時間を過ぎて強制終了したアプリは使えないため、すべて終了して次に必要になったときに起動し直す
                stopAllApps
            }
        }
    } finally {
        # 中止・続けられないエラーの場合もここは実行される。
        # 後片付けも数十秒かかることがあるため、何をしているかを画面に伝える（進み具合の数はそのまま残す）
        # 画面は「成功 = 処理済み - 失敗」と出すため、取り込まなかった（元ファイルが無くなった）分は数に入れない
        $processed = $run.SuccessCount + $ledger.Failures.Count
        $reporter.Progress(${indexingPhaseFinish}, $processed, 0, $ledger.Failures.Count, "Officeアプリを終了しています…")
        if ($dispatcher.UsesThreads) {
            $dispatcher.Stop()
        } else {
            $script:officeWatchdog.Stop()
            $script:officeWatchdog.State.Cancel = $null
            stopAllApps
            waitKeptApps ({ [bool]$channel.Stop }.GetNewClosure())
        }
        $script:officePidSink = $null
        $script:officeRecordDir = $null
        $script:officeOwnDir = $null
        $script:onOfficeHandOver = $null
        removeTmpDir
        $ledger.RemoveIngestingFile()
        # 取り込んだ TSV は、中止したときも残さず本文インデックスのファイルに入れる（残すとインデックスの容量が倍になる）
        $reporter.Progress(${indexingPhaseFinish}, $processed, 0, $ledger.Failures.Count, "インデックスをまとめています…")
        flushPendingPublish $pending -All
        $reporter.Progress(${indexingPhaseFinish}, $processed, 0, $ledger.Failures.Count, "更新の記録を書き直しています…")
        # 取り込みの直前に無くなっていたファイルの行は除く（次回の検索でも見つからず、インデックスも削除済み）
        writeStatusKeepingDropped $ledger $folders @($rows | Where-Object { $_ -and !$ledger.DroppedRows.Contains([string]$_.相対パス) }) $keptDropped $previous
        # 初めて取り込んだインデックスは、最初に書き出した時点ではまだフォルダが無いため、ここでもう一度書く
        # （work\content_index\<インデックス名>\source_folder.txt。インデックス 1 個だけをコピーしても元のファイルの場所が分かる）
        writeSourceFolderFile $folders
        # 高速検索用の システムインデックスを作り直す。中止したとき・フォルダが見えなくなったときは作らない
        # （作り直していないフォルダは反映待ちのままのため、検索ではそのフォルダを照合する）
        if (!$dispatcher.Stopped -and !$dispatcher.FolderLost) {
            $reporter.Progress(${indexingPhaseFinish}, $processed, 0, $ledger.Failures.Count, "システムインデックス（高速検索用）を作っています…")
            try {
                [void](updateSystemIndexes -shouldStop { $channel.Stop })
            } catch {
                writeIndexerLog "システムインデックスを作れませんでした（次のインデックス作成で作り直します）: $($_.Exception.Message)" "Yellow"
            }
        }
        # 画面が終わり方（成功・失敗の件数）を読めるよう、進み具合は消さずに最後の状態を残す
        $reporter.Progress(${indexingPhaseFinish}, $processed, $dispatcher.Remaining, $ledger.Failures.Count, "")
    }

    $stopped = $dispatcher.Stopped
    if ($dispatcher.FolderLost) {
        # クロール対象フォルダが見えなくなった場合は、続けられないエラーとして画面に知らせる（残りは未取り込みのまま）
        $message = "クロール対象フォルダが見つからなくなったため、インデックス作成を中止しました: $($dispatcher.FolderLost)" +
            "（残り $($dispatcher.Remaining) 件は取り込まずに残しました。フォルダを使えるようにしてから、もう一度更新してください）"
        writeIndexerLog ""
        writeIndexerLog $message "Red"
        $channel.Error = $message
        return 1
    }

    if (!$stopped) {
        completeRoundIndexes $roundNames
    }
    writeIndexerLog ""
    if ($stopped) {
        writeIndexerLog "インデックス作成を中止しました。（成功: $($run.SuccessCount) 件 / 失敗: $($ledger.Failures.Count) 件）" "Yellow"
    } else {
        writeIndexerLog "インデックス作成が完了しました。（成功: $($run.SuccessCount) 件 / 失敗: $($ledger.Failures.Count) 件）" "Green"
    }
    if ($ledger.DroppedRows.Count -gt 0) {
        writeIndexerLog "取り込みの直前に元のファイルが無くなった $($ledger.DroppedRows.Count) 件は、取り込まずに一覧から除きました。" "Yellow"
    }
    writeIndexerLog "各ファイルの状態・更新日時は $(Split-Path $workspace.StatusFile -Leaf) で確認できます。"
    if ($ledger.Failures.Count -gt 0) {
        # インデックス作成中の表示は流れて見えなくなるため、失敗したファイルと原因を最後にまとめて表示する
        writeIndexerLog ""
        writeIndexerLog "＜取り込みに失敗したファイルと原因＞" "Yellow"
        foreach ($failure in @($ledger.Failures | Select-Object -First $failureListLimit)) {
            writeIndexerLog "  $($failure.RelPath)"
            writeIndexerLog "    → $($failure.Message)" "Red"
        }
        if ($ledger.Failures.Count -gt $failureListLimit) {
            writeIndexerLog "  ほか $($ledger.Failures.Count - $failureListLimit) 件（原因は一覧の「エラー」列）"
        }
        writeIndexerLog ""
        if ($ledger.Failures.Count -gt $run.PostponedCount) {
            writeIndexerLog "失敗したファイルは、画面で「失敗分も再取り込みする」を選んで取り込むと再取り込みします。" "Yellow"
        }
        if ($run.PostponedCount -gt 0) {
            # 一覧の状態は「未取り込み」のため、PowerPoint を閉じてからの次の更新で選ばなくても取り込む
            writeIndexerLog "PowerPoint のために取り込めなかった $($run.PostponedCount) 件は、PowerPoint を閉じてからもう一度更新すると、そのまま取り込みます。" "Yellow"
        }
    }

    if ($stopped) {
        return 2
    }
    return 0
}
