# インデックス作成の本体（invokeIndexer）。画面を使わない起動口（indexer.ps1）と、画面のインデクサのスレッド（IndexingSession）が呼ぶ。
# 設計は docs/design/indexing/parallel.md、docs/design/structure/threads.md「画面とインデクサの受け渡し」。
#
# ・司令のスレッド（invokeIndexer を呼んだスレッド）が、クロール・確認・取り込み一覧・本文インデックスのファイルとシステムインデックスの書き出しを行う。
#   Office の COM には触らない
# ・1 ファイルの取り込み（Office で開いて TSV にする）は、取り込みのスレッド（STA）が行う。スレッドごとに自分の Office を持つ。
#   スレッドの数が 0 のときは、司令のスレッドで取り込む（テストで、途中に割り込むため）
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

# 取り込みのスレッドで動かすスクリプト。自分のレーンの列（tasks）から 1 ファイルずつ取り出して取り込み、結果を results に入れる。
# Office のレーン（Excel・Word・PowerPoint）は STA で、そのアプリを 1 つ持つ。読み取りのレーンは Office を持たない。
# 列が閉じられたら（CompleteAdding）、Office を終了して終わる
${ingestWorkerScript} = {
    param ($settings, $tasks, $results, $number)
    $ErrorActionPreference = "Stop"
    # indexerLib の部品（Prelude）は addIngestTask が、この param ブロックのあとにつなぐ（joinWorkerScript）
    # 置き場所は司令のスレッドと同じにする（読み込み直すと設定から決め直してしまうため）。一時フォルダはスレッドごとに分ける。
    # 部品を読み込んだのと同じスコープ（取り込みのスレッドでは global）に置く
    $own = [Workspace]::new($settings.WorkDir)
    $own.PublishDir = Join-Path $settings.PublishDir "w$number"
    Set-Variable -Name workspace -Value $own
    Set-Variable -Name tmpDir -Value (newWorkerTmpDir $settings.TmpDir $number)
    Set-Variable -Name tmpDirReason -Value $settings.TmpDirReason
    [System.IO.Directory]::CreateDirectory($workspace.PublishDir) | Out-Null
    $script:officePidSink = $settings.OfficePids
    $script:officeRecordDir = $settings.OfficeRecordDir
    $script:officeOwnDir = resolveOwnDir $tmpDir
    $script:onOfficeHandOver = ${officeHandOverNotice}
    $script:officeUnavailable = ($settings.Lane -eq ${laneReader})
    [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::BelowNormal

    # レーンの状態（制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたか）。司令と共有する
    $script:laneState = $settings.LaneState
    if (!$script:officeUnavailable) {
        $script:officeWatchdog.State.Abandon = $settings.LaneState
        if ($settings.UserOfficePids) {
            $script:officeWatchdog.State.Users = $settings.UserOfficePids
        }
        $script:officeWatchdog.Start()
    }
    try {
        runIngestWorker $tasks $results $settings.FileTimeoutMinutes $settings.RestartInterval
    } finally {
        if (!$script:officeUnavailable) {
            $script:officeWatchdog.Stop()
            # 利用者が使っている PowerPoint を諦めたレーンは、終了させない（開いているファイルの保存されていない変更を失わせないため）
            if (!$settings.LaneState.Abandoned) {
                stopAllApps
                # 渡し切れずに持ち続けているアプリは、ここで上限まで仕上げ直す（中止・画面を閉じるときは待たない）
                waitKeptApps ({ [bool]$settings.Channel.Stop }.GetNewClosure())
            }
        }
    }
}

# 司令が並びの先を見る件数（空いているレーン向けのファイルを先に渡すため）。大きくすると、本文インデックスのファイルの書き出しが遅れるフォルダが増える
${ingestLookAhead} = 200

function runIngestWorker {
    # 取り込みのスレッドの繰り返し（ingestWorkerScript が呼ぶ）。列から 1 件ずつ取り込み、結果を results に入れる。
    # Office を持つスレッドは、決まった数を取り込むたびに Office を起動し直す
    param (
        $tasks,
        $results,
        [int]$fileTimeoutMinutes,
        [int]$restartInterval
    )

    $done = 0
    foreach ($task in $tasks.GetConsumingEnumerable()) {
        $result = invokeIngestTask $task $fileTimeoutMinutes
        $results.Add($result)
        if ($result.Postponed) {
            # 後回し（利用者のPowerPointが起動していて使えなかった）は、取り込んだ件数に数えない
            # （件数が変わらないまま後回しが続いたときに、起動し直しを早めないため）
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
    # 読み取りのスレッド（Office を使わずに読むファイル）の数を決める。requested が 0 以上ならその数
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

function getIngestLaneCapacity {
    # レーンに同時に渡しておく数。Office のレーンは取り込み中の 1 件と次の 1 件、読み取りのレーンはスレッドの数の 2 倍
    param (
        [string]$lane,
        [int]$readers
    )

    if ($lane -eq ${laneReader}) {
        return [Math]::Max(1, $readers) * 2
    }
    return 2
}

function testLaneAbandoned {
    # このレーンが、制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたものか（取り込みのスレッドの $script:laneState）
    $state = $script:laneState
    return ($null -ne $state -and [bool]$state.Abandoned)
}

function invokeIngestTask {
    # 1 ファイルを取り込み、結果 @{ RelPath; Ok; Reroute; TsvCount; Message; TimedOut; ExtractVersion; Log } を返す（例外は投げない）。
    # 取り込みのスレッド、またはスレッドの数が 0 のときは司令のスレッドで動く。表示内容は Log に貯め、司令がログに書く
    param (
        [hashtable]$task,
        [int]$fileTimeoutMinutes
    )

    # Reroute: Office を使わずに読めなかった（中身が旧形式・パスワード付き）。司令が Word・PowerPoint のレーンに回し直す
    # Postponed: 利用者のPowerPointが起動していて使えなかった。司令が「未取り込み」のまま次回に回す
    $result = @{ RelPath = $task.RelPath; Ok = $false; Reroute = $false; Postponed = $false; TsvCount = 0; Message = ""; TimedOut = $false; ExtractVersion = ""; Log = "" }
    if (!$tmpDir) {
        # 取り込みの作業フォルダを置けない（selectTmpDir の Brackets・TooLong）ときは、
        # どのファイルも中間 TSV などをこのフォルダに作るため、テキストファイルを含めこのファイルの取り込みをスキップし、取り込みの失敗として記録する
        $result.Message = "$(getTmpDirUnavailableMessage $tmpDirReason)このファイルの取り込みをスキップしました。"
        return $result
    }
    if (testLaneAbandoned) {
        # 利用者が使っている PowerPoint を諦めたレーン。取り込まず、後回しにする（司令は結果を受け取らずに後回しにしている）
        $result.Postponed = $true
        return $result
    }
    $log = New-Object System.IO.StringWriter
    $previousLog = $script:indexerLog
    $script:indexerLog = $log
    try {
        clearTmpDir
        $script:officeWatchdog.State.TimedOut = $false
        $script:officeWatchdog.State.Deadline = (Get-Date).AddMinutes($fileTimeoutMinutes)
        try {
            $result.TsvCount = ingestFile $task.SourcePath
        } finally {
            $script:officeWatchdog.State.Deadline = [datetime]::MaxValue
        }
        if (testLaneAbandoned) {
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
        if (testLaneAbandoned) {
            # 諦めたレーンで、利用者の PowerPoint が閉じられるなどして戻ってきた。失敗にも、アプリの終了にもしない
            $result.Postponed = $true
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
        if (!$script:officeUnavailable -and $appName -and !(testLaneAbandoned)) {
            try { [void](handOverForeignApp $appName) } catch {}
        }
        $result.TimedOut = [bool]$script:officeWatchdog.State.TimedOut
        $script:indexerLog = $previousLog
        $result.Log = $log.ToString()
    }
    return $result
}

function newIngestPool {
    # 取り込みのスレッドの入れ物を作る。スレッドは、そのレーンのファイルを初めて渡すときに始める（addIngestTask）。
    # 渡すファイルが無いレーンは、スレッドも Office も作らない。@{ Queues; Results; Workers; Readers; Settings; Count }
    param (
        [int]$readers,
        [hashtable]$settings
    )

    $queues = @{}
    $laneStates = @{}
    foreach ($lane in ${laneExcel}, ${laneWord}, ${lanePowerPoint}, ${laneReader}) {
        $queues[$lane] = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        # Abandoned: 制限時間を過ぎても利用者が使っている PowerPoint を止めずに、このレーンを諦めた（取り込みのスレッドの監視が立てる）。
        # Handled: 司令が、取り込み中だったファイルの後回しを済ませた
        $laneStates[$lane] = [hashtable]::Synchronized(@{ Abandoned = $false; Handled = $false })
    }
    return @{
        Queues   = $queues
        LaneStates = $laneStates
        Results  = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        Workers  = New-Object System.Collections.Generic.List[hashtable]
        Readers  = $readers
        Settings = $settings
        Started  = @{}
    }
}

function addIngestTask {
    # レーンの列にファイルを渡す。そのレーンのスレッドがまだ無ければ始める
    # （Office のレーンは STA のスレッド 1 つ、読み取りのレーンは MTA のスレッドを読み取りの数だけ）
    param (
        [hashtable]$pool,
        [string]$lane,
        [hashtable]$task
    )

    if (!$pool.Started.ContainsKey($lane)) {
        $pool.Started[$lane] = $true
        $count = if ($lane -eq ${laneReader}) { $pool.Readers } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $settings = $pool.Settings.Clone()
            $settings.Lane = $lane
            $settings.LaneState = $pool.LaneStates[$lane]
            $runspace = [runspacefactory]::CreateRunspace($settings.Lib.State)
            # Office の COM は、作ったスレッドから呼ぶ（STA）。読み取りのスレッドは COM を使わない（MTA）
            $runspace.ApartmentState = if ($lane -eq ${laneReader}) { [System.Threading.ApartmentState]::MTA } else { [System.Threading.ApartmentState]::STA }
            $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $runspace.Open()
            $ps = [powershell]::Create()
            $ps.Runspace = $runspace
            $number = $pool.Workers.Count + 1
            $joined = if ($settings.Lib.Prelude) { joinWorkerScript $settings.Lib.Prelude ${ingestWorkerScript}.ToString() } else { ${ingestWorkerScript}.ToString() }
            [void]$ps.AddScript($joined).AddArgument($settings).AddArgument($pool.Queues[$lane]).AddArgument($pool.Results).AddArgument($number)
            $pool.Workers.Add(@{ PowerShell = $ps; Runspace = $runspace; Handle = $ps.BeginInvoke(); Ended = $false; Lane = $lane })
        }
    }
    $pool.Queues[$lane].Add($task)
}

function getUnhandledAbandonedLanes {
    # 諦めた（Abandoned）のに、司令がまだ後回しを済ませていない（Handled でない）レーンの名前の並び。取り込みのスレッドが無いレーンは含まない
    param (
        [hashtable]$pool
    )

    return @($pool.Started.Keys | Where-Object { $pool.LaneStates[$_].Abandoned -and !$pool.LaneStates[$_].Handled })
}

function receiveIngestResult {
    # 取り込みのスレッドの結果を 1 つ受け取る。スレッドが途中で止まっていたら（読み込めない等）、その理由で例外にする。
    # 諦めたレーンがあって司令が後回しを済ませていないときは、結果を待たずに $null を返す
    param (
        [hashtable]$pool
    )

    $result = $null
    while (!$pool.Results.TryTake([ref]$result, 500)) {
        if (@(getUnhandledAbandonedLanes $pool).Count -gt 0) {
            # 制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたレーンがある。戻らない結果を待たずに、司令に後回しをさせる
            return $null
        }
        foreach ($worker in $pool.Workers) {
            if ($worker.Ended -or !$worker.Handle.IsCompleted -or $pool.LaneStates[$worker.Lane].Abandoned) {
                continue
            }
            $worker.Ended = $true
            $reason = "取り込みのスレッドが止まりました。"
            try {
                [void]$worker.PowerShell.EndInvoke($worker.Handle)
            } catch {
                $exception = $_.Exception
                while ($exception.InnerException) {
                    $exception = $exception.InnerException
                }
                $reason = "取り込みのスレッドが止まりました：$($exception.Message)"
            }
            throw $reason
        }
    }
    return $result
}

function stopIngestWorkers {
    # 取り込みのスレッドに終わりを伝え、取り込み中のファイルが終わるのを待ってから片づける（Office のスレッドは Office を終了してから終わる）
    param (
        [hashtable]$pool
    )

    foreach ($queue in $pool.Queues.Values) {
        $queue.CompleteAdding()
    }
    $abandoned = $false
    foreach ($worker in $pool.Workers) {
        if ($pool.LaneStates[$worker.Lane].Abandoned -and !$worker.Handle.IsCompleted) {
            # 諦めたレーンのスレッドは、固まった PowerPoint の呼び出しから戻らないことがある。待たず、止める依頼（BeginStop）だけを出す。
            # PowerShell・ランスペースと列は Dispose しない（Dispose は動いている間は止まるのを待つ）。後片付けはプロセスの終わりに任せる
            try { [void]$worker.PowerShell.BeginStop($null, $null) } catch {}
            $worker.Ended = $true
            $abandoned = $true
            continue
        }
        if (!$worker.Ended) {
            try {
                [void]$worker.PowerShell.EndInvoke($worker.Handle)
            } catch {
            }
            $worker.Ended = $true
        }
        $worker.PowerShell.Dispose()
        $worker.Runspace.Dispose()
    }
    if (!$abandoned) {
        foreach ($queue in $pool.Queues.Values) {
            $queue.Dispose()
        }
        $pool.Results.Dispose()
    }
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
            $mutex = newAppMutex "indexer" $workspace.Dir
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

function invokeIndexerBody {
    # invokeIndexer の本体（鍵を取り、ログを開いた後）。終了コードを返す。続けられないエラーは例外にする。
    # 取り込み一覧・取り込み中のファイル・今回の失敗と消えた行は StatusLedger（$ledger）、
    # 本文インデックスへの書き出し待ちは PendingPublish（$pending）、進み具合・画面の確認待ちは IndexingReporter（$reporter）にまとめる。
    # 3 つは互いを呼ばず、ここ（司令）がつなぐ（docs/design/structure/classes.md「クラスのつなぎ方」）
    param (
        $channel
    )

    $ledger = [StatusLedger]::new($workspace)
    $pending = [PendingPublish]::new()
    $reporter = [IndexingReporter]::new($channel)

    # 書き出し待ちのフォルダを本文インデックスへまとめる（書き出しは司令であるここが呼ぶ。PendingPublish は決めて返すだけ）。
    # -All は、取り込み中・まだ渡していない数を見ずに、書き出し待ちのフォルダをすべて書き出す（クロールの前後・インデックス作成の終わり）
    function flushPending {
        param ([switch]$All)

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

    # 利用者が使っている PowerPoint のために取り込めなかったファイルを、失敗として記録する（件数は呼び出す側が数える）。
    # 取り込み一覧の行は「未取り込み」のまま、理由を「エラー」列に書く（「失敗」にすると、次の更新で飛ばされるため）。
    # 前回の行があれば、TSV 数・取り込み日時・抽出版は前回のまま残す（インデックスは変えていない）
    function recordUserPowerPointPostponed {
        param ([int]$number, $row, [string]$reason)

        $relPath = [string]$row.相対パス
        $old = $null
        if ($previous.TryGetValue($relPath, [ref]$old) -and $old) {
            $row.TSV数 = $old.TSV数
            $row.取り込み日時 = $old.取り込み日時
            $row.抽出版 = $old.抽出版
        }
        $row.状態 = ${stateNew}
        $row.エラー = $reason
        $ledger.AddFailure($relPath, $reason)
        writeIndexerLog ("[{0}/{1}] {2}" -f $number, $total, $relPath)
        writeIndexerLog "    取り込めませんでした（次の更新で取り込みます）: $reason" "Yellow"
    }
    # 取り込み中のファイルを、取り込み中から外して残す（書き出し待ちを済ませ、失敗として記録する。releaseLane: レーンの取り込み中の数も減らす）
    function postponeInflightForUser {
        param ([string]$relKey, [string]$reason, [switch]$releaseLane)

        $lost = $inflight[$relKey]
        [void]$inflight.Remove($relKey)
        if ($releaseLane) {
            $laneBusy[$lost.Lane] = $laneBusy[$lost.Lane] - 1
        }
        $pending.Complete($lost.Folder)
        recordUserPowerPointPostponed $lost.Number $lost.Row $reason
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
    # content_index\ を作る前に行う（片付けに失敗したら content_index\ を作らずに終わる）
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

    [System.IO.Directory]::CreateDirectory($workspace.IndexDir) | Out-Null
    # 前の版から続けて使うワークスペース・別のドライブへ写したワークスペース・取り込んだワークスペースもここで付く
    # （継ぐかどうかに頼らず、ここで一度だけ content_index 全体に付ける。入れた直後にはそれぞれの場所でも付ける）。
    # 根フォルダに未だ付いていない回（主に初回）は下をすべてたどるため時間がかかるので、先に 1 行残しておく
    writeIndexerLog "content_index を Windows Search の対象から外しています…"
    $nci = setNotContentIndexed $workspace.IndexDir -Recurse
    if (!$nci.Ok) {
        writeIndexerLog "content_index を Windows Search の対象から外せませんでした（$($nci.Reason)）。高速検索が効くまで時間がかかることがあります。" "Yellow"
    }
    $selected = initTmpDir
    $script:tmpDir = $selected.Dir
    $script:tmpDirReason = $selected.Reason
    [System.IO.Directory]::CreateDirectory($workspace.PublishDir) | Out-Null

    # クロール対象フォルダごとにインデックス名（work\content_index 直下のフォルダ名）を決める。前回と同じフォルダは同じ名前を使う
    $status = $ledger.ReadStatus()
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
    $leftover = findIndexFoldersWithBooks $workspace.IndexDir
    if ($leftover.Count -gt 0) {
        writeIndexerLog "集約ファイルに入れていないインデックス（$($leftover.Count) フォルダ）をまとめています…"
        $reporter.Progress(${indexingPhaseCrawl}, 0, 0, 0, "集約ファイルに入れていないインデックスをまとめています…")
        foreach ($folder in $leftover) {
            $pending.MarkFolder($folder)
        }
        flushPending -All
    }

    # 新しく割り当てたインデックス名を設定に保存する
    $planner.SaveNames($targetFolders, { param($assigned) $catalog.StoreAssignedNames($assigned) }.GetNewClosure())
    $folders = $planner.Folders

    writeIndexerLog "出力先フォルダ: $($workspace.IndexDir)"
    writeIndexerLog ""
    writeIndexerLog "クロールしています..."
    # 取り込み一覧の「済」に対してインデックス（TSV）が残っているかを調べるため、今あるTSVの数を数えておく
    # （利用者が work\content_index を直接削除した場合に、「済」のまま検索できなくなるのを防ぐ）
    $reporter.Progress(${indexingPhaseCrawl}, 0, 0, 0, "更新済みのインデックスを確認しています…")
    $indexCounts = getIndexTsvCounts
    if ($null -eq $indexCounts) {
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
    $planner.Crawl($indexCounts, $onCrawlFolder)
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
    if ($channel.ConfirmTargets) {
        $answer = $reporter.WaitForApproval(${indexingPhaseConfirm}, $plan.ToArray(), $targets.Count, $failed.Count, $approvalTimeoutMinutes)
        if ($null -eq $answer) {
            # 取りやめ。取り込み対象にした行は前回の記録のまま（一覧に無かったファイルは記録しない）にする。
            # 取り込み一覧自体は書き直す（無くなったファイルの削除を反映する必要があるため）
            $keep = $planner.KeepForCancel()
            writeIndexerLog ""
            writeIndexerLog "画面で取りやめたため、取り込みません。（取り込み一覧は前回のままです）" "Yellow"
            writeStatusKeepingDropped $ledger $folders $keep $keptDropped $previous
            writeSourceFolderFile $folders
            $reporter.Progress(${indexingPhaseFinish}, 0, 0, 0, "インデックス作成を取りやめました")
            removeTmpDir
            return 2
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

    if ($targets.Count -eq 0) {
        writeIndexerLog ""
        if ($cloudSkipped -eq 0) {
            writeIndexerLog "取り込みが必要なファイルはありません。（一覧: $(Split-Path $workspace.StatusFile -Leaf)）" "Green"
        } else {
            writeIndexerLog "取り込めるファイルはありません。（一覧: $(Split-Path $workspace.StatusFile -Leaf)）" "Green"
        }
        # 元のファイルが無くなったフォルダは、本文インデックスのファイルから外す
        flushPending -All
        # 取り込むファイルが無くても、システムインデックスがまだ無いフォルダ（この版に上げた直後など）は作る
        $reporter.Progress(${indexingPhaseFinish}, 0, 0, 0, "システムインデックス（高速検索用）を確かめています…")
        try {
            [void](updateSystemIndexes -shouldStop { $channel.Stop })
        } catch {
            writeIndexerLog "システムインデックスを作れませんでした（次のインデックス作成で作り直します）: $($_.Exception.Message)" "Yellow"
        }
        $reporter.Progress(${indexingPhaseFinish}, 0, 0, 0, $(if ($cloudSkipped -eq 0) { "更新が必要なファイルはありませんでした" } else { "クラウドにだけあるファイルは、ダウンロードせずに残しました" }))
        removeTmpDir
        return 0
    }

    $total = $targets.Count
    $readers = getIngestWorkerCount $channel.Workers $total ([int](readSettings).ingestThreads)
    writeIndexerLog ""
    if ($readers -gt 0) {
        writeIndexerLog "$total 件のファイルを取り込みます。（Excel・Word・PowerPoint はそれぞれ 1 つずつ、Office を使わずに読むファイルは $readers 個のスレッドで並べて取り込みます。1ファイルの取り込みに ${fileTimeoutMinutes} 分以上かかった場合は、そのファイルを失敗として次のファイルへ進みます）"
    } else {
        writeIndexerLog "$total 件のファイルを取り込みます。（1ファイルの取り込みに ${fileTimeoutMinutes} 分以上かかった場合は、そのファイルを失敗として次のファイルへ進みます）"
    }

    $successCount = 0
    $postponedCount = 0 # 利用者のPowerPointのために取り込めなかった件数（失敗の件数に含む。一覧は「未取り込み」のまま次回に回す。終わりの案内で、再取り込みの選択が要る失敗と分ける）
    $stopped = $false
    # 今回失敗したファイル（$ledger.Failures。@{ RelPath; Message } の並び）・
    # 取り込みの直前に元のファイルが無くなっていた相対パス（$ledger.DroppedRows。取り込み一覧から除く）は StatusLedger が持つ
    $folderLost = ""  # インデックス作成中に見えなくなったクロール対象フォルダ（見つかったら中止する）
    $remaining = 0
    # 取り込み中のファイル: 相対パス → @{ Row; Count; Folder; Number; Lane; Task }
    $inflight = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    # フォルダごとの取り込み中の数と、まだ渡していない数（どちらも 0 になったら本文インデックスのファイルに書き出す）は $pending が持つ
    # レーンごとの取り込み中の数（レーンに渡せる数は ingestLaneCapacity）
    $laneBusy = @{}
    $lanes = New-Object string[] $total
    $bookFolders = New-Object string[] $total
    $dispatched = New-Object bool[] $total
    $undispatched = $total
    for ($i = 0; $i -lt $total; $i++) {
        $lanes[$i] = getIngestLane $targets[$i].相対パス
        $bookFolders[$i] = [System.IO.Path]::GetDirectoryName((getBookDir $targets[$i].相対パス))
        $pending.AddPending($bookFolders[$i])
    }
    # 起動した Office の PID の記録を置く場所（開始時のワークスペースから決め、取り込みのスレッドへ文字列で渡す。途中でワークスペースが変わっても、書く所と消す所が食い違わない）
    $officeRecordDir = getOfficePidDir $workspace
    $pool = $null
    $inlineResult = $null
    $currentPath = ""  # 最後に取り込みのスレッドに渡したファイル（画面に「取り込み中のファイル」として出す）
    $next = 0          # まだ渡していない最初のファイル

    try {
        if ($readers -gt 0) {
            $pool = newIngestPool $readers @{
                Lib = (getPartLoad indexerLib)
                WorkDir = $workspace.Dir; TmpDir = ${tmpDir}; TmpDirReason = ${tmpDirReason}; PublishDir = $workspace.PublishDir
                FileTimeoutMinutes = $fileTimeoutMinutes; RestartInterval = $restartInterval; OfficePids = $channel.OfficePids
                UserOfficePids = $channel.UserOfficePids
                OfficeRecordDir = $officeRecordDir; Channel = $channel
            }
        } else {
            $script:officePidSink = $channel.OfficePids
            $script:officeRecordDir = $officeRecordDir
            $script:officeOwnDir = resolveOwnDir ${tmpDir}
            $script:onOfficeHandOver = ${officeHandOverNotice}
            # 司令のスレッドで取り込むときは、利用者が使っている PowerPoint を止めないだけ（諦めるレーンが無いため Abandon は渡さない）
            $script:officeWatchdog.State.Users = $channel.UserOfficePids
            $script:officeWatchdog.Start()
        }

        while ($true) {
            # 制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたレーンは、取り込み中だったファイルを後回しにする
            # （固まった呼び出しは戻らないことがあるため、結果を待たない。戻ってきた結果は捨てる）
            if ($pool) {
                foreach ($abandonedLane in @(getUnhandledAbandonedLanes $pool)) {
                    $pool.LaneStates[$abandonedLane].Handled = $true
                    foreach ($relKey in @($inflight.Keys)) {
                        if ($inflight[$relKey].Lane -ne $abandonedLane) {
                            continue
                        }
                        postponeInflightForUser $relKey ${postponedReasonInUse} -releaseLane
                        $postponedCount++
                    }
                    $ledger.WriteIngestingFiles((getIngestingEntries $inflight $carried))
                    $reporter.Progress(${indexingPhaseIngest}, ($successCount + $ledger.Failures.Count), ($total - $successCount - $ledger.Failures.Count), $ledger.Failures.Count, $currentPath)
                    flushPending
                }
            }
            # 空いているレーンに、次のファイルを渡す。並びの先（ingestLookAhead 件まで）も見て、空いているレーン向けのものを先に渡す
            # （Excel のファイルが続いても、Office を使わずに読むファイルを待たせない）
            if (!$stopped -and !$folderLost -and $undispatched -gt 0 -and $channel.Stop) {
                # 画面から中止を求められたら、次のファイルを渡さずに終える（残りは「未取り込み」のまま、次回取り込みする）
                $stopped = $true
                $remaining = $undispatched
                writeIndexerLog ""
                writeIndexerLog "中止の要求を受けたため、インデックス作成を中止します。（残り ${remaining} 件は次回取り込みします）" "Yellow"
            }
            $limit = [Math]::Min($total, $next + ${ingestLookAhead})
            for ($i = $next; !$stopped -and !$folderLost -and $i -lt $limit; $i++) {
                if ($dispatched[$i]) {
                    continue
                }
                $lane = $lanes[$i]
                if ($pool -and $pool.LaneStates[$lane].Abandoned) {
                    # 諦めたレーンのファイルは、渡さずに「未取り込み」のまま残す
                    $dispatched[$i] = $true
                    $undispatched--
                    $pending.Skip($bookFolders[$i])
                    $postponedCount++
                    recordUserPowerPointPostponed ($i + 1) $targets[$i] ${postponedReasonInUse}
                    $reporter.Progress(${indexingPhaseIngest}, ($successCount + $ledger.Failures.Count), ($total - $successCount - $ledger.Failures.Count), $ledger.Failures.Count, $currentPath)
                    continue
                }
                if ($pool) {
                    if ([int]$laneBusy[$lane] -ge (getIngestLaneCapacity $lane $readers)) {
                        continue
                    }
                } elseif ($inflight.Count -ge 1) {
                    break
                }
                $row = $targets[$i]
                $relPath = $row.相対パス
                $bookFolder = $bookFolders[$i]
                $parts = splitIndexRelPath $relPath
                $sourceFolder = $folderByName[$parts.Name]
                $sourcePath = Join-Path $sourceFolder $parts.Rest

                # 取り込み対象を調べてから取り込むまでの間に、元のファイルが移動・削除されることがある。
                # 「失敗」として記録すると、再取り込みを選ぶまで残ってしまうため、無くなったファイルは一覧・インデックスから除く
                if (![System.IO.File]::Exists((toLongPath $sourcePath))) {
                    if (!(Test-Path -LiteralPath $sourceFolder -PathType Container)) {
                        # クロール対象フォルダごと見えなくなった（ネットワークの切断・USBメモリの取り外し等）。
                        # 残りのファイルをすべて失敗にしないよう、「未取り込み」のまま中止する（次回、続きから取り込める）
                        $folderLost = $sourceFolder
                        $remaining = $undispatched
                        break
                    }
                    $dispatched[$i] = $true
                    $undispatched--
                    $pending.Skip($bookFolder)
                    writeIndexerLog ("[{0}/{1}] {2}" -f ($i + 1), $total, $relPath)
                    writeIndexerLog "    元のファイルが無くなったため、取り込まずに一覧から除きます。（移動・削除・名前変更された）" "Yellow"
                    removeBookDir (getBookDir $relPath)
                    $pending.Add($relPath, $true)
                    $ledger.AddDropped($relPath)
                    continue
                }

                # 強制終了したときに、どのファイルの取り込み中に止まったか次回分かるよう記録する
                $count = 1
                if ($carried.ContainsKey($relPath)) {
                    $count = $carried[$relPath] + 1
                    [void]$carried.Remove($relPath)
                }
                $task = @{ RelPath = $relPath; SourcePath = $sourcePath }
                $dispatched[$i] = $true
                $undispatched--
                $inflight[$relPath] = @{ Row = $row; Count = $count; Folder = $bookFolder; Number = $i + 1; Lane = $lane; Task = $task }
                $pending.Dispatch($bookFolder)
                $laneBusy[$lane] = [int]$laneBusy[$lane] + 1
                $ledger.WriteIngestingFiles((getIngestingEntries $inflight $carried))
                # 画面はこの 1 行から進み具合を作る（取り込み一覧は読まない）
                $currentPath = $relPath
                $reporter.Progress(${indexingPhaseIngest}, ($successCount + $ledger.Failures.Count), ($total - $successCount - $ledger.Failures.Count), $ledger.Failures.Count, $currentPath)
                if ($pool) {
                    addIngestTask $pool $lane $task
                } else {
                    $inlineResult = invokeIngestTask $task $fileTimeoutMinutes
                }
            }
            while ($next -lt $total -and $dispatched[$next]) {
                $next++
            }
            if ($inflight.Count -eq 0) {
                if ($stopped -or $folderLost -or $undispatched -eq 0) {
                    break
                }
                continue
            }

            # 取り込みの結果を 1 つ受け取って記録する
            if ($pool) {
                $result = receiveIngestResult $pool
            } else {
                $result = $inlineResult
                $inlineResult = $null
            }
            if ($null -eq $result) {
                # 諦めたレーンがある。ループの先頭で後回しにする
                continue
            }
            $entry = $null
            if (!$inflight.TryGetValue($result.RelPath, [ref]$entry)) {
                # 諦めたレーンのスレッドが、後回しにしたあとで返した結果。捨てる
                continue
            }
            if ($result.Reroute) {
                # Office を使わずに読めなかった（中身が旧形式・パスワード付き）。Word・PowerPoint のレーンに回し直す
                # （取り込み中のまま。中止を求められていても最後まで取り込む）
                $laneBusy[$entry.Lane] = $laneBusy[$entry.Lane] - 1
                $entry.Lane = getOfficeLane $result.RelPath
                if ($pool.LaneStates[$entry.Lane].Abandoned) {
                    # 回し先が諦めたレーンのときは、渡さずに後回しにする
                    postponeInflightForUser $result.RelPath ${postponedReasonInUse}
                    $postponedCount++
                    $ledger.WriteIngestingFiles((getIngestingEntries $inflight $carried))
                    flushPending
                    continue
                }
                $laneBusy[$entry.Lane] = [int]$laneBusy[$entry.Lane] + 1
                addIngestTask $pool $entry.Lane $entry.Task
                continue
            }
            if ($result.Postponed) {
                # 利用者のPowerPointが起動していて使えなかった。失敗として理由を見せ、取り込み一覧は「未取り込み」のまま残して次回のインデックス作成で取り込む
                # （インデックスは変えない。$ledger.AddRow・$pending.Add はしない）
                [void]$inflight.Remove($result.RelPath)
                $laneBusy[$entry.Lane] = $laneBusy[$entry.Lane] - 1
                $pending.Complete($entry.Folder)
                $postponedCount++
                # 諦めたレーンのスレッドが返した後回しは「使っている方がいた」、始めから起動していたときは「起動していた」
                $postponedReason = if ($pool -and $pool.LaneStates[$entry.Lane].Abandoned) { ${postponedReasonInUse} } else { ${postponedReasonRunning} }
                recordUserPowerPointPostponed $entry.Number $entry.Row $postponedReason
                $ledger.WriteIngestingFiles((getIngestingEntries $inflight $carried))
                $reporter.Progress(${indexingPhaseIngest}, ($successCount + $ledger.Failures.Count), ($total - $successCount - $ledger.Failures.Count), $ledger.Failures.Count, $currentPath)
                flushPending
                continue
            }
            [void]$inflight.Remove($result.RelPath)
            $laneBusy[$entry.Lane] = $laneBusy[$entry.Lane] - 1
            $row = $entry.Row
            writeIndexerLog ("[{0}/{1}] {2}" -f $entry.Number, $total, $result.RelPath)
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
                $successCount++
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
            $ledger.WriteIngestingFiles((getIngestingEntries $inflight $carried))
            $pending.Complete($entry.Folder)
            $reporter.Progress(${indexingPhaseIngest}, ($successCount + $ledger.Failures.Count), ($total - $successCount - $ledger.Failures.Count), $ledger.Failures.Count, $currentPath)

            if (!$pool -and ($result.TimedOut -or (($successCount + $ledger.Failures.Count) % $restartInterval) -eq 0)) {
                # 制限時間を過ぎて強制終了したアプリは使えないため、すべて終了して次に必要になったときに起動し直す
                stopAllApps
            }
            # 取り込みが揃ったフォルダ（取り込み中が無く、まだ渡していないファイルも無い）を、本文インデックスのファイルに書き出す
            flushPending
        }
    } finally {
        # 中止・続けられないエラーの場合もここは実行される。
        # 後片付けも数十秒かかることがあるため、何をしているかを画面に伝える（進み具合の数はそのまま残す）
        # 画面は「成功 = 処理済み - 失敗」と出すため、取り込まなかった（元ファイルが無くなった）分は数に入れない
        $processed = $successCount + $ledger.Failures.Count
        $reporter.Progress(${indexingPhaseFinish}, $processed, 0, $ledger.Failures.Count, "Officeアプリを終了しています…")
        if ($pool) {
            stopIngestWorkers $pool
        } else {
            $script:officeWatchdog.Stop()
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
        flushPending -All
        $reporter.Progress(${indexingPhaseFinish}, $processed, 0, $ledger.Failures.Count, "更新の記録を書き直しています…")
        # 取り込みの直前に無くなっていたファイルの行は除く（次回の検索でも見つからず、インデックスも削除済み）
        writeStatusKeepingDropped $ledger $folders @($rows | Where-Object { $_ -and !$ledger.DroppedRows.Contains([string]$_.相対パス) }) $keptDropped $previous
        # 初めて取り込んだインデックスは、最初に書き出した時点ではまだフォルダが無いため、ここでもう一度書く
        # （work\content_index\<インデックス名>\source_folder.txt。インデックス 1 個だけをコピーしても元のファイルの場所が分かる）
        writeSourceFolderFile $folders
        # 高速検索用の システムインデックスを作り直す。中止したとき・フォルダが見えなくなったときは作らない
        # （作り直していないフォルダは反映待ちのままのため、検索ではそのフォルダを照合する）
        if (!$stopped -and !$folderLost) {
            $reporter.Progress(${indexingPhaseFinish}, $processed, 0, $ledger.Failures.Count, "システムインデックス（高速検索用）を作っています…")
            try {
                [void](updateSystemIndexes -shouldStop { $channel.Stop })
            } catch {
                writeIndexerLog "システムインデックスを作れませんでした（次のインデックス作成で作り直します）: $($_.Exception.Message)" "Yellow"
            }
        }
        # 画面が終わり方（成功・失敗の件数）を読めるよう、進み具合は消さずに最後の状態を残す
        $reporter.Progress(${indexingPhaseFinish}, $processed, $remaining, $ledger.Failures.Count, "")
    }

    if ($folderLost) {
        # クロール対象フォルダが見えなくなった場合は、続けられないエラーとして画面に知らせる（残りは未取り込みのまま）
        $message = "クロール対象フォルダが見つからなくなったため、インデックス作成を中止しました: ${folderLost}" +
            "（残り ${remaining} 件は取り込まずに残しました。フォルダを使えるようにしてから、もう一度更新してください）"
        writeIndexerLog ""
        writeIndexerLog $message "Red"
        $channel.Error = $message
        return 1
    }

    writeIndexerLog ""
    if ($stopped) {
        writeIndexerLog "インデックス作成を中止しました。（成功: ${successCount} 件 / 失敗: $($ledger.Failures.Count) 件）" "Yellow"
    } else {
        writeIndexerLog "インデックス作成が完了しました。（成功: ${successCount} 件 / 失敗: $($ledger.Failures.Count) 件）" "Green"
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
        if ($ledger.Failures.Count -gt $postponedCount) {
            writeIndexerLog "失敗したファイルは、画面で「失敗分も再取り込みする」を選んで取り込むと再取り込みします。" "Yellow"
        }
        if ($postponedCount -gt 0) {
            # 一覧の状態は「未取り込み」のため、PowerPoint を閉じてからの次の更新で選ばなくても取り込む
            writeIndexerLog "PowerPoint のために取り込めなかった ${postponedCount} 件は、PowerPoint を閉じてからもう一度更新すると、そのまま取り込みます。" "Yellow"
        }
    }

    if ($stopped) {
        return 2
    }
    return 0
}

function getIngestingEntries {
    # 取り込み中のファイルの記録（writeIngestingFiles に渡す形）。取り込み中のものと、最後に回してまだ始めていないもの
    param (
        $inflight,
        $carried
    )

    $entries = New-Object System.Collections.Generic.List[hashtable]
    foreach ($key in $inflight.Keys) {
        $entries.Add(@{ RelPath = $key; Count = $inflight[$key].Count })
    }
    foreach ($key in $carried.Keys) {
        $entries.Add(@{ RelPath = $key; Count = $carried[$key] })
    }
    return , $entries.ToArray()
}
