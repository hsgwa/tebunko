# インデックス作成の本体（invokeIndexer）。画面を使わない起動口（indexer.ps1）と、画面のインデクサのスレッド（IndexingSession）が呼ぶ。
# 設計は docs/design/architecture/threads.md「インデックス作成の並列化」「画面とインデクサの受け渡し」。
#
# ・司令のスレッド（invokeIndexer を呼んだスレッド）が、クロール・確認・取り込み一覧・集約ファイルとシステムインデックスの書き出しを行う。
#   Office の COM には触らない
# ・1 ファイルの取り込み（Office で開いて TSV にする）は、取り込みのスレッド（STA）が行う。スレッドごとに自分の Office を持つ。
#   スレッドの数が 0 のときは、司令のスレッドで取り込む（テストで、途中に割り込むため）
# ・画面とのやり取りは受け渡しの口（newIndexerChannel）で行う。表示内容は インデックス作成ログ.txt に書く

# 取り込みのスレッドが読み込む部品（indexer_lib.ps1）
${indexerLibPath} = "$PSScriptRoot\indexer_lib.ps1"

# 取り込みのスレッドで動かすスクリプト。自分のレーンの列（tasks）から 1 ファイルずつ取り出して取り込み、結果を results に入れる。
# Office のレーン（Excel・Word・PowerPoint）は STA で、そのアプリを 1 つ持つ。読み取りのレーンは Office を持たない。
# 列が閉じられたら（CompleteAdding）、Office を終了して終わる
${ingestWorkerScript} = {
    param ($settings, $tasks, $results, $number)
    $ErrorActionPreference = "Stop"
    . $settings.Lib
    # 置き場所は司令のスレッドと同じにする（読み込み直すと設定から決め直してしまうため）。一時フォルダはスレッドごとに分ける。
    # 部品を読み込んだのと同じスコープ（取り込みのスレッドでは global）に置く
    $own = [Workspace]::new($settings.WorkDir)
    $own.PublishDir = Join-Path $settings.PublishDir "w$number"
    Set-Variable -Name workspace -Value $own
    Set-Variable -Name tmpDir -Value (Join-Path $settings.TmpDir "w$number")
    [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
    [System.IO.Directory]::CreateDirectory($workspace.PublishDir) | Out-Null
    $script:officePidSink = $settings.OfficePids
    $script:officeUnavailable = ($settings.Lane -eq ${laneReader})
    [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::BelowNormal

    if (!$script:officeUnavailable) {
        startWatchdog
    }
    try {
        runIngestWorker $tasks $results $settings.FileTimeoutMinutes $settings.RestartInterval
    } finally {
        if (!$script:officeUnavailable) {
            stopWatchdog
            stopAllApps
        }
    }
}

# 司令が並びの先を見る件数（空いているレーン向けのファイルを先に渡すため）。大きくすると、集約ファイルの書き出しが遅れるフォルダが増える
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

function invokeIngestTask {
    # 1 ファイルを取り込み、結果 @{ RelPath; Ok; Reroute; TsvCount; Message; TimedOut; ExtractVersion; Log } を返す（例外は投げない）。
    # 取り込みのスレッド、またはスレッドの数が 0 のときは司令のスレッドで動く。表示内容は Log に貯め、司令がログに書く
    param (
        [hashtable]$task,
        [int]$fileTimeoutMinutes
    )

    # Reroute: Office を使わずに読めなかった（中身が旧形式・パスワード付き）。司令が Word・PowerPoint のレーンに回し直す
    $result = @{ RelPath = $task.RelPath; Ok = $false; Reroute = $false; TsvCount = 0; Message = ""; TimedOut = $false; ExtractVersion = ""; Log = "" }
    $log = New-Object System.IO.StringWriter
    $previousLog = $script:indexerLog
    $script:indexerLog = $log
    try {
        clearTmpDir
        $script:watchdog.TimedOut = $false
        $script:watchdog.Deadline = (Get-Date).AddMinutes($fileTimeoutMinutes)
        try {
            $result.TsvCount = ingestFile $task.SourcePath
        } finally {
            $script:watchdog.Deadline = [datetime]::MaxValue
        }
        publishTsv (getBookDir $task.RelPath)
        $result.ExtractVersion = [string](getExtractVersion $task.RelPath)
        $result.Ok = $true
    } catch {
        $base = $_.Exception.GetBaseException()
        if ($base -is [System.OperationCanceledException] -and $base.Message -eq ${officeRequiredMessage}) {
            $result.Reroute = $true
            return $result
        }
        $message = describeIngestError $_.Exception
        if ($script:watchdog.TimedOut) {
            $message = "${fileTimeoutMinutes} 分以内に取り込みが終わらなかったため中止しました（Officeアプリを強制終了しました）"
        }
        $result.Message = $message
        # アプリが不安定になっている可能性があるため終了する（次に必要になったときに起動し直す）
        if (!$script:officeUnavailable) {
            try { stopApp (getAppName $task.RelPath) } catch {}
        }
    } finally {
        $result.TimedOut = [bool]$script:watchdog.TimedOut
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
    foreach ($lane in ${laneExcel}, ${laneWord}, ${lanePowerPoint}, ${laneReader}) {
        $queues[$lane] = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
    }
    return @{
        Queues   = $queues
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
            $runspace = [runspacefactory]::CreateRunspace()
            # Office の COM は、作ったスレッドから呼ぶ（STA）。読み取りのスレッドは COM を使わない（MTA）
            $runspace.ApartmentState = if ($lane -eq ${laneReader}) { [System.Threading.ApartmentState]::MTA } else { [System.Threading.ApartmentState]::STA }
            $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $runspace.Open()
            $ps = [powershell]::Create()
            $ps.Runspace = $runspace
            $number = $pool.Workers.Count + 1
            [void]$ps.AddScript(${ingestWorkerScript}.ToString()).AddArgument($settings).AddArgument($pool.Queues[$lane]).AddArgument($pool.Results).AddArgument($number)
            $pool.Workers.Add(@{ PowerShell = $ps; Runspace = $runspace; Handle = $ps.BeginInvoke(); Ended = $false; Lane = $lane })
        }
    }
    $pool.Queues[$lane].Add($task)
}

function receiveIngestResult {
    # 取り込みのスレッドの結果を 1 つ受け取る。スレッドが途中で止まっていたら（読み込めない等）、その理由で例外にする
    param (
        [hashtable]$pool
    )

    $result = $null
    while (!$pool.Results.TryTake([ref]$result, 500)) {
        foreach ($worker in $pool.Workers) {
            if ($worker.Ended -or !$worker.Handle.IsCompleted) {
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
    foreach ($worker in $pool.Workers) {
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
    foreach ($queue in $pool.Queues.Values) {
        $queue.Dispose()
    }
    $pool.Results.Dispose()
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
            $writer = New-Object System.IO.StreamWriter($workspace.IndexingLogFile, $false, ${utf8Bom})
            $writer.AutoFlush = $true
            $script:indexerLog = $writer
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
    # invokeIndexer の本体（鍵を取り、ログを開いた後）。終了コードを返す。続けられないエラーは例外にする
    param (
        $channel
    )

    writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "インデックス作成の準備をしています…"

    $restartInterval = 100  # Officeアプリを再起動する間隔（取り込みのスレッドごとのファイル数）。メモリ肥大化対策（再起動 1 回で約 2 秒かかるため、間隔を詰めすぎない）
    $fileTimeoutMinutes = 10  # 1ファイルの取り込みの制限時間（分）。超えたらOfficeアプリを強制終了し、そのファイルは失敗とする
    $approvalTimeoutMinutes = 60  # 画面の返事を待つ制限時間（分）。画面が返事をしない場合に待ち続けないよう打ち切る
    $interruptLimit = 2       # 取り込み中に続けて強制終了した回数がこれに達したファイルは、失敗として以降スキップする
    $failureListLimit = 50    # 終了時に失敗したファイルと原因を表示する最大件数（残りは取り込み一覧で確認する）

    # 集約ファイル（content.<拡張子>.tsv）に書き出す前のフォルダ: フォルダ（フルパス）→ 無くなった元のファイル名の集まり。
    # 取り込んだ TSV は元のファイルごとのフォルダに一時的に置き、同じフォルダの取り込みが終わったらまとめて書き出す
    # （元のファイル 1 つごとに書き出すと、フォルダの大きさ × ファイルの数だけ書き直すことになるため）
    $script:pendingPublish = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)

    $targetFolders = @(getTargetFolders)
    if ($targetFolders.Count -eq 0) {
        throw "クロール対象フォルダがありません。画面でクロール対象フォルダを追加してください。"
    }
    if (@($targetFolders | Where-Object { $_.Enabled }).Count -eq 0) {
        throw "チェックの付いたクロール対象フォルダがありません。画面で取り込むフォルダにチェックを付けてください。"
    }

    [System.IO.Directory]::CreateDirectory($workspace.IndexDir) | Out-Null
    removeStaleTmpDirs
    [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
    [System.IO.Directory]::CreateDirectory($workspace.PublishDir) | Out-Null

    # クロール対象フォルダごとにインデックス名（work\index 直下のフォルダ名）を決める。前回と同じフォルダは同じ名前を使う
    $status = readStatusFile
    $folders = @(assignIndexNames $targetFolders $status.Folders)
    $previous = $status.Rows
    removeDroppedFolders $folders $status.Folders
    # 前回のインデックス作成が途中で止まり、集約ファイルに入れていない TSV（元のファイルごとのフォルダ）が残っていれば、先に入れる
    $leftover = findIndexFoldersWithBooks $workspace.IndexDir
    if ($leftover.Count -gt 0) {
        writeIndexerLog "集約ファイルに入れていないインデックス（$($leftover.Count) フォルダ）をまとめています…"
        writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "集約ファイルに入れていないインデックスをまとめています…"
        foreach ($folder in $leftover) {
            $script:pendingPublish[$folder] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        }
        flushPendingPublish
    }

    # 新しく割り当てたインデックス名を設定に保存する（インデックスの「名前」と「置き場所」を設定で分けて持つため。
    # 名前が設定にあれば、フォルダを移してパスを書き換えても同じインデックスとして扱える）
    if (@($targetFolders | Where-Object { -not $_.Name }).Count -gt 0) {
        # 保存した一覧（読み直して名前だけを足したもの）の名前で続ける。この間に画面で外した・名前を付けられなかったフォルダは、今回は取り込まない
        $saved = @(saveAssignedIndexNames $folders)
        $folders = @($saved | Where-Object {
            $savedPath = $_.Path
            $_.Name -and @($targetFolders | Where-Object { $_.Path -eq $savedPath }).Count -gt 0
        })
        writeIndexerLog "インデックス名を設定に保存しました: $((@($folders | ForEach-Object { $_.Name }) -join '、'))"
    }

    writeIndexerLog "出力先フォルダ: $($workspace.IndexDir)"
    writeIndexerLog ""
    writeIndexerLog "クロールしています..."
    # 取り込み一覧の「済」に対してインデックス（TSV）が残っているかを調べるため、今あるTSVの数を数えておく
    # （利用者が work\index を直接削除した場合に、「済」のまま検索できなくなるのを防ぐ）
    writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "取り込み済みのインデックスを確認しています…"
    $indexCounts = getIndexTsvCounts
    if ($null -eq $indexCounts) {
        writeIndexerLog "  インデックスのフォルダを調べられないため、インデックスが残っているかの確認は行いません。" "Yellow"
    }
    $rows = New-Object System.Collections.Generic.List[object]
    $targets = New-Object System.Collections.Generic.List[object]
    $failed = New-Object System.Collections.Generic.List[object]
    $plan = New-Object System.Collections.Generic.List[object]   # 画面の確認に出す、インデックスごとの件数
    foreach ($folder in $folders) {
        if (-not $folder.Enabled) {
            writeIndexerLog "  [$($folder.Name)] $($folder.Path) … チェックなしのため取り込みません（インデックスはそのまま残します）"
            $plan.Add((newIngestPlanRow $folder.Name $folder.Path ${planKindUnchecked}))
        } elseif (!(Test-Path -LiteralPath $folder.Path -PathType Container)) {
            writeIndexerLog "  [$($folder.Name)] $($folder.Path) … フォルダが見つからないため取り込みません" "Yellow"
            $plan.Add((newIngestPlanRow $folder.Name $folder.Path ${planKindMissing}))
        } else {
            writeIndexerLog "  [$($folder.Name)] $($folder.Path)"
            # 大きいフォルダ・ネットワーク越しでは時間がかかるため、どのフォルダを見ているかを画面に伝える
            writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "[$($folder.Name)] のOfficeファイルを探しています… $($folder.Path)"
            $list = createTargetList $folder $previous $indexCounts
            $rows.AddRange($list.Rows)
            $targets.AddRange($list.Targets)
            $failed.AddRange($list.Failed)
            $plan.Add($list.Plan)
            foreach ($removedPath in $list.Removed) {
                addPendingPublish $removedPath $true
            }
            continue
        }

        # 取り込まないフォルダは前回の結果をそのまま残す
        $prefix = "$($folder.Name)\"
        foreach ($key in @($previous.Keys)) {
            if ($key.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $rows.Add($previous[$key])
            }
        }
    }

    # 画面から始めた場合は、数えた件数を画面に出して、取り込むかどうかの返事を待つ。
    # 「更新不要」かどうかも、この件数を見て画面が知らせる（取り込み対象が 0 件でも、前回失敗の再取り込みを選べる）
    $retryTargets = [bool]$channel.RetryFailed
    if ($channel.ConfirmTargets) {
        $answer = waitForIndexingApproval $channel $plan.ToArray() $targets.Count $failed.Count $approvalTimeoutMinutes
        if ($null -eq $answer) {
            # 取りやめ。1件も取り込んでいないため、取り込み対象にした行は前回の記録のまま（一覧に無かったファイルは記録しない）にする。
            # 「未取り込み」で記録すると、次回［インデックス作成を開始］が［続きから再開］になり、中断したように見えるため。
            # 取り込み一覧自体は書き直す（無くなったファイルの削除を反映する必要があるため）
            $targetPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($row in $targets) {
                [void]$targetPaths.Add($row.相対パス)
            }
            $keep = New-Object System.Collections.Generic.List[object]
            foreach ($row in $rows) {
                if (!$targetPaths.Contains($row.相対パス)) {
                    $keep.Add($row)
                    continue
                }
                $old = $null
                if ($previous.TryGetValue($row.相対パス, [ref]$old)) {
                    $keep.Add($old)
                }
            }
            writeIndexerLog ""
            writeIndexerLog "画面で取りやめたため、取り込みません。（取り込み一覧は前回のままです）" "Yellow"
            writeStatusFile $folders $keep
            writeSourceFolderFile $folders
            writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "インデックス作成を取りやめました"
            removeTmpDir
            return 2
        }
        $retryTargets = $answer.RetryFailed
    }

    if ($failed.Count -gt 0) {
        writeIndexerLog ""
        if ($retryTargets) {
            writeIndexerLog "前回取り込みに失敗し、その後更新されていないファイル $($failed.Count) 件も再取り込みします。"
            $targets.AddRange($failed)
        } else {
            writeIndexerLog "前回取り込みに失敗し、その後更新されていないファイル $($failed.Count) 件はスキップします。（パスワード付きなど）"
        }
    }

    # 前回、取り込み中に強制終了した（ウィンドウを閉じた・PCが停止した等）ファイルは、同じファイルで止まり続けないよう最後に回す。
    # 続けて $interruptLimit 回強制終了したファイルは、応答しなくなるファイルとみなして失敗とする。
    # 最後に回したファイルは、取り込みを始めるまで記録を残す（その前にまた止まっても、回数が分かるように）
    $carried = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in (readIngestingFiles)) {
        $row = @($targets | Where-Object { $_.相対パス -eq $entry.RelPath }) | Select-Object -First 1
        if (!$row) {
            continue
        }
        [void]$targets.Remove($row)
        if ($entry.Count -ge $interruptLimit) {
            $row.状態 = ${stateFailed}
            $row.TSV数 = ""
            $row.取り込み日時 = formatFileTime (Get-Date)
            $row.エラー = "取り込み中に $($entry.Count) 回続けて強制終了されたため、取り込みを中止しました（Officeアプリが応答しなくなる可能性があります）"
            writeIndexerLog ""
            writeIndexerLog "取り込み中に $($entry.Count) 回続けて強制終了したファイルは、失敗としてスキップします: $($row.相対パス)" "Yellow"
        } else {
            $targets.Add($row)
            $carried[$row.相対パス] = $entry.Count
            writeIndexerLog ""
            writeIndexerLog "前回、取り込み中に強制終了したファイルは最後に取り込みます: $($row.相対パス)" "Yellow"
        }
    }
    writeIngestingFiles @($carried.Keys | ForEach-Object { @{ RelPath = $_; Count = $carried[$_] } })
    writeIndexingProgress ${indexingPhaseCrawl} 0 $targets.Count 0 "取り込み一覧を書き出しています…"
    writeStatusFile $folders $rows
    # インデックスのフォルダごと別の場所・PCへコピーしても元のファイルの場所が分かるよう、インデックス名とクロール対象フォルダの対応を置く
    writeSourceFolderFile $folders

    # インデックス名 → クロール対象フォルダ
    $folderByName = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($folder in $folders) {
        $folderByName[$folder.Name] = $folder.Path.TrimEnd("\")
    }

    if ($targets.Count -eq 0) {
        writeIndexerLog ""
        writeIndexerLog "取り込みが必要なファイルはありません。（一覧: $(Split-Path $workspace.StatusFile -Leaf)）" "Green"
        # 元のファイルが無くなったフォルダは、集約ファイルから外す
        flushPendingPublish
        # 取り込むファイルが無くても、システムインデックスがまだ無いフォルダ（この版に上げた直後など）は作る
        writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "システムインデックス（高速検索用）を確かめています…"
        try {
            [void](updateSystemIndexes -shouldStop { $channel.Stop })
        } catch {
            writeIndexerLog "システムインデックスを作れませんでした（次のインデックス作成で作り直します）: $($_.Exception.Message)" "Yellow"
        }
        writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "取り込みが必要なファイルはありませんでした"
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
    $stopped = $false
    $failures = New-Object System.Collections.Generic.List[object]  # 今回失敗したファイル: @{ RelPath; Message }
    # 取り込みの直前に元のファイルが無くなっていたファイルの相対パス。取り込み一覧から除く
    $droppedRows = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $folderLost = ""  # インデックス作成中に見えなくなったクロール対象フォルダ（見つかったら中止する）
    $remaining = 0
    # 取り込み中のファイル: 相対パス → @{ Row; Count; Folder; Number; Lane; Task }
    $inflight = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    # フォルダごとの取り込み中の数と、まだ渡していない数（どちらも 0 になったら集約ファイルに書き出す）
    $folderBusy = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    $folderPending = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    # レーンごとの取り込み中の数（レーンに渡せる数は ingestLaneCapacity）
    $laneBusy = @{}
    $lanes = New-Object string[] $total
    $bookFolders = New-Object string[] $total
    $dispatched = New-Object bool[] $total
    $undispatched = $total
    for ($i = 0; $i -lt $total; $i++) {
        $lanes[$i] = getIngestLane $targets[$i].相対パス
        $bookFolders[$i] = [System.IO.Path]::GetDirectoryName((getBookDir $targets[$i].相対パス))
        $folderPending[$bookFolders[$i]] = [int]$folderPending[$bookFolders[$i]] + 1
    }
    $pool = $null
    $inlineResult = $null
    $currentPath = ""  # 最後に取り込みのスレッドに渡したファイル（画面に「取り込み中のファイル」として出す）
    $next = 0          # まだ渡していない最初のファイル

    try {
        if ($readers -gt 0) {
            $pool = newIngestPool $readers @{
                Lib = ${indexerLibPath}
                WorkDir = $workspace.Dir; TmpDir = ${tmpDir}; PublishDir = $workspace.PublishDir
                FileTimeoutMinutes = $fileTimeoutMinutes; RestartInterval = $restartInterval; OfficePids = $channel.OfficePids
            }
        } else {
            $script:officePidSink = $channel.OfficePids
            startWatchdog
        }

        while ($true) {
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
                    $folderPending[$bookFolder] = $folderPending[$bookFolder] - 1
                    writeIndexerLog ("[{0}/{1}] {2}" -f ($i + 1), $total, $relPath)
                    writeIndexerLog "    元のファイルが無くなったため、取り込まずに一覧から除きます。（移動・削除・名前変更された）" "Yellow"
                    removeBookDir (getBookDir $relPath)
                    addPendingPublish $relPath $true
                    [void]$droppedRows.Add($relPath)
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
                $folderPending[$bookFolder] = $folderPending[$bookFolder] - 1
                $inflight[$relPath] = @{ Row = $row; Count = $count; Folder = $bookFolder; Number = $i + 1; Lane = $lane; Task = $task }
                $folderBusy[$bookFolder] = [int]$folderBusy[$bookFolder] + 1
                $laneBusy[$lane] = [int]$laneBusy[$lane] + 1
                writeIngestingFiles (getIngestingEntries $inflight $carried)
                # 画面はこの 1 行から進み具合を作る（取り込み一覧は読まない）
                $currentPath = $relPath
                writeIndexingProgress ${indexingPhaseIngest} ($successCount + $failures.Count) ($total - $successCount - $failures.Count) $failures.Count $currentPath
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
            $entry = $inflight[$result.RelPath]
            if ($result.Reroute) {
                # Office を使わずに読めなかった（中身が旧形式・パスワード付き）。Word・PowerPoint のレーンに回し直す
                # （取り込み中のまま。中止を求められていても最後まで取り込む）
                $laneBusy[$entry.Lane] = $laneBusy[$entry.Lane] - 1
                $entry.Lane = getOfficeLane $result.RelPath
                $laneBusy[$entry.Lane] = [int]$laneBusy[$entry.Lane] + 1
                addIngestTask $pool $entry.Lane $entry.Task
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
                addPendingPublish $result.RelPath
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
                $failures.Add(@{ RelPath = $result.RelPath; Message = $result.Message })
            }

            # 中断しても結果が残るよう、1件ごとに取り込み一覧へ追記する（最後に1ファイル1行にまとめ直す）
            $row.取り込み日時 = formatFileTime (Get-Date)
            addStatusRow $row
            writeIngestingFiles (getIngestingEntries $inflight $carried)
            $folderBusy[$entry.Folder] = $folderBusy[$entry.Folder] - 1
            writeIndexingProgress ${indexingPhaseIngest} ($successCount + $failures.Count) ($total - $successCount - $failures.Count) $failures.Count $currentPath

            if (!$pool -and ($result.TimedOut -or (($successCount + $failures.Count) % $restartInterval) -eq 0)) {
                # 制限時間を過ぎて強制終了したアプリは使えないため、すべて終了して次に必要になったときに起動し直す
                stopAllApps
            }
            # 取り込みが揃ったフォルダ（取り込み中が無く、まだ渡していないファイルも無い）を、集約ファイルに書き出す
            flushPendingPublish (@($folderBusy.Keys | Where-Object { $folderBusy[$_] -gt 0 }) + @($folderPending.Keys | Where-Object { $folderPending[$_] -gt 0 }))
        }
    } finally {
        # 中止・続けられないエラーの場合もここは実行される。
        # 後片付けも数十秒かかることがあるため、何をしているかを画面に伝える（進み具合の数はそのまま残す）
        # 画面は「成功 = 処理済み - 失敗」と出すため、取り込まなかった（元ファイルが無くなった）分は数に入れない
        $processed = $successCount + $failures.Count
        writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "Officeアプリを終了しています…"
        if ($pool) {
            stopIngestWorkers $pool
        } else {
            stopWatchdog
            stopAllApps
        }
        $script:officePidSink = $null
        removeTmpDir
        removeIngestingFile
        # 取り込んだ TSV は、中止したときも残さず集約ファイルに入れる（残すとインデックスの容量が倍になる）
        writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "インデックスをまとめています…"
        flushPendingPublish
        writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "取り込み一覧を書き直しています…"
        # 取り込みの直前に無くなっていたファイルの行は除く（次回の検索でも見つからず、インデックスも削除済み）
        writeStatusFile $folders @($rows | Where-Object { $_ -and !$droppedRows.Contains([string]$_.相対パス) })
        # 初めて取り込んだインデックスは、最初に書き出した時点ではまだフォルダが無いため、ここでもう一度書く
        # （work\index\<インデックス名>\元のフォルダ.txt。インデックス 1 個だけをコピーしても元のファイルの場所が分かる）
        writeSourceFolderFile $folders
        # 高速検索用の システムインデックスを作り直す。中止したとき・フォルダが見えなくなったときは作らない
        # （作り直していないフォルダは反映待ちのままのため、検索ではそのフォルダを照合する）
        if (!$stopped -and !$folderLost) {
            writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "システムインデックス（高速検索用）を作っています…"
            try {
                [void](updateSystemIndexes -shouldStop { $channel.Stop })
            } catch {
                writeIndexerLog "システムインデックスを作れませんでした（次のインデックス作成で作り直します）: $($_.Exception.Message)" "Yellow"
            }
        }
        # 画面が終わり方（成功・失敗の件数）を読めるよう、進み具合は消さずに最後の状態を残す
        writeIndexingProgress ${indexingPhaseFinish} $processed $remaining $failures.Count ""
    }

    if ($folderLost) {
        # クロール対象フォルダが見えなくなった場合は、続けられないエラーとして画面に知らせる（残りは未取り込みのまま）
        $message = "クロール対象フォルダが見つからなくなったため、インデックス作成を中止しました: ${folderLost}" +
            "（残り ${remaining} 件は未取り込みのまま残しました。フォルダを使えるようにしてから、もう一度取り込んでください）"
        writeIndexerLog ""
        writeIndexerLog $message "Red"
        $channel.Error = $message
        return 1
    }

    writeIndexerLog ""
    if ($stopped) {
        writeIndexerLog "インデックス作成を中止しました。（成功: ${successCount} 件 / 失敗: $($failures.Count) 件）" "Yellow"
    } else {
        writeIndexerLog "インデックス作成が完了しました。（成功: ${successCount} 件 / 失敗: $($failures.Count) 件）" "Green"
    }
    if ($droppedRows.Count -gt 0) {
        writeIndexerLog "取り込みの直前に元のファイルが無くなった $($droppedRows.Count) 件は、取り込まずに一覧から除きました。" "Yellow"
    }
    writeIndexerLog "各ファイルの状態・更新日時は $(Split-Path $workspace.StatusFile -Leaf) で確認できます。"
    if ($failures.Count -gt 0) {
        # インデックス作成中の表示は流れて見えなくなるため、失敗したファイルと原因を最後にまとめて表示する
        writeIndexerLog ""
        writeIndexerLog "＜取り込みに失敗したファイルと原因＞" "Yellow"
        foreach ($failure in @($failures | Select-Object -First $failureListLimit)) {
            writeIndexerLog "  $($failure.RelPath)"
            writeIndexerLog "    → $($failure.Message)" "Red"
        }
        if ($failures.Count -gt $failureListLimit) {
            writeIndexerLog "  ほか $($failures.Count - $failureListLimit) 件（一覧の「状態」が「${stateFailed}」の行。原因は「エラー」列）"
        }
        writeIndexerLog ""
        writeIndexerLog "失敗したファイルは、画面で「失敗分も再取り込みする」を選んで取り込むと再取り込みします。" "Yellow"
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

function addPendingPublish {
    # 取り込んだ・無くなった元のファイルのフォルダを、書き出し待ちにする
    param (
        [string]$relPath,
        [bool]$removed = $false
    )

    $folder = [System.IO.Path]::GetDirectoryName((getBookDir $relPath))
    if (!$script:pendingPublish.ContainsKey($folder)) {
        $script:pendingPublish[$folder] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }
    if ($removed) {
        [void]$script:pendingPublish[$folder].Add([System.IO.Path]::GetFileName($relPath))
    }
}

function flushPendingPublish {
    # 書き出し待ちのフォルダを、集約ファイル・システムインデックスに書き出す。keepFolders は、まだ取り込みが続くため除く
    # （取り込み中のファイルがあるフォルダと、まだ取り込みのスレッドに渡していないファイルがあるフォルダ）。
    # 書き出せなかったフォルダは TSV が残るため、次のインデックス作成の始めに書き出す
    param (
        [string[]]$keepFolders = @()
    )

    $keep = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($folder in @($keepFolders)) {
        if ($folder) {
            [void]$keep.Add($folder)
        }
    }
    $flush = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($folder in @($script:pendingPublish.Keys)) {
        if (!$keep.Contains($folder)) {
            $flush[$folder] = $script:pendingPublish[$folder]
            [void]$script:pendingPublish.Remove($folder)
        }
    }
    if ($flush.Count -eq 0) {
        return
    }
    try {
        [void](publishIndexFolders $flush)
    } catch {
        writeIndexerLog "    インデックスをまとめられませんでした（次のインデックス作成でまとめ直します）: $($_.Exception.Message)" "Yellow"
    }
}
