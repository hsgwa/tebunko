# 取り込むファイルを、レーンごとの取り込みのスレッドに振り分ける（IngestDispatcher）。
# 取り込みのスレッドの入れ物（プール）・レーンごとの取り込み中の数・並びの先読み・取り込み中のファイルを持つ。
# 取り込みのスレッドとは、今までどおり hashtable と .NET のコレクション（BlockingCollection）で受け渡す。
# 司令（invokeIndexerBody）が、次のように呼ぶ。
#   Start → [StopIfRequested → TakeAbandonedLane → BeginPass → DispatchNext（Submit）→ Receive] を繰り返す → Stop
# 記録・ログ・進み具合には触らない。DispatchNext・TakeAbandonedLane・Receive は、司令がそれらに反映するための結果（hashtable）を返す。
# クラスの本体では $script: 変数を直接読まない（docs/design/structure/classes.md「クラスと関数の使い分け」）ため、
# レーンの名前などは呼び出し元（司令）から options で受け取る。
# 取り込みのスレッドには、インスタンスではなく素の値と hashtable を渡す（スレッドをまたいでインスタンスを渡さない）。
# メソッドの中のローカル変数は、プロパティ名（Pool など）と大文字・小文字だけの違いにしない（$state にする。同名への代入はプロパティへの代入と見なされ、$this. を要求するエラーになる。indexing_reporter.ps1 と同じ）

class IngestDispatcher {
    # 取り込みのスレッドの入れ物 @{ Queues; LaneStates; Results; Workers; Readers; Settings; Started }。スレッドを使わないときは $null
    [object]$Pool
    # 取り込みのスレッドを使うか（Start のあと $true。使わないときは $false）
    [bool]$UsesThreads
    # 中止を受けて振り分けをやめた
    [bool]$Stopped
    # インデックス作成中に見えなくなったクロール対象フォルダ（見つかったら振り分けをやめる）。無ければ空
    [string]$FolderLost
    # 中止・フォルダが見えなくなったことで振り分けなかった件数
    [int]$Remaining
    # まだ渡していない件数
    [int]$Undispatched

    hidden [hashtable]$Options
    hidden $Targets
    hidden [string[]]$Lanes
    hidden [string[]]$BookFolders
    hidden $FolderByName
    hidden $Carried
    hidden [bool[]]$Dispatched
    hidden [int]$Total
    hidden [int]$Readers
    # まだ渡していない最初のファイル
    hidden [int]$Next
    # 今の回で、次に見るファイルと、見る範囲の終わり
    hidden [int]$Cursor
    hidden [int]$Limit
    # 取り込み中のファイル: 相対パス → @{ Row; Count; Folder; Number; Lane; Task }
    hidden $Inflight
    # レーンごとの取り込み中の数
    hidden [hashtable]$LaneBusy
    # スレッドを使わないとき、SetInlineResult で預けた結果（Receive が取り出す）
    hidden $InlineResult

    IngestDispatcher($targets, [string[]]$lanes, [string[]]$bookFolders, $folderByName, $carried, [hashtable]$options) {
        # targets: 取り込む行の並び（相対パス列を持つ）。lanes・bookFolders: targets と同じ並びのレーン名・本文インデックスのフォルダ。
        # folderByName: インデックス名 → クロール対象フォルダ。carried: 前回強制終了したファイル（相対パス → 回数）。
        # options: LaneExcel・LaneWord・LanePowerPoint・LaneReader（レーン名）、LookAhead（先読みの件数）、WorkerScript（取り込みのスレッドのスクリプト）、
        #   PostponedReasonInUse・PostponedReasonRunning（後回しの理由）
        #   省略できる TestPath … パスを調べる scriptblock { param($path, $kind) }（$kind は File か Folder。あれば $true）。省略すると本物のファイルシステムを調べる
        foreach ($key in "LaneExcel", "LaneWord", "LanePowerPoint", "LaneReader", "LookAhead", "WorkerScript", "PostponedReasonInUse", "PostponedReasonRunning") {
            if (!$options.ContainsKey($key)) {
                throw "IngestDispatcher の options に ${key} がありません。"
            }
        }
        $this.Options = $options
        $this.Targets = $targets
        $this.Lanes = $lanes
        $this.BookFolders = $bookFolders
        $this.FolderByName = $folderByName
        $this.Carried = $carried
        $this.Total = $lanes.Length
        $this.Undispatched = $lanes.Length
        $this.Dispatched = New-Object bool[] $lanes.Length
        $this.FolderLost = ""
        $this.Inflight = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        $this.LaneBusy = @{}
    }

    [void] Start([int]$readers, [hashtable]$settings) {
        # 取り込みのスレッドの入れ物を作る。スレッドは、そのレーンのファイルを初めて渡すときに始める（Submit）。
        # 渡すファイルが無いレーンは、スレッドも Office も作らない
        $this.Readers = $readers
        $queues = @{}
        $laneStates = @{}
        foreach ($lane in $this.Options.LaneExcel, $this.Options.LaneWord, $this.Options.LanePowerPoint, $this.Options.LaneReader) {
            $queues[$lane] = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
            # Abandoned: 制限時間を過ぎても利用者が使っている PowerPoint を止めずに、このレーンを諦めた（取り込みのスレッドの監視が立てる）。
            # Handled: 司令が、取り込み中だったファイルの後回しを済ませた
            $laneStates[$lane] = [hashtable]::Synchronized(@{ Abandoned = $false; Handled = $false })
        }
        $this.UsesThreads = $true
        $this.Pool = @{
            Queues     = $queues
            LaneStates = $laneStates
            Results    = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
            Workers    = New-Object System.Collections.Generic.List[hashtable]
            Readers    = $readers
            Settings   = $settings
            Started    = @{}
        }
    }

    [bool] StopIfRequested([bool]$requested) {
        # 画面から中止を求められたら、次のファイルを渡さずに終える（残りは「未取り込み」のまま、次回取り込みする）。
        # 今回初めて止めたときだけ $true を返す（残りの件数は Remaining）
        if (!$this.Stopped -and !$this.FolderLost -and $this.Undispatched -gt 0 -and $requested) {
            $this.Stopped = $true
            $this.Remaining = $this.Undispatched
            return $true
        }
        return $false
    }

    [object] TakeAbandonedLane() {
        # 諦めたのに、まだ後回しを済ませていないレーンがあれば、そのレーンで取り込み中だったファイルを取り込み中から外して返す。
        # 固まった呼び出しは戻らないことがあるため結果を待たない（戻ってきた結果は Receive が捨てる）。
        # @{ Lane; Entries（取り込み中の記録 @{ Row; Count; Folder; Number; Lane; Task } の並び） }。無ければ $null
        if (!$this.Pool) {
            return $null
        }
        $lane = $null
        foreach ($name in $this.unhandledAbandonedLanes()) {
            $lane = $name
            break
        }
        if ($null -eq $lane) {
            return $null
        }
        $this.Pool.LaneStates[$lane].Handled = $true
        $entries = New-Object System.Collections.Generic.List[object]
        foreach ($relPath in @($this.Inflight.Keys)) {
            $entry = $this.Inflight[$relPath]
            if ($entry.Lane -ne $lane) {
                continue
            }
            [void]$this.Inflight.Remove($relPath)
            $this.LaneBusy[$lane] = [int]$this.LaneBusy[$lane] - 1
            $entries.Add($entry)
        }
        return @{ Lane = $lane; Entries = $entries.ToArray() }
    }

    [void] BeginPass() {
        # 並びの先（LookAhead 件まで）も見て、空いているレーン向けのファイルを先に渡す
        # （Excel のファイルが続いても、Office を使わずに読むファイルを待たせない）。回の始めに呼ぶ
        $this.Cursor = $this.Next
        $this.Limit = [Math]::Min($this.Total, $this.Next + $this.Options.LookAhead)
    }

    [object] DispatchNext() {
        # 次に渡すファイルを 1 つ決め、結果を返す。回の終わりは $null。結果の Kind:
        #   Abandoned  … 諦めたレーンのファイル。渡さずに「未取り込み」のまま残す（司令が後回しとして記録する）
        #   Vanished   … 取り込む直前に元のファイルが無くなっていた（司令が一覧・インデックスから除く）
        #   Dispatched … 取り込みに渡す（司令が記録してから Submit する）
        # どれも @{ Kind; Number; Row; RelPath; BookFolder; Lane; Task }（Kind によっては一部だけ）。
        # クロール対象フォルダごと見えなくなったときは FolderLost に入れて回を終える
        while (!$this.Stopped -and !$this.FolderLost -and $this.Cursor -lt $this.Limit) {
            $i = $this.Cursor
            $this.Cursor = $i + 1
            if ($this.Dispatched[$i]) {
                continue
            }
            $lane = $this.Lanes[$i]
            $row = $this.Targets[$i]
            $bookFolder = $this.BookFolders[$i]
            if ($this.Pool -and $this.Pool.LaneStates[$lane].Abandoned) {
                # 諦めたレーンのファイルは、渡さずに「未取り込み」のまま残す
                $this.Dispatched[$i] = $true
                $this.Undispatched = $this.Undispatched - 1
                return @{ Kind = "Abandoned"; Number = $i + 1; Row = $row; BookFolder = $bookFolder }
            }
            if ($this.Pool) {
                if ([int]$this.LaneBusy[$lane] -ge (getIngestLaneCapacity $lane $this.Readers)) {
                    continue
                }
            } elseif ($this.Inflight.Count -ge 1) {
                break
            }
            $relPath = $row.相対パス
            $parts = splitIndexRelPath $relPath
            $sourceFolder = $this.FolderByName[$parts.Name]
            $sourcePath = Join-Path $sourceFolder $parts.Rest

            # 取り込み対象を調べてから取り込むまでの間に、元のファイルが移動・削除されることがある。
            # 「失敗」として記録すると、再取り込みを選ぶまで残ってしまうため、無くなったファイルは一覧・インデックスから除く
            if (!$this.pathExists($sourcePath, "File")) {
                if (!$this.pathExists($sourceFolder, "Folder")) {
                    # クロール対象フォルダごと見えなくなった（ネットワークの切断・USBメモリの取り外し等）。
                    # 残りのファイルをすべて失敗にしないよう、「未取り込み」のまま中止する（次回、続きから取り込める）
                    $this.FolderLost = $sourceFolder
                    $this.Remaining = $this.Undispatched
                    break
                }
                $this.Dispatched[$i] = $true
                $this.Undispatched = $this.Undispatched - 1
                return @{ Kind = "Vanished"; Number = $i + 1; RelPath = $relPath; BookFolder = $bookFolder }
            }

            # 強制終了したときに、どのファイルの取り込み中に止まったか次回分かるよう、回数を数える（記録は司令が書く）
            $count = 1
            if ($this.Carried.ContainsKey($relPath)) {
                $count = $this.Carried[$relPath] + 1
                [void]$this.Carried.Remove($relPath)
            }
            $task = @{ RelPath = $relPath; SourcePath = $sourcePath }
            $this.Dispatched[$i] = $true
            $this.Undispatched = $this.Undispatched - 1
            $this.Inflight[$relPath] = @{ Row = $row; Count = $count; Folder = $bookFolder; Number = $i + 1; Lane = $lane; Task = $task }
            $this.LaneBusy[$lane] = [int]$this.LaneBusy[$lane] + 1
            return @{ Kind = "Dispatched"; Number = $i + 1; RelPath = $relPath; BookFolder = $bookFolder; Lane = $lane; Task = $task }
        }
        while ($this.Next -lt $this.Total -and $this.Dispatched[$this.Next]) {
            $this.Next = $this.Next + 1
        }
        return $null
    }

    [void] Submit($dispatched) {
        # DispatchNext が Dispatched で返したファイルを、取り込みのスレッドに渡す（Pool があるとき）
        $this.addTask($dispatched.Lane, $dispatched.Task)
    }

    [void] SetInlineResult($result) {
        # スレッドを使わないとき、司令が自分のスレッドで取り込んだ結果を預ける（次の Receive が返す）
        $this.InlineResult = $result
    }

    [int] InflightCount() {
        return $this.Inflight.Count
    }

    [object[]] IngestingEntries() {
        # 取り込み中のファイルの記録（StatusLedger.WriteIngestingFiles に渡す形）。取り込み中のものと、最後に回してまだ始めていないもの
        $entries = New-Object System.Collections.Generic.List[object]
        foreach ($key in $this.Inflight.Keys) {
            $entries.Add(@{ RelPath = $key; Count = $this.Inflight[$key].Count })
        }
        foreach ($key in $this.Carried.Keys) {
            $entries.Add(@{ RelPath = $key; Count = $this.Carried[$key] })
        }
        return $entries.ToArray()
    }

    [object] Receive() {
        # 取り込みの結果を 1 つ受け取り、取り込み中から外して返す。司令が反映することが無いときは $null
        # （諦めたレーンがある・諦めたレーンのスレッドが後回しにしたあとで返した結果を捨てた・回し直しを渡した）。Kind:
        #   Completed       … 取り込んだ（成功・失敗）。@{ Kind; Entry; Result }
        #   Postponed       … 利用者の PowerPoint が起動していて使えなかった。@{ Kind; Entry; Reason }
        #   RerouteAbandoned … 回し直し先のレーンが諦めたレーンだった。渡さずに後回しにする。@{ Kind; Entry; Reason }
        $result = $null
        if ($this.Pool) {
            $result = $this.receiveFromPool()
        } else {
            $result = $this.InlineResult
            $this.InlineResult = $null
        }
        if ($null -eq $result) {
            return $null
        }
        $entry = $null
        if (!$this.Inflight.TryGetValue($result.RelPath, [ref]$entry)) {
            # 諦めたレーンのスレッドが、後回しにしたあとで返した結果。捨てる
            return $null
        }
        if ($result.Reroute) {
            # Office を使わずに読めなかった（中身が旧形式・パスワード付き）。Word・PowerPoint のレーンに回し直す
            # （取り込み中のまま。中止を求められていても最後まで取り込む）
            $this.LaneBusy[$entry.Lane] = $this.LaneBusy[$entry.Lane] - 1
            $entry.Lane = getOfficeLane $result.RelPath
            if ($this.Pool.LaneStates[$entry.Lane].Abandoned) {
                # 回し先が諦めたレーンのときは、渡さずに後回しにする
                [void]$this.Inflight.Remove($result.RelPath)
                return @{ Kind = "RerouteAbandoned"; Entry = $entry; Reason = $this.Options.PostponedReasonInUse }
            }
            $this.LaneBusy[$entry.Lane] = [int]$this.LaneBusy[$entry.Lane] + 1
            $this.addTask($entry.Lane, $entry.Task)
            return $null
        }
        [void]$this.Inflight.Remove($result.RelPath)
        $this.LaneBusy[$entry.Lane] = $this.LaneBusy[$entry.Lane] - 1
        if ($result.Postponed) {
            # 諦めたレーンのスレッドが返した後回しは「使っている方がいた」、始めから起動していたときは「起動していた」
            $reason = if ($this.Pool -and $this.Pool.LaneStates[$entry.Lane].Abandoned) { $this.Options.PostponedReasonInUse } else { $this.Options.PostponedReasonRunning }
            return @{ Kind = "Postponed"; Entry = $entry; Reason = $reason }
        }
        return @{ Kind = "Completed"; Entry = $entry; Result = $result }
    }

    [void] Stop() {
        # 取り込みのスレッドに終わりを伝え、取り込み中のファイルが終わるのを待ってから片づける（Office のスレッドは Office を終了してから終わる）
        $state = $this.Pool
        foreach ($queue in $state.Queues.Values) {
            $queue.CompleteAdding()
        }
        $abandoned = $false
        foreach ($worker in $state.Workers) {
            if ($state.LaneStates[$worker.Lane].Abandoned -and !$worker.Handle.IsCompleted) {
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
            foreach ($queue in $state.Queues.Values) {
                $queue.Dispose()
            }
            $state.Results.Dispose()
        }
    }

    hidden [bool] pathExists([string]$path, [string]$kind) {
        # 元のファイル・フォルダがあるか。options の TestPath があれば、それで調べる（差し替え用）
        if ($this.Options.ContainsKey("TestPath")) {
            return [bool](& $this.Options.TestPath $path $kind)
        }
        if ($kind -eq "File") {
            return [System.IO.File]::Exists((toLongPath $path))
        }
        return [bool](Test-Path -LiteralPath $path -PathType Container)
    }

    hidden [object[]] unhandledAbandonedLanes() {
        # 諦めた（Abandoned）のに、司令がまだ後回しを済ませていない（Handled でない）レーンの名前の並び。取り込みのスレッドが無いレーンは含まない
        $state = $this.Pool
        return @($state.Started.Keys | Where-Object { $state.LaneStates[$_].Abandoned -and !$state.LaneStates[$_].Handled })
    }

    hidden [void] addTask([string]$lane, [hashtable]$task) {
        # レーンの列にファイルを渡す。そのレーンのスレッドがまだ無ければ始める
        # （Office のレーンは STA のスレッド 1 つ、読み取りのレーンは MTA のスレッドを読み取りの数だけ）
        $state = $this.Pool
        $isReader = ($lane -eq $this.Options.LaneReader)
        if (!$state.Started.ContainsKey($lane)) {
            $state.Started[$lane] = $true
            $count = if ($isReader) { $state.Readers } else { 1 }
            for ($i = 0; $i -lt $count; $i++) {
                $settings = $state.Settings.Clone()
                $settings.Lane = $lane
                $settings.LaneState = $state.LaneStates[$lane]
                $runspace = [runspacefactory]::CreateRunspace($settings.Lib.State)
                # Office の COM は、作ったスレッドから呼ぶ（STA）。読み取りのスレッドは COM を使わない（MTA）
                $runspace.ApartmentState = if ($isReader) { [System.Threading.ApartmentState]::MTA } else { [System.Threading.ApartmentState]::STA }
                $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
                $runspace.Open()
                $ps = [powershell]::Create()
                $ps.Runspace = $runspace
                $number = $state.Workers.Count + 1
                $joined = if ($settings.Lib.Prelude) { joinWorkerScript $settings.Lib.Prelude $this.Options.WorkerScript } else { $this.Options.WorkerScript }
                [void]$ps.AddScript($joined).AddArgument($settings).AddArgument($state.Queues[$lane]).AddArgument($state.Results).AddArgument($number)
                $state.Workers.Add(@{ PowerShell = $ps; Runspace = $runspace; Handle = $ps.BeginInvoke(); Ended = $false; Lane = $lane })
            }
        }
        $state.Queues[$lane].Add($task)
    }

    hidden [object] receiveFromPool() {
        # 取り込みのスレッドの結果を 1 つ受け取る。スレッドが途中で止まっていたら（読み込めない等）、その理由で例外にする。
        # 諦めたレーンがあって司令が後回しを済ませていないときは、結果を待たずに $null を返す
        $state = $this.Pool
        $result = $null
        while (!$state.Results.TryTake([ref]$result, 500)) {
            if (@($this.unhandledAbandonedLanes()).Count -gt 0) {
                # 制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたレーンがある。戻らない結果を待たずに、司令に後回しをさせる
                return $null
            }
            foreach ($worker in $state.Workers) {
                if ($worker.Ended -or !$worker.Handle.IsCompleted -or $state.LaneStates[$worker.Lane].Abandoned) {
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
}
