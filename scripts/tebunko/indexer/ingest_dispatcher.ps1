# 取り込むファイルを、取り込みのワーカーに振り分ける（IngestDispatcher）。
# 取り込みのワーカーは 1 種類（どれも STA で、同じスクリプト）。ワーカーごとの列・渡してある数・持っている Office のアプリ・並びの先読み・取り込み中のファイルを持つ。
# 渡し先は判断層の selectIngestWorker が決める。Office の COM は作ったスレッドからしか呼べないため、
# Office のアプリは初めてそのアプリのファイルを受けたワーカーが持ち、そのアプリのファイルはそのワーカーにだけ渡す。
# 取り込みのワーカーとは、今までどおり hashtable と .NET のコレクション（BlockingCollection）で受け渡す。
# 司令（invokeIndexerBody）が、次のように呼ぶ。
#   Start → [StopIfRequested → TakeAbandonedWorker → BeginPass → DispatchNext（Submit）→ Receive] を繰り返す → Stop
# 記録・ログ・進み具合には触らない。DispatchNext・TakeAbandonedWorker・Receive は、司令がそれらに反映するための結果（hashtable）を返す。
# クラスの本体では $script: 変数を直接読まない（docs/design/structure/classes.md「クラスと関数の使い分け」）ため、
# 先読みの件数などは呼び出し元（司令）から options で受け取る。
# 取り込みのワーカーには、インスタンスではなく素の値と hashtable を渡す（スレッドをまたいでインスタンスを渡さない）。
# メソッドの中のローカル変数は、プロパティ名（Pool など）と大文字・小文字だけの違いにしない（$state にする。同名への代入はプロパティへの代入と見なされ、$this. を要求するエラーになる。indexing_reporter.ps1 と同じ）

class IngestDispatcher {
    # 取り込みのワーカーの入れ物 @{ Slots; Results; Settings }。スレッドを使わないときは $null。
    # Slots はワーカーごとの @{ Queue; State; Busy; Owns; Worker }。Worker は、スレッドを始めるまで $null
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
    hidden [string[]]$Apps
    hidden [string[]]$BookFolders
    hidden $FolderByName
    hidden $Carried
    hidden [bool[]]$Dispatched
    hidden [int]$Total
    # まだ渡していない最初のファイル
    hidden [int]$Next
    # 今の回で、次に見るファイルと、見る範囲の終わり
    hidden [int]$Cursor
    hidden [int]$Limit
    # 取り込み中のファイル: 相対パス → @{ Row; Count; Folder; Number; App; Worker; Task }（Worker は渡したワーカーの番号。司令のスレッドで取り込むときは -1）
    hidden $Inflight
    # スレッドを使わないとき、SetInlineResult で預けた結果（Receive が取り出す）
    hidden $InlineResult
    # 直近の Receive が受け取った取り込みの結果（諦めたワーカーが後回しにしたあとで返した結果は入れない。司令が種類ごとの時間を足すために読む）。
    [object]$LastResult

    IngestDispatcher($targets, [string[]]$apps, [string[]]$bookFolders, $folderByName, $carried, [hashtable]$options) {
        # targets: 取り込む行の並び（相対パス列を持つ）。apps・bookFolders: targets と同じ並びの、最初に要る Office のアプリ（要らなければ空）・本文インデックスのフォルダ。
        # folderByName: インデックス名 → クロール対象フォルダ。carried: 前回強制終了したファイル（相対パス → 回数）。
        # options: LookAhead（先読みの件数）、WorkerScript（取り込みのワーカーのスクリプト）、
        #   PostponedReasonInUse・PostponedReasonRunning（後回しの理由）
        #   省略できる TestPath … パスを調べる scriptblock { param($path, $kind) }（$kind は File か Folder。あれば $true）。省略すると本物のファイルシステムを調べる
        foreach ($key in "LookAhead", "WorkerScript", "PostponedReasonInUse", "PostponedReasonRunning") {
            if (!$options.ContainsKey($key)) {
                throw "IngestDispatcher の options に ${key} がありません。"
            }
        }
        $this.Options = $options
        $this.Targets = $targets
        $this.Apps = $apps
        $this.BookFolders = $bookFolders
        $this.FolderByName = $folderByName
        $this.Carried = $carried
        $this.Total = $apps.Length
        $this.Undispatched = $apps.Length
        $this.Dispatched = New-Object bool[] $apps.Length
        $this.FolderLost = ""
        $this.Inflight = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    }

    [void] Start([int]$workerCount, [hashtable]$settings) {
        # 取り込みのワーカーの入れ物を作る。スレッドは、そのワーカーにファイルを初めて渡すときに始める（Submit）。
        # ファイルを渡さないワーカーは、スレッドも Office も作らない
        $slots = New-Object System.Collections.Generic.List[hashtable]
        for ($i = 0; $i -lt $workerCount; $i++) {
            # Abandoned: 制限時間を過ぎても利用者が使っている PowerPoint を止めずに、このワーカーを諦めた（ワーカーの監視が立てる）。
            # Handled: 司令が、取り込み中だったファイルの後回しを済ませた
            $slots.Add(@{
                Queue  = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
                State  = [hashtable]::Synchronized(@{ Abandoned = $false; Handled = $false })
                Busy   = 0
                Owns   = New-Object System.Collections.Generic.List[string]
                Worker = $null
            })
        }
        $this.UsesThreads = $true
        $this.Pool = @{
            Slots    = $slots
            Results  = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
            Settings = $settings
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

    [object] TakeAbandonedWorker() {
        # 諦めたのに、まだ後回しを済ませていないワーカーがあれば、そのワーカーで取り込み中だったファイルを取り込み中から外して返す。
        # 固まった呼び出しは戻らないことがあるため結果を待たない（戻ってきた結果は Receive が捨てる）。
        # @{ Worker（ワーカーの番号）; Entries（取り込み中の記録 @{ Row; Count; Folder; Number; App; Worker; Task } の並び） }。無ければ $null
        if (!$this.Pool) {
            return $null
        }
        $index = $this.unhandledAbandonedWorker()
        if ($index -lt 0) {
            return $null
        }
        $slot = $this.Pool.Slots[$index]
        $slot.State.Handled = $true
        $entries = New-Object System.Collections.Generic.List[object]
        foreach ($relPath in @($this.Inflight.Keys)) {
            $entry = $this.Inflight[$relPath]
            if ($entry.Worker -ne $index) {
                continue
            }
            [void]$this.Inflight.Remove($relPath)
            $slot.Busy = $slot.Busy - 1
            $entries.Add($entry)
        }
        return @{ Worker = $index; Entries = $entries.ToArray() }
    }

    [void] BeginPass() {
        # 並びの先（LookAhead 件まで）も見て、空いているワーカー向けのファイルを先に渡す
        # （Excel のファイルが続いても、Office を使わずに読むファイルを待たせない）。回の始めに呼ぶ
        $this.Cursor = $this.Next
        $this.Limit = [Math]::Min($this.Total, $this.Next + $this.Options.LookAhead)
    }

    [object] DispatchNext() {
        # 次に渡すファイルを 1 つ決め、結果を返す。回の終わりは $null。結果の Kind:
        #   Abandoned  … 諦めたワーカーのアプリのファイル。渡さずに「未取り込み」のまま残す（司令が後回しとして記録する）
        #   Vanished   … 取り込む直前に元のファイルが無くなっていた（司令が一覧・インデックスから除く）
        #   Dispatched … 取り込みに渡す（司令が記録してから Submit する）
        # どれも @{ Kind; Number; Row; RelPath; BookFolder; Worker; Task }（Kind によっては一部だけ）。
        # クロール対象フォルダごと見えなくなったときは FolderLost に入れて回を終える
        while (!$this.Stopped -and !$this.FolderLost -and $this.Cursor -lt $this.Limit) {
            $i = $this.Cursor
            $this.Cursor = $i + 1
            if ($this.Dispatched[$i]) {
                continue
            }
            $app = $this.Apps[$i]
            $row = $this.Targets[$i]
            $bookFolder = $this.BookFolders[$i]
            $worker = -1
            if ($this.Pool) {
                $pick = selectIngestWorker $this.workerViews() $app 2
                if ($pick.Kind -eq "Abandoned") {
                    # 諦めたワーカーのアプリのファイルは、渡さずに「未取り込み」のまま残す
                    $this.Dispatched[$i] = $true
                    $this.Undispatched = $this.Undispatched - 1
                    return @{ Kind = "Abandoned"; Number = $i + 1; Row = $row; BookFolder = $bookFolder }
                }
                if ($pick.Kind -eq "Wait") {
                    continue
                }
                $worker = $pick.Index
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
            if ($this.Pool) {
                # ワーカーは、Office が要らないファイルでは Office を使わずに読み、読めなければ Office が要ると返す
                $task.Office = [bool]$app
            }
            $this.Dispatched[$i] = $true
            $this.Undispatched = $this.Undispatched - 1
            $this.Inflight[$relPath] = @{ Row = $row; Count = $count; Folder = $bookFolder; Number = $i + 1; App = $app; Worker = $worker; Task = $task }
            $this.assign($worker, $app)
            return @{ Kind = "Dispatched"; Number = $i + 1; RelPath = $relPath; BookFolder = $bookFolder; Worker = $worker; Task = $task }
        }
        while ($this.Next -lt $this.Total -and $this.Dispatched[$this.Next]) {
            $this.Next = $this.Next + 1
        }
        return $null
    }

    [void] Submit($dispatched) {
        # DispatchNext が Dispatched で返したファイルを、取り込みのワーカーに渡す（Pool があるとき）
        $this.addTask($dispatched.Worker, $dispatched.Task)
    }

    [void] SetInlineResult($result) {
        # スレッドを使わないとき、司令が自分のスレッドで取り込んだ結果を預ける（次の Receive が返す）
        $this.InlineResult = $result
    }

    [int] InflightCount() {
        return $this.Inflight.Count
    }

    [int] GetInflightCount([string]$relPath) {
        # 取り込み中のファイルを続けて取り込みを始めた回数（今の分を含む）。ジャーナルの「始」に書く。取り込み中でなければ 0
        $entry = $null
        if ($this.Inflight.TryGetValue($relPath, [ref]$entry)) {
            return [int]$entry.Count
        }
        return 0
    }

    [object] Receive() {
        return $this.Receive($false)
    }

    [object] Receive([bool]$stopRequested) {
        # 取り込みの結果を 1 つ受け取り、取り込み中から外して返す。司令が反映することが無いときは $null
        # （諦めたワーカーがある・諦めたワーカーが後回しにしたあとで返した結果を捨てた・回し直しを渡した）。Kind:
        #   Completed       … 取り込んだ（成功・失敗）。@{ Kind; Entry; Result }
        #   Postponed       … 利用者の PowerPoint が起動していて使えなかった。@{ Kind; Entry; Reason }
        #   RerouteAbandoned … 回し直し先のアプリの持ち主が諦めたワーカーだった。渡さずに後回しにする。@{ Kind; Entry; Reason }
        #   Stopped         … 中止を求められて、取り込まなかった・打ち切った（stopRequested のとき、回し直しも回さずにこれにする）。
        #                     @{ Kind; Entry; RelPath; First }。First は、今回初めて止めたとき（残りの件数は Remaining）
        # stopRequested: 画面から中止を求められているか（channel の Stop）
        $this.LastResult = $null
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
            # 諦めたワーカーが、後回しにしたあとで返した結果。捨てる
            return $null
        }
        $this.LastResult = $result
        if ($result.Stopped -or ($result.Reroute -and $stopRequested)) {
            # 中止を求められて、取り込まなかった・打ち切った。失敗にはしない（取り込み一覧は未取り込みのまま、インデックスは変えない）
            [void]$this.Inflight.Remove($result.RelPath)
            $this.release($entry.Worker)
            $first = !$this.Stopped
            if ($first) {
                $this.Stopped = $true
                $this.Remaining = $this.Undispatched
            }
            $this.Remaining = $this.Remaining + 1
            return @{ Kind = "Stopped"; Entry = $entry; RelPath = $result.RelPath; First = $first }
        }
        if ($result.Reroute) {
            # Office を使わずに読めなかった（中身が旧形式・パスワード付き）。Word・PowerPoint の持ち主に回し直す
            # （取り込み中のまま。中止を求められていれば、上で止めた扱いにしているため、ここへは来ない）
            $this.release($entry.Worker)
            $entry.App = getRerouteApp $result.RelPath
            # 回し直しは、渡してある数の上限を超えても渡す（持ち主が取り込み中でも列に足す）
            $pick = selectIngestWorker $this.workerViews() $entry.App 1000
            if ($pick.Kind -ne "Worker") {
                # 回し先の持ち主が諦めたワーカーのときは、渡さずに後回しにする
                [void]$this.Inflight.Remove($result.RelPath)
                return @{ Kind = "RerouteAbandoned"; Entry = $entry; Reason = $this.Options.PostponedReasonInUse }
            }
            $entry.Worker = $pick.Index
            $entry.Task.Office = $true
            $this.assign($entry.Worker, $entry.App)
            $this.addTask($entry.Worker, $entry.Task)
            return $null
        }
        [void]$this.Inflight.Remove($result.RelPath)
        $this.release($entry.Worker)
        if ($result.Postponed) {
            # 諦めたワーカーが返した後回しは「使っている方がいた」、始めから起動していたときは「起動していた」
            $reason = if ($this.Pool -and $this.Pool.Slots[$entry.Worker].State.Abandoned) { $this.Options.PostponedReasonInUse } else { $this.Options.PostponedReasonRunning }
            return @{ Kind = "Postponed"; Entry = $entry; Reason = $reason }
        }
        return @{ Kind = "Completed"; Entry = $entry; Result = $result }
    }

    [void] Stop() {
        # 取り込みのワーカーに終わりを伝え、取り込み中のファイルが終わるのを待ってから片づける（Office を持つワーカーは Office を終了してから終わる）
        $state = $this.Pool
        foreach ($slot in $state.Slots) {
            $slot.Queue.CompleteAdding()
        }
        $abandoned = $false
        foreach ($slot in $state.Slots) {
            $worker = $slot.Worker
            if ($null -eq $worker) {
                continue
            }
            if ($slot.State.Abandoned -and !$worker.Handle.IsCompleted) {
                # 諦めたワーカーは、固まった PowerPoint の呼び出しから戻らないことがある。待たず、止める依頼（BeginStop）だけを出す。
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
            foreach ($slot in $state.Slots) {
                $slot.Queue.Dispose()
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

    hidden [object[]] workerViews() {
        # selectIngestWorker に渡す、ワーカーごとの状態（渡してある数・持っているアプリ・諦めたか）
        $views = New-Object System.Collections.Generic.List[object]
        foreach ($slot in $this.Pool.Slots) {
            $views.Add(@{ Busy = $slot.Busy; Owns = $slot.Owns.ToArray(); Abandoned = [bool]$slot.State.Abandoned })
        }
        return $views.ToArray()
    }

    hidden [void] assign([int]$worker, [string]$app) {
        # ワーカーに渡してある数を 1 つ増やし、そのアプリが初めてなら、そのワーカーを持ち主にする（司令のスレッドで取り込むとき、worker は -1）
        if ($worker -lt 0) {
            return
        }
        $slot = $this.Pool.Slots[$worker]
        $slot.Busy = $slot.Busy + 1
        if ($app -and !$slot.Owns.Contains($app)) {
            $slot.Owns.Add($app)
        }
    }

    hidden [void] release([int]$worker) {
        # ワーカーに渡してある数を 1 つ減らす
        if ($worker -ge 0) {
            $this.Pool.Slots[$worker].Busy = $this.Pool.Slots[$worker].Busy - 1
        }
    }

    hidden [int] unhandledAbandonedWorker() {
        # 諦めた（Abandoned）のに、司令がまだ後回しを済ませていない（Handled でない）ワーカーの番号。無ければ -1。スレッドを始めていないワーカーは含まない
        $slots = $this.Pool.Slots
        for ($i = 0; $i -lt $slots.Count; $i++) {
            if ($null -ne $slots[$i].Worker -and $slots[$i].State.Abandoned -and !$slots[$i].State.Handled) {
                return $i
            }
        }
        return -1
    }

    hidden [void] addTask([int]$index, [hashtable]$task) {
        # ワーカーの列にファイルを渡す。そのワーカーのスレッドがまだ無ければ始める（どのワーカーも STA のスレッド 1 つ）
        $state = $this.Pool
        $slot = $state.Slots[$index]
        if ($null -eq $slot.Worker) {
            $settings = $state.Settings.Clone()
            $settings.WorkerState = $slot.State
            $runspace = [runspacefactory]::CreateRunspace($settings.Lib.State)
            # Office の COM は、作ったスレッドから呼ぶ（STA）
            $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
            $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $runspace.Open()
            $ps = [powershell]::Create()
            $ps.Runspace = $runspace
            $joined = if ($settings.Lib.Prelude) { joinWorkerScript $settings.Lib.Prelude $this.Options.WorkerScript } else { $this.Options.WorkerScript }
            [void]$ps.AddScript($joined).AddArgument($settings).AddArgument($slot.Queue).AddArgument($state.Results).AddArgument($index + 1)
            $slot.Worker = @{ PowerShell = $ps; Runspace = $runspace; Handle = $ps.BeginInvoke(); Ended = $false }
        }
        $slot.Queue.Add($task)
    }

    hidden [object] receiveFromPool() {
        # 取り込みのワーカーの結果を 1 つ受け取る。ワーカーが途中で止まっていたら（読み込めない等）、その理由で例外にする。
        # 諦めたワーカーがあって司令が後回しを済ませていないときは、結果を待たずに $null を返す
        $state = $this.Pool
        $result = $null
        while (!$state.Results.TryTake([ref]$result, 500)) {
            if ($this.unhandledAbandonedWorker() -ge 0) {
                # 制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたワーカーがある。戻らない結果を待たずに、司令に後回しをさせる
                return $null
            }
            foreach ($slot in $state.Slots) {
                $worker = $slot.Worker
                if ($null -eq $worker -or $worker.Ended -or !$worker.Handle.IsCompleted -or $slot.State.Abandoned) {
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
