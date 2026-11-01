# 検索の司令のスレッド（SearchService）。画面を開いている間 1 つだけ動かし、検索の要求（newSearchRequest）を順に実行する。
# lib.ps1 の読み込みと照合のプールの用意は、スレッドを始めたときに 1 回だけ行う（検索のたびに行わない）。
# 設計は docs/design/structure/threads.md「スレッドの一覧」「寿命」。
#
# SearchService は画面のスレッドだけから呼ぶ（PowerShell 5.1 のクラスのメソッドは、定義したランスペースで動くため）。
# スレッドをまたいで使う要求は、Synchronized の hashtable（newSearchRequest）にする。

# 司令のスレッドで動かすスクリプト。要求の列が閉じられる（Close）まで、要求を 1 つずつ実行する。
# lib の部品（Prelude）は SearchService.Start が、この param ブロックのあとにつなぐ（joinWorkerScript）
${searchServiceScript} = {
    param ($requests, $cache, $workers, $traceInfo = $null)
    # 読み込めなければスレッドを終える（画面は IsRunning・GetFailure で知り、次の要求で作り直す）
    $ErrorActionPreference = "Stop"
    initWorkspace
    $ErrorActionPreference = "Continue"
    # 前の検索の結果（最後まで検索し終えた要求の記録。語が長くなったとき、前のヒットだけを照合し直すのに使う）
    $narrowState = @{ Record = $null }
    $pool = if ($workers -gt 1) { newContentIndexWorkerPool $workers } else { $null }
    # 処理の記録（shared/core/trace.ps1）を、画面と同じ指定・同じ起点で入れる。用意が済んだ時刻は画面のスレッドが読む（GetReadyTicks）
    if ($null -ne $traceInfo) {
        initTrace $traceInfo.Spec $traceInfo.OriginTicks
        $traceInfo.ReadyTicks = [DateTime]::UtcNow.Ticks
    }
    # 溢れを書いた（かもしれない）tmp の根。ワークスペースを切り替えると、要求の WorkDir が変わるため、起動時の $workspace とは限らない
    $overflowRoots = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    try {
        foreach ($request in $requests.GetConsumingEnumerable()) {
            if ($request.Kind -eq "Export") {
                # 結果の書き出し。溢れ（検索の続きの置き場所）には触らず、キャッシュも縮めない
                invokeSearchExport $request
                continue
            }
            $requestRoot = try { [Workspace]::new($request.WorkDir).TmpRoot } catch { $null }
            # 前の検索で使った別のワークスペースの溢れは、この検索の始めに消す（今の要求の根は invokeSearchRequest が消す）
            foreach ($root in @($overflowRoots)) {
                if ($root -ne $requestRoot) {
                    try { [void](removeSearchOverflow (getSearchOverflowPidDir $root) $root) } catch { }
                    [void]$overflowRoots.Remove($root)
                }
            }
            if ($requestRoot) { [void]$overflowRoots.Add($requestRoot) }
            invokeSearchRequest $request $pool $cache $narrowState
            # 絞った要求は何も読まないので、整理しない（直前の全体の検索で使った内容を追い出さないため）
            if ($cache -and !$request.Narrowed) {
                [void](trimTsvTextCache $cache)
            }
        }
    } finally {
        if ($pool) {
            $pool.Close()
        }
        # この画面の検索の溢れ（search_overflow\<PID>）を、使った根のすべてと起動時の根で消す（照合のプールを閉じた後なので、書き手はいない。画面のスレッドでは消さない）
        [void]$overflowRoots.Add($workspace.TmpRoot)
        foreach ($root in @($overflowRoots)) {
            try { [void](removeSearchOverflow (getSearchOverflowPidDir $root) $root) } catch { }
        }
    }
}

# 閉じるときに、実行中の検索が止まるのを待つ時間（ミリ秒）。過ぎたらスレッドを止める
${searchServiceCloseMilliseconds} = 5000

class SearchService {
    # 照合のプールのスレッドの数
    [int]$Workers
    hidden [string]$Script
    hidden [hashtable]$LibLoad
    hidden $Cache
    hidden [int]$CloseMilliseconds
    hidden [System.Collections.Concurrent.BlockingCollection[hashtable]]$Requests
    hidden [System.Management.Automation.Runspaces.Runspace]$Runspace
    hidden [powershell]$PowerShell
    hidden [System.IAsyncResult]$Handle
    # いま実行している（最後に渡した）要求
    hidden [hashtable]$Current
    # いま実行している（最後に渡した）書き出しの要求
    hidden [hashtable]$CurrentExport

    SearchService([string]$script, [hashtable]$libLoad, $cache, [int]$workers, [int]$closeMilliseconds) {
        $this.Script = $script
        $this.LibLoad = $libLoad
        $this.Cache = $cache
        $this.Workers = [Math]::Max(1, $workers)
        $this.CloseMilliseconds = $closeMilliseconds
        $this.Start()
    }

    hidden [void] Start() {
        $this.Requests = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $rs = [runspacefactory]::CreateRunspace($this.LibLoad.State)
        $rs.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        $joined = if ($this.LibLoad.Prelude) { joinWorkerScript $this.LibLoad.Prelude $this.Script } else { $this.Script }
        [void]$ps.AddScript($joined).AddArgument($this.Requests).AddArgument($this.Cache).AddArgument($this.Workers).AddArgument($this.LibLoad.Trace)
        $this.Runspace = $rs
        $this.PowerShell = $ps
        $this.Handle = $ps.BeginInvoke()
    }

    [bool] IsRunning() {
        # 司令のスレッドが動いているか（lib.ps1 を読み込めなかった等で止まっていれば $false）
        return $null -ne $this.PowerShell -and !$this.Handle.IsCompleted
    }

    [long] GetReadyTicks() {
        # 司令のスレッドの用意（lib.ps1 の読み込みと照合のプール）が済んだ時刻（UTC の Ticks）。処理の記録が入っていない・まだ済んでいなければ 0
        $info = $this.LibLoad.Trace
        if ($null -eq $info) {
            return 0
        }
        return [long]$info.ReadyTicks
    }

    [string] GetFailure() {
        # 司令のスレッドが止まっていれば、その理由（無ければ空）
        if ($null -eq $this.PowerShell -or !$this.Handle.IsCompleted) {
            return ""
        }
        if ($this.PowerShell.Streams.Error.Count -gt 0) {
            return $this.PowerShell.Streams.Error[0].ToString()
        }
        return "検索のスレッドが止まりました。"
    }

    [hashtable] Request([hashtable]$request) {
        # 要求を渡す。実行中の要求は取り消す（次の要求が先に始まる）。司令のスレッドが止まっていれば作り直す
        $this.Cancel()
        if (!$this.IsRunning()) {
            $this.Shutdown()
            $this.Start()
        }
        $this.Current = $request
        $this.Requests.Add($request)
        return $request
    }

    [hashtable] Export([hashtable]$request) {
        # 結果の書き出しの要求（newExportRequest）を渡す。実行中の検索は取り消さない（溢れは検索の続きの置き場所なので残す）。
        # 司令のスレッドが止まっていれば作り直す（溢れは前のスレッドの終わりに消えているので、書き出しは表の分だけになる）
        if (!$this.IsRunning()) {
            $this.Shutdown()
            $this.Start()
        }
        $this.CurrentExport = $request
        $this.Requests.Add($request)
        return $request
    }

    [void] Cancel() {
        # 実行中の要求（検索・書き出し）を取り消す（照合の作業の切れ目で止まる）
        if ($this.Current) {
            $this.Current.Stop = $true
        }
        if ($this.CurrentExport) {
            $this.CurrentExport.Stop = $true
        }
    }

    [void] Close() {
        # 検索を取り消し、司令のスレッドと照合のプールを片づける。何度呼んでもよい
        $this.Cancel()
        $this.Shutdown()
        $this.Current = $null
        $this.CurrentExport = $null
    }

    hidden [void] Shutdown() {
        if ($this.Requests) {
            $this.Requests.CompleteAdding()
        }
        if ($this.PowerShell) {
            if (!$this.Handle.AsyncWaitHandle.WaitOne($this.CloseMilliseconds)) {
                try {
                    $this.PowerShell.Stop()
                } catch {
                }
            }
            $this.PowerShell.Dispose()
            $this.PowerShell = $null
        }
        if ($this.Runspace) {
            $this.Runspace.Dispose()
            $this.Runspace = $null
        }
        if ($this.Requests) {
            $this.Requests.Dispose()
            $this.Requests = $null
        }
    }
}

function newSearchService {
    # 検索の司令のスレッドを始める（画面を開いたときに 1 回）。閉じるときは Close を呼ぶ
    param (
        $cache = $null,
        [int]$workers = (getWorkerCount)
    )

    $libLoad = getPartLoad lib
    # 処理の記録を入れているときだけ、司令のスレッドへ指定と起点を渡す
    if (testTraceActive) {
        $libLoad.Trace = [hashtable]::Synchronized(@{ Spec = (getTraceSpec).Spec; OriginTicks = (getTraceSpec).OriginTicks; ReadyTicks = [long]0 })
    }
    return [SearchService]::new(${searchServiceScript}.ToString(), $libLoad, $cache, $workers, ${searchServiceCloseMilliseconds})
}
