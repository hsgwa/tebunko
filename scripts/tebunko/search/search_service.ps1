# 検索の司令のスレッド（SearchService）。画面を開いている間 1 つだけ動かし、検索の要求（newSearchRequest）を順に実行する。
# lib.ps1 の読み込みと照合のプールの用意は、スレッドを始めたときに 1 回だけ行う（検索のたびに行わない）。
# 設計は docs/design/structure/threads.md「スレッドの一覧」「寿命」。
#
# SearchService は画面のスレッドだけから呼ぶ（PowerShell 5.1 のクラスのメソッドは、定義したランスペースで動くため）。
# スレッドをまたいで使う要求は、Synchronized の hashtable（newSearchRequest）にする。

# 司令のスレッドで動かすスクリプト。要求の列が閉じられる（Close）まで、要求を 1 つずつ実行する。
# lib の部品（Prelude）は SearchService.Start が、この param ブロックのあとにつなぐ（joinWorkerScript）
${searchServiceScript} = {
    param ($requests, $cache, $workers)
    # 読み込めなければスレッドを終える（画面は IsRunning・GetFailure で知り、次の要求で作り直す）
    $ErrorActionPreference = "Stop"
    initWorkspace
    $ErrorActionPreference = "Continue"
    $pool = if ($workers -gt 1) { newContentIndexWorkerPool $workers } else { $null }
    # 溢れを書いた（かもしれない）tmp の根。ワークスペースを切り替えると、要求の WorkDir が変わるため、起動時の $workspace とは限らない
    $overflowRoots = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    try {
        foreach ($request in $requests.GetConsumingEnumerable()) {
            $requestRoot = try { [Workspace]::new($request.WorkDir).TmpRoot } catch { $null }
            # 前の検索で使った別のワークスペースの溢れは、この検索の始めに消す（今の要求の根は invokeSearchRequest が消す）
            foreach ($root in @($overflowRoots)) {
                if ($root -ne $requestRoot) {
                    try { [void](removeSearchOverflow (getSearchOverflowPidDir $root) $root) } catch { }
                    [void]$overflowRoots.Remove($root)
                }
            }
            if ($requestRoot) { [void]$overflowRoots.Add($requestRoot) }
            invokeSearchRequest $request $pool $cache
            if ($cache) {
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

    SearchService([string]$script, [hashtable]$libLoad, $cache, [int]$workers, [int]$closeMilliseconds) {
        $this.Script = $script
        $this.LibLoad = $libLoad
        $this.Cache = $cache
        $this.Workers = [Math]::Max(1, $workers)
        $this.CloseMilliseconds = $closeMilliseconds
        $this.Start()
    }

    SearchService([string]$script, [hashtable]$libLoad, $cache, [int]$workers, [int]$closeMilliseconds, [bool]$deferred) {
        # deferred が $true のときは、スレッドを始めない（Prepare か最初の Request で始める）
        $this.Script = $script
        $this.LibLoad = $libLoad
        $this.Cache = $cache
        $this.Workers = [Math]::Max(1, $workers)
        $this.CloseMilliseconds = $closeMilliseconds
        if (!$deferred) {
            $this.Start()
        }
    }

    [void] Prepare() {
        # まだ始めていなければ、司令のスレッドを始める（最初の検索の前に済ませておくため。始めてあれば何もしない）
        if ($null -eq $this.PowerShell) {
            $this.Start()
        }
    }

    hidden [void] Start() {
        $this.Requests = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $rs = [runspacefactory]::CreateRunspace($this.LibLoad.State)
        $rs.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        $joined = if ($this.LibLoad.Prelude) { joinWorkerScript $this.LibLoad.Prelude $this.Script } else { $this.Script }
        [void]$ps.AddScript($joined).AddArgument($this.Requests).AddArgument($this.Cache).AddArgument($this.Workers)
        $this.Runspace = $rs
        $this.PowerShell = $ps
        $this.Handle = $ps.BeginInvoke()
    }

    [bool] IsRunning() {
        # 司令のスレッドが動いているか（lib.ps1 を読み込めなかった等で止まっていれば $false）
        return $null -ne $this.PowerShell -and !$this.Handle.IsCompleted
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

    [void] Cancel() {
        # 実行中の要求を取り消す（照合の作業の切れ目で止まる）
        if ($this.Current) {
            $this.Current.Stop = $true
        }
    }

    [void] Close() {
        # 検索を取り消し、司令のスレッドと照合のプールを片づける。何度呼んでもよい
        $this.Cancel()
        $this.Shutdown()
        $this.Current = $null
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
        [int]$workers = (getWorkerCount),
        # スレッドを始めずに作る（画面を出すまでの時間を減らすため。画面が出た後に Prepare で始める）
        [switch]$Deferred
    )

    $libLoad = getPartLoad lib
    return [SearchService]::new(${searchServiceScript}.ToString(), $libLoad, $cache, $workers, ${searchServiceCloseMilliseconds}, [bool]$Deferred)
}
