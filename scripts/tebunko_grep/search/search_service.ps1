# 検索の司令のスレッド（SearchService）。画面を開いている間 1 つだけ動かし、検索の要求（newSearchRequest）を順に実行する。
# lib.ps1 の読み込みと照合のプールの用意は、スレッドを始めたときに 1 回だけ行う（検索のたびに行わない）。
# 設計は docs/00_共通_4_プロセスとスレッド.md の 7.2・7.3。
#
# SearchService は画面のスレッドだけから呼ぶ（PowerShell 5.1 のクラスのメソッドは、定義したランスペースで動くため）。
# スレッドをまたいで使う要求は、Synchronized の hashtable（newSearchRequest）にする。

# 司令のスレッドで動かすスクリプト。要求の列が閉じられる（Close）まで、要求を 1 つずつ実行する
${searchServiceScript} = {
    param ($libPath, $requests, $cache, $workers)
    # 読み込めなければスレッドを終える（画面は IsRunning・GetFailure で知り、次の要求で作り直す）
    $ErrorActionPreference = "Stop"
    . $libPath
    $ErrorActionPreference = "Continue"
    $pool = if ($workers -gt 1) { newPackWorkerPool $workers } else { $null }
    try {
        foreach ($request in $requests.GetConsumingEnumerable()) {
            invokeSearchRequest $request $pool $cache
        }
    } finally {
        if ($pool) {
            $pool.Close()
        }
    }
}

# 閉じるときに、実行中の検索が止まるのを待つ時間（ミリ秒）。過ぎたらスレッドを止める
${searchServiceCloseMilliseconds} = 5000

class SearchService {
    # 照合のプールのスレッドの数
    [int]$Workers
    hidden [string]$Script
    hidden [string]$LibPath
    hidden $Cache
    hidden [int]$CloseMilliseconds
    hidden [System.Collections.Concurrent.BlockingCollection[hashtable]]$Requests
    hidden [powershell]$PowerShell
    hidden [System.IAsyncResult]$Handle
    # いま実行している（最後に渡した）要求
    hidden [hashtable]$Current

    SearchService([string]$script, [string]$libPath, $cache, [int]$workers, [int]$closeMilliseconds) {
        $this.Script = $script
        $this.LibPath = $libPath
        $this.Cache = $cache
        $this.Workers = [Math]::Max(1, $workers)
        $this.CloseMilliseconds = $closeMilliseconds
        $this.Start()
    }

    hidden [void] Start() {
        $this.Requests = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $ps = [powershell]::Create()
        [void]$ps.AddScript($this.Script).AddArgument($this.LibPath).AddArgument($this.Requests).AddArgument($this.Cache).AddArgument($this.Workers)
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
        if ($this.Requests) {
            $this.Requests.Dispose()
            $this.Requests = $null
        }
    }
}

function newSearchService {
    # 検索の司令のスレッドを始める（画面を開いたときに 1 回）。閉じるときは Close を呼ぶ
    param (
        [string]$libPath,
        $cache = $null,
        [int]$workers = (getWorkerCount)
    )

    return [SearchService]::new(${searchServiceScript}.ToString(), $libPath, $cache, $workers, ${searchServiceCloseMilliseconds})
}
