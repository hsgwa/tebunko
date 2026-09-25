# スレッドのプール（WorkerPool）。ランスペースと PowerShell のインスタンスを使い回し、仕事のたびに作って捨てない。
# 設計は docs/00_共通_4_プロセスとスレッド.md の 7.3・7.8。
#
# WorkerPool は、作ったランスペースのスレッドだけから呼ぶ（PowerShell 5.1 のクラスのメソッドは、クラスを定義したランスペースで動くため）。
# 仕事のスクリプトは文字列で渡す（スクリプトブロックのまま渡すと、作ったランスペースに結び付いたまま別のスレッドで動いてしまう）。

function newWorkerState {
    # プールの各スレッドに読み込む関数・値だけを入れた InitialSessionState を作る
    # （lib.ps1 全体を読み込むと、スレッドを用意するだけで時間がかかるため）
    param (
        [string[]]$functions = @(),
        [string[]]$variables = @()
    )

    $state = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    foreach ($name in $functions) {
        $state.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($name, (Get-Command $name -CommandType Function).Definition))
    }
    foreach ($name in $variables) {
        $state.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new($name, (Get-Variable $name -ValueOnly), ""))
    }
    return $state
}

function getWorkerCount {
    # プールのスレッドの数の既定（画面のために 1 コアを残す。1〜max）
    param (
        [int]$max = 4,
        [int]$processors = [Environment]::ProcessorCount
    )

    return [Math]::Max(1, [Math]::Min($processors - 1, $max))
}

class WorkerPool {
    # スレッドの数
    [int]$Size
    # 各スレッドの優先度（Normal・BelowNormal など）。仕事を始めるたびに、そのスレッドに設定する
    [string]$Priority
    hidden [System.Management.Automation.Runspaces.RunspacePool]$Pool
    # 使い終わって空いている PowerShell のインスタンス
    hidden [System.Collections.Generic.Stack[powershell]]$Idle

    WorkerPool([int]$size, [System.Management.Automation.Runspaces.InitialSessionState]$state, [System.Management.Automation.Host.PSHost]$hostUi, [string]$priority) {
        $this.Size = [Math]::Max(1, $size)
        $this.Priority = [string][System.Threading.ThreadPriority]$priority
        $this.Idle = New-Object 'System.Collections.Generic.Stack[powershell]'
        # host は省けない（$null は受け付けない）。渡されなければ、このランスペースのものを使う
        if ($null -eq $hostUi) {
            $hostUi = Get-Host
        }
        $this.Pool = [runspacefactory]::CreateRunspacePool(1, $this.Size, $state, $hostUi)
        # 既定では仕事のたびに新しいスレッドを作る。ランスペースごとに 1 つのスレッドを使い続ける
        $this.Pool.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $this.Pool.Open()
    }

    [hashtable] Submit([string]$script, [object[]]$arguments) {
        # 仕事を 1 つ始め、@{ PowerShell; Handle } を返す（Receive・Cancel に渡す）
        if ($null -eq $this.Pool) {
            throw "WorkerPool は閉じています。"
        }
        if ($this.Idle.Count -gt 0) {
            $ps = $this.Idle.Pop()
        } else {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $this.Pool
        }
        # 優先度を設定してから仕事のスクリプトを呼ぶ（AddStatement で文を分けると、引数付きのスクリプトが終わらなくなるため、1 つのスクリプトにする）
        [void]$ps.AddScript("[System.Threading.Thread]::CurrentThread.Priority = '$($this.Priority)'`r`n& {`r`n$script`r`n} @args")
        foreach ($argument in $arguments) {
            [void]$ps.AddArgument($argument)
        }
        return @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
    }

    [object] Receive([hashtable]$job) {
        # 仕事の終わりを待って出力を返す。PowerShell のインスタンスは次の仕事に使い回す。仕事の例外はそのまま投げる
        $ps = $job.PowerShell
        try {
            $output = $ps.EndInvoke($job.Handle)
        } catch {
            $ps.Dispose()
            throw
        }
        $ps.Commands.Clear()
        $ps.Streams.ClearStreams()
        $this.Idle.Push($ps)
        # クラスのメソッドは戻り値を展開しない（PSDataCollection のまま返る）
        return $output
    }

    [void] Cancel([hashtable]$job) {
        # 仕事を止めて捨てる（止めたインスタンスは使い回さない）
        try {
            $job.PowerShell.Stop()
        } catch {
        }
        $job.PowerShell.Dispose()
    }

    [void] Close() {
        # スレッドを止めて片づける。何度呼んでもよい
        while ($this.Idle.Count -gt 0) {
            $this.Idle.Pop().Dispose()
        }
        if ($this.Pool) {
            $this.Pool.Dispose()
            $this.Pool = $null
        }
    }
}
