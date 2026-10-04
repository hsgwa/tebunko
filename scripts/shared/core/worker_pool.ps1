# スレッドのプール（WorkerPool）。ランスペースと PowerShell のインスタンスを使い回し、仕事のたびに作って捨てない。
# 設計は docs/design/structure/threads.md「寿命」、docs/design/structure/classes.md「クラスと関数の使い分け」。
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

function joinWorkerScript {
    # 部品の読み込み（Prelude）を、本体のスクリプト（script）につなぐ。呼び出し側の書き換えをせず、
    # 意味を変えない形でつなぐため、挿入する場所は本体の構文解析で決める:
    #   param ブロックがあればその直後、続けて最初の文が `$ErrorActionPreference = "Stop"` ならそのあとにも進める
    #   （部品が読めないときにスレッドを終える今の動きを保つため。tebunko-pr のテストが確かめる）。
    # Prelude は、同じランスペースで仕事のたびに呼ばれても 1 回だけ実行する（$global:workerPreludeDone で守る。
    # ランスペースを使い回さない呼び出し元でも害はない）
    param (
        [string]$prelude,
        [string]$script
    )

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($script, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) {
        throw "joinWorkerScript: つなぐ本体の構文エラーです: $($errors[0].Message)"
    }
    $insertAt = if ($ast.ParamBlock) { $ast.ParamBlock.Extent.EndOffset } else { 0 }
    $statements = @($ast.EndBlock.Statements)
    if ($statements.Count -gt 0 -and $statements[0].Extent.Text -match '^\$ErrorActionPreference\s*=\s*"Stop"$') {
        $insertAt = $statements[0].Extent.EndOffset
    }
    $head = $script.Substring(0, $insertAt)
    $tail = $script.Substring($insertAt)
    $guarded = "`r`nif (!`$global:workerPreludeDone) {`r`n$prelude`r`n`$global:workerPreludeDone = `$true`r`n}"
    return "$head$guarded$tail"
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
    # 各スレッドで、最初の仕事の前に 1 回だけ実行するスクリプト（lib.ps1 の読み込みなど）。ランスペースの全体（global）で実行する
    [string]$Prelude = ""
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
        $text = "[System.Threading.Thread]::CurrentThread.Priority = '$($this.Priority)'`r`n"
        $body = "& {`r`n$script`r`n} @args"
        if ($this.Prelude) {
            # if の中は新しいスコープにならないため、prelude で読み込んだ関数はランスペースに残り、次の仕事でも使える
            $body = joinWorkerScript $this.Prelude $body
        }
        [void]$ps.AddScript("$text$body")
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

class BackgroundQueue {
    # 画面から頼まれる短い仕事（プレビューの読み込み・状態の読み直し・プロセスの一覧など）を、使い回すスレッドで実行する。
    # 各スレッドでは、最初の仕事の前に 1 回だけ prelude（lib.ps1 の読み込みなど）を実行する（仕事のたびに読み込まない）。
    # 終わった仕事は Poll（画面のタイマー）で受け取り、画面のスレッドで onDone { param($output, $errorText) } を呼ぶ。
    # 画面のスレッドだけから呼ぶ
    hidden [WorkerPool]$Pool
    hidden [System.Collections.Generic.List[hashtable]]$Jobs

    BackgroundQueue([int]$size, [hashtable]$load, [System.Management.Automation.Host.PSHost]$hostUi) {
        # load は getPartLoad の戻り値（@{ State; Prelude }）。呼び出し元が Prelude に続けて
        # 呼びたい文（initWorkspace など）を足したいときは、渡す前に $load.Prelude を書き換える
        $this.Jobs = New-Object 'System.Collections.Generic.List[hashtable]'
        $this.Pool = [WorkerPool]::new($size, $load.State, $hostUi, "Normal")
        $this.Pool.Prelude = $load.Prelude
    }

    [void] Post([string]$script, [object[]]$arguments, [scriptblock]$onDone) {
        # 仕事を 1 つ始める（スレッドが空いていなければ、空くのを待ってから始まる）
        $job = $this.Pool.Submit($script, $arguments)
        $job.OnDone = $onDone
        $this.Jobs.Add($job)
    }

    [int] Poll() {
        # 終わった仕事の onDone を呼び、まだ終わっていない仕事の数を返す
        foreach ($job in $this.Jobs.ToArray()) {
            if (!$job.Handle.IsCompleted) {
                continue
            }
            [void]$this.Jobs.Remove($job)
            $output = $null
            $errorText = $null
            try {
                if ($job.PowerShell.Streams.Error.Count -gt 0) {
                    $errorText = $job.PowerShell.Streams.Error[0].ToString()
                }
                $output = $this.Pool.Receive($job)
            } catch {
                # EndInvoke の呼び出しの例外に包まれているため、仕事が投げた元の例外の文面にする
                $exception = $_.Exception
                while ($exception.InnerException) {
                    $exception = $exception.InnerException
                }
                $errorText = $exception.Message
            }
            if ($job.OnDone) {
                & $job.OnDone $output $errorText
            }
        }
        return $this.Jobs.Count
    }

    [void] Close() {
        # 終わっていない仕事を止め、スレッドを片づける。何度呼んでもよい
        foreach ($job in $this.Jobs.ToArray()) {
            $this.Pool.Cancel($job)
        }
        $this.Jobs.Clear()
        $this.Pool.Close()
    }

    [void] Abandon() {
        # 画面を閉じるときだけに使う。Close と違い、止まった仕事（OS の呼び出しで戻らない届かない共有など）を待たずに戻る。
        # 終わった仕事は Close と同じく片づける（待たされないため）。終わっていない仕事には止める依頼（BeginStop）だけを出し、
        # その PowerShell のインスタンスと、この列のプール（RunspacePool）は Dispose しない
        # （PowerShell.Dispose() は動いている間は中で Stop を呼んで待ち、RunspacePool.Dispose() も止まったランスペースを待つため）。
        # 後始末はプロセスの終わりに任せる
        foreach ($job in $this.Jobs.ToArray()) {
            if ($job.Handle.IsCompleted) {
                try {
                    [void]$this.Pool.Receive($job)
                } catch {
                }
            } else {
                try {
                    [void]$job.PowerShell.BeginStop($null, $null)
                } catch {
                }
            }
        }
        $this.Jobs.Clear()
    }
}