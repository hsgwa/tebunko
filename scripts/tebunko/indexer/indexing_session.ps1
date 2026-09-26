# 画面のインデックス作成 1 回分（IndexingSession）。インデクサの司令のスレッドを作り、indexer.ps1 を受け渡しの口（newIndexerChannel）付きで実行する。
# 設計は docs/00_共通_4_プロセスとスレッド.md の 7.2・7.5・7.6。
#
# IndexingSession は画面のスレッドだけから呼ぶ（PowerShell 5.1 のクラスのメソッドは、定義したランスペースで動くため）。
# 画面とインデクサのやり取りは、すべて受け渡しの口で行う（中止・確認の返事・進み具合・終了コード）

# インデクサの司令のスレッドで動かすスクリプト。インデックス作成の処理のため、スレッドの優先度を下げる
${indexingSessionScript} = {
    param ($indexerPath, $channel)
    [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::BelowNormal
    & $indexerPath -Channel $channel
}

class IndexingSession {
    # 受け渡しの口
    [hashtable]$Channel
    # 始めた時刻
    [datetime]$Started
    hidden [System.Management.Automation.Runspaces.Runspace]$Runspace
    hidden [powershell]$PowerShell
    hidden [System.IAsyncResult]$Handle

    IndexingSession([string]$script, [string]$indexerPath, [hashtable]$channel) {
        $this.Channel = $channel
        $this.Started = Get-Date
        # 司令のスレッドは COM に触らないため MTA（Office は取り込みのスレッド（STA）が扱う）
        $this.Runspace = [runspacefactory]::CreateRunspace()
        $this.Runspace.ApartmentState = [System.Threading.ApartmentState]::MTA
        $this.Runspace.Open()
        $this.PowerShell = [powershell]::Create()
        $this.PowerShell.Runspace = $this.Runspace
        [void]$this.PowerShell.AddScript($script).AddArgument($indexerPath).AddArgument($channel)
        $this.Handle = $this.PowerShell.BeginInvoke()
    }

    [bool] IsRunning() {
        return $null -ne $this.PowerShell -and !$this.Handle.IsCompleted
    }

    [void] Stop() {
        # 中止を求める（取り込み中のファイルが終わったところで止まる。確認を待っていれば取りやめる）
        $this.Channel.Stop = $true
        [void]$this.Channel.Answered.Set()
    }

    [bool] Wait([int]$milliseconds) {
        # 終わるのを待つ。終わっていれば $true
        if ($null -eq $this.PowerShell) {
            return $true
        }
        return $this.Handle.AsyncWaitHandle.WaitOne($milliseconds)
    }

    [int] GetExitCode() {
        # 終了コード（0 完了 / 1 エラー / 2 中止）。インデクサが終了コードを入れずに止まったとき（読み込めない等）は 1
        if ($null -ne $this.Channel.ExitCode) {
            return [int]$this.Channel.ExitCode
        }
        return 1
    }

    [string] GetError() {
        # 続けられないエラーの内容。インデクサが入れていなければ、スレッドが止まった理由
        if ($this.Channel.Error) {
            return [string]$this.Channel.Error
        }
        $this.Collect()
        return $this.Failure
    }

    hidden [string]$Failure = ""
    hidden [bool]$Ended = $false

    hidden [void] Collect() {
        # 終わったスレッドの結果を受け取り、止まった理由（例外・エラー）を Failure に残す（1 回だけ）
        if ($this.Ended -or $null -eq $this.PowerShell -or !$this.Handle.IsCompleted) {
            return
        }
        $this.Ended = $true
        try {
            [void]$this.PowerShell.EndInvoke($this.Handle)
            if ($this.PowerShell.Streams.Error.Count -gt 0) {
                $this.Failure = $this.PowerShell.Streams.Error[0].ToString()
            }
        } catch {
            $exception = $_.Exception
            while ($exception.InnerException) {
                $exception = $exception.InnerException
            }
            $this.Failure = $exception.Message
        }
    }

    [int] KillOffice() {
        # インデックス作成が起動して PID を記録した Office だけを止め、止めた数を返す（閉じるときに終わらない場合の最後の手段）
        $count = 0
        foreach ($id in @($this.Channel.OfficePids.Keys)) {
            try {
                $process = [System.Diagnostics.Process]::GetProcessById($id)
                # 終了済みで ID が別のプロセスに使われている場合に備え、記録したプロセス名と同じか確かめる
                if ($process.ProcessName -eq $this.Channel.OfficePids[$id]) {
                    $process.Kill()
                    $count++
                }
            } catch {
            }
        }
        return $count
    }

    [void] Close() {
        # スレッドを片づける。終わっていなければ中止を求めて待つ（呼ぶ側が終わるのを待ってから呼ぶ）。何度呼んでもよい
        if ($null -eq $this.PowerShell) {
            return
        }
        if (!$this.Handle.IsCompleted) {
            $this.Stop()
            if (!$this.Wait(5000)) {
                try {
                    $this.PowerShell.Stop()
                } catch {
                }
            }
        }
        $this.Collect()
        $this.PowerShell.Dispose()
        $this.PowerShell = $null
        $this.Runspace.Dispose()
        $this.Runspace = $null
    }
}

function newIndexingSession {
    # インデックス作成を始める（画面の［インデックス作成を開始］）。終わったら Close を呼ぶ
    param (
        [string]$indexerPath,
        [hashtable]$channel
    )

    return [IndexingSession]::new(${indexingSessionScript}.ToString(), $indexerPath, $channel)
}
