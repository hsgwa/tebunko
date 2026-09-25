# スレッドのプール（shared\core\worker_pool.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "getWorkerCount" -Tag Unit {
    It "画面のために 1 コアを残し、1〜上限に収める" {
        getWorkerCount 4 8 | Should Be 4
        getWorkerCount 4 3 | Should Be 2
        getWorkerCount 4 1 | Should Be 1
        getWorkerCount 2 16 | Should Be 2
    }
}

Describe "newWorkerState" -Tag Unit {
    It "指定した関数と値だけを各スレッドに読み込む" {
        function workerPoolTestDouble($x) { $x * 2 }
        $workerPoolTestValue = 5
        $state = newWorkerState @("workerPoolTestDouble") @("workerPoolTestValue")
        @($state.Commands | Where-Object { $_.Name -eq "workerPoolTestDouble" }).Count | Should Be 1
        @($state.Variables | Where-Object { $_.Name -eq "workerPoolTestValue" })[0].Value | Should Be 5
    }
}

Describe "WorkerPool" -Tag Unit {
    BeforeEach {
        function workerPoolTestDouble($x) { $x * 2 }
        $script:pool = [WorkerPool]::new(2, (newWorkerState @("workerPoolTestDouble")), $null, "BelowNormal")
    }
    AfterEach {
        $script:pool.Close()
    }

    It "仕事の出力を返す（関数と引数を使える）" {
        $job = $script:pool.Submit('param ($a, $b) workerPoolTestDouble ($a + $b)', @(20, 1))
        @($script:pool.Receive($job)) | Should Be 42
    }

    It "配列・hashtable・null の引数をそのまま渡す" {
        $job = $script:pool.Submit('param ($list, $table, $none) "$($list.Count) $($table.K) $($null -eq $none)"', @(@(1, 2, 3), @{ K = "v" }, $null))
        @($script:pool.Receive($job)) | Should Be "3 v True"
    }

    It "仕事のスレッドを指定した優先度にする" {
        $job = $script:pool.Submit('[System.Threading.Thread]::CurrentThread.Priority.ToString()', @())
        @($script:pool.Receive($job)) | Should Be "BelowNormal"
    }

    It "PowerShell のインスタンスを使い回す" {
        $first = $script:pool.Submit('1', @())
        [void]$script:pool.Receive($first)
        $second = $script:pool.Submit('2', @())
        [void]$script:pool.Receive($second)
        [object]::ReferenceEquals($first.PowerShell, $second.PowerShell) | Should Be $true
    }

    It "同時に複数の仕事を進める" {
        $jobs = @(1..4 | ForEach-Object { $script:pool.Submit('param ($n) Start-Sleep -Milliseconds 50; $n', @($_)) })
        @($jobs | ForEach-Object { @($script:pool.Receive($_))[0] }) -join "," | Should Be "1,2,3,4"
    }

    It "仕事の例外を投げ、次の仕事は続けられる" {
        $job = $script:pool.Submit('throw "失敗"', @())
        { $script:pool.Receive($job) } | Should Throw "失敗"
        @($script:pool.Receive($script:pool.Submit('"続き"', @()))) | Should Be "続き"
    }

    It "止めた仕事は使い回さない" {
        $job = $script:pool.Submit('Start-Sleep -Seconds 30', @())
        $script:pool.Cancel($job)
        $next = $script:pool.Submit('"次"', @())
        [object]::ReferenceEquals($job.PowerShell, $next.PowerShell) | Should Be $false
        @($script:pool.Receive($next)) | Should Be "次"
    }

    It "閉じた後は仕事を受け付けない。Close は何度呼んでもよい" {
        $script:pool.Close()
        $script:pool.Close()
        { $script:pool.Submit('1', @()) } | Should Throw "閉じています"
    }
}

Describe "WorkerPool（Prelude）" -Tag Unit {
    It "各スレッドで最初の仕事の前に 1 回だけ実行し、定義した関数を次の仕事でも使える" {
        $pool = [WorkerPool]::new(1, (newWorkerState), $null, "Normal")
        try {
            $pool.Prelude = 'function preludeDefined { "定義済み" }; $global:preludeCount = [int]$global:preludeCount + 1'
            @($pool.Receive($pool.Submit('preludeDefined', @()))) | Should Be "定義済み"
            @($pool.Receive($pool.Submit('"$(preludeDefined) $global:preludeCount"', @()))) | Should Be "定義済み 1"
        } finally {
            $pool.Close()
        }
    }
}

Describe "BackgroundQueue" -Tag Unit {
    BeforeEach {
        $script:queue = [BackgroundQueue]::new(2, 'function queueHelper($x) { "[$x]" }', $null)
        $script:done = New-Object System.Collections.Generic.List[string]
    }
    AfterEach {
        $script:queue.Close()
    }

    function waitQueue {
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        while ($script:queue.Poll() -gt 0 -and $watch.Elapsed.TotalSeconds -lt 30) {
            Start-Sleep -Milliseconds 20
        }
    }

    It "終わった仕事の出力を onDone に渡す（prelude の関数を使える）" {
        $script:queue.Post('param ($a) queueHelper $a', @("x"), { param ($output, $errorText) $script:done.Add("$($output[0])|$errorText") })
        waitQueue
        $script:done -join "," | Should Be "[x]|"
    }

    It "仕事の失敗は errorText で渡す" {
        $script:queue.Post('throw "読めません"', @(), { param ($output, $errorText) $script:done.Add("$errorText") })
        $script:queue.Post('Write-Error "警告付き"; "続き"', @(), { param ($output, $errorText) $script:done.Add("$($output[0])|$errorText") })
        waitQueue
        ($script:done | Sort-Object) -join "," | Should Be "続き|警告付き,読めません"
    }

    It "終わっていない仕事の数を返す" {
        $script:queue.Post('Start-Sleep -Milliseconds 500', @(), $null)
        $script:queue.Poll() | Should Be 1
        waitQueue
        $script:queue.Poll() | Should Be 0
    }

    It "閉じると終わっていない仕事を止める。Close は何度呼んでもよい" {
        $script:queue.Post('Start-Sleep -Seconds 30', @(), { param ($output, $errorText) $script:done.Add("呼ばれた") })
        $script:queue.Close()
        $script:queue.Close()
        $script:queue.Poll() | Should Be 0
        $script:done.Count | Should Be 0
    }
}