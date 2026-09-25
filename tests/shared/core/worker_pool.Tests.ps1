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
