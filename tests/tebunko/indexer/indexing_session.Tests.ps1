# 画面のインデックス作成 1 回分（tebunko\indexer\indexing_session.ps1 の IndexingSession）のテスト。
# invokeIndexerMain（本物は indexer_main.ps1）の代わりに、受け渡しの口（newIndexerChannel）だけを使う偽の本体を動かす
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    # Mock invokeIndexerMain で差し替えるため、本物の関数（呼ばなければ依存の indexer_lib.ps1 は要らない）を読み込んでおく
    . "${scriptsDir}\tebunko\indexer\indexer_main.ps1"

    # 本物の invokeIndexerMain（indexer_main.ps1）の代わりに、受け渡しの口（$Channel）だけを使う偽の本体を
    # 関数として State に登録する（indexerLib の部品の読み込み方と同じ仕組み。tebunko\core\parts.ps1）。
    # newIndexingSession は getPartLoad indexerLib（本物の indexerLib）を使うため、
    # 偽のインデクサを使うテストでは IndexingSession を直接作る
    function newFakeIndexerLoad([string]$body) {
        $state = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
        $entry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry("invokeIndexerMain", "param (`$Channel)`r`n$body")
        $state.Commands.Add($entry)
        return @{ State = $state; Prelude = "" }
    }

    function newFakeIndexingSession([string]$body, [hashtable]$channel) {
        return [IndexingSession]::new(${indexingSessionScript}.ToString(), (newFakeIndexerLoad $body), $channel)
    }
}

Describe "IndexingSession" -Tag Io {
    It "インデクサを別のスレッド（MTA・優先度を下げる）で動かし、終わったら終了コードを返す" {
        $body = @'
$Channel.Seen = @{ Thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId; Priority = [string][System.Threading.Thread]::CurrentThread.Priority; Apartment = [string][System.Threading.Thread]::CurrentThread.GetApartmentState() }
$Channel.ExitCode = 0
'@
        $session = newFakeIndexingSession $body (newIndexerChannel)
        try {
            $session.Wait(30000) | Should -Be $true
            $session.IsRunning() | Should -Be $false
            $session.GetExitCode() | Should -Be 0
            $session.Channel.Seen.Thread | Should -Not -Be ([System.Threading.Thread]::CurrentThread.ManagedThreadId)
            $session.Channel.Seen.Priority | Should -Be "BelowNormal"
            $session.Channel.Seen.Apartment | Should -Be "MTA"
        } finally {
            $session.Close()
        }
    }

    It "中止を求めると、受け渡しの口の Stop を立てて返事を待つのをやめさせる" {
        $body = @'
while (!$Channel.Answered.WaitOne(20)) { }
$Channel.ExitCode = if ($Channel.Stop) { 2 } else { 0 }
'@
        $session = newFakeIndexingSession $body (newIndexerChannel)
        try {
            $session.IsRunning() | Should -Be $true
            $session.Stop()
            $session.Wait(30000) | Should -Be $true
            $session.GetExitCode() | Should -Be 2
        } finally {
            $session.Close()
        }
    }

    It "終了コードを入れずに止まったら 1 とし、止まった理由を返す" {
        $session = newFakeIndexingSession 'throw "読み込めませんでした"' (newIndexerChannel)
        try {
            $session.Wait(30000) | Should -Be $true
            $session.GetExitCode() | Should -Be 1
            $session.GetError() | Should -Match "読み込めませんでした"
        } finally {
            $session.Close()
        }
    }

    It "インデクサが入れたエラーの内容を返す" {
        $session = newFakeIndexingSession '$Channel.Error = "クロール対象フォルダがありません。"; $Channel.ExitCode = 1' (newIndexerChannel)
        try {
            [void]$session.Wait(30000)
            $session.GetError() | Should -Be "クロール対象フォルダがありません。"
        } finally {
            $session.Close()
        }
    }

    It "Close は、動いていれば中止を求めて終わりを待ち、片づける。何度呼んでもよい" {
        $body = @'
$Channel.Started = $true
while (!$Channel.Stop) { Start-Sleep -Milliseconds 20 }
$Channel.ExitCode = 2
'@
        $session = newFakeIndexingSession $body (newIndexerChannel)
        # インデクサが動き始めてから閉じる（PC が混んでいると、スレッドが動き始めるまでに時間がかかる）
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        while (!$session.Channel.Started -and $watch.Elapsed.TotalSeconds -lt 30) { Start-Sleep -Milliseconds 20 }
        $session.Close()
        $session.Close()
        $session.IsRunning() | Should -Be $false
        $session.GetExitCode() | Should -Be 2
        $session.Wait(0) | Should -Be $true
    }

    It "KillOffice は、記録した PID のうちプロセス名が同じものだけを止める" {
        $session = newFakeIndexingSession '$Channel.ExitCode = 0' (newIndexerChannel)
        # Office の代わりに、このテストが起動したプロセスを使う
        $target = Start-Process -FilePath "ping.exe" -ArgumentList "-n 30 127.0.0.1" -WindowStyle Hidden -PassThru
        $other = Start-Process -FilePath "ping.exe" -ArgumentList "-n 30 127.0.0.1" -WindowStyle Hidden -PassThru
        try {
            $session.Channel.OfficePids[$target.Id] = $target.ProcessName
            $session.Channel.OfficePids[$other.Id] = "EXCEL"   # 名前が違う（ID が別のプロセスに使われた）ものは止めない
            $session.KillOffice() | Should -Be 1
            $target.WaitForExit(10000) | Should -Be $true
            $other.HasExited | Should -Be $false
        } finally {
            foreach ($process in @($target, $other)) {
                if (!$process.HasExited) { $process.Kill() }
            }
            $session.Close()
        }
    }

    It "GetNotice・GetPostponed は受け渡しの口の値を返す。入っていなければ空・0" {
        $session = newFakeIndexingSession '$Channel.Notice = "PowerPoint が起動していたため、2 件を取り込まずに残しました。"; $Channel.Postponed = 2; $Channel.ExitCode = 0' (newIndexerChannel)
        try {
            [void]$session.Wait(30000)
            $session.GetNotice() | Should -Be "PowerPoint が起動していたため、2 件を取り込まずに残しました。"
            $session.GetPostponed() | Should -Be 2
        } finally {
            $session.Close()
        }

        $emptySession = newFakeIndexingSession '$Channel.ExitCode = 0' (newIndexerChannel)
        try {
            [void]$emptySession.Wait(30000)
            $emptySession.GetNotice() | Should -Be ""
            $emptySession.GetPostponed() | Should -Be 0
        } finally {
            $emptySession.Close()
        }
    }

    It "インデクサが止まらずにエラーだけを書いて終わったら、その内容を理由として返す" {
        $session = newFakeIndexingSession 'Write-Error "読み込めないファイルがありました"' (newIndexerChannel)
        try {
            [void]$session.Wait(30000)
            $session.GetExitCode() | Should -Be 1
            $session.GetError() | Should -Match "読み込めないファイルがありました"
        } finally {
            $session.Close()
        }
    }

    It "Close で中止を求めても終わらなければ、スレッドを止めて片づける" {
        # 中止の要求を見ないインデクサ（Office が応答しないまま、など）
        $session = newFakeIndexingSession 'Start-Sleep -Seconds 60' (newIndexerChannel)
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $session.Close()
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 30
        $session.IsRunning() | Should -Be $false
    }
}

Describe "インデクサの司令のスクリプト（indexingSessionScript）" -Tag Io {
    It "スレッドの優先度を下げて、invokeIndexerMain に受け渡しの口を渡して動かす" {
        Mock invokeIndexerMain {
            $Channel.Seen = [string][System.Threading.Thread]::CurrentThread.Priority
            $Channel.ExitCode = 0
        }
        $channel = newIndexerChannel
        $priority = [System.Threading.Thread]::CurrentThread.Priority
        try {
            & ${indexingSessionScript} $channel
        } finally {
            [System.Threading.Thread]::CurrentThread.Priority = $priority
        }
        $channel.Seen | Should -Be "BelowNormal"
        $channel.ExitCode | Should -Be 0
        Should -Invoke invokeIndexerMain -Times 1 -Exactly
    }
}
