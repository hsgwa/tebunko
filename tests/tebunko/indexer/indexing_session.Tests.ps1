# 画面のインデックス作成 1 回分（tebunko\indexer\indexing_session.ps1 の IndexingSession）のテスト。
# indexer.ps1 の代わりに、受け渡しの口（newIndexerChannel）だけを使う偽のスクリプトを動かす
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function newFakeIndexer([string]$name, [string]$body) {
        $path = Join-Path $TestDrive "$name.ps1"
        [System.IO.File]::WriteAllText($path, "param (`$Channel)`r`n$body", $utf8Bom)
        return $path
    }
}

Describe "IndexingSession" -Tag Io {
    It "インデクサを別のスレッド（MTA・優先度を下げる）で動かし、終わったら終了コードを返す" {
        $fake = newFakeIndexer "done" @'
$Channel.Seen = @{ Thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId; Priority = [string][System.Threading.Thread]::CurrentThread.Priority; Apartment = [string][System.Threading.Thread]::CurrentThread.GetApartmentState() }
$Channel.ExitCode = 0
'@
        $session = newIndexingSession $fake (newIndexerChannel)
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
        $fake = newFakeIndexer "wait" @'
while (!$Channel.Answered.WaitOne(20)) { }
$Channel.ExitCode = if ($Channel.Stop) { 2 } else { 0 }
'@
        $session = newIndexingSession $fake (newIndexerChannel)
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
        $fake = newFakeIndexer "throw" 'throw "読み込めませんでした"'
        $session = newIndexingSession $fake (newIndexerChannel)
        try {
            $session.Wait(30000) | Should -Be $true
            $session.GetExitCode() | Should -Be 1
            $session.GetError() | Should -Match "読み込めませんでした"
        } finally {
            $session.Close()
        }
    }

    It "インデクサが入れたエラーの内容を返す" {
        $fake = newFakeIndexer "error" '$Channel.Error = "クロール対象フォルダがありません。"; $Channel.ExitCode = 1'
        $session = newIndexingSession $fake (newIndexerChannel)
        try {
            [void]$session.Wait(30000)
            $session.GetError() | Should -Be "クロール対象フォルダがありません。"
        } finally {
            $session.Close()
        }
    }

    It "Close は、動いていれば中止を求めて終わりを待ち、片づける。何度呼んでもよい" {
        $fake = newFakeIndexer "close" @'
$Channel.Started = $true
while (!$Channel.Stop) { Start-Sleep -Milliseconds 20 }
$Channel.ExitCode = 2
'@
        $session = newIndexingSession $fake (newIndexerChannel)
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
        $fake = newFakeIndexer "office" '$Channel.ExitCode = 0'
        $session = newIndexingSession $fake (newIndexerChannel)
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

    It "インデクサが止まらずにエラーだけを書いて終わったら、その内容を理由として返す" {
        $fake = newFakeIndexer "writeerror" 'Write-Error "読み込めないファイルがありました"'
        $session = newIndexingSession $fake (newIndexerChannel)
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
        $fake = newFakeIndexer "hang" 'Start-Sleep -Seconds 60'
        $session = newIndexingSession $fake (newIndexerChannel)
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $session.Close()
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 30
        $session.IsRunning() | Should -Be $false
    }
}

Describe "インデクサの司令のスクリプト（indexingSessionScript）" -Tag Io {
    It "スレッドの優先度を下げて、indexer.ps1 に受け渡しの口を渡して動かす" {
        $fake = newFakeIndexer "direct" '$Channel.Seen = [string][System.Threading.Thread]::CurrentThread.Priority; $Channel.ExitCode = 0'
        $channel = newIndexerChannel
        $priority = [System.Threading.Thread]::CurrentThread.Priority
        try {
            & ${indexingSessionScript} $fake $channel
        } finally {
            [System.Threading.Thread]::CurrentThread.Priority = $priority
        }
        $channel.Seen | Should -Be "BelowNormal"
        $channel.ExitCode | Should -Be 0
    }
}
