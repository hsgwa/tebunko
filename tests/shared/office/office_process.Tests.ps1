# Office プロセス（shared\office\office_process.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getOfficeProcesses / stopOfficeProcesses" -Tag Io {
    It "ウィンドウを持たないプロセスをバックグラウンドとし、表示名を付ける" {
        Mock Get-Process {
            @(
                [pscustomobject]@{ Id = 1; ProcessName = "EXCEL"; MainWindowHandle = [IntPtr]::Zero; StartTime = [datetime]"2030-01-01"; WorkingSet64 = 10MB; MainWindowTitle = "" }
                [pscustomobject]@{ Id = 2; ProcessName = "WINWORD"; MainWindowHandle = [IntPtr]100; StartTime = [datetime]"2030-01-01"; WorkingSet64 = 20MB; MainWindowTitle = "文書 - Word" }
            )
        }
        $processes = @(getOfficeProcesses)
        $processes.Count | Should -Be 2
        $processes[0].AppName | Should -Be "Excel"
        $processes[0].Background | Should -Be $true
        $processes[1].AppName | Should -Be "Word"
        $processes[1].Background | Should -Be $false
        $processes[1].MemoryMB | Should -Be 20
    }

    It "終了できなかったプロセスは理由を返す" {
        Mock Stop-Process { if ($Id -eq 2) { throw "アクセスが拒否されました" } }
        $results = @(stopOfficeProcesses @(1, 2))
        $results.Count | Should -Be 2
        $results[0].Stopped | Should -Be $true
        $results[1].Stopped | Should -Be $false
        $results[1].Message | Should -Be "アクセスが拒否されました"
    }
}
