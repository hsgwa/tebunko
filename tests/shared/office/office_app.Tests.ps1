# Office アプリの起動・終了（shared\office\office_app.ps1）のテスト。
# Excel・Word・PowerPoint は使わない。New-Object -ComObject と Get-Process を Mock して、偽のアプリとプロセスを返す。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\office\office_app.ps1"

$log = New-Object System.Collections.ArrayList

# 偽のアプリ。設定されなかったプロパティは「未設定」のまま残る
function newFakeApp([int]$openDocuments = 0) {
    $app = New-Object psobject -Property @{
        Visible = "未設定"; DisplayAlerts = "未設定"; EnableEvents = "未設定"; ScreenUpdating = "未設定"
        AskToUpdateLinks = "未設定"; AutomationSecurity = "未設定"
        Documents = @{ Count = $openDocuments }; Presentations = @{ Count = $openDocuments }
    }
    $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { [void]$log.Add("Quit") }
    return $app
}

# Get-Process -Name の 1 回目（起動前）と 2 回目（起動後）に返すプロセス ID
function setProcesses([int[]]$before, [int[]]$after) {
    $processes.Calls = 0
    $processes.Before = $before
    $processes.After = $after
}

$processes = @{ Calls = 0; Before = @(); After = @() }

Describe "getApp・stopApp（偽の Office アプリ）" -Tag Unit {
    Mock Get-Process {
        $processes.Calls++
        $ids = if ($processes.Calls -eq 1) { $processes.Before } else { $processes.After }
        return @($ids | ForEach-Object { [pscustomobject]@{ Id = $_ } })
    } -ParameterFilter { $Name }

    # 終了待ち（5 秒）で終わらなかったプロセス
    Mock Get-Process {
        $process = [pscustomobject]@{ Id = $Id[0] }
        $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { [void]$log.Add("WaitForExit:$($args[0])"); return $false }
        $process | Add-Member -MemberType ScriptMethod -Name Kill -Value { [void]$log.Add("Kill:$($this.Id)") }
        return $process
    } -ParameterFilter { $Id }

    BeforeEach {
        # 前のテストで起動したことになっているアプリを片付ける
        setProcesses @() @()
        foreach ($name in @("Excel", "Word", "PowerPoint")) { stopApp $name }
        $log.Clear()
    }

    It "Excel は非表示・警告なし・イベントなし・リンクを更新しない・マクロ無効で起動する" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @(100) @(100, 200)

        $app = getApp "Excel"
        $app.Visible | Should Be $false
        $app.DisplayAlerts | Should Be $false
        $app.EnableEvents | Should Be $false
        $app.ScreenUpdating | Should Be $false
        $app.AskToUpdateLinks | Should Be $false
        $app.AutomationSecurity | Should Be 3
    }

    It "2 回目からは起動済みのアプリを返す" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @() @(200)

        [void](getApp "Excel")
        [void](getApp "Excel")
        Assert-MockCalled New-Object -Times 1 -Exactly -Scope It -ParameterFilter { $ComObject -eq "Excel.Application" }
    }

    It "PID の入れ物があれば、自分で起動したアプリの PID とプロセス名を入れ、終了したら外す" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @(100) @(100, 300)
        $script:officePidSink = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,string]'
        try {
            [void](getApp "Word")
            $script:officePidSink[300] | Should Be "WINWORD"
            stopApp "Word"
            $script:officePidSink.Count | Should Be 0
        } finally {
            $script:officePidSink = $null
        }
    }

    It "自分で起動したアプリは、制限時間を過ぎたら強制終了してよいプロセスに入れる" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @(100) @(100, 300)

        [void](getApp "Word")
        $script:watchdog.Pids | Should Be @(300)
    }

    It "PowerPoint はウィンドウを隠さず（Visible を変えない）、警告なし・マクロ無効で起動する" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setProcesses @() @(400)

        $app = getApp "PowerPoint"
        $app.Visible | Should Be "未設定"
        $app.DisplayAlerts | Should Be 1
        $app.AutomationSecurity | Should Be 3
    }

    It "自分で起動したアプリは Quit し、待ち時間（Excel は 1 秒、Word・PowerPoint は 5 秒）で終わらなければプロセスを強制終了する" {
        # Excel は抽出中に取り出した COM オブジェクトが残って Quit では終わらないため、待ち時間を短くしてある（office_app.ps1 の $appInfo）
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @(100) @(100, 200)
        [void](getApp "Excel")

        stopApp "Excel"
        $log -join "|" | Should Be "Quit|WaitForExit:1000|Kill:200"
        @($script:watchdog.Pids).Count | Should Be 0

        $log.Clear()
        $fakeWord = newFakeApp
        Mock New-Object { $fakeWord } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @(100) @(100, 300)
        [void](getApp "Word")

        stopApp "Word"
        $log -join "|" | Should Be "Quit|WaitForExit:5000|Kill:300"
    }

    It "起動中の利用者のアプリに接続した場合は、終了させない" {
        # 新しいプロセスが増えない = 既に起動していたアプリに接続した
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setProcesses @(500) @(500)
        [void](getApp "PowerPoint")
        @($script:watchdog.Pids).Count | Should Be 0

        stopApp "PowerPoint"
        @($log).Count | Should Be 0
    }

    It "インデックス作成中に利用者が同じ Word で文書を開いた場合は、終了させない" {
        $fake = newFakeApp -openDocuments 1
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @() @(600)
        [void](getApp "Word")

        stopApp "Word"
        @($log).Count | Should Be 0
    }

    It "起動の前後で同じアプリのプロセスが複数増えた（どれが自分のか分からない）ときは、強制終了の対象にしない" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @(100) @(100, 700, 800)
        [void](getApp "Excel")
        @($script:watchdog.Pids).Count | Should Be 0

        # 自分で起動したことにはなるため Quit はするが、プロセスを指定した強制終了はしない
        stopApp "Excel"
        $log -join "|" | Should Be "Quit"
    }

    It "起動していないアプリの stopApp は何もしない" {
        { stopApp "Excel" } | Should Not Throw
        @($log).Count | Should Be 0
    }
}

Describe "stopAllApps" -Tag Unit {
    It "1 つのアプリの終了に失敗しても、残りのアプリは終了させる" {
        $script:apps["Excel"] = @{ Com = $null; Pid = 0; Shared = $true }
        $script:apps["Word"] = @{ Com = $null; Pid = 0; Shared = $true }
        $script:stopped = New-Object System.Collections.ArrayList
        Mock stopApp {
            [void]$script:stopped.Add($name)
            $script:apps.Remove($name)
            if ($name -eq "Excel") { throw "終了できません" }
        }
        Mock Write-Host {}

        { stopAllApps } | Should Not Throw
        @($script:stopped | Sort-Object) -join "," | Should Be "Excel,Word"
        $script:apps.Count | Should Be 0
        Assert-MockCalled Write-Host -Times 1 -Exactly -Scope It -ParameterFilter { "$Object" -match "Excel の終了に失敗しました: 終了できません" }
    }

    It "起動しているアプリが無ければ何もしない" {
        $script:apps.Clear()
        { stopAllApps } | Should Not Throw
    }
}

Describe "releaseComObject" -Tag Unit {
    It "`$null は何もしない" {
        { releaseComObject $null } | Should Not Throw
    }

    It "COM オブジェクトを解放する（解放した後は使えない）" {
        $com = New-Object -ComObject Scripting.Dictionary
        releaseComObject $com
        { $com.Add("a", 1) } | Should Throw
    }
}

Describe "startWatchdog / stopWatchdog（1 ファイルの制限時間の監視）" -Tag Unit {
    AfterEach {
        stopWatchdog
        $script:watchdog.Stop = $false
        $script:watchdog.TimedOut = $false
        $script:watchdog.Deadline = [datetime]::MaxValue
        $script:watchdog.Pids = @()
    }

    It "制限時間を過ぎたら TimedOut を立て、制限時刻を戻す。Office 以外のプロセスは終了させない" {
        # Office アプリではない自分自身のプロセス（ID が再利用された場合に当たる）
        $script:watchdog.Pids = @($PID)
        startWatchdog
        $script:watchdog.Deadline = [datetime]::Now.AddSeconds(-1)
        $deadline = [datetime]::Now.AddSeconds(10)
        while (-not $script:watchdog.TimedOut -and [datetime]::Now -lt $deadline) {
            Start-Sleep -Milliseconds 100
        }
        $script:watchdog.TimedOut | Should Be $true
        $script:watchdog.Deadline | Should Be ([datetime]::MaxValue)
        (Get-Process -Id $PID).HasExited | Should Be $false
    }

    It "制限時間内なら TimedOut を立てない。止めると監視のスレッドを片付ける" {
        startWatchdog
        $script:watchdog.Deadline = [datetime]::Now.AddMinutes(10)
        Start-Sleep -Milliseconds 700
        $script:watchdog.TimedOut | Should Be $false
        stopWatchdog
        $script:watchdogThread | Should Be $null
    }

    It "監視を始めていなければ、止めても何もしない" {
        { stopWatchdog } | Should Not Throw
    }
}

Describe "getAppName" -Tag Unit {
    It "拡張子から抽出に使うアプリを決める（大文字でも同じ）" {
        getAppName "a.xlsx" | Should Be "Excel"
        getAppName "a.XLSB" | Should Be "Excel"
        getAppName "a.xls" | Should Be "Excel"
        getAppName "a.docm" | Should Be "Word"
        getAppName "a.DOC" | Should Be "Word"
        getAppName "a.pptx" | Should Be "PowerPoint"
        getAppName "a.pptm" | Should Be "PowerPoint"
        getAppName "a.txt" | Should Be $null
    }
}

Describe "lockOfficeProcess（待たずに試す）" -Tag Unit {
    It "ほかのスレッドが鍵を持っていれば、待たずに `$null を返す。放されたら取れる" {
        $holder = holdOfficeLock "PowerPoint"
        try {
            lockOfficeProcess "PowerPoint" 0 | Should BeNullOrEmpty
        } finally {
            releaseOfficeLock $holder
        }
        $lock = lockOfficeProcess "PowerPoint" 0
        $lock | Should Not BeNullOrEmpty
        # 同じスレッドは続けて取れる（取った数だけ放す）
        $again = lockOfficeProcess "PowerPoint" 0
        $again | Should Not BeNullOrEmpty
        unlockOfficeProcess $again
        unlockOfficeProcess $lock
    }
}

Describe "testComDisconnected" -Tag Unit {
    It "相手の Office のプロセスが終わったときの例外だけを「つながっていない」とする" {
        foreach ($code in "80010108", "800706BA", "800706BE", "800401FD") {
            testComDisconnected (New-Object System.Runtime.InteropServices.COMException("切れた", [Convert]::ToInt32($code, 16))) | Should Be $true
        }
        testComDisconnected (New-Object System.Runtime.InteropServices.COMException("開けない", [Convert]::ToInt32("800A03EC", 16))) | Should Be $false
        testComDisconnected (New-Object System.InvalidOperationException("ほか")) | Should Be $false
    }
}

Describe "PowerPoint の共有（偽の PowerPoint）" -Tag Unit {
    Mock Get-Process {
        $processes.Calls++
        $ids = if ($processes.Calls -eq 1) { $processes.Before } else { $processes.After }
        return @($ids | ForEach-Object { [pscustomobject]@{ Id = $_ } })
    } -ParameterFilter { $Name }
    # 動いているプロセス（共有の PowerPoint がまだ動いているかの確認・終了の待ち）
    Mock Get-Process {
        $process = [pscustomobject]@{ Id = $Id[0]; ProcessName = "POWERPNT" }
        $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { [void]$log.Add("WaitForExit"); return $true }
        $process | Add-Member -MemberType ScriptMethod -Name Kill -Value { [void]$log.Add("Kill") }
        return $process
    } -ParameterFilter { $Id }

    BeforeEach {
        $script:powerPointShare = newPowerPointShare
        $script:powerPointHeld = $false
        setProcesses @() @()
        foreach ($name in @("Excel", "Word", "PowerPoint")) { $script:apps.Remove($name) }
        $log.Clear()
    }
    AfterEach {
        $script:powerPointShare = $null
        $script:powerPointHeld = $false
        foreach ($name in @("Excel", "Word", "PowerPoint")) { $script:apps.Remove($name) }
    }

    It "起動した PowerPoint を共有の状態に記録し、ほかのスレッドは同じ PowerPoint につなぐ（自分のものにしない）" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setProcesses @() @(400)
        [void](getApp "PowerPoint")
        $script:powerPointShare.Pid | Should Be 400
        $script:powerPointShare.Owned | Should Be $true

        # ほかのスレッドの代わりに、つながりを放してからつなぎ直す（新しいプロセスは起動しない）
        stopApp "PowerPoint"
        setProcesses @(400) @(400)
        [void](getApp "PowerPoint")
        $script:apps["PowerPoint"].Shared | Should Be $true
        $script:powerPointShare.Pid | Should Be 400
        $script:powerPointShare.Owned | Should Be $true
    }

    It "共有しているときの stopApp は、PowerPoint を終了せずにつながりを放すだけ" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setProcesses @() @(400)
        [void](getApp "PowerPoint")
        stopApp "PowerPoint"
        $log -contains "Quit" | Should Be $false
        $script:apps.ContainsKey("PowerPoint") | Should Be $false
    }

    It "共有の PowerPoint は、このスレッドが使っている間だけ監視の対象に入れる" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setProcesses @() @(400)
        [void](getApp "PowerPoint")
        @($script:watchdog.Pids) -contains 400 | Should Be $false
        $script:powerPointHeld = $true
        updateWatchedPids
        @($script:watchdog.Pids) -contains 400 | Should Be $true
    }

    It "利用者の PowerPoint につないだときは、自分のものにせず監視の対象にも入れない" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setProcesses @(500) @(500)
        [void](getApp "PowerPoint")
        $script:powerPointShare.Owned | Should Be $false
        $script:powerPointHeld = $true
        updateWatchedPids
        @($script:watchdog.Pids).Count | Should Be 0
    }

    It "最後の取り込みのスレッドが終わるときだけ、インデックス作成が起動した PowerPoint を終了する" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        $share = $script:powerPointShare
        $share.Pid = 400
        $share.Owned = $true
        enterPowerPointShare $share
        enterPowerPointShare $share
        exitPowerPointShare $share
        $log -contains "Quit" | Should Be $false
        exitPowerPointShare $share
        $log -contains "Quit" | Should Be $true
        $share.Pid | Should Be 0
        $share.Owned | Should Be $false
    }

    It "利用者の PowerPoint（自分のものでない）は、最後でも終了しない" {
        Mock New-Object { newFakeApp } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        $share = $script:powerPointShare
        $share.Pid = 500
        $share.Owned = $false
        enterPowerPointShare $share
        exitPowerPointShare $share
        Assert-MockCalled New-Object -Times 0 -Exactly -Scope It -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
    }

    It "利用者がファイルを開いている PowerPoint は、最後でも終了しない" {
        $fake = newFakeApp 1
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        $share = $script:powerPointShare
        $share.Pid = 400
        $share.Owned = $true
        closeSharedPowerPoint $share
        $log -contains "Quit" | Should Be $false
        $log -contains "Kill" | Should Be $false
    }
}
