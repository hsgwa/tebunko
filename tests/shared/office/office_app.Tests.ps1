# Office アプリの起動・終了（shared\office\office_app.ps1）のテスト。
# Excel・Word・PowerPoint は使わない。New-Object -ComObject と Get-Process を Mock して、偽のアプリとプロセスを返す。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\shared\office\office_app.ps1"

    $log = New-Object System.Collections.ArrayList
    # 自分のセッションと、ほかの利用者・ほかの作業フォルダの tebunko を想定した別のセッション
    $ownSession = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    $otherSession = $ownSession + 1

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

    # Get-Process -Name が1回返す並び（自分のセッション・ほかのセッションの両方を混ぜて返せる）
    function newProcessSet([int[]]$ownIds = @(), [int[]]$otherIds = @()) {
        $own = @($ownIds | ForEach-Object { [pscustomobject]@{ Id = $_; SessionId = $ownSession } })
        $other = @($otherIds | ForEach-Object { [pscustomobject]@{ Id = $_; SessionId = $otherSession } })
        return , @($own + $other)
    }

    # Get-Process -Name の呼び出しごとに、渡した並びを順に返す
    function setProcessSequence {
        param ([object[]]$sequence)
        $script:processSequence = New-Object 'System.Collections.Generic.Queue[object]'
        foreach ($item in $sequence) { $script:processSequence.Enqueue($item) }
    }

    # Excel・Word（起動前・起動後の確認だけ）用: Get-Process -Name を2回呼ぶ
    function setProcesses {
        param ([int[]]$before, [int[]]$after, [int[]]$otherBefore = @(), [int[]]$otherAfter = @())
        setProcessSequence @((newProcessSet $before $otherBefore), (newProcessSet $after $otherAfter))
    }

    # PowerPoint（起動する前に既に自分のセッションにいないかも確かめる）用: Get-Process -Name を3回呼ぶ
    # （事前の確認・起動前の確認・起動後の確認。事前の確認と起動前の確認の間には何もしないため同じ値になる）
    function setSingleInstanceProcesses {
        param ([int[]]$before, [int[]]$after, [int[]]$otherBefore = @(), [int[]]$otherAfter = @())
        setProcessSequence @((newProcessSet $before $otherBefore), (newProcessSet $before $otherBefore), (newProcessSet $after $otherAfter))
    }

    # PowerPoint が事前の確認で例外になり、起動を試さない場合用: Get-Process -Name を1回だけ呼ぶ
    function setPrecheckProcesses {
        param ([int[]]$ownIds, [int[]]$otherIds = @())
        setProcessSequence @((newProcessSet $ownIds $otherIds))
    }
}

Describe "getApp・stopApp（偽の Office アプリ）" -Tag Unit {
    BeforeAll {
        Mock Get-Process {
            return $script:processSequence.Dequeue()
        } -ParameterFilter { $Name }

        # 終了待ち（5 秒）で終わらなかったプロセス
        Mock Get-Process {
            $process = [pscustomobject]@{ Id = $Id[0] }
            $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { [void]$log.Add("WaitForExit:$($args[0])"); return $false }
            $process | Add-Member -MemberType ScriptMethod -Name Kill -Value { [void]$log.Add("Kill:$($this.Id)") }
            return $process
        } -ParameterFilter { $Id }
    }

    BeforeEach {
        # 前のテストで起動したことになっているアプリを片付ける
        foreach ($name in @("Excel", "Word", "PowerPoint")) { stopApp $name }
        $log.Clear()
    }

    It "Excel は非表示・警告なし・イベントなし・リンクを更新しない・マクロ無効で起動する" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @(100) @(100, 200)

        $app = getApp "Excel"
        $app.Visible | Should -Be $false
        $app.DisplayAlerts | Should -Be $false
        $app.EnableEvents | Should -Be $false
        $app.ScreenUpdating | Should -Be $false
        $app.AskToUpdateLinks | Should -Be $false
        $app.AutomationSecurity | Should -Be 3
    }

    It "2 回目からは起動済みのアプリを返す" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @() @(200)

        $first = getApp "Excel"
        $second = getApp "Excel"
        [object]::ReferenceEquals($first, $second) | Should -Be $true
        Should -Invoke New-Object -Times 1 -Exactly -Scope It -ParameterFilter { $ComObject -eq "Excel.Application" }
    }

    It "PID の入れ物があれば、自分で起動したアプリの PID とプロセス名を入れ、終了したら外す" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @(100) @(100, 300)
        $script:officePidSink = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,string]'
        try {
            [void](getApp "Word")
            $script:officePidSink[300] | Should -Be "WINWORD"
            stopApp "Word"
            $script:officePidSink.Count | Should -Be 0
        } finally {
            $script:officePidSink = $null
        }
    }

    It "自分で起動したアプリは、制限時間を過ぎたら強制終了してよいプロセスに入れる" -TestCases @(
        @{ Name = "Word"; Prog = "Word.Application"; SingleInstance = $false }
        @{ Name = "PowerPoint"; Prog = "PowerPoint.Application"; SingleInstance = $true }
    ) {
        param ($Name, $Prog, $SingleInstance)
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq $Prog }
        if ($SingleInstance) {
            setSingleInstanceProcesses @() @(300)
        } else {
            setProcesses @(100) @(100, 300)
        }

        [void](getApp $Name)
        $script:watchdog.Pids | Should -Be @(300)
    }

    It "PowerPoint はウィンドウを隠さず（Visible を変えない）、警告なし・マクロ無効で起動する" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setSingleInstanceProcesses @() @(400)

        $app = getApp "PowerPoint"
        $app.Visible | Should -Be "未設定"
        $app.DisplayAlerts | Should -Be 1
        $app.AutomationSecurity | Should -Be 3
    }

    It "自分で起動したアプリは Quit し、待ち時間（Excel は 1 秒、Word・PowerPoint は 5 秒）で終わらなければプロセスを強制終了し、同じ待ち時間で終わるのを待つ" {
        # Excel は抽出中に取り出した COM オブジェクトが残って Quit では終わらないため、待ち時間を短くしてある（office_app.ps1 の $appInfo）
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @(100) @(100, 200)
        [void](getApp "Excel")

        stopApp "Excel"
        $log -join "|" | Should -Be "Quit|WaitForExit:1000|Kill:200|WaitForExit:1000"
        @($script:watchdog.Pids).Count | Should -Be 0

        $log.Clear()
        $fakeWord = newFakeApp
        Mock New-Object { $fakeWord } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @(100) @(100, 300)
        [void](getApp "Word")

        stopApp "Word"
        $log -join "|" | Should -Be "Quit|WaitForExit:5000|Kill:300|WaitForExit:5000"
    }

    It "新しいプロセスが増えず起動中の利用者のアプリに接続した場合（Excel・Word）は、終了させない" {
        # 新しいプロセスが増えない = 既に起動していたアプリに接続した（PowerPoint はこの状態にならないよう、下で別に確かめる）
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @(500) @(500)
        [void](getApp "Excel")
        @($script:watchdog.Pids).Count | Should -Be 0

        stopApp "Excel"
        @($log).Count | Should -Be 0
    }

    It "インデックス作成中に利用者が同じ Word で文書を開いた場合は、終了させない" {
        $fake = newFakeApp -openDocuments 1
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @() @(600)
        [void](getApp "Word")

        stopApp "Word"
        @($log).Count | Should -Be 0
    }

    It "起動の前後で同じアプリのプロセスが複数増えた（どれが自分のか分からない）ときは、強制終了の対象にしない" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        setProcesses @(100) @(100, 700, 800)
        [void](getApp "Excel")
        @($script:watchdog.Pids).Count | Should -Be 0

        # 自分で起動したことにはなるため Quit はするが、プロセスを指定した強制終了はしない
        stopApp "Excel"
        $log -join "|" | Should -Be "Quit"
    }

    It "起動の前後にほかのセッションのプロセスが増えても、自分のものと取り違えない（Pid・見張りの対象に入らない）" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @() @() @() @(999)
        [void](getApp "Word")

        $script:apps["Word"].Pid | Should -Be 0
        $script:apps["Word"].Shared | Should -Be $true
        @($script:watchdog.Pids) | Should -Not -Contain 999
    }

    It "PowerPoint は自分のセッションに既に起動していれば、New-Object を呼ばずに例外を投げる（利用者が開いている・強制終了で終わらずに残った場合を含む）" {
        Mock New-Object { newFakeApp } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setPrecheckProcesses @(500)

        { getApp "PowerPoint" } | Should -Throw -ExpectedMessage "*PowerPoint が起動しているため*" -ExceptionType ([System.InvalidOperationException])
        Should -Invoke New-Object -Times 0 -Exactly -Scope It -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        $script:apps.ContainsKey("PowerPoint") | Should -Be $false
    }

    It "PowerPoint はほかのセッションにだけ起動していれば、接続せず新しく起動する" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setSingleInstanceProcesses @() @(410) @(999) @(999)

        $app = getApp "PowerPoint"
        $app.DisplayAlerts | Should -Be 1
        Should -Invoke New-Object -Times 1 -Exactly -Scope It -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
    }

    It "PowerPoint は起動後に新しいプロセスが増えなければ、設定を変える前に解放して例外を投げる（`$script:apps に残さない）" {
        # 事前の確認から New-Object の間に、利用者が先にアプリを起動した場合に当たる
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        setSingleInstanceProcesses @() @()

        { getApp "PowerPoint" } | Should -Throw -ExpectedMessage "*PowerPoint が起動しているため*" -ExceptionType ([System.InvalidOperationException])
        $fake.AutomationSecurity | Should -Be "未設定"
        $fake.DisplayAlerts | Should -Be "未設定"
        @($log).Count | Should -Be 0
        $script:apps.ContainsKey("PowerPoint") | Should -Be $false
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

        { stopAllApps } | Should -Not -Throw
        @($script:stopped | Sort-Object) -join "," | Should -Be "Excel,Word"
        $script:apps.Count | Should -Be 0
        Should -Invoke Write-Host -Times 1 -Exactly -Scope It -ParameterFilter { "$Object" -match "Excel の終了に失敗しました: 終了できません" }
    }
}

Describe "releaseComObject" -Tag Unit {
    It "`$null は何もしない" {
        { releaseComObject $null } | Should -Not -Throw
    }

    It "COM オブジェクトを解放する（解放した後は使えない）" {
        $com = New-Object -ComObject Scripting.Dictionary
        releaseComObject $com
        { $com.Add("a", 1) } | Should -Throw
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
        $script:watchdog.TimedOut | Should -Be $true
        $script:watchdog.Deadline | Should -Be ([datetime]::MaxValue)
        (Get-Process -Id $PID).HasExited | Should -Be $false
    }

    It "制限時間内なら TimedOut を立てない。止めると監視のスレッドを片付ける" {
        startWatchdog
        $script:watchdog.Deadline = [datetime]::Now.AddMinutes(10)
        Start-Sleep -Milliseconds 700
        $script:watchdog.TimedOut | Should -Be $false
        stopWatchdog
        $script:watchdogThread | Should -Be $null
    }
}

Describe "getAppName" -Tag Unit {
    It "拡張子から抽出に使うアプリを決める（大文字でも同じ）" {
        getAppName "a.xlsx" | Should -Be "Excel"
        getAppName "a.XLSB" | Should -Be "Excel"
        getAppName "a.xls" | Should -Be "Excel"
        getAppName "a.docm" | Should -Be "Word"
        getAppName "a.DOC" | Should -Be "Word"
        getAppName "a.pptx" | Should -Be "PowerPoint"
        getAppName "a.pptm" | Should -Be "PowerPoint"
        getAppName "a.txt" | Should -Be $null
    }
}

Describe "getApp（起動したプロセスの優先度）" -Tag Unit {
    BeforeAll {
        Mock Get-Process {
            return $script:processSequence.Dequeue()
        } -ParameterFilter { $Name }
        Mock Get-Process {
            if (!$script:started.ContainsKey($Id[0])) {
                $script:started[$Id[0]] = [pscustomobject]@{ Id = $Id[0]; ProcessName = "X"; PriorityClass = "Normal" }
            }
            return $script:started[$Id[0]]
        } -ParameterFilter { $Id }
    }

    It "起動した Office の優先度は変えない（利用者とプロセスを共有しうるため）" {
        $script:started = @{}
        Mock New-Object { newFakeApp } -ParameterFilter { $ComObject -eq "PowerPoint.Application" -or $ComObject -eq "Excel.Application" }
        foreach ($name in @("Excel", "PowerPoint")) { $script:apps.Remove($name) }
        setSingleInstanceProcesses @() @(410)
        [void](getApp "PowerPoint")
        setProcesses @() @(420)
        [void](getApp "Excel")
        foreach ($id in 410, 420) {
            (!$script:started.ContainsKey($id) -or [string]$script:started[$id].PriorityClass -eq "Normal") | Should -Be $true
        }
        foreach ($name in @("Excel", "PowerPoint")) { $script:apps.Remove($name) }
    }
}
