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

Describe "起動した Office の PID の記録（getApp・stopApp）" -Tag Io {
    BeforeAll {
        Mock Get-Process { return $script:processSequence.Dequeue() } -ParameterFilter { $Name }
        # PID から引くプロセス。$script:byId に入れたものだけが「いる」。WaitForExit は $script:waitResult を返す
        Mock Get-Process { $script:byId[[int]$Id[0]] } -ParameterFilter { $Id }

        function newRecordProcess([int]$id, [string]$name, $start) {
            $p = [pscustomobject]@{ Id = $id; ProcessName = $name; StartTime = $start }
            $p | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { return $script:waitResult }
            $p | Add-Member -MemberType ScriptMethod -Name Kill -Value {}
            return $p
        }
    }

    BeforeEach {
        $script:officeRecordDir = $null
        foreach ($appName in @("Excel", "Word", "PowerPoint")) { stopApp $appName }
        $log.Clear()
        $script:byId = @{}
        $script:waitResult = $true
        $script:dir = Join-Path $TestDrive "office_pids\$([guid]::NewGuid())"
        $script:officeRecordDir = $dir
    }

    AfterEach {
        $script:officeRecordDir = $null
    }

    It "新しいプロセスがちょうど 1 つ増えたら、PID・名前・起動時刻・持ち主（このプロセス）の記録を書く" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        $start = [datetime]"2030-01-01T00:00:00Z"
        $script:byId[200] = newRecordProcess 200 "EXCEL" $start
        setProcesses @(100) @(100, 200)
        [void](getApp "Excel")

        $fields = ([System.IO.File]::ReadAllText("$dir\200.txt") -split "`t")
        $fields[0] | Should -Be "EXCEL"
        [long]$fields[1] | Should -Be $start.ToUniversalTime().Ticks
        [int]$fields[2] | Should -Be $PID
    }

    It "<name>ときは、記録を書かない" -TestCases @(
        @{ name = "起動時刻が読めない"; kind = "unreadable" }
        @{ name = "利用者のアプリに接続した（新しいプロセスが増えない）"; kind = "shared" }
        @{ name = "プロセスが複数増えた（どれか分からない）"; kind = "many" }
        @{ name = "記録の場所が決まっていない"; kind = "nodir" }
    ) {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        $script:byId[200] = newRecordProcess 200 "EXCEL" ([datetime]"2030-01-01T00:00:00Z")
        if ($kind -eq "unreadable") {
            $script:byId[200] | Add-Member -MemberType ScriptProperty -Name StartTime -Value { throw "アクセスが拒否されました" } -Force
        }
        if ($kind -eq "nodir") { $script:officeRecordDir = $null }
        $after = switch ($kind) { "shared" { @(100) } "many" { @(100, 200, 300) } default { @(100, 200) } }
        setProcesses @(100) $after
        [void](getApp "Excel")

        Test-Path $dir | Should -Be $false
    }

    It "終了を確かめたら記録を消す" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        $script:byId[200] = newRecordProcess 200 "EXCEL" ([datetime]"2030-01-01T00:00:00Z")
        setProcesses @(100) @(100, 200)
        [void](getApp "Excel")
        Test-Path "$dir\200.txt" | Should -Be $true

        stopApp "Excel"
        Test-Path "$dir\200.txt" | Should -Be $false
    }

    It "強制終了しても終わらなかったら、記録を残す" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        $script:byId[200] = newRecordProcess 200 "EXCEL" ([datetime]"2030-01-01T00:00:00Z")
        setProcesses @(100) @(100, 200)
        [void](getApp "Excel")
        $script:waitResult = $false

        stopApp "Excel"
        Test-Path "$dir\200.txt" | Should -Be $true
    }

    It "もうプロセスが無ければ、記録を消す" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Excel.Application" }
        $script:byId[200] = newRecordProcess 200 "EXCEL" ([datetime]"2030-01-01T00:00:00Z")
        setProcesses @(100) @(100, 200)
        [void](getApp "Excel")
        $script:byId.Remove(200)

        stopApp "Excel"
        Test-Path "$dir\200.txt" | Should -Be $false
    }

    It "利用者が文書を開いて止めなかったときは、利用者に渡したものとして記録を消す" {
        $fake = newFakeApp
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Word.Application" }
        $script:byId[300] = newRecordProcess 300 "WINWORD" ([datetime]"2030-01-01T00:00:00Z")
        setProcesses @(100) @(100, 300)
        [void](getApp "Word")
        $fake.Documents = @{ Count = 1 }

        stopApp "Word"
        Test-Path "$dir\300.txt" | Should -Be $false
    }
}

Describe "利用者が開いたブックの見分けと、利用者への引き渡し（isUnderDir・getForeignWorkbookCount・handOverApp・stopApp）" -Tag Unit {
    BeforeAll {
        function newBook([string]$fullName) {
            $book = [pscustomobject]@{ FullName = $fullName }
            $book | Add-Member -MemberType ScriptMethod -Name Close -Value { [void]$log.Add("Close:$($this.FullName)|$($args[0])") }
            return $book
        }
        function newKeptState($com, [string]$name = "Excel") {
            return @{ Name = $name; Com = $com; BooksClosed = $false; Books = $null; Done = @{} }
        }
        function newHandOverApp([object[]]$books = @(), [int]$documents = 0) {
            $app = New-Object psobject -Property @{
                Visible = $false; UserControl = $false; DisplayAlerts = $false; EnableEvents = $false; ScreenUpdating = $false
                AskToUpdateLinks = $false; AutomationSecurity = 3
                Workbooks = $books; Documents = @{ Count = $documents }; Presentations = @{ Count = $documents }
            }
            $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { [void]$log.Add("Quit") }
            return $app
        }
        $own = "C:\work\tmp\w1"
        # アプリごとの、利用者のブック・文書があるときの扱い（true = Excel のように渡す）。アプリを足したら、ここにも足す
        $handOverExpect = [ordered]@{ Excel = $true; Word = $false; PowerPoint = $false }
    }

    BeforeEach {
        foreach ($name in @("Excel", "Word", "PowerPoint")) { $script:apps.Remove($name) }
        $log.Clear()
        $script:officeOwnDir = $own
        $script:officePidSink = $null
        $script:officeRecordDir = $null
        $script:onOfficeHandOver = $null
        $script:officeKeptApps = @()
        updateWatchedPids
        Mock Get-Process { $null }
    }

    AfterEach {
        $script:officeOwnDir = $null
        $script:officePidSink = $null
        $script:officeRecordDir = $null
        $script:onOfficeHandOver = $null
        foreach ($name in @("Excel", "Word", "PowerPoint")) { $script:apps.Remove($name) }
        updateWatchedPids
    }

    It "isUnderDir: <name>" -TestCases @(
        @{ name = "置き場の下"; path = "C:\work\tmp\w1\a.xlsx"; dir = "C:\work\tmp\w1"; expected = $true }
        @{ name = "大文字小文字だけが違う"; path = "c:\WORK\TMP\W1\A.xlsx"; dir = "C:\work\tmp\w1"; expected = $true }
        @{ name = "置き場の末尾に区切りがある"; path = "C:\work\tmp\w1\a.xlsx"; dir = "C:\work\tmp\w1\"; expected = $true }
        @{ name = "置き場の名前を前置に持つ別フォルダ"; path = "C:\work\tmp\w10\a.xlsx"; dir = "C:\work\tmp\w1"; expected = $false }
        @{ name = "置き場の外"; path = "C:\docs\x\a.xlsx"; dir = "C:\work\tmp\w1"; expected = $false }
        @{ name = "URL（正規化できない）"; path = "https://example.com/a.xlsx"; dir = "C:\work\tmp\w1"; expected = $false }
    ) {
        param ($name, $path, $dir, $expected)
        isUnderDir $path $dir | Should -Be $expected
    }

    It "getForeignWorkbookCount: <name>" -TestCases @(
        @{ name = "ブックが無ければ 0"; books = @(); ownDir = "C:\work\tmp\w1"; expected = 0 }
        @{ name = "置き場の下だけなら 0"; books = @("C:\work\tmp\w1\a.xlsx"); ownDir = "C:\work\tmp\w1"; expected = 0 }
        @{ name = "外のブックがあれば数える"; books = @("C:\work\tmp\w1\a.xlsx", "C:\docs\x\山田.xlsx"); ownDir = "C:\work\tmp\w1"; expected = 1 }
        @{ name = "置き場が決まっていなければ全部数える"; books = @("C:\work\tmp\w1\a.xlsx", "C:\docs\x\山田.xlsx"); ownDir = $null; expected = 2 }
        @{ name = "tmp と tmp2 を取り違えない"; books = @("C:\work\tmp2\a.xlsx"); ownDir = "C:\work\tmp"; expected = 1 }
    ) {
        param ($name, $books, $ownDir, $expected)
        $script:officeOwnDir = $ownDir
        $com = newHandOverApp @($books | ForEach-Object { newBook $_ })
        getForeignWorkbookCount $com | Should -Be $expected
    }

    It "getForeignWorkbookCount: Workbooks が例外なら -1（分からない）" {
        # 本物の COM は Workbooks の取り出しで例外を投げる（偽物の ScriptProperty は例外を握りつぶすため、一覧を読む処理を差し替える）
        Mock getWorkbookSplit { throw "呼び出しが拒否されました" }
        getForeignWorkbookCount (newHandOverApp) | Should -Be -1
    }

    It "getForeignWorkbookCount: 置き場の下にある 8.3 の短い名前のブックは自分のものと数える（ファイルごとに窓を出さない）" {
        Mock Get-Item { [pscustomobject]@{ FullName = "C:\work\tmp\w1\長い名前のブック.xlsx" } } -ParameterFilter { $LiteralPath -eq "C:\work\tmp\w1\長い名前~1.xlsx" }
        $com = newHandOverApp @((newBook "C:\work\tmp\w1\長い名前~1.xlsx"))
        getForeignWorkbookCount $com | Should -Be 0
    }

    It "stopApp: Workbooks が読めないとき、<name>" -TestCases @(
        @{ name = "見える窓があれば渡す（Quit も強制終了もしない）"; mainWindow = 777; handed = $true }
        @{ name = "見える窓が無ければ今までどおり Quit する"; mainWindow = 0; handed = $false }
    ) {
        param ($name, $mainWindow, $handed)
        $process = [pscustomobject]@{ ProcessName = "EXCEL"; MainWindowHandle = [IntPtr]$mainWindow }
        $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { $true }
        Mock Get-Process { $process } -ParameterFilter { $Id -eq 4242 }
        Mock getWorkbookSplit { throw "呼び出しが拒否されました" }
        $com = newHandOverApp @()
        $script:apps["Excel"] = @{ Com = $com; Pid = 4242; Shared = $false }

        stopApp "Excel"

        $script:apps.ContainsKey("Excel") | Should -Be $false
        ($com.Visible) | Should -Be $handed
        (@($log) -contains "Quit") | Should -Be (-not $handed)
    }

    It "getForeignWorkbookCount: 短い名前（~ を含む）は長い名前にしてから比べる。読めなければ利用者のものと数える" {
        Mock Get-Item { [pscustomobject]@{ FullName = "C:\work\tmp\w1\長い名前.xlsx" } } -ParameterFilter { $LiteralPath -eq "C:\work\tmp\w1\LONGNA~1.xlsx" }
        Mock Get-Item { throw "読めません" } -ParameterFilter { $LiteralPath -eq "C:\work\tmp\w1\SHORT~1.xlsx" }
        Mock Get-Item { throw "読めません" } -ParameterFilter { $LiteralPath -eq "C:\other\SHORT~1.xlsx" }
        $com = newHandOverApp @((newBook "C:\work\tmp\w1\LONGNA~1.xlsx"))
        getForeignWorkbookCount $com | Should -Be 0
        # 置き場の下にあるように見えても、読めなければ置き場の下と確かめられない
        $com = newHandOverApp @((newBook "C:\work\tmp\w1\SHORT~1.xlsx"), (newBook "C:\other\SHORT~1.xlsx"))
        getForeignWorkbookCount $com | Should -Be 2
    }

    It "getForeignWorkbookCount: 取り出した Workbooks の参照も放す" {
        Mock releaseComObject {}
        $books = @((newBook "C:\work\tmp\w1\a.xlsx"))
        $com = newHandOverApp $books
        $workbooks = $com.Workbooks
        [void](getForeignWorkbookCount $com)
        Should -Invoke releaseComObject -ParameterFilter { [object]::ReferenceEquals($object, $workbooks) } -Times 1 -Exactly -Scope It
    }

    It "testProcessHasWindow: <name>" -TestCases @(
        @{ name = "名前が合い、窓が見える"; id = 4242; processName = "EXCEL"; handle = 777; expected = $true }
        @{ name = "窓が無い"; id = 4242; processName = "EXCEL"; handle = 0; expected = $false }
        @{ name = "PID が別の名前のプロセスに再利用された"; id = 4242; processName = "notepad"; handle = 777; expected = $false }
        @{ name = "PID を控えていない（0）"; id = 0; processName = "EXCEL"; handle = 777; expected = $false }
    ) {
        param ($name, $id, $processName, $handle, $expected)
        $process = [pscustomobject]@{ ProcessName = $processName; MainWindowHandle = [IntPtr]$handle }
        Mock Get-Process { $process }

        testProcessHasWindow $id "EXCEL" | Should -Be $expected
    }

    It "resolveOwnDir: <name>" -TestCases @(
        @{ name = "長い名前にできればその名前"; dir = "C:\TEST~1\tmp"; mocked = "C:\Test Long\tmp"; expected = "C:\Test Long\tmp" }
        @{ name = "読めなければ元の名前"; dir = "C:\TEST~1\tmp"; mocked = $null; expected = "C:\TEST~1\tmp" }
        @{ name = "置き場が無ければ `$null"; dir = ""; mocked = $null; expected = $null }
    ) {
        param ($name, $dir, $mocked, $expected)
        if ($mocked) { Mock Get-Item { [pscustomobject]@{ FullName = $mocked } } } else { Mock Get-Item { throw "読めません" } }
        resolveOwnDir $dir | Should -Be $expected
    }

    It "handOverApp: 見張り・一括終了・記録から外し、設定を戻して窓を出し、Quit も強制終了もしない" {
        $com = newHandOverApp @((newBook "C:\docs\x\山田.xlsx"))
        $script:apps["Excel"] = @{ Com = $com; Pid = 4242; Shared = $false }
        updateWatchedPids
        $script:watchdog.Pids | Should -Contain 4242
        $script:officePidSink = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,string]'
        $script:officePidSink[4242] = "EXCEL"
        $script:officeRecordDir = Join-Path $TestDrive "hand_over_records"
        [System.IO.Directory]::CreateDirectory($script:officeRecordDir) | Out-Null
        [System.IO.File]::WriteAllText("$($script:officeRecordDir)\4242.txt", "EXCEL`t1`t1")
        Mock Stop-Process {}

        handOverApp "Excel"

        $script:apps.ContainsKey("Excel") | Should -Be $false
        @($script:watchdog.Pids) | Should -Not -Contain 4242
        $script:officePidSink.ContainsKey(4242) | Should -Be $false
        Test-Path "$($script:officeRecordDir)\4242.txt" | Should -Be $false
        $com.Visible | Should -Be $true
        $com.UserControl | Should -Be $true
        $com.DisplayAlerts | Should -Be $true
        $com.EnableEvents | Should -Be $true
        $com.ScreenUpdating | Should -Be $true
        $com.AskToUpdateLinks | Should -Be $true
        $com.AutomationSecurity | Should -Be 1
        @($log).Count | Should -Be 0
        Should -Invoke Get-Process -Times 0 -Exactly -Scope It
        Should -Invoke Stop-Process -Times 0 -Exactly -Scope It
    }

    It "handOverApp: 自分が開いたブック（置き場の下）は Close(`$false) し、利用者のブックは閉じない" {
        $com = newHandOverApp @((newBook "C:\work\tmp\w1\壊れた.xlsx"), (newBook "C:\docs\x\山田.xlsx"))
        $script:apps["Excel"] = @{ Com = $com; Pid = 4242; Shared = $false }

        handOverApp "Excel"

        @($log) | Should -Be @("Close:C:\work\tmp\w1\壊れた.xlsx|False")
    }

    It "handOverApp: 置き場の下と確かめられない（読めない ~ の）ブックは閉じない" {
        Mock Get-Item { throw "読めません" }
        $com = newHandOverApp @((newBook "C:\work\tmp\w1\SHORT~1.xlsx"), (newBook "C:\docs\x\予算~案.xlsx"))
        $script:apps["Excel"] = @{ Com = $com; Pid = 4242; Shared = $false }

        handOverApp "Excel"

        @($log).Count | Should -Be 0
    }

    It "handOverApp: 窓や設定を戻せたら `$true を返し、通知に窓を出せたことを伝える" {
        $script:onOfficeHandOver = { param ($name, $shown) [void]$log.Add("通知:${name}:${shown}") }
        $script:apps["Excel"] = @{ Com = (newHandOverApp); Pid = 4242; Shared = $false }

        (handOverApp "Excel") | Should -Be $true

        @($log) | Should -Be @("通知:Excel:True")
    }

    It "handOverApp: 窓を出せなかったら `$false を返し、COM を放さずに持ち続け（止めも強制終了もしない）、通知に伝える" {
        Mock releaseComObject {}
        Mock Stop-Process {}
        $script:officeKeptApps = @()
        $script:onOfficeHandOver = { param ($name, $shown) [void]$log.Add("通知:${name}:${shown}") }
        $com = newHandOverApp @((newBook "C:\docs\x\山田.xlsx"))
        $com.PSObject.Properties.Remove("Visible")
        $com | Add-Member -MemberType ScriptProperty -Name Visible -Value { $false } -SecondValue { throw "設定できません" }
        $script:apps["Excel"] = @{ Com = $com; Pid = 4242; Shared = $false }
        updateWatchedPids

        (handOverApp "Excel") | Should -Be $false

        @($log) | Should -Be @("通知:Excel:False")
        $script:apps.ContainsKey("Excel") | Should -Be $false
        @($script:watchdog.Pids) | Should -Not -Contain 4242
        @($script:officeKeptApps).Count | Should -Be 1
        [object]::ReferenceEquals($script:officeKeptApps[0].Com, $com) | Should -Be $true
        Should -Invoke releaseComObject -ParameterFilter { [object]::ReferenceEquals($object, $com) } -Times 0 -Exactly -Scope It
        Should -Invoke Stop-Process -Times 0 -Exactly -Scope It
        $script:officeKeptApps = @()
    }

    It "handOverApp: 一度失敗しても、やり直して窓を出せたら `$true" {
        $script:attempts = 0
        $com = newHandOverApp @((newBook "C:\docs\x\山田.xlsx"))
        $com.PSObject.Properties.Remove("UserControl")
        $com | Add-Member -MemberType ScriptProperty -Name UserControl -Value { $script:uc } -SecondValue {
            $script:attempts++
            if ($script:attempts -lt 2) { throw "呼び出しが拒否されました" }
            $script:uc = $args[0]
        }
        $script:apps["Excel"] = @{ Com = $com; Pid = 4242; Shared = $false }

        (handOverApp "Excel") | Should -Be $true

        $script:attempts | Should -Be 2
    }

    It "handOverApp: 設定を戻すところで例外が出ても、見張り・一括終了・記録からは外れている" {
        $com = New-Object psobject -Property @{ Workbooks = @() }
        foreach ($property in "Visible", "UserControl", "DisplayAlerts", "EnableEvents", "ScreenUpdating", "AskToUpdateLinks", "AutomationSecurity") {
            $com | Add-Member -MemberType ScriptProperty -Name $property -Value { 0 } -SecondValue { throw "設定できません" }
        }
        $script:apps["Excel"] = @{ Com = $com; Pid = 4242; Shared = $false }
        $script:officePidSink = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,string]'
        $script:officePidSink[4242] = "EXCEL"
        updateWatchedPids

        { handOverApp "Excel" } | Should -Not -Throw

        $script:apps.ContainsKey("Excel") | Should -Be $false
        @($script:watchdog.Pids) | Should -Not -Contain 4242
        $script:officePidSink.ContainsKey(4242) | Should -Be $false
    }

    It "stopAllApps: 起動し直しやスレッドの終わりの一括終了でも、渡したことを知らせる" {
        $script:onOfficeHandOver = { param ($name, $shown) [void]$log.Add("通知:${name}:${shown}") }
        $script:apps["Excel"] = @{ Com = (newHandOverApp @((newBook "C:\docs\x\山田.xlsx"))); Pid = 4242; Shared = $false }

        stopAllApps

        @($log) | Should -Be @("通知:Excel:True")
        @($log) -contains "Quit" | Should -Be $false
    }

    It "handOverApp: 利用者の Excel に接続したもの（Shared）には当てない（設定も変えない）" {
        $com = newHandOverApp @((newBook "C:\docs\x\山田.xlsx"))
        $script:apps["Excel"] = @{ Com = $com; Pid = 0; Shared = $true }

        handOverApp "Excel"

        $script:apps.ContainsKey("Excel") | Should -Be $true
        $com.Visible | Should -Be $false
        $com.DisplayAlerts | Should -Be $false
        $com.AutomationSecurity | Should -Be 3
    }

    It "stopApp の表が appInfo と一致する（アプリを足して行が欠けたら落ちる）" {
        @($handOverExpect.Keys | Sort-Object) | Should -Be @($appInfo.Keys | Sort-Object)
    }

    It "stopApp: 利用者が開いたブック・文書があるとき、<app> は<expected>" -TestCases @(
        @{ app = "Excel"; expected = "渡し（Quit も強制終了もしない）" }
        @{ app = "Word"; expected = "止めない" }
        @{ app = "PowerPoint"; expected = "止めない" }
    ) {
        param ($app, $expected)
        $com = newHandOverApp @((newBook "C:\docs\x\山田.xlsx")) 1
        $script:apps[$app] = @{ Com = $com; Pid = 4242; Shared = $false }

        stopApp $app

        $script:apps.ContainsKey($app) | Should -Be $false
        @($log) -contains "Quit" | Should -Be $false
        $com.Visible | Should -Be $handOverExpect[$app]
        Should -Invoke Get-Process -Times 0 -Exactly -Scope It
    }

    It "retryKeptApps: 全部通って初めて参照を放して「渡した」と知らせる。任意の設定も戻し、自分のブックだけ閉じ、通らなければ持ち続ける（待たない・止めない）" {
        Mock releaseComObject {}
        Mock Stop-Process {}
        $script:onOfficeHandOver = { param ($name, $shown) [void]$log.Add("通知:${name}:${shown}") }
        $good = newHandOverApp @((newBook "C:\work\tmp\w1\壊れた.xlsx"), (newBook "C:\docs\x\山田.xlsx"))
        $bad = newHandOverApp
        $bad.PSObject.Properties.Remove("Visible")
        $bad | Add-Member -MemberType ScriptProperty -Name Visible -Value { $false } -SecondValue { throw "設定できません" }
        $script:officeKeptApps = @((newKeptState $good), (newKeptState $bad))

        retryKeptApps

        $good.Visible | Should -Be $true
        $good.UserControl | Should -Be $true
        $good.DisplayAlerts | Should -Be $true
        $good.EnableEvents | Should -Be $true
        $good.ScreenUpdating | Should -Be $true
        $good.AutomationSecurity | Should -Be 1
        # 自分のブックだけ Close($false)。利用者のブックは閉じない
        @($log) | Should -Be @("Close:C:\work\tmp\w1\壊れた.xlsx|False", "通知:Excel:True")
        Should -Invoke releaseComObject -ParameterFilter { [object]::ReferenceEquals($object, $good) } -Times 1 -Exactly -Scope It
        Should -Invoke releaseComObject -ParameterFilter { [object]::ReferenceEquals($object, $bad) } -Times 0 -Exactly -Scope It
        Should -Invoke Stop-Process -Times 0 -Exactly -Scope It
        @($script:officeKeptApps).Count | Should -Be 1
        [object]::ReferenceEquals($script:officeKeptApps[0].Com, $bad) | Should -Be $true
    }

    It "retryKeptApps: 任意の設定が通らないうちは「渡した」としない（必須が通っていても持ち続ける）" {
        Mock releaseComObject {}
        $com = newHandOverApp
        $com.PSObject.Properties.Remove("ScreenUpdating")
        $com | Add-Member -MemberType ScriptProperty -Name ScreenUpdating -Value { $false } -SecondValue { throw "設定できません" }
        $script:officeKeptApps = @((newKeptState $com))

        retryKeptApps

        @($script:officeKeptApps).Count | Should -Be 1
        Should -Invoke releaseComObject -ParameterFilter { [object]::ReferenceEquals($object, $com) } -Times 0 -Exactly -Scope It
    }

    It "retryKeptApps: 最初にブックの一覧を読めなくて閉じられなかった自分のブックも、後で閉じる" {
        Mock releaseComObject {}
        $script:splitCalls = 0
        $book = newBook "C:\work\tmp\w1\壊れた.xlsx"
        $com = newHandOverApp @($book)
        $state = newKeptState $com
        Mock getWorkbookSplit { $script:splitCalls++; if ($script:splitCalls -eq 1) { throw "呼び出しが拒否されました" }; @{ Own = @($book); Foreign = @() } }

        (restoreHandedOverApp $state 1) | Should -Be $false
        @($log).Count | Should -Be 0
        $script:officeKeptApps = @($state)
        retryKeptApps

        @($log) | Should -Be @("Close:C:\work\tmp\w1\壊れた.xlsx|False")
        @($script:officeKeptApps).Count | Should -Be 0
        Should -Invoke releaseComObject -ParameterFilter { [object]::ReferenceEquals($object, $book) } -Times 1 -Exactly -Scope It
    }

    Context "waitKeptApps（取り込みのスレッドの終わりの待ち。時計は Start-Sleep の差し替えで進める）" {
        BeforeAll {
            # Visible を、指定した回数だけ拒んでから通す偽の Excel
            function newSlowApp([int]$failures) {
                $script:failsLeft = $failures
                $app = newHandOverApp
                $app.PSObject.Properties.Remove("Visible")
                $app | Add-Member -MemberType ScriptProperty -Name Visible -Value { $false } -SecondValue {
                    if ($script:failsLeft -gt 0) { $script:failsLeft--; throw "設定できません" }
                }
                return $app
            }
        }
        BeforeEach {
            Mock Start-Sleep { $script:slept += $Milliseconds }
            Mock releaseComObject {}
            Mock Stop-Process {}
            $script:slept = 0
            $script:onOfficeHandOver = { param ($name, $shown) [void]$log.Add("通知:${name}:${shown}") }
        }
        It "待つうちに通れば、そこで渡して抜ける" {
            $script:officeKeptApps = @((newKeptState (newSlowApp 2)))

            waitKeptApps

            @($script:officeKeptApps).Count | Should -Be 0
            @($log) | Should -Be @("通知:Excel:True")
            $script:slept | Should -Be (2 * $script:officeKeptWaitStep)
        }

        It "上限まで通らなければ諦める（上限の秒数だけ待ち、Quit・強制終了・参照の解放はしない）" {
            $script:officeKeptApps = @((newKeptState (newSlowApp 100000)))

            waitKeptApps

            $script:slept | Should -Be ($script:officeKeptWaitSeconds * 1000)
            @($script:officeKeptApps).Count | Should -Be 1
            @($log).Count | Should -Be 0
            Should -Invoke releaseComObject -ParameterFilter { [object]::ReferenceEquals($object, $script:officeKeptApps[0].Com) } -Times 0 -Exactly -Scope It
            Should -Invoke Stop-Process -Times 0 -Exactly -Scope It
        }

        It "持ち続けが無ければ待たない" {
            $script:officeKeptApps = @()

            waitKeptApps

            $script:slept | Should -Be 0
            Should -Invoke Start-Sleep -Times 0 -Exactly -Scope It
        }

        It "中止の合図が来ていれば、待たずに抜ける" {
            $script:officeKeptApps = @((newKeptState (newSlowApp 100000)))

            waitKeptApps { $true }

            $script:slept | Should -Be 0
            @($script:officeKeptApps).Count | Should -Be 1
        }
    }

    It "持ち続けている Excel は、<when>に設定し直される" -TestCases @(
        @{ when = "次の Excel のファイルの確かめ（handOverForeignApp）"; run = { [void](handOverForeignApp "Excel") } }
        @{ when = "一括終了（stopAllApps）"; run = { stopAllApps } }
    ) {
        param ($when, $run)
        Mock releaseComObject {}
        $kept = newHandOverApp
        $script:officeKeptApps = @((newKeptState $kept))

        & $run

        $kept.Visible | Should -Be $true
        @($script:officeKeptApps).Count | Should -Be 0
    }

    It "stopApp: Excel に利用者が開いたブックが無ければ、今までどおり Quit する" {
        $com = newHandOverApp @((newBook "C:\work\tmp\w1\a.xlsx"))
        $script:apps["Excel"] = @{ Com = $com; Pid = 0; Shared = $false }

        stopApp "Excel"

        @($log) | Should -Be @("Quit")
        $com.Visible | Should -Be $false
    }
}
