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
        # Excel は変換中に取り出した COM オブジェクトが残って Quit では終わらないため、待ち時間を短くしてある（office_app.ps1 の $appInfo）
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

    It "変換中に利用者が同じ Word で文書を開いた場合は、終了させない" {
        $fake = newFakeApp -openDocuments 1
        Mock New-Object { $fake } -ParameterFilter { $ComObject -eq "Word.Application" }
        setProcesses @() @(600)
        [void](getApp "Word")

        stopApp "Word"
        @($log).Count | Should Be 0
    }

    It "起動していないアプリの stopApp は何もしない" {
        { stopApp "Excel" } | Should Not Throw
        @($log).Count | Should Be 0
    }
}

Describe "getAppName" -Tag Unit {
    It "拡張子から変換に使うアプリを決める（大文字でも同じ）" {
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
