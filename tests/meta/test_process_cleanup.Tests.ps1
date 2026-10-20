# テストが自分で起こしたプロセスの後始末の決まり（docs/design/testing/run.md「テストの書き方」）を確かめる。
#   ・tests/ の中で、プロセスを止める・待つ書き方（Stop-Process・Wait-Process・.Kill()・名前で止める形）が
#     tests/helpers/process.ps1 の外に無いこと（自分で終わるプロセスとの競合でテストが揺れるのを防ぐ）
#   ・tests/helpers/process.ps1 の stopStartedProcess・waitProcessIdExit が、握る例外を限り、止まらなければ投げること
BeforeAll {
    $here = (Resolve-Path "$PSScriptRoot\..").Path
    . "$here\helpers\process.ps1"

    # 置き換えない後始末。File は tests からの相対パス、Pattern は呼び出しの全文に当てる正規表現
    $script:cleanupExceptions = @(
        # 子まで木ごと止めるため（Process の Kill では子が残る）。終了コードを見ないので、競合しても落ちない
        @{ File = "tools\run_isolated.Tests.ps1"; Pattern = '^& taskkill\.exe /PID \$p\.Id /T /F( \| Out-Null)?$' }
    )

    function getProcessCleanupViolations {
        # コードを AST で読み、後始末の直書きを "種類: 呼び出しの全文" の一覧で返す。コードを文字列で渡せるので、落ちる形を自己テストで確かめられる
        param ([string]$code, [string]$file = "(見本)")

        $ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$null, [ref]$null)
        $nodes = $ast.FindAll({
            param ($node)
            ($node -is [System.Management.Automation.Language.CommandAst]) -or
            ($node -is [System.Management.Automation.Language.InvokeMemberExpressionAst])
        }, $true)
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($node in $nodes) {
            $text = $node.Extent.Text
            if ($node -is [System.Management.Automation.Language.InvokeMemberExpressionAst]) {
                if ($node.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $node.Member.Value -eq "Kill") {
                    $found.Add("Kill: $text")
                }
                continue
            }
            # Mock・Should の引数に現れる名前（Mock Stop-Process {}・Should -Invoke Stop-Process）は、呼び出しでなく CommandAst の名前にならない
            $name = $node.GetCommandName()
            if (-not $name) { continue }
            switch -Regex ($name) {
                '^(Stop-Process|Wait-Process|spps|kill)$' { $found.Add("$($Matches[1]): $text") }
                '^taskkill(\.exe)?$' {
                    if ($text -match '(?i)\s/IM\b') {
                        $found.Add("taskkill /IM（名前で止める）: $text")
                    } else {
                        $allowed = $false
                        foreach ($exception in $script:cleanupExceptions) {
                            if ($file -eq $exception.File -and $text -match $exception.Pattern) { $allowed = $true }
                        }
                        if (-not $allowed) { $found.Add("taskkill: $text") }
                    }
                }
            }
        }
        return @($found.ToArray())
    }

    function newFakeProcess {
        # Kill・WaitForExit・HasExited・Id を持つ偽の Process。KillBehavior は Kill の動き（None・Invalid・Win32）。HasExited は Kill を呼んだ後の値
        param ([string]$KillBehavior, [bool]$ExitedAfterKill, [bool]$WaitResult, [bool]$ExitedAtStart = $false)
        $fake = [pscustomobject]@{ Id = 4242; HasExited = $ExitedAtStart; Killed = 0; Waited = 0 }
        $fake | Add-Member -MemberType ScriptMethod -Name Kill -Value {
            $this.Killed++
            $this.HasExited = $ExitedAfterKill
            if ($KillBehavior -eq "Invalid") { throw (New-Object System.InvalidOperationException "終わったプロセス") }
            if ($KillBehavior -eq "Win32") { throw (New-Object System.ComponentModel.Win32Exception 5) }
        }.GetNewClosure()
        $fake | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value {
            param ($ms)
            $this.Waited++
            return $WaitResult
        }.GetNewClosure()
        return $fake
    }
}

Describe "テストの後始末の書き方" -Tag Meta {
    It "tests/ の中に、tests/helpers/process.ps1 の外でプロセスを止める・待つ書き方が無い" {
        $helper = "$here\helpers\process.ps1"
        $violations = New-Object System.Collections.Generic.List[string]
        foreach ($file in Get-ChildItem -LiteralPath $here -Recurse -Filter *.ps1 -File) {
            if ($file.FullName -eq $helper) { continue }
            $relative = $file.FullName.Substring($here.Length + 1)
            $code = [System.IO.File]::ReadAllText($file.FullName)
            foreach ($violation in (getProcessCleanupViolations $code $relative)) {
                $violations.Add("${relative}: $violation")
            }
        }
        $violations | Should -BeNullOrEmpty -Because "後始末は tests/helpers/process.ps1 の stopStartedProcess・waitProcessIdExit を使う"
    }

    It "見本（落ちる形）: <name>" -TestCases @(
        @{ name = "HasExited の後の Stop-Process -Id"; code = 'if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }' }
        @{ name = "`$p.Kill()"; code = 'if (-not $p.HasExited) { $p.Kill() }' }
        @{ name = "Wait-Process -Id"; code = 'Wait-Process -Id 1 -Timeout 5' }
        @{ name = "spps"; code = 'spps -Id 1' }
        @{ name = "kill"; code = 'kill 1' }
        @{ name = "taskkill /IM x.exe"; code = 'taskkill /IM x.exe /F' }
        @{ name = "taskkill.exe /IM x.exe"; code = '& taskkill.exe /IM x.exe /F' }
        @{ name = "taskkill /PID（例外の表に無い）"; code = 'taskkill /PID 1 /F' }
        @{ name = "Get-Process -Name x | Stop-Process"; code = 'Get-Process -Name x | Stop-Process -Force' }
        @{ name = "Mock の本体の中の Stop-Process"; code = 'Mock Get-Process { Stop-Process -Id 1 }' }
    ) {
        @(getProcessCleanupViolations $code).Count | Should -BeGreaterThan 0
    }

    It "見本（通る形）: <name>" -TestCases @(
        @{ name = "Mock Stop-Process"; code = 'Mock Stop-Process {}'; file = "(見本)" }
        @{ name = "Should -Invoke Stop-Process"; code = 'Should -Invoke Stop-Process -Times 1'; file = "(見本)" }
        @{ name = "Add-Member -Name Kill"; code = 'Add-Member -MemberType ScriptMethod -Name Kill -Value {}'; file = "(見本)" }
        @{ name = "文字列の Stop-Process"; code = '$x = ''Stop-Process''; "Wait-Process $x $p.Kill()"'; file = "(見本)" }
        @{ name = "stopStartedProcess"; code = 'stopStartedProcess $p'; file = "(見本)" }
        @{ name = "例外の表の taskkill.exe /PID（呼び出しの名前に .exe）"; code = 'if (!$p.HasExited) { & taskkill.exe /PID $p.Id /T /F | Out-Null }'; file = "tools\run_isolated.Tests.ps1" }
    ) {
        @(getProcessCleanupViolations $code $file).Count | Should -Be 0
    }

    It "例外の表の taskkill は、別のファイルでは落ちる" {
        @(getProcessCleanupViolations '& taskkill.exe /PID $p.Id /T /F' "tools\other.Tests.ps1").Count | Should -Be 1
    }
}

Describe "stopStartedProcess（偽の Process）" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "Kill の間に終わった（InvalidOperationException・HasExited が真）→ 投げない"; behavior = "Invalid"; exited = $true; wait = $true; throws = $false }
        @{ name = "Kill が InvalidOperationException で、まだ終わっていない → 投げる"; behavior = "Invalid"; exited = $false; wait = $true; throws = $true }
        @{ name = "Kill が Win32Exception（アクセス拒否など）→ 投げる"; behavior = "Win32"; exited = $false; wait = $true; throws = $true }
        @{ name = "Kill が通っても WaitForExit が偽（止められずに残った）→ 投げる"; behavior = "None"; exited = $false; wait = $false; throws = $true }
        @{ name = "普通に止まる → 投げない"; behavior = "None"; exited = $true; wait = $true; throws = $false }
    ) {
        $fake = newFakeProcess -KillBehavior $behavior -ExitedAfterKill $exited -WaitResult $wait
        if ($throws) {
            { stopStartedProcess $fake 100 } | Should -Throw
        } else {
            { stopStartedProcess $fake 100 } | Should -Not -Throw
        }
        $fake.Killed | Should -Be 1
    }

    It "もう終わっているプロセスには Kill を呼ばない" {
        $fake = newFakeProcess -KillBehavior None -ExitedAfterKill $true -WaitResult $true -ExitedAtStart $true
        { stopStartedProcess $fake 100 } | Should -Not -Throw
        $fake.Killed | Should -Be 0
    }

    It "止められずに残ったときの例外に、PID と待った時間が入る" {
        $fake = newFakeProcess -KillBehavior None -ExitedAfterKill $false -WaitResult $false
        { stopStartedProcess $fake 123 } | Should -Throw "*4242*123*"
    }
}

Describe "stopStartedProcess・waitProcessIdExit（本物のプロセス）" -Tag Io {
    BeforeAll {
        function startChild {
            param ([string]$FileName, [string]$Arguments)
            $psi = New-Object System.Diagnostics.ProcessStartInfo($FileName, $Arguments)
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            return [System.Diagnostics.Process]::Start($psi)
        }
        $script:cmdPath = "$env:windir\System32\cmd.exe"
        $script:powershellPath = "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe"
    }

    It "すぐ終わるプロセスを、終わった後・終わる最中に止めようとしても、例外を出さない（20 回）" {
        $failures = New-Object System.Collections.Generic.List[string]
        1..20 | ForEach-Object {
            $p = startChild $script:cmdPath "/d /c exit 0"
            try {
                if ($_ % 2 -eq 0) { [void]$p.WaitForExit(5000) }
                stopStartedProcess $p
                if (-not $p.HasExited) { $failures.Add("$_ 回目: 終わっていない") }
            } catch {
                $failures.Add("$_ 回目: $($_.Exception.Message)")
            } finally {
                $p.Dispose()
            }
        }
        $failures | Should -BeNullOrEmpty
    }

    It "動き続けるプロセスは止まる" {
        $p = startChild $script:powershellPath "-NoProfile -Command Start-Sleep 30"
        try {
            stopStartedProcess $p
            $p.HasExited | Should -BeTrue
        } finally {
            $p.Dispose()
        }
    }

    It "居ない PID を待っても例外を出さない" {
        $p = startChild $script:cmdPath "/d /c exit 0"
        [void]$p.WaitForExit(5000)
        $deadId = $p.Id
        $p.Dispose()
        { waitProcessIdExit $deadId 500 } | Should -Not -Throw
    }

    It "自分で終わるプロセスの PID を待てる" {
        $p = startChild $script:powershellPath "-NoProfile -Command Start-Sleep 1"
        try {
            { waitProcessIdExit $p.Id 20000 } | Should -Not -Throw
            $p.HasExited | Should -BeTrue
        } finally {
            $p.Dispose()
        }
    }

    It "動き続けるプロセスの PID を短い待ち時間で待つと、投げる（その後で止める）" {
        $p = startChild $script:powershellPath "-NoProfile -Command Start-Sleep 30"
        try {
            { waitProcessIdExit $p.Id 300 } | Should -Throw "*終わらなかった*"
        } finally {
            stopStartedProcess $p
            $p.Dispose()
        }
    }
}
