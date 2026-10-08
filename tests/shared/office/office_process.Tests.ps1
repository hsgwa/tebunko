# Office プロセス（shared\office\office_process.ps1）のテスト
# 止めてよいのは「記録があり、記録と今のプロセスが合うもの」だけ。PID が使い回された・名前や起動時刻が違うものを止めないことを確かめる。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    # Get-Process の結果に見立てた値。start に $null を渡すと「起動時刻が読めない」ものになる
    function newFakeProcess {
        param ([int]$id, [string]$name = "EXCEL", $start = [datetime]"2030-01-01T00:00:00Z", [int]$session = 0, [bool]$window = $false)
        $p = [pscustomobject]@{
            Id = $id; ProcessName = $name; SessionId = $session
            MainWindowHandle = $(if ($window) { [IntPtr]100 } else { [IntPtr]::Zero })
            StartTime = $start
        }
        if ($null -eq $start) {
            $p | Add-Member -MemberType ScriptProperty -Name StartTime -Value { throw "アクセスが拒否されました" } -Force
        }
        return $p
    }
}

Describe "記録の書き込みと読み込み" -Tag Io {
    BeforeAll {
        $script:t0 = ([datetime]"2030-01-01T00:00:00Z").ToUniversalTime().Ticks
        $script:t1 = ([datetime]"2030-01-02T00:00:00Z").ToUniversalTime().Ticks
    }
    BeforeEach {
        $script:dir = Join-Path $TestDrive "office_pids\key1"
        $script:procs = @{}
        Mock Get-Process { if ($Id -and $script:procs.ContainsKey([int]$Id[0])) { $script:procs[[int]$Id[0]] } }
    }

    It "書いた記録を読み戻せ、書きかけの .tmp は残らない" {
        $script:procs[10] = newFakeProcess 10
        addOfficeRecord $dir 10 "EXCEL" $t0 55 $t1 | Should -Be $true
        @(Get-ChildItem $dir).Name | Should -Be @("10.txt")
        $r = @(readOfficeRecords $dir)
        $r.Count | Should -Be 1
        $r[0].Id | Should -Be 10
        $r[0].ProcessName | Should -Be "EXCEL"
        $r[0].OwnerId | Should -Be 55
        $r[0].OwnerStartTicks | Should -Be $t1
        $r[0].Owned | Should -Be $true
    }

    It "持ち主を省くと、書いたプロセス自身になる" {
        $script:procs[10] = newFakeProcess 10
        addOfficeRecord $dir 10 "EXCEL" $t0 | Should -Be $true
        (readOfficeRecords $dir)[0].OwnerId | Should -Be $PID
    }

    It "場所が空、または起動時刻が 0 以下なら書かない" {
        addOfficeRecord "" 10 "EXCEL" $t0 | Should -Be $false
        addOfficeRecord (Join-Path $TestDrive "no_write") 10 "EXCEL" 0 | Should -Be $false
        Test-Path (Join-Path $TestDrive "no_write") | Should -Be $false
    }

    It "書けない場所（フォルダの場所にファイルがある）は `$false を返し、例外にしない" {
        Set-Content (Join-Path $TestDrive "file_as_dir") "x"
        addOfficeRecord (Join-Path $TestDrive "file_as_dir\sub") 10 "EXCEL" $t0 | Should -Be $false
    }

    It ".tmp は読まず、1 時間より古いものだけ消す" {
        New-Item -ItemType Directory $dir -Force | Out-Null
        Set-Content "$dir\20.tmp" "EXCEL`t$t0`t1`t1"
        Set-Content "$dir\21.tmp" "EXCEL`t$t0`t1`t1"
        (Get-Item "$dir\21.tmp").LastWriteTimeUtc = [datetime]::UtcNow.AddHours(-2)
        @(readOfficeRecords $dir).Count | Should -Be 0
        Test-Path "$dir\20.tmp" | Should -Be $true
        Test-Path "$dir\21.tmp" | Should -Be $false
    }

    It "<name>は、記録を消し、返さない" -TestCases @(
        @{ name = "PID のプロセスが無い"; kind = "none" }
        @{ name = "PID が別の名前のプロセスに使い回された"; kind = "name" }
        @{ name = "PID が起動時刻の違うプロセスに使い回された"; kind = "start" }
    ) {
        if ($kind -eq "name") { $script:procs[10] = newFakeProcess 10 "notepad" }
        if ($kind -eq "start") { $script:procs[10] = newFakeProcess 10 "EXCEL" ([datetime]"2031-05-05T00:00:00Z") }
        addOfficeRecord $dir 10 "EXCEL" $t0 | Out-Null
        @(readOfficeRecords $dir).Count | Should -Be 0
        Test-Path "$dir\10.txt" | Should -Be $false
    }

    It "起動時刻が読めないプロセスは Owned にせず、記録は残す" {
        $script:procs[10] = newFakeProcess 10 "EXCEL" $null
        addOfficeRecord $dir 10 "EXCEL" $t0 | Out-Null
        $r = @(readOfficeRecords $dir)
        $r.Count | Should -Be 1
        $r[0].Owned | Should -Be $false
        Test-Path "$dir\10.txt" | Should -Be $true
    }

    It "形の違う記録（<name>）は使わず、消さない" -TestCases @(
        @{ name = "項目が足りない"; text = "EXCEL`t1" }
        @{ name = "Office のプロセス名でない"; text = "notepad`t100`t1`t1" }
        @{ name = "起動時刻が数でない"; text = "EXCEL`tabc`t1`t1" }
        @{ name = "起動時刻が 0"; text = "EXCEL`t0`t1`t1" }
        @{ name = "空"; text = "" }
    ) {
        $script:procs[10] = newFakeProcess 10
        New-Item -ItemType Directory $dir -Force | Out-Null
        [System.IO.File]::WriteAllText("$dir\10.txt", $text)
        @(readOfficeRecords $dir).Count | Should -Be 0
        Test-Path "$dir\10.txt" | Should -Be $true
    }

    It "ファイル名が PID でない記録は読まない" {
        New-Item -ItemType Directory $dir -Force | Out-Null
        Set-Content "$dir\abc.txt" "EXCEL`t$t0`t1`t1"
        @(readOfficeRecords $dir).Count | Should -Be 0
    }

    It "フォルダが無い・場所が空のときは空の配列" {
        @(readOfficeRecords (Join-Path $TestDrive "none")).Count | Should -Be 0
        @(readOfficeRecords "").Count | Should -Be 0
    }

    It "ほかの PC の鍵のフォルダは、読まず、消さない" {
        $other = Join-Path $TestDrive "office_pids\key2"
        addOfficeRecord $other 10 "EXCEL" $t0 | Out-Null   # PID 10 のプロセスは無いが、別の鍵なので触らない
        @(readOfficeRecords $dir).Count | Should -Be 0
        Test-Path "$other\10.txt" | Should -Be $true
    }

    It "removeOfficeRecord は、無い記録・空の場所でも失敗にしない" {
        { removeOfficeRecord $dir 99 } | Should -Not -Throw
        { removeOfficeRecord "" 99 } | Should -Not -Throw
    }
}

Describe "getOfficeRecordOwner" -Tag Io {
    BeforeAll {
        $script:t0 = ([datetime]"2030-01-01T00:00:00Z").ToUniversalTime().Ticks
        $script:t1 = ([datetime]"2030-01-02T00:00:00Z").ToUniversalTime().Ticks
    }
    BeforeEach {
        $script:procs = @{}
        Mock Get-Process { if ($Id -and $script:procs.ContainsKey([int]$Id[0])) { $script:procs[[int]$Id[0]] } }
    }

    It "<name>" -TestCases @(
        @{ name = "持ち主がこの画面のプロセス → Self"; ownerId = 100; ownerTicks = 7; alive = "none"; expected = "Self" }
        @{ name = "持ち主が生きていて起動時刻が同じ → Other"; ownerId = 200; ownerTicks = 630000000000000000; alive = "same"; expected = "Other" }
        @{ name = "持ち主の PID が無い → Gone"; ownerId = 200; ownerTicks = 630000000000000000; alive = "none"; expected = "Gone" }
        @{ name = "持ち主の PID が起動時刻の違うプロセスに使い回された → Gone"; ownerId = 200; ownerTicks = 1; alive = "same"; expected = "Gone" }
        @{ name = "持ち主の起動時刻が読めない（生きている）→ Other（安全な側）"; ownerId = 200; ownerTicks = 1; alive = "unreadable"; expected = "Other" }
        @{ name = "記録に持ち主の起動時刻が無い（0）で生きている → Other"; ownerId = 200; ownerTicks = 0; alive = "same"; expected = "Other" }
    ) {
        if ($alive -eq "same") { $script:procs[200] = newFakeProcess 200 "x" ([datetime]::new(630000000000000000, [DateTimeKind]::Utc)) }
        if ($alive -eq "unreadable") { $script:procs[200] = newFakeProcess 200 "x" $null }
        $record = [pscustomobject]@{ OwnerId = $ownerId; OwnerStartTicks = $ownerTicks }
        getOfficeRecordOwner $record 100 7 | Should -Be $expected
    }
}

Describe "getOfficeProcesses" -Tag Io {
    BeforeAll {
        $script:t0 = ([datetime]"2030-01-01T00:00:00Z").ToUniversalTime().Ticks
        $script:t1 = ([datetime]"2030-01-02T00:00:00Z").ToUniversalTime().Ticks
    }
    BeforeEach {
        $script:dir = Join-Path $TestDrive "office_pids\key1"
        $script:all = @()
        Mock Get-Process {
            if ($Id) { $script:all | Where-Object { $_.Id -eq $Id[0] } } else { $script:all | Where-Object { $Name -contains $_.ProcessName } }
        }
        $script:session = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    }

    It "窓の有無・表示名・記録の有無（Owned）・持ち主の状態を返す" {
        $script:all = @(
            (newFakeProcess 1 "EXCEL" ([datetime]"2030-01-01T00:00:00Z") $session $false)
            (newFakeProcess 2 "WINWORD" ([datetime]"2030-01-01T00:00:00Z") $session $true)
            (newFakeProcess 3 "EXCEL" ([datetime]"2030-01-01T00:00:00Z") $session $false)
        )
        addOfficeRecord $dir 1 "EXCEL" $t0 900 $t1 | Out-Null
        $p = @(getOfficeProcesses $dir 100 5)
        $p.Count | Should -Be 3
        $p[0].AppName | Should -Be "Excel"
        $p[0].Owned | Should -Be $true
        $p[0].Owner | Should -Be "Gone"
        $p[0].HasWindow | Should -Be $false
        $p[1].AppName | Should -Be "Word"
        $p[1].HasWindow | Should -Be $true
        $p[1].Owned | Should -Be $false
        $p[1].Owner | Should -Be ""
        $p[2].Owned | Should -Be $false
    }

    It "ほかのセッションのプロセスは返さない" {
        $script:all = @((newFakeProcess 1 "EXCEL" ([datetime]"2030-01-01T00:00:00Z") ($session + 1) $false))
        @(getOfficeProcesses $dir 100 5).Count | Should -Be 0
    }

    It "記録の場所を省くと、どれも Owned にならない" {
        $script:all = @((newFakeProcess 1 "EXCEL" ([datetime]"2030-01-01T00:00:00Z") $session $false))
        (@(getOfficeProcesses))[0].Owned | Should -Be $false
    }

    It "起動時刻が読めなくても返す（StartTime は `$null）" {
        $script:all = @((newFakeProcess 1 "EXCEL" $null $session $false))
        $p = @(getOfficeProcesses $dir 100 5)
        $p.Count | Should -Be 1
        $p[0].StartTime | Should -Be $null
        $p[0].Owned | Should -Be $false
    }
}

Describe "stopOfficeProcesses" -Tag Io {
    BeforeAll {
        $script:t0 = ([datetime]"2030-01-01T00:00:00Z").ToUniversalTime().Ticks
    }
    BeforeEach {
        $script:dir = Join-Path $TestDrive "office_pids\key1"
        $script:procs = @{}
        Mock Get-Process { if ($Id -and $script:procs.ContainsKey([int]$Id[0])) { $script:procs[[int]$Id[0]] } }
        Mock Stop-Process {}
    }

    It "<name>なら止めず、理由を返し、記録は消さない" -TestCases @(
        @{ name = "名前が違う（PID が別のアプリに使い回された）"; procName = "notepad"; procStart = "2030-01-01T00:00:00Z"; targetName = "EXCEL"; targetStart = "2030-01-01T00:00:00Z" }
        @{ name = "起動時刻が違う"; procName = "EXCEL"; procStart = "2031-01-01T00:00:00Z"; targetName = "EXCEL"; targetStart = "2030-01-01T00:00:00Z" }
        @{ name = "今のプロセスの起動時刻が読めない"; procName = "EXCEL"; procStart = $null; targetName = "EXCEL"; targetStart = "2030-01-01T00:00:00Z" }
        @{ name = "確認した起動時刻が無い"; procName = "EXCEL"; procStart = "2030-01-01T00:00:00Z"; targetName = "EXCEL"; targetStart = $null }
        @{ name = "対象が Office のプロセス名でない"; procName = "notepad"; procStart = "2030-01-01T00:00:00Z"; targetName = "notepad"; targetStart = "2030-01-01T00:00:00Z" }
    ) {
        $script:procs[10] = newFakeProcess 10 $procName $(if ($procStart) { [datetime]$procStart } else { $null })
        addOfficeRecord $dir 10 "EXCEL" $t0 | Out-Null
        $target = @{ Id = 10; ProcessName = $targetName; StartTime = $(if ($targetStart) { [datetime]$targetStart } else { $null }) }
        $r = @(stopOfficeProcesses @($target) $dir)
        Should -Invoke Stop-Process -Times 0 -Exactly
        $r.Count | Should -Be 1
        $r[0].Stopped | Should -Be $false
        $r[0].Status | Should -Be "Changed"
        $r[0].Reason | Should -Be "PID 10 は確認の後に別のプロセスに変わったため、終了しませんでした。"
        Test-Path "$dir\10.txt" | Should -Be $true
    }

    It "名前も起動時刻も同じなら止め、記録を消す" {
        $script:procs[10] = newFakeProcess 10
        addOfficeRecord $dir 10 "EXCEL" $t0 | Out-Null
        $r = @(stopOfficeProcesses @(@{ Id = 10; ProcessName = "EXCEL"; StartTime = [datetime]"2030-01-01T00:00:00Z" }) $dir)
        Should -Invoke Stop-Process -Times 1 -Exactly -ParameterFilter { $Id -eq 10 }
        $r[0].Stopped | Should -Be $true
        $r[0].Status | Should -Be "Stopped"
        Test-Path "$dir\10.txt" | Should -Be $false
    }

    It "もう無いプロセスは止めず、止めた扱いにも失敗にもせず、記録を消す" {
        addOfficeRecord $dir 10 "EXCEL" $t0 | Out-Null
        $r = @(stopOfficeProcesses @(@{ Id = 10; ProcessName = "EXCEL"; StartTime = [datetime]"2030-01-01T00:00:00Z" }) $dir)
        Should -Invoke Stop-Process -Times 0 -Exactly
        $r[0].Stopped | Should -Be $false
        $r[0].Status | Should -Be "Gone"
        $r[0].Reason | Should -Be ""
        Test-Path "$dir\10.txt" | Should -Be $false
    }

    It "止められなかったものは理由を返し、記録を残す。ほかの対象は続けて止める" {
        $script:procs[1] = newFakeProcess 1
        $script:procs[2] = newFakeProcess 2
        Mock Stop-Process { if ($Id -eq 1) { throw "アクセスが拒否されました" } }
        addOfficeRecord $dir 1 "EXCEL" $t0 | Out-Null
        $time = [datetime]"2030-01-01T00:00:00Z"
        $r = @(stopOfficeProcesses @(@{ Id = 1; ProcessName = "EXCEL"; StartTime = $time }, @{ Id = 2; ProcessName = "EXCEL"; StartTime = $time }) $dir)
        $r.Count | Should -Be 2
        $r[0].Status | Should -Be "Failed"
        $r[0].Reason | Should -Be "アクセスが拒否されました"
        $r[1].Status | Should -Be "Stopped"
        Test-Path "$dir\1.txt" | Should -Be $true
    }
}
