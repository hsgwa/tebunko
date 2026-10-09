# 画面のスモークテストの道具（gui_helpers.ps1）のうち、画面を起動せずに確かめられる部分。
# 失敗の材料を集める関数（イベントの当て方・Summary の 1 行・終了の記録・起動のコマンド）が、材料を取りこぼさないことを止める。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"

    function newFakeEvent {
        param ([string]$Provider, [datetime]$Time, [string]$Message, [int]$Id = 1000)
        return [pscustomobject]@{ ProviderName = $Provider; TimeCreated = $Time; Message = $Message; Id = $Id }
    }
}

Describe "matchGuiEventProcessId" -Tag Unit {
    It "<Name>" -TestCases @(
        @{ Name = "10 進数の PID に当たる"; Message = "Faulting process id: 4242"; ProcessId = 4242; Expected = $true }
        @{ Name = "16 進数（小文字）に当たる"; Message = "Faulting process id: 0x1092"; ProcessId = 4242; Expected = $true }
        @{ Name = "16 進数（大文字）に当たる"; Message = "Faulting process id: 0X1A2B"; ProcessId = 6699; Expected = $true }
        @{ Name = "16 進数の頭の 0 があっても当たる"; Message = "process id: 0x00001092"; ProcessId = 4242; Expected = $true }
        @{ Name = "別の数字の一部（PID 123 に対する 1234）には当たらない"; Message = "id: 1234"; ProcessId = 123; Expected = $false }
        @{ Name = "別の数字の一部（前に数字が続く 51234）には当たらない"; Message = "id: 51234"; ProcessId = 1234; Expected = $false }
        @{ Name = "16 進数の一部（0x1A2B3 に対する 6699）には当たらない"; Message = "id: 0x1A2B3"; ProcessId = 6699; Expected = $false }
        @{ Name = "本文が空なら当たらない"; Message = ""; ProcessId = 4242; Expected = $false }
    ) {
        (matchGuiEventProcessId $Message $ProcessId) | Should -Be $Expected
    }
}

Describe "selectGuiCrashEvents" -Tag Unit {
    BeforeAll {
        $script:since = [datetime]"2026-10-10 10:00:00"
    }

    It "<Name>" -TestCases @(
        @{ Name = "Application Error で、PID が 16 進数で入っているものを拾う"; Provider = "Application Error"; Offset = 1; Message = "pid 0x1092"; Expected = 1 }
        @{ Name = ".NET Runtime を拾う"; Provider = ".NET Runtime"; Offset = 1; Message = "pid 4242"; Expected = 1 }
        @{ Name = "Windows Error Reporting を拾う"; Provider = "Windows Error Reporting"; Offset = 1; Message = "pid 4242"; Expected = 1 }
        @{ Name = "Application Hang を拾う"; Provider = "Application Hang"; Offset = 1; Message = "pid 4242"; Expected = 1 }
        @{ Name = "種類の外（MsiInstaller）は拾わない"; Provider = "MsiInstaller"; Offset = 1; Message = "pid 4242"; Expected = 0 }
        @{ Name = "起動の前の記録は拾わない"; Provider = "Application Error"; Offset = -1; Message = "pid 4242"; Expected = 0 }
        @{ Name = "別の PID の記録は拾わない"; Provider = "Application Error"; Offset = 1; Message = "pid 1234"; Expected = 0 }
    ) {
        $fake = newFakeEvent $Provider $script:since.AddMinutes($Offset) $Message
        @(selectGuiCrashEvents @($fake) 4242 $script:since).Count | Should -Be $Expected
    }

    It "イベントが 0 件（null）でも例外にならない" {
        @(selectGuiCrashEvents $null 4242 $script:since).Count | Should -Be 0
    }
}

BeforeDiscovery {
    # -TestCases は Discovery のときに評価されるので、表の中で使う本文はここで用意する（BeforeAll より前）
    $script:runtimeBody = "Application: powershell.exe`nFramework Version: v4.0.30319`nException Info: System.InvalidOperationException"
}

Describe "selectGuiCrashEventsWithoutPid・getGuiCrashInfo" -Tag Unit {
    BeforeAll {
        $script:since = [datetime]"2026-10-10 10:00:00"
        $script:runtimeBody = "Application: powershell.exe`nFramework Version: v4.0.30319`nException Info: System.InvalidOperationException"
    }

    It "<Name>" -TestCases @(
        @{ Name = "PID を書かない .NET Runtime は、PID を問わない拾い方で拾う"; Provider = ".NET Runtime"; Offset = 1; Message = $script:runtimeBody; Expected = 1 }
        @{ Name = "PID が入っているものは、こちらには入れない（PID で当たる側に出る）"; Provider = "Application Error"; Offset = 1; Message = "pid 0x1092"; Expected = 0 }
        @{ Name = "種類の外は拾わない"; Provider = "MsiInstaller"; Offset = 1; Message = "x"; Expected = 0 }
        @{ Name = "起動の前の記録は拾わない"; Provider = ".NET Runtime"; Offset = -1; Message = "x"; Expected = 0 }
    ) {
        $Message | Should -Not -BeNullOrEmpty   # 表の本文が Discovery で空にならないこと
        $fake = newFakeEvent $Provider $script:since.AddMinutes($Offset) $Message 1026
        @(selectGuiCrashEventsWithoutPid @($fake) 4242 $script:since).Count | Should -Be $Expected
    }

    It "GitHub Actions のランナーでは、PID を書かない記録の本文まで書く（PID で当たったものとは分ける）" {
        $savedCi = $env:GITHUB_ACTIONS
        try {
            $env:GITHUB_ACTIONS = "true"
            Mock Get-WinEvent { @(
                (newFakeEvent ".NET Runtime" ([datetime]::Now.AddMinutes(1)) $script:runtimeBody 1026),
                (newFakeEvent "Application Error" ([datetime]::Now.AddMinutes(1)) "pid 0x1092" 1000)
            ) }
            $text = (getGuiCrashInfo 4242 ([datetime]::Now.AddMinutes(-5))) -join "`n"
        } finally {
            $env:GITHUB_ACTIONS = $savedCi
        }
        $text | Should -Match "PID で当たった記録: 1 件"
        $text | Should -Match "PID を問わず時間で拾った記録.*: 1 件"
        $text | Should -Match "InvalidOperationException"
    }

    It "手元の実行では、PID を書かない記録は件数だけで、本文を残さない" {
        $savedCi = $env:GITHUB_ACTIONS
        try {
            $env:GITHUB_ACTIONS = $null
            Mock Get-WinEvent { @((newFakeEvent ".NET Runtime" ([datetime]::Now.AddMinutes(1)) $script:runtimeBody 1026)) }
            $text = (getGuiCrashInfo 4242 ([datetime]::Now.AddMinutes(-5))) -join "`n"
        } finally {
            $env:GITHUB_ACTIONS = $savedCi
        }
        $text | Should -Match "PID で当たった記録: 該当なし"
        $text | Should -Match "PID を問わず時間で拾った記録: 1 件"
        $text | Should -Not -Match "InvalidOperationException"
    }

    It "「該当が無い」の例外（NoMatchingEventsFound）は 0 件として、該当なしと書く" {
        Mock Get-WinEvent { Write-Error "No events were found" -ErrorId NoMatchingEventsFound -Category ObjectNotFound }
        $text = (getGuiCrashInfo 4242 ([datetime]::Now.AddMinutes(-5))) -join "`n"
        $text | Should -Match "PID で当たった記録: 該当なし"
        $text | Should -Not -Match "読めなかった"
    }

    It "読めなかったとき（ほかの例外）は、該当なしと書かずに、読めなかったと書く" {
        Mock Get-WinEvent { throw "アクセスが拒否された" }
        $text = (getGuiCrashInfo 4242 ([datetime]::Now.AddMinutes(-5))) -join "`n"
        $text | Should -Match "イベントログを読めなかった"
        $text | Should -Match "アクセスが拒否された"
        $text | Should -Not -Match "該当なし"
    }

    It "どちらも無ければ、探した範囲つきで該当なしと書く" {
        Mock Get-WinEvent { @() }
        $text = (getGuiCrashInfo 4242 ([datetime]::Now.AddMinutes(-5))) -join "`n"
        $text | Should -Match "PID で当たった記録: 該当なし（探した範囲"
        $text | Should -Match "PID を問わず時間で拾った記録: 該当なし（探した範囲"
    }
}

Describe "formatGuiSummaryRow・addGuiSummaryRow" -Tag Unit {
    It "場面・手順・文言を表の 1 行にする（| はエスケープし、改行は空白にする）" {
        formatGuiSummaryRow "S3" "閉じる" "終了コード (5)`r`n続き | あり" | Should -Be '| S3 | 閉じる | 終了コード (5) 続き \| あり |'
    }

    It "GITHUB_STEP_SUMMARY が無ければ何もしない" {
        $saved = $env:GITHUB_STEP_SUMMARY
        try {
            $env:GITHUB_STEP_SUMMARY = $null
            { addGuiSummaryRow "S1" "閉じる" "x" } | Should -Not -Throw
        } finally {
            $env:GITHUB_STEP_SUMMARY = $saved
        }
    }

    It "GITHUB_STEP_SUMMARY があれば、見出しの行を 1 回だけ書いて、失敗のたびに 1 行足す" {
        $saved = $env:GITHUB_STEP_SUMMARY
        $file = Join-Path $TestDrive "summary.md"
        try {
            $env:GITHUB_STEP_SUMMARY = $file
            addGuiSummaryRow "S1" "閉じる" "画面の終了コードが 0 ではない（5）"
            addGuiSummaryRow "S2b" "閉じる" "画面の終了コードが 0 ではない（5）"
        } finally {
            $env:GITHUB_STEP_SUMMARY = $saved
        }
        $lines = @(Get-Content -LiteralPath $file -Encoding UTF8)
        @($lines | Where-Object { $_ -like "| 場面 | 手順 |*" }).Count | Should -Be 1
        $lines | Should -Contain "| S1 | 閉じる | 画面の終了コードが 0 ではない（5） |"
        $lines | Should -Contain "| S2b | 閉じる | 画面の終了コードが 0 ではない（5） |"
    }
}

Describe "getGuiExitRecord" -Tag Unit {
    It "終わった様子・呼び出し元に戻った印・閉じる順番の記録を 1 つにまとめる" {
        $dir = Join-Path $TestDrive "tool"
        $work = Join-Path $dir "work"
        [void][IO.Directory]::CreateDirectory($work)
        [IO.File]::WriteAllText("$dir\gui_returned_4242.txt", "returned=2026-10-10T10:00:00 ok=True LASTEXITCODE= Error=")
        [IO.File]::WriteAllText("$work\close_trace.txt", "09:00:00.100`tPID 777`tClosing に入った`tスレッド 30`r`n10:00:00.100`tPID 4242`tClosing に入った`tスレッド 30`r`n10:00:00.200`tPID 4242`tShowDialog から戻った`tスレッド 28`r`n")
        $exited = Get-Date
        $S = @{
            Scene = "S1"; Step = "閉じる"; Tool = @{ Dir = $dir; Work = $work }
            Process = [pscustomobject]@{ Id = 4242; ExitCode = 5; ExitTime = $exited }
            ClosingAt = $exited.AddSeconds(-1.5); ClosingState = @("スレッドの数: 31", "子のプロセス: なし"); Children = @()
        }

        $text = (getGuiExitRecord $S) -join "`n"

        $text | Should -Match "PID: 4242　終了コード: 5"
        $text | Should -Match "閉じる操作から終わるまで: 1\.5 秒"
        $text | Should -Match "閉じる直前の スレッドの数: 31"
        $text | Should -Match "returned=2026-10-10T10:00:00 ok=True"
        $text | Should -Match "ShowDialog から戻った"
        $text | Should -Not -Match "PID 777"
    }

    It "記録のファイルが無くても、無いと書いて最後まで作る" {
        $dir = Join-Path $TestDrive "tool2"
        [void][IO.Directory]::CreateDirectory("$dir\work")
        $S = @{ Scene = "S1"; Step = "閉じる"; Tool = @{ Dir = $dir; Work = "$dir\work" }; Process = [pscustomobject]@{ Id = 1; ExitCode = 5; ExitTime = Get-Date } }

        $text = (getGuiExitRecord $S) -join "`n"

        $text | Should -Match "gui_returned_1\.txt（呼び出し元に戻った印.*\n無い"
        $text | Should -Match "close_trace\.txt（閉じる順番の記録.*\n無い"
        $text | Should -Match "控えていない"
    }

    It "ワークスペースを切り替える場面のように、作業フォルダの外にある閉じる順番の記録も、場所を付けて写す" {
        # 実際の配置: 切り替え先（別のワークスペース）は、ツールのフォルダ（tool）の隣にある
        $root = Join-Path $TestDrive "tool3root"
        $dir = Join-Path $root "tool"
        [void][IO.Directory]::CreateDirectory("$dir\work")
        [void][IO.Directory]::CreateDirectory("$root\別のワークスペース")
        [IO.File]::WriteAllText("$root\別のワークスペース\close_trace.txt", "10:00:00.100`tPID 1`t別のワークスペースの節目`tスレッド 30`r`n10:00:00.100`tPID 2`t前の起動の節目`tスレッド 30`r`n")
        [IO.File]::WriteAllText("$dir\setting.config", ("{""workspaceFolder"": " + (ConvertTo-Json "$root\別のワークスペース") + "}"))
        $S = @{ Scene = "S2c"; Step = "閉じる"; Tool = @{ Dir = $dir; Work = "$dir\work"; Config = "$dir\setting.config"; DefaultWorkspace = "$root\default_workspace" }; Process = [pscustomobject]@{ Id = 1; ExitCode = 5; ExitTime = Get-Date } }

        $text = (getGuiExitRecord $S) -join "`n"

        $text | Should -Match "場所: 別のワークスペース"
        $text | Should -Match "別のワークスペースの節目"
        $text | Should -Not -Match "前の起動の節目"
    }
}

Describe "getGuiCloseTraceFiles" -Tag Unit {
    # 探す先の決め。どの先にも close_trace.txt を置き、どれが選ばれるかを場所（フォルダ名）で見る。
    # 実機の既定のワークスペースは、パスの文字列を渡すだけ（作らない・読まない・書かない）。assertNotRealWorkspace を例外にして、外れることだけを見る
    It "<Name>" -TestCases @(
        @{ Name = "設定の場所が作業フォルダと同じなら、1 回だけ"; Folder = "{WORK}"; Expected = @("work", "default_workspace") }
        @{ Name = "workspaceFolder が空なら、既定のワークスペースを探す"; Folder = ""; Expected = @("work", "default_workspace") }
        @{ Name = "相対パスは、ツールのフォルダ基準"; Folder = "..\ws"; Expected = @("ws", "work", "default_workspace") }
        @{ Name = "実機の既定のワークスペースを指す設定は外して、残りを続ける"; Folder = "Z:\実機の既定\tebunko_ws"; Expected = @("work", "default_workspace") }
    ) {
        $root = Join-Path $TestDrive ("gc" + [guid]::NewGuid().ToString("N"))
        $tool = "$root\tool"
        foreach ($d in "$tool\work", "$root\default_workspace", "$root\ws") {
            [void][IO.Directory]::CreateDirectory($d)
            [IO.File]::WriteAllText("$d\close_trace.txt", "10:00:00.100`tPID 1`tx`tスレッド 1`r`n")
        }
        $folderValue = $Folder.Replace("{WORK}", "$tool\work")
        [IO.File]::WriteAllText("$tool\setting.config", (ConvertTo-Json @{ workspaceFolder = $folderValue }))
        Mock assertNotRealWorkspace { if ($Path -like "Z:\実機の既定*") { throw "実機の既定のワークスペース" } }
        $S = @{ Tool = @{ Dir = $tool; Work = "$tool\work"; Config = "$tool\setting.config"; DefaultWorkspace = "$root\default_workspace" } }

        $places = @(getGuiCloseTraceFiles $S | ForEach-Object { $_.Place })

        ($places -join ",") | Should -Be ($Expected -join ",")
    }
}

Describe "getGuiProcessCommand" -Tag Io {
    BeforeAll {
        function newStubTool {
            param ([string]$Name, [string]$Body)
            $dir = Join-Path $TestDrive $Name
            [void][IO.Directory]::CreateDirectory($dir)
            $gui = Join-Path $dir "it's gui.ps1"   # 単一引用符を含む名前でも起動できる
            [IO.File]::WriteAllText($gui, $Body, (New-Object System.Text.UTF8Encoding($true)))
            return @{ Dir = $dir; Gui = $gui }
        }
        function runCommand {
            param ($Tool)
            $p = Start-Process powershell.exe -ArgumentList @("-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-Command", (getGuiProcessCommand $Tool)) -PassThru -WindowStyle Hidden
            $null = $p.Handle
            try {
                if (-not $p.WaitForExit(60000)) { throw "起動した powershell.exe が 60 秒で終わらなかった" }
                return $p.ExitCode
            } finally {
                # 終わらなかったときは、起動した PID だけを止める
                if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
            }
        }
    }

    It "起動したスクリプトが正常に戻ると、終了コードは 0 で、戻った印を書く。閉じる順番の記録の印は、起動したプロセスの中にだけ立つ" {
        $tool = newStubTool "ok" '[IO.File]::WriteAllText("$PSScriptRoot\trace_env.txt", "$env:TEBUNKO_CLOSE_TRACE")'

        (runCommand $tool) | Should -Be 0

        (Get-Content -LiteralPath (Get-ChildItem "$($tool.Dir)\gui_returned_*.txt").FullName -Raw) | Should -Match "returned=.* ok=True"
        (Get-Content -LiteralPath "$($tool.Dir)\trace_env.txt" -Raw) | Should -Be "1"
        $env:TEBUNKO_CLOSE_TRACE | Should -BeNullOrEmpty
    }

    It "起動したスクリプトが失敗して戻ると、終了コードは 1 で、失敗した印（ok=False と LASTEXITCODE）を書く" {
        $tool = newStubTool "ng" 'exit 3'

        (runCommand $tool) | Should -Be 1

        $returned = Get-Content -LiteralPath (Get-ChildItem "$($tool.Dir)\gui_returned_*.txt").FullName -Raw
        $returned | Should -Match "ok=False"
        $returned | Should -Match "LASTEXITCODE=3"
    }

    It "起動したスクリプトが戻らずに例外で止まると、終了コードは 1 で、戻った印は残らない" {
        $tool = newStubTool "throw" 'throw "わざと止める"'

        (runCommand $tool) | Should -Be 1

        @(Get-ChildItem "$($tool.Dir)\gui_returned_*.txt" -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
}
