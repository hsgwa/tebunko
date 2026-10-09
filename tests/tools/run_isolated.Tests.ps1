# tools\isolation\isolation_common.ps1（本物の既定のワークスペースを守る関数）と tools\run_isolated.ps1 のテスト。
# 本物の場所には触れない。本物の代わりに、$TestDrive の下の場所を渡して確かめる。
BeforeAll {
    . "$PSScriptRoot\..\..\tools\isolation\isolation_common.ps1"
    $script:runIsolated = (Resolve-Path "$PSScriptRoot\..\..\tools\run_isolated.ps1").Path
}

Describe "assertNotRealWorkspace" -Tag Unit {
    It "<Case>" -TestCases @(
        @{ Case = "別の場所は通る"; Path = "C:\temp\a"; Real = "C:\Users\test\Documents\tebunko_ws"; Throws = $false }
        @{ Case = "本物と同じなら止める"; Path = "C:\Users\test\Documents\tebunko_ws"; Real = "C:\Users\test\Documents\tebunko_ws"; Throws = $true }
        @{ Case = "大文字小文字違いも止める"; Path = "c:\users\TEST\documents\TEBUNKO_WS\"; Real = "C:\Users\test\Documents\tebunko_ws"; Throws = $true }
        @{ Case = "本物の下も止める"; Path = "C:\Users\test\Documents\tebunko_ws\sub"; Real = "C:\Users\test\Documents\tebunko_ws"; Throws = $true }
        @{ Case = "本物を含む上の場所も止める"; Path = "C:\Users\test\Documents"; Real = "C:\Users\test\Documents\tebunko_ws"; Throws = $true }
        @{ Case = "名前が似ているだけなら通る"; Path = "C:\Users\test\Documents\tebunko_ws2"; Real = "C:\Users\test\Documents\tebunko_ws"; Throws = $false }
        @{ Case = "空なら止める"; Path = ""; Real = "C:\Users\test\Documents\tebunko_ws"; Throws = $true }
    ) {
        if ($Throws) {
            { assertNotRealWorkspace $Path $Real } | Should -Throw
        } else {
            { assertNotRealWorkspace $Path $Real } | Should -Not -Throw
        }
    }
}

Describe "getIsolationSnapshot と compareIsolationSnapshot" -Tag Io {
    It "変わらなければ違いは空、増減・変更は一覧に出る" {
        $dir = Join-Path $TestDrive "snap"
        New-Item -ItemType Directory -Path $dir | Out-Null
        [System.IO.File]::WriteAllText("$dir\a.txt", "a")
        [System.IO.File]::WriteAllText("$dir\b.txt", "b")
        $before = getIsolationSnapshot $dir
        @(compareIsolationSnapshot $before (getIsolationSnapshot $dir) "x").Count | Should -Be 0

        [System.IO.File]::WriteAllText("$dir\a.txt", "aaa")
        Remove-Item "$dir\b.txt"
        [System.IO.File]::WriteAllText("$dir\c.txt", "c")
        $diffs = @(compareIsolationSnapshot $before (getIsolationSnapshot $dir) "x")
        $diffs.Count | Should -Be 3
        ($diffs -join "`n") | Should -BeLike "*変わった: \a.txt*"
        ($diffs -join "`n") | Should -BeLike "*消えた: \b.txt*"
        ($diffs -join "`n") | Should -BeLike "*増えた: \c.txt*"
    }

    It "無いフォルダは「無い」と控え、後から作られると違いに出る" {
        $dir = Join-Path $TestDrive "nothing"
        $before = getIsolationSnapshot $dir
        $before[""] | Should -Be "(無い)"
        New-Item -ItemType Directory -Path $dir | Out-Null
        @(compareIsolationSnapshot $before (getIsolationSnapshot $dir) "x").Count | Should -BeGreaterThan 0
    }

    It "中身は読まず、名前・大きさ・更新時刻だけを控える（値に本文を含まない）" {
        $dir = Join-Path $TestDrive "secret"
        New-Item -ItemType Directory -Path $dir | Out-Null
        [System.IO.File]::WriteAllText("$dir\a.txt", "SECRET_BODY")
        ((getIsolationSnapshot $dir).Values -join " ") | Should -Not -Match "SECRET_BODY"
    }
}

Describe "startWorkspaceGuard と stopWorkspaceGuard" -Tag Io {
    BeforeEach {
        $script:savedEnv = $env:TEBUNKO_DEFAULT_WORKSPACE
    }
    AfterEach {
        $env:TEBUNKO_DEFAULT_WORKSPACE = $script:savedEnv
    }

    It "使い捨ての場所を環境変数に入れ、終わると元に戻して消し、本物（の代わり）に違いが無ければ空を返す" {
        $real = Join-Path $TestDrive "real1"
        New-Item -ItemType Directory -Path $real | Out-Null
        $guard = startWorkspaceGuard $real
        $env:TEBUNKO_DEFAULT_WORKSPACE | Should -Be $guard.Isolated
        Test-Path -LiteralPath $guard.Isolated | Should -BeTrue
        $diffs = @(stopWorkspaceGuard $guard)
        $diffs.Count | Should -Be 0
        $env:TEBUNKO_DEFAULT_WORKSPACE | Should -Be $script:savedEnv
        Test-Path -LiteralPath $guard.Isolated | Should -BeFalse
    }

    It "本物（の代わり）に書かれると、違いの一覧を返す" {
        $real = Join-Path $TestDrive "real2"
        New-Item -ItemType Directory -Path $real | Out-Null
        $guard = startWorkspaceGuard $real
        [System.IO.File]::WriteAllText("$real\ingesting.txt", "x")
        $diffs = @(stopWorkspaceGuard $guard)
        $diffs.Count | Should -Be 1
        $diffs[0] | Should -BeLike "*増えた: \ingesting.txt*"
    }

    It "使い捨ての場所が本物の下になる設定のときは、環境変数を変えずに止める" {
        $env:TEBUNKO_DEFAULT_WORKSPACE = "C:\keep"
        # 本物を一時フォルダの上の場所にすると、使い捨ての場所はその下になる
        { startWorkspaceGuard ([System.IO.Path]::GetTempPath()) } | Should -Throw
        $env:TEBUNKO_DEFAULT_WORKSPACE | Should -Be "C:\keep"
    }
}

Describe "getIsolationReport" -Tag Unit {
    It "違いが無ければ終了コード 0 で、表示する行は空" {
        $r = getIsolationReport @()
        $r.ExitCode | Should -Be 0
        @($r.Lines).Count | Should -Be 0
    }

    It "違いがあれば終了コード 1 で、見出し・違い・外が書いた場合の注意を出す" {
        $r = getIsolationReport @("x 増えた: \a.txt (大きさ 1 更新 2026-01-01 00:00:00)")
        $r.ExitCode | Should -Be 1
        $r.Lines[0] | Should -BeLike "*違いがありました:"
        ($r.Lines -join "`n") | Should -BeLike "*増えた: \a.txt*"
        ($r.Lines -join "`n") | Should -BeLike "*このプロセスの外が書いた場合*"
    }

    It "空の文字列だけの一覧は違いなしと数える" {
        (getIsolationReport @("", $null)).ExitCode | Should -Be 0
    }
}

Describe "変わったものの更新時刻" -Tag Io {
    It "増えた・変わったものは、大きさと更新時刻を一覧に出す" {
        $dir = Join-Path $TestDrive "mt"
        New-Item -ItemType Directory -Path $dir | Out-Null
        [System.IO.File]::WriteAllText("$dir\a.txt", "a")
        $before = getIsolationSnapshot $dir
        [System.IO.File]::WriteAllText("$dir\a.txt", "aaa")
        [System.IO.File]::WriteAllText("$dir\n.txt", "n")
        $text = (compareIsolationSnapshot $before (getIsolationSnapshot $dir) "x") -join "`n"
        $text | Should -Match "変わった: \\a.txt .*大きさ 1 更新 \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}.* -> 大きさ 3 更新"
        $text | Should -Match "増えた: \\n.txt .*大きさ 1 更新 \d{4}-"
    }
}

Describe "newIsolationWatch・getIsolationWatchDiffs・finishWorkspaceGuard" -Tag Io {
    BeforeEach { $script:savedEnv = $env:TEBUNKO_DEFAULT_WORKSPACE }
    AfterEach { $env:TEBUNKO_DEFAULT_WORKSPACE = $script:savedEnv }

    It "見張った場所が変わると、表示の名前つきで違いを返す" {
        $a = Join-Path $TestDrive "wa"
        $b = Join-Path $TestDrive "wb"
        New-Item -ItemType Directory -Path $a, $b | Out-Null
        $watch = newIsolationWatch ([ordered]@{ "場所A" = $a; "場所B" = $b })
        @(getIsolationWatchDiffs $watch).Count | Should -Be 0
        [System.IO.File]::WriteAllText("$b\x.txt", "x")
        $diffs = @(getIsolationWatchDiffs $watch)
        $diffs.Count | Should -Be 1
        $diffs[0] | Should -BeLike "場所B 増えた: \x.txt*"
    }

    It "finishWorkspaceGuard は、違いが無ければ 0、本物（の代わり）に書かれていれば 1 を返す" -TestCases @(
        @{ Write = $false; Code = 0 }
        @{ Write = $true; Code = 1 }
    ) {
        $real = Join-Path $TestDrive "fr-$Code"
        New-Item -ItemType Directory -Path $real | Out-Null
        $guard = startWorkspaceGuard $real
        if ($Write) { [System.IO.File]::WriteAllText("$real\ingesting.txt", "x") }
        $report = finishWorkspaceGuard $guard
        $report.ExitCode | Should -Be $Code
        Test-Path -LiteralPath $guard.Isolated | Should -BeFalse
    }
}

Describe "run_isolated.ps1" -Tag Io {
    BeforeAll {
        function invokeRunIsolated([hashtable]$arguments) {
            $global:LASTEXITCODE = 0
            $out = & $script:runIsolated @arguments *>&1 | Out-String
            $script:exitCode = $LASTEXITCODE
            return $out
        }
    }

    It "使い捨てのワークスペースだけに書く確かめは、終了コード 0" {
        $real = Join-Path $TestDrive "real3"
        New-Item -ItemType Directory -Path $real | Out-Null
        $text = invokeRunIsolated @{ RealWorkspace = $real; Command = { param ($ToolDir, $Workspace) [System.IO.File]::WriteAllText("$Workspace\a.txt", "x") } }
        $script:exitCode | Should -Be 0
        $text | Should -Match "本物のフォルダに違いなし"
    }

    It "本物（の代わり）に書く確かめは、違いの一覧を出して終了コード 1" {
        $real = Join-Path $TestDrive "real4"
        New-Item -ItemType Directory -Path $real | Out-Null
        $cmd = [scriptblock]::Create("param (`$ToolDir, `$Workspace) [System.IO.File]::WriteAllText('$real\ingesting.txt', 'x')")
        $text = invokeRunIsolated @{ RealWorkspace = $real; Command = $cmd }
        $script:exitCode | Should -Be 1
        $text | Should -BeLike "*増えた: \ingesting.txt*"
    }

    It "設定でワークスペースが使い捨ての外を指すなら、確かめを流さずに終了コード 1" {
        $real = Join-Path $TestDrive "real5"
        $outside = Join-Path $TestDrive "outside"   # 架空の場所（本物のフォルダではない）
        New-Item -ItemType Directory -Path $real | Out-Null
        $marker = Join-Path $TestDrive "ran.txt"
        $cmd = [scriptblock]::Create("param (`$ToolDir, `$Workspace) [System.IO.File]::WriteAllText('$marker', 'x')")
        $text = invokeRunIsolated @{ RealWorkspace = $real; Settings = @{ workspaceFolder = $outside }; Command = $cmd }
        $script:exitCode | Should -Be 1
        Test-Path -LiteralPath $marker | Should -BeFalse
        Test-Path -LiteralPath $outside | Should -BeFalse
        $text | Should -Match "起動しません"
    }

    It "確かめが例外で終わっても、本物（の代わり）の違いを比べて出す" {
        $real = Join-Path $TestDrive "real6"
        New-Item -ItemType Directory -Path $real | Out-Null
        $cmd = [scriptblock]::Create("param (`$ToolDir, `$Workspace) [System.IO.File]::WriteAllText('$real\late.txt', 'x'); throw 'boom'")
        $out = New-Object System.Collections.Generic.List[string]
        try { & $script:runIsolated -RealWorkspace $real -Command $cmd *>&1 | ForEach-Object { $out.Add([string]$_) } } catch { $out.Add("EX: $($_.Exception.Message)") }
        ($out -join "`n") | Should -BeLike "*boom*"
        ($out -join "`n") | Should -BeLike "*増えた: \late.txt*"
    }
}

Describe "run_isolated.ps1 の既定の経路（画面を開く）" -Tag Gui {
    BeforeAll {
        function invokeRunIsolatedGui([hashtable]$arguments) {
            $global:LASTEXITCODE = 0
            $out = & $script:runIsolated @arguments -WaitSeconds 10 *>&1 | Out-String
            $script:exitCode = $LASTEXITCODE
            return $out
        }
        $script:source = Join-Path $TestDrive "src"
        New-Item -ItemType Directory -Path $script:source | Out-Null
    }

    It "-Settings で取り込み先を渡した画面を開き、設定もワークスペースも使い捨ての中、終了コード 0" {
        $real = Join-Path $TestDrive "realg1"
        New-Item -ItemType Directory -Path $real | Out-Null
        $text = invokeRunIsolatedGui @{ RealWorkspace = $real; Settings = @{ targetFolders = @(@{ name = "資料"; path = $script:source; enabled = $true }) } }
        $script:exitCode | Should -Be 0
        $text | Should -Match "起動したプロセスの PID: \d+"
        $text | Should -Match "本物のフォルダに違いなし"
        $tmp = [regex]::Escape([System.IO.Path]::GetTempPath().TrimEnd([char]92))
        $text | Should -Match "設定ファイルの場所: $tmp"
        $text | Should -Match "work の場所: $tmp"
    }

    It "-Single で単一 .ps1 版の画面を開き、終了コード 0" {
        $real = Join-Path $TestDrive "realg2"
        New-Item -ItemType Directory -Path $real | Out-Null
        $text = invokeRunIsolatedGui @{ RealWorkspace = $real; Single = $true }
        $script:exitCode | Should -Be 0
        $text | Should -Match "起動したプロセスの PID: \d+"
        $text | Should -Match "本物のフォルダに違いなし"
    }
}
