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

Describe "getFolderSnapshot と compareFolderSnapshot" -Tag Io {
    It "変わらなければ違いは空、増減・変更は一覧に出る" {
        $dir = Join-Path $TestDrive "snap"
        New-Item -ItemType Directory -Path $dir | Out-Null
        [System.IO.File]::WriteAllText("$dir\a.txt", "a")
        [System.IO.File]::WriteAllText("$dir\b.txt", "b")
        $before = getFolderSnapshot $dir
        @(compareFolderSnapshot $before (getFolderSnapshot $dir) "x").Count | Should -Be 0

        [System.IO.File]::WriteAllText("$dir\a.txt", "aaa")
        Remove-Item "$dir\b.txt"
        [System.IO.File]::WriteAllText("$dir\c.txt", "c")
        $diffs = @(compareFolderSnapshot $before (getFolderSnapshot $dir) "x")
        $diffs.Count | Should -Be 3
        ($diffs -join "`n") | Should -BeLike "*変わった: \a.txt*"
        ($diffs -join "`n") | Should -BeLike "*消えた: \b.txt*"
        ($diffs -join "`n") | Should -BeLike "*増えた: \c.txt*"
    }

    It "無いフォルダは「無い」と控え、後から作られると違いに出る" {
        $dir = Join-Path $TestDrive "nothing"
        $before = getFolderSnapshot $dir
        $before[""] | Should -Be "(無い)"
        New-Item -ItemType Directory -Path $dir | Out-Null
        @(compareFolderSnapshot $before (getFolderSnapshot $dir) "x").Count | Should -BeGreaterThan 0
    }

    It "中身は読まず、名前・大きさ・更新時刻だけを控える（値に本文を含まない）" {
        $dir = Join-Path $TestDrive "secret"
        New-Item -ItemType Directory -Path $dir | Out-Null
        [System.IO.File]::WriteAllText("$dir\a.txt", "SECRET_BODY")
        ((getFolderSnapshot $dir).Values -join " ") | Should -Not -Match "SECRET_BODY"
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

    It "ワークスペースの指定が使い捨てのフォルダの外なら、確かめを流さずに終了コード 1" {
        $real = Join-Path $TestDrive "real5"
        $outside = Join-Path $TestDrive "outside"
        New-Item -ItemType Directory -Path $real, $outside | Out-Null
        $marker = Join-Path $TestDrive "ran.txt"
        $cmd = [scriptblock]::Create("param (`$ToolDir, `$Workspace) [System.IO.File]::WriteAllText('$marker', 'x')")
        $text = invokeRunIsolated @{ RealWorkspace = $real; Workspace = $outside; Command = $cmd }
        $script:exitCode | Should -Be 1
        Test-Path -LiteralPath $marker | Should -BeFalse
        $text | Should -Match "起動しません"
    }
}
