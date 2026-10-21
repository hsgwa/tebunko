# 起動の段階の記録（ui\startup_trace.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\..\scripts\tebunko\ui\startup_trace.ps1"
}

Describe "startStartupTrace・markStartup" -Tag Unit {
    BeforeEach {
        startStartupTrace ""
    }

    It "環境変数が 1 でなければ、何も控えない（<flag>）" -TestCases @(
        @{ flag = "" }
        @{ flag = $null }
        @{ flag = "0" }
        @{ flag = "true" }
    ) {
        startStartupTrace $flag
        markStartup "段階"
        $script:startupTrace | Should -BeNullOrEmpty
    }

    It "1 のときは最初の行（PowerShell の起動）を控え、段階を足すたびに、経過ミリ秒が戻らない行を足す" {
        startStartupTrace "1"
        markStartup "段階 1"
        markStartup "段階 2"

        $script:startupTrace.Count | Should -Be 3
        $times = @($script:startupTrace | ForEach-Object { [double]($_ -split "`t")[1] })
        ($script:startupTrace[0] -split "`t")[0] | Should -Be "PowerShell の起動"
        ($script:startupTrace[2] -split "`t")[0] | Should -Be "段階 2"
        $times[0] | Should -BeGreaterOrEqual 0
        $times[1] | Should -BeGreaterOrEqual $times[0]
        $times[2] | Should -BeGreaterOrEqual $times[1]
    }
}

Describe "writeStartupTrace" -Tag Io {
    BeforeEach {
        startStartupTrace ""
    }

    It "記録を始めていないときは、ファイルを作らない" {
        $file = Join-Path $TestDrive "trace0.tsv"
        writeStartupTrace $file
        Test-Path -LiteralPath $file | Should -BeFalse
    }

    It "記録を始めているときは、控えた行をタブ区切りで書き、書いたあとは控えを捨てる（2 回目は書かない）" {
        startStartupTrace "1"
        markStartup "一覧が出る"
        $file = Join-Path $TestDrive "trace1.tsv"

        writeStartupTrace $file

        $lines = @([System.IO.File]::ReadAllLines($file))
        $lines.Count | Should -Be 2
        $lines[0] | Should -Match "^PowerShell の起動`t\d+$"
        $lines[1] | Should -Match "^一覧が出る`t\d+$"
        Remove-Item -LiteralPath $file
        writeStartupTrace $file
        Test-Path -LiteralPath $file | Should -BeFalse
    }

    It "書き先が空・書けない場所でも、例外を出さない" -TestCases @(
        @{ path = "" }
        @{ path = "Z:\no_such_drive_for_tebunko\trace.tsv" }
    ) {
        startStartupTrace "1"
        { writeStartupTrace $path } | Should -Not -Throw
    }
}
