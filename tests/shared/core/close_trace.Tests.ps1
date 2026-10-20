# 閉じる処理の記録（shared\core\close_trace.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "writeCloseTrace" -Tag Io {
    It "記録先が空のときは、何も書かない（例外も出さない）" -TestCases @(
        @{ Path = "" }
        @{ Path = $null }
    ) {
        { writeCloseTrace -Path $Path -Step "段" } | Should -Not -Throw
    }

    It "記録先があるときは、段の名前を 1 行ずつ足して書く" {
        $file = Join-Path $TestDrive "trace\close_trace.txt"
        [void][System.IO.Directory]::CreateDirectory((Split-Path $file))

        writeCloseTrace -Path $file -Step "Closing"
        writeCloseTrace -Path $file -Step "InvokeShutdown"

        $lines = @([System.IO.File]::ReadAllLines($file))
        $lines.Count | Should -Be 2
        $lines[0] | Should -Match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} Closing$'
        $lines[1] | Should -Match ' InvokeShutdown$'
    }

    It "書けない記録先（フォルダが無い）でも、例外を出さない" {
        $file = Join-Path $TestDrive "no_such_folder\close_trace.txt"

        { writeCloseTrace -Path $file -Step "段" } | Should -Not -Throw
        Test-Path -LiteralPath $file | Should -Be $false
    }
}
