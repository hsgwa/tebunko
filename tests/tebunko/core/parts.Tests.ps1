# 部品（lib・indexerLib・indexer）の読み込み先（tebunko\core\parts.ps1）のテスト。
# 設計は docs/design/structure/single-script.md
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getPartLoad（zip 版。`${bundledScriptPath} が無い）" -Tag Unit {
    It "<name> は scripts 配下の実際のファイルを指し、引数は空" -ForEach @(
        @{ Name = "lib"; File = "tebunko\lib.ps1" }
        @{ Name = "indexerLib"; File = "tebunko\indexer\indexer_lib.ps1" }
        @{ Name = "indexer"; File = "tebunko\indexer.ps1" }
    ) {
        $load = getPartLoad $Name
        $load.Path | Should -Be (Resolve-Path "$scriptsDir\$File").Path
        Test-Path -LiteralPath $load.Path | Should -Be $true
        $load.Args.Count | Should -Be 0
    }
}

Describe "getPartLoad（単一 .ps1 版。`${bundledScriptPath} がある）" -Tag Unit {
    BeforeEach {
        $global:bundledScriptPath = "C:\fake\single.ps1"
    }
    AfterEach {
        Remove-Variable -Name bundledScriptPath -Scope Global -ErrorAction SilentlyContinue
    }

    It "<name> は自分自身のパスを指し、引数に Part を渡す" -ForEach @(
        @{ Name = "lib" }
        @{ Name = "indexerLib" }
        @{ Name = "indexer" }
    ) {
        $load = getPartLoad $Name
        $load.Path | Should -Be "C:\fake\single.ps1"
        $load.Args.Part | Should -Be $Name
    }
}

Describe "getPartInitScript" -Tag Unit {
    It "引数が無ければ、パスだけの 1 行にする" {
        getPartInitScript @{ Path = "C:\scripts\lib.ps1"; Args = @{} } | Should -Be ". 'C:\scripts\lib.ps1'"
    }

    It "引数があれば、-<キー> '<値>' を並べる" {
        getPartInitScript @{ Path = "C:\fake\single.ps1"; Args = @{ Part = "lib" } } | Should -Be ". 'C:\fake\single.ps1' -Part 'lib'"
    }

    It "パス・引数の値に含まれる ' は '' にする（文字列として壊れないように）" {
        getPartInitScript @{ Path = "C:\it's\lib.ps1"; Args = @{ Part = "it's" } } | Should -Be ". 'C:\it''s\lib.ps1' -Part 'it''s'"
    }
}
