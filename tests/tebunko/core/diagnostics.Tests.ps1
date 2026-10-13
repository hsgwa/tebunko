# tebunko のログの見出し（tebunko\core\diagnostics.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getToolDiagnosticHeader" -Tag Io {
    It "ツールの名前と版の行を持つ" {
        $text = (getToolDiagnosticHeader) -join "`n"
        $text | Should -Match "==== tebunko の診断情報 ===="
        $text | Should -Match "tebunko の版: "
    }
}

Describe "writeIndexerLog" -Tag Unit {
    AfterEach {
        $script:indexerLog = $null
        $script:indexerEcho = $false
    }

    It "ログを開いていればログに書く。開いていなければ何もしない" {
        { writeIndexerLog "何もしない" } | Should -Not -Throw
        $script:indexerLog = New-Object System.IO.StringWriter
        writeIndexerLog "1 行目"
        writeIndexerLog "2 行目" "Yellow"
        $script:indexerLog.ToString() | Should -Be "1 行目`r`n2 行目`r`n"
    }

    It "コンソールに出すときは色を付ける" {
        Mock Write-Host {}
        $script:indexerEcho = $true
        writeIndexerLog "注意" "Yellow"
        writeIndexerLog "普通"
        Should -Invoke Write-Host -Times 1 -Exactly -Scope It -ParameterFilter { "$Object" -eq "注意" -and $ForegroundColor -eq "Yellow" }
        Should -Invoke Write-Host -Times 1 -Exactly -Scope It -ParameterFilter { "$Object" -eq "普通" -and !$ForegroundColor }
    }

    It "ログに書けなくても止めない" {
        $writer = New-Object System.IO.StringWriter
        $writer.Dispose()
        $script:indexerLog = $writer
        { writeIndexerLog "書けない" } | Should -Not -Throw
    }
}
