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
