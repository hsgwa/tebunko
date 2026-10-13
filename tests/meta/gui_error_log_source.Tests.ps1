# 画面のエラーの記録先の登録（scripts/tebunko/gui.ps1）が、消えたり前後したりしていないことのテスト。
# 共有の app_host.ps1 は、記録先と見出しをツール側が登録した式から得る。登録が無いと、
# 画面エラー（gui_error_log.txt）が黙って書かれなくなるため、起動口の構文木で確かめる。
Describe "画面のエラーの記録先の登録（gui.ps1）" -Tag Meta {
    BeforeAll {
        . "$PSScriptRoot\..\helpers\load.ps1"
        $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\gui.ps1", [ref]$null, [ref]$null)
        $commands = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true))
        $appHostLoad = @($commands | Where-Object {
            $_.InvocationOperator -eq [System.Management.Automation.Language.TokenKind]::Dot -and $_.Extent.Text -like "*app_host.ps1*"
        })
        $registrations = @($commands | Where-Object { $_.GetCommandName() -eq "setGuiErrorLogSource" })
    }

    It "app_host.ps1 を読み込む行がある" {
        $appHostLoad.Count | Should -Be 1
    }

    It "app_host.ps1 を読み込んだあとに、記録先と見出しを渡して setGuiErrorLogSource を呼ぶ" {
        $registrations.Count | Should -Be 1
        $registration = $registrations[0]
        $registration.Extent.StartOffset | Should -BeGreaterThan $appHostLoad[0].Extent.EndOffset

        $arguments = @($registration.CommandElements | Select-Object -Skip 1)
        $arguments.Count | Should -Be 2
        $arguments[0].Extent.Text | Should -Match "getToolGuiErrorLogFile"
        $arguments[1].Extent.Text | Should -Match "getToolGuiErrorLogHeader"
    }
}
