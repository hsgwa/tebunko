# 画面のスモークテスト（骨組み。ランナーで画面を開けるかの確認用）
BeforeAll {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $script:tool = Join-Path $TestDrive "tool"
    Copy-Item -LiteralPath (Resolve-Path "$PSScriptRoot\..\..\scripts").Path -Destination "$script:tool\scripts" -Recurse
}

Describe "画面のスモークテスト" -Tag Gui {
    It "起動して本体のウィンドウが見つかり、閉じると終了する" {
        $gui = "$script:tool\scripts\tebunko\gui.ps1"
        $p = Start-Process powershell.exe -ArgumentList '-NoProfile', '-STA', '-ExecutionPolicy', 'RemoteSigned', '-File', "`"$gui`"" -PassThru
        $null = $p.Handle
        try {
            $found = $null
            $sw = [Diagnostics.Stopwatch]::StartNew()
            while ($sw.Elapsed.TotalSeconds -lt 90 -and !$found -and !$p.HasExited) {
                $cond = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $p.Id)
                foreach ($w in [Windows.Automation.AutomationElement]::RootElement.FindAll("Children", $cond)) {
                    $tabs = $w.FindFirst("Descendants", (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "Tabs")))
                    if ($tabs) { $found = $w }
                }
                Start-Sleep -Milliseconds 100
            }
            Write-Host ("窓が見つかるまで {0} 秒 / 終了済み {1}" -f $sw.Elapsed.TotalSeconds, $p.HasExited)
            $found | Should -Not -BeNullOrEmpty
            $found.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close()
            $p.WaitForExit(30000) | Should -BeTrue
            $p.ExitCode | Should -Be 0
        } finally {
            if (!$p.HasExited) { Stop-Process -Id $p.Id -Force }
        }
    }
}
