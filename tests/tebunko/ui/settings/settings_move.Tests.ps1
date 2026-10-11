# ワークスペースの移動を始める処理（tebunko\ui\settings\settings.ps1 の startWorkspaceMove・updateMainContentOperable）の、
# 移している間のナビ・中身の操作を止める動きのテスト。関数だけをファイルから取り出し、画面の部品は偽物にする。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index_view.ps1"

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\settings\settings.ps1", [ref]$tokens, [ref]$errors)
    foreach ($name in "updateMainContentOperable", "startWorkspaceMove") {
        $definition = $ast.Find({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }

    function updateIndexingButton { }
    function setStatus { param ($text) }
    function showIndexingWait { param ($kind) }
    function closeIndexingWait { }
    function getWorkspaceJobQueue { param ($folders) return $null }
    function finishWorkspaceMove { }
    function startJob {
        param ($work, $arguments, $onDone, $queue)
        $global:moveProbe.EnabledAtStart = @($ui.NavHost.IsEnabled, $ui.ContentHost.IsEnabled)
        if ($global:moveProbe.Throw) { throw "始められない" }
    }
}

Describe "ワークスペースの移動中のナビ・中身の操作" -Tag Unit {
    BeforeEach {
        $script:ui = @{ NavHost = [pscustomobject]@{ IsEnabled = $true }; ContentHost = [pscustomobject]@{ IsEnabled = $true } }
        $ui = $script:ui
        $global:moveProbe = @{ EnabledAtStart = $null; Throw = $false }
        $script:workspaceMoving = $false
    }
    AfterAll {
        Remove-Variable -Name moveProbe -Scope Global -ErrorAction SilentlyContinue
    }

    It "updateMainContentOperable は、移している間だけナビと中身を使えなくする" {
        $script:workspaceMoving = $true
        updateMainContentOperable
        $script:ui.NavHost.IsEnabled | Should -BeFalse
        $script:ui.ContentHost.IsEnabled | Should -BeFalse
        $script:workspaceMoving = $false
        updateMainContentOperable
        $script:ui.NavHost.IsEnabled | Should -BeTrue
        $script:ui.ContentHost.IsEnabled | Should -BeTrue
    }

    It "startWorkspaceMove は、裏の仕事を始める前にナビと中身を使えなくする" {
        startWorkspaceMove "move" "C:\前" "C:\後"
        @($global:moveProbe.EnabledAtStart) | Should -Be @($false, $false)
        $script:workspaceMoving | Should -BeTrue
    }

    It "startWorkspaceMove は、始められなかったら、印を下ろしてナビと中身を元に戻し、例外を返す" {
        $global:moveProbe.Throw = $true
        { startWorkspaceMove "move" "C:\前" "C:\後" } | Should -Throw
        $script:workspaceMoving | Should -BeFalse
        $script:ui.NavHost.IsEnabled | Should -BeTrue
        $script:ui.ContentHost.IsEnabled | Should -BeTrue
    }
}