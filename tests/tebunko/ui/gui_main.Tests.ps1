# 窓を前に出したときの読み直し（tebunko\ui\gui_main.ps1 の onWindowActivated）の、移動中の守りのテスト。
# 関数だけをファイルから取り出し、読み直しの関数は呼ばれた順を残す偽物にする。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index_view.ps1"

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\gui_main.ps1", [ref]$tokens, [ref]$errors)
    $definition = $ast.Find({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "onWindowActivated" }, $true)
    . ([scriptblock]::Create($definition.Extent.Text))

    function safe { param ([scriptblock]$block) & $block }
    function getTargetFolders { $global:activatedLog.Add("getTargetFolders"); return @() }
    function getTargetsKey { param ($folders) return "" }
    function loadTargets { $global:activatedLog.Add("loadTargets") }
    function setStatus { param ($text) }
    function refreshFolderStatus { $global:activatedLog.Add("refreshFolderStatus") }
    function isIndexing { return $false }
    function refreshIndexingState { $global:activatedLog.Add("refreshIndexingState") }
    function shouldRefreshFastSearchStatus { return $false }
    function updateSearchTarget { $global:activatedLog.Add("updateSearchTarget") }
}

Describe "onWindowActivated" -Tag Unit {
    BeforeEach {
        $global:activatedLog = New-Object System.Collections.Generic.List[string]
        $script:startupLoaded = $true
        $script:workspaceMoving = $false
        $script:savedTargets = ""
    }
    AfterAll {
        Remove-Variable -Name activatedLog -Scope Global -ErrorAction SilentlyContinue
    }

    It "ワークスペースを移している間は、何も読み直さない" {
        $script:workspaceMoving = $true
        onWindowActivated
        @($global:activatedLog).Count | Should -Be 0
    }

    It "起動時の読み込みが済むまでは、何も読み直さない" {
        $script:startupLoaded = $false
        onWindowActivated
        @($global:activatedLog).Count | Should -Be 0
    }

    It "移していなければ、読み直す（守りが常に効いているのではない）" {
        onWindowActivated
        @($global:activatedLog) | Should -Contain "refreshFolderStatus"
        @($global:activatedLog) | Should -Contain "updateSearchTarget"
    }
}