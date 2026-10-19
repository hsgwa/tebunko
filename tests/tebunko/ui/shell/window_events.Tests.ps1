# 窓の入口（tebunko\ui\shell\window_events.ps1 の onWindowActivated・onWindowPreviewKeyDown）の、移動中の守りのテスト。
# 関数だけをファイルから取り出し、読み直しや画面の部品は呼ばれた順を残す偽物にする。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit_view.ps1"
    Add-Type -AssemblyName PresentationCore, WindowsBase

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\shell\window_events.ps1", [ref]$tokens, [ref]$errors)
    foreach ($name in "onWindowActivated", "onWindowPreviewKeyDown") {
        $definition = $ast.Find({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }

    function getShortcutAction { param ($key, $ctrl, $shift, $screen) $global:activatedLog.Add("getShortcutAction:$key"); return "action" }
    function getCurrentScreen { return "SearchTab" }
    function invokeShortcutAction { param ($action) $global:activatedLog.Add("invokeShortcutAction"); return $global:shortcutHandled }

    function safe { param ([scriptblock]$block) & $block }
    function getTargetFolders { $global:activatedLog.Add("getTargetFolders"); return @() }
    function getTargetsKey { param ($folders) return "" }
    function loadTargets { $global:activatedLog.Add("loadTargets") }
    function setStatus { param ($text) }
    function refreshFolderStatus { $global:activatedLog.Add("refreshFolderStatus") }
    function isIndexing { return $false }
    function refreshIndexingState { $global:activatedLog.Add("refreshIndexingState") }
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
        Remove-Variable -Name activatedLog, shortcutHandled -Scope Global -ErrorAction SilentlyContinue
    }

    It "<name>は、何も読み直さない" -TestCases @(
        @{ name = "ワークスペースを移している間"; moving = $true; loaded = $true }
        @{ name = "起動時の読み込みが済むまで"; moving = $false; loaded = $false }
    ) {
        $script:workspaceMoving = $moving
        $script:startupLoaded = $loaded
        onWindowActivated
        @($global:activatedLog).Count | Should -Be 0
    }

    It "移していなければ、読み直す（守りが常に効いているのではない）" {
        onWindowActivated
        @($global:activatedLog) | Should -Contain "refreshFolderStatus"
        @($global:activatedLog) | Should -Contain "updateSearchTarget"
    }
}

Describe "onWindowPreviewKeyDown" -Tag Unit {
    BeforeEach {
        $global:activatedLog = New-Object System.Collections.Generic.List[string]
        $global:shortcutHandled = $true
        $script:workspaceMoving = $false
        $script:fakeKey = [pscustomobject]@{ Key = "Tab"; Handled = $false }
    }

    It "ワークスペースを移している間は、キーを捨て（Handled）、ショートカットを呼ばない" {
        $script:workspaceMoving = $true
        onWindowPreviewKeyDown $null $script:fakeKey ([System.Windows.Input.ModifierKeys]::Control)
        $script:fakeKey.Handled | Should -BeTrue
        @($global:activatedLog).Count | Should -Be 0
    }

    It "移していなければ、ショートカットを呼び、扱えたら Handled にする" {
        onWindowPreviewKeyDown $null $script:fakeKey ([System.Windows.Input.ModifierKeys]::Control)
        $script:fakeKey.Handled | Should -BeTrue
        @($global:activatedLog) | Should -Be @("getShortcutAction:Tab", "invokeShortcutAction")
    }

    It "移していなくて、ショートカットが扱わなければ、Handled にしない" {
        $global:shortcutHandled = $false
        onWindowPreviewKeyDown $null $script:fakeKey ([System.Windows.Input.ModifierKeys]::None)
        $script:fakeKey.Handled | Should -BeFalse
        @($global:activatedLog) | Should -Contain "invokeShortcutAction"
    }

    It "Alt が混ざったキーは、移している間でも触らない（窓を閉じる Alt+F4 を止めない）" {
        $script:workspaceMoving = $true
        onWindowPreviewKeyDown $null $script:fakeKey ([System.Windows.Input.ModifierKeys]::Alt)
        $script:fakeKey.Handled | Should -BeFalse
        @($global:activatedLog).Count | Should -Be 0
    }
}