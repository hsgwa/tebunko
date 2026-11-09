# 窓の入口（tebunko\ui\shell\window_events.ps1 の onWindowActivated・onWindowPreviewKeyDown）の、移動中の守りのテスト。
# 関数だけをファイルから取り出し、読み直しや画面の部品は呼ばれた順を残す偽物にする。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit_view.ps1"
    Add-Type -AssemblyName PresentationCore, WindowsBase

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\shell\window_events.ps1", [ref]$tokens, [ref]$errors)
    foreach ($name in "reloadTargetsIfChanged", "checkSettingsStamp", "onWindowActivated", "onWindowPreviewKeyDown") {
        $definition = $ast.Find({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }

    function getShortcutAction { param ($key, $ctrl, $shift, $screen) $global:activatedLog.Add("getShortcutAction:$key"); return "action" }
    function getCurrentScreen { return "SearchTab" }
    function invokeShortcutAction { param ($action) $global:activatedLog.Add("invokeShortcutAction"); return $global:shortcutHandled }

    function safe { param ([scriptblock]$block) & $block }
    function getTargetFolders { $global:activatedLog.Add("getTargetFolders"); return @() }
    function getTargetsKey { param ($folders) return $global:targetsKey }
    function testFileWatchSignaled { param ($Watch) if ($null -eq $Watch) { return $false }; return [bool]$global:signals[$Watch] }
    function getFileStamp { param ($Path) return $global:currentStamp }
    function startJob { param ($scriptBlock, $arguments, $onDone, $queue) $global:activatedLog.Add("startJob:$queue") }
    function getWorkspaceJobQueue { param ($paths) return "network" }
    function loadTargets { $global:activatedLog.Add("loadTargets") }
    function setStatus { param ($text) }
    function refreshFolderStatus { $global:activatedLog.Add("refreshFolderStatus") }
    function isIndexing { return [bool]$global:indexingNow }
    function setIndexingNow { param ($value) $global:indexingNow = $value }
    function refreshIndexingState { param ([switch]$SkipIfUnchanged) $global:activatedLog.Add("refreshIndexingState:$SkipIfUnchanged") }
    function updateSearchTarget { $global:activatedLog.Add("updateSearchTarget") }
}

Describe "onWindowActivated" -Tag Unit {
    BeforeEach {
        $global:activatedLog = New-Object System.Collections.Generic.List[string]
        $script:startupLoaded = $true
        $script:workspaceMoving = $false
        $script:savedTargets = ""
        $global:targetsKey = ""
        $global:signals = @{ settings = $false; status = $false }
        $global:currentStamp = "s0"
        $script:settingsStamp = "s0"
        $script:settingsOnNetwork = $false
        $script:settingsWatch = "settings"
        $script:statusWatch = "status"
        $script:lastStampCheck = Get-Date
        $global:indexingNow = $false
    }
    AfterAll {
        Remove-Variable -Name indexingNow, activatedLog, signals, targetsKey, currentStamp, shortcutHandled -Scope Global -ErrorAction SilentlyContinue
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

    It "移していなければ、フォルダの有無を調べる（守りが常に効いているのではない）" {
        onWindowActivated
        @($global:activatedLog) | Should -Contain "refreshFolderStatus"
    }

    It "合図なし・60 秒以内なら、設定も取り込み一覧も読み直さない" {
        onWindowActivated
        @($global:activatedLog) | Should -Not -Contain "getTargetFolders"
        @($global:activatedLog | Where-Object { $_ -like "refreshIndexingState*" }).Count | Should -Be 0
        @($global:activatedLog) | Should -Not -Contain "updateSearchTarget"
        @($global:activatedLog) | Should -Not -Contain "loadTargets"
    }

    It "<name>" -TestCases @(
        @{ name = "設定の合図ありで印が違い、一覧も違えば、読み直して対象を更新する"; stamp = "s1"; key = "x"; reload = $true; getCalls = 1 }
        @{ name = "設定の合図ありで印が違っても、一覧が同じなら読み直さない"; stamp = "s1"; key = ""; reload = $false; getCalls = 1 }
        @{ name = "設定の合図ありでも印が同じなら、getTargetFolders を呼ばない"; stamp = "s0"; key = "x"; reload = $false; getCalls = 0 }
    ) {
        $global:signals.settings = $true
        $global:currentStamp = $stamp
        $global:targetsKey = $key
        onWindowActivated
        @($global:activatedLog | Where-Object { $_ -eq "getTargetFolders" }).Count | Should -Be $getCalls
        @($global:activatedLog -contains "loadTargets") | Should -Be $reload
        @($global:activatedLog -contains "updateSearchTarget") | Should -Be $reload
    }

    It "取り込み一覧の合図があれば、印が同じなら飛ばす指定で集計を頼む" {
        $global:signals.status = $true
        onWindowActivated
        @($global:activatedLog) | Should -Contain "refreshIndexingState:True"
    }

    It "設定の見張りが無ければ（`$null）、合図なしでも印を比べる" {
        $script:settingsWatch = $null
        $global:currentStamp = "s1"
        $global:targetsKey = "x"
        onWindowActivated
        @($global:activatedLog) | Should -Contain "loadTargets"
    }

    It "取り込み一覧の見張りが無ければ（`$null）、合図なしでも集計を頼む" {
        $script:statusWatch = $null
        onWindowActivated
        @($global:activatedLog) | Should -Contain "refreshIndexingState:True"
    }

    It "60 秒過ぎていれば、合図なしでも設定の印を比べ、集計を頼む" {
        $script:lastStampCheck = (Get-Date).AddSeconds(-61)
        $global:currentStamp = "s1"
        $global:targetsKey = "x"
        onWindowActivated
        @($global:activatedLog) | Should -Contain "loadTargets"
        @($global:activatedLog) | Should -Contain "refreshIndexingState:True"
    }

    It "印を比べたら、60 秒の起点を書き直す（続けて比べない）" {
        $script:lastStampCheck = (Get-Date).AddSeconds(-61)
        onWindowActivated
        $global:activatedLog.Clear()
        onWindowActivated
        @($global:activatedLog | Where-Object { $_ -like "refreshIndexingState*" }).Count | Should -Be 0
    }

    It "更新中は、合図があっても集計を頼まない" {
        $global:signals.status = $true
        setIndexingNow $true
        onWindowActivated
        @($global:activatedLog | Where-Object { $_ -like "refreshIndexingState*" }).Count | Should -Be 0
    }

    It "設定ファイルがネットワークにあれば、画面のスレッドで印を取らず、60 秒過ぎたときだけ裏で調べる" {
        $script:settingsOnNetwork = $true
        $script:settingsWatch = $null
        onWindowActivated
        @($global:activatedLog) | Should -Not -Contain "startJob:network"
        $script:lastStampCheck = (Get-Date).AddSeconds(-61)
        onWindowActivated
        @($global:activatedLog) | Should -Contain "startJob:network"
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