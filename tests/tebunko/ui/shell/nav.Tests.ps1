# ナビ（左の欄）の画面の切り替えとキー操作の振り分け（tebunko\ui\shell\nav.ps1）のテスト。
# 画面のコントロールと、切り替えで呼ぶ関数は、呼ばれた順を $global:navLog に残す偽のものに差し替える。
# （本物のキーは、Gui テストが画面を別プロセスで動かすため送れない。キーを決める判断は nav_view.Tests.ps1）
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\shell\nav_view.ps1"

    function newFakeBox {
        param ([string]$name)
        $box = [pscustomobject]@{ Name = $name }
        $box | Add-Member ScriptMethod Focus { $global:navLog.Add("$($this.Name).Focus"); return $true }
        $box | Add-Member ScriptMethod SelectAll { $global:navLog.Add("$($this.Name).SelectAll") }
        $box | Add-Member ScriptMethod Start { $global:navLog.Add("$($this.Name).Start") }
        $box | Add-Member ScriptMethod Stop { $global:navLog.Add("$($this.Name).Stop") }
        return $box
    }

    function newFakeUi {
        $content = [pscustomobject]@{ Content = $null; Margin = $null }
        $content | Add-Member ScriptMethod UpdateLayout { $global:navLog.Add("UpdateLayout") }
        $navList = [pscustomobject]@{ SelectedItem = $null }
        $navList | Add-Member ScriptMethod Add_SelectionChanged { param($handler) }
        $fake = @{ NavPaneHost = [pscustomobject]@{ Visibility = "Collapsed" }; IndexFooterFolders = [pscustomobject]@{ Visibility = "Collapsed" }; IndexFooterFiles = [pscustomobject]@{ Visibility = "Collapsed" }; ContentHost = $content; NavList = $navList; WordBox = (newFakeBox "WordBox"); FilterBox = (newFakeBox "FilterBox") }
        foreach ($name in "SearchTab", "IndexTab", "SettingsTab") { $fake[$name] = [pscustomobject]@{ Name = $name } }
        return $fake
    }

    function safe { param ([scriptblock]$block) & $block }
    function refreshIndexingState { $global:navLog.Add("refreshIndexingState") }
    function refreshIndexSummary { $global:navLog.Add("refreshIndexSummary") }
    function loadIndexTree { $global:navLog.Add("loadIndexTree") }
    function cancelSearch { $global:navLog.Add("cancelSearch") }

    $global:navLog = New-Object System.Collections.Generic.List[string]
    $script:ui = newFakeUi
    . "${scriptsDir}\tebunko\ui\shell\nav.ps1"

    function resetNav {
        $global:navLog = New-Object System.Collections.Generic.List[string]
        $script:ui = newFakeUi
        $script:currentScreen = $null
        $script:startupLoaded = $true
        $script:search = $null
        $script:screenContents = @{ SearchTab = "search"; IndexTab = "index"; SettingsTab = "settings" }
    }
}

AfterAll {
    Remove-Variable -Name navLog -Scope Global -ErrorAction SilentlyContinue
}

Describe "selectScreen" -Tag Unit {
    BeforeEach { resetNav }

    It "左の欄の検索対象のツリー（NavPaneHost）は、検索の画面のときだけ出す" {
        selectScreen "SearchTab"
        $script:ui.NavPaneHost.Visibility | Should -Be "Visible"
        selectScreen "IndexTab"
        $script:ui.NavPaneHost.Visibility | Should -Be "Collapsed"
    }


    It "登録の件数（IndexFooterFolders・IndexFooterFiles）は、インデックス管理の画面のときだけステータスバーに出す" {
        selectScreen "IndexTab"
        $script:ui.IndexFooterFolders.Visibility | Should -Be "Visible"
        $script:ui.IndexFooterFiles.Visibility | Should -Be "Visible"
        selectScreen "SettingsTab"
        $script:ui.IndexFooterFolders.Visibility | Should -Be "Collapsed"
        $script:ui.IndexFooterFiles.Visibility | Should -Be "Collapsed"
    }


    It "選んだ画面の中身を ContentHost に差し、ナビの選択も合わせる" {
        selectScreen "SettingsTab"
        $ui.ContentHost.Content | Should -Be "settings"
        $ui.NavList.SelectedItem | Should -Be $ui["SettingsTab"]
        (getCurrentScreen) | Should -Be "SettingsTab"
    }

    It "画面ごとに ContentHost の余白を合わせる" {
        selectScreen "SearchTab"
        $ui.ContentHost.Margin | Should -Be "0,0,0,0"
        selectScreen "SettingsTab"
        $ui.ContentHost.Margin | Should -Be "16,12,16,12"
    }

    It "知らない名前は何もしない" {
        selectScreen "IndexTab"
        $global:navLog.Clear()
        selectScreen "Tabs"
        (getCurrentScreen) | Should -Be "IndexTab"
        $global:navLog.Count | Should -Be 0
    }

    It "いまと同じ画面は、中身も読み直しも時計も触らない" {
        selectScreen "IndexTab"
        $global:navLog.Clear()
        $ui.ContentHost.Content = "そのまま"
        selectScreen "IndexTab"
        $ui.ContentHost.Content | Should -Be "そのまま"
        $global:navLog.Count | Should -Be 0
    }

    It "起動の読み込みが済むまでは、読み直さず、時計にも触らない" {
        $script:startupLoaded = $false
        selectScreen "SettingsTab"
        (getCurrentScreen) | Should -Be "SettingsTab"
        $global:navLog.Count | Should -Be 0
    }

    It "インデックス管理へ移ると、状態を読み直す" {
        selectScreen "IndexTab"
        @($global:navLog) | Should -Be @("refreshIndexingState")
    }

    It "検索・設定へ移っても、何も読み直さない" {
        selectScreen "SearchTab"
        selectScreen "SettingsTab"
        $global:navLog.Count | Should -Be 0
    }
}

Describe "invokeShortcutAction" -Tag Unit {
    BeforeEach {
        resetNav
        Mock selectScreen { $global:navLog.Add("selectScreen $name") }
    }

    It "FocusSearchWord は、画面を切り替え、配置を済ませてから、検索ワードの欄にフォーカスして全選択する" {
        (invokeShortcutAction @{ Action = "FocusSearchWord"; Screen = "SearchTab" }) | Should -Be $true
        @($global:navLog) | Should -Be @("selectScreen SearchTab", "UpdateLayout", "WordBox.Focus", "WordBox.SelectAll")
    }

    It "FocusFilter は、画面を切り替え、配置を済ませてから、絞り込みの欄にフォーカスする" {
        (invokeShortcutAction @{ Action = "FocusFilter"; Screen = "SearchTab" }) | Should -Be $true
        @($global:navLog) | Should -Be @("selectScreen SearchTab", "UpdateLayout", "FilterBox.Focus")
    }

    It "SwitchScreen は、画面を切り替えるだけ" {
        (invokeShortcutAction @{ Action = "SwitchScreen"; Screen = "IndexTab" }) | Should -Be $true
        @($global:navLog) | Should -Be @("selectScreen IndexTab")
    }

    It "Refresh は、<screen> では <expected> を読み直す" -TestCases @(
        @{ screen = "IndexTab"; expected = @("refreshIndexingState", "refreshIndexSummary", "loadIndexTree") }
        @{ screen = "SettingsTab"; expected = @("refreshIndexingState", "refreshIndexSummary", "loadIndexTree") }
    ) {
        (invokeShortcutAction @{ Action = "Refresh"; Screen = $screen }) | Should -Be $true
        @($global:navLog) | Should -Be $expected
    }

    It "CancelSearch は、検索中だけ止めてキーを受ける。検索中でなければ何もせずキーを渡す" {
        (invokeShortcutAction @{ Action = "CancelSearch"; Screen = "SearchTab" }) | Should -Be $false
        $global:navLog.Count | Should -Be 0
        $script:search = [pscustomobject]@{ Running = $true }
        (invokeShortcutAction @{ Action = "CancelSearch"; Screen = "SearchTab" }) | Should -Be $true
        @($global:navLog) | Should -Be @("cancelSearch")
    }

    It "None は何もしない" {
        (invokeShortcutAction @{ Action = "None"; Screen = "" }) | Should -Be $false
        $global:navLog.Count | Should -Be 0
    }
}
