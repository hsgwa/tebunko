# 待つ間のダイアログ（showWaitDialog・updateWaitDialog・closeWaitDialog）を、本物の WPF の窓で確かめる。
# 画面の本体は起動せず、このテストのプロセスの中で窓を開く（STA）。親の窓は空の窓、暗幕は Border 1 枚で代わりにする。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\message_view.ps1"
    . "${scriptsDir}\shared\ui\app_host.ps1"
    . "${scriptsDir}\shared\ui\shell.ps1"

    ${sharedXamlDir} = "${scriptsDir}\shared\xaml"
    ${iconFile}      = "${scriptsDir}\tebunko\tebunko.ico"
    ${themeFile}     = "${sharedXamlDir}\theme.xaml"
    ${appTitle}      = "tebunko"
}

Describe "待つ間のダイアログ" -Tag Gui {
    BeforeEach {
        $script:owner = New-Object System.Windows.Window
        $script:owner.Width = 600
        $script:owner.Height = 400
        $script:owner.ShowInTaskbar = $false
        $script:owner.Show()
        $script:scrim = New-Object System.Windows.Controls.Border
        $script:scrim.Visibility = "Collapsed"
        setDialogScrim $script:scrim
        $script:dialogDepth = 0
        $script:handles = @()
    }
    AfterEach {
        foreach ($h in $script:handles) { closeWaitDialog $h }
        $script:owner.Close()
        setDialogScrim $null
    }

    It "出すと見出し・説明・棒・秒読み・押せないボタンが並び、暗幕が掛かる。閉じると暗幕が外れる" {
        $handle = showWaitDialog -heading "中止しています…" -description "説明" -countdown "あと 15 秒で強制終了します" -buttonText "閉じない" -owner $script:owner
        $script:handles += $handle
        $handle.Dialog.IsVisible | Should -BeTrue
        $handle.HeadingText.Text | Should -Be "中止しています…"
        $handle.DescriptionText.Visibility | Should -Be "Visible"
        $handle.CountdownText.Text | Should -Be "あと 15 秒で強制終了します"
        $handle.WaitBar.IsIndeterminate | Should -BeTrue
        $handle.WaitButton.Content | Should -Be "閉じない"
        $handle.WaitButton.IsEnabled | Should -BeFalse
        $script:scrim.Visibility | Should -Be "Visible"
        $handle.Dialog.Owner | Should -Be $script:owner
        $script:dialogDepth | Should -Be 1

        closeWaitDialog $handle
        $handle.Dialog.IsVisible | Should -BeFalse
        $script:scrim.Visibility | Should -Be "Collapsed"
        $script:dialogDepth | Should -Be 0
        { closeWaitDialog $handle } | Should -Not -Throw  # 2 回目は何もしない
    }

    It "利用者の閉じる操作では閉じない" {
        $handle = showWaitDialog -heading "中止しています…" -owner $script:owner
        $script:handles += $handle
        $handle.Dialog.Close()
        $handle.Dialog.IsVisible | Should -BeTrue
    }

    It "説明・秒読み・ボタンが空なら隠す。updateWaitDialog で棒と文字を書き換えられる" {
        $handle = showWaitDialog -heading "待っています" -owner $script:owner
        $script:handles += $handle
        $handle.DescriptionText.Visibility | Should -Be "Collapsed"
        $handle.CountdownText.Visibility | Should -Be "Collapsed"
        $handle.WaitButton.Visibility | Should -Be "Collapsed"

        updateWaitDialog $handle @{ Heading = "続けています"; Countdown = "あと 3 秒で強制終了します"; Indeterminate = $false; Value = 0.5 }
        $handle.HeadingText.Text | Should -Be "続けています"
        $handle.CountdownText.Visibility | Should -Be "Visible"
        $handle.WaitBar.IsIndeterminate | Should -BeFalse
        $handle.WaitBar.Value | Should -Be 0.5

        updateWaitDialog $handle @{ Countdown = "" }
        $handle.CountdownText.Visibility | Should -Be "Collapsed"
    }

    It "ほかのダイアログと重なっても、暗幕は全部閉じるまで残る" {
        $first = showWaitDialog -heading "1 つ目" -owner $script:owner
        $second = showWaitDialog -heading "2 つ目" -owner $script:owner
        $script:handles += @($first, $second)
        closeWaitDialog $second
        $script:scrim.Visibility | Should -Be "Visible"
        closeWaitDialog $first
        $script:scrim.Visibility | Should -Be "Collapsed"
    }

    It "showConfirm を重ねても、確認を閉じたあと待つ間のダイアログを閉じるまで暗幕が残る" {
        $handle = showWaitDialog -heading "待っています" -owner $script:owner
        $script:handles += $handle
        $real = (Get-Command showOwnedDialog -CommandType Function).ScriptBlock
        $script:seen = @{ Depth = -1; Scrim = ""; Owner = $null }
        $seen = $script:seen
        # 確認が開いたら、そのときの暗幕の数と親を控えて閉じる（本物の showOwnedDialog は、そのまま呼ぶ）
        function showOwnedDialog {
            param ($dialog)
            $seen = $script:seen
            $scrim = $script:scrim
            $getDepth = { $script:dialogDepth }
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromMilliseconds(200)
            $timer.Add_Tick({
                $timer.Stop()
                $seen.Depth = & $getDepth
                $seen.Scrim = [string]$scrim.Visibility
                $seen.Owner = $dialog.Owner
                $dialog.Close()
            }.GetNewClosure())
            $timer.Start()
            return (& $real $dialog)
        }
        $answer = showConfirm -heading "確認" -choices @(@{ Text = "実行する"; Value = "go" }) -owner $script:owner
        $answer | Should -BeNullOrEmpty
        $seen.Depth | Should -Be 2
        $seen.Scrim | Should -Be "Visible"
        $seen.Owner | Should -Be $script:owner
        $script:dialogDepth | Should -Be 1
        $script:scrim.Visibility | Should -Be "Visible"
        closeWaitDialog $handle
        $script:dialogDepth | Should -Be 0
        $script:scrim.Visibility | Should -Be "Collapsed"
    }

    It '閉じたあとの updateWaitDialog・$null の入れ物は何もしない' {
        $handle = showWaitDialog -heading "待っています" -owner $script:owner
        closeWaitDialog $handle
        { updateWaitDialog $handle @{ Heading = "x" } } | Should -Not -Throw
        { updateWaitDialog $null @{ Heading = "x" } } | Should -Not -Throw
        { closeWaitDialog $null } | Should -Not -Throw
    }
}
