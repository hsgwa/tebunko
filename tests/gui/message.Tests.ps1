# メッセージの画面（showMessage）を、種類ごとに本物の WPF の窓で確かめる（アイコン・ボタン・既定のボタン・戻り値・Esc）。
# 画面の本体は起動せず、このテストのプロセスの中で窓を開く（STA の Dispatcher を使う）。窓は DispatcherTimer が押して閉じる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\message_view.ps1"
    . "${scriptsDir}\shared\ui\app_host.ps1"
    . "${scriptsDir}\shared\ui\shell.ps1"

    ${sharedXamlDir} = "${scriptsDir}\shared\xaml"
    ${fontsDir}      = "${scriptsDir}\shared\fonts"
    ${iconFile}      = "${scriptsDir}\tebunko\tebunko.ico"
    ${themeFile}     = "${sharedXamlDir}\theme.xaml"
    ${appTitle}      = "tebunko"
    $window = $null

    # タイマーのクロージャからも書き換えられるよう、入れ物（ハッシュ表）を 1 つ共有する
    $script:state = @{ Seen = @{}; Press = ""; Esc = $false; Error = "" }

    # 窓が開いたら（ShowDialog の直前）、見える形を控え、Esc か指定のボタンを押すタイマーを掛ける
    function showOwnedDialog {
        param ($dialog)
        $state = $script:state
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(300)
        $timer.Add_Tick({
            $timer.Stop()
            try {
                $iconHost = $dialog.FindName("HeadingIconHost")
                $buttons = @($dialog.FindName("ButtonPanel").Children)
                $state.Seen = @{
                    Icon    = if ($iconHost.Visibility -eq "Visible") { [System.Windows.Automation.AutomationProperties]::GetName($iconHost) } else { "" }
                    Heading = $dialog.FindName("HeadingText").Text
                    Buttons = @($buttons | ForEach-Object { [string]$_.Content })
                    Default = (@($buttons | Where-Object { $_.IsDefault } | ForEach-Object { [string]$_.Content }) -join ",")
                }
                $target = if ($state.Esc) { $buttons | Where-Object { $_.IsCancel } | Select-Object -First 1 }
                          elseif ($state.Press -ne "") { $buttons | Where-Object { [string]$_.Content -eq $state.Press } | Select-Object -First 1 }
                          else { $buttons | Where-Object { $_.IsDefault } | Select-Object -First 1 }
                if ($null -eq $target) {
                    # Esc で押せる［キャンセル］が無い形（はい／いいえ）は、窓を閉じたときの戻り値を確かめる
                    $dialog.Close()
                    return
                }
                $target.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
            } catch {
                $state.Error = "$_"
                $dialog.Close()  # 失敗してもテストが止まらないよう閉じる
            }
        }.GetNewClosure())
        $timer.Start()
        return $dialog.ShowDialog()
    }
}

Describe "メッセージの画面（showMessage）" -Tag Gui {
    BeforeEach { $script:state.Error = "" }
    AfterEach { $script:state.Error | Should -Be "" }  # 窓の中の確かめで例外が出ていないこと

    # 種類ごとのアイコンの名前・ボタン・既定のボタン・押したときの戻り値
    It "<Icon>・<Buttons>: アイコン「<Kind>」・ボタン「<Shown>」・既定「<Default>」・<Press> で <Expect>" -TestCases @(
        @{ Icon = "Information"; Buttons = "OK";          Kind = "お知らせ"; Shown = "OK";                  Default = "OK";     Press = "OK";         Expect = "OK" }
        @{ Icon = "Warning";     Buttons = "OK";          Kind = "警告";     Shown = "OK";                  Default = "OK";     Press = "OK";         Expect = "OK" }
        @{ Icon = "Error";       Buttons = "OK";          Kind = "エラー";   Shown = "OK";                  Default = "OK";     Press = "OK";         Expect = "OK" }
        @{ Icon = "Question";    Buttons = "YesNo";       Kind = "確認";     Shown = "はい,いいえ";         Default = "はい";   Press = "いいえ";     Expect = "No" }
        @{ Icon = "Warning";     Buttons = "OKCancel";    Kind = "警告";     Shown = "OK,キャンセル";       Default = "OK";     Press = "キャンセル"; Expect = "Cancel" }
        @{ Icon = "Question";    Buttons = "YesNoCancel"; Kind = "確認";     Shown = "はい,いいえ,キャンセル"; Default = "はい"; Press = "はい";       Expect = "Yes" }
    ) {
        param ($Icon, $Buttons, $Kind, $Shown, $Default, $Press, $Expect)
        $script:state.Press = $Press; $script:state.Esc = $false
        $result = showMessage "見出し`n`n補足" $Buttons $Icon
        [string]$result | Should -Be $Expect
        $script:state.Seen.Icon | Should -Be $Kind
        ($script:state.Seen.Buttons -join ",") | Should -Be $Shown
        $script:state.Seen.Default | Should -Be $Default
        $script:state.Seen.Heading | Should -Be "見出し"
    }

    It "既定のボタンを指定すると、そのボタンが Enter で押される（<Buttons> の <DefaultName> は <Expect>）" -TestCases @(
        @{ Buttons = "YesNo";    DefaultName = "No";     Expect = "No" }
        @{ Buttons = "OKCancel"; DefaultName = "Cancel"; Expect = "Cancel" }
    ) {
        param ($Buttons, $DefaultName, $Expect)
        $script:state.Press = ""; $script:state.Esc = $false
        [string](showMessage "見出し" $Buttons "Question" $DefaultName) | Should -Be $Expect
    }

    It "Esc か窓を閉じたときの戻り値は、［OK］だけなら OK、はい／いいえなら No、それ以外は Cancel" -TestCases @(
        @{ Buttons = "OK";          Expect = "OK" }
        @{ Buttons = "YesNo";       Expect = "No" }
        @{ Buttons = "OKCancel";    Expect = "Cancel" }
        @{ Buttons = "YesNoCancel"; Expect = "Cancel" }
    ) {
        param ($Buttons, $Expect)
        $script:state.Press = ""; $script:state.Esc = $true
        [string](showMessage "見出し" $Buttons "Warning") | Should -Be $Expect
    }

}
