# 画面の共通部品（scripts/shared/ui/shell.ps1）のテスト。特に、画面のスレッドの Dispatcher で
# 捕まえていない例外を受ける reportUnexpectedError・registerUnhandledErrorHandler を確かめる。
# 画面（$window）は出さず、$ui は偽物、showMessage は Mock にする。CI は shell: powershell（5.1。STA）で
# 動くため、今のスレッドの Dispatcher が使える。
Describe "shell.ps1" -Tag Unit {
    BeforeAll {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

        $here       = (Resolve-Path "$PSScriptRoot\..\..").Path
        $scriptsDir = (Resolve-Path "$here\..\scripts").Path
        . "$scriptsDir\shared\ui\app_host.ps1"
        . "$scriptsDir\shared\ui\shell.ps1"

        $dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher

        # ---- 画面の偽物 ----
        $ui = @{ StatusText = [pscustomobject]@{ Text = ""; ToolTip = "" } }

        function getGuiErrorLogFile {
            Join-Path $TestDrive "画面エラー.txt"
        }

        # Dispatcher に積んだ action を実行させ、終わるまで待つ（保険の 5 秒で必ず抜ける。入れ子でも使う）。
        # 戻り値は「保険で抜けたのではなく、積んだ処理が最後まで終わったか」
        function pumpDispatcherUntilDone {
            param (
                [scriptblock[]]$actions
            )

            $frame = New-Object System.Windows.Threading.DispatcherFrame
            $done = @{ Value = $false }
            $watchdog = New-Object System.Windows.Threading.DispatcherTimer
            $watchdog.Interval = [TimeSpan]::FromSeconds(5)
            $watchdog.Add_Tick({ $frame.Continue = $false }.GetNewClosure())
            $watchdog.Start()

            foreach ($action in $actions) {
                $dispatcher.BeginInvoke([action]$action) | Out-Null
            }
            $dispatcher.BeginInvoke([action]{ $done.Value = $true; $frame.Continue = $false }.GetNewClosure()) | Out-Null

            [System.Windows.Threading.Dispatcher]::PushFrame($frame)
            $watchdog.Stop()
            return $done.Value
        }

        # 記録の見出し行（==== ... ====）の数
        function getErrorLogHeadingCount {
            if (!(Test-Path (getGuiErrorLogFile))) {
                return 0
            }
            @(Get-Content (getGuiErrorLogFile) | Where-Object { $_ -match "^====" }).Count
        }

        # 実際に届く例外と同じ形（Handled を立てられる・Exception を持つ）の偽物の DispatcherUnhandledExceptionEventArgs
        function newFakeUnhandledArgs([System.Exception]$exception) {
            [pscustomobject]@{ Exception = $exception; Handled = $false }
        }
    }

    BeforeEach {
        $ui.StatusText.Text = ""
        $ui.StatusText.ToolTip = ""
        if (Test-Path (getGuiErrorLogFile)) {
            Remove-Item (getGuiErrorLogFile) -Force
        }
        $script:unhandledErrorTable = [ordered]@{}
        $script:unhandledDialogShowing = $false
        $script:reportUnexpectedErrorNow = { Get-Date }
        $handler = registerUnhandledErrorHandler $dispatcher
    }

    AfterEach {
        $dispatcher.Remove_UnhandledException($handler)
    }

    Describe "registerUnhandledErrorHandler" {
        It "例外が Dispatcher の外へ出ず、記録・ステータス・ダイアログで 1 回知らせる" {
            Mock showMessage { }

            $done = pumpDispatcherUntilDone @({ throw "テストの例外" })

            $done | Should -BeTrue
            getErrorLogHeadingCount | Should -Be 1
            (Get-Content (getGuiErrorLogFile) -Raw) | Should -Match "テストの例外"
            $ui.StatusText.Text | Should -Be "エラーが発生しました：テストの例外"
            Should -Invoke showMessage -Times 1 -Exactly -ParameterFilter {
                $message -eq "エラーが発生しました。`nテストの例外`n`n詳しい内容は 画面エラー.txt に残しています。"
            }
        }

        It "例外の種類（<Label>）によらず、中の例外のメッセージで記録・表示する" -TestCases @(
            @{ Label = "スクリプトの throw"; Message = "テストの例外A"; Build = {
                try { throw "テストの例外A" } catch { $_.Exception }
            } }
            @{ Label = ".NET の例外"; Message = "WPF の例外"; Build = {
                [System.InvalidOperationException]::new("WPF の例外")
            } }
            @{ Label = "TargetInvocationException に包んだ例外"; Message = "中の例外"; Build = {
                [System.Reflection.TargetInvocationException]::new([System.InvalidOperationException]::new("中の例外"))
            } }
        ) {
            param ($Label, $Message, $Build)

            Mock showMessage { }

            & $handler $null (newFakeUnhandledArgs (& $Build))

            getErrorLogHeadingCount | Should -Be 1
            (Get-Content (getGuiErrorLogFile) -Raw) | Should -Match ([regex]::Escape($Message))
            $ui.StatusText.Text | Should -Be "エラーが発生しました：$Message"
            Should -Invoke showMessage -Times 1 -Exactly
        }

        It "同じ例外を繰り返し知らせない（ダイアログは 1 回・記録は 60 秒に 1 回）" {
            Mock showMessage { }
            $script:reportUnexpectedErrorNow = { $script:fakeNow }
            $script:fakeNow = Get-Date

            function raiseSameException {
                try { throw "繰り返す例外" } catch {
                    & $handler $null (newFakeUnhandledArgs $_.Exception)
                }
            }

            raiseSameException
            raiseSameException
            raiseSameException

            getErrorLogHeadingCount | Should -Be 1
            Should -Invoke showMessage -Times 1 -Exactly

            $script:fakeNow = $script:fakeNow.AddSeconds(61)
            raiseSameException

            getErrorLogHeadingCount | Should -Be 2
            (Get-Content (getGuiErrorLogFile) -Raw) | Should -Match "同じ例外が2回起きました"
            Should -Invoke showMessage -Times 1 -Exactly
        }

        It "ダイアログを出している間に起きた別の例外は、記録されるがダイアログを出さない" {
            Mock showMessage {
                # MessageBox が表示中もメッセージを回すのと同じ状況を、入れ子の DispatcherFrame で作る
                $done = pumpDispatcherUntilDone @({
                    try { throw "重なった別の例外" } catch {
                        & $handler $null (newFakeUnhandledArgs $_.Exception)
                    }
                })
                $done | Should -BeTrue
            }

            try { throw "最初の例外" } catch {
                & $handler $null (newFakeUnhandledArgs $_.Exception)
            }

            getErrorLogHeadingCount | Should -Be 2
            Should -Invoke showMessage -Times 1 -Exactly
        }

        It "知らせる処理（ダイアログの表示）が例外を投げても、Dispatcher の外へ例外が出ない" {
            Mock showMessage { throw "showMessage 自体の例外" }

            $done = pumpDispatcherUntilDone @({ throw "テストの例外" })

            $done | Should -BeTrue
        }
    }

    Describe "safe（動きが変わらないこと）" {
        It "記録・ステータス・showMessage の文言が今までと同じで、同じ例外でも毎回ダイアログを出す" {
            Mock showMessage { }

            safe { throw "x" }
            safe { throw "x" }

            getErrorLogHeadingCount | Should -Be 2
            $ui.StatusText.Text | Should -Be "エラーが発生しました：x"
            Should -Invoke showMessage -Times 2 -Exactly -ParameterFilter {
                $message -eq "エラーが発生しました。`nx`n`n詳しい内容は 画面エラー.txt に残しています。"
            }
        }
    }
}
