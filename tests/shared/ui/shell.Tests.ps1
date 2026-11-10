# 画面の共通部品（scripts/shared/ui/shell.ps1）のテスト。特に、画面のスレッドの Dispatcher で
# 捕まえていない例外を受ける reportUnexpectedError・registerUnhandledErrorHandler を確かめる。
# 画面（$window）は出さず、$ui は偽物、showErrorDialog は Mock にする。CI は shell: powershell（5.1。STA）で
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
        # ステータスの文の出し先は、ツール側が登録する形（製品では tebunko/ui/shell/services.ps1）
        $statusText = $ui.StatusText
        setStatusSink { param ([string]$text) $statusText.Text = $text }.GetNewClosure()

        # 記録先は、製品が使う登録の口で、テスト用のフォルダにする
        setGuiErrorLogSource { Join-Path $TestDrive "gui_error_log.txt" } $null

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
            Mock showErrorDialog { }

            $done = pumpDispatcherUntilDone @({ throw "テストの例外" })

            $done | Should -BeTrue
            getErrorLogHeadingCount | Should -Be 1
            (Get-Content (getGuiErrorLogFile) -Raw) | Should -Match "テストの例外"
            $ui.StatusText.Text | Should -Be "エラーが発生しました：テストの例外"
            Should -Invoke showErrorDialog -Times 1 -Exactly -ParameterFilter {
                $heading -eq "予期しないエラーが起きました。" -and $detail -match "テストの例外" -and $logFile -match "gui_error_log.txt"
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

            Mock showErrorDialog { }

            & $handler $null (newFakeUnhandledArgs (& $Build))

            getErrorLogHeadingCount | Should -Be 1
            (Get-Content (getGuiErrorLogFile) -Raw) | Should -Match ([regex]::Escape($Message))
            $ui.StatusText.Text | Should -Be "エラーが発生しました：$Message"
            Should -Invoke showErrorDialog -Times 1 -Exactly
        }

        It "同じ例外を繰り返し知らせない（ダイアログは 1 回・記録は 60 秒に 1 回）" {
            Mock showErrorDialog { }
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
            Should -Invoke showErrorDialog -Times 1 -Exactly

            $script:fakeNow = $script:fakeNow.AddSeconds(61)
            raiseSameException

            getErrorLogHeadingCount | Should -Be 2
            (Get-Content (getGuiErrorLogFile) -Raw) | Should -Match "同じ例外が2回起きました"
            Should -Invoke showErrorDialog -Times 1 -Exactly
        }

        It "ダイアログを出している間に起きた別の例外は、記録されるがダイアログを出さない" {
            Mock showErrorDialog {
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
            Should -Invoke showErrorDialog -Times 1 -Exactly
        }

        It "見分けの表は 100 件まで。古いキーが消えたら、その例外はまた 1 回だけダイアログを出す" {
            Mock showErrorDialog { }

            function raiseNumberedException([int]$n) {
                try { throw "例外$n" } catch {
                    & $handler $null (newFakeUnhandledArgs $_.Exception)
                }
            }

            raiseNumberedException 0
            for ($i = 1; $i -le 100; $i++) {
                raiseNumberedException $i
            }
            # ここまでで 101 種類の例外。上限は 100 件のため、最初の「例外0」のキーはすでに消えている
            Should -Invoke showErrorDialog -Times 101 -Exactly

            raiseNumberedException 0

            Should -Invoke showErrorDialog -Times 102 -Exactly
        }

        It "知らせる処理（ダイアログの表示）が例外を投げても、Dispatcher の外へ例外が出ない" {
            Mock showErrorDialog { throw "showErrorDialog 自体の例外" }

            $done = pumpDispatcherUntilDone @({ throw "テストの例外" })

            $done | Should -BeTrue
        }
    }

    Describe "safe（動きが変わらないこと）" {
        It "記録・ステータス・ダイアログの文言が今までと同じで、同じ例外でも毎回ダイアログを出す" {
            Mock showErrorDialog { }

            safe { throw "x" }
            safe { throw "x" }

            getErrorLogHeadingCount | Should -Be 2
            $ui.StatusText.Text | Should -Be "エラーが発生しました：x"
            Should -Invoke showErrorDialog -Times 2 -Exactly -ParameterFilter {
                $heading -eq "予期しないエラーが起きました。" -and $detail -match ": x$" -and $logFile -match "gui_error_log.txt"
            }
        }
    }

    Describe "setStatus（出し先の登録）" {
        It "登録した出し先に文を渡す" {
            setStatus "確かめ中です"
            $ui.StatusText.Text | Should -Be "確かめ中です"
        }

        It "出し先を登録していなければ、何もせず例外も出さない" {
            $saved = $script:statusSink
            try {
                $script:statusSink = $null
                { setStatus "捨てられる" } | Should -Not -Throw
                $ui.StatusText.Text | Should -Be ""
            } finally {
                $script:statusSink = $saved
            }
        }
    }

    Describe "確認ダイアログの形" {
        It "form=<form> 選択肢 <count> 件・Danger=<danger> は <expected>" -ForEach @(
            @{ form = "";       count = 1; danger = $false; expected = "normal" }
            @{ form = "";       count = 1; danger = $true;  expected = "danger" }
            @{ form = "";       count = 2; danger = $false; expected = "choice" }
            @{ form = "";       count = 0; danger = $false; expected = "normal" }
            @{ form = "danger"; count = 2; danger = $false; expected = "danger" }
        ) {
            $choices = @(for ($i = 1; $i -le $count; $i++) { @{ Text = "x"; Value = "v$i"; Danger = $danger } })
            getConfirmForm $choices $form | Should -Be $expected
        }

        It "幅は <form> が <width>" -ForEach @(
            @{ form = "normal";   width = 520 }
            @{ form = "danger";   width = 520 }
            @{ form = "progress"; width = 520 }
            @{ form = "error";    width = 520 }
            @{ form = "choice";   width = 620 }
        ) {
            getConfirmWidth $form | Should -Be $width
        }
    }

    Describe "showOwnedDialog（暗幕）" {
        BeforeEach {
            # 偽の暗幕（Visibility を持つ）と、閉じるまでの間の暗幕の状態を覚える偽のダイアログ
            $scrim = [pscustomobject]@{ Visibility = "Collapsed" }
            setDialogScrim $scrim
        }
        AfterAll {
            setDialogScrim $null
        }

        It "開いている間は Visible、閉じた後は Collapsed に戻り、結果を返す" {
            $dialog = [pscustomobject]@{ Seen = $null }
            $dialog | Add-Member -MemberType ScriptMethod -Name ShowDialog -Value { $this.Seen = $scrim.Visibility; return $true }
            showOwnedDialog $dialog | Should -BeTrue
            $dialog.Seen | Should -Be "Visible"
            $scrim.Visibility | Should -Be "Collapsed"
        }

        It "ShowDialog が例外を投げても Collapsed に戻る" {
            $dialog = [pscustomobject]@{}
            $dialog | Add-Member -MemberType ScriptMethod -Name ShowDialog -Value { throw "開けない" }
            { showOwnedDialog $dialog } | Should -Throw "*開けない*"
            $scrim.Visibility | Should -Be "Collapsed"
        }

        It "暗幕を覚えていなければ、暗幕には触れずにダイアログだけ出す" {
            setDialogScrim $null
            $dialog = [pscustomobject]@{}
            $dialog | Add-Member -MemberType ScriptMethod -Name ShowDialog -Value { return $false }
            showOwnedDialog $dialog | Should -BeFalse
            $scrim.Visibility | Should -Be "Collapsed"
        }
    }

    Context "startUiWork（画面のスレッドの分割）" {
        BeforeAll {
            function tickRunning {
                # 動いているタイマーの Tick を手で 1 回ずつ呼ぶ
                foreach ($fakeTimer in $script:fakeTimers.ToArray()) {
                    if ($fakeTimer.Running) { & $fakeTimer.OnTick $fakeTimer $null }
                }
            }
            function newTimer {
                # タイマーは動かさない偽物。Tick の処理は取っておき、tickRunning が手で呼ぶ
                param ([int]$milliseconds, [scriptblock]$onTick)
                $timer = [pscustomobject]@{ Tag = $null; Interval = $milliseconds; OnTick = $onTick; Running = $false }
                $timer | Add-Member ScriptMethod Start { $this.Running = $true }
                $timer | Add-Member ScriptMethod Stop { $this.Running = $false }
                $script:fakeTimers.Add($timer)
                return $timer
            }
            function runningTimerCount { ($script:fakeTimers.ToArray() | Where-Object { $_.Running } | Measure-Object).Count }
        }

        BeforeEach {
            stopAllUiWork
            $script:fakeTimers = New-Object System.Collections.Generic.List[object]
            $script:seen = New-Object System.Collections.Generic.List[object]
            $script:dones = New-Object System.Collections.Generic.List[object]
            $script:stops = 0
            $script:clock = 0
            $script:clockStep = 25   # 1 回の予算（60）で 3 件回る（25・50・75 の 3 件目で予算切れ）
            Mock getUiWorkMilliseconds { $script:clock += $script:clockStep; $script:clock }
            Mock reportUnexpectedError { }
        }

        AfterEach {
            stopAllUiWork
        }

        It "予算内に収まる件数は、その場で全部回り、OnDone が 1 回呼ばれ、タイマーは作らない" {
            $script:clockStep = 1

            startUiWork -Name "t" -Items @(1, 2, 3) -OnItem { param ($item, $state) $script:seen.Add($item) } -OnDone { param ($state) $script:dones.Add($state) }

            ($script:seen -join ",") | Should -Be "1,2,3"
            $script:dones.Count | Should -Be 1
            $script:fakeTimers.Count | Should -Be 0
            testUiWorkRunning "t" | Should -BeFalse
        }

        It "予算を超える件数は、1 回目で途中まで回り、タイマーの 1 回ごとに続き、全部順に回ったあと OnDone が 1 回呼ばれる" {
            startUiWork -Name "t" -Items @(1..7) -OnItem { param ($item, $state) $script:seen.Add($item) } -OnDone { param ($state) $script:dones.Add($state) }

            ($script:seen -join ",") | Should -Be "1,2,3"
            $script:dones.Count | Should -Be 0
            testUiWorkRunning "t" | Should -BeTrue
            $script:fakeTimers.Count | Should -Be 1
            $script:fakeTimers[0].Interval | Should -Be 100
            $script:fakeTimers[0].Tag | Should -Be "t"
            $script:fakeTimers[0].Running | Should -BeTrue

            tickRunning
            ($script:seen -join ",") | Should -Be "1,2,3,4,5,6"
            $script:dones.Count | Should -Be 0
            tickRunning

            ($script:seen -join ",") | Should -Be "1,2,3,4,5,6,7"
            $script:dones.Count | Should -Be 1
            $script:fakeTimers[0].Running | Should -BeFalse
            testUiWorkRunning "t" | Should -BeFalse
        }

        It "-State の値が OnBegin・OnItem・OnDone に渡る（OnBegin は 1 回だけ、回し始める前）" {
            $state = @{ Mark = "印"; Begun = 0 }

            startUiWork -Name "t" -Items @(1..7) -State $state `
                -OnBegin { param ($s) $s.Begun++; $script:seen.Add("begin:$($s.Mark)") } `
                -OnItem { param ($item, $s) $script:seen.Add("$($s.Mark)$item") } `
                -OnDone { param ($s) $script:dones.Add($s) }
            completeUiWork "t"

            $script:seen[0] | Should -Be "begin:印"
            $script:seen[1] | Should -Be "印1"
            $state.Begun | Should -Be 1
            [object]::ReferenceEquals($script:dones[0], $state) | Should -BeTrue
        }

        It "Restart: 呼び直すと、前の OnDone は呼ばれず、前の残りも回らない" {
            startUiWork -Name "t" -Items @(1..7) -OnItem { param ($item, $state) $script:seen.Add("a$item") } -OnDone { param ($state) $script:dones.Add("a") }

            startUiWork -Name "t" -Items @(1..2) -OnItem { param ($item, $state) $script:seen.Add("b$item") } -OnDone { param ($state) $script:dones.Add("b") }
            completeUiWork "t"

            ($script:seen -join ",") | Should -Be "a1,a2,a3,b1,b2"
            ($script:dones -join ",") | Should -Be "b"
        }

        It "Requeue: 回っている間にタイマーの 1 回ごとに呼び直し続けても、今の分は k 回以内のタイマーで OnDone まで進み、呼び直しが止んだあと 2k + 1 回以内で終わり、最後の State が渡る" {
            $k = 3   # 7 件は 1 回に 3 件ずつで 3 回（開始 + タイマー 2 回）で回る
            $onItem = { param ($item, $state) }
            $onDone = { param ($state) $script:dones.Add($state.N) }
            startUiWork -Name "t" -Items @(1..7) -OnItem $onItem -OnDone $onDone -State @{ N = 0 } -WhenBusy Requeue
            $ticks = 0
            $firstDoneTick = -1
            $calls = 0
            while ((testUiWorkRunning "t") -and $ticks -lt 40 -and $calls -lt 6) {
                $calls++
                startUiWork -Name "t" -Items @(1..7) -OnItem $onItem -OnDone $onDone -State @{ N = $calls } -WhenBusy Requeue
                tickRunning
                $ticks++
                if ($firstDoneTick -lt 0 -and $script:dones.Count -gt 0) { $firstDoneTick = $ticks }
            }
            $afterTicks = 0
            while ((testUiWorkRunning "t") -and $afterTicks -lt 40) {
                tickRunning
                $afterTicks++
            }

            $firstDoneTick | Should -BeGreaterThan 0
            $firstDoneTick | Should -BeLessOrEqual $k
            $script:dones[0] | Should -Be 0
            testUiWorkRunning "t" | Should -BeFalse
            $afterTicks | Should -BeLessOrEqual (2 * $k + 1)
            $script:dones[$script:dones.Count - 1] | Should -Be $calls
        }

        It "Requeue: 控えから回し始めるとき、OnBegin をその時に呼ぶ" {
            $onBegin = { param ($s) $script:seen.Add("begin$($s.N)") }
            $onItem = { param ($item, $s) }
            startUiWork -Name "t" -Items @(1..7) -OnItem $onItem -OnBegin $onBegin -State @{ N = 1 } -WhenBusy Requeue
            startUiWork -Name "t" -Items @(1..7) -OnItem $onItem -OnBegin $onBegin -State @{ N = 2 } -WhenBusy Requeue
            ($script:seen -join ",") | Should -Be "begin1"

            completeUiWork "t"

            ($script:seen -join ",") | Should -Be "begin1,begin2"
        }

        It "stopUiWork・stopAllUiWork のあとは、何も回らない（控えも回らない）" -TestCases @(
            @{ how = "stopUiWork"; stop = { stopUiWork "t" } }
            @{ how = "stopAllUiWork"; stop = { stopAllUiWork } }
        ) {
            $onItem = { param ($item, $state) $script:seen.Add($item) }
            $onDone = { param ($state) $script:dones.Add(1) }
            startUiWork -Name "t" -Items @(1..7) -OnItem $onItem -OnDone $onDone -WhenBusy Requeue
            startUiWork -Name "t" -Items @(1..7) -OnItem $onItem -OnDone $onDone -WhenBusy Requeue

            & $stop
            tickRunning
            completeUiWork "t"

            ($script:seen -join ",") | Should -Be "1,2,3"
            $script:dones.Count | Should -Be 0
            runningTimerCount | Should -Be 0
            testUiWorkRunning "t" | Should -BeFalse
        }

        It "OnItem の中で同じ名前の仕事を捨てると（stopUiWork・Restart）、捨てた仕事の残りの項目はその 1 件で止まる" -TestCases @(
            @{ how = "stopUiWork"; act = { stopUiWork "t" } }
            @{ how = "Restart"; act = { startUiWork -Name "t" -Items @(9) -OnItem { param ($item, $state) $script:seen.Add("new$item") } -OnDone { param ($state) $script:dones.Add("new") } } }
        ) {
            $script:clockStep = 1   # 予算では止まらない
            $script:act = $act
            startUiWork -Name "t" -Items @(1..5) -OnItem { param ($item, $state) $script:seen.Add($item); if ($item -eq 2) { & $script:act } } -OnDone { param ($state) $script:dones.Add("old") }

            $seenText = $script:seen -join ","
            if ($how -eq "stopUiWork") { $seenText | Should -Be "1,2" } else { $seenText | Should -Be "1,2,new9" }
            $script:dones.Contains("old") | Should -BeFalse
        }

        It "completeUiWork で、残りと控えが全部回り、OnDone が呼ばれる" {
            $onItem = { param ($item, $state) $script:seen.Add($item) }
            $onDone = { param ($state) $script:dones.Add(1) }
            startUiWork -Name "t" -Items @(1..7) -OnItem $onItem -OnDone $onDone -WhenBusy Requeue
            startUiWork -Name "t" -Items @(11..12) -OnItem $onItem -OnDone $onDone -WhenBusy Requeue

            completeUiWork "t"

            ($script:seen -join ",") | Should -Be "1,2,3,4,5,6,7,11,12"
            $script:dones.Count | Should -Be 2
            runningTimerCount | Should -Be 0
        }

        It "OnItem が例外を投げると、reportUnexpectedError と OnStop が呼ばれ、その仕事は止まる" {
            startUiWork -Name "t" -Items @(1..7) -OnItem { param ($item, $state) if ($item -eq 5) { throw "壊れた" }; $script:seen.Add($item) } `
                -OnDone { param ($state) $script:dones.Add(1) } -OnStop { param ($state) $script:stops++ }

            tickRunning

            ($script:seen -join ",") | Should -Be "1,2,3,4"
            Should -Invoke reportUnexpectedError -Times 1 -Exactly
            $script:stops | Should -Be 1
            $script:dones.Count | Should -Be 0
            testUiWorkRunning "t" | Should -BeFalse
            runningTimerCount | Should -Be 0
        }

        It "OnDone が例外を投げても、reportUnexpectedError と OnStop が呼ばれ、控えは回らない" {
            $script:clockStep = 1
            startUiWork -Name "t" -Items @(1..7) -OnItem { param ($item, $state) } -OnDone { param ($state) throw "壊れた" } -OnStop { param ($state) $script:stops++ }

            Should -Invoke reportUnexpectedError -Times 1 -Exactly
            $script:stops | Should -Be 1
            testUiWorkRunning "t" | Should -BeFalse
        }

        It "呼び元が渡した一覧をあとで変えても、回る中身は変わらない" {
            $items = @(1..7)
            startUiWork -Name "t" -Items $items -OnItem { param ($item, $state) $script:seen.Add($item) }

            $items[5] = 99
            completeUiWork "t"

            ($script:seen -join ",") | Should -Be "1,2,3,4,5,6,7"
        }

        It "Items が空でも OnDone は呼ばれる" {
            startUiWork -Name "t" -Items @() -OnItem { param ($item, $state) } -OnDone { param ($state) $script:dones.Add(1) }

            $script:dones.Count | Should -Be 1
        }
    }
}
