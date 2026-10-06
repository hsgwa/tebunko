# 起動時の「前回残った Office の確認」の画面（画面層）。gui_main.ps1 が読み込む。
# 文言・出すかどうか・対象の選び方は判断層（leftover_view.ps1）、プロセスの一覧・終了は shared/office/office_process.ps1。
# ここは、それらをつなぎ、確認のダイアログ（dialog_leftover.xaml）とステータスの文を出すだけ。
# 流れ: 起動時の読み込みの終わりに、記録の読み取り（getOfficeProcesses）を裏の仕事で行う → 出す時機を決める（getLeftoverPromptTiming）→
#       確認 → ［終了する］なら、PID を読み直して確認に出したものだけを選び（getLeftoverTargets）、裏の仕事で終了する → ステータスに結果を出す。
# 止めるのは、起動のときに記録した PID の Office だけ（名前で探して止めない）。画面のスレッドでは記録を読まず、プロセスにも触れない。

$script:leftoverNoticesOpen = $false   # 起動時のお知らせ（メッセージボックス）を開いている間は true。閉じたら resumeLeftoverPrompt で決め直す
$script:leftoverPending = $null        # お知らせが閉じるのを待っている、確認に出す予定の残り物（getOfficeProcesses の要素）
$script:leftoverFakeFile = $env:TEBUNKO_GUI_LEFTOVER_FILE   # 写真・画面のテストの場面から偽の行を差し込むときだけ入る（leftover_view.ps1 の「偽の行」）

function startLeftoverCheck {
    # 起動時の読み込みの終わりに呼ぶ。残った Office の記録を裏で読み、確認を出すか決める
    $mode = getLeftoverSourceMode $script:leftoverFakeFile
    startJob {
        param ($recordDir, $fakeFile)
        if ($fakeFile) {
            $parsed = Get-Content -LiteralPath $fakeFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $rows = @($parsed | ForEach-Object { $_ })   # PowerShell 5.1 の ConvertFrom-Json は配列を 1 つの値で返すため、1 段ほどく
            return @{ Processes = $rows }
        }
        return @{ Processes = @(getOfficeProcesses $recordDir) }
    } @((getOfficePidDir $workspace), $(if ($mode -eq "Fake") { $script:leftoverFakeFile } else { "" })) {
        param ($output, $errorText)
        if ($errorText -or $null -eq $output -or $output.Count -eq 0) {
            if ($errorText) { setStatus (getLeftoverFailureText "Check" $errorText) }
            return
        }
        safe {
            $processes = @($output[0].Processes)
            if ((getLeftoverSourceMode $script:leftoverFakeFile) -eq "Fake") {
                $processes = @(convertLeftoverFakeRows $processes)
            }
            decideLeftoverPrompt $processes
        }
    } (getOfficePidQueue $workspace.Dir)
}

function decideLeftoverPrompt {
    # 残り物を確認に出すか、見送るか、お知らせが閉じるまで待つかを決め、出すときは出す
    param (
        [object[]]$processes
    )

    $prompt = getLeftoverPrompt $processes
    $script:leftoverPending = $null
    if ($null -eq $prompt) {
        return
    }
    $closing = [bool]$script:closeWaiting -or [bool]$script:closeReady
    $timing = getLeftoverPromptTiming $prompt.Count $closing (isIndexing) $script:leftoverNoticesOpen @($window.OwnedWindows).Count
    if ($timing -eq "Defer") {
        $script:leftoverPending = $processes
        return
    }
    if ($timing -ne "Show") {
        return
    }
    $answer = showLeftoverDialog $prompt
    if ($answer -ne "stop") {
        return
    }
    stopLeftoverProcesses $prompt
}

function resumeLeftoverPrompt {
    # 起動時のお知らせが閉じたあとに、待っていた確認をもう一度決める
    $pending = $script:leftoverPending
    if ($null -ne $pending) {
        decideLeftoverPrompt $pending
    }
}

function showLeftoverDialog {
    # 確認のダイアログを出す。［終了する］を選べば "stop"、それ以外（今回は終了しない・閉じる）は $null
    param (
        $prompt
    )

    $dialog = loadWindow "${xamlDir}\dialog_leftover.xaml" ${fontsDir}
    $dialog.Owner = $window
    $dialog.Title = $prompt.Title
    $ctrl = @{}
    foreach ($name in @("HeadingText", "FactsList", "DetailToggleButton", "DetailBox", "DetailList", "HintText", "LeftoverCancelButton", "LeftoverStopButton")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $ctrl.HeadingText.Text = $prompt.Heading
    $ctrl.FactsList.ItemsSource = @($prompt.Facts | ForEach-Object { factKept $_.Title $_.Detail })
    $ctrl.DetailList.ItemsSource = @($prompt.Details | ForEach-Object { [pscustomobject]$_ })
    $ctrl.HintText.Text = $prompt.Hint
    $ctrl.LeftoverCancelButton.Content = $prompt.CancelText
    $ctrl.LeftoverStopButton.Content = @($prompt.Choices)[0].Text
    $ctrl.DetailToggleButton.Content = getLeftoverDetailToggleText $false
    $chosen = @{ Value = $null }
    # GetNewClosure() したハンドラは、呼び出しが入れ子だと関数を名前で解決できないため、実体を変数に取り込んで & で呼ぶ
    $getToggleText = ${function:getLeftoverDetailToggleText}
    $ctrl.LeftoverStopButton.Add_Click({
        $chosen.Value = "stop"
        $dialog.DialogResult = $true
    }.GetNewClosure())
    # マウス・キーボードでも UI オートメーションの Toggle でも IsChecked が変わるので、Click ではなく Checked / Unchecked で受ける
    $applyToggle = {
        $open = [bool]$ctrl.DetailToggleButton.IsChecked
        $ctrl.DetailBox.Visibility = $(if ($open) { "Visible" } else { "Collapsed" })
        $ctrl.DetailToggleButton.Content = & $getToggleText $open
    }.GetNewClosure()
    $ctrl.DetailToggleButton.Add_Checked($applyToggle)
    $ctrl.DetailToggleButton.Add_Unchecked($applyToggle)
    # うっかり Enter で終了しないよう、既定は［今回は終了しない］
    $dialog.Add_ContentRendered({ $ctrl.LeftoverCancelButton.Focus() | Out-Null }.GetNewClosure())
    $null = showOwnedDialog $dialog
    return $chosen.Value
}

function stopLeftoverProcesses {
    # ［終了する］のあと。PID を読み直し、確認に出したものだけを終了して、結果をステータスに出す
    param (
        $prompt
    )

    $script:leftoverPrompt = $prompt
    if ((getLeftoverSourceMode $script:leftoverFakeFile) -eq "Fake") {
        # 偽の行は、結果を作るだけ（プロセスには触れない）
        $targets = @(getLeftoverTargets @($prompt.Targets | ForEach-Object { $_.Id }) @($prompt.Targets))
        finishLeftoverStop (getLeftoverFakeStopResults $targets)
        return
    }
    startJob {
        param ($recordDir)
        return @{ Processes = @(getOfficeProcesses $recordDir) }
    } @((getOfficePidDir $workspace)) {
        param ($output, $errorText)
        if ($errorText -or $null -eq $output -or $output.Count -eq 0) {
            if ($errorText) { setStatus (getLeftoverFailureText "Stop" $errorText) }
            return
        }
        safe {
            $shown = @($script:leftoverPrompt.Targets | ForEach-Object { $_.Id })
            $targets = @(getLeftoverTargets $shown @($output[0].Processes))
            if ($targets.Count -eq 0) {
                return
            }
            startJob {
                param ($targets, $recordDir)
                return @{ Results = @(stopOfficeProcesses $targets $recordDir) }
            } @($targets, (getOfficePidDir $workspace)) {
                param ($output, $errorText)
                if ($errorText -or $null -eq $output -or $output.Count -eq 0) {
                    if ($errorText) { setStatus (getLeftoverFailureText "Stop" $errorText) }
                    return
                }
                safe { finishLeftoverStop @($output[0].Results) }
            } (getOfficePidQueue $workspace.Dir)
        }
    } (getOfficePidQueue $workspace.Dir)
}

function finishLeftoverStop {
    # 終了の結果をステータスに出す（言うことが無ければ変えない）
    param (
        [object[]]$results
    )

    $text = getLeftoverResultText $results
    if ($null -ne $text) {
        setStatus $text
    }
}
