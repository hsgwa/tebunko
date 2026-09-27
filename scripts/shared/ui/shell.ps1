# 画面の共通部品（状態表示・メッセージ・確認ダイアログ・タイマー・別スレッド）。

function setStatus {
    param (
        [string]$text
    )

    $ui.StatusText.Text = $text
    $ui.StatusText.ToolTip = $text
}

function showMessage {
    param (
        [string]$message,
        [string]$buttons = "OK",
        [string]$icon = "Information",
        [string]$default = "None",
        [System.Windows.Window]$owner = $window  # ダイアログを開いているときは、そのダイアログを親にする
    )

    # 親を指定した表示に失敗しても、知らせること自体は止めない。
    # （親のウィンドウが閉じかけている・別のスレッドから呼ばれた等で失敗することがある。
    #   ここで例外が出ると、元のエラーが「Show の呼び出しに失敗」という別のエラーに化けて分からなくなる）
    if ($null -ne $owner) {
        try {
            return [System.Windows.MessageBox]::Show($owner, $message, ${appTitle}, $buttons, $icon, $default)
        } catch {
            writeErrorLog "メッセージを親付きで表示できませんでした" $_
        }
    }
    return [System.Windows.MessageBox]::Show($message, ${appTitle}, $buttons, $icon, $default)
}

# 確認ダイアログに並べる「実行するとこうなります」の 1 行を作る
function factKept { param ([string]$title, [string]$detail = "") [ConfirmFact]@{ Mark = "✓"; MarkBrush = ${okBrush};   Title = $title; Detail = $detail } }

function factGone { param ([string]$title, [string]$detail = "") [ConfirmFact]@{ Mark = "✗"; MarkBrush = ${ngBrush};   Title = $title; Detail = $detail } }

function factNext { param ([string]$title, [string]$detail = "") [ConfirmFact]@{ Mark = "→"; MarkBrush = ${infoBrush}; Title = $title; Detail = $detail } }

function factWarn { param ([string]$title, [string]$detail = "") [ConfirmFact]@{ Mark = "!"; MarkBrush = ${warnBrush}; Title = $title; Detail = $detail } }

function newChoiceContent {
    # 選択肢ボタンの中身。1 行目に動作、2 行目にその結果を置く
    param (
        [string]$text,
        [string]$detail
    )

    $panel = New-Object System.Windows.Controls.StackPanel
    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = $text
    $title.FontWeight = [System.Windows.FontWeights]::SemiBold
    $title.TextWrapping = "Wrap"
    $panel.Children.Add($title) | Out-Null
    if ($detail -ne "") {
        $line = New-Object System.Windows.Controls.TextBlock
        $line.Text = $detail
        $line.FontSize = 12
        $line.Foreground = ${grayBrush}
        $line.TextWrapping = "Wrap"
        $line.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 3, 0, 0
        $panel.Children.Add($line) | Out-Null
    }
    return $panel
}

function showConfirm {
    # 確認ダイアログ。「見出し（何をするか）」「こうなります（何が消えて何が残るか）」「選択肢のボタン」で、
    # 文章を読まなくても押す前に結果が分かるようにする。選んだ Value を返す（キャンセル・閉じるは $null）。
    #   choices: @{ Text = "削除する"; Detail = "ボタンの下に出す補足"; Value = "delete"; Danger = $true; Careful = $true } の配列
    #            1 つなら［実行］＋［キャンセル］、2 つ以上なら選択肢ボタンを縦に並べる
    #            Danger は赤いボタン、Careful は色はそのままでキャンセルを既定にする
    #   facts:   factGone / factKept / factNext で作った行
    #   hint:    読まなくても操作できる補足（別のやり方の案内など）
    param (
        [string]$heading,
        [object[]]$choices,
        [object[]]$facts = @(),
        [string]$hint = "",
        [string]$cancelText = "キャンセル",
        [System.Windows.Window]$owner = $window
    )

    $dialog = loadWindow "${sharedXamlDir}\dialog_confirm.xaml"
    $dialog.Title = ${appTitle}
    $dialog.Owner = $owner
    $ctrl = @{}
    foreach ($name in @("HeadingText", "FactsPanel", "FactsList", "ChoicePanel", "HintText", "ButtonPanel")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $chosen = @{ Value = $null }  # ボタンの Click から書き換えるため、入れ物ごとクロージャに渡す

    $ctrl.HeadingText.Text = $heading
    if ($facts.Count -gt 0) {
        $ctrl.FactsList.ItemsSource = $facts
        $ctrl.FactsPanel.Visibility = "Visible"
    }
    if ($hint -ne "") {
        $ctrl.HintText.Text = $hint
        $ctrl.HintText.Visibility = "Visible"
    }

    $onChoice = {
        param ($sender, $e)
        $chosen.Value = $sender.Tag
        $dialog.DialogResult = $true
    }.GetNewClosure()
    $focusTarget = $null
    $cautious = $false
    foreach ($choice in $choices) {
        $button = New-Object System.Windows.Controls.Button
        $button.Tag = $choice.Value
        $button.Add_Click($onChoice)
        if ($choices.Count -eq 1) {
            $cautious = [bool]$choice.Danger -or [bool]$choice.Careful
            $button.Style = $dialog.FindResource($(if ($choice.Danger) { "Danger" } else { "Primary" }))
            $button.Content = $choice.Text
            $button.IsDefault = !$cautious
            $ctrl.ButtonPanel.Children.Add($button) | Out-Null
        } else {
            $button.Style = $dialog.FindResource("Choice")
            $button.Content = newChoiceContent $choice.Text $choice.Detail
            $ctrl.ChoicePanel.Children.Add($button) | Out-Null
        }
        if ($null -eq $focusTarget) {
            $focusTarget = $button
        }
    }

    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = $cancelText
    $cancel.IsCancel = $true
    $ctrl.ButtonPanel.Children.Add($cancel) | Out-Null
    if ($cautious -or $choices.Count -gt 1) {
        # 消す操作・選択肢が複数の操作は、うっかり Enter で進まないようキャンセルを既定にする
        $cancel.IsDefault = $true
        $focusTarget = $cancel
    }

    $dialog.Add_ContentRendered({ $focusTarget.Focus() | Out-Null }.GetNewClosure())
    $null = $dialog.ShowDialog()
    return $chosen.Value
}

# 予期しない例外を繰り返し知らせないための、同じ例外の見分け（型・メッセージ・発生場所）と抑え方の状態。
# registerUnhandledErrorHandler から reportUnexpectedError を呼ぶときだけ効かせる（safe・Closing の呼び出しは見ない）
$script:unhandledErrorTable = [ordered]@{}
${maxUnhandledErrorKeys} = 100  # メッセージにパスなどが入ると際限なく増えるため、古いものから消す
$script:unhandledDialogShowing = $false
# 時刻の取得はここから行う（テストで差し替えて、60 秒の判定を進められるようにする）
$script:reportUnexpectedErrorNow = { Get-Date }

function reportUnexpectedError {
    # 予期しない例外を記録・ステータス・ダイアログで知らせる（内容・順は今までの safe の catch と同じ）。
    # -unhandled のときだけ、同じ例外（型・メッセージ・発生場所が同じ）の繰り返しを抑える。
    #   ダイアログ：画面を開いている間に 1 回だけ出す。別の例外でも、ダイアログを出している間は重ねない
    #   記録：60 秒に 1 回まで。飛ばした間の回数を、次に記録するときに添える
    param (
        [string]$context,
        [System.Management.Automation.ErrorRecord]$record,
        [switch]$unhandled
    )

    $message = $record.Exception.Message
    $logContext = $context
    $skipLog = $false
    $skipDialog = $false
    $entry = $null

    if ($unhandled) {
        $key = "$($record.Exception.GetType().FullName)|$message|$($record.InvocationInfo.PositionMessage)"
        if (!$script:unhandledErrorTable.Contains($key)) {
            if ($script:unhandledErrorTable.Count -ge ${maxUnhandledErrorKeys}) {
                $script:unhandledErrorTable.Remove(@($script:unhandledErrorTable.Keys)[0])
            }
            $script:unhandledErrorTable[$key] = @{ DialogShown = $false; LastLogged = $null; Skipped = 0 }
        }
        $entry = $script:unhandledErrorTable[$key]

        $now = & $script:reportUnexpectedErrorNow
        if ($null -eq $entry.LastLogged -or ($now - $entry.LastLogged).TotalSeconds -ge 60) {
            if ($entry.Skipped -gt 0) {
                $logContext = "${context}（前の記録の後に同じ例外が$($entry.Skipped)回起きました）"
            }
            $entry.LastLogged = $now
            $entry.Skipped = 0
        } else {
            $entry.Skipped++
            $skipLog = $true
        }
        $skipDialog = $entry.DialogShown -or $script:unhandledDialogShowing
    }

    if (!$skipLog) {
        writeErrorLog $logContext $record
    }
    setStatus "エラーが発生しました：$message"
    if ($skipDialog) {
        return
    }
    $dialogText = "エラーが発生しました。`n$message`n`n詳しい内容は $(Split-Path -Leaf (getGuiErrorLogFile)) に残しています。"
    if ($unhandled) {
        $entry.DialogShown = $true
        $script:unhandledDialogShowing = $true
        try {
            showMessage $dialogText "OK" "Error" | Out-Null
        } finally {
            $script:unhandledDialogShowing = $false
        }
        return
    }
    showMessage $dialogText "OK" "Error" | Out-Null
}

function registerUnhandledErrorHandler {
    # 画面のスレッドの Dispatcher で捕まえていない例外を受け、画面を落とさず reportUnexpectedError で知らせる。
    # 登録した処理（デリゲート）を返す（テストで Remove_UnhandledException するため）
    param (
        [System.Windows.Threading.Dispatcher]$dispatcher
    )

    $handler = {
        param ($sender, $e)
        $e.Handled = $true
        try {
            $exception = $e.Exception
            # WPF がハンドラーを Delegate.DynamicInvoke で呼ぶ経路では、これに包まれて届く
            while ($exception -is [System.Reflection.TargetInvocationException] -and $exception.InnerException) {
                $exception = $exception.InnerException
            }
            if ($exception -is [System.Management.Automation.IContainsErrorRecord]) {
                # PowerShell のスクリプトブロックが投げた例外は、元の ErrorRecord をそのまま使う
                $record = $exception.ErrorRecord
            } else {
                # WPF が投げた .NET の例外には ErrorRecord が無いため、記録に残せるよう作る
                $record = New-Object System.Management.Automation.ErrorRecord(
                    $exception, $exception.GetType().FullName, ([System.Management.Automation.ErrorCategory]::NotSpecified), $null)
            }
            reportUnexpectedError "画面の操作中" $record -unhandled
        } catch {
            # 知らせる処理そのものが失敗しても、画面は落とさない
        }
    }
    $dispatcher.Add_UnhandledException($handler)
    return $handler
}

function safe {
    # イベント処理で例外が起きても画面を落とさず、内容を表示する
    param (
        [scriptblock]$block
    )

    try {
        & $block
    } catch {
        reportUnexpectedError "画面の操作中" $_
    }
}

function newTimer {
    param (
        [int]$milliseconds,
        [scriptblock]$onTick
    )

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds($milliseconds)
    $timer.Add_Tick($onTick)
    return $timer
}

function formatTime {
    # 当日なら HH:mm、それ以前は M/d HH:mm
    param (
        $time
    )

    if ($null -eq $time) {
        return ""
    }
    if ($time.Date -eq (Get-Date).Date) {
        return $time.ToString("H:mm")
    }
    return $time.ToString("M/d H:mm")
}

function readTextShared {
    # インデクサが書き込み中でも妨げないよう、共有を許して読む
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) {
        return ""
    }
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
}

# ---- 別スレッドの処理（インデックスの件数など） ----

# 仕事を受けるスレッド（BackgroundQueue）は、読み込み口（gui.ps1）が $script:backgroundQueue に用意する。
# 仕事のスクリプトでは、そのツールの関数（lib.ps1）をそのまま使える（スレッドを始めたときに 1 回だけ読み込む）

# ネットワークのパスだけを調べる専用の列。届かない共有で止まっても、プレビュー・状態の読み直し（$script:backgroundQueue）が
# 待たされないようにする。初めて使うときにだけ作る（ネットワークのパスが無い利用者には、スレッドも lib.ps1 の読み込みも増えない）。
# 列を作る式（BackgroundQueue の生成）は gui.ps1 が持ち、setNetworkQueueFactory で渡す
# （shared/ が tebunko/gui.ps1 の変数に直接頼らないようにするため）。
$script:networkQueue = $null
$script:networkQueueFactory = $null

function setNetworkQueueFactory {
    # ネットワークを調べる列を作る式を登録する（gui.ps1 が起動時に 1 回呼ぶ）
    param ([scriptblock]$factory)
    $script:networkQueueFactory = $factory
}

function ensureNetworkQueue {
    # ネットワークを調べる列を、初めて使うときだけ作って返す。両方のスレッドを空の仕事で温める
    # （1 回目の［開く］が冷えたスレッドに当たって lib.ps1 の読み込みを待たないようにする）
    if ($null -ne $script:networkQueue) {
        return $script:networkQueue
    }
    if ($null -eq $script:networkQueueFactory) {
        throw "ensureNetworkQueue: setNetworkQueueFactory が呼ばれていない"
    }
    $script:networkQueue = & $script:networkQueueFactory
    # 温める回数は、作った列自身のスレッドの数から取る（gui.ps1 の ${backgroundWorkers} に頼らない。
    # 無いと 0 回になって黙って温めないままになるため）
    for ($i = 0; $i -lt $script:networkQueue.Pool.Size; $i++) {
        $script:networkQueue.Post('$null', @(), $null)
    }
    $script:jobTimer.Start()
    return $script:networkQueue
}

function startJob {
    # scriptBlock を別スレッドで実行し、終わったら画面のスレッドで onDone { param($output, $errorText) } を呼ぶ。
    #   queue: "default"（既定。今までの列）・"network"（届かない共有を調べる専用の列。無ければここで作る）
    param (
        [scriptblock]$scriptBlock,
        [object[]]$arguments,
        [scriptblock]$onDone,
        [string]$queue = "default"
    )

    $target = if ($queue -eq "network") { ensureNetworkQueue } else { $script:backgroundQueue }
    $target.Post($scriptBlock.ToString(), $arguments, $onDone)
    $script:jobTimer.Start()
}

# 終わった仕事を受け取る間隔。プレビューの読み込みも通るため短くする（両方の列に仕事が無ければ止める）
$script:jobTimer = newTimer 50 {
    safe {
        $pending = $script:backgroundQueue.Poll()
        if ($null -ne $script:networkQueue) {
            $pending += $script:networkQueue.Poll()
        }
        if ($pending -eq 0) {
            $script:jobTimer.Stop()
        }
    }
}