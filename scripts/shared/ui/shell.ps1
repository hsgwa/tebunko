# 画面の共通部品（状態表示・メッセージ・確認ダイアログ・タイマー・別スレッド）。
# メッセージ（showMessage）・確認（showConfirm）・エラーの知らせ（showErrorDialog）は、同じ窓（dialog_confirm.xaml）で出す。

function setStatus {
    param (
        [string]$text
    )

    $ui.StatusText.Text = $text
    $ui.StatusText.ToolTip = $text
}

function showMessage {
    # お知らせ・警告・誤り・確認のメッセージ。確認（showConfirm）・エラー（showErrorDialog）と同じ自前の画面で出す
    # （枠・題・余白・ボタンの形・種類ごとのアイコンと色は同じ。種類とボタンの決め方は message_view.ps1）。
    #   buttons: OK・OKCancel・YesNo・YesNoCancel ／ icon: Information・Warning・Error・Question・None ／ default: 最初に Enter で押すボタンの戻り値の名前
    # 戻り値は [System.Windows.MessageBoxResult]（OK・Yes・No・Cancel）。OS 標準の MessageBox.Show と同じ
    param (
        [string]$message,
        [string]$buttons = "OK",
        [string]$icon = "Information",
        [string]$default = "None",
        [System.Windows.Window]$owner = $window  # ダイアログを開いているときは、そのダイアログを親にする
    )

    # 自前の画面を出せないとき（別のスレッドから呼ばれた・親の窓が閉じかけている・画面の定義が読めない等）も、知らせること自体は止めない。
    # ここで例外が出ると、元のエラーが別のエラーに化けて分からなくなるため、記録だけ残して OS 標準のメッセージボックスで出す
    try {
        return showMessageDialog $message $buttons $icon $default $owner
    } catch {
        writeErrorLog "メッセージを自前の画面で表示できませんでした" $_
    }
    if ($null -ne $owner) {
        try {
            return [System.Windows.MessageBox]::Show($owner, $message, ${appTitle}, $buttons, $icon, $default)
        } catch {
            writeErrorLog "メッセージを親付きで表示できませんでした" $_
        }
    }
    return [System.Windows.MessageBox]::Show($message, ${appTitle}, $buttons, $icon, $default)
}

function setDialogLook {
    # メッセージの画面の見出しの左のアイコン。種類（getMessageLook）のアイコンと色にする。アイコンが無い種類は隠す
    param (
        $dialog,
        [hashtable]$look
    )

    $mark = $dialog.FindName("HeadingIcon")
    $host_ = $dialog.FindName("HeadingIconHost")
    if ($look.Icon -eq "") {
        $host_.Visibility = "Collapsed"
        return
    }
    $mark.Data = $dialog.FindResource($look.Icon)
    [System.Windows.Automation.AutomationProperties]::SetName($host_, $look.Label)  # 読み上げの名前（お知らせ・警告・エラー・確認）
    $mark.Stroke = $dialog.FindResource($look.Brush)
    $host_.Visibility = "Visible"
}

function showMessageDialog {
    param (
        [string]$message,
        [string]$buttons,
        [string]$icon,
        [string]$default,
        [System.Windows.Window]$owner
    )

    $dialog = loadWindow "${sharedXamlDir}\dialog_confirm.xaml" ${fontsDir}
    $dialog.Title = ${appTitle}
    $dialog.Width = getConfirmWidth "normal"
    if ($null -ne $owner) {
        $dialog.Owner = $owner
    } else {
        $dialog.WindowStartupLocation = "CenterScreen"
    }
    setDialogLook $dialog (getMessageLook $icon)
    $parts = getMessageParts $message
    $dialog.FindName("HeadingText").Text = $parts.Heading
    if ($parts.Hint -ne "") {
        $hint = $dialog.FindName("HintText")
        $hint.Text = $parts.Hint
        $hint.Visibility = "Visible"
    }
    if ($parts.Detail -ne "") {
        $dialog.FindName("DetailText").Text = $parts.Detail
        $dialog.FindName("DetailBox").Visibility = "Visible"
    }
    $panel = $dialog.FindName("ButtonPanel")
    $chosen = @{ Value = getMessageCloseResult $buttons }  # ボタンの Click から書き換えるため、入れ物ごとクロージャに渡す
    $onClick = {
        param ($sender, $e)
        $chosen.Value = $sender.Tag
        $dialog.DialogResult = $true
    }.GetNewClosure()
    $focusTarget = $null
    foreach ($spec in (getMessageButtons $buttons $default)) {
        $button = New-Object System.Windows.Controls.Button
        $button.Content = $spec.Text
        $button.Tag = $spec.Result
        $button.IsDefault = $spec.IsDefault
        $button.IsCancel = $spec.IsCancel
        if ($spec.Primary) {
            $button.Style = $dialog.FindResource("Primary")
            $focusTarget = $button
        }
        $button.Add_Click($onClick)
        $panel.Children.Add($button) | Out-Null
    }
    $dialog.Add_ContentRendered({ if ($null -ne $focusTarget) { $focusTarget.Focus() | Out-Null } }.GetNewClosure())
    $null = showOwnedDialog $dialog
    return [System.Windows.MessageBoxResult]$chosen.Value
}

# 確認ダイアログに並べる「実行するとこうなります」の 1 行を作る
function factKept { param ([string]$title, [string]$detail = "") [ConfirmFact]@{ Mark = "✓"; MarkBrush = ${okBrush};   Title = $title; Detail = $detail } }

function factGone { param ([string]$title, [string]$detail = "") [ConfirmFact]@{ Mark = "✗"; MarkBrush = ${ngBrush};   Title = $title; Detail = $detail } }

function factNext { param ([string]$title, [string]$detail = "") [ConfirmFact]@{ Mark = "→"; MarkBrush = ${infoBrush}; Title = $title; Detail = $detail } }

function factWarn { param ([string]$title, [string]$detail = "") [ConfirmFact]@{ Mark = "!"; MarkBrush = ${warnBrush}; Title = $title; Detail = $detail } }

# 確認ダイアログの形。呼び出し側が form を指定しなければ、選択肢から決める
#   choice: 選択肢が 2 つ以上（［キャンセル］の右に選択肢のボタンを横に並べる）、danger: 1 つで Danger、normal: それ以外
function getConfirmForm {
    param (
        [object[]]$choices,
        [string]$form = ""
    )

    if ($form -ne "") {
        return $form
    }
    if (@($choices).Count -gt 1) {
        return "choice"
    }
    if (@($choices).Count -eq 1 -and [bool]@($choices)[0].Danger) {
        return "danger"
    }
    return "normal"
}

# 形ごとの窓の幅。選択肢のボタンを並べる choice は広く、ほかは 520
function getConfirmWidth {
    param (
        [string]$form
    )

    switch ($form) {
        "choice" { return 620 }
        default  { return 520 }
    }
}

# 暗幕にする部品（窓いっぱいに掛ける Border）。shared/ はツールの部品の名前を知らないので、起動側が渡す
$script:dialogScrim = $null
$script:dialogDepth = 0  # いま開いているダイアログ（showOwnedDialog）の数

function setDialogScrim {
    param (
        $scrim
    )

    $script:dialogScrim = $scrim
}

function showOwnedDialog {
    # 本体の窓を親にするダイアログを出す。出している間は、本体の上に暗幕を掛ける（閉じる・例外のときも外す）。
    # 戻り値は ShowDialog の戻り値（DialogResult）
    param (
        $dialog
    )

    # ダイアログの上にさらにダイアログ（メッセージなど）を重ねるときは、中のほうが閉じても暗幕は外さない（外側のダイアログがまだ開いている）
    $scrim = $script:dialogScrim
    $script:dialogDepth++
    if ($null -ne $scrim) {
        $scrim.Visibility = "Visible"
    }
    try {
        return $dialog.ShowDialog()
    } finally {
        $script:dialogDepth--
        if ($null -ne $scrim -and $script:dialogDepth -le 0) {
            $scrim.Visibility = "Collapsed"
        }
    }
}

function showConfirm {
    # 確認ダイアログ。「見出し（何をするか）」「こうなります（何が消えて何が残るか）」「選択肢のボタン」で、
    # 文章を読まなくても押す前に結果が分かるようにする。選んだ Value を返す（キャンセル・閉じるは $null）。
    #   choices: @{ Text = "削除する"; Value = "delete"; Danger = $true; Careful = $true } の配列
    #            ［キャンセル］の右に横に並べる。最後の 1 つが主なボタン（Danger は赤）。ほかは枠だけのボタン
    #            Danger・Careful は、うっかり Enter で進まないようキャンセルを既定にする
    #   facts:   factGone / factKept / factNext で作った行
    #   hint:    読まなくても操作できる補足（別のやり方の案内など。見出しの下に淡く出す）
    #   detail:  詳しい内容（エラーの文面など）。枠の中に、選べる文字で出す
    #   title:   窓の題（OS のタイトルバー）。既定はアプリ名
    #   form:    normal・danger・choice。省略すると選択肢から決める（getConfirmForm）。幅と、danger の赤い印・赤い実行ボタンが変わる
    param (
        [string]$heading,
        [object[]]$choices,
        [object[]]$facts = @(),
        [string]$hint = "",
        [string]$detail = "",
        [string]$cancelText = "キャンセル",
        [string]$title = "",
        [string]$form = "",
        [System.Windows.Window]$owner = $window
    )

    $form = getConfirmForm $choices $form
    $dialog = loadWindow "${sharedXamlDir}\dialog_confirm.xaml" ${fontsDir}
    $dialog.Title = if ($title -ne "") { $title } else { ${appTitle} }
    $dialog.Width = getConfirmWidth $form
    $dialog.Owner = $owner
    $ctrl = @{}
    foreach ($name in @("HeadingIcon", "HeadingText", "FactsPanel", "FactsList", "DetailBox", "DetailText", "HintText", "ButtonPanel")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $chosen = @{ Value = $null }  # ボタンの Click から書き換えるため、入れ物ごとクロージャに渡す

    $ctrl.HeadingText.Text = $heading
    # 見出しの左のアイコン。取り消せない操作（danger）は誤りと同じ赤い「!」、ほかの確認は「?」（種類は message_view.ps1）
    setDialogLook $dialog (getMessageLook $(if ($form -eq "danger") { "Error" } else { "Question" }))
    if ($facts.Count -gt 0) {
        $ctrl.FactsList.ItemsSource = $facts
        $ctrl.FactsPanel.Visibility = "Visible"
    }
    if ($detail -ne "") {
        $ctrl.DetailText.Text = $detail
        $ctrl.DetailBox.Visibility = "Visible"
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
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = $cancelText
    $cancel.IsCancel = $true
    $ctrl.ButtonPanel.Children.Add($cancel) | Out-Null  # ［キャンセル］は選択肢のボタンの左
    for ($i = 0; $i -lt @($choices).Count; $i++) {
        $choice = @($choices)[$i]
        $button = New-Object System.Windows.Controls.Button
        $button.Tag = $choice.Value
        $button.Add_Click($onChoice)
        $button.Content = $choice.Text
        $isMain = ($i -eq @($choices).Count - 1)
        if ($isMain) {
            $button.Style = $dialog.FindResource($(if ($choice.Danger) { "Danger.Filled" } else { "Primary" }))
            $cautious = [bool]$choice.Danger -or [bool]$choice.Careful
            $button.IsDefault = !$cautious -and @($choices).Count -eq 1
        }
        $ctrl.ButtonPanel.Children.Add($button) | Out-Null
        if ($null -eq $focusTarget) {
            $focusTarget = $button
        }
    }
    if ($cautious -or @($choices).Count -gt 1) {
        # 消す操作・選択肢が複数の操作は、うっかり Enter で進まないようキャンセルを既定にする
        $cancel.IsDefault = $true
        $focusTarget = $cancel
    }

    $dialog.Add_ContentRendered({ $focusTarget.Focus() | Out-Null }.GetNewClosure())
    $null = showOwnedDialog $dialog
    return $chosen.Value
}

function showErrorDialog {
    # エラーの知らせ。見出し・エラーの文面・［内容をコピー］［ログを開く］［閉じる］。
    # logFile を渡さないときは［ログを開く］を出さない（開く先が無いため）
    param (
        [string]$heading,
        [string]$detail,
        [string]$logFile = "",
        [System.Windows.Window]$owner = $window
    )

    $dialog = loadWindow "${sharedXamlDir}\dialog_confirm.xaml" ${fontsDir}
    $dialog.Title = "エラー"
    $dialog.Width = getConfirmWidth "error"
    if ($null -ne $owner) {
        $dialog.Owner = $owner
    }
    setDialogLook $dialog (getMessageLook "Error")
    $dialog.FindName("HeadingText").Text = $heading
    $dialog.FindName("DetailText").Text = $detail
    $dialog.FindName("DetailBox").Visibility = "Visible"
    $panel = $dialog.FindName("ButtonPanel")

    $copy = New-Object System.Windows.Controls.Button
    $copy.Content = "内容をコピー"
    $copy.Add_Click({
        try {
            [System.Windows.Clipboard]::SetText($detail)
        } catch {
            # クリップボードを他が使っているときは、あきらめる（もう一度押せる）
        }
    }.GetNewClosure())
    $panel.Children.Add($copy) | Out-Null
    if ($logFile -ne "") {
        $open = New-Object System.Windows.Controls.Button
        $open.Content = "ログを開く"
        $open.Add_Click({
            if (Test-Path -LiteralPath $logFile) {
                Start-Process -FilePath "$env:SystemRoot\System32\notepad.exe" -ArgumentList "`"$logFile`""
            }
        }.GetNewClosure())
        $panel.Children.Add($open) | Out-Null
    }
    $close = New-Object System.Windows.Controls.Button
    $close.Content = "閉じる"
    $close.Style = $dialog.FindResource("Primary")
    $close.IsDefault = $true
    $close.IsCancel = $true
    $panel.Children.Add($close) | Out-Null
    $null = showOwnedDialog $dialog
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
    $heading = "予期しないエラーが起きました。"
    $logFile = getGuiErrorLogFile
    if ($unhandled) {
        $entry.DialogShown = $true
        $script:unhandledDialogShowing = $true
        try {
            showErrorDialog $heading "$($record.Exception.GetType().FullName): $message" $logFile
        } finally {
            $script:unhandledDialogShowing = $false
        }
        return
    }
    showErrorDialog $heading "$($record.Exception.GetType().FullName): $message" $logFile
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