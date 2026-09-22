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

function safe {
    # イベント処理で例外が起きても画面を落とさず、内容を表示する
    param (
        [scriptblock]$block
    )

    try {
        & $block
    } catch {
        writeErrorLog "画面の操作中" $_
        setStatus "エラーが発生しました：$($_.Exception.Message)"
        showMessage "エラーが発生しました。`n$($_.Exception.Message)`n`n詳しい内容は $(Split-Path -Leaf ${guiErrorLogFile}) に残しています。" "OK" "Error" | Out-Null
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

function testSamePath {
    # フォルダ選択ダイアログの中で、2つのパスが同じ書き方かを見る（末尾の \ ・大文字と小文字の違いは無視する）
    param (
        [string]$a,
        [string]$b
    )

    return [string]::Equals(([string]$a).TrimEnd("\"), ([string]$b).TrimEnd("\"), [System.StringComparison]::OrdinalIgnoreCase)
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

$script:jobs = New-Object System.Collections.ArrayList

function startJob {
    # scriptBlock を別スレッドで実行し、終わったら画面のスレッドで onDone { param($output, $errorText) } を呼ぶ
    param (
        [scriptblock]$scriptBlock,
        [object[]]$arguments,
        [scriptblock]$onDone
    )

    $ps = [powershell]::Create()
    [void]$ps.AddScript($scriptBlock.ToString())
    foreach ($argument in $arguments) {
        [void]$ps.AddArgument($argument)
    }
    [void]$script:jobs.Add(@{ PS = $ps; Handle = $ps.BeginInvoke(); OnDone = $onDone })
    $script:jobTimer.Start()
}


$script:jobTimer = newTimer 200 {
    safe {
        foreach ($job in @($script:jobs.ToArray())) {
            if (!$job.Handle.IsCompleted) {
                continue
            }
            $script:jobs.Remove($job)
            $output = $null
            $errorText = $null
            try {
                $output = $job.PS.EndInvoke($job.Handle)
                if ($job.PS.Streams.Error.Count -gt 0) {
                    $errorText = $job.PS.Streams.Error[0].ToString()
                }
            } catch {
                $errorText = $_.Exception.Message
            } finally {
                $job.PS.Dispose()
            }
            & $job.OnDone $output $errorText
        }
        if ($script:jobs.Count -eq 0) {
            $script:jobTimer.Stop()
        }
    }
}
