# 画面のスモークテストの共通の関数（tests\gui\*.Tests.ps1 の BeforeAll が読み込む）。
#
# 本物の画面（scripts\tebunko\gui.ps1）を別のプロセスで起動し、UI オートメーションで操作する。
#   ・ツールは scripts\ を $TestDrive に写して起動し、設定ファイル・ワークスペースも写した先に置く（利用者の環境に触らない）
#   ・部品は AutomationId（= x:Name）で探す。名前の無い部品（確認ダイアログの選択肢・メッセージボックスのボタン）は表示の文字（Name）で探す
#   ・操作はパターン（Invoke・Value・SelectionItem・Toggle・Window）で行い、マウス・キーボードの合成は使わない
#   ・待ちは 100ms ごとに条件を調べる。上限を超えるか、予定していないエラーの窓が出たら、すぐ失敗にする
#   ・失敗したら、画面の画像とログを work\test\gui\<場面>\ に残す（CI は成果物に上げる）
# 詳細は docs\design\testing\gui-smoke.md「画面のスモークテスト」。

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms

# ---- 待ちの上限（秒） ----
${guiStartTimeout}  = 90    # 起動して本体の窓が出るまで
${guiIndexTimeout}  = 120   # インデックスの取り込みが終わるまで
${guiDefaultTimeout} = 30   # ほかの待ち

# 予定していない窓（本体が出す異常の窓）の文言。同じプロセスにこれらが出たら、待たずにその文言で失敗にする
${guiErrorPatterns} = @("予期しないエラー", "エラーが発生しました", "すでに開いています", "起動できません")

function getGuiRepoRoot {
    return (Resolve-Path "$PSScriptRoot\..\..").Path
}

# ---- 利用者の環境を変えていないことの確かめ ----

function getGuiEnvSnapshot {
    # 流す前後で比べる。作業ツリーの setting.config・work\content_index、%LOCALAPPDATA%\tebunko、利用者の既定のワークスペース、Office のプロセスの数。
    # 既定のワークスペースは、CI（GITHUB_ACTIONS）では S6 が使うので調べない
    $root = getGuiRepoRoot
    $list = {
        param ([string]$path)
        if (!(Test-Path -LiteralPath $path)) { return "(無い)" }
        $item = Get-Item -LiteralPath $path
        if (!$item.PSIsContainer) { return "$($item.Length):$($item.LastWriteTimeUtc.Ticks)" }
        return (@(Get-ChildItem -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { "$($_.FullName.Substring($path.Length)):$($_.Length):$($_.LastWriteTimeUtc.Ticks)" }) -join "`n")
    }
    $snapshot = [ordered]@{
        "作業ツリーの setting.config" = (& $list "$root\setting.config")
        "作業ツリーの work\content_index" = (& $list "$root\work\content_index")
        "LOCALAPPDATA\tebunko"       = (& $list (Join-Path ([Environment]::GetFolderPath("LocalApplicationData")) "tebunko"))
        "Office のプロセスの数"       = @(Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue).Count
    }
    if ($env:GITHUB_ACTIONS -ne "true") {
        $snapshot["既定のワークスペース"] = (& $list (Join-Path ([Environment]::GetFolderPath("UserProfile")) "Documents\tebunko_ws"))
    }
    return $snapshot
}

function compareGuiEnvSnapshot {
    # 前後の違いの一覧（無ければ空）。
    # Office のプロセスの数は、流す前に Office が動いていたら比べない（手元で利用者や、ほかのテストが Office を使っていることがあり、数が変わるため。
    # CI・Office の無い機械では、増えていないことを確かめる）
    param ($Before, $After)
    return @($Before.Keys | Where-Object {
        if ($_ -eq "Office のプロセスの数" -and [int]$Before[$_] -gt 0) { return $false }
        [string]$Before[$_] -ne [string]$After[$_]
    })
}

# ---- 起動する ----

function newGuiTool {
    # scripts\ を $Dir\tool\scripts に写し、設定ファイル（ワークスペースは写した先の work）を書く。
    #   settings: 設定ファイルに足す項目（既定のワークスペースを使う場面は workspaceFolder に "" を渡す）
    param (
        [string]$Dir,
        [hashtable]$Settings = @{}
    )

    $tool = Join-Path $Dir "tool"
    if (!(Test-Path -LiteralPath "$tool\scripts")) {
        [void][IO.Directory]::CreateDirectory($tool)
        Copy-Item -LiteralPath "$(getGuiRepoRoot)\scripts" -Destination "$tool\scripts" -Recurse
    }
    $work = Join-Path $tool "work"
    [void][IO.Directory]::CreateDirectory($work)
    $config = [ordered]@{ workspaceFolder = $work }
    foreach ($key in $Settings.Keys) { $config[$key] = $Settings[$key] }
    writeGuiConfig $tool $config
    return @{ Dir = $tool; Work = $work; Config = "$tool\setting.config"; Gui = "$tool\scripts\tebunko\gui.ps1" }
}

function writeGuiConfig {
    param ([string]$Tool, $Config)
    [IO.File]::WriteAllText("$Tool\setting.config", (ConvertTo-Json -InputObject $Config -Depth 5), (New-Object Text.UTF8Encoding($false)))
}

function readGuiConfig {
    param ($Tool)
    return (Get-Content -LiteralPath $Tool.Config -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function startGuiProcess {
    # tebunko.bat と同じ、呼び出し演算子 & での起動にする（-File で直接起動すると、実物の tebunko.bat
    # （powershell -Command "...; & 'gui.ps1'"）より入れ子が 1 段浅くなり、その 1 段の違いで
    # .GetNewClosure() したスクリプトブロックが関数を名前で解決できなくなる不具合（#149 で見つかった）を
    # このテストがすり抜けてしまうため）
    param ($Tool)
    $command = "& '$($Tool.Gui.Replace("'", "''"))'"
    $p = Start-Process powershell.exe -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "RemoteSigned", "-Command", $command) -PassThru -WindowStyle Hidden
    $null = $p.Handle   # ExitCode を取るため、起動の直後にハンドルを持つ
    return $p
}

function startGui {
    # 画面を起動する。プロセスを起こしたら、待たずにすぐ $S を返す（以降の操作の「場面」）。
    # 本体の窓（Tabs を持つ窓）を待つのは invokeGuiScene の中で行う。起動そのものが失敗しても
    # （XAML の読み込み例外など）、そこで失敗の材料を残してからプロセスを止められるようにするため
    param (
        $Tool,
        [string]$Scene
    )

    $pool = [runspacefactory]::CreateRunspacePool(1, 4)
    $pool.Open()
    $S = @{ Tool = $Tool; Scene = $Scene; Step = "起動"; Async = New-Object System.Collections.ArrayList; Pool = $pool; Window = $null; Timing = [ordered]@{}; Extra = @() }
    $S.Process = startGuiProcess $Tool
    return $S
}

function waitGuiStarted {
    # 本体の窓（Tabs を持つ窓）が出るまで待つ。invokeGuiScene が Body の前に呼ぶ
    param ($S)

    if ($S.Window) {
        return
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $S.Window = waitGui $S "本体の窓（Tabs）" ${guiStartTimeout} {
        foreach ($w in @(getGuiTopWindows $S)) {
            if (findGui $w -Id "Tabs") { return $w }
        }
    }
    $S.Timing["起動"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
}

function stopGui {
    # 画面が残っていれば止め、裏の呼び出しを片づける（終了コードは、閉じる操作の側で確かめる）
    param ($S)

    foreach ($extra in @($S.Extra)) {
        if ($extra -and !$extra.HasExited) { Stop-Process -Id $extra.Id -Force -ErrorAction SilentlyContinue }
    }
    if ($S.Process -and !$S.Process.HasExited) {
        Stop-Process -Id $S.Process.Id -Force -ErrorAction SilentlyContinue
        [void]$S.Process.WaitForExit(10000)
    }
    foreach ($item in @($S.Async)) {
        try {
            if ($item.Result.IsCompleted) { $item.Shell.Dispose() } else { [void]$item.Shell.BeginStop($null, $null) }
        } catch { }
    }
    try { $S.Pool.Close() } catch { }
}

function closeGui {
    # 本体を閉じて、終了コード 0 で終わるまで確かめる
    param ($S, [int]$Timeout = ${guiDefaultTimeout})

    setGuiStep $S "閉じる"
    $processId = $S.Process.Id
    $pattern = $S.Window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
    $pattern.Close()
    waitGui $S "画面が終了する" $Timeout -AllowExited { $S.Process.HasExited } | Out-Null
    if ($S.Process.ExitCode -ne 0) {
        # trap（gui.ps1）を通らない終了（native の障害など）は原因が分からないため、Windows のイベントログを材料に残す
        $S.CrashInfo = getGuiCrashInfo $processId
        throw "画面の終了コードが 0 ではない（$($S.Process.ExitCode)）"
    }
}

function getGuiCrashInfo {
    # 終了コードが 0 でないとき（1 の trap の exit も含む）、Windows のイベントログ（Application）からその
    # プロセス ID に関する直近の記録を探す。原因不明の終了（アクセス違反・COM の例外など）を追う材料にする
    param ([int]$ProcessId)

    try {
        $events = Get-WinEvent -FilterHashtable @{ LogName = "Application"; StartTime = (Get-Date).AddMinutes(-5) } -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -like "*$ProcessId*" } | Select-Object -First 5
        return @($events | ForEach-Object { "[$($_.TimeCreated)] $($_.ProviderName)（ID $($_.Id)）: $($_.Message)" })
    } catch {
        return @("イベントログを読めなかった: $($_.Exception.Message)")
    }
}

# ---- 待つ ----

function setGuiStep {
    param ($S, [string]$Step)
    $S.Step = $Step
}

function waitGui {
    # Condition が $null・$false 以外を返すまで、100ms ごとに調べる。返した値を返す。
    # 待つ間に、画面のプロセスが終わったとき（AllowExited でなければ）と、予定していないエラーの窓が出たときは、すぐ失敗にする
    param (
        $S,
        [string]$What,
        [int]$Timeout,
        [scriptblock]$Condition,
        [switch]$AllowExited
    )

    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        $result = & $Condition
        if ($result) { return $result }
        if (!$AllowExited -and $S.Process.HasExited) {
            throw "画面が終了した（終了コード $($S.Process.ExitCode)）。待っていたもの: $What"
        }
        if (!$AllowExited) {
            $bad = getGuiErrorWindowText $S
            if ($bad) { throw "予定していない窓が出た: 「$bad」。待っていたもの: $What" }
        }
        if ($sw.Elapsed.TotalSeconds -gt $Timeout) {
            throw "$What が $Timeout 秒以内に見つからなかった"
        }
        Start-Sleep -Milliseconds 100
    }
}

function getGuiErrorWindowText {
    # 本体以外の窓に、異常の文言（guiErrorPatterns）があれば、その文字を返す。
    # 本体の窓がまだ見つかっていない（起動を待っている間）は、そのプロセスのすべての窓（起動中の表示は除く）を調べる。
    # そうしないと、起動時の XAML の読み込み例外などで trap が出すメッセージボックスに気づけず、90 秒待ってから
    # 「本体の窓が見つからない」というだけの失敗になり、材料も残らない
    param ($S)

    $windows = if ($S.Window) { @(getGuiOtherWindows $S) } else { @(getGuiTopWindows $S | Where-Object { !(findGui $_ -Id "SplashProgress") }) }
    foreach ($w in $windows) {
        foreach ($text in @(getGuiTexts $w)) {
            foreach ($pattern in ${guiErrorPatterns}) {
                if ($text -like "*$pattern*") { return $text }
            }
        }
    }
    return $null
}

# ---- 部品を探す ----

function getGuiKey {
    param ($Element)
    try { return ($Element.GetRuntimeId() -join ",") } catch { return "" }
}

function newGuiCondition {
    param ([string]$Id, [string]$Name, [string]$Type)

    $ae = [Windows.Automation.AutomationElement]
    $list = New-Object System.Collections.ArrayList
    if ($Id) { [void]$list.Add((New-Object Windows.Automation.PropertyCondition($ae::AutomationIdProperty, $Id))) }
    if ($Name) { [void]$list.Add((New-Object Windows.Automation.PropertyCondition($ae::NameProperty, $Name))) }
    if ($Type) { [void]$list.Add((New-Object Windows.Automation.PropertyCondition($ae::ControlTypeProperty, [Windows.Automation.ControlType]::$Type))) }
    if ($list.Count -eq 0) { return [Windows.Automation.Condition]::TrueCondition }
    if ($list.Count -eq 1) { return $list[0] }
    return New-Object Windows.Automation.AndCondition([Windows.Automation.Condition[]]$list.ToArray())
}

function findGui {
    # Root の下（子孫）から、AutomationId・Name・種類（ControlType 名。Button など）に合う最初の部品。無ければ $null
    param ($Root, [string]$Id, [string]$Name, [string]$Type)
    try {
        return $Root.FindFirst("Descendants", (newGuiCondition -Id $Id -Name $Name -Type $Type))
    } catch {
        return $null
    }
}

function findAllGui {
    param ($Root, [string]$Id, [string]$Name, [string]$Type)
    try {
        return @($Root.FindAll("Descendants", (newGuiCondition -Id $Id -Name $Name -Type $Type)) | ForEach-Object { $_ })
    } catch {
        return @()
    }
}

function getGuiTopWindows {
    # 画面のプロセスの、いちばん上の窓（本体・ダイアログ・メッセージボックス・メニュー・OS のフォルダ選択）。
    # ツールヒント（結果の見出し行などの ToolTip="{Binding ...}"）は、マウスを動かさなくても、UI オートメーションで
    # 部品を選ぶ・フォーカスするだけで出ることがある（WPF の既定 ShowsToolTipOnKeyboardFocus）。
    # クラス名が Popup の窓（WPF の Popup の入れ物。ControlType は Window で、ToolTip では出ない）を対象から外す
    param ($S)
    $ae = [Windows.Automation.AutomationElement]
    $condition = New-Object Windows.Automation.AndCondition(
        (New-Object Windows.Automation.PropertyCondition($ae::ProcessIdProperty, $S.Process.Id)),
        (New-Object Windows.Automation.NotCondition((New-Object Windows.Automation.PropertyCondition($ae::ClassNameProperty, "Popup")))))
    try {
        return @($ae::RootElement.FindAll("Children", $condition) | ForEach-Object { $_ })
    } catch {
        return @()
    }
}

function getGuiOtherWindows {
    # 本体以外の窓（ダイアログ・確認・メッセージボックス・メニュー・OS のフォルダ選択）。
    # 親（Owner）のある窓は UI オートメーションの木では親の窓の子として出るため、窓の子の窓を順にたどる
    param ($S)
    $mainKey = getGuiKey $S.Window
    $found = New-Object System.Collections.ArrayList
    $queue = New-Object System.Collections.Queue
    foreach ($w in @(getGuiTopWindows $S)) { $queue.Enqueue($w) }
    while ($queue.Count -gt 0) {
        $w = $queue.Dequeue()
        if ((getGuiKey $w) -ne $mainKey) { [void]$found.Add($w) }
        try {
            foreach ($child in $w.FindAll("Children", (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Window)))) {
                $queue.Enqueue($child)
            }
        } catch { }
    }
    return @($found.ToArray())
}

function getGuiTexts {
    # 窓の中の文字の一覧。WPF の文言は ControlType.Text、OS 標準のメッセージボックスの文言・ボタンは
    # ControlType.Pane（パターンを持たない Win32 の Static・Button）で出るため、両方から拾う
    param ($Window)
    try {
        $found = @(findAllGui $Window -Type Text) + @(findAllGui $Window -Type Pane)
        return @($found | ForEach-Object { $_.Current.Name } | Where-Object { $_ })
    } catch {
        return @()
    }
}

function waitGuiById {
    # 部品が出るまで待つ（Root は本体の窓、またはダイアログ）
    param ($S, $Root, [string]$Id, [int]$Timeout = ${guiDefaultTimeout})
    return waitGui $S "部品 $Id" $Timeout { findGui $Root -Id $Id }
}

function waitGuiByName {
    param ($S, $Root, [string]$Name, [string]$Type = "", [int]$Timeout = ${guiDefaultTimeout})
    return waitGui $S "部品「$Name」" $Timeout { findGui $Root -Name $Name -Type $Type }
}

function waitGuiWindow {
    # 本体以外の窓が出るまで待つ。Id を渡すと、その AutomationId の部品を持つ窓（Text があれば、その文字を含むものだけ）。
    # Id が無ければ、Text を含む文字のある窓（メッセージボックス・確認ダイアログ）。
    # Guard を渡すと、待つたびに呼ぶ（時間に頼る場面で、先に終わってしまったことにすぐ気づいて分かりやすい文言で
    # 失敗させるため。例: 取り込みが終わってしまい、確認ダイアログが二度と出ない場合）
    param ($S, [string]$What, [string]$Id = "", [string]$Text = "", [int]$Timeout = ${guiDefaultTimeout}, [scriptblock]$Guard = $null)

    return waitGui $S $What $Timeout {
        if ($Guard) { & $Guard }
        foreach ($w in @(getGuiOtherWindows $S)) {
            if ($Id) {
                $e = findGui $w -Id $Id
                if ($e -and (!$Text -or $e.Current.Name -like "*$Text*")) { return $w }
            } else {
                foreach ($t in @(getGuiTexts $w)) {
                    if ($t -like "*$Text*") { return $w }
                }
            }
        }
    }
}

function waitGuiWindowClosed {
    param ($S, $Window, [string]$What, [int]$Timeout = ${guiDefaultTimeout})
    $key = getGuiKey $Window
    waitGui $S "$What が閉じる" $Timeout { @(getGuiOtherWindows $S | Where-Object { (getGuiKey $_) -eq $key }).Count -eq 0 } | Out-Null
}

# ---- 操作する（パターンで行う） ----

function waitGuiEnabled {
    param ($S, $Element, [string]$What, [int]$Timeout = ${guiDefaultTimeout})
    waitGui $S "$What が押せる（有効になる）" $Timeout { try { $Element.Current.IsEnabled } catch { $false } } | Out-Null
}

function invokeGuiPatternAsync {
    # UI オートメーションのパターンの操作（Invoke・Close など）を、別のスレッド（ランスペース）から呼ぶ。
    # ダイアログ・メッセージボックスなどのモーダルを開く操作は、戻らないことがあるので、2 秒待って戻らなければ
    # 「モーダルが開いた」とみなして先へ進む（呼び出し自体の失敗はここで例外にする。$ErrorPrefix は「〇〇を押せなかった」など）。
    # invokeGui（InvokePattern）・closeGuiWindowAsync（WindowPattern.Close）が使う共通の土台
    param ($S, $Element, [string]$ErrorPrefix, [scriptblock]$Action)

    $shell = [powershell]::Create()
    $shell.RunspacePool = $S.Pool
    [void]$shell.AddScript($Action).AddArgument($Element)
    $result = $shell.BeginInvoke()
    [void]$S.Async.Add(@{ Shell = $shell; Result = $result })
    if ($result.AsyncWaitHandle.WaitOne(2000)) {
        try {
            [void]$shell.EndInvoke($result)
        } catch {
            throw "${ErrorPrefix}: $($_.Exception.InnerException.Message)"
        }
    }
}

function invokeGui {
    # ［ボタン］などを押す（InvokePattern）
    param ($S, $Element, [string]$What, [switch]$NoWait)

    if (!$NoWait) { waitGuiEnabled $S $Element $What }
    invokeGuiPatternAsync $S $Element "$What を押せなかった" {
        param ($element)
        Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
        $element.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
    }
}

function clickGui {
    # AutomationId の部品を探して押す（有効になるまで待つ）
    param ($S, $Root, [string]$Id, [string]$What = "")
    if (!$What) { $What = "［$Id］" }
    $e = waitGuiById $S $Root $Id
    invokeGui $S $e $What
}

function clickGuiByName {
    # 名前の無い部品（確認ダイアログの選択肢・メッセージボックスのボタン）を、表示の文字で探して押す
    param ($S, $Root, [string]$Name, [string]$Type = "Button")
    $e = waitGuiByName $S $Root $Name $Type
    invokeGui $S $e "［$Name］"
}

function setGuiText {
    # 文字を入れる欄（TextBox。ValuePattern）
    param ($S, $Element, [string]$Text)
    $Element.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Text)
}

function getGuiValue {
    # 欄の文字（ValuePattern）。無ければ Name
    param ($Element)
    $pattern = $null
    if ($Element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) { return $pattern.Current.Value }
    return $Element.Current.Name
}

function selectGui {
    # タブ・一覧の行を選ぶ（SelectionItemPattern）
    param ($Element)
    $Element.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
}

function isGuiSelected {
    param ($Element)
    return $Element.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Current.IsSelected
}

function toggleGui {
    param ($Element)
    $Element.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern).Toggle()
}

function getGuiToggleState {
    param ($Element)
    return [string]$Element.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern).Current.ToggleState
}

function selectGuiTab {
    # タブを選び、選ばれるまで待つ。Id は TabItem の AutomationId（IndexTab・SearchTab・SettingsTab・KillTab）
    param ($S, [string]$Id, [string]$ContentId = "")

    $tab = waitGuiById $S $S.Window $Id
    selectGui $tab
    waitGui $S "タブ $Id が選ばれる" ${guiDefaultTimeout} { isGuiSelected $tab } | Out-Null
    if ($ContentId) { [void](waitGuiById $S $S.Window $ContentId) }
}

function getGuiSelectedTab {
    # いま選ばれているタブの AutomationId
    param ($S)
    foreach ($id in "IndexTab", "SearchTab", "SettingsTab", "KillTab") {
        $tab = findGui $S.Window -Id $id
        if ($tab -and (isGuiSelected $tab)) { return $id }
    }
    return ""
}

function getGuiText {
    # 部品（TextBlock など）に出ている文字
    param ($Element)
    return $Element.Current.Name
}

# ---- 失敗の材料 ----

function saveGuiEvidence {
    # 失敗したときの画面の画像と、写した先のログを、作業ツリーの work\test\gui\<場面>\ に置く
    param ($S)

    try {
        $dest = Join-Path (getGuiRepoRoot) "work\test\gui\$($S.Scene)"
        [void][IO.Directory]::CreateDirectory($dest)
        try {
            $bounds = [Windows.Forms.Screen]::PrimaryScreen.Bounds
            $bitmap = New-Object Drawing.Bitmap($bounds.Width, $bounds.Height)
            $graphics = [Drawing.Graphics]::FromImage($bitmap)
            $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.Size)
            $bitmap.Save("$dest\画面.png", [Drawing.Imaging.ImageFormat]::Png)
            $graphics.Dispose(); $bitmap.Dispose()
        } catch {
            "画像を撮れなかった: $($_.Exception.Message)" | Set-Content -LiteralPath "$dest\画像なし.txt" -Encoding UTF8
        }
        foreach ($name in "画面エラー.txt", "インデックス作成ログ.txt") {
            $file = Join-Path $S.Tool.Work $name
            if (Test-Path -LiteralPath $file) { Copy-Item -LiteralPath $file -Destination $dest -Force }
        }
        if ($S.CrashInfo) {
            $S.CrashInfo | Set-Content -LiteralPath "$dest\クラッシュ情報.txt" -Encoding UTF8
        }
        $tree = New-Object System.Collections.ArrayList
        foreach ($w in @(getGuiTopWindows $S)) {
            [void]$tree.Add("[窓] $($w.Current.Name) | class=$($w.Current.ClassName) | type=$($w.Current.ControlType.ProgrammaticName) | offscreen=$($w.Current.IsOffscreen)")
            foreach ($t in @(getGuiTexts $w)) { [void]$tree.Add("    $t") }
        }
        $tree | Set-Content -LiteralPath "$dest\窓の一覧.txt" -Encoding UTF8
    } catch { }
}

function invokeGuiScene {
    # 場面 1 つ分の操作（Body）を実行する。まず本体の窓を待ち（waitGuiStarted）、それから Body を実行する。
    # 失敗したら、どの手順で止まったかを文言に足し、失敗の材料を残して例外にする（起動そのものの失敗も含む）。
    # 終わったら（失敗しても）画面を止める
    param ($S, [scriptblock]$Body)

    try {
        waitGuiStarted $S
        & $Body
    } catch {
        $message = "[$($S.Scene)・手順: $($S.Step)] $($_.Exception.Message)"
        saveGuiEvidence $S
        throw $message
    } finally {
        stopGui $S
    }
}

# ---- OS のフォルダ選択（Win32 のダイアログ。ボタンと欄は UI オートメーションのパターンを持たないため、窓のメッセージで操作する） ----

if (-not ('TebunkoGuiNative' -as [type])) {
    Add-Type -Namespace "" -Name TebunkoGuiNative -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
public static extern System.IntPtr SendMessage(System.IntPtr hWnd, uint msg, System.IntPtr wParam, string lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool PostMessage(System.IntPtr hWnd, uint msg, System.IntPtr wParam, System.IntPtr lParam);
'@
}

function clickGuiNativeButton {
    # パターンを持たない Win32 のボタンを、BM_CLICK で押す
    param ($Element)
    $handle = [IntPtr]$Element.Current.NativeWindowHandle
    if ($handle -eq [IntPtr]::Zero) { throw "ボタンの窓のハンドルが取れない（$($Element.Current.Name)）" }
    [void][TebunkoGuiNative]::PostMessage($handle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
}

function setGuiNativeText {
    # パターンを持たない Win32 の欄の文字を、WM_SETTEXT で入れる
    param ($Element, [string]$Text)
    $handle = [IntPtr]$Element.Current.NativeWindowHandle
    if ($handle -eq [IntPtr]::Zero) { throw "欄の窓のハンドルが取れない" }
    [void][TebunkoGuiNative]::SendMessage($handle, 0x000C, [IntPtr]::Zero, $Text)
}

function findGuiFolderPicker {
    # OS のフォルダ選択の窓（クラス #32770 で、フォルダ名の欄 1152 を持つ。予備のダイアログは 1148）
    param ($S)
    foreach ($w in @(getGuiOtherWindows $S)) {
        if ($w.Current.ClassName -ne "#32770") { continue }
        if ((findGui $w -Id "1152") -or (findGui $w -Id "1148")) { return $w }
    }
    return $null
}

function useGuiFolderPicker {
    # OS のフォルダ選択を、開いて閉じるまで本物で動かす。Path を渡すとその欄に入れて［フォルダーの選択］（ボタン ID 1）、
    # 渡さないと［キャンセル］（ボタン ID 2）。ボタンと欄は言語に依らない ID で探す（ランナーは英語、手元は日本語）
    param ($S, [string]$Path = "")

    $picker = waitGui $S "OS のフォルダ選択" ${guiDefaultTimeout} { findGuiFolderPicker $S }
    if ($Path) {
        $edit = waitGui $S "フォルダ名の欄" ${guiDefaultTimeout} { $e = findGui $picker -Id "1152"; if (!$e) { $e = findGui $picker -Id "1148" }; $e }
        setGuiNativeText $edit $Path
        $button = findGui $picker -Id "1" -Type Pane
    } else {
        $button = findGui $picker -Id "2" -Type Pane
    }
    if (!$button) { throw "フォルダ選択のボタンが見つからない" }
    clickGuiNativeButton $button
    waitGuiWindowClosed $S $picker "OS のフォルダ選択"
}

function useGuiFileOpenPicker {
    # OS のファイルを開くダイアログ（OpenFileDialog。selectZipFile）を、開いて閉じるまで本物で動かす。
    # フォルダ選択と同じクラス（#32770）・ファイル名の欄（1148）・［開く］ボタン（1）を使う
    param ($S, [string]$Path)

    $picker = waitGui $S "OS のファイルを開くダイアログ" ${guiDefaultTimeout} { findGuiFolderPicker $S }
    $edit = waitGui $S "ファイル名の欄" ${guiDefaultTimeout} { findGui $picker -Id "1148" }
    setGuiNativeText $edit $Path
    $button = findGui $picker -Id "1" -Type Pane
    if (!$button) { throw "［開く］ボタンが見つからない" }
    clickGuiNativeButton $button
    waitGuiWindowClosed $S $picker "OS のファイルを開くダイアログ"
}

# ---- テストデータ ----

function newGuiSampleIndex {
    # 検索できるインデックス（集約ファイル）を、写した先のワークスペースに作る。Root は TSV の置き場所。
    # 元のファイル（見積.xlsx・議事録.docx）は実在しない（元のファイルが無い行の確かめに使う）。
    # 呼び出す側で tests\helpers\load.ps1 を読み込んでおく（newTsv・newPackIndex・toIndexFileName）
    param ($Tool, [string]$Root, [string]$Name = "営業")

    $tsvRoot = Join-Path $Root "tsv_$Name"
    newTsv "$tsvRoot\$Name\見積.xlsx\$(toIndexFileName "見積")" @("品名`t数量`t単価", "", "りんご`t10`t100", "ABC`tabc")
    newTsv "$tsvRoot\$Name\議事録.docx\$(toIndexFileName "ページ001")" @("見積の方針", "単価は据え置き")
    [void](newPackIndex $tsvRoot "$($Tool.Work)\content_index")
}

function findGuiAnywhere {
    # 本体・本体以外の窓（メニューを含む）のどこかにある部品
    param ($S, [string]$Id, [string]$Name, [string]$Type)

    $found = findGui $S.Window -Id $Id -Name $Name -Type $Type
    if ($found) { return $found }
    foreach ($w in @(getGuiOtherWindows $S)) {
        $found = findGui $w -Id $Id -Name $Name -Type $Type
        if ($found) { return $found }
    }
    return $null
}

function newGuiSourceFolder {
    # 元のフォルダ（取り込むファイル）。tests\testdata\office\ の .docx・.pptx（Office を使わずに直接読める）を、Copies 組ずつ写す。
    # .xlsx は Excel が要るので使わない。Broken を付けると、ZIP としては開けるが XML が壊れている .docx（Office を起動せずに失敗になる）も 1 つ置く。
    # ZIP でないファイル（旧形式・パスワード付きなど）は、Word・PowerPoint に回し直されて Office が動くので使わない
    param ([string]$Path, [int]$Copies = 1, [switch]$Broken)

    [void][IO.Directory]::CreateDirectory($Path)
    $office = Join-Path (getGuiRepoRoot) "tests\testdata\office"
    $files = @("$office\Word\基本.docx", "$office\Word\表.docx", "$office\PowerPoint\基本.pptx")
    for ($i = 1; $i -le $Copies; $i++) {
        foreach ($file in $files) {
            $suffix = if ($Copies -gt 1) { "_$i" } else { "" }
            $name = [IO.Path]::GetFileNameWithoutExtension($file) + $suffix + [IO.Path]::GetExtension($file)
            Copy-Item -LiteralPath $file -Destination (Join-Path $Path $name)
        }
    }
    if ($Broken) { newGuiBrokenDocx (Join-Path $Path "壊れた文書.docx") }
}

function newGuiBrokenDocx {
    # ZIP としては開けるが、本文の XML（word\document.xml）が壊れている .docx
    param ([string]$Path)

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $entries = [ordered]@{
            "[Content_Types].xml" = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
            "_rels/.rels"         = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
            "word/document.xml"   = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>途中で壊れた'
        }
        foreach ($name in $entries.Keys) {
            $entry = $zip.CreateEntry($name)
            $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
            $writer.Write($entries[$name])
            $writer.Dispose()
        }
    } finally {
        $zip.Dispose()
    }
}

function getGuiGridRows {
    # 表（DataGrid）の行（DataItem）。列見出しは含めない
    param ($Grid)
    return @(findAllGui $Grid -Type DataItem)
}

function getGuiHitRows {
    # 検索結果の表（ResultGrid）の、ヒットの行（ファイルごとの見出しの行を除く）。
    # 見出しの行は AutomationProperties.Name をファイル名にし、セルの Text を持たない（FileHeaderRow のテンプレート）ため、
    # 中に Text が 1 つ以上ある行をヒットの行とする（読み上げの名前が場所の表示に変わっても壊れないようにする）
    param ($Grid)
    return @(getGuiGridRows $Grid | Where-Object { @(findAllGui $_ -Type Text).Count -gt 0 })
}

function getGuiRowTexts {
    # 行の中の文字（セルの Text）
    param ($Row)
    return @(findAllGui $Row -Type Text | ForEach-Object { $_.Current.Name })
}

function waitGuiButtonLike {
    # 表示の文字（1 行目の動作）がパターンに合うボタンが出るまで待つ。選択肢が 2 つ以上の確認ダイアログのボタンは、
    # 動作と補足を子の TextBlock 2 つに分けて持ち、ボタン自身の Name は空になる（newChoiceContent）。
    # そのため、ボタンの Name だけでなく、子の文字（1 行目）でも探す
    param ($S, $Root, [string]$Pattern, [int]$Timeout = ${guiDefaultTimeout})
    return waitGui $S "ボタン「$Pattern」" $Timeout {
        findAllGui $Root -Type Button | Where-Object {
            $_.Current.Name -like $Pattern -or ((@(findAllGui $_ -Type Text) | Select-Object -First 1).Current.Name -like $Pattern)
        } | Select-Object -First 1
    }
}

function clickGuiByNameLike {
    param ($S, $Root, [string]$Pattern)
    $button = waitGuiButtonLike $S $Root $Pattern
    invokeGui $S $button "［$Pattern］"
}

function answerGuiConfirm {
    # 確認ダイアログ（見出しの文字で探す）が出るのを待ち、選択肢を押して、閉じるまで待つ。
    # 「押す → 確認を待つ → 選択肢を押す → 閉じるのを待つ」の繰り返しをまとめたもの（トリガーの操作は呼び出し側で行う）。
    #   Like を付けると、選択肢は clickGuiByNameLike（1 行目の前方一致）で押す。押した後にダイアログが閉じない場合は使わない
    #   Guard は waitGuiWindow に渡す（時間に頼る場面で、先に終わってしまったことに早く気づかせるため）
    param ($S, [string]$What, [string]$Heading, [string]$Choice, [switch]$Like, [int]$Timeout = ${guiDefaultTimeout}, [scriptblock]$Guard = $null)

    $confirm = waitGuiWindow $S $What -Id "HeadingText" -Text $Heading -Timeout $Timeout -Guard $Guard
    if ($Like) { clickGuiByNameLike $S $confirm $Choice } else { clickGuiByName $S $confirm $Choice }
    waitGuiWindowClosed $S $confirm $What
}

function closeGuiMessage {
    # メッセージボックス（Text に文言が出ている窓）を待って、その文言を返し、［OK］で閉じる。
    # OS 標準のメッセージボックスの［OK］はパターンを持たないため、OS のフォルダ選択と同じくネイティブのクリックで押す
    param ($S, [string]$Text, [string]$What)
    $window = waitGuiWindow $S $What -Text $Text
    $found = @(getGuiTexts $window) -join " "
    $button = waitGuiByName $S $window "OK"
    clickGuiNativeButton $button
    waitGuiWindowClosed $S $window $What
    return $found
}

function closeGuiWindowAsync {
    # 窓を閉じる操作（WindowPattern.Close）。閉じるときに確認のダイアログが出て戻らないことがあるので、別のスレッドから呼ぶ
    param ($S, $Window, [string]$What = "窓を閉じる")

    invokeGuiPatternAsync $S $Window "$What ができなかった" {
        param ($element)
        Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
        $element.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close()
    }
}

function startGuiIndexing {
    # ［1 インデックス管理］の［インデックス作成を開始］を押し、確認のダイアログで［インデックス作成を開始］を押して、取り込みを始める
    param ($S)

    clickGui $S $S.Window "IndexingButton" "［インデックス作成を開始］"
    $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
    clickGui $S $confirm "StartButton" "確認の［インデックス作成を開始］"
    waitGuiWindowClosed $S $confirm "取り込みの確認"
}

function testGuiIndexing {
    # 取り込みの最中か（［インデックス作成中…］のボタンが出ている）
    param ($S)
    $button = findGui $S.Window -Id "IndexingButton"
    return ($button -and $button.Current.Name -eq "インデックス作成中…")
}
