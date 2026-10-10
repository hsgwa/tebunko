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

# 既定のワークスペースを求める・前後を比べる関数（tools\run_isolated.ps1・tests\run.ps1 と同じもの）
. "$PSScriptRoot\..\..\tools\isolation\isolation_common.ps1"

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
    # 既定のワークスペースは、いつも調べる（S6 も差し替えた既定で流すので、利用者の既定のワークスペースには触れない）
    $root = getGuiRepoRoot
    $list = {
        param ([string]$path)
        $map = getIsolationSnapshot $path
        return (@($map.Keys | Sort-Object | ForEach-Object { "${_}:$($map[$_])" }) -join "`n")
    }
    $snapshot = [ordered]@{
        "作業ツリーの setting.config" = (& $list "$root\setting.config")
        "作業ツリーの work\content_index" = (& $list "$root\work\content_index")
        "LOCALAPPDATA\tebunko"       = (& $list (Join-Path ([Environment]::GetFolderPath("LocalApplicationData")) "tebunko"))
        "Office のプロセスの数"       = @(Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue).Count
    }
    $snapshot["既定のワークスペース"] = (& $list (getRealDefaultWorkspace))
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

    # 環境変数 TEBUNKO_GUI_SINGLE=1 のときは、すべての画面のテストを単一 .ps1 版で流す（通しの確かめ用）
    if ($env:TEBUNKO_GUI_SINGLE -eq "1") {
        return newGuiSingleScriptTool $Dir $Settings
    }
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
    return @{ Dir = $tool; Work = $work; Config = "$tool\setting.config"; Gui = "$tool\scripts\tebunko\gui.ps1"; DefaultWorkspace = (newGuiDefaultWorkspace $Dir) }
}

function newGuiSingleScriptTool {
    # 展開せずに動く単一 .ps1 版（試験版）を $Dir\tool に組み立て、同じフォルダに設定ファイル（ワークスペースは
    # 同じフォルダの work）を書く。scripts\ の写しは使わず、tools\new_single_script.ps1 でその場で作る
    # （${rootDir} が .ps1 自身の置き場所になるため、setting.config・work もそこにできる。shared/core/paths.ps1）。
    #   settings: 設定ファイルに足す項目
    param (
        [string]$Dir,
        [hashtable]$Settings = @{}
    )

    $tool = Join-Path $Dir "tool"
    [void][IO.Directory]::CreateDirectory($tool)
    $scriptPath = Join-Path $tool "tebunko-test.ps1"
    if (!(Test-Path -LiteralPath $scriptPath)) {
        & "$(getGuiRepoRoot)\tools\new_single_script.ps1" -Version "v0.0.0-test" -OutFile $scriptPath | Out-Null
    }
    $work = Join-Path $tool "work"
    [void][IO.Directory]::CreateDirectory($work)
    $config = [ordered]@{ workspaceFolder = $work }
    foreach ($key in $Settings.Keys) { $config[$key] = $Settings[$key] }
    writeGuiConfig $tool $config
    return @{ Dir = $tool; Work = $work; Config = "$tool\setting.config"; Gui = $scriptPath; DefaultWorkspace = (newGuiDefaultWorkspace $Dir) }
}

function newGuiDefaultWorkspace {
    # 起動する画面の既定のワークスペース（環境変数 TEBUNKO_DEFAULT_WORKSPACE に渡す）。$Dir の中に作る（利用者の既定のワークスペースにしない）
    param ([string]$Dir)
    $path = Join-Path $Dir "default_workspace"
    [void][IO.Directory]::CreateDirectory($path)
    return $path
}

function writeGuiConfig {
    param ([string]$Tool, $Config)
    [IO.File]::WriteAllText("$Tool\setting.config", (ConvertTo-Json -InputObject $Config -Depth 5), (New-Object Text.UTF8Encoding($false)))
}

function readGuiConfig {
    param ($Tool)
    return (Get-Content -LiteralPath $Tool.Config -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function getGuiProcessCommand {
    # 画面を起動する powershell.exe の -Command の文字列。tebunko.bat と同じ、呼び出し演算子 & での起動にする
    # （-File で直接起動すると、実物の tebunko.bat（powershell -Command "...; & 'gui.ps1'"）より入れ子が 1 段浅くなり、
    # その 1 段の違いで .GetNewClosure() したスクリプトブロックが関数を名前で解決できなくなる不具合（#149 で見つかった）を
    # このテストがすり抜けてしまうため。入れ子の深さは変えない）。
    # 起動したスクリプトから戻ったら、その時刻・$?・$LASTEXITCODE・$Error の先頭 3 件を、ツールのフォルダの gui_returned_<PID>.txt に書く（多重起動の 2 つ目のプロセスが 1 つ目の印を上書きしないよう、プロセスごとのファイルにする）
    # （終了コードが 0 でないとき、スクリプトの外で決まったのかを見分ける材料。終了コードの決まり方は変えない: $? が偽なら exit 1、真なら何もしない）。
    # 閉じる順番の記録（本体の TEBUNKO_CLOSE_TRACE）は、画面のプロセスの中でだけ立てる（テストのプロセスには立てない）。
    # 引用符は単一引用符だけにする（Start-Process の引数として渡すため）
    param ($Tool)
    $gui = $Tool.Gui.Replace("'", "''")
    $returned = (Join-Path $Tool.Dir "gui_returned_").Replace("'", "''")
    $write = "try { `$e = (@(`$Error | Select-Object -First 3 | ForEach-Object { [string]`$_ }) -join ' / '); " +
        "`$t = (@((Get-Process -Id `$PID).Threads | ForEach-Object { [string]`$_.Id + ':' + [string]`$_.ThreadState + ':' + [string]`$_.WaitReason }) -join ','); " +
        "`$rs = (@(Get-Runspace | ForEach-Object { [string]`$_.Id + '/' + [string]`$_.Name + '/' + [string]`$_.RunspaceStateInfo.State + '/' + [string]`$_.RunspaceAvailability }) -join ','); " +
        "`$pr = Get-Process -Id `$PID; `$ev = (@(`$pr.Threads | Where-Object { [string]`$_.WaitReason -eq 'EventPairLow' } | ForEach-Object { [string][int](`$_.StartTime - `$pr.StartTime).TotalSeconds }) -join ','); " +
        "[IO.File]::WriteAllText('$returned' + `$PID + '.txt', ('returned=' + (Get-Date).ToString('o') + ' ok=' + `$r + ' LASTEXITCODE=' + `$c + ' Error=' + `$e + ' THREADS=' + `$t + ' RUNSPACES(' + @(Get-Runspace).Count + ')=' + `$rs + ' EVENTPAIR_START_SEC=' + `$ev)) } catch { }"
    # 【一時】終わり方の比べ（環境変数 TEBUNKO_TEST_EXIT_VARIANT が立っているときだけ。立っていなければ変わらない）
    #   envexit: 戻ったあと [Environment]::Exit(0) で終わる / exiting: PowerShell.Exiting の時刻を gui_exiting_<PID>.txt に書く
    $variants = @(([string]$env:TEBUNKO_TEST_EXIT_VARIANT) -split ',' | Where-Object { $_ })
    $head = ""
    $tail = ""
    # 戻ったあとの終わらせ方（Exiting の記録とは重ねて指定できる）
    if ($variants -contains 'envexit') { $tail = "; [Environment]::Exit(0)" }
    if ($variants -contains 'exit0') { $tail = "; exit 0" }
    if ($variants -contains 'dispatcher') { $tail = "; [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()" }
    if ($variants -contains 'rsdispose') { $tail = "; Get-Runspace | Where-Object { `$_.Id -ne [runspace]::DefaultRunspace.Id } | ForEach-Object { try { `$_.Dispose() } catch { } }" }
    if ($variants -contains 'gc') { $tail = "; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()" }
    if ($variants -contains 'exiting') {
        $exiting = (Join-Path $Tool.Dir "gui_exiting_").Replace("'", "''")
        # Exiting の中では待たない・重い処理をしない（結果を変えないため）。スレッド数・Runspace の数と状態（前景のスレッドの数は PowerShell からは取れないため、取らない）を書くだけ
        $action = '{ $o = ''thr='' + @((Get-Process -Id $PID).Threads).Count; try { $o += '' rs='' + (@(Get-Runspace | ForEach-Object { [string]$_.RunspaceStateInfo.State }) -join ''/''); $a = [System.Windows.Application]::Current; $o += '' App='' + ($null -ne $a) } catch { $o += '' 例外'' }; [IO.File]::WriteAllText(''EXITINGPATH'' + $PID + ''.txt'', (Get-Date).ToString(''o'') + '' '' + $o) }'.Replace('EXITINGPATH', $exiting)
        $head = "[void](Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action $action); "
    }
    return "`$env:TEBUNKO_CLOSE_TRACE = '1'; $head& '$gui'; `$r = `$?; `$c = `$LASTEXITCODE; $write; if (!`$r) { exit 1 }$tail"
}

function startGuiProcess {
    param ($Tool)
    $command = getGuiProcessCommand $Tool
    # 既定のワークスペースを、$TestDrive の中に差し替えて起動する（起動した子のプロセスに引き継がれる。外れた値は例外にして起動しない）
    assertNotRealWorkspace $Tool.DefaultWorkspace
    $previousWorkspace = $env:TEBUNKO_DEFAULT_WORKSPACE
    $env:TEBUNKO_DEFAULT_WORKSPACE = $Tool.DefaultWorkspace
    try {
        $p = Start-Process powershell.exe -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "RemoteSigned", "-Command", $command) -PassThru -WindowStyle Hidden
    } finally {
        $env:TEBUNKO_DEFAULT_WORKSPACE = $previousWorkspace
    }
    $null = $p.Handle   # ExitCode を取るため、起動の直後にハンドルを持つ
    return $p
}

function startGui {
    # 画面を起動する。プロセスを起こしたら、待たずにすぐ $S を返す（以降の操作の「場面」）。
    # 本体の窓（NavList を持つ窓）を待つのは invokeGuiScene の中で行う。起動そのものが失敗しても
    # （XAML の読み込み例外など）、そこで失敗の材料を残してからプロセスを止められるようにするため
    param (
        $Tool,
        [string]$Scene
    )

    $pool = [runspacefactory]::CreateRunspacePool(1, 4)
    $pool.Open()
    $S = @{ Tool = $Tool; Scene = $Scene; Step = "起動"; Async = New-Object System.Collections.ArrayList; Pool = $pool; Window = $null; Timing = [ordered]@{}; Extra = @(); StartedAt = (Get-Date) }
    $S.Process = startGuiProcess $Tool
    return $S
}

function waitGuiStarted {
    # 本体の窓（NavList を持つ窓）が出るまで待つ。invokeGuiScene が Body の前に呼ぶ
    param ($S)

    if ($S.Window) {
        return
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $S.Window = waitGui $S "本体の窓（NavList）" ${guiStartTimeout} {
        foreach ($w in @(getGuiTopWindows $S)) {
            if (findGui $w -Id "NavList") { return $w }
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
    markGuiClosing $S
    $pattern = $S.Window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
    $S.ClosePath = New-Object System.Collections.ArrayList
    $pre = ""
    try {
        $pre = "WindowInteractionState=" + $pattern.Current.WindowInteractionState + " IsModal=" + $pattern.Current.IsModal + " IsEnabled=" + $S.Window.Current.IsEnabled + " IsOffscreen=" + $S.Window.Current.IsOffscreen
    } catch { $pre = "直前の状態を取れなかった: " + $_.Exception.Message }
    $callAt = (Get-Date).ToString('HH:mm:ss.fff')
    $caught = $null
    try { $pattern.Close() } catch { $caught = $_ }
    $doneAt = (Get-Date).ToString('HH:mm:ss.fff')
    [void]$S.ClosePath.Add("test 手段=WindowPattern.Close 呼んだ=" + $callAt + " 戻った=" + $doneAt + " 結果=" + $(if ($caught) { "例外" } else { "戻った（値なし）" }) + " 例外=" + $(if ($caught) { $caught.Exception.GetType().Name + ": " + $caught.Exception.Message } else { "なし" }) + " 直前=" + $pre)
    if ($caught) { writeGuiClosePathMaterial $S "例外"; throw $caught }
    $S.Samples = New-Object System.Collections.ArrayList
    $S.NextSampleAt = $S.ClosingAt.AddSeconds(1)
    try {
        waitGui $S "画面が終了する" $Timeout -AllowExited {
            if ($S.Process.HasExited) { return $true }
            sampleGuiThreadsIfDue $S
            if (!$S.HangCaptured -and ((Get-Date) - $S.ClosingAt).TotalSeconds -ge 20) { $S.HangCaptured = $true; captureGuiHangMaterial $S }
            return $false
        } | Out-Null
    } finally {
        # 【一時】終わらないときの材料: 待つ間のスレッドの様子（成功のときも、待ちが 1 秒を超えたら残る）
        foreach ($line in @($S.Samples)) { Write-Host ("GUI-HANG 場面=" + $S.Scene + " " + $line) }
    }
    assertGuiExited $S
}

function captureGuiHangMaterial {
    # 【一時】閉じて 20 秒たっても終わらないとき、PID で止める前に、戻った印・Exiting の印と、ランナーの上の cdb（環境変数 TEBUNKO_CDB）で
    # スレッドごとの呼び出し元（スタックの文字だけ。ダンプのファイルは作らない・メモリの中身は出さない）を取る
    param ($S)

    $out = New-Object System.Collections.ArrayList
    $procId = $S.Process.Id
    foreach ($n in "gui_returned_", "gui_exiting_") {
        $f = Join-Path $S.Tool.Dir "$n$procId.txt"
        [void]$out.Add("---- $n$procId.txt ----")
        [void]$out.Add($(if (Test-Path -LiteralPath $f) { [IO.File]::ReadAllText($f) } else { "無い" }))
    }
    [void]$out.Add("---- 20 秒の時点の窓 ----")
    foreach ($l in @(getGuiWindowStates $S)) { [void]$out.Add($l) }
    [void]$out.Add("---- 閉じる道の跡 ----")
    foreach ($l in @(getGuiClosePathLines $S)) { [void]$out.Add($l) }
    $cdb = $env:TEBUNKO_CDB
    if (!$cdb -or !(Test-Path -LiteralPath $cdb)) {
        [void]$out.Add("---- cdb ---- 使えない（TEBUNKO_CDB が無い）")
    } else {
        try {
            $sym = Join-Path $(if ($env:RUNNER_TEMP) { $env:RUNNER_TEMP } else { $S.Tool.Dir }) "sym"
            $outFile = Join-Path $S.Tool.Dir "cdb_out_$procId.txt"
            $errFile = Join-Path $S.Tool.Dir "cdb_err_$procId.txt"
            # ~*e !clrstack は、スレッド 0 で GetContextState が失敗した（0x8007001F）ところで止まったため、スレッドを 1 本ずつ選んで取る（0 は最後）
            $perThread = (@(1..27 | ForEach-Object { "~${_}s; !clrstack" }) -join "; ")
            $cmds = ".symfix+ $sym; ~*kn 40; .loadby sos clr; .cordll -ve -u -l; !threads; !eestack -short; $perThread; ~0s; !clrstack; q"
            $argText = "-pv -p $procId -y `"srv*$sym*https://msdl.microsoft.com/download/symbols`" -c `"$cmds`""
            $proc = Start-Process -FilePath $cdb -ArgumentList $argText -RedirectStandardOutput $outFile -RedirectStandardError $errFile -NoNewWindow -PassThru
            if (!$proc.WaitForExit(120000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue; [void]$out.Add("---- cdb ---- 120 秒で止めた") }
            [void]$out.Add("---- cdb のスタック（フレームの行だけ） ----")
            $keep = @([IO.File]::ReadAllLines($outFile) | Where-Object { $_ -match '!|Id:|OS Thread Id|Child|Unable|Failed|failed|^\s*[0-9a-f]{2}\s|^[0-9a-f]{16}\s+[0-9a-f]{16}\s|^\s*\d+\s+\d+\s+[0-9a-f]+\s' } | Select-Object -First 1500)
            foreach ($l in $keep) { [void]$out.Add($l) }
        } catch {
            [void]$out.Add("---- cdb ---- 失敗: " + $_.Exception.Message)
        }
    }
    $S.HangMaterial = $out
    foreach ($l in $out) { Write-Host ("GUI-HANGMAT 場面=" + $S.Scene + " " + $l) }
}

function getGuiClosePathLines {
    # 【一時】閉じる道の跡（テスト側の記録と、本体が PID 入りのファイルに書いた記録）
    param ($S)

    $lines = New-Object System.Collections.ArrayList
    foreach ($l in @($S.ClosePath)) { [void]$lines.Add([string]$l) }
    $f = Join-Path $S.Tool.Dir "gui_closepath_$($S.Process.Id).txt"
    if (Test-Path -LiteralPath $f) {
        foreach ($l in @([IO.File]::ReadAllLines($f))) { [void]$lines.Add("本体 " + $l) }
    } else {
        [void]$lines.Add("本体 gui_closepath_$($S.Process.Id).txt が無い（本体は閉じる道のどこにも入っていない）")
    }
    return @($lines.ToArray())
}

function writeGuiClosePathMaterial {
    # 【一時】閉じる道の跡を CI のログに出す（失敗した回と、成功の見本）
    param ($S, [string]$Label)

    foreach ($l in @(getGuiClosePathLines $S)) { Write-Host ("GUI-CLOSEPATH " + $Label + " 場面=" + $S.Scene + " PID=" + $S.Process.Id + " " + $l) }
}

function getGuiWindowStates {
    # 【一時】画面のプロセスの上位の窓すべて（題・クラス名・表示か・有効か・モーダルか・持ち主の窓があるか）。
    # 持ち主のある窓は UI オートメーションの木では持ち主の窓の子として出るので、「持ち主あり」は子として見つかったものを指す
    param ($S)

    $rows = New-Object System.Collections.ArrayList
    $describe = {
        param ($w, $kind)
        $c = $w.Current
        $modal = "?"; $state = "?"
        try { $wp = $w.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern); $modal = $wp.Current.IsModal; $state = $wp.Current.WindowInteractionState } catch { }
        "窓 " + $kind + " 題=" + $c.Name + " クラス=" + $c.ClassName + " 表示=" + (!$c.IsOffscreen) + " 有効=" + $c.IsEnabled + " モーダル=" + $modal + " 状態=" + $state
    }
    try {
        foreach ($w in @(getGuiTopWindows $S)) { [void]$rows.Add((& $describe $w "上位")) }
        foreach ($w in @(getGuiOtherWindows $S)) { [void]$rows.Add((& $describe $w "本体以外（持ち主あり、または別の上位）")) }
        [void]$rows.Add("本体の有効=" + $S.Window.Current.IsEnabled)
    } catch {
        [void]$rows.Add("窓の状態を取れなかった: " + $_.Exception.Message)
    }
    return @($rows.ToArray())
}

function sampleGuiThreadsIfDue {
    # 【一時】閉じる操作の 1 秒後から 5 秒おきに、画面のプロセスのスレッドごとの状態・待ちの理由・開始アドレス（モジュール名+オフセット）・CPU 時間を控える
    param ($S)

    if ((Get-Date) -lt $S.NextSampleAt) { return }
    $S.NextSampleAt = (Get-Date).AddSeconds(5)
    try {
        $p = Get-Process -Id $S.Process.Id -ErrorAction Stop
        $mods = @(); try { $mods = @($p.Modules) } catch { }
        $parts = foreach ($t in $p.Threads) {
            $addr = ""
            try {
                $a = $t.StartAddress.ToInt64()
                $m = $mods | Where-Object { $a -ge $_.BaseAddress.ToInt64() -and $a -lt ($_.BaseAddress.ToInt64() + $_.ModuleMemorySize) } | Select-Object -First 1
                $addr = if ($m) { $m.ModuleName + "+0x" + ($a - $m.BaseAddress.ToInt64()).ToString("x") } else { "0x" + $a.ToString("x") }
            } catch { $addr = "?" }
            $cpu = ""; try { $cpu = [Math]::Round($t.TotalProcessorTime.TotalMilliseconds) } catch { }
            [string]$t.Id + ":" + [string]$t.ThreadState + ":" + [string]$t.WaitReason + ":" + $addr + ":cpu" + $cpu
        }
        $sec = [Math]::Round(((Get-Date) - $S.ClosingAt).TotalSeconds, 1)
        [void]$S.Samples.Add("閉じてから ${sec} 秒 スレッド $($p.Threads.Count) 動作中 " + (@($parts) -join ' | '))
    } catch {
        [void]$S.Samples.Add("様子を取れなかった: " + $_.Exception.Message)
    }
}

function markGuiClosing {
    # 閉じる操作の直前に、画面のプロセスの様子（スレッドの数・子のプロセス）と時刻を控える（assertGuiExited が終了の記録に使う）。
    # 閉じる操作を closeGui の外で行う場面（取り込みを止めて閉じる など）は、その操作の直前に呼ぶ
    param ($S)

    $S.ClosingAt = Get-Date
    $S.ClosingState = @()
    $S.Children = @()
    try {
        $S.Process.Refresh()
        $S.ClosingState += "スレッドの数: $($S.Process.Threads.Count)"
        $children = @(Get-CimInstance -ClassName Win32_Process -Filter "ParentProcessId = $($S.Process.Id)" -ErrorAction Stop)
        $S.Children = @($children | ForEach-Object { @{ Id = [int]$_.ProcessId; Name = [string]$_.Name } })
        $names = if ($S.Children.Count -eq 0) { "なし" } else { (@($S.Children | ForEach-Object { "$($_.Name)（PID $($_.Id)）" })) -join ", " }
        $S.ClosingState += "子のプロセス: $names"
    } catch {
        $S.ClosingState += "様子を取れなかった: $($_.Exception.Message)"
    }
}

function assertGuiExited {
    # 画面のプロセスが終わったあとの判定を 1 つにまとめる。終了コードが 0 でなければ、原因を追う材料
    # （クラッシュ情報・終了の記録）を集めてから失敗にする。閉じ方の違う場面（closeGui・取り込みを止めて閉じる）が、どれもここを通る
    param ($S)

    if ($S.Process.ExitCode -eq 0) {
        try {
            $diagFile = Join-Path $S.Tool.Dir "gui_returned_$($S.Process.Id).txt"
            $diagText = if (Test-Path -LiteralPath $diagFile) { [IO.File]::ReadAllText($diagFile) } else { "returned 無し" }
            $diagThreads = ([regex]::Match($diagText, 'THREADS=([^ ]*)').Groups[1].Value -split ',').Count
            $diagText = ($diagText -replace ' Error=.*?(?= THREADS=)', '') -replace 'THREADS=[^ ]*', "THREADS_COUNT=$diagThreads"
            $exitingFile = Join-Path $S.Tool.Dir "gui_exiting_$($S.Process.Id).txt"
            $exitingText = if (Test-Path -LiteralPath $exitingFile) { [IO.File]::ReadAllText($exitingFile) } else { "無し" }
            $traceTail = ""
            try {
                foreach ($f in @(getGuiCloseTraceFiles $S)) {
                    $mine = @([IO.File]::ReadAllLines($f.Path) | Where-Object { $_ -like "*PID $($S.Process.Id)`t*" })
                    if ($mine.Count -gt 0) { $traceTail += " TRACE=" + ((@($mine | Select-Object -Last 4) | ForEach-Object { ($_ -split "`t")[0] + " " + ($_ -split "`t")[2] }) -join ' ; ') }
                }
            } catch { }
            Write-Host ("GUI-DIAG exit=0 場面=" + $S.Scene + " EXITING=" + $exitingText + $traceTail + " " + $diagText)
            if (!$script:guiClosePathSampled) { $script:guiClosePathSampled = $true; writeGuiClosePathMaterial $S "成功の見本" }
        } catch { Write-Host ("GUI-DIAG 失敗 " + $_.Exception.Message) }
        return
    }
    collectGuiExitMaterial $S
    throw "画面の終了コードが 0 ではない（$($S.Process.ExitCode)）"
}

function collectGuiExitMaterial {
    # 終了コードが 0 でない（または予定外に終わった）ときの材料を $S に集める。reportStartupFailure（gui.ps1）を通らない終了
    # （native の障害など）は原因が分からないため、Windows のイベントログと、終了の様子を残す
    param ($S)

    $since = if ($S.StartedAt) { [datetime]$S.StartedAt } else { (Get-Date).AddMinutes(-5) }
    $S.CrashInfo = getGuiCrashInfo $S.Process.Id $since
    $S.ExitRecord = getGuiExitRecord $S
}

function matchGuiEventProcessId {
    # イベントの本文に、そのプロセス ID が 10 進数でも 16 進数（0x 付き。大文字・小文字は問わない）でも入っているか。
    # 別の数字・英数字の一部（PID 123 に対する 1234、0x1A2B3）には当たらない
    param ([string]$Message, [int]$ProcessId)

    if ([string]::IsNullOrEmpty($Message)) { return $false }
    $decimal = [string]$ProcessId
    $hex = $ProcessId.ToString("x")
    return ($Message -match "(?<![0-9A-Za-z])$decimal(?![0-9A-Za-z])") -or ($Message -match "(?i)(?<![0-9A-Za-z])0x0*$hex(?![0-9A-Za-z])")
}

${guiCrashProviders} = @("Application Error", ".NET Runtime", "Windows Error Reporting", "Application Hang")

function selectGuiCrashEvents {
    # イベントのうち、記録の種類（guiCrashProviders）が合い、起動の時刻以降で、本文にそのプロセス ID があるものだけを返す
    param ($Events, [int]$ProcessId, [datetime]$Since)

    return @(@($Events) | Where-Object {
        $_ -and $_.ProviderName -in ${guiCrashProviders} -and $_.TimeCreated -ge $Since -and (matchGuiEventProcessId ([string]$_.Message) $ProcessId)
    })
}

function readGuiApplicationEvents {
    # Application ログの、起動の時刻以降の記録。「該当が無い」（NoMatchingEventsFound）だけを 0 件として扱い、読めなかった（権限・ログの不具合）ときは例外にする
    # （読めなかったのに「該当なし」と書くと、探したが無かったのか、読めなかったのかが区別できない）
    param ([datetime]$Since)

    try {
        return @(Get-WinEvent -FilterHashtable @{ LogName = "Application"; StartTime = $Since } -ErrorAction Stop)
    } catch {
        if ($_.FullyQualifiedErrorId -like "NoMatchingEventsFound*") { return @() }
        throw
    }
}

function selectGuiCrashEventsWithoutPid {
    # 種類（guiCrashProviders）と時刻（起動以降）は合うが、本文にそのプロセス ID が無いイベントを返す。
    # .NET Runtime（1026）・Windows Error Reporting（1001）・Application Hang（1002）は、本文にプロセス ID を書かないことがあるため、
    # PID だけで当てると全部「該当なし」になる。PID で当たったものとは分けて扱う
    param ($Events, [int]$ProcessId, [datetime]$Since)

    return @(@($Events) | Where-Object {
        $_ -and $_.ProviderName -in ${guiCrashProviders} -and $_.TimeCreated -ge $Since -and -not (matchGuiEventProcessId ([string]$_.Message) $ProcessId)
    })
}

function getGuiCrashInfo {
    # 終了コードが 0 でないとき（reportStartupFailure が返す 1 の exit も含む）、Windows のイベントログ（Application）から、
    # そのプロセスが動いていた間（起動の時刻から今まで）の記録のうち、障害の種類のものを探す。2 段に分けて書く。
    #   1. 本文にそのプロセス ID があるもの（確実に当たる）
    #   2. PID を書かない種類のために、同じ時間の範囲・同じ種類で、PID が本文に無いもの
    # 2 は、GitHub Actions のランナー（GITHUB_ACTIONS が立つ。使い捨ての機械）では先頭 3 件の本文まで書く。
    # 手元の実行では、メンテナの実機のほかのアプリの記録の本文をファイルに残さないよう、件数だけにする。
    # 見つからなくても「該当なし」と探した範囲を返す（探したが無かったのか、探していないのかを区別できるように）
    param ([int]$ProcessId, [datetime]$Since)

    $range = "Application ログ、$($Since.ToString('yyyy-MM-dd HH:mm:ss')) 以降、種類: $(${guiCrashProviders} -join '・')"
    try {
        $events = readGuiApplicationEvents $Since
        $result = New-Object System.Collections.ArrayList
        $byPid = @(selectGuiCrashEvents $events $ProcessId $Since)
        if ($byPid.Count -eq 0) {
            [void]$result.Add("PID で当たった記録: 該当なし（探した範囲: $range、本文の PID $ProcessId（10 進数・0x 付きの 16 進数））")
        } else {
            [void]$result.Add("PID で当たった記録: $($byPid.Count) 件（先頭 5 件）")
            foreach ($e in @($byPid | Select-Object -First 5)) { [void]$result.Add("[$($e.TimeCreated)] $($e.ProviderName)（ID $($e.Id)）: $($e.Message)") }
        }
        $byTime = @(selectGuiCrashEventsWithoutPid $events $ProcessId $Since)
        if ($byTime.Count -eq 0) {
            [void]$result.Add("PID を問わず時間で拾った記録: 該当なし（探した範囲: $range）")
        } elseif ($env:GITHUB_ACTIONS) {
            [void]$result.Add("PID を問わず時間で拾った記録（本文に PID が無い。別のプロセスの記録を含み得る）: $($byTime.Count) 件（先頭 3 件）")
            foreach ($e in @($byTime | Select-Object -First 3)) { [void]$result.Add("[$($e.TimeCreated)] $($e.ProviderName)（ID $($e.Id)）: $($e.Message)") }
        } else {
            [void]$result.Add("PID を問わず時間で拾った記録: $($byTime.Count) 件（手元の実行では、ほかのアプリの記録の本文を残さないため、件数だけ。本文は GitHub Actions の実行でだけ書く）")
        }
        return @($result)
    } catch {
        return @("イベントログを読めなかった（探す範囲: $range）: $($_.Exception.Message)")
    }
}
function getGuiCloseTraceFiles {
    # 閉じる順番の記録（close_trace.txt）がある、その場面のワークスペースを探す。探す先は、設定（setting.config）に書かれたワークスペース・
    # 作業フォルダ（Work）・既定のワークスペースの 3 つだけ（切り替えた先は、ツールのフォルダの外の $TestDrive の中にある）。
    # 利用者の実機の既定のワークスペースの中は探さない（assertNotRealWorkspace を通らないものは外す）。Path と Place（フォルダ名）を返す
    param ($S)

    $dirs = New-Object System.Collections.ArrayList
    try {
        $folder = ([string](readGuiConfig $S.Tool).workspaceFolder).Trim()
        if ($folder -ne "") { [void]$dirs.Add([IO.Path]::GetFullPath([IO.Path]::Combine($S.Tool.Dir, [Environment]::ExpandEnvironmentVariables($folder)))) }
    } catch { }
    foreach ($d in @($S.Tool.Work, $S.Tool.DefaultWorkspace)) { if ($d) { [void]$dirs.Add([IO.Path]::GetFullPath($d)) } }
    $result = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($dir in $dirs) {
        $key = $dir.TrimEnd('\').ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        try { assertNotRealWorkspace $dir } catch { continue }
        $file = Join-Path $dir "close_trace.txt"
        if (Test-Path -LiteralPath $file) { [void]$result.Add(@{ Path = $file; Place = (Split-Path -Leaf $dir) }) }
    }
    return @($result)
}

function getGuiExitRecord {
    # 終了の様子の記録（終了の記録.txt の中身）。閉じる前の様子・終わった時刻・閉じる操作から終わるまでの秒数・子がまだ動いているか・
    # 呼び出し元に戻った印（gui_returned_<PID>.txt）・閉じる順番の記録（設定のワークスペース・作業フォルダ・既定のワークスペースの close_trace.txt のうち、この PID の行）
    param ($S)

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("場面: $($S.Scene)　手順: $($S.Step)")
    $part = {
        param ([string]$Name, [scriptblock]$Body)
        try { & $Body } catch { [void]$lines.Add("$Name を取れなかった: $($_.Exception.Message)") }
    }
    & $part "プロセスの終わり" {
        [void]$lines.Add("PID: $($S.Process.Id)　終了コード: $($S.Process.ExitCode)")
        $exitTime = $S.Process.ExitTime
        [void]$lines.Add("終わった時刻: $($exitTime.ToString('HH:mm:ss.fff'))")
        if ($S.ClosingAt) {
            [void]$lines.Add("閉じる操作から終わるまで: $([Math]::Round(($exitTime - $S.ClosingAt).TotalSeconds, 2)) 秒")
        } else {
            [void]$lines.Add("閉じる操作から終わるまで: 閉じる操作の前に終わった（または控えていない）")
        }
    }
    & $part "閉じる直前の様子" {
        foreach ($line in @($S.ClosingState)) { [void]$lines.Add("閉じる直前の $line") }
        if (@($S.ClosingState).Count -eq 0) { [void]$lines.Add("閉じる直前の様子: 控えていない") }
    }
    & $part "子のプロセス" {
        foreach ($child in @($S.Children)) {
            $alive = [bool](Get-Process -Id $child.Id -ErrorAction SilentlyContinue)
            [void]$lines.Add("終わったあとの子 $($child.Name)（PID $($child.Id)）: $(if ($alive) { 'まだ動いている' } else { '終わっている' })")
        }
    }
    & $part "gui_returned_<PID>.txt" {
        $file = Join-Path $S.Tool.Dir "gui_returned_$($S.Process.Id).txt"
        [void]$lines.Add("---- gui_returned_$($S.Process.Id).txt（呼び出し元に戻った印。無ければ戻る前にプロセスが終わった） ----")
        if (Test-Path -LiteralPath $file) { [void]$lines.Add([IO.File]::ReadAllText($file)) } else { [void]$lines.Add("無い") }
    }
    & $part "gui_exiting_<PID>.txt" {
        $file = Join-Path $S.Tool.Dir "gui_exiting_$($S.Process.Id).txt"
        if (Test-Path -LiteralPath $file) { [void]$lines.Add("PowerShell.Exiting の時刻（gui_exiting_$($S.Process.Id).txt）: " + [IO.File]::ReadAllText($file)) }
    }
    & $part "close_trace.txt" {
        # 同じワークスペースを何度も起こすと前の起動の行が残るので、この PID の行だけを写す。ワークスペースを切り替える場面もあるため、場所（ワークスペースのフォルダ名）を付ける
        [void]$lines.Add("---- close_trace.txt（閉じる順番の記録。この PID の行だけ。最後の行が、止まる前に着いた節目） ----")
        $files = @(getGuiCloseTraceFiles $S)
        if ($files.Count -eq 0) { [void]$lines.Add("無い") }
        foreach ($f in $files) {
            $mine = @([IO.File]::ReadAllLines($f.Path) | Where-Object { $_ -like "*`tPID $($S.Process.Id)`t*" })
            [void]$lines.Add("場所: $($f.Place)")
            if ($mine.Count -eq 0) { [void]$lines.Add("この PID の行は無い") } else { [void]$lines.AddRange($mine) }
        }
    }
    return @($lines)
}

function formatGuiSummaryRow {
    # GitHub Actions の Summary（Markdown の表）に足す 1 行。文言の | はエスケープし、改行は空白にする
    param ([string]$Scene, [string]$Step, [string]$Message)

    $cell = { param ([string]$Text) return ($Text -replace "\r?\n", " ").Replace("|", "\|") }
    return "| $(& $cell $Scene) | $(& $cell $Step) | $(& $cell $Message) |"
}

function addGuiSummaryRow {
    # 場面が失敗したとき、環境変数 GITHUB_STEP_SUMMARY があるときだけ、場面・手順・失敗の文言を表の 1 行として足す
    # （受け入れのときに、ログを開かなくても「閉じるときの終了コード 5」だと分かるようにする）
    param ([string]$Scene, [string]$Step, [string]$Message)

    $path = $env:GITHUB_STEP_SUMMARY
    if (!$path) { return }
    try {
        $text = ""
        if (!(Test-Path -LiteralPath $path) -or (Get-Item -LiteralPath $path).Length -eq 0) {
            $text = "### 画面のスモークテストで失敗した場面`r`n`r`n| 場面 | 手順 | 失敗の文言 |`r`n|---|---|---|`r`n"
        }
        $text += (formatGuiSummaryRow $Scene $Step $Message) + "`r`n"
        [IO.File]::AppendAllText($path, $text, (New-Object Text.UTF8Encoding($false)))
    } catch { }
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
            collectGuiExitMaterial $S
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
    # そうしないと、起動時の XAML の読み込み例外などで reportStartupFailure が出すメッセージボックスに気づけず、90 秒待ってから
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

function checkGuiRow {
    # 行のチェックを付ける（初めは付いていない。付いていれば何もしない）。［アクション ▾］の項目は、チェックを付けた行に対して動く
    param ($S, $Row)
    $check = findGui $Row -Type CheckBox
    if ((getGuiToggleState $check) -ne "On") {
        toggleGui $check
    }
    waitGui $S "行のチェックが付く" ${guiDefaultTimeout} { (getGuiToggleState (findGui $Row -Type CheckBox)) -eq "On" } | Out-Null
}

function clickGuiAction {
    # インデックス一覧の見出しの［アクション ▾］を押して開くメニューから、項目（AutomationId）を押す。
    # ［更新］［エクスポート…］［削除…］は、チェックを付けた行に対して動く（先に checkGuiRow で付ける）。
    # メニューは別の窓（ポップアップ）で開くので、本体の外から探す
    param ($S, [string]$Id, [string]$What = "")
    if (!$What) { $What = "［$Id］" }
    $button = waitGuiById $S $S.Window "ActionsButton"
    invokeGui $S $button "［アクション ▾］"
    $menu = waitGuiWindow $S "アクションのメニュー" -Id $Id
    $item = waitGuiById $S $menu $Id
    invokeGui $S $item $What
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
    # タブを選び、選ばれるまで待つ。Id は TabItem の AutomationId（IndexTab・SearchTab・SettingsTab）
    param ($S, [string]$Id, [string]$ContentId = "")

    $tab = waitGuiById $S $S.Window $Id
    selectGui $tab
    waitGui $S "タブ $Id が選ばれる" ${guiDefaultTimeout} { isGuiSelected $tab } | Out-Null
    if ($ContentId) { [void](waitGuiById $S $S.Window $ContentId) }
}

function getGuiSelectedTab {
    # いま選ばれているタブの AutomationId
    param ($S)
    foreach ($id in "IndexTab", "SearchTab", "SettingsTab") {
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
    # 失敗したときの画面の画像と、写した先のログ、終了の記録を、作業ツリーの work\test\gui\<場面>\ に置く。
    # 部分ごとに別の try にして、1 つが失敗しても残りを書く（失敗した部分は 証拠の保存の失敗.txt に理由を 1 行ずつ残す）
    param ($S)

    try {
        $dest = Join-Path (getGuiRepoRoot) "work\test\gui\$($S.Scene)"
        [void][IO.Directory]::CreateDirectory($dest)
    } catch { return }
    $failed = New-Object System.Collections.ArrayList
    $save = {
        param ([string]$Name, [scriptblock]$Body)
        try { & $Body } catch { [void]$failed.Add("${Name}: $($_.Exception.Message)") }
    }
    & $save "画面の画像" {
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
    }
    & $save "ログの写し" {
        foreach ($name in "gui_error_log.txt", "indexing_log.txt") {
            $file = Join-Path $S.Tool.Work $name
            if (Test-Path -LiteralPath $file) { Copy-Item -LiteralPath $file -Destination $dest -Force }
        }
        # 閉じる順番の記録は、ワークスペースを切り替える場面でも拾えるよう、その場面のワークスペースを探す（場所はワークスペースのフォルダ名。全部の PID の行を写す）
        foreach ($f in @(getGuiCloseTraceFiles $S)) {
            Copy-Item -LiteralPath $f.Path -Destination (Join-Path $dest "close_trace（$($f.Place)）.txt") -Force
        }
    }
    & $save "クラッシュ情報" {
        # 終了コードの確認まで進まなかった失敗でも、必ず書く（「探したが無かった」と「探していない」を区別できるように）
        $info = if ($S.CrashInfo) { $S.CrashInfo } else { @("探していない（終了コードの確認まで進まなかった失敗）") }
        # Windows Error Reporting などは遅れて書かれることがあるため、終了直後に読んだものに加えて、保存の時点でもう一度読む
        if ($S.CrashInfo -and $S.Process) {
            $since = if ($S.StartedAt) { [datetime]$S.StartedAt } else { (Get-Date).AddMinutes(-5) }
            $info = @("---- 終了の直後に読んだ ----") + @($info) + @("---- 保存の時点で読み直した ----") + @(getGuiCrashInfo $S.Process.Id $since)
        }
        $info | Set-Content -LiteralPath "$dest\クラッシュ情報.txt" -Encoding UTF8
    }
    & $save "終了の記録" {
        if ($S.ExitRecord) { $S.ExitRecord | Set-Content -LiteralPath "$dest\終了の記録.txt" -Encoding UTF8 }
    }
    & $save "ハングの材料" {
        if ($S.HangMaterial) { $S.HangMaterial | Set-Content -LiteralPath "$dest\ハングの材料.txt" -Encoding UTF8 }
    }
    & $save "窓の一覧" {
        $tree = New-Object System.Collections.ArrayList
        foreach ($w in @(getGuiTopWindows $S)) {
            [void]$tree.Add("[窓] $($w.Current.Name) | class=$($w.Current.ClassName) | type=$($w.Current.ControlType.ProgrammaticName) | offscreen=$($w.Current.IsOffscreen)")
            foreach ($t in @(getGuiTexts $w)) { [void]$tree.Add("    $t") }
        }
        if ($tree.Count -eq 0) { [void]$tree.Add("窓は無い（画面の窓も、ほかの窓も見つからなかった）") }
        $tree | Set-Content -LiteralPath "$dest\窓の一覧.txt" -Encoding UTF8
    }
    if ($failed.Count -gt 0) {
        try { $failed | Set-Content -LiteralPath "$dest\証拠の保存の失敗.txt" -Encoding UTF8 } catch { }
    }
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
        addGuiSummaryRow $S.Scene $S.Step $_.Exception.Message
        throw $message
    } finally {
        stopGui $S
    }
}

# ---- OS のフォルダ選択（Win32 のダイアログ。ボタンと欄は UI オートメーションのパターンを持たないため、窓のメッセージで操作する） ----

if (-not ('TebunkoGuiNative' -as [type])) {
    # RECT の構造体と P/Invoke の宣言を 1 つの Add-Type にまとめる（別々の Add-Type にすると、
    # -MemberDefinition の側から前に定義した構造体の型が見えない）
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public struct TebunkoGuiRect {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
public delegate bool TebunkoEnumWindowsProc(IntPtr hWnd, IntPtr lParam);
public static class TebunkoGuiNative {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out TebunkoGuiRect rect);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(TebunkoEnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, uint dwAttribute, out TebunkoGuiRect pvAttribute, int cbAttribute);
    [DllImport("user32.dll")]
    public static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);
}
'@
}

# DWMWA_EXTENDED_FRAME_BOUNDS。GetWindowRect は、見えない大きさ調整用の枠（DWM の影）を含めて少し大きく返すことがあり、
# そのまま撮ると窓の外側が数ピクセル写り込む。DwmGetWindowAttribute のこの属性は、実際に見えている枠を返す
${dwmwaExtendedFrameBounds} = 9

# ---- 写真を撮る道具（tools\capture_screens.ps1）が使う、窓を動かす関数 ----

function getGuiRect {
    # 部品の画面上の範囲（BoundingRectangle）を System.Drawing.Rectangle にして返す。
    # 塗りつぶす部品（本体・ダイアログの中の部品）に使う
    param ($Element)
    $r = $Element.Current.BoundingRectangle
    return New-Object Drawing.Rectangle([int][Math]::Round($r.X), [int][Math]::Round($r.Y), [int][Math]::Round($r.Width), [int][Math]::Round($r.Height))
}

function getGuiWindowRect {
    # 窓（本体・ダイアログ・メニュー・起動中の表示）の画面上の範囲を、ネイティブの呼び出しで返す。
    # UI オートメーションの BoundingRectangle は、窓がまだ IsOffscreen（描画が済んでいない）の間 0 x 0 になることがあり、
    # 起動中の表示（一瞬で消える）はそのまま消えてしまうことがあるため、窓そのものの範囲はこちらを使う。
    # DwmGetWindowAttribute（見えている枠）を優先し、失敗したら GetWindowRect（見えない調整用の枠を含むことがある）
    param ($Element)
    $handle = [IntPtr]$Element.Current.NativeWindowHandle
    if ($handle -eq [IntPtr]::Zero) {
        return New-Object Drawing.Rectangle(0, 0, 0, 0)
    }
    $rect = New-Object TebunkoGuiRect
    if ([TebunkoGuiNative]::DwmGetWindowAttribute($handle, ${dwmwaExtendedFrameBounds}, [ref]$rect, 16) -eq 0) {
        return New-Object Drawing.Rectangle($rect.Left, $rect.Top, ($rect.Right - $rect.Left), ($rect.Bottom - $rect.Top))
    }
    if (!([TebunkoGuiNative]::GetWindowRect($handle, [ref]$rect))) {
        return New-Object Drawing.Rectangle(0, 0, 0, 0)
    }
    return New-Object Drawing.Rectangle($rect.Left, $rect.Top, ($rect.Right - $rect.Left), ($rect.Bottom - $rect.Top))
}

function redrawGuiWindow {
    # 窓を、子の部品も含めて強制的に描き直す（RedrawWindow。RDW_INVALIDATE・RDW_ERASE・RDW_FRAME・
    # RDW_ALLCHILDREN・RDW_UPDATENOW = 0x585）。確認が続けて出るときなど、前の窓の絵の一部が残って
    # 見えることがあるため、撮る直前に呼ぶ
    param ($Window)
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    if ($handle -ne [IntPtr]::Zero) {
        [void][TebunkoGuiNative]::RedrawWindow($handle, [IntPtr]::Zero, [IntPtr]::Zero, 0x585)
    }
}

function getGuiNativeProcessWindows {
    # プロセスの、目に見える窓のハンドルと範囲の一覧（Win32 の EnumWindows・GetWindowRect・IsWindowVisible だけで調べる）。
    # 起動中の表示（SplashProgress）は、出てから一瞬で閉じる一方、UI オートメーションの provider がまだこの窓を
    # 認識していないことがあり、AutomationElement 経由（getGuiTopWindows・BoundingRectangle）では見つからない・
    # 大きさが 0 x 0 のままのことがある。ネイティブの Win32 の呼び出しだけなら、窓ができた直後から正しい範囲が取れる
    param ([int]$ProcessId)

    $found = New-Object System.Collections.Generic.List[PSObject]
    $callback = [TebunkoEnumWindowsProc] {
        param ($hWnd, $lParam)
        $pid2 = 0
        [void][TebunkoGuiNative]::GetWindowThreadProcessId($hWnd, [ref]$pid2)
        if ($pid2 -eq $ProcessId -and [TebunkoGuiNative]::IsWindowVisible($hWnd)) {
            $rect = New-Object TebunkoGuiRect
            if ([TebunkoGuiNative]::GetWindowRect($hWnd, [ref]$rect)) {
                $w = $rect.Right - $rect.Left
                $h = $rect.Bottom - $rect.Top
                if ($w -gt 0 -and $h -gt 0) {
                    $found.Add([PSCustomObject]@{ Handle = $hWnd; Rect = (New-Object Drawing.Rectangle($rect.Left, $rect.Top, $w, $h)) })
                }
            }
        }
        return $true
    }
    [void][TebunkoGuiNative]::EnumWindows($callback, [IntPtr]::Zero)
    return @($found.ToArray())
}

function activateGuiWindow {
    # 窓を前に出す（撮る前に呼ぶ。Window を渡さなければ本体の窓。SetForegroundWindow はネイティブの窓、
    # SetFocus は UI オートメーションの部品に効く。起動中の表示（$S.Window がまだ無い）にも使えるよう Window を渡せる）
    param ($S, $Window = $null)
    if (!$Window) { $Window = $S.Window }
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    [void][TebunkoGuiNative]::SetForegroundWindow($handle)
    try { $Window.SetFocus() } catch { }
    Start-Sleep -Milliseconds 200
}

function pressGuiEnterKey {
    # 窓に Enter キー（既定のボタンを押す）を、WM_KEYDOWN・WM_KEYUP（VK_RETURN = 13）のメッセージだけで送る
    # （マウス・キーボードの合成はしない。SendInput・keybd_event は使わない）。
    # OS 標準のメッセージボックスの一部は、ボタンへの BM_CLICK（clickGuiNativeButton）では閉じないことがあるため、
    # そのときはこちらを使う（起動時にごく早く出るメッセージボックスなど）
    param ($Window)
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    [void][TebunkoGuiNative]::PostMessage($handle, 0x0100, [IntPtr]13, [IntPtr]::Zero)
    [void][TebunkoGuiNative]::PostMessage($handle, 0x0101, [IntPtr]13, [IntPtr]::Zero)
}

function pressGuiKey {
    # 窓に、修飾キーなしのキー（VK_F5 = 0x74 など）を、WM_KEYDOWN・WM_KEYUP のメッセージだけで送る（SendInput・keybd_event は使わない）。
    # Ctrl・Shift を押した形は送れない（ハンドラは本物のキーボードの状態を読むため）
    param ($Window, [int]$VirtualKey)
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    [void][TebunkoGuiNative]::PostMessage($handle, 0x0100, [IntPtr]$VirtualKey, [IntPtr]::Zero)
    [void][TebunkoGuiNative]::PostMessage($handle, 0x0101, [IntPtr]$VirtualKey, [IntPtr]::Zero)
}

function pressGuiMessageOk {
    # メッセージの画面（自前の画面。OS 標準のメッセージボックスにも使える）の［OK］を押す。
    # 自前の画面のボタンは Invoke パターンで押し、パターンを持たない OS 標準のボタンは BM_CLICK で押す。Enter キーも合わせて送る
    param ($Window)
    $ok = findGui $Window -Name "OK"
    if ($ok) {
        $pattern = $null
        if ($ok.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) { $pattern.Invoke() } else { clickGuiNativeButton $ok }
    }
    pressGuiEnterKey $Window
}

function closeGuiNativeMessage {
    # OK だけのメッセージの画面を、確実に閉じるまで閉じ続ける（もとは OS 標準のメッセージボックス用）。
    # ボタンへの BM_CLICK（clickGuiNativeButton）だけでは閉じないことがあるため、Enter キー（pressGuiEnterKey）も
    # 合わせて送り、閉じるまで両方を送り直す（写真を撮る道具が、実機で BM_CLICK だけでは閉じなかった場面があったため）
    param ($S, $Window, [string]$What, [int]$Timeout = ${guiDefaultTimeout})
    $key = getGuiKey $Window
    waitGui $S "$What が閉じる" $Timeout {
        try {
            pressGuiMessageOk $Window
        } catch { }
        Start-Sleep -Milliseconds 300
        !(@(getGuiOtherWindows $S) | Where-Object { (getGuiKey $_) -eq $key })
    } | Out-Null
}

function resizeGuiWindow {
    # 本体の窓の大きさを変える（TransformPattern.Resize）
    param ($S, [int]$Width, [int]$Height)
    $pattern = $S.Window.GetCurrentPattern([Windows.Automation.TransformPattern]::Pattern)
    $pattern.Resize($Width, $Height)
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
    # メッセージの画面（Text に文言が出ている窓）を待って、その文言を返し、［OK］で閉じる。
    # 自前の画面の［OK］は Invoke パターンで押す（OS 標準のメッセージボックスのときは、パターンが無いのでネイティブのクリックで押す）
    # Kind は、見出しの左のアイコンの名前（お知らせ・警告・エラー・確認）。ボタンは［OK］だけで、メッセージの画面の枠（見出しの部品）であることも確かめる
    param ($S, [string]$Text, [string]$What, [string]$Kind = "警告")
    $window = waitGuiWindow $S $What -Text $Text
    $found = @(getGuiTexts $window) -join " "
    $mark = findGui $window -Id "HeadingIcon"
    if (!$mark -or $mark.Current.Name -ne $Kind) { throw "${What}: 見出しのアイコンが「$Kind」でない（$(if ($mark) { $mark.Current.Name } else { 'アイコンなし' })）" }
    $buttonNames = @(findAllGui $window -Type Button | ForEach-Object { $_.Current.Name })
    if (($buttonNames -join ",") -ne "OK") { throw "${What}: ボタンが［OK］だけでない（$($buttonNames -join ',')）" }
    $button = waitGuiByName $S $window "OK"
    $pattern = $null
    if ($button.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) { $pattern.Invoke() } else { clickGuiNativeButton $button }
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
    # ［インデックス管理］の［すべて更新］を押し、確認のダイアログで［更新を開始］を押して、取り込みを始める
    param ($S)

    clickGui $S $S.Window "IndexingButton" "［すべて更新］"
    $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
    clickGui $S $confirm "StartButton" "確認の［更新を開始］"
    waitGuiWindowClosed $S $confirm "取り込みの確認"
}

function getGuiIndexingBannerText {
    # 更新の帯（IndexingProgressText）の文字。帯が隠れているときは空文字列（隠れた部品は UI Automation に出ない）
    param ($S)
    $e = findGui $S.Window -Id "IndexingProgressText"
    if ($e) { return [string]$e.Current.Name }
    return ""
}

function testGuiIndexing {
    # 取り込みの最中か（［更新中…］のボタンが出ている）
    param ($S)
    $button = findGui $S.Window -Id "IndexingButton"
    return ($button -and $button.Current.Name -eq "更新中…")
}
