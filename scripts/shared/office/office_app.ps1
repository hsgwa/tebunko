# Officeアプリ（Excel・Word・PowerPoint）の起動・終了と、1ファイルの制限時間の監視。
# COM を使う処理だけをまとめる。使う側が dot-source する（shared.ps1 では読み込まない）。

# ----------------------------------------------------------------------------
# Officeアプリ（Excel・Word・PowerPoint）の起動・終了
# ----------------------------------------------------------------------------

# 起動中のアプリ: 名前 → @{ Com; Pid; Shared }
$script:apps = @{}

# ExitWait: Quit の後、終わるのを待つ時間（ミリ秒）。過ぎたら強制終了する（stopApp）。
#   Excel は抽出で取り出したCOMオブジェクトが解放されきらないため、インデクサが動いている間は Quit しても終わらず、
#   待ちの上限まで待ってから強制終了していた（実測: 毎回 5 秒待って強制終了）。待つだけ無駄なため短くする。
#   GC で解放を促して自分で終わらせる方法は、インデクサの終了が COM の解放待ちで約 60 秒止まることがあった（実測 25 回中 1〜2 回）ため採らない
$appInfo = @{
    Excel      = @{ ProgId = "Excel.Application";      Process = "EXCEL";    ExitWait = 1000 }
    Word       = @{ ProgId = "Word.Application";       Process = "WINWORD";  ExitWait = 5000 }
    # SingleInstance: 1つのセッションに1つのプロセスしか持てないアプリ（PowerPointだけ）。
    # 既に起動している（利用者が開いている）ときは、接続せずに使わない（下の getApp）
    PowerPoint = @{ ProgId = "PowerPoint.Application"; Process = "POWERPNT"; ExitWait = 5000; SingleInstance = $true }
}

# SingleInstance のアプリが、既に自分のセッションで起動している（利用者が使用中の）ときに投げる例外の文言。
# 呼ぶ側は "<アプリ名>${officeAppInUseMessage}" の形で使う。Reroute（extract_office.ps1 の officeRequiredMessage・
# OperationCanceledException）と取り違えないよう、型（InvalidOperationException）でも区別する
${officeAppInUseMessage} = " が起動しているため、更新に使用できません"

# 起動した Office のプロセスの優先度は下げない（Normal のまま）。利用者がダブルクリックしたファイルがインデックス作成の Excel・Word で開くことがあり、
# PowerPoint は 1 つのプロセスしか持てないため、利用者とプロセスを共有しないと確実には言えない。利用者の操作を遅くしないよう、
# 優先度を下げるのは、利用者と共有しないインデックス作成のスレッドだけにする（docs/design/structure/threads.md「スレッドの一覧」）
# 起動したアプリの PID を入れる入れ物（ConcurrentDictionary[int,string]。$null なら入れない）。
# 画面が閉じるときに、インデックス作成が起動した Office を PID で止めるために使う
$script:officePidSink = $null
# 起動した Office の PID の記録を置くフォルダ（<ワークスペース>\office_pids\<PC の鍵>。$null なら書かない）。
# shared はツールを知らないため、使う側（インデックス作成）が場所を決めて入れる（office_process.ps1 の addOfficeRecord・removeOfficeRecord）。
# 記録は、次に画面を起動したときに、残った Office を確認して止めるために使う（書けなくても取り込みは続ける。その Office は止める対象にならない）
$script:officeRecordDir = $null
# 自分が開くブックの置き場（インデックス作成の作業領域の一時フォルダ。$null なら、開いているブックをすべて利用者のものとみなす）。
# 使う側が入れる。置き場の外のブックが Excel に開かれたら、利用者が開いたブックとして扱う（getForeignWorkbookCount・handOverApp）
$script:officeOwnDir = $null
# Excel を利用者に渡したときに呼ぶ処理（スクリプトブロック。引数はアプリ名と、窓を出せたか。$null なら何もしない）。ログを書く使う側が入れる
$script:onOfficeHandOver = $null
# 渡そうとして窓を出せなかった Excel（COM の参照を放さずに持ち続ける）
$script:officeKeptApps = @()

function resolveLongName {
    # 8.3 の短い名前（TEST~1 など）を含むパスを長い名前にする。読めなければ $null
    param ([string]$path)

    try {
        return (Get-Item -LiteralPath $path -ErrorAction Stop).FullName
    } catch {
        return $null
    }
}

function resolveOwnDir {
    # 自分が開くブックの置き場（一時フォルダ）を、比べる形にする。短い名前を長くできなければ、そのまま使う。
    # 置き場が決まっていなければ $null（そのときは開いたブックをすべて利用者のものと数える）
    param ([string]$dir)

    if (-not $dir) {
        return $null
    }
    $long = resolveLongName $dir
    return $(if ($long) { $long } else { $dir })
}

function isUnderDir {
    # path が dir の下にあるか。両方を正規化し、dir は区切りで終わる形にして、大文字小文字を区別せずに前置で比べる
    # （tmp と tmp2 を取り違えない）。正規化できない（URL など）ときは、下に無いものとする
    param ([string]$path, [string]$dir)

    try {
        $full = [System.IO.Path]::GetFullPath($path)
        $base = [System.IO.Path]::GetFullPath($dir).TrimEnd('\') + '\'
        return $full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)
    } catch {
        return $false
    }
}

function getWorkbookSplit {
    # Excel で開いているブックを、自分が開いたもの（置き場の下）と利用者が開いたもの（外）に分ける。@{ Own; Foreign }
    # 開いたブックの FullName を控えて照合する方法は、失敗の途中で Open が戻らなかったブックを控えられないため採らない。
    # 短い名前（~ を含む）は長い名前にしてから比べる。読めなければ、置き場の下と確かめられないので利用者のものと数える
    # （データを失わない側。自分のものは置き場の下にあり、読める）
    param ($com)

    $own = New-Object System.Collections.ArrayList
    $foreign = New-Object System.Collections.ArrayList
    $workbooks = $com.Workbooks
    try {
        foreach ($book in @($workbooks)) {
            if ($null -eq $book) { continue }
            $full = [string]$book.FullName
            if ($full.Contains('~')) {
                $full = resolveLongName $full
            }
            if ($full -and $script:officeOwnDir -and (isUnderDir $full $script:officeOwnDir)) {
                [void]$own.Add($book)
            } else {
                [void]$foreign.Add($book)
            }
        }
    } finally {
        try { releaseComObject $workbooks } catch {}
    }
    return @{ Own = $own; Foreign = $foreign }
}

function getForeignWorkbookCount {
    # Excel で利用者が開いたブック（置き場の外）の数。読めない（COM の例外）ときは -1（分からない。handOverForeignApp が窓で判断する）
    param ($com)

    try {
        $split = getWorkbookSplit $com
        $count = @($split.Foreign).Count
        # 取り出したブックの参照を残すと Excel が終わらなくなるため、数えたら放す
        foreach ($book in @($split.Own) + @($split.Foreign)) { try { releaseComObject $book } catch {} }
        return $count
    } catch {
        return -1
    }
}

function testProcessHasWindow {
    # 起動で控えた PID のプロセスに、見える窓があるか。PID だけを見る（窓を中身で探さない）。控えていなければ（0）偽
    param ([int]$processId, [string]$processName)

    if ($processId -le 0) {
        return $false
    }
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    return [bool]($process -and $process.ProcessName -eq $processName -and $process.MainWindowHandle -ne 0)
}

function getOwnSessionProcessIds {
    # 自分のセッションで動いている、指定した名前のプロセスのIDの一覧。
    # ほかのセッション（同じPCの別の利用者・別の作業フォルダの tebunko 等）のプロセスは、自分のものと取り違えないため数えない
    param (
        [string]$processName,
        [int]$sessionId
    )

    return @(Get-Process -Name $processName -ErrorAction SilentlyContinue |
        Where-Object { $_.SessionId -eq $sessionId } | ForEach-Object { $_.Id })
}

function getApp {
    # アプリのCOMオブジェクトを返す。起動していなければ起動する。
    # SingleInstance のアプリ（PowerPoint）が既に自分のセッションで起動しているときは、新しいプロセスができず
    # 利用者のアプリに接続することになるため、接続せずに例外を投げる（呼ぶ側は後回しにする）
    param (
        [string]$name
    )

    if (-not $script:apps.ContainsKey($name)) {
        $info = $appInfo[$name]
        $sessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId

        if ($info.SingleInstance -and (getOwnSessionProcessIds $info.Process $sessionId).Count -gt 0) {
            throw (New-Object System.InvalidOperationException "${name}${officeAppInUseMessage}")
        }

        # 終了できなかった場合に強制終了するため、新しく起動したプロセスのIDを控えておく
        $before = getOwnSessionProcessIds $info.Process $sessionId
        $com = New-Object -ComObject $info.ProgId
        $after = getOwnSessionProcessIds $info.Process $sessionId
        $newIds = @($after | Where-Object { $before -notcontains $_ })

        if ($info.SingleInstance -and $newIds.Count -eq 0) {
            # 事前の確認から New-Object の間に、利用者が先にアプリを起動した（先に接続した）。
            # DisplayAlerts・AutomationSecurity などの設定を変える前に解放し、利用者のアプリに触らない
            try { releaseComObject $com } catch {}
            throw (New-Object System.InvalidOperationException "${name}${officeAppInUseMessage}")
        }
        if ($newIds.Count -eq 1 -and $script:officePidSink) {
            $script:officePidSink[[int]$newIds[0]] = $info.Process
        }
        if ($newIds.Count -eq 1 -and $script:officeRecordDir) {
            # 起動時刻が読めたときだけ記録する（読めなければ、残っても止める対象にならない）
            $started = Get-Process -Id ([int]$newIds[0]) -ErrorAction SilentlyContinue
            $startTicks = $(if ($started -and $started.ProcessName -eq $info.Process) { getOfficeStartTicks $started } else { $null })
            if ($null -ne $startTicks) {
                [void](addOfficeRecord $script:officeRecordDir ([int]$newIds[0]) $info.Process $startTicks)
            }
        }

        switch ($name) {
            "Excel" {
                $com.Visible = $false
                $com.DisplayAlerts = $false
                $com.EnableEvents = $false
                $com.ScreenUpdating = $false
                $com.AskToUpdateLinks = $false
            }
            "Word" {
                $com.Visible = $false
                $com.DisplayAlerts = 0  # wdAlertsNone
            }
            "PowerPoint" {
                # PowerPointはウィンドウを隠せないため、ファイルをウィンドウ無しで開く（Visible は変更しない）
                $com.DisplayAlerts = 1  # ppAlertsNone
            }
        }
        $com.AutomationSecurity = 3  # msoAutomationSecurityForceDisable（マクロ無効）

        # Excel・Word は起動中のアプリに接続することがある（新しいプロセスができない場合。PowerPoint は上で防いでいるため
        # ここには来ない）。その場合は利用者のものなので終了させない
        $script:apps[$name] = @{
            Com    = $com
            Pid    = $(if ($newIds.Count -eq 1) { $newIds[0] } else { 0 })
            Shared = ($newIds.Count -eq 0)
        }
        updateWatchedPids
    }
    return $script:apps[$name].Com
}

function handOverApp {
    # 起動した Excel を、利用者に渡す。Quit も強制終了もしない（利用者のブックを閉じない）。渡した Excel はもう使わない。
    # 見張り・終了時の一括終了・次の起動の残り物の確認のどれからも外すので、誰も止めない。
    # Shared の Excel（利用者の Excel に接続したもの）には当てない（止めず、設定も変えない）
    param ([string]$name)

    $app = $script:apps[$name]
    if ($null -eq $app -or $app.Shared) {
        return $false
    }
    # (a) 見張りの対象から外す (b) 一括終了の対象から外す (c) 残り物の確認の記録を消す。それぞれ、ほかが失敗しても行う
    try { $script:apps.Remove($name) } catch {}
    try { updateWatchedPids } catch {}
    try {
        if ($app.Pid -and $script:officePidSink) {
            $removed = $null
            [void]$script:officePidSink.TryRemove([int]$app.Pid, [ref]$removed)
        }
    } catch {}
    try {
        if ($app.Pid -and $script:officeRecordDir) {
            removeOfficeRecord $script:officeRecordDir ([int]$app.Pid)
        }
    } catch {}

    $com = $app.Com
    $books = $null
    # (0) 自分が開いたブック（失敗して開いたままの一時コピー）は、利用者に見せず、一時フォルダの掃除をロックで失敗させないよう閉じる。
    # 閉じるのは置き場の下と確かめられたものだけ（利用者のブックは閉じない）
    try {
        $books = getWorkbookSplit $com
        foreach ($book in $books.Own) {
            try { $book.Close($false) } catch {}
        }
    } catch {}
    # (d) 起動時に変えた設定を、利用者が起動したときの状態に戻す。(e) 窓を出し、COM の参照を放しても利用者が閉じるまで残るようにする。
    # 窓・UserControl・DisplayAlerts が戻らないまま参照を放すと、Excel が保存の確認なしに終わりうるため、成功を確かめ、数回やり直す
    $required = @(@("DisplayAlerts", $true), @("Visible", $true), @("UserControl", $true))
    $optional = @(@("EnableEvents", $true), @("ScreenUpdating", $true), @("AskToUpdateLinks", $true), @("AutomationSecurity", 1))
    $pending = @($required)
    for ($try = 0; $try -lt 3 -and $pending.Count -gt 0; $try++) {
        if ($try -gt 0) { Start-Sleep -Milliseconds 200 }
        $failed = @()
        foreach ($setting in $pending) {
            try { $com.($setting[0]) = $setting[1] } catch { $failed += , $setting }
        }
        $pending = $failed
    }
    foreach ($setting in $optional) {
        try { $com.($setting[0]) = $setting[1] } catch {}
    }
    $ok = ($pending.Count -eq 0)

    if ($ok) {
        # (f) このスレッドが持つ COM を解放しきる。WaitForPendingFinalizers は長く止まることがあるため使わない
        try {
            if ($books) {
                foreach ($book in @($books.Own) + @($books.Foreign)) { try { releaseComObject $book } catch {} }
            }
            releaseComObject $com
            [GC]::Collect()
        } catch {}
    } else {
        # 窓を出せなかった Excel は、参照を放すと終わるおそれがあるため、放さずに持ち続ける（止めも強制終了もしない）。
        # 次の確かめの時機に retryKeptApps が設定し直す。インデックス作成が終わると取り込みのスレッドごと参照が切れる
        $script:officeKeptApps += , @{ Name = $name; Com = $com; Pending = $pending }
    }

    if ($script:onOfficeHandOver) {
        try { & $script:onOfficeHandOver $name $ok } catch {}
    }
    return $ok
}

function retryKeptApps {
    # 窓を出せずに持ち続けている Excel の設定をもう一度戻し、通ったら参照を放して「渡した」ログを書く。待たない（通らなければ次の時機に回す）
    if (@($script:officeKeptApps).Count -eq 0) {
        return
    }
    $remaining = @()
    foreach ($kept in @($script:officeKeptApps)) {
        $ok = $true
        foreach ($setting in @($kept.Pending)) {
            try { $kept.Com.($setting[0]) = $setting[1] } catch { $ok = $false }
        }
        if ($ok) {
            try { releaseComObject $kept.Com; [GC]::Collect() } catch {}
            if ($script:onOfficeHandOver) {
                try { & $script:onOfficeHandOver $kept.Name $true } catch {}
            }
        } else {
            $remaining += , $kept
        }
    }
    $script:officeKeptApps = $remaining
}

function handOverForeignApp {
    # 起動した Excel に利用者が開いたブックがあれば、利用者に渡して $true を返す。無ければ何もせず $false
    param ([string]$name)

    # 前に窓を出せなかった Excel があれば、ここで設定し直す
    retryKeptApps
    $app = $script:apps[$name]
    if ($name -ne "Excel" -or $null -eq $app -or $app.Shared) {
        return $false
    }
    $count = getForeignWorkbookCount $app.Com
    if ($count -eq 0) {
        return $false
    }
    # ブックの一覧を読めない（COM が呼び出しを拒んだ。利用者が操作中のことがある）ときは、見える窓があれば利用者が使っているとみなし、止めない
    if ($count -lt 0 -and -not (testProcessHasWindow ([int]$app.Pid) $appInfo["Excel"].Process)) {
        return $false
    }
    [void](handOverApp $name)
    return $true
}

function stopApp {
    param (
        [string]$name
    )

    $app = $script:apps[$name]
    if ($null -eq $app) {
        return
    }
    # インデックス作成中に利用者が Excel でブックを開いた場合は、閉じずに利用者に渡す
    if (handOverForeignApp $name) {
        return
    }
    $script:apps.Remove($name)
    updateWatchedPids

    # インデックス作成中に利用者が同じアプリでファイルを開いた場合は、終了させない
    $inUse = $app.Shared
    if (-not $inUse) {
        try {
            switch ($name) {
                "Word"       { $inUse = ($app.Com.Documents.Count -gt 0) }
                "PowerPoint" { $inUse = ($app.Com.Presentations.Count -gt 0) }
            }
        } catch {}
    }

    if (-not $inUse) {
        try { $app.Com.Quit() } catch {}
    }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app.Com) } catch {}

    # GC::WaitForPendingFinalizers() はCOMの解放待ちで長時間（約60秒）止まることがあるため使わず、
    # 終了しなかったアプリはプロセスIDを指定して強制終了する
    $exited = $true
    if (-not $inUse -and $app.Pid) {
        $process = Get-Process -Id $app.Pid -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit($appInfo[$name].ExitWait)) {
            # 終了処理中のプロセスは Kill() が「アクセス拒否」で失敗することがあるが、そのまま終了するため無視する
            try { $process.Kill() } catch {}
            # 強制終了は非同期のため、同じ上限で終わるのを待つ。それでも残った場合は、次の getApp が
            # 利用者のものとみなして後回しにする（安全な側に倒れる）
            $exited = [bool]$process.WaitForExit($appInfo[$name].ExitWait)
        }
    }
    # 記録は、プロセスが終わったと確かめてから消す。終わらなければ残す（次の起動の確認に任せる）。
    # 利用者がファイルを開いていて止めなかったとき（inUse）は、利用者に渡したものとして消す
    if ($app.Pid -and $script:officeRecordDir -and ($inUse -or $exited)) {
        removeOfficeRecord $script:officeRecordDir ([int]$app.Pid)
    }
    if ($app.Pid -and $script:officePidSink) {
        $removed = $null
        [void]$script:officePidSink.TryRemove([int]$app.Pid, [ref]$removed)
    }
}

function stopAllApps {
    retryKeptApps
    # 1つのアプリの終了に失敗しても、残りのアプリは終了させる
    foreach ($name in @($script:apps.Keys)) {
        try {
            stopApp $name
        } catch {
            Write-Host "    ${name} の終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function getAppName {
    # 拡張子から、抽出に使うアプリの名前を返す
    param (
        [string]$path
    )

    switch -Regex ([System.IO.Path]::GetExtension($path).ToLower()) {
        "^\.xls" { return "Excel" }
        "^\.doc" { return "Word" }
        "^\.ppt" { return "PowerPoint" }
    }
    return $null
}

function releaseComObject($object) {
    if ($null -ne $object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object)
    }
}

# ----------------------------------------------------------------------------
# 1ファイルの制限時間の監視
# ----------------------------------------------------------------------------

# 取り込み中のCOM呼び出しは応答が無いと戻らず、Ctrl+C も効かないため、別スレッドで制限時間を監視する。
# 制限時間を過ぎたら、自分で起動したOfficeアプリを強制終了する（COM呼び出しが例外で戻り、そのファイルは失敗になる）。
#   Deadline: 取り込み中のファイルの制限時刻（取り込み中でなければ MaxValue）
#   Pids    : 強制終了してよいプロセスID（自分で起動したOfficeアプリ）
#   TimedOut: 制限時間を過ぎて強制終了した
$script:watchdog = [hashtable]::Synchronized(@{ Deadline = [datetime]::MaxValue; Pids = @(); TimedOut = $false; Stop = $false })
$script:watchdogThread = $null

function updateWatchedPids {
    # 強制終了してよいプロセスIDを、起動中のアプリのうち自分で起動したものにする（利用者のアプリは終了させない）
    $script:watchdog.Pids = @($script:apps.Values | Where-Object { -not $_.Shared -and $_.Pid } | ForEach-Object { $_.Pid })
}

function startWatchdog {
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript({
        param($watch, [string[]]$processNames)
        while (-not $watch.Stop) {
            Start-Sleep -Milliseconds 500
            if ([datetime]::Now -lt $watch.Deadline) {
                continue
            }
            $watch.Deadline = [datetime]::MaxValue
            $watch.TimedOut = $true
            foreach ($id in @($watch.Pids)) {
                try {
                    # 終了済みでIDが別のプロセスに再利用されている場合に備え、Officeアプリであることを確かめる
                    $process = [System.Diagnostics.Process]::GetProcessById($id)
                    if ($processNames -contains $process.ProcessName) {
                        $process.Kill()
                    }
                } catch {}
            }
        }
    }).AddArgument($script:watchdog).AddArgument([string[]]@($appInfo.Values | ForEach-Object { $_.Process }))
    $script:watchdogThread = @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
}

function stopWatchdog {
    if ($null -eq $script:watchdogThread) {
        return
    }
    $script:watchdog.Stop = $true
    try { [void]$script:watchdogThread.PowerShell.EndInvoke($script:watchdogThread.Handle) } catch {}
    $script:watchdogThread.PowerShell.Dispose()
    $script:watchdogThread = $null
}
