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
    PowerPoint = @{ ProgId = "PowerPoint.Application"; Process = "POWERPNT"; ExitWait = 5000 }
}

# 起動した Office のプロセスの優先度は下げない（Normal のまま）。利用者がダブルクリックしたファイルがインデックス作成の Excel・Word で開くことがあり、
# PowerPoint は 1 つのプロセスしか持てないため、利用者とプロセスを共有しないと確実には言えない。利用者の操作を遅くしないよう、
# 優先度を下げるのは、利用者と共有しないインデックス作成のスレッドだけにする（docs/00_共通_4_プロセスとスレッド.md 7.2）
# 起動したアプリの PID を入れる入れ物（ConcurrentDictionary[int,string]。$null なら入れない）。
# 取り込みを複数のスレッドで行うとき、画面が閉じるときに、インデックス作成が起動した Office を PID で止めるために使う
$script:officePidSink = $null

function lockOfficeProcess {
    # プロセスの中で 1 つずつ行う Office の操作の鍵（名前付きミューテックス）を取る。解放は unlockOfficeProcess。
    #   start     : 起動（起動の前後のプロセスの一覧の差で PID を調べるため、同時に起動すると取り違える）
    #   PowerPoint: PowerPoint を使う操作（PowerPoint は 1 つのプロセスしか持てず、ほかのスレッドと共有になるため）
    # timeout（ミリ秒。-1 は取れるまで待つ）までに取れなければ $null を返す。同じスレッドは続けて取れる（取った数だけ解放する）
    param (
        [string]$name,
        [int]$timeout = -1
    )

    $mutex = New-Object System.Threading.Mutex($false, "Local\tebunko_office_${name}_${PID}")
    try {
        if (!$mutex.WaitOne($timeout)) {
            $mutex.Dispose()
            return $null
        }
    } catch [System.Threading.AbandonedMutexException] {
        # 持っていたスレッドが解放せずに終わった。鍵は取れている
    }
    return $mutex
}

function unlockOfficeProcess {
    param (
        [System.Threading.Mutex]$mutex
    )

    $mutex.ReleaseMutex()
    $mutex.Dispose()
}

# 取り込みのスレッドの間で PowerPoint を共有するときの状態（newPowerPointShare。$null なら共有しない）。
# PowerPoint は 1 つのプロセスしか持てないため、一度起動したら、インデックス作成が終わるまで同じプロセスを使い回す。
# COM の部品（つながり）は作ったスレッドでしか使えないため、スレッドごとにつなぐ
$script:powerPointShare = $null
# PowerPoint がほかの取り込みのスレッドで使われているときの例外の文言（後回しにする。invokeIngestTask）
${powerPointBusyMessage} = "PowerPoint はほかの取り込みで使用中です。"

function newPowerPointShare {
    # PowerPoint を共有するときの状態を作る（司令が作り、取り込みのスレッドに渡す）。
    #   Pid: 使っている PowerPoint の PID（0 は無い）/ Owned: インデックス作成が起動したものか（利用者の PowerPoint なら $false）/
    #   Workers: 動いている取り込みのスレッドの数（最後のスレッドが終了する）
    return [hashtable]::Synchronized(@{ Pid = 0; Owned = $false; Workers = 0 })
}

function enterPowerPointShare {
    # 取り込みのスレッドが始まるときに呼ぶ
    param (
        [hashtable]$share
    )

    [System.Threading.Monitor]::Enter($share.SyncRoot)
    try {
        $share.Workers = $share.Workers + 1
    } finally {
        [System.Threading.Monitor]::Exit($share.SyncRoot)
    }
}

function exitPowerPointShare {
    # 取り込みのスレッドが終わるときに呼ぶ。最後のスレッドなら、インデックス作成が起動した PowerPoint を終了する
    param (
        [hashtable]$share
    )

    [System.Threading.Monitor]::Enter($share.SyncRoot)
    try {
        $share.Workers = $share.Workers - 1
        $last = ($share.Workers -le 0)
    } finally {
        [System.Threading.Monitor]::Exit($share.SyncRoot)
    }
    if ($last) {
        closeSharedPowerPoint $share
    }
}

function closeSharedPowerPoint {
    # インデックス作成が起動した PowerPoint を終了する（利用者の PowerPoint・利用者がファイルを開いている PowerPoint は終了しない）
    param (
        [hashtable]$share
    )

    $lock = lockOfficeProcess "PowerPoint"
    try {
        if (!$share.Owned -or !$share.Pid) {
            return
        }
        $process = Get-Process -Id $share.Pid -ErrorAction SilentlyContinue
        if ($process -and $process.ProcessName -eq $appInfo.PowerPoint.Process) {
            $inUse = $false
            $com = $null
            try {
                # 起動済みの PowerPoint につなぐ（PowerPoint は 1 つのプロセスしか持てないため、新しくは起動しない）
                $com = New-Object -ComObject $appInfo.PowerPoint.ProgId
                $inUse = ($com.Presentations.Count -gt 0)
                if (!$inUse) {
                    $com.Quit()
                }
            } catch {
            } finally {
                try { releaseComObject $com } catch {}
            }
            if (!$inUse -and !$process.WaitForExit($appInfo.PowerPoint.ExitWait)) {
                try { $process.Kill() } catch {}
            }
        }
        if ($script:officePidSink) {
            $removed = $null
            [void]$script:officePidSink.TryRemove([int]$share.Pid, [ref]$removed)
        }
        $share.Pid = 0
        $share.Owned = $false
    } finally {
        unlockOfficeProcess $lock
    }
}

function testComDisconnected {
    # COM の呼び出しの例外が「つながっていない」（相手の Office のプロセスが終わった等）ものか
    param (
        [System.Exception]$exception
    )

    $base = $exception.GetBaseException()
    # RPC_E_DISCONNECTED / RPC_S_SERVER_UNAVAILABLE / RPC_S_CALL_FAILED / CO_E_OBJNOTCONNECTED
    # （PowerShell 5.1 は 0x8… の値を HResult と同じ Int32（負の数）として読む）
    return @(0x80010108, 0x800706BA, 0x800706BE, 0x800401FD) -contains $base.HResult
}

function getApp {
    # アプリのCOMオブジェクトを返す。起動していなければ起動する
    param (
        [string]$name
    )

    if (-not $script:apps.ContainsKey($name)) {
        $info = $appInfo[$name]

        # 終了できなかった場合に強制終了するため、新しく起動したプロセスのIDを控えておく。
        # 取り込みのスレッドが同時に起動すると、どれが自分の起動したものか分からなくなるため、起動は 1 つずつ行う
        $lock = lockOfficeProcess "start"
        try {
            $before = @(Get-Process -Name $info.Process -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
            $com = New-Object -ComObject $info.ProgId
            $after = @(Get-Process -Name $info.Process -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        } finally {
            unlockOfficeProcess $lock
        }
        $newIds = @($after | Where-Object { $before -notcontains $_ })
        if ($newIds.Count -eq 1) {
            if ($script:officePidSink) {
                $script:officePidSink[[int]$newIds[0]] = $info.Process
            }
        }
        if ($name -eq "PowerPoint" -and $script:powerPointShare) {
            # 共有する PowerPoint を記録する（PowerPoint を使う操作は鍵を取って行うため、ここは 1 つずつ通る）。
            # 起動しなかったときは、ほかのスレッドが起動したものか、利用者の PowerPoint につないでいる
            $share = $script:powerPointShare
            if ($newIds.Count -eq 1) {
                $share.Pid = [int]$newIds[0]
                $share.Owned = $true
            } elseif (!$share.Pid -or !(Get-Process -Id $share.Pid -ErrorAction SilentlyContinue)) {
                $share.Pid = 0
                $share.Owned = $false
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

        # PowerPoint等は起動中のアプリに接続することがある。その場合は利用者のものなので終了させない
        $script:apps[$name] = @{
            Com    = $com
            Pid    = $(if ($newIds.Count -eq 1) { $newIds[0] } else { 0 })
            Shared = ($newIds.Count -eq 0)
        }
        updateWatchedPids
    }
    return $script:apps[$name].Com
}

function stopApp {
    param (
        [string]$name
    )

    $app = $script:apps[$name]
    if ($null -eq $app) {
        return
    }
    $script:apps.Remove($name)
    updateWatchedPids

    if ($name -eq "PowerPoint" -and $script:powerPointShare) {
        # 共有している PowerPoint は、ほかのスレッドも使うため終了しない。このスレッドのつながりだけを放す
        # （終了は最後のスレッドが行う。closeSharedPowerPoint）
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app.Com) } catch {}
        return
    }

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
    if (-not $inUse -and $app.Pid) {
        $process = Get-Process -Id $app.Pid -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit($appInfo[$name].ExitWait)) {
            # 終了処理中のプロセスは Kill() が「アクセス拒否」で失敗することがあるが、そのまま終了するため無視する
            try { $process.Kill() } catch {}
        }
    }
    if ($app.Pid -and $script:officePidSink) {
        $removed = $null
        [void]$script:officePidSink.TryRemove([int]$app.Pid, [ref]$removed)
    }
}

function stopAllApps {
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
# このスレッドが共有の PowerPoint を使っている（鍵を持っている）間 $true。そのときだけ共有の PowerPoint を監視の対象に入れる
$script:powerPointHeld = $false

function updateWatchedPids {
    # 強制終了してよいプロセスIDを、起動中のアプリのうち自分で起動したものにする（利用者のアプリは終了させない）。
    # 共有の PowerPoint は、このスレッドが使っている間だけ入れる（ほかのスレッドの変換の途中で止めないため）
    $share = $script:powerPointShare
    $ids = @($script:apps.GetEnumerator() | Where-Object {
            -not ($share -and $_.Key -eq "PowerPoint") -and -not $_.Value.Shared -and $_.Value.Pid
        } | ForEach-Object { $_.Value.Pid })
    if ($share -and $script:powerPointHeld -and $share.Owned -and $share.Pid) {
        $ids += $share.Pid
    }
    $script:watchdog.Pids = $ids
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
