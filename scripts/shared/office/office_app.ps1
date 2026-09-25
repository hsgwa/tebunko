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

# 起動した Office のプロセスの優先度。利用者の操作・検索を先に動かすため下げる（docs/00_共通_4_プロセスとスレッド.md 7.2）
$officePriority = [System.Diagnostics.ProcessPriorityClass]::BelowNormal
# 起動したアプリの PID を入れる入れ物（ConcurrentDictionary[int,string]。$null なら入れない）。
# 取り込みを複数のスレッドで行うとき、画面が閉じるときに、インデックス作成が起動した Office を PID で止めるために使う
$script:officePidSink = $null

function lockOfficeProcess {
    # プロセスの中で 1 つずつ行う Office の操作の鍵（名前付きミューテックス）を取る。解放は unlockOfficeProcess。
    #   start     : 起動（起動の前後のプロセスの一覧の差で PID を調べるため、同時に起動すると取り違える）
    #   PowerPoint: PowerPoint を使う操作（PowerPoint は 1 つのプロセスしか持てず、ほかのスレッドと共有になるため）
    param (
        [string]$name
    )

    $mutex = New-Object System.Threading.Mutex($false, "Local\tebunko_office_${name}_${PID}")
    try {
        [void]$mutex.WaitOne()
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
            try { (Get-Process -Id $newIds[0]).PriorityClass = $officePriority } catch {}
            if ($script:officePidSink) {
                $script:officePidSink[[int]$newIds[0]] = $info.Process
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
