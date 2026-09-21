# Officeアプリ（Excel・Word・PowerPoint）の起動・終了と、1ファイルの制限時間の監視。
# COM を使う処理だけをまとめる。使う側が dot-source する（shared.ps1 では読み込まない）。

# ----------------------------------------------------------------------------
# Officeアプリ（Excel・Word・PowerPoint）の起動・終了
# ----------------------------------------------------------------------------

# 起動中のアプリ: 名前 → @{ Com; Pid; Shared }
$script:apps = @{}

$appInfo = @{
    Excel      = @{ ProgId = "Excel.Application";      Process = "EXCEL" }
    Word       = @{ ProgId = "Word.Application";       Process = "WINWORD" }
    PowerPoint = @{ ProgId = "PowerPoint.Application"; Process = "POWERPNT" }
}

function getApp {
    # アプリのCOMオブジェクトを返す。起動していなければ起動する
    param (
        [string]$name
    )

    if (-not $script:apps.ContainsKey($name)) {
        $info = $appInfo[$name]

        # 終了できなかった場合に強制終了するため、新しく起動したプロセスのIDを控えておく
        $before = @(Get-Process -Name $info.Process -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $com = New-Object -ComObject $info.ProgId
        $after = @(Get-Process -Name $info.Process -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $newIds = @($after | Where-Object { $before -notcontains $_ })

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

    # 変換中に利用者が同じアプリでファイルを開いた場合は、終了させない
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
        if ($process -and -not $process.WaitForExit(5000)) {
            # 終了処理中のプロセスは Kill() が「アクセス拒否」で失敗することがあるが、そのまま終了するため無視する
            try { $process.Kill() } catch {}
        }
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
    # 拡張子から、変換に使うアプリの名前を返す
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

# 変換中のCOM呼び出しは応答が無いと戻らず、Ctrl+C も効かないため、別スレッドで制限時間を監視する。
# 制限時間を過ぎたら、自分で起動したOfficeアプリを強制終了する（COM呼び出しが例外で戻り、そのファイルは失敗になる）。
#   Deadline: 変換中のファイルの制限時刻（変換中でなければ MaxValue）
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
