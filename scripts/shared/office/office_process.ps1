# 実行中の Excel・Word・PowerPoint の一覧と強制終了。

# ----------------------------------------------------------------------------
# 検索・Officeプロセス・画面（gui.ps1）で共有する処理
# ----------------------------------------------------------------------------

# 強制終了の対象: プロセス名 → 表示名
${officeProcessNames} = [ordered]@{ EXCEL = "Excel"; WINWORD = "Word"; POWERPNT = "PowerPoint" }

function getOfficeProcesses {
    # 実行中の Excel・Word・PowerPoint を返す。
    # ウィンドウを持たない（MainWindowHandle が 0）プロセスは、インデクサなどでバックグラウンド起動されたものとする
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($process in @(Get-Process -Name @(${officeProcessNames}.Keys) -ErrorAction SilentlyContinue)) {
        $startTime = $null
        try {
            $startTime = $process.StartTime
        } catch {
            # 権限の無いプロセスは起動時刻を取得できない
        }
        $result.Add([pscustomobject]@{
            Id          = $process.Id
            ProcessName = $process.ProcessName
            AppName     = ${officeProcessNames}[$process.ProcessName]
            Background  = ($process.MainWindowHandle -eq [IntPtr]::Zero)
            StartTime   = $startTime
            MemoryMB    = [math]::Round($process.WorkingSet64 / 1MB)
            Title       = $process.MainWindowTitle
        })
    }
    return $result.ToArray()
}

function stopOfficeProcesses {
    # 指定したプロセスを保存せずに終了し、@{ Id; Stopped; Message } の配列を返す
    param (
        [int[]]$ids
    )

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($id in $ids) {
        try {
            Stop-Process -Id $id -Force -ErrorAction Stop
            $results.Add(@{ Id = $id; Stopped = $true; Message = "" })
        } catch {
            $results.Add(@{ Id = $id; Stopped = $false; Message = $_.Exception.Message })
        }
    }
    return $results.ToArray()
}
