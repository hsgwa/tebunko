# ファイルの変更の見張り（FileSystemWatcher）と、更新の印。
#
# 窓を前面にするたびにファイルを読み直さず、変わったときだけ読み直すために使う（docs/design/gui/responsiveness.md）。
# 合図は Register-ObjectEvent の -Action なしで、PowerShell のイベントの列に積むだけにする
# （Watcher のイベントは thread pool のスレッドで起き、Runspace が無いため、そこでスクリプトブロックは動かせない）。
# 見るのはメモリの中の列だけで、ディスクには触らない。ネットワークの場所は見張らない（通知が届かないことがあり、作るときに共有に触れるため）。
# 作れない・Error になったときは、呼び出し側の保険（更新の印の比較）に任せる。

function newFileWatch {
    # Path のファイルを見張る。見張れない（ネットワーク・フォルダが無い・作れない）ときは $null を返す
    param ([string]$Path)

    if ([string]::IsNullOrEmpty($Path)) {
        return $null
    }
    try {
        $full = [System.IO.Path]::GetFullPath($Path)
        $dir = [System.IO.Path]::GetDirectoryName($full)
        if ([string]::IsNullOrEmpty($dir) -or (testNetworkPath $full)) {
            return $null
        }
        if (!(Test-Path -LiteralPath $dir -PathType Container)) {
            return $null
        }
        $watcher = New-Object System.IO.FileSystemWatcher
        $id = "FileWatch-" + [guid]::NewGuid().ToString("N")
        $watch = [pscustomobject]@{ Watcher = $watcher; SourceId = $id; Subscribers = @() }
        try {
            $watcher.Path = $dir
            $watcher.Filter = [System.IO.Path]::GetFileName($full)
            $watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size -bor [System.IO.NotifyFilters]::FileName
            foreach ($eventName in "Changed", "Created", "Deleted", "Renamed", "Error") {
                $watch.Subscribers += Register-ObjectEvent -InputObject $watcher -EventName $eventName -SourceIdentifier "$id-$eventName"
            }
            $watcher.EnableRaisingEvents = $true
        } catch {
            closeFileWatch $watch
            return $null
        }
        return $watch
    } catch {
        return $null
    }
}

function testFileWatchSignaled {
    # 合図が来ていれば $true を返して合図を消す。見張りが無い（$null）ときは $false（呼び出し側が保険で見る）
    param ($Watch)

    if ($null -eq $Watch) {
        return $false
    }
    $signaled = $false
    foreach ($eventName in "Changed", "Created", "Deleted", "Renamed", "Error") {
        $events = @(Get-Event -SourceIdentifier "$($Watch.SourceId)-$eventName" -ErrorAction SilentlyContinue)
        if ($events.Count -gt 0) {
            $signaled = $true
            $events | Remove-Event -ErrorAction SilentlyContinue
        }
    }
    return $signaled
}

function getFileStamp {
    # ファイルの更新の印（更新日時と長さの 1 行）。無い・読めないときは空
    param ([string]$Path)

    try {
        $info = New-Object System.IO.FileInfo $Path
        if (!$info.Exists) {
            return ""
        }
        return "{0}|{1}" -f $info.LastWriteTimeUtc.Ticks, $info.Length
    } catch {
        return ""
    }
}

function closeFileWatch {
    # 見張りを止めて片づける。$null・2 回目でも失敗しない
    param ($Watch)

    if ($null -eq $Watch) {
        return
    }
    try { $Watch.Watcher.EnableRaisingEvents = $false } catch { Write-Verbose $_.Exception.Message }
    foreach ($eventName in "Changed", "Created", "Deleted", "Renamed", "Error") {
        try { Unregister-Event -SourceIdentifier "$($Watch.SourceId)-$eventName" -ErrorAction SilentlyContinue } catch { Write-Verbose $_.Exception.Message }
        try { Get-Event -SourceIdentifier "$($Watch.SourceId)-$eventName" -ErrorAction SilentlyContinue | Remove-Event -ErrorAction SilentlyContinue } catch { Write-Verbose $_.Exception.Message }
    }
    try { $Watch.Watcher.Dispose() } catch { Write-Verbose $_.Exception.Message }
}
