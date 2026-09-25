# Office の操作の鍵（shared\office\office_app.ps1 の lockOfficeProcess）を、ほかのスレッドが持っている状態を作る。
#   $holder = holdOfficeLock "PowerPoint"   # 別のスレッドが鍵を取るまで待って返す
#   releaseOfficeLock $holder               # 鍵を放させ、スレッドを片づける
# 鍵の名前は lockOfficeProcess と同じ（プロセスごと）

function holdOfficeLock {
    param ([string]$name)

    $acquired = New-Object System.Threading.ManualResetEvent($false)
    $release = New-Object System.Threading.ManualResetEvent($false)
    $ps = [powershell]::Create()
    [void]$ps.AddScript({
        param ($mutexName, $acquired, $release)
        $mutex = New-Object System.Threading.Mutex($false, $mutexName)
        [void]$mutex.WaitOne()
        [void]$acquired.Set()
        [void]$release.WaitOne()
        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }).AddArgument("Local\tebunko_office_${name}_${PID}").AddArgument($acquired).AddArgument($release)
    $holder = @{ PowerShell = $ps; Handle = $ps.BeginInvoke(); Release = $release }
    if (!$acquired.WaitOne(10000)) {
        throw "別のスレッドが鍵を取れませんでした"
    }
    return $holder
}

function releaseOfficeLock {
    param ([hashtable]$holder)

    [void]$holder.Release.Set()
    try { [void]$holder.PowerShell.EndInvoke($holder.Handle) } catch {}
    $holder.PowerShell.Dispose()
}
