# テストが自分で起こしたプロセスの後始末。自分で終わるプロセスとの競合（HasExited を見てから止めるまでの間に終わる）で
# テストが揺れないよう、後始末はここの関数だけで行う（tests/meta/test_process_cleanup.Tests.ps1 が、ここ以外の書き方を落とす）。
# 他の helpers に依存しない。

# メソッド呼び出しの例外は PowerShell の例外（MethodInvocationException など）に包まれて届くので、包みをはずして型を見る
function getUnwrappedException {
    param($ErrorRecord)
    $e = $ErrorRecord.Exception
    while ($e -is [System.Management.Automation.RuntimeException] -and $e.InnerException) { $e = $e.InnerException }
    return $e
}
function isInnerException {
    param($ErrorRecord, [type]$Type)
    return ((getUnwrappedException $ErrorRecord) -is $Type)
}

# 自分で起こした Process オブジェクト（Process.Start で起こしてハンドルを持っているもの）を止めて、終わるのを待つ。止める相手は Process オブジェクトで決め、PID から引き直さない
# （PID の使い回しで別のプロセスを止めないため）。握るのは「Kill の間に終わった」InvalidOperationException だけ。
# 止められずに残ったときは、PID と待った時間を入れた例外を投げる（取り残しを見えるようにする）。
# 型は付けない（Kill・WaitForExit・HasExited を持つ偽の物でも確かめられるように）
function stopStartedProcess {
    param($Process, [int]$TimeoutMs = 10000)
    if (-not $Process.HasExited) {
        try {
            $Process.Kill()
        } catch {
            # Kill の間に終わった競合だけを握る。.NET Framework の Kill は、終わったかを確かめてから止めるまでの間に相手が終わると
            # InvalidOperationException でなく Win32Exception（アクセス拒否 = 5）を投げることがある。どちらも HasExited が真のときだけ握る
            $e = getUnwrappedException $_
            $raced = ($e -is [System.InvalidOperationException]) -or (($e -is [System.ComponentModel.Win32Exception]) -and $e.NativeErrorCode -eq 5)
            if (-not $raced -or -not $Process.HasExited) { throw }
        }
    }
    if (-not $Process.WaitForExit($TimeoutMs)) {
        throw "起こしたプロセス（PID $($Process.Id)）が $TimeoutMs ミリ秒待っても終わらなかった"
    }
}

# PID のプロセスが自分で終わるのを待つだけ（止めない）。止めないので PID から引き直してよい（使い回しの害は待つ時間だけ）。
# もう居ない・引く前後に終わったときは何もしない。時間内に終わらなければ例外を投げる
function waitProcessIdExit {
    param([int]$ProcessId, [int]$TimeoutMs = 5000)
    $target = $null
    try {
        $target = [System.Diagnostics.Process]::GetProcessById($ProcessId)
    } catch {
        if (-not (isInnerException $_ ([System.ArgumentException]))) { throw }
        return
    }
    try {
        try {
            if (-not $target.WaitForExit($TimeoutMs)) {
                throw "PID $ProcessId のプロセスが $TimeoutMs ミリ秒待っても終わらなかった"
            }
        } catch {
            # 待つ間に終わって、情報が引けなくなった（それ以外の例外はそのまま投げる）
            if (-not (isInnerException $_ ([System.InvalidOperationException]))) { throw }
            return
        }
    } finally {
        $target.Dispose()
    }
}
