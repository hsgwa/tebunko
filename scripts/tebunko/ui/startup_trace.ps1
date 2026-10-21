# 起動の段階の記録（どの段階に何ミリ秒かかったかを、画面が出るまでの間、メモリに控える）。
#
# 環境変数 TEBUNKO_STARTUP_TRACE が 1 のときだけ動く（無ければ何も控えず、何も書かない）。
# 控えるのは「段階の名前」と「PowerShell のプロセスが始まってからの経過ミリ秒」の組。段階の時間は、隣り合う行の差で求める
# （最初の行はプロセスが始まってから gui.ps1 が動き出すまで）。計測の道具（tools\perf\measure_startup.ps1）が読む。
# 書き出しは、起動時の読み込み（loadStartupData）が済んだ後に 1 回だけ、ワークスペース直下の起動の記録（Workspace の StartupTraceFile）へ行う。
# gui.ps1 の最初に読み込む（shared.ps1 より前から測るため、読み込み口の外の部品だけを使う）。
# 書けなくても例外を出さない（記録のせいで起動を止めない）。

$script:startupTrace = $null
$script:startupWatch = $null
$script:startupTraceOffset = 0

function startStartupTrace {
    # flag は環境変数 TEBUNKO_STARTUP_TRACE の値。"1" のときだけ記録を始め、最初の行（PowerShell の起動）を控える
    param (
        [string]$flag
    )

    $script:startupTrace = $null
    if ($flag -ne "1") {
        return
    }
    try {
        $script:startupWatch = [System.Diagnostics.Stopwatch]::StartNew()
        $started = [System.Diagnostics.Process]::GetCurrentProcess().StartTime
        $script:startupTraceOffset = [long]([DateTime]::Now - $started).TotalMilliseconds
        $script:startupTrace = New-Object System.Collections.Generic.List[string]
        markStartup "PowerShell の起動"
    } catch {
        $script:startupTrace = $null
    }
}

function markStartup {
    # 段階の名前と、プロセスが始まってからの経過ミリ秒を控える（記録を始めていなければ何もしない）
    param (
        [string]$name
    )

    if ($null -eq $script:startupTrace) {
        return
    }
    $elapsed = $script:startupTraceOffset + $script:startupWatch.ElapsedMilliseconds
    $script:startupTrace.Add("${name}`t${elapsed}")
}

function writeStartupTrace {
    # 控えた記録を path に書く（上書き）。記録を始めていない・path が空のときは何もしない。書き終えたら、控えを捨てる（1 回だけ書く）
    param (
        [string]$path
    )

    if ($null -eq $script:startupTrace -or [string]::IsNullOrEmpty($path)) {
        return
    }
    $lines = $script:startupTrace
    $script:startupTrace = $null
    try {
        [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        # 記録のせいで起動を止めない（書けなければ残さない）
        $null = $_.Exception
    }
}
