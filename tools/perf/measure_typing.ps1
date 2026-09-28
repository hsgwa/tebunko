# 入力しながらの検索（インクリメンタルサーチ）の速さを測る。tools\measure_perf.ps1 が別のプロセスで起動する。
# 語の前から1文字ずつ伸ばした語（getTypingSteps）を、画面と同じ検索の司令のスレッド（SearchService）へ、決まった間隔で渡す。
# 新しい要求は前の要求を取り消す（SearchService.Request）。これを1回の入力として、-Count回繰り返す。
# 測るのは、最後の要求を渡す直前（0の時刻）から、前の要求がすべて終わるまで（WaitPrevMs）・最初のヒットまで（FirstHitMs）・
# 検索が終わるまで（FinishMs）。要求を渡す時間（RequestMs）も、これらに含める。
# 待つ間はビジーループにせず Thread.Sleep(1) で待つ（実際には Windows のタイマーの細かさの分、約 15.6 ms ごとになる）。
# 検索の司令のスレッドが無い版（#102 より前）は、入力しながらの検索が無いので測らない（Mode を none にし、Skipped に理由を書く。失敗にはしない）。
# 結果（1 回の入力ごとの記録・リソースの記録）を -Out の JSON に書く。
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Index,
    [Parameter(Mandatory = $true)][string]$WordFile,
    [int]$Count = 20,
    [int]$IntervalMs = 150,
    [int]$MinLength = 3,
    [switch]$Fast,
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$SampleMs = 200
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\perf_common.ps1"
$lib = resolveTebunkoLib $Tool
. $lib
assertTebunkoFunctions @("newSearchRequest", "newTsvTextCache")

$spec = [System.IO.File]::ReadAllText($WordFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$Name = [string]$spec.Name
$Word = [string]$spec.Word
$Regex = [bool]$spec.Regex

function roundMs($value) { if ($null -eq $value) { return $null }; return [Math]::Round([double]$value, 1) }

function writeTypingOut([hashtable]$obj) {
    [System.IO.File]::WriteAllText($Out, ($obj | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($false)))
}

$hasService = [bool](Get-Command newSearchService -CommandType Function -ErrorAction SilentlyContinue)
$mode = if ($hasService) { "service" } else { "none" }
if (!$hasService) {
    writeTypingOut ([ordered]@{
        Name = $Name; Word = $Word; Regex = $Regex; Mode = $mode; Count = $Count; IntervalMs = $IntervalMs; MinLength = $MinLength; Fast = [bool]$Fast
        Steps = 0; Attempts = @(); Samples = @(); PeakWorkingSetMB = 0.0
        Skipped = "検索の司令のスレッド（newSearchService）が無い版のため、入力しながらの検索は測れません"
    })
    return
}

if ($Fast) {
    $wantIndex = (Join-Path $Work "content_index").TrimEnd("\")
    if (!$Index.TrimEnd("\").Equals($wantIndex, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "-TypingFast では、-Index を <Work>\content_index にしてください。"
    }
    if (!(Get-Command newSearchRequest -CommandType Function).Parameters.ContainsKey("workDir")) {
        throw "測る版は、高速検索のワークスペースを要求で渡せません。"
    }
}

$steps = getTypingSteps $Word $Regex $MinLength
if ($steps.Count -eq 0) {
    writeTypingOut ([ordered]@{
        Name = $Name; Word = $Word; Regex = $Regex; Mode = $mode; Count = $Count; IntervalMs = $IntervalMs; MinLength = $MinLength; Fast = [bool]$Fast
        Steps = 0; Attempts = @(); Samples = @(); PeakWorkingSetMB = 0.0
        Skipped = "送る語がありません（すべて正規表現として不正、または空の文字列に一致する）"
    })
    return
}

# 1 回の入力を送る。前の要求（最後を除く）が終わるまで・最後の要求の最初のヒットまで・終わるまでの時間（ms）と、
# 要求ごとの間隔・遅れ・渡す時間、最後の要求の件数などを返す
function invokeTypingAttempt {
    param (
        $service,
        [string[]]$steps,
        [bool]$isRegex,
        [int]$intervalMs,
        [string]$index,
        [bool]$useFast,
        [string]$workDir,
        [int]$n
    )

    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $requests = New-Object System.Collections.Generic.List[object]
    $intervals = New-Object System.Collections.Generic.List[double]
    $lags = New-Object System.Collections.Generic.List[double]
    $reqTimes = New-Object System.Collections.Generic.List[double]
    $prevSendMs = 0.0
    $zeroMs = 0.0
    for ($k = 0; $k -lt $steps.Count; $k++) {
        $planned = $k * $intervalMs
        while ($watch.Elapsed.TotalMilliseconds -lt $planned) { [System.Threading.Thread]::Sleep(1) }
        $isLast = ($k -eq $steps.Count - 1)
        if ($isLast) { $zeroMs = $watch.Elapsed.TotalMilliseconds }
        $sendWatch = [System.Diagnostics.Stopwatch]::StartNew()
        # 古い版（#102〜#107）の newSearchRequest には workDir が無いため、-Fast でないときは今までどおり 4 引数だけで呼ぶ
        $request = if ($useFast) {
            newSearchRequest -word $steps[$k] -simpleMatch (!$isRegex) -folders @($index) -limit 10000 -useFast $true -workDir $workDir
        } else {
            newSearchRequest $steps[$k] (!$isRegex) @($index) 10000
        }
        $req = $service.Request($request)
        $reqTimes.Add($sendWatch.Elapsed.TotalMilliseconds)
        $sendMs = $watch.Elapsed.TotalMilliseconds
        $lags.Add($sendMs - $planned)
        if ($k -gt 0) { $intervals.Add($sendMs - $prevSendMs) }
        $prevSendMs = $sendMs
        $requests.Add($req)
    }

    $last = $requests[$requests.Count - 1]
    $prevRequests = @($requests | Select-Object -First ($requests.Count - 1))
    $waitPrevMs = $null; $firstHitMs = $null; $finishMs = $null; $failure = $null
    while ($true) {
        $elapsed = $watch.Elapsed.TotalMilliseconds - $zeroMs
        if ($null -eq $waitPrevMs -and (@($prevRequests | Where-Object { !$_.Finished }).Count -eq 0)) { $waitPrevMs = $elapsed }
        if ($null -eq $firstHitMs -and !$last.Queue.IsEmpty) { $firstHitMs = $elapsed }
        if ($last.Finished) { $finishMs = $watch.Elapsed.TotalMilliseconds - $zeroMs; break }
        if (!$service.IsRunning()) { $failure = $service.GetFailure(); break }
        [System.Threading.Thread]::Sleep(1)
    }
    if ($failure) { throw "検索のスレッドが止まりました: $failure" }
    if ($last.Error) { throw "検索に失敗しました: $($last.Error)" }

    $cancelledSteps = @($prevRequests | Where-Object { $_.Cancelled }).Count
    $hits = 0
    $hit = $null
    while ($last.Queue.TryDequeue([ref]$hit)) { $hits++ }

    $attempt = [ordered]@{
        N = $n; Steps = $steps.Count
        FirstHitMs = (roundMs $firstHitMs); FinishMs = (roundMs $finishMs); WaitPrevMs = (roundMs $waitPrevMs)
        IntervalMs = @($intervals.ToArray() | ForEach-Object { roundMs $_ })
        LagMs = @($lags.ToArray() | ForEach-Object { roundMs $_ })
        RequestMs = @($reqTimes.ToArray() | ForEach-Object { roundMs $_ })
        Hits = $hits; Truncated = [bool]$last.Truncated; Packs = [int]$last.IndexTotal; FastUsed = [bool]$last.FastUsed
        CancelledSteps = $cancelledSteps
    }
    $snap = getProcessSnapshot
    foreach ($key in $snap.Keys) { $attempt[$key] = $snap[$key] }
    return $attempt
}

$monitor = startResourceMonitor $SampleMs
$attempts = New-Object System.Collections.Generic.List[object]
$cache = newTsvTextCache
$service = $null
try {
    setMonitorPhase $monitor "入力しながらの検索"
    $service = newSearchService $lib $cache
    for ($n = 1; $n -le $Count; $n++) {
        $attempts.Add((invokeTypingAttempt $service $steps $Regex $IntervalMs $Index ([bool]$Fast) $Work $n))
        Write-Host ("  ${Name} 入力 {0} / {1} 回目" -f $n, $Count)
    }
} finally {
    if ($service) { $service.Close() }
    $samples = stopResourceMonitor $monitor
}
writeTypingOut ([ordered]@{
    Name = $Name; Word = $Word; Regex = $Regex; Mode = $mode; Count = $Count; IntervalMs = $IntervalMs; MinLength = $MinLength; Fast = [bool]$Fast
    Steps = $steps.Count; Attempts = $attempts.ToArray(); Samples = $samples
    PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
    Skipped = $null
})
