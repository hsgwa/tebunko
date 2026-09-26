# 1 つの語で、同じプロセスのまま -Count 回続けて検索し、1 回ごとの時間とメモリを測る。tools\measure_perf.ps1 が語ごとに別のプロセスで起動する。
# 検索は、測る tebunko の画面と同じ流れで行う。
#   service  … #102 以降。検索の司令のスレッド（newSearchService）をプロセスで 1 つ作り、要求（newSearchRequest）を順に送る。
#              lib.ps1 の読み込みと照合のプールの用意は、司令のスレッドが始めに 1 回だけ行う（1 回目の検索の時間に含まれる）
#   runspace … #102 より前。検索のたびに新しい Runspace で lib.ps1 を読み込み、列挙して照合する（newSearchService が無い版）
# 検索のキャッシュ（newTsvTextCache）は、画面と同じくプロセスで 1 つを持ち続ける。1 回目は「画面を開き直した直後の検索」に当たる。
# 検索する語は -WordFile の JSON（Name・Word・Regex）で受け取る（日本語や正規表現を引数で渡すと崩れることがあるため）。
# 結果（検索 1 回ごとの時間・ヒット件数・メモリと、リソースの記録）を -Out の JSON に書く。
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Index,
    [Parameter(Mandatory = $true)][string]$WordFile,
    [int]$Count = 20,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$SampleMs = 200
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\perf_common.ps1"
$lib = resolveTebunkoLib $Tool
. $lib
assertTebunkoFunctions @("getIndexPackFiles", "searchPackIndex", "newTsvTextCache")
$spec = [System.IO.File]::ReadAllText($WordFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$Name = [string]$spec.Name
$Word = [string]$spec.Word
$Regex = [bool]$spec.Regex
$mode = if (Get-Command newSearchService -CommandType Function -ErrorAction SilentlyContinue) { "service" } else { "runspace" }

# service: 要求を 1 つ送り、終わるまで待つ。列挙が終わった時刻（IndexTotal が -1 から件数になった時刻）で、列挙と照合の時間を分ける。
# 待つ間隔は短くするが、Windows のタイマーの細かさ（約 15 ミリ秒）より細かくは測れない
function invokeServiceSearch($service) {
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $request = $service.Request((newSearchRequest $Word (!$Regex) @($Index) 10000))
    $listMs = $null
    while (!$request.Finished) {
        if ($null -eq $listMs -and $request.IndexTotal -ge 0) { $listMs = $watch.Elapsed.TotalMilliseconds }
        if (!$service.IsRunning()) { throw "検索のスレッドが止まりました: $($service.GetFailure())" }
        [System.Threading.Thread]::Sleep(5)
    }
    $totalMs = $watch.Elapsed.TotalMilliseconds
    if ($request.Error) { throw "検索に失敗しました: $($request.Error)" }
    if ($null -eq $listMs) { $listMs = $totalMs }
    $hits = 0
    $hit = $null
    while ($request.Queue.TryDequeue([ref]$hit)) { $hits++ }
    return @{ TotalMs = $totalMs; LoadMs = $null; ListMs = $listMs; MatchMs = $totalMs - $listMs; Packs = [int]$request.IndexTotal; Hits = $hits; Truncated = [bool]$request.Truncated }
}

# runspace: #102 より前の画面の流れ（検索のたびに新しい Runspace で lib.ps1 を読み込む）
$runspaceScript = {
    param ($lib, $index, $word, $simple, $cache)
    $total = [System.Diagnostics.Stopwatch]::StartNew(); $step = [System.Diagnostics.Stopwatch]::StartNew()
    . $lib
    $load = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $found = getIndexPackFiles @($index)
    $list = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $r = searchPackIndex $word $found.Packs $simple 10000 -cache $cache
    @{ LoadMs = $load; ListMs = $list; MatchMs = $step.Elapsed.TotalMilliseconds; TotalMs = $total.Elapsed.TotalMilliseconds; Packs = $found.Packs.Count; Hits = $r.Hits.Count; Truncated = [bool]$r.Truncated }
}
function invokeRunspaceSearch($cache) {
    $ps = [powershell]::Create()
    try {
        [void]$ps.AddScript($runspaceScript.ToString()).AddArgument($lib).AddArgument($Index).AddArgument($Word).AddArgument(!$Regex).AddArgument($cache)
        return $ps.Invoke()[0]
    } finally {
        $ps.Dispose()
    }
}

function roundMs($value) { if ($null -eq $value) { return $null }; return [Math]::Round([double]$value, 1) }

$monitor = startResourceMonitor $SampleMs
$searches = New-Object System.Collections.Generic.List[object]
$service = $null
try {
    $cache = newTsvTextCache
    setMonitorPhase $monitor "検索"
    # 画面を開いたときに当たる（司令のスレッドを始める。lib.ps1 の読み込みは司令のスレッドで行われ、1 回目の検索の時間に含まれる）
    if ($mode -eq "service") { $service = newSearchService $lib $cache }
    for ($i = 1; $i -le $Count; $i++) {
        $r = if ($mode -eq "service") { invokeServiceSearch $service } else { invokeRunspaceSearch $cache }
        $row = [ordered]@{ N = $i; TotalMs = (roundMs $r.TotalMs); LoadMs = (roundMs $r.LoadMs); ListMs = (roundMs $r.ListMs)
            MatchMs = (roundMs $r.MatchMs); Packs = $r.Packs; Hits = $r.Hits; Truncated = $r.Truncated }
        $snap = getProcessSnapshot
        foreach ($k in $snap.Keys) { $row[$k] = $snap[$k] }
        $searches.Add($row)
        Write-Host ("  {0} {1}: {2:N0} ms（{3}{4} 件）" -f $Name, $i, $r.TotalMs, $r.Hits, $(if ($r.Truncated) { "+" } else { "" }))
    }
} finally {
    if ($service) { $service.Close() }
    $samples = stopResourceMonitor $monitor
}
$result = [ordered]@{
    Name = $Name; Word = $Word; Regex = $Regex; Mode = $mode; Searches = $searches.ToArray(); Samples = $samples
    PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
}
[System.IO.File]::WriteAllText($Out, ($result | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
