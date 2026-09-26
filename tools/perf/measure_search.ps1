# 1 つの語で、同じプロセスのまま -Count 回続けて検索し、1 回ごとの時間とメモリを測る。tools\measure_perf.ps1 が語ごとに別のプロセスで起動する。
# 1 回の検索は画面の検索スレッドと同じ流れ（新しい Runspace で lib.ps1 を読み込む → 集約ファイルを列挙する → 照合する。上限 1 万件）。
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

$searchScript = {
    param ($lib, $index, $word, $simple, $cache)
    $total = [System.Diagnostics.Stopwatch]::StartNew(); $step = [System.Diagnostics.Stopwatch]::StartNew()
    . $lib
    $load = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $found = getIndexPackFiles @($index)
    $list = $step.Elapsed.TotalMilliseconds; $step.Restart()
    $r = searchPackIndex $word $found.Packs $simple 10000 -cache $cache
    @{ LoadMs = $load; ListMs = $list; MatchMs = $step.Elapsed.TotalMilliseconds; TotalMs = $total.Elapsed.TotalMilliseconds; Packs = $found.Packs.Count; Hits = $r.Hits.Count; Truncated = [bool]$r.Truncated }
}

$monitor = startResourceMonitor $SampleMs
$searches = New-Object System.Collections.Generic.List[object]
try {
    $cache = newTsvTextCache
    setMonitorPhase $monitor "検索"
    for ($i = 1; $i -le $Count; $i++) {
        $ps = [powershell]::Create()
        try {
            [void]$ps.AddScript($searchScript.ToString()).AddArgument($lib).AddArgument($Index).AddArgument($Word).AddArgument(!$Regex).AddArgument($cache)
            $r = $ps.Invoke()[0]
        } finally {
            $ps.Dispose()
        }
        $row = [ordered]@{ N = $i; TotalMs = [Math]::Round($r.TotalMs, 1); LoadMs = [Math]::Round($r.LoadMs, 1); ListMs = [Math]::Round($r.ListMs, 1)
            MatchMs = [Math]::Round($r.MatchMs, 1); Packs = $r.Packs; Hits = $r.Hits; Truncated = $r.Truncated }
        $snap = getProcessSnapshot
        foreach ($k in $snap.Keys) { $row[$k] = $snap[$k] }
        $searches.Add($row)
        Write-Host ("  {0} {1}: {2:N0} ms（{3}{4} 件）" -f $Name, $i, $r.TotalMs, $r.Hits, $(if ($r.Truncated) { "+" } else { "" }))
    }
} finally {
    $samples = stopResourceMonitor $monitor
}
$result = [ordered]@{
    Name = $Name; Word = $Word; Regex = [bool]$Regex; Searches = $searches.ToArray(); Samples = $samples
    PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
}
[System.IO.File]::WriteAllText($Out, ($result | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
