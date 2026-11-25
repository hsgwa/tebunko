# 前と後の 2 つの ref の検索の時間を、同じ機械で交互に測って比べる（perf.yml の入力 base_ref・rounds・scattered。手元でも使える）。
# 測るのは tools\measure_perf.ps1 で、ref ごと・データごとに起動する（このスクリプトは tebunko の lib を読まない。結果の result.json を読むだけ）。
#
#   .\tools\perf\measure_compare.ps1 -BaseTool <前の tebunko> -TargetTool <後の tebunko> -PerfdataIndex <TSV> -PerfdataWords <words.tsv>
#       [-ScatteredIndex <TSV> -ScatteredWords <語の表>] -Rounds 5 -Count 20 -Work <作業フォルダ> -Out <結果のフォルダ>
#   -BaseTool・-BaseRef     … 前の tebunko のフォルダと、見出しに出す ref の名前。省くと後だけを測る（比は出さない）
#   -TargetTool・-TargetRef … 後の tebunko のフォルダと、見出しに出す ref の名前
#   -PerfdataIndex・-PerfdataWords   … perfdata の TSV（tebunko-perfdata の new_index.ps1 で作る）と語の表。省くと測らない
#   -ScatteredIndex・-ScatteredWords … 散らばる語の TSV と語の表（tools\perf\new_scattered_data.ps1 で作る）。省くと測らない
#   -Rounds … 前 → 後 を交互に測る組の数。-Count … 語ごとに続けて検索する回数（measure_perf.ps1 の -Count）
#   -Work   … ref ごと・データごとの置き場 <Work>\<base|target>\<perfdata|scattered> を作る。元の TSV は写して使い、消さない
#   -Out    … 起動ごとの結果を <Out>\compare\<組>\<base|target>\<perfdata|scattered> に、比べた結果を compare.json・compare.md に書く
#   -MeasureScript … 起動する計測の道具（既定は tools\measure_perf.ps1。テストでは偽の道具を渡す）
#   -RunId・-Scale … 実行の情報（measure_perf.ps1 に渡す）
#
# 流れ:
#   1. ref ごと・データごとに、元の TSV の写しを置き場の本文インデックスのフォルダ（測る版の Workspace の IndexDir）の下に作る
#      （置き場の形を perf.yml の今の計測と同じ <W>\<IndexDir>\<データの名前> にし、gram 索引も作られるようにする）
#   2. 組ごとに、前の perfdata → 前の散らばる語 → 後の perfdata → 後の散らばる語 の順に measure_perf.ps1 を起動する。
#      1 組目は TSV を pack にするので、インデックス作成も測る。2 組目からは pack だけの置き場なので、検索だけを測る
#   3. 語ごと・ref ごとに、組ごとの TotalMs の中央値を並べ、その中央値を「その ref の値」にする。
#      揺れは (最大 − 最小) ÷ 中央値、比は 後 ÷ 前。件数（Hits）は組と ref をまたいで同じかを見る。GC は組ごとの値の中央値
param (
    [string]$MeasureScript,
    [string]$BaseTool,
    [string]$BaseRef = "",
    [Parameter(Mandatory = $true)][string]$TargetTool,
    [string]$TargetRef = "",
    [string]$PerfdataIndex,
    [string]$PerfdataWords,
    [string]$ScatteredIndex,
    [string]$ScatteredWords,
    [int]$Rounds = 5,
    [int]$Count = 20,
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$RunId = "",
    [string]$Scale = ""
)

$ErrorActionPreference = "Stop"
. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))
# Windows PowerShell 5.1 は -File で起動すると param の既定値の $PSScriptRoot が空になるため、既定の道具はここで決める
if (!$MeasureScript) { $MeasureScript = Join-Path (Split-Path $PSScriptRoot -Parent) "measure_perf.ps1" }
if ($Rounds -lt 1) { throw "-Rounds は 1 以上にしてください。" }
if (!$PerfdataIndex -and !$ScatteredIndex) { throw "-PerfdataIndex・-ScatteredIndex の少なくとも 1 つを指定してください。" }
if (($PerfdataIndex -and !$PerfdataWords) -or ($ScatteredIndex -and !$ScatteredWords)) { throw "TSV を指定したデータには、語の表（-PerfdataWords・-ScatteredWords）も指定してください。" }

function resolvePsPath([string]$path) { return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path).TrimEnd("\") }

function getShortSha([string]$tool) {
    # 見出しに出す短いコミット。git が使えなければ空
    $saved = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try { $sha = [string](& git -C $tool rev-parse HEAD 2>$null); if ($LASTEXITCODE -or $sha.Length -lt 7) { return "" }; return $sha.Substring(0, 7) } catch { return "" } finally { $ErrorActionPreference = $saved }
}

function newSides {
    # 測る ref の並び（前があれば前 → 後）
    $sides = New-Object System.Collections.Generic.List[object]
    if ($BaseTool) {
        $tool = (Resolve-Path -LiteralPath $BaseTool).ProviderPath
        $sides.Add(@{ Key = "base"; Tool = $tool; Ref = $BaseRef; Sha = (getShortSha $tool) })
    }
    $tool = (Resolve-Path -LiteralPath $TargetTool).ProviderPath
    $sides.Add(@{ Key = "target"; Tool = $tool; Ref = $TargetRef; Sha = (getShortSha $tool) })
    return , $sides.ToArray()
}

function newDataSets {
    # 測るデータの並び（perfdata → 散らばる語）
    $data = New-Object System.Collections.Generic.List[object]
    if ($PerfdataIndex) { $data.Add(@{ Key = "perfdata"; Source = (Resolve-Path -LiteralPath $PerfdataIndex).ProviderPath; Words = (Resolve-Path -LiteralPath $PerfdataWords).ProviderPath }) }
    if ($ScatteredIndex) { $data.Add(@{ Key = "scattered"; Source = (Resolve-Path -LiteralPath $ScatteredIndex).ProviderPath; Words = (Resolve-Path -LiteralPath $ScatteredWords).ProviderPath }) }
    return , $data.ToArray()
}

function copyDataSets($sides, $dataSets, [string]$work) {
    # ref ごと・データごとの置き場に TSV を写す。置き場（<Work>\<ref>\<データ>）と、その本文インデックスのフォルダ（Index）を返す
    $places = @{}
    foreach ($side in $sides) {
        foreach ($data in $dataSets) {
            $root = Join-Path $work "$($side.Key)\$($data.Key)"
            if ([System.IO.Directory]::Exists($root)) { [System.IO.Directory]::Delete($root, $true) }
            $index = Join-Path (getToolIndexDir $side.Tool $root) $data.Key
            [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($index))
            Copy-Item -LiteralPath $data.Source -Destination $index -Recurse
            $places["$($side.Key)|$($data.Key)"] = @{ Work = $root; Index = $index }
        }
    }
    return $places
}

function getInfoArgs($info) {
    # 実行の情報の引数。空の値は渡さない（powershell.exe -File は空の引数を落とし、名前だけが残って止まるため）
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($name in @("RunId", "Ref", "Sha", "Scale")) {
        if ($info[$name]) { $list.Add("-$name"); $list.Add([string]$info[$name]) }
    }
    return , $list.ToArray()
}

function invokeRounds($sides, $dataSets, $places) {
    # 組ごとに、ref ごと・データごとに measure_perf.ps1 を起動し、語ごとの値を集める（キーは「データ|語」、その下に ref ごとの組の並び）
    $records = [ordered]@{}
    for ($r = 1; $r -le $Rounds; $r++) {
        foreach ($side in $sides) {
            foreach ($data in $dataSets) {
                Write-Information "組 $r / $($Rounds): $($side.Key) の $($data.Key) を測ります。" -InformationAction Continue
                $place = $places["$($side.Key)|$($data.Key)"]
                $runOut = Join-Path $Out "compare\$r\$($side.Key)\$($data.Key)"
                $label = "$($side.Key) $($data.Key) round $r ($($side.Ref) $($side.Sha))"
                $runArgs = @("-Tool", $side.Tool, "-Index", $place.Index, "-Work", $place.Work, "-Words", $data.Words, "-Count", $Count, "-Out", $runOut, "-Label", $label) +
                    (getInfoArgs @{ RunId = $RunId; Ref = $side.Ref; Sha = $side.Sha; Scale = $Scale })
                invokeChildScript $MeasureScript $runArgs
                $result = [System.IO.File]::ReadAllText((Join-Path $runOut "result.json"), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
                addRoundRecords $records $data.Key $side.Key @($result.Search)
            }
        }
    }
    return $records
}

function addRoundRecords($records, [string]$data, [string]$side, $search) {
    # 1 回の起動の結果（result.json の Search）から、語ごとの TotalMs の中央値・件数・GC を足す
    foreach ($s in $search) {
        $key = "$data|$($s.Name)"
        if (!$records.Contains($key)) { $records[$key] = @{ Data = $data; Name = [string]$s.Name; base = (New-Object System.Collections.Generic.List[object]); target = (New-Object System.Collections.Generic.List[object]) } }
        $records[$key][$side].Add(@{ Ms = [double]$s.TotalMs.Median; Hits = $s.Hits; Gc0 = $s.Gc0; Gc1 = $s.Gc1; Gc2 = $s.Gc2 })
    }
}

function getMedianOf($rows, [string]$key) {
    $values = [double[]]@($rows | ForEach-Object { $_[$key] } | Where-Object { $null -ne $_ })
    $stats = getStats $values
    if ($null -eq $stats) { return $null }
    return $stats.Median
}

function getSideSummary($rows) {
    # ref 1 つの組の並びから、その ref の値（中央値）・揺れ・件数・GC を返す。組が無ければ $null
    if ($rows.Count -eq 0) { return $null }
    $ms = [double[]]@($rows | ForEach-Object { $_.Ms })
    $stats = getStats $ms
    $spread = $(if ($stats.Median -gt 0) { [Math]::Round(($stats.Max - $stats.Min) / $stats.Median, 3) } else { $null })
    return @{
        Ms = $stats.Median; Spread = $spread; Rounds = $ms; Hits = @($rows | ForEach-Object { $_.Hits })
        Gc = [ordered]@{ Gc0 = (getMedianOf $rows "Gc0"); Gc1 = (getMedianOf $rows "Gc1"); Gc2 = (getMedianOf $rows "Gc2") }
    }
}

function newWordResults($records) {
    # 語ごとの比べた結果（compare.json の Words）
    $words = New-Object System.Collections.Generic.List[object]
    foreach ($rec in $records.Values) {
        $base = getSideSummary $rec.base
        $target = getSideSummary $rec.target
        $allHits = @(@($(if ($base) { $base.Hits }) + $target.Hits) | ForEach-Object { [string]$_ } | Select-Object -Unique)
        $words.Add([ordered]@{
            Data = $rec.Data; Name = $rec.Name
            BaseMs = $(if ($base) { $base.Ms }); TargetMs = $target.Ms
            Ratio = $(if ($base -and $base.Ms -gt 0) { [Math]::Round($target.Ms / $base.Ms, 3) })
            BaseSpread = $(if ($base) { $base.Spread }); TargetSpread = $target.Spread
            BaseRounds = $(if ($base) { $base.Rounds }); TargetRounds = $target.Rounds
            BaseHits = $(if ($base) { $base.Hits[0] }); TargetHits = $target.Hits[0]; HitsMatch = ($allHits.Count -eq 1)
            BaseGc = $(if ($base) { $base.Gc }); TargetGc = $target.Gc
        })
    }
    return , $words.ToArray()
}

function formatMs($value) { if ($null -eq $value) { return "–" }; return ("{0:N0} ms" -f $value) }
function formatPlain($value) { if ($null -eq $value) { return "–" }; return [string]$value }
function formatGc($gc) { if ($null -eq $gc) { return "–" }; return "$(formatPlain $gc.Gc0)/$(formatPlain $gc.Gc1)/$(formatPlain $gc.Gc2)" }
function formatRounds($values) { if ($null -eq $values) { return "–" }; return (@($values | ForEach-Object { "{0:N0}" -f $_ }) -join "・") }

function newCompareMarkdown($sides, $words) {
    # compare.md（Summary に足す節）
    $base = @($sides | Where-Object { $_.Key -eq "base" })
    $target = @($sides | Where-Object { $_.Key -eq "target" })[0]
    $md = New-Object System.Collections.Generic.List[string]
    if ($base.Count) {
        $md.Add("## 前と後の比べ（前: $($base[0].Ref) ($($base[0].Sha))・後: $($target.Ref) ($($target.Sha))）")
        $md.Add("")
        $md.Add("同じジョブの中で、前 → 後 の順に $Rounds 組、交互に測った（データごと・ref ごとに measure_perf.ps1 を起動。1 組の中の検索は $Count 回）。値は、組ごとの「数え終わるまで」（TotalMs）の中央値を $Rounds 個並べた、その中央値。揺れは (最大 − 最小) ÷ 中央値、比は 後 ÷ 前。1 組目はインデックス作成のすぐ後に測る。GC 回数は検索の間の 0/1/2 世代の、組ごとの値の中央値。")
    } else {
        $md.Add("## 散らばる語の検索（$($target.Ref) ($($target.Sha))）")
        $md.Add("")
        $md.Add("場所ごとに 1 件ずつ散らばる語のデータを、ほかの計測と別の置き場に作って測った（measure_perf.ps1 を $Rounds 回起動。1 回の中の検索は $Count 回）。値は「数え終わるまで」（TotalMs）の中央値。GC 回数は検索の間の 0/1/2 世代。")
    }
    $md.Add("")
    $md.Add("| データ | 語 | 前 | 後 | 比 | 前の揺れ | 前の組ごと | 後の組ごと | 件数（前／後） | GC（前／後） |")
    $md.Add("|---|---|---|---|---|---|---|---|---|---|")
    foreach ($w in $words) {
        $hits = "$(formatPlain $w.BaseHits)／$(formatPlain $w.TargetHits)$(if (!$w.HitsMatch) { '（違う）' })"
        $md.Add("| $($w.Data) | $($w.Name) | $(formatMs $w.BaseMs) | $(formatMs $w.TargetMs) | $(formatPlain $w.Ratio) | $(formatPlain $w.BaseSpread) | $(formatRounds $w.BaseRounds) | $(formatRounds $w.TargetRounds) | $hits | $(formatGc $w.BaseGc)／$(formatGc $w.TargetGc) |")
    }
    return (($md -join "`n") + "`n")
}

$Work = resolvePsPath $Work
$Out = resolvePsPath $Out
[void][System.IO.Directory]::CreateDirectory($Work)
[void][System.IO.Directory]::CreateDirectory($Out)
$sides = newSides
$dataSets = newDataSets
$places = copyDataSets $sides $dataSets $Work
$records = invokeRounds $sides $dataSets $places
$words = newWordResults $records
$compare = [ordered]@{
    Rounds = $Rounds; Count = $Count; RunId = $RunId; Scale = $Scale; Scattered = [bool]$ScatteredIndex
    Base = $(if ($BaseTool) { [ordered]@{ Ref = $sides[0].Ref; Sha = $sides[0].Sha } }); Target = [ordered]@{ Ref = $sides[-1].Ref; Sha = $sides[-1].Sha }
    Words = $words
}
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $Out "compare.json"), ($compare | ConvertTo-Json -Depth 6), $utf8)
[System.IO.File]::WriteAllText((Join-Path $Out "compare.md"), (newCompareMarkdown $sides $words), $utf8)
Write-Information "前と後の比べを書きました（語 $($words.Count)・組 $Rounds）。" -InformationAction Continue
