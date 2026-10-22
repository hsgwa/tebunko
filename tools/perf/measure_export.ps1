# 検索結果をファイル（.xlsx・.csv・.tsv）に書き出す部分（invokeSearchExport。検索の司令のスレッドが実行する書き手）の速さを測る。
# tools\measure_perf.ps1 が別のプロセスで起動する。測る元は、その場で作る合成のヒット -Hits 件（Excel・Word・複数行・タブを含むセルを混ぜる）。
# 表に出ている分だけを書く（上限を超えた溢れは無い）。形式ごとに -Repeat 回続けて書き、1 回ごとの時間を出す（比べるのは中央値）。
# 結果（時間・書いた件数・ファイルの大きさ・リソースの記録）を -Out の JSON に書く。件数は、測る版どうしで同じ結果を書いたかの確かめに使う。
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Repeat = 3,
    [int]$Hits = 10000,
    [int]$SampleMs = 200
)

$ErrorActionPreference = "Stop"
. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))
$lib = resolveTebunkoLib $Tool
. ([WildcardPattern]::Escape($lib.Replace('`', '``')))
assertTebunkoFunctions @("newExportRequest", "invokeSearchExport")

function newExportPerfHits([int]$count) {
    # 合成のヒット。種類を回して、Excel の表・Word の段落・複数行のセル・タブ区切りのセルを混ぜる
    $list = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $count; $i++) {
        switch ($i % 4) {
            0 { $hit = @{ RelDir = "営業部\見積"; Book = "見積$($i % 50).xlsx"; Location = "Sheet1!B$i"; LineNumber = $i; Line = "商品名`t数量`t金額 $i" } }
            1 { $hit = @{ RelDir = "総務"; Book = "規程$($i % 30).docx"; Location = "本文"; LineNumber = $i; Line = "第 $i 条 検索の対象となる文字列を含む段落です。" } }
            2 { $hit = @{ RelDir = "総務"; Book = "議事録$($i % 30).docx"; Location = "表 1"; LineNumber = $i; Line = "1 行目`n2 行目 $i`r`n`"引用符`"を含む,コンマ" } }
            default { $hit = @{ RelDir = ""; Book = "メモ$($i % 20).txt"; Location = ""; LineNumber = $i; Line = ("あいうえお" * 20) + " $i" } }
        }
        $list.Add($hit)
    }
    return $list.ToArray()
}

$dir = Join-Path $Work "export"
[void][System.IO.Directory]::CreateDirectory($dir)
$shown = newExportPerfHits $Hits
$source = @{ WorkDir = (Join-Path $dir "work"); Truncated = $false }
[void][System.IO.Directory]::CreateDirectory($source.WorkDir)

$monitor = startResourceMonitor $SampleMs
$formats = New-Object System.Collections.Generic.List[object]
try {
    foreach ($format in @("xlsx", "csv", "tsv")) {
        setMonitorPhase $monitor $format
        $path = Join-Path $dir "result.$format"
        $runs = New-Object System.Collections.Generic.List[object]
        for ($i = 1; $i -le $Repeat; $i++) {
            $request = newExportRequest $source $path $shown "" "測る語" $format $true
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            invokeSearchExport $request
            $ms = $watch.Elapsed.TotalMilliseconds
            Write-Information ("  {0} {1} 回目: {2:N0} ms（{3}）" -f $format, $i, $ms, $request.Outcome) -InformationAction Continue
            $runs.Add([ordered]@{ Ms = [Math]::Round($ms, 1); Outcome = [string]$request.Outcome; Written = [int]$request.Written })
        }
        $sizeMB = 0
        if ([System.IO.File]::Exists($path)) { $sizeMB = [Math]::Round(([System.IO.FileInfo]$path).Length / 1MB, 2) }
        $formats.Add([ordered]@{ Format = $format; SizeMB = $sizeMB; Runs = $runs.ToArray() })
    }
} finally {
    $samples = stopResourceMonitor $monitor
}
$result = [ordered]@{ Repeat = $Repeat; Hits = $Hits; Formats = $formats.ToArray() }
$result.Samples = $samples
$result.PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
[System.IO.File]::WriteAllText($Out, ($result | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
