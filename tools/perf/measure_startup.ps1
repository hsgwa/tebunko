# 起動と終了の速さを、zip 版（scripts\ をそのまま使う形）と単一 .ps1 版で交互に測って比べる。
# 測るのは「起動して主画面が出る（操作できる）まで」と「閉じて終わるまで」だけ。取り込みと検索は測らない。
# 設計は docs/design/structure/single-script.md「速さの測り方」。起動は .github/workflows/perf.yml（入力 startup）。
#
#   .\tools\perf\measure_startup.ps1 -Work <作業フォルダ> -Out <結果の出力先> [-Runs 5]
#
# 毎回新しい powershell.exe（-NoProfile -STA -ExecutionPolicy RemoteSigned -File <入口>）を起こし、
# Stopwatch を起こす直前に始める。「主画面が出た」は UI オートメーションで、そのプロセスの窓の中に
# AutomationId "Tabs"（主画面のタブ）が現れたときとする（起動中の表示の窓は数えない）。
# 閉じるのは WindowPattern.Close()。終了コードが 0 でない回があっても値は捨てず、回数を結果に出す。
# 先頭の 1 回（ウォームアップ）は中央値に入れない。zip 版と単一 .ps1 版は別々の空のフォルダに置く。
param (
    [Parameter(Mandatory = $true)]
    [string]$Work,
    [Parameter(Mandatory = $true)]
    [string]$Out,
    [int]$Runs = 5,
    [int]$StartTimeoutSeconds = 120,
    [int]$CloseTimeoutSeconds = 60
)

$ErrorActionPreference = "Stop"

$rootDir = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
. "$PSScriptRoot\startup_common.ps1"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

# ---- 準備: zip 版と単一 .ps1 版を、別々の空のフォルダに置く ----
$zipDir = Join-Path $Work "zip"
$singleDir = Join-Path $Work "single"
foreach ($dir in @($zipDir, $singleDir, $Out)) {
    if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force }
    [void](New-Item -ItemType Directory -Force -Path $dir)
}
Copy-Item -LiteralPath (Join-Path $rootDir "scripts") -Destination (Join-Path $zipDir "scripts") -Recurse
$singleEntry = & (Join-Path $rootDir "tools\new_single_script.ps1") -Version "v0.0.0-startup" -OutFile (Join-Path $singleDir "tebunko.ps1")
$singleEntry = $singleEntry | Select-Object -Last 1
$zipEntry = Join-Path $zipDir "scripts\tebunko\gui.ps1"

# 1 回測る。起動から主画面が出るまで・閉じてから終わるまで（ミリ秒）と終了コードを返す
function measureOnce {
    param (
        [string]$entry
    )

    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $p = Start-Process powershell.exe -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "RemoteSigned", "-File", $entry) -PassThru -WindowStyle Hidden
    $null = $p.Handle   # ExitCode を取るため、起動の直後にハンドルを持つ
    $window = $null
    try {
        $idCondition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "Tabs")
        $pidCondition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $p.Id)
        while ($null -eq $window -and $watch.Elapsed.TotalSeconds -lt $StartTimeoutSeconds -and !$p.HasExited) {
            foreach ($w in @([Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children, $pidCondition))) {
                if ($w.FindFirst([Windows.Automation.TreeScope]::Descendants, $idCondition)) { $window = $w; break }
            }
            if ($null -eq $window) { [System.Threading.Thread]::Sleep(20) }
        }
        if ($null -eq $window) {
            throw "主画面が出ませんでした（$entry。終了済み: $($p.HasExited)）"
        }
        $startMs = $watch.Elapsed.TotalMilliseconds

        $watch.Restart()
        $window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close()
        $exited = $p.WaitForExit($CloseTimeoutSeconds * 1000)
        $closeMs = $watch.Elapsed.TotalMilliseconds
        return [ordered]@{ StartMs = [Math]::Round($startMs, 1); CloseMs = [Math]::Round($closeMs, 1); ExitCode = $(if ($exited) { $p.ExitCode } else { $null }) }
    } finally {
        if (!$p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
        $p.Dispose()
    }
}

# ---- 測る: zip 版と単一 .ps1 版を交互に。先頭の 1 組はウォームアップ ----
$results = [ordered]@{ zip = @(); single = @() }
for ($i = 0; $i -le $Runs; $i++) {
    foreach ($kind in @("zip", "single")) {
        $entry = if ($kind -eq "zip") { $zipEntry } else { $singleEntry }
        $r = measureOnce $entry
        $r.Run = $i
        $results[$kind] += $r
        Write-Host ("{0} 回目 {1}: 起動 {2} ms・終了 {3} ms・終了コード {4}{5}" -f $i, $kind, $r.StartMs, $r.CloseMs, $r.ExitCode, $(if ($i -eq 0) { "（ウォームアップ）" } else { "" }))
    }
}

# ---- 比べる ----
$start = compareStartup ([double[]]@($results.zip | ForEach-Object { $_.StartMs })) ([double[]]@($results.single | ForEach-Object { $_.StartMs }))
$close = compareStartup ([double[]]@($results.zip | ForEach-Object { $_.CloseMs })) ([double[]]@($results.single | ForEach-Object { $_.CloseMs }))
$nonZero = [ordered]@{
    zip    = countNonZeroExit @($results.zip | ForEach-Object { $_.ExitCode })
    single = countNonZeroExit @($results.single | ForEach-Object { $_.ExitCode })
}

$json = [ordered]@{ Runs = $Runs; Warmup = 1; Start = $start; Close = $close; NonZeroExit = $nonZero; Results = $results }
[System.IO.File]::WriteAllText((Join-Path $Out "startup.json"), ($json | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))

function formatVerdict {
    param ($c)
    if ($null -eq $c.Pass) { return "判定できない" }
    if ($c.Pass) { return "1.2 倍以内" } else { return "1.2 倍を超えた" }
}
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("## 起動と終了の速さ（zip 版と単一 .ps1 版。各 $Runs 回・ウォームアップ 1 回を捨てた中央値）")
$lines.Add("")
$lines.Add("| | zip 版の中央値 (ms) | 単一 .ps1 版の中央値 (ms) | 倍率 | 判定 |")
$lines.Add("|---|---|---|---|---|")
foreach ($row in @(@("起動（主画面が出るまで）", $start), @("終了（閉じてから終わるまで）", $close))) {
    $c = $row[1]
    $ratioText = if ($null -ne $c.Ratio) { $c.Ratio.ToString("0.00", [System.Globalization.CultureInfo]::InvariantCulture) } else { "-" }
    $lines.Add("| $($row[0]) | $($c.ZipMedian) | $($c.SingleMedian) | $ratioText | $(formatVerdict $c) |")
}
$lines.Add("")
$lines.Add("終了コードが 0 でなかった回数（ウォームアップを含む全 $($Runs + 1) 回のうち）: zip 版 $($nonZero.zip)・単一 .ps1 版 $($nonZero.single)")
$lines.Add("")
$lines.Add("| 回 | 版 | 起動 (ms) | 終了 (ms) | 終了コード |")
$lines.Add("|---|---|---|---|---|")
foreach ($kind in @("zip", "single")) {
    foreach ($r in $results[$kind]) { $lines.Add("| $($r.Run) | $kind | $($r.StartMs) | $($r.CloseMs) | $($r.ExitCode) |") }
}
$summary = ($lines -join "`r`n") + "`r`n"
[System.IO.File]::WriteAllText((Join-Path $Out "summary.md"), $summary, (New-Object System.Text.UTF8Encoding($false)))
Write-Host $summary
