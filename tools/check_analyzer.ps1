# PSScriptAnalyzer を scripts\・tools\・tests\ にかけ、ベースラインより増えた指摘があれば失敗にする。
#
#   .\tools\check_analyzer.ps1                  検査する（CI の test が呼ぶ）
#   .\tools\check_analyzer.ps1 -UpdateBaseline  ベースライン（tests\meta\PSScriptAnalyzer.baseline.tsv）を今の指摘で書き直す
#
# 設定は scripts\・tools\ が tests\meta\PSScriptAnalyzer.psd1、tests\ が tests\meta\PSScriptAnalyzer.tests.psd1。
# 除く規則とその理由は設定ファイルのコメントと docs\safety\scans.md にある。残った指摘は、ファイル・規則ごとの件数でベースラインに記録する。
# 件数が増えたときも減ったときも失敗にする（減ったまま通すと、減った分の余りに新しい指摘が隠れるため）。
# 件数は PowerShell の版で変わりうるので、Windows PowerShell 5.1 でだけ動かす。
# -UpdateBaseline で直したベースラインは、人が中身を見て PR に入れる。
# PSScriptAnalyzer の版は環境変数 PSSA_VERSION があればその版を読み込む（CI が版を固定している）。
param (
    [switch]$UpdateBaseline,
    [string]$Root = (Join-Path $PSScriptRoot "..")
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\analyzer_baseline.ps1"

if ($PSVersionTable.PSEdition -ne "Desktop") {
    Write-Host "この検査は Windows PowerShell 5.1 でだけ動かします（指摘の件数が PowerShell の版で変わり、ベースラインとずれるため）。powershell.exe で実行してください。" -ForegroundColor Red
    exit 1
}

$rootDir = [System.IO.Path]::GetFullPath($Root).TrimEnd("\")
$baselinePath = Join-Path $rootDir "tests\meta\PSScriptAnalyzer.baseline.tsv"
$utf8 = New-Object System.Text.UTF8Encoding($false)

if ($env:PSSA_VERSION) {
    Import-Module PSScriptAnalyzer -RequiredVersion $env:PSSA_VERSION
} else {
    Import-Module PSScriptAnalyzer
}

$targets = @(
    @{ Path = "scripts"; Settings = "tests\meta\PSScriptAnalyzer.psd1" }
    @{ Path = "tools"; Settings = "tests\meta\PSScriptAnalyzer.psd1" }
    @{ Path = "tests"; Settings = "tests\meta\PSScriptAnalyzer.tests.psd1" }
)
$findings = New-Object System.Collections.Generic.List[object]
foreach ($t in $targets) {
    foreach ($f in @(Invoke-ScriptAnalyzer -Path (Join-Path $rootDir $t.Path) -Recurse -Settings (Join-Path $rootDir $t.Settings))) {
        # TypeNotFound は規則でなく構文解析の指摘なので ExcludeRules では除けない。
        # 継承元の型が別ファイルにあるため 1 ファイルだけでは解決できないもの（読み込む順で解決することは tests\meta\structure.Tests.ps1 の「型の読み込み」が確かめる）
        if ($f.RuleName -eq "TypeNotFound") { continue }
        # tests\testdata\text\ の .ps1・.psm1・.psd1 は、テキスト抽出の見本（中身は文章）。チェックアウトの改行などで指摘が変わるため調べない
        if ($f.ScriptPath.Replace("/", "\").ToLowerInvariant().Contains("\tests\testdata\text\")) { continue }
        $findings.Add($f)
    }
}
$counts = getAnalyzerCounts $findings.ToArray() $rootDir

if ($UpdateBaseline) {
    [System.IO.File]::WriteAllText($baselinePath, (formatAnalyzerBaseline $counts), $utf8)
    Write-Host "ベースラインを書き直しました（$($counts.Count) 行）: $baselinePath"
    exit 0
}

$baseline = if (Test-Path -LiteralPath $baselinePath) { readAnalyzerBaseline ([System.IO.File]::ReadAllText($baselinePath, $utf8)) } else { @{} }
$result = compareAnalyzerBaseline $baseline $counts

if ($result.Increased.Count -gt 0) {
    Write-Host "PSScriptAnalyzer の指摘がベースラインより増えています。直すか、理由があれば設定で除いてください。" -ForegroundColor Red
    $result.Increased | ForEach-Object { Write-Host "  $_" }
}
if ($result.Decreased.Count -gt 0) {
    Write-Host "PSScriptAnalyzer の指摘がベースラインより減っています。ベースラインを実際に合わせてください。" -ForegroundColor Red
    $result.Decreased | ForEach-Object { Write-Host "  $_" }
}
if ($result.Increased.Count -gt 0 -or $result.Decreased.Count -gt 0) { exit 1 }
Write-Host "PSScriptAnalyzer の検査: ベースラインどおり（指摘 $($findings.Count) 件）"
exit 0
