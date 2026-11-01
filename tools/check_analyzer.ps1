# PSScriptAnalyzer を scripts\・tools\・tests\ にかけ、ベースラインより増えた指摘があれば失敗にする。
#
#   .\tools\check_analyzer.ps1                  検査する（CI の test が呼ぶ）
#   .\tools\check_analyzer.ps1 -UpdateBaseline  ベースライン（tests\meta\PSScriptAnalyzer.baseline.tsv）を今の指摘で書き直す
#   .\tools\check_analyzer.ps1 -Path a.ps1,b.ps1  指定したファイル（リポジトリからの相対パス）だけを検査する（コミット前のフックが呼ぶ）
#       そのファイルの（ファイル・規則）の件数だけをベースラインと比べ、増えても減っても失敗にする。
#       -UpdateBaseline とは一緒に使えない（一部のファイルだけで書き直すと、ほかのファイルの行が消えるため。終了コード 2）
#
# 設定は scripts\・tools\ が tests\meta\PSScriptAnalyzer.psd1、tests\ が tests\meta\PSScriptAnalyzer.tests.psd1。
# 除く規則とその理由は設定ファイルのコメントと docs\safety\scans.md にある。残った指摘は、ファイル・規則ごとの件数でベースラインに記録する。
# 件数が増えたときも減ったときも失敗にする（減ったまま通すと、減った分の余りに新しい指摘が隠れるため）。
# 件数は PowerShell の版で変わりうるので、Windows PowerShell 5.1 でだけ動かす。
# -UpdateBaseline で直したベースラインは、人が中身を見て PR に入れる。
# PSScriptAnalyzer の版は環境変数 PSSA_VERSION があればその版を読み込む（CI が版を固定している）。
param (
    [switch]$UpdateBaseline,
    [string[]]$Path,   # 検査するファイル（既定は scripts・tools・tests の全部）
    [string]$Root = (Join-Path $PSScriptRoot "..")
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\analyzer_baseline.ps1"

# powershell.exe -File で呼ぶと、-Path a,b は配列にならず 1 つの文字列で渡るため、カンマで分ける
$onlyFiles = @($Path | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$limited = $onlyFiles.Count -gt 0
if ($onlyFiles.Count -gt 0 -and $UpdateBaseline) {
    Write-Warning "-Path と -UpdateBaseline は一緒には使えません（一部のファイルだけでベースラインを書き直すと、ほかのファイルの行が消えるため）。"
    exit 2
}

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
if ($limited) {
    # 作業ツリーにある .ps1 だけを、置き場所に合う設定でかける（消した・名前を変えて無いファイルは飛ばす）
    $onlyFiles = @(selectAnalyzerFiles $onlyFiles $rootDir)
    foreach ($rel in $onlyFiles) {
        $t = $targets | Where-Object { $rel -like "$($_.Path)/*" } | Select-Object -First 1
        foreach ($f in @(Invoke-ScriptAnalyzer -Path (Join-Path $rootDir $rel) -Settings (Join-Path $rootDir $t.Settings))) {
            $findings.Add($f)
        }
    }
} else {
    foreach ($t in $targets) {
        foreach ($f in @(Invoke-ScriptAnalyzer -Path (Join-Path $rootDir $t.Path) -Recurse -Settings (Join-Path $rootDir $t.Settings))) {
            $findings.Add($f)
        }
    }
}
$findings = @(selectAnalyzerFindings $findings.ToArray())
$counts = getAnalyzerCounts $findings $rootDir

if ($UpdateBaseline) {
    [System.IO.File]::WriteAllText($baselinePath, (formatAnalyzerBaseline $counts), $utf8)
    Write-Host "ベースラインを書き直しました（$($counts.Count) 行）: $baselinePath"
    exit 0
}

$baseline = if (Test-Path -LiteralPath $baselinePath) { readAnalyzerBaseline ([System.IO.File]::ReadAllText($baselinePath, $utf8)) } else { @{} }
if ($limited) { $baseline = filterAnalyzerBaseline $baseline $onlyFiles }
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
