# PDF の読み取り（readPdfFile）の速さを、版（コミット）を並べて同じ機械で測る。結果は表（summary.md）と result.json に書く。
#
#   .\tools\perf\measure_pdf.ps1 -Work <作業フォルダ> -Base origin/worktree-pdf
#   -Base     … 基準の版（ブランチ・コミット）。ほかの版と同じ PDF・同じ機械で測り、比べる
#   -Versions … 測る版（コミット）を測る順に。省略すると、Base の次から HEAD までのコミットのうち、pdf_file.ps1 を変えたものを古い順に全部測る
#   -Scale    … 測る PDF の量（new_pdf_data.ps1。既定 1）
#   -Repeat   … 版ごとに readPdfFile を測る回数（既定 3）
#   -Pdf      … 測る PDF。省略すると new_pdf_data.ps1 で作る
#   -Out      … 結果を書くフォルダ。既定は <Work>\result
#
# 版ごとに、そのコミットの scripts\shared\core を作業フォルダに取り出し、別のプロセス（measure_pdf_version.ps1）で読み込んで測る
# （版どうしの定義が混ざらない）。基準と最後の版は、関数ごとの時間も 1 回測る（関数を包んで数えるので、そのぶん遅くなる。順位をつけるための値）。
# 読み取った文字の SHA-256 が基準と違う版は、表に印を付ける（結果が変わっている）。
param (
    [Parameter(Mandatory = $true)][string]$Work,
    [string]$Repo = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$Base = "origin/worktree-pdf",
    [string[]]$Versions,
    [double]$Scale = 1.0,
    [int]$Repeat = 3,
    [string]$Pdf,
    [string]$Out
)

$ErrorActionPreference = "Stop"
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object System.Text.UTF8Encoding($false)
$Repo = (Resolve-Path -LiteralPath $Repo).ProviderPath.TrimEnd("\")
[void][System.IO.Directory]::CreateDirectory($Work)
$Work = (Resolve-Path -LiteralPath $Work).ProviderPath.TrimEnd("\")
if (!$Out) { $Out = Join-Path $Work "result" }
[void][System.IO.Directory]::CreateDirectory($Out)

function invokeGit([string[]]$arguments) {
    $output = & git -C $Repo @arguments
    if ($LASTEXITCODE) { throw "git $($arguments -join ' ') が失敗しました（終了コード $LASTEXITCODE）。" }
    return $output
}

function invokeChild([string]$script, [string[]]$arguments) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot $script) @arguments
    if ($LASTEXITCODE) { throw "$script が失敗しました（終了コード $LASTEXITCODE）。" }
}

# 測る版の一覧（基準が先頭）
$baseSha = [string](invokeGit @("rev-parse", "$Base^{commit}"))
$refs = New-Object System.Collections.Generic.List[string]
$refs.Add($baseSha)
if ($Versions) {
    foreach ($v in $Versions) { $refs.Add([string](invokeGit @("rev-parse", "$v^{commit}"))) }
} else {
    foreach ($line in @(invokeGit @("rev-list", "--reverse", "$baseSha..HEAD", "--", "scripts/shared/core/pdf_file.ps1"))) {
        if ($line) { $refs.Add([string]$line) }
    }
}

# 測る PDF
if (!$Pdf) {
    $Pdf = Join-Path $Work "perf.pdf"
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    invokeChild "new_pdf_data.ps1" @("-Dest", $Pdf, "-Scale", $Scale.ToString($inv))
    Write-Host ("PDF を作りました（{0:N1} 秒）" -f $watch.Elapsed.TotalSeconds)
}
$Pdf = (Resolve-Path -LiteralPath $Pdf).ProviderPath

$results = New-Object System.Collections.Generic.List[hashtable]
for ($i = 0; $i -lt $refs.Count; $i++) {
    $sha = $refs[$i]
    $short = $sha.Substring(0, 7)
    $subject = [string](invokeGit @("log", "-1", "--format=%s", $sha))
    $dir = Join-Path $Work "version\$short"
    if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force }
    [void][System.IO.Directory]::CreateDirectory($dir)
    $zip = Join-Path $dir "scripts.zip"
    [void](invokeGit @("archive", "--format=zip", "-o", $zip, $sha, "scripts/shared/core"))
    Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $dir "tree") -Force
    $json = Join-Path $dir "result.json"
    $arguments = @("-Scripts", (Join-Path $dir "tree\scripts"), "-Pdf", $Pdf, "-Repeat", $Repeat, "-Out", $json)
    if ($i -eq 0 -or $i -eq $refs.Count - 1) { $arguments += "-Profile" }
    Write-Host "測っています: $short"
    invokeChild "measure_pdf_version.ps1" $arguments
    $data = [System.IO.File]::ReadAllText($json, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    $sorted = @($data.Seconds | Sort-Object)
    $median = $sorted[[int][Math]::Floor($sorted.Count / 2)]
    $results.Add(@{ Sha = $short; Subject = $subject; Median = [double]$median; Min = [double]$sorted[0]; Hash = [string]$data.Hash; Profile = @($data.Profile); ProfileTotal = [double]$data.ProfileTotal; Bytes = [long]$data.Bytes; Pages = [int]$data.Pages })
}

# 表
function formatSeconds([double]$v) { return $v.ToString("0.00", $inv) }
function formatRatio([double]$v) { return $v.ToString("0.00", $inv) }
$baseResult = $results[0]
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("## PDF の読み取りの速さ")
[void]$sb.AppendLine()
[void]$sb.AppendLine(("測った PDF: {0:N2} MB・{1} ページ・各版 {2} 回（中央値と最小）" -f ($baseResult.Bytes / 1MB), $baseResult.Pages, $Repeat))
[void]$sb.AppendLine()
[void]$sb.AppendLine("| 版 | 内容 | 中央値（秒） | 最小（秒） | 基準との比 | 直前との比 | 結果 |")
[void]$sb.AppendLine("|---|---|---:|---:|---:|---:|---|")
for ($i = 0; $i -lt $results.Count; $i++) {
    $r = $results[$i]
    $label = if ($i -eq 0) { "基準" } else { "+$i" }
    $ratio = if ($i -eq 0) { "1.00" } else { formatRatio ($r.Median / $baseResult.Median) }
    $prev = if ($i -eq 0) { "-" } else { formatRatio ($r.Median / $results[$i - 1].Median) }
    $same = if ($r.Hash -eq $baseResult.Hash) { "基準と同じ" } else { "**基準と違う**" }
    $subject = $r.Subject.Replace("|", "/")
    [void]$sb.AppendLine("| $label $($r.Sha) | $subject | $(formatSeconds $r.Median) | $(formatSeconds $r.Min) | $ratio | $prev | $same |")
}
foreach ($index in @(0, ($results.Count - 1) | Select-Object -Unique)) {
    $r = $results[$index]
    $title = if ($index -eq 0) { "基準" } else { "最後の版" }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("### 関数ごとの時間（$title $($r.Sha)。包んで数えるため全体より遅い。入れ子の呼び出しは外側だけ数える）")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("| 関数 | 回数 | 時間（秒） | 全体に対する割合 |")
    [void]$sb.AppendLine("|---|---:|---:|---:|")
    foreach ($row in @($r.Profile | Sort-Object { - [double]$_.Seconds })) {
        $share = if ($r.ProfileTotal -gt 0) { ([double]$row.Seconds / $r.ProfileTotal).ToString("P0", $inv) } else { "-" }
        [void]$sb.AppendLine("| $($row.Name) | $($row.Calls) | $(formatSeconds ([double]$row.Seconds)) | $share |")
    }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("全体（包んだとき）: $(formatSeconds $r.ProfileTotal) 秒")
}
[System.IO.File]::WriteAllText((Join-Path $Out "summary.md"), $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
[System.IO.File]::WriteAllText((Join-Path $Out "result.json"), (ConvertTo-Json @($results) -Depth 6), $utf8)
Write-Host $sb.ToString()
