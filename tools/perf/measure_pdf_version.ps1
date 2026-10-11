# PDF の読み取り 1 つの版の速さを、1 つのプロセスで測る（tools\perf\measure_pdf.ps1 が版ごとに起動する）。
#
#   powershell.exe -NoProfile -File .\tools\perf\measure_pdf_version.ps1 -Scripts <scripts フォルダ> -Pdf <PDF> -Repeat 3 -Out <結果の JSON>
#   -Scripts … 測る版の scripts フォルダ（shared\core\pdf_crypt.ps1・pdf_file.ps1 を読む）
#   -Repeat  … readPdfFile を測る回数。1 回目は読み込み直後の時間で、2 回目以降もそのまま出す
#   -Profile … 時間を測ったあと、関数ごとの時間を 1 回だけ測る（関数を包んで数えるため、そのぶん遅くなる。順位をつけるための値）
# 出力の JSON: Seconds（1 回ごとの秒）・Hash（読み取った文字の SHA-256。版どうしで結果が変わっていないかを見る）・Pages・Profile
param (
    [Parameter(Mandatory = $true)][string]$Scripts,
    [Parameter(Mandatory = $true)][string]$Pdf,
    [int]$Repeat = 3,
    [switch]$Profile,
    [Parameter(Mandatory = $true)][string]$Out
)

$ErrorActionPreference = "Stop"

# ファイルの読み取りに使う toLongPath だけを、測る対象の外で用意する（長いパスは使わない）
function toLongPath([string]$path) { return $path }
. ([WildcardPattern]::Escape((Join-Path $Scripts "shared\core\pdf_crypt.ps1").Replace('`', '``')))
. ([WildcardPattern]::Escape((Join-Path $Scripts "shared\core\pdf_file.ps1").Replace('`', '``')))

function getPagesHash([object[]]$pages) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes(($pages -join "`f"))
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "")
    } finally { $sha.Dispose() }
}

$seconds = New-Object System.Collections.Generic.List[double]
$hash = ""
$pageCount = 0
for ($i = 0; $i -lt $Repeat; $i++) {
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $pages = readPdfFile $Pdf
    $watch.Stop()
    $seconds.Add($watch.Elapsed.TotalSeconds)
    $hash = getPagesHash $pages
    $pageCount = $pages.Count
}

$profileRows = @()
if ($Profile) {
    $names = @(
        "readPdfPages", "readPdfXref", "readPdfIndirectAt", "getPdfObject", "getPdfStreamBytes", "expandPdfFlate", "undoPdfPredictor",
        "readPdfToken", "readPdfValue", "readPdfToUnicode", "newPdfFont", "getPdfFont", "readPdfPageText", "convertPdfStringToText",
        "convertPdfCode", "convertPdfUtf16", "getPdfGlyphCount", "addPdfText", "testPdfRef", "resolvePdf"
    )
    $stats = @{}
    function newProfiledFunction([string]$name, [scriptblock]$original) {
        $entry = @{ Calls = 0L; Ticks = 0L; Depth = 0 }
        $stats[$name] = $entry
        return {
            $entry.Calls++
            $started = 0L
            if ($entry.Depth -eq 0) { $started = [System.Diagnostics.Stopwatch]::GetTimestamp() }
            $entry.Depth++
            try { & $original @args } finally {
                $entry.Depth--
                if ($entry.Depth -eq 0) { $entry.Ticks += [System.Diagnostics.Stopwatch]::GetTimestamp() - $started }
            }
        }.GetNewClosure()
    }
    foreach ($name in $names) {
        $command = Get-Command $name -CommandType Function -ErrorAction SilentlyContinue
        if ($null -eq $command) { continue }
        Set-Item -Path "function:$name" -Value (newProfiledFunction $name $command.ScriptBlock)
    }
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    [void](readPdfFile $Pdf)
    $watch.Stop()
    $total = $watch.Elapsed.TotalSeconds
    $profileRows = @($stats.Keys | Sort-Object | ForEach-Object {
        @{ Name = $_; Calls = $stats[$_].Calls; Seconds = $stats[$_].Ticks / [double][System.Diagnostics.Stopwatch]::Frequency }
    })
    $profileTotal = $total
} else {
    $profileTotal = 0.0
}

$result = @{
    Seconds = @($seconds)
    Hash = $hash
    Pages = $pageCount
    Bytes = ([System.IO.FileInfo]::new($Pdf)).Length
    Profile = $profileRows
    ProfileTotal = $profileTotal
}
[System.IO.File]::WriteAllText($Out, (ConvertTo-Json $result -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
