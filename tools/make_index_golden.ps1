# 前の版のファイル（見本・golden。tests/testdata/compat/index/<見本の名前>/）を作る手伝い。
#
#   .\tools\make_index_golden.ps1 -WorkspaceDir C:\tebunko_golden\ws -ExportZip C:\tebunko_golden\export.zip `
#       -SampleDir tests\testdata\compat\index\v0.3.1+english-names
#
# 手順（詳しくは tests/testdata/README.md の「前の版のファイル（compat\）」）:
#   1. source/ に見本のファイルを置く（このスクリプトは触らない。手で置く）
#   2. setting.config のクロール対象フォルダに source/ を指したフォルダ（利用者名を含まない固定のパス）を登録し、
#      indexer.ps1 を 1 回動かして workspaceFolder（固定のパス）にワークスペースを作る
#   3. exportIndex でワークスペースを zip に書き出す
#   4. このスクリプトで、ワークスペースの前の版と同じ形で残る部分（content_index\・system_index\・
#      system_index_state.tsv・ingest_status.tsv）と zip を見本へコピーし、file_times.tsv を作る
#   5. expected.json は手で書く（このスクリプトは作らない）
#
# file_times.tsv: source/ のファイルの更新日時を、ingest_status.tsv に記録された文字列（秒まで・タイムゾーンの情報なし）
# から作った Ticks で保存する。git は取り出すときにファイルの更新日時を今の日時にしてしまうため、
# tests/tebunko/indexer/index_compat.Tests.ps1 がこのファイルを読み、source/ を写した先でファイルの更新日時を
# この Ticks に戻す（File.SetLastWriteTime。Ticks は DateTimeKind を持たないため、実行する PC のタイムゾーンに関わらず
# 同じ文字列になる＝タイムゾーンに依存しない）。indexer が取り込み済みと判断する条件
# （indexer_decide.ps1 の getIngestDecision）が、記録した更新日時の文字列と、そのとき実際にファイルから読んだ
# 更新日時の文字列が等しいことのため、Ticks の値そのもの（UTC かどうか）ではなく、文字列を再現できることが要る
param (
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceDir,
    [Parameter(Mandatory = $true)]
    [string]$ExportZip,
    [Parameter(Mandatory = $true)]
    [string]$SampleDir
)

$ErrorActionPreference = "Stop"

$statusFileName = "ingest_status.tsv"
# 見本に残す、ワークスペースの中の前の版と同じ形で残る部分（tmp\・publish\・ingesting.txt・indexing_log.txt・
# gui_error_log.txt・search_results.txt・setting.config は見本に含めない。tests/testdata/README.md に理由を書く）
$keepEntries = @("content_index", "system_index", "system_index_state.tsv", $statusFileName)

function readIngestStatusTimes {
    # ingest_status.tsv を読み、@{ 相対パス（source/ からの相対パス。/ 区切り） = 更新日時の文字列 } を返す
    param ([string]$path)

    $bytes = [System.IO.File]::ReadAllBytes($path)
    $text = (New-Object System.Text.UTF8Encoding($true)).GetString($bytes)
    $lines = @($text -split "`r`n" | Where-Object { $_ -ne "" })
    if ($lines.Count -lt 2) {
        throw "${path} の形が想定と違います（行が足りません）"
    }
    $folderLine = $lines[0].Split("`t")
    if ($folderLine.Count -ne 3) {
        throw "${path} の 1 行目（クロール対象フォルダ）の形が想定と違います"
    }
    $indexName = $folderLine[2]
    $prefix = "${indexName}\"
    $times = [ordered]@{}
    for ($i = 2; $i -lt $lines.Count; $i++) {
        $fields = $lines[$i].Split("`t")
        if ($fields.Count -lt 2) { continue }
        $relative = $fields[0]
        if ($relative.StartsWith($prefix)) {
            $relative = $relative.Substring($prefix.Length)
        }
        $times[$relative.Replace("\", "/")] = $fields[1]
    }
    return $times
}

if (!(Test-Path -LiteralPath $WorkspaceDir)) {
    throw "ワークスペース ${WorkspaceDir} が見つかりません"
}
$statusPath = Join-Path $WorkspaceDir $statusFileName
if (!(Test-Path -LiteralPath $statusPath)) {
    throw "${statusPath} が見つかりません（先に indexer.ps1 を動かしてください）"
}
if (!(Test-Path -LiteralPath $ExportZip)) {
    throw "zip ${ExportZip} が見つかりません（先に exportIndex で書き出してください）"
}
if (!(Test-Path -LiteralPath (Join-Path $SampleDir "source"))) {
    throw "${SampleDir}\source が見つかりません（先に見本のファイルを置いてください）"
}

$wsDest = Join-Path $SampleDir "ws"
if (Test-Path -LiteralPath $wsDest) {
    Remove-Item -LiteralPath $wsDest -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $wsDest | Out-Null

foreach ($entry in $keepEntries) {
    $src = Join-Path $WorkspaceDir $entry
    if (!(Test-Path -LiteralPath $src)) { continue }
    $dest = Join-Path $wsDest $entry
    if ((Get-Item -LiteralPath $src) -is [System.IO.DirectoryInfo]) {
        Copy-Item -LiteralPath $src -Destination $dest -Recurse -Force
    } else {
        Copy-Item -LiteralPath $src -Destination $dest -Force
    }
}

Copy-Item -LiteralPath $ExportZip -Destination (Join-Path $SampleDir "export.zip") -Force

$times = readIngestStatusTimes $statusPath
$timesLines = New-Object System.Collections.Generic.List[string]
$timesLines.Add("# source/ からの相対パス（/ 区切り）<TAB>Ticks（DateTimeKind なし。更新日時の文字列をそのまま解いた値）")
foreach ($relative in $times.Keys) {
    $parsed = [datetime]::ParseExact($times[$relative], "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
    $timesLines.Add("${relative}`t$($parsed.Ticks)")
}
$timesPath = Join-Path $SampleDir "file_times.tsv"
[System.IO.File]::WriteAllText($timesPath, (($timesLines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))

Write-Host "見本を作りました: ${SampleDir}"
Write-Host "  ws\ … $((Get-ChildItem -LiteralPath $wsDest -Recurse -File).Count) ファイル"
Write-Host "  export.zip"
Write-Host "  file_times.tsv … $($times.Count) 件"
