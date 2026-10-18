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
#   4. このスクリプトで、ワークスペースの形のまま残す部分（インデックスの置き場・システムインデックス・
#      システムインデックスの状態・取り込み一覧。置き場の名前は Workspace クラスから取る）と zip を見本へコピーし、file_times.tsv を作る
#   5. expected.json は手で書く（このスクリプトは作らない）
#
# file_times.tsv: 見本のファイルの更新日時の記録。git は取り出すときにファイルの更新日時を今の日時にしてしまうため、
# tests/tebunko/indexer/index_compat.Tests.ps1 がこのファイルを読み、source/・ws/ を写した先でファイルの更新日時を戻す。
#   source/ … ingest_status.tsv に記録された文字列（秒まで・タイムゾーンの情報なし）から作った Ticks。File.SetLastWriteTime で戻す
#             （Ticks は DateTimeKind を持たないため、実行する PC のタイムゾーンに関わらず同じ文字列になる）。
#             indexer が取り込み済みと判断する条件（indexer_decide.ps1 の getIngestDecision）が、記録した更新日時の文字列と
#             実際にファイルから読んだ更新日時の文字列が等しいことのため、文字列を再現できることが要る
#   ws/     … 本文インデックスとシステムインデックスのフォルダのファイルの実際の更新日時（UTC の Ticks）。File.SetLastWriteTimeUtc で戻す
param (
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceDir,
    [Parameter(Mandatory = $true)]
    [string]$ExportZip,
    [Parameter(Mandatory = $true)]
    [string]$SampleDir
)

$ErrorActionPreference = "Stop"

# ワークスペースの中の置き場の名前は、本体の定義（クラスだけのファイル）から取る
. ([WildcardPattern]::Escape("$PSScriptRoot\..\scripts\tebunko\core\workspace.ps1".Replace('`', '``')))

function getWorkspaceRelativePath {
    # ワークスペースの中の場所（full）の、ワークスペースからの相対パスを返す
    param ([string]$workspaceDir, [string]$full)

    return $full.Substring($workspaceDir.TrimEnd("\").Length + 1)
}

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
$sourceWorkspace = [Workspace]::new($WorkspaceDir)
# 取り込み一覧は、今の置き場に無ければ前の版の置き場（前の版の見本を作るとき）にある
$statusPath = $sourceWorkspace.StatusFile
if (!(Test-Path -LiteralPath $statusPath)) { $statusPath = $sourceWorkspace.OldStatusFile }
if (!(Test-Path -LiteralPath $statusPath)) {
    throw "取り込み一覧が見つかりません（先に indexer.ps1 を動かしてください）: ${WorkspaceDir}"
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
[void][System.IO.Directory]::CreateDirectory($wsDest)

# 見本に残す、ワークスペースの中の形のまま残る部分（tmp\・publish\・store\journal\・indexing_log.txt・
# gui_error_log.txt・search_results.txt・setting.config は見本に含めない。tests/testdata/README.md に理由を書く）
$keepEntries = @($sourceWorkspace.StoreDir, $sourceWorkspace.OldIndexDir, $sourceWorkspace.OldStatusFile,
    $sourceWorkspace.SystemIndexDir, $sourceWorkspace.SystemIndexStateFile) | ForEach-Object { getWorkspaceRelativePath $WorkspaceDir $_ }
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
# store\ ごと写した中の、取り込みのジャーナル（store\journal\）は、取り込み中の記録で、前の版との互換の確かめには要らないため消す
$journalDest = Join-Path $wsDest (getWorkspaceRelativePath $WorkspaceDir (Split-Path -Parent $sourceWorkspace.JournalFile))
if (Test-Path -LiteralPath $journalDest) {
    Remove-Item -LiteralPath $journalDest -Recurse -Force
}

Copy-Item -LiteralPath $ExportZip -Destination (Join-Path $SampleDir "export.zip") -Force

$times = readIngestStatusTimes $statusPath
$timesLines = New-Object System.Collections.Generic.List[string]
$timesLines.Add("# 見本のフォルダからの相対パス（/ 区切り）<TAB>Ticks。source/ は DateTimeKind なし（更新日時の文字列をそのまま解いた値）、ws/ は UTC")
foreach ($relative in $times.Keys) {
    $parsed = [datetime]::ParseExact($times[$relative], "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
    $timesLines.Add("source/${relative}`t$($parsed.Ticks)")
}
# ws/ の本文インデックスとシステムインデックスのフォルダのファイルは、実際の更新日時（UTC の Ticks）を残す。
# システムインデックスの作り直しの判定（system_index.ps1）が、本文インデックスの TSV と system_index.txt の更新日時の前後と、
# system_index_state.tsv の「反映待ち」に書いた system_index.txt の更新日時（UTC の Ticks）を見るため、取り出した直後の日時のままでは
# 「古い」と見なして作り直してしまう
$wsFileCount = 0
$timedDirs = @($sourceWorkspace.IndexDir, $sourceWorkspace.OldIndexDir, $sourceWorkspace.SystemIndexDir) | ForEach-Object { getWorkspaceRelativePath $WorkspaceDir $_ }
foreach ($sub in $timedDirs) {
    $subDir = Join-Path $wsDest $sub
    if (!(Test-Path -LiteralPath $subDir)) { continue }
    foreach ($file in @(Get-ChildItem -LiteralPath $subDir -Recurse -File | Sort-Object FullName)) {
        $relative = $file.FullName.Substring($wsDest.Length + 1).Replace("\", "/")
        $timesLines.Add("ws/${relative}`t$($file.LastWriteTimeUtc.Ticks)")
        $wsFileCount++
    }
}
$timesPath = Join-Path $SampleDir "file_times.tsv"
[System.IO.File]::WriteAllText($timesPath, (($timesLines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))

Write-Host "見本を作りました: ${SampleDir}"
Write-Host "  ws\ … $((Get-ChildItem -LiteralPath $wsDest -Recurse -File).Count) ファイル"
Write-Host "  export.zip"
Write-Host "  file_times.tsv … source $($times.Count) 件・ws ${wsFileCount} 件"
