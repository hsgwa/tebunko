# Office からの取り込みの速さを測るデータ（Office ファイルの束）を作る。tools\measure_perf.ps1 の -Office に渡す。
#
#   .\tools\perf\new_ingest_data.ps1 -Dest <置くフォルダ（空か、まだ無いフォルダ）> -Docx 200 -Pptx 200
#   -Docx / -Pptx  … .docx / .pptx の数（既定 200）。テストデータの複製。Office を使わずに読む（読み取りのスレッド）
#   -Doc / -Ppt    … .doc / .ppt の数（既定 0）。テストデータの複製。Word / PowerPoint で新形式に変換して読む
#   -Xlsx          … .xlsx の数（既定 0）。-Books に渡したフォルダから、パスの順に先頭 N 冊を写す。Excel で読む
#   -Books         … tebunko-perfdata の new_books.ps1 で作ったブックのフォルダ（-Xlsx が 1 以上のとき）
#   -Source        … 元にするテストデータ（既定はこのリポジトリの tests\testdata\office）
#
# ファイルは 50 個ごとにフォルダを分ける（フォルダ001\資料0001.docx の形。フォルダごとの集約ファイルの書き出しも一緒に測るため）。
# 種類は xlsx・docx・pptx・doc・ppt の順に並べ、番号は全体で通す。同じ引数からは、同じ構成・同じ中身のデータができる。
# 元のファイルは作成者名などを除いたテストデータなので、作ったデータにも個人情報は入らない。
param (
    [Parameter(Mandatory = $true)][string]$Dest,
    [int]$Docx = 200,
    [int]$Pptx = 200,
    [int]$Doc = 0,
    [int]$Ppt = 0,
    [int]$Xlsx = 0,
    [string]$Books,
    [string]$Source = (Join-Path (Split-Path $PSScriptRoot -Parent | Split-Path -Parent) "tests\testdata\office")
)

$ErrorActionPreference = "Stop"

# 元のファイルの場所（Source からの相対）。テストデータの名前が変わったらここを直す
$templates = [ordered]@{
    docx = "Word\長文.docx"
    pptx = "PowerPoint\多数スライド.pptx"
    doc  = "Word\形式\旧形式.doc"
    ppt  = "PowerPoint\形式\旧形式.ppt"
}
$perFolder = 50

$Source = (Resolve-Path -LiteralPath $Source).ProviderPath.TrimEnd("\")
if ([System.IO.Directory]::Exists($Dest) -and @([System.IO.Directory]::GetFileSystemEntries($Dest)).Count -gt 0) {
    throw "置き場所が空ではありません。空のフォルダか、まだ無いフォルダを指定してください: $Dest"
}
[void][System.IO.Directory]::CreateDirectory($Dest)
$Dest = (Resolve-Path -LiteralPath $Dest).ProviderPath.TrimEnd("\")

# 種類ごとの元のファイルの一覧（xlsx は複数、それ以外は同じ 1 つを複製する）
$plan = New-Object System.Collections.Generic.List[object]
if ($Xlsx -gt 0) {
    if (!$Books) { throw "-Xlsx を使うときは、-Books に new_books.ps1 で作ったブックのフォルダを指定してください。" }
    $files = [string[]][System.IO.Directory]::GetFiles((Resolve-Path -LiteralPath $Books).ProviderPath, "*.xlsx", "AllDirectories")
    [Array]::Sort($files, [System.StringComparer]::Ordinal)
    if ($files.Count -lt $Xlsx) { throw ".xlsx が足りません（-Xlsx $Xlsx に対して $($files.Count) 冊）。" }
    foreach ($f in @($files | Select-Object -First $Xlsx)) { $plan.Add(@{ Ext = "xlsx"; From = $f }) }
}
foreach ($pair in @(@("docx", $Docx), @("pptx", $Pptx), @("doc", $Doc), @("ppt", $Ppt))) {
    if ($pair[1] -le 0) { continue }
    $from = Join-Path $Source $templates[$pair[0]]
    if (!(Test-Path -LiteralPath $from)) { throw "元のファイルがありません: $($templates[$pair[0]])（Source を確かめてください）" }
    for ($i = 0; $i -lt $pair[1]; $i++) { $plan.Add(@{ Ext = $pair[0]; From = $from }) }
}

$n = 0
foreach ($item in $plan) {
    $n++
    $folder = Join-Path $Dest ("フォルダ{0:000}" -f [int][Math]::Ceiling($n / $perFolder))
    [void][System.IO.Directory]::CreateDirectory($folder)
    [System.IO.File]::Copy($item.From, (Join-Path $folder ("資料{0:0000}.{1}" -f $n, $item.Ext)))
}
Write-Host ("データを作りました: {0} ファイル（xlsx {1}・docx {2}・pptx {3}・doc {4}・ppt {5}）" -f $n, $Xlsx, $Docx, $Pptx, $Doc, $Ppt)
