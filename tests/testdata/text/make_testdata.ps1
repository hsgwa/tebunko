# tests\testdata\text\ のサンプルを作り直す。
#   対象の拡張子（shared\core\text_file.ps1 の ${textOpenExtensions} + ${textNotepadExtensions}。75 個）を
#   1 つずつ、拡張子を足した順番の番号「TXT<NN>」を付けたファイルにする。
#   実行・登録が既定の動作になる 7 個（.bat .cmd .ps1 .vbs .js .reg .sh）は、
#   実行しても何も起きない中身（コメントだけ）にする。
#   文字コードの違い・NUL を含むバイナリ・64KB を超えるファイルも、下の「特別なファイル」で作る。
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\testdata\text\make_testdata.ps1

$here = (Resolve-Path "$PSScriptRoot\..\..").Path
. "$here\..\scripts\shared\core\text_file.ps1"

$outDir = $PSScriptRoot
$abnormalDir = Join-Path $outDir "異常系"

if (Test-Path $outDir) {
    Get-ChildItem $outDir -File | Where-Object { $_.Name -ne "make_testdata.ps1" } | Remove-Item -Force
}
if (Test-Path $abnormalDir) { Remove-Item $abnormalDir -Recurse -Force }
New-Item -ItemType Directory -Path $abnormalDir -Force | Out-Null

function toUtf8Bytes([string]$text, [bool]$bom = $false) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    if ($bom) { return , (@(0xEF, 0xBB, 0xBF) + $bytes) }
    return , $bytes
}

function toUtf16LeBytes([string]$text, [bool]$bom = $false) {
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($text)
    if ($bom) { return , (@(0xFF, 0xFE) + $bytes) }
    return , $bytes
}

function toUtf16BeBytes([string]$text, [bool]$bom = $false) {
    $bytes = [System.Text.Encoding]::BigEndianUnicode.GetBytes($text)
    if ($bom) { return , (@(0xFE, 0xFF) + $bytes) }
    return , $bytes
}

function toCodePageBytes([int]$codePage, [string]$text) {
    return , ([System.Text.Encoding]::GetEncoding($codePage).GetBytes($text))
}

# 拡張子ごとの、既定と違う中身・文字コード（キーは拡張子。小文字）。
# ここに無い拡張子は既定（UTF-8・BOM 無し・"TXT-<NN> サンプル（<拡張子>）" ＋ "山田 太郎"）になる。
$overrides = @{
    ".txt" = @{
        text = { param($id, $ext) "TXT-$id サンプル（$ext）`r`n山田 太郎`r`n" }
        bytes = { param($text) toUtf8Bytes $text $true }  # UTF-8（BOM 付き）
    }
    ".csv" = @{
        text = { param($id, $ext) "TXT-$id,サンプル,$ext`r`n山田太郎,やまだたろう`r`n" }
        bytes = { param($text) toCodePageBytes 51932 $text }  # EUC-JP（かなを含む）
    }
    ".ini" = @{
        text = { param($id, $ext) "TXT-$id サンプル（$ext）`r`n[設定]`r`n名前=山田太郎`r`n" }
        bytes = { param($text) toCodePageBytes 932 $text }  # Shift_JIS
    }
    ".yml" = @{
        text = { param($id, $ext) "TXT-$id サンプル（$ext）`r`n名前: 山田太郎`r`n" }
        bytes = { param($text) toUtf16LeBytes $text $true }  # UTF-16LE（BOM 付き）
    }
    ".properties" = @{
        text = { param($id, $ext) "TXT-$id サンプル（$ext）`r`n名前=山田太郎`r`n" }
        bytes = { param($text) toUtf16BeBytes $text $true }  # UTF-16BE（BOM 付き）
    }
    ".graphql" = @{
        text = { param($id, $ext) "TXT-$id サンプル（$ext）`r`nこれは日本語のテストです。ひらがなとカタカナと漢字が入っています。山田`r`n" }
        bytes = { param($text) toCodePageBytes 50220 $text }  # ISO-2022-JP
    }
    ".toml" = @{
        text = { param($id, $ext) "TXT-$id サンプル（$ext）`r`nこれは日本語のテストです。ひらがなとカタカナと漢字が入っています。山田`r`n" }
        bytes = { param($text) toUtf16LeBytes $text $false }  # UTF-16LE（BOM 無し。かなの割合で判定）
    }
    # 実行・登録が既定の動作になる 7 個。中身はコメントだけ（実行しても何も起きない）
    ".bat" = @{ text = { param($id, $ext) "REM TXT-$id サンプル（$ext）`r`nREM 山田 太郎`r`n" } }
    ".cmd" = @{ text = { param($id, $ext) "REM TXT-$id サンプル（$ext）`r`nREM 山田 太郎`r`n" } }
    # .ps1 は AGENTS.md の決まり（コミット前の検査）で BOM 付き UTF-8 が必須
    ".ps1" = @{
        text = { param($id, $ext) "# TXT-$id サンプル（$ext）`r`n# 山田 太郎`r`n" }
        bytes = { param($text) toUtf8Bytes $text $true }  # UTF-8（BOM 付き）
    }
    ".vbs" = @{ text = { param($id, $ext) "' TXT-$id サンプル（$ext）`r`n' 山田 太郎`r`n" } }
    ".js"  = @{ text = { param($id, $ext) "// TXT-$id サンプル（$ext）`r`n// 山田 太郎`r`n" } }
    ".reg" = @{ text = { param($id, $ext) "; TXT-$id サンプル（$ext）`r`n; 山田 太郎`r`n" } }
    ".sh"  = @{ text = { param($id, $ext) "# TXT-$id サンプル（$ext）`r`n# 山田 太郎`r`n" } }
}

$allExt = @(${textOpenExtensions}) + @(${textNotepadExtensions})
$created = @()
for ($i = 0; $i -lt $allExt.Count; $i++) {
    $ext = $allExt[$i]
    $id = "{0:D2}" -f ($i + 1)
    $override = $overrides[$ext]
    if ($null -ne $override -and $override.ContainsKey("text")) {
        $text = & $override.text $id $ext
    } else {
        $text = "TXT-$id サンプル（$ext）`r`n山田 太郎`r`n"
    }
    if ($null -ne $override -and $override.ContainsKey("bytes")) {
        $bytes = & $override.bytes $text
    } else {
        $bytes = toUtf8Bytes $text $false
    }
    $path = Join-Path $outDir "TXT$id$ext"
    [System.IO.File]::WriteAllBytes($path, [byte[]]$bytes)
    $created += [pscustomobject]@{ Id = $id; Ext = $ext; Path = $path }
}

# 特別なファイル（異常系・大きいファイル）
# バイナリ.log: NUL を含む、テキストとして取り込めないファイル
$binary = New-Object byte[] 256
for ($i = 0; $i -lt 256; $i++) { $binary[$i] = [byte]$i }  # 0x00（NUL）を含む
[System.IO.File]::WriteAllBytes((Join-Path $abnormalDir "バイナリ.log"), $binary)

# 末尾NUL.log: 64KB（先頭のサンプル範囲）の外にだけ NUL が 1 つある、70000 バイトのログらしいファイル。
# 強制終了などで末尾が NUL で埋まった状態を表す。ファイル全体を見て NUL を検出するため、取り込めない
$logLine = "2026-10-03 00:00:00 INFO サンプルのログ行です。`r`n"
$logBytes = [System.Text.Encoding]::UTF8.GetBytes($logLine)
$repeat = [Math]::Ceiling(70000.0 / $logBytes.Length)
$logAll = New-Object System.Collections.Generic.List[byte]
for ($i = 0; $i -lt $repeat; $i++) { $logAll.AddRange($logBytes) }
$tailNul = $logAll.ToArray() + [byte]0
[System.IO.File]::WriteAllBytes((Join-Path $abnormalDir "末尾NUL.log"), $tailNul)

# 大きいテキスト.log: 64KB を超えるが、NUL を含まない普通のテキスト（成功するケース）
$bigText = ""
for ($i = 1; $i -le 1500; $i++) { $bigText += "TXT-大 サンプルのログ行 $i 行目。山田 太郎`r`n" }
[System.IO.File]::WriteAllBytes((Join-Path $outDir "大きいテキスト.log"), (toUtf8Bytes $bigText $false))

Write-Output ("作成したファイル: {0} 個（拡張子ごと）＋ 異常系 2 個 ＋ 大きいテキスト.log 1 個" -f $created.Count)
