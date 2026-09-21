# 生成したテストデータから、実行した人の情報を取り除く（make_testdata.ps1 が最後に呼ぶ）。
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\testdata\scrub_personal.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\testdata\scrub_personal.ps1 -Names "山田 太郎"
#
# リポジトリは公開しているため、Office が埋め込む作成者名・保存先の絶対パス・アカウント ID をコミットできない。
# 埋め込まれる場所の一覧は README.md の「生成」にある。
#
# 名前は決め打ちにせず、次から集めて test に置き換える（実行する人ごとに違うため）。
#   - Windows のユーザー名（$env:USERNAME）と、プロファイルのフォルダ名（C:\Users\<名前>）
#   - Office のユーザー名・頭文字（レジストリの HKCU\Software\Microsoft\Office\Common\UserInfo）
#   - -Names で渡した名前（Microsoft アカウントの表示名など）
#
# あわせて、名前を知らなくても消せるものは形で消す。
#   - docProps/core.xml の作成者・最終更新者、ODF の meta.xml の作成者 … test
#   - docProps/app.xml の会社名・管理者 … 空
#   - コメント・変更履歴に付く Microsoft アカウント ID（userId 属性）… 同じ長さの 0
#
# 形式によって置き換え方が違う。
#   - ZIP（OOXML・ODF）: 中の XML を書き換える。書き換えが要らないエントリは元の圧縮データのまま残す
#     （ODF の mimetype は無圧縮で先頭に置く決まりがあるため、開かずにそのまま残す）
#   - 旧形式（複合ドキュメント）・PDF・ZIP の中のバイナリ（.xlsb の xl/workbook.bin など）:
#     レコード長やオフセットが崩れるため、バイト長を変えずに置き換える（test_ のように _ で長さを揃える）
#   - テキスト（.txt など）: そのまま test に置き換える
param (
    [string[]]$Path = @("$PSScriptRoot\office", "$PSScriptRoot\設定例"),
    [string[]]$Names = @(),
    [switch]$NoAutoNames   # 自動で名前を集めない（テスト用）
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$replacement = "test"
$textExtensions = @(".txt", ".csv", ".tsv", ".xml", ".md", ".json", ".config", ".ini")
$zipXmlPattern = '(?i)\.(xml|rels|vml)$'

$encodings = @(
    [System.Text.Encoding]::GetEncoding(932),
    (New-Object System.Text.UTF8Encoding($false)),
    [System.Text.Encoding]::Unicode
)

# 置き換える名前を集める
function getNames {
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($name in $Names) { $found.Add($name) }
    if (!$NoAutoNames) {
        $found.Add($env:USERNAME)
        $found.Add((Split-Path $env:USERPROFILE -Leaf))
        $userInfo = "HKCU:\Software\Microsoft\Office\Common\UserInfo"
        if (Test-Path $userInfo) {
            $info = Get-ItemProperty $userInfo
            $found.Add([string]$info.UserName)
            $found.Add([string]$info.UserInitials)
        }
    }
    # 短すぎる名前・ありふれた語は関係ない文字列まで置き換えてしまうため除く（user だと C:\Users まで変わる）。test 自体も除く
    $common = '^(?i)(test[_0-9]*|users?|admin(istrator)?|owner|guest|default|public|windows|microsoft|office|pc)$'
    $result = @($found | ForEach-Object { if ($_) { $_.Trim() } } |
        Where-Object { $_ -and $_.Length -ge 2 -and $_ -notmatch $common } |
        Sort-Object Length -Descending | Select-Object -Unique)
    # 1 文字の差で別の単語と重なりやすい 2 文字の名前は、日本語（頭文字ではない）のときだけ使う
    return @($result | Where-Object { $_.Length -ge 3 -or $_ -match '[^\x00-\x7F]' })
}

# 同じバイト長の置き換え文字列（test_ のように _ で埋める。長い名前より短ければ切る）
function sameLengthBytes([System.Text.Encoding]$encoding, [int]$byteLength) {
    $unit = $encoding.GetByteCount("a")
    $chars = [Math]::Floor($byteLength / $unit)
    $text = ($replacement + ("_" * $chars)).Substring(0, $chars)
    return $encoding.GetBytes($text)
}

# ASCII の英字は大文字・小文字を区別しない（パスの中の名前は先頭が大文字のこともある）
function foldByte([byte]$value) {
    if ($value -ge 0x41 -and $value -le 0x5A) { return [byte]($value + 0x20) }
    return $value
}

# $pattern は小文字にしたもの（foldByte 済み）を渡す
function indexOfBytes([byte[]]$data, [byte[]]$pattern, [int]$start) {
    $lower = $pattern[0]
    $upper = if ($lower -ge 0x61 -and $lower -le 0x7A) { [byte]($lower - 0x20) } else { $lower }
    $last = $data.Length - $pattern.Length
    $i = $start
    while ($i -le $last) {
        # 先頭のバイトの候補を IndexOf で探し（速い）、残りを 1 バイトずつ比べる
        $a = [Array]::IndexOf($data, $lower, $i, $last - $i + 1)
        $b = if ($upper -ne $lower) { [Array]::IndexOf($data, $upper, $i, $last - $i + 1) } else { -1 }
        if ($a -lt 0) { $i = $b } elseif ($b -lt 0) { $i = $a } else { $i = [Math]::Min($a, $b) }
        if ($i -lt 0) { return -1 }
        $match = $true
        for ($k = 1; $k -lt $pattern.Length; $k++) {
            if ((foldByte $data[$i + $k]) -ne $pattern[$k]) { $match = $false; break }
        }
        if ($match) { return $i }
        $i++
    }
    return -1
}

# バイト長を変えずに名前を置き換える。置き換えた数を返す
function scrubBytesInPlace([byte[]]$data, [string[]]$names) {
    $count = 0
    foreach ($name in $names) {
        foreach ($encoding in $encodings) {
            # 大文字・小文字の違いは ASCII の英字だけ見る（indexOfBytes）
            $pattern = [byte[]]@($encoding.GetBytes($name) | ForEach-Object { foldByte $_ })
            if ($pattern.Length -eq 0) { continue }
            $fill = sameLengthBytes $encoding $pattern.Length
            $position = indexOfBytes $data $pattern 0
            while ($position -ge 0) {
                [Array]::Copy($fill, 0, $data, $position, $fill.Length)
                # 半端なバイト（UTF-16 で奇数長になる場合など）は _ で埋める
                for ($k = $fill.Length; $k -lt $pattern.Length; $k++) { $data[$position + $k] = 0x5F }
                $count++
                $position = indexOfBytes $data $pattern ($position + $pattern.Length)
            }
        }
    }
    return $count
}

# XML（ZIP の中）の文字列を置き換える。長さは変わってよい
function scrubXml([string]$text, [string[]]$names) {
    foreach ($name in $names) {
        $text = [regex]::Replace($text, [regex]::Escape($name), $replacement, "IgnoreCase")
    }
    $text = [regex]::Replace($text, '<(dc:creator|cp:lastModifiedBy|meta:initial-creator)>[^<]*</\1>', "<`$1>$replacement</`$1>")
    $text = [regex]::Replace($text, '<(Company|Manager)>[^<]*</\1>', '<$1></$1>')
    $text = [regex]::Replace($text, '(?i)\b((?:w15:)?userId=")([^"]*)(")', {
            param($match)
            $match.Groups[1].Value + ("0" * $match.Groups[2].Value.Length) + $match.Groups[3].Value
        })
    return $text
}

# ZIP のエントリ 1 つ分の中身を置き換えた結果を返す。変わらなければ $null
function scrubEntry([string]$entryName, [byte[]]$bytes, [string[]]$names) {
    if ($entryName -match $zipXmlPattern) {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $offset = if ($hasBom) { 3 } else { 0 }
        $text = $utf8.GetString($bytes, $offset, $bytes.Length - $offset)
        $new = scrubXml $text $names
        if ($new -ceq $text) { return $null }
        $body = $utf8.GetBytes($new)
        if (!$hasBom) { return ,$body }
        return ,([byte[]](@(0xEF, 0xBB, 0xBF) + $body))
    }
    $copy = [byte[]]$bytes.Clone()
    if ((scrubBytesInPlace $copy $names) -eq 0) { return $null }
    return ,$copy
}

function readEntry($entry) {
    $buffer = New-Object System.IO.MemoryStream
    $stream = $entry.Open()
    try { $stream.CopyTo($buffer) } finally { $stream.Dispose() }
    return ,$buffer.ToArray()
}

# ZIP を書き換える。置き換えたエントリの数を返す
function scrubZip([string]$file, [string[]]$names) {
    # 1. 読み取りで、書き換えが要るエントリだけを選ぶ
    $changes = @{}
    $archive = [System.IO.Compression.ZipFile]::OpenRead($file)
    try {
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName.EndsWith("/")) { continue }
            $new = scrubEntry $entry.FullName (readEntry $entry) $names
            if ($null -ne $new) { $changes[$entry.FullName] = $new }
        }
    } finally {
        $archive.Dispose()
    }
    if ($changes.Count -eq 0) { return 0 }

    # 2. 更新モードで、そのエントリだけを書き換える（開かなかったエントリは元の圧縮データのまま書き戻される）
    $stream = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    $archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        foreach ($entry in @($archive.Entries)) {
            if (!$changes.ContainsKey($entry.FullName)) { continue }
            $content = $changes[$entry.FullName]
            $writer = $entry.Open()
            try {
                $writer.SetLength(0)
                $writer.Write($content, 0, $content.Length)
            } finally {
                $writer.Dispose()
            }
        }
    } finally {
        $archive.Dispose()
        $stream.Dispose()
    }
    return $changes.Count
}

function isZip([string]$file) {
    $stream = [System.IO.File]::OpenRead($file)
    try {
        $head = New-Object byte[] 4
        $read = $stream.Read($head, 0, 4)
        return $read -eq 4 -and $head[0] -eq 0x50 -and $head[1] -eq 0x4B -and $head[2] -eq 0x03 -and $head[3] -eq 0x04
    } finally {
        $stream.Dispose()
    }
}

# テキストは長さを変えて置き換える。ASCII 以外を含む名前は、元の文字コードのまま置き換えるためバイト長を保つ
function scrubText([byte[]]$data, [string[]]$names) {
    $asciiNames = @($names | Where-Object { $_ -notmatch '[^\x00-\x7F]' })
    $otherNames = @($names | Where-Object { $_ -match '[^\x00-\x7F]' })
    $count = 0
    if ($otherNames.Count -gt 0) { $count += scrubBytesInPlace $data $otherNames }
    if ($asciiNames.Count -eq 0) { return @{ Count = $count; Data = $data } }
    # ASCII の名前は Latin1 で読んで置き換えればどの文字コードでも壊れない（1 バイト＝1 文字で往復できる）
    $latin1 = [System.Text.Encoding]::GetEncoding(28591)
    $text = $latin1.GetString($data)
    foreach ($name in $asciiNames) {
        $count += [regex]::Matches($text, [regex]::Escape($name), "IgnoreCase").Count
        $text = [regex]::Replace($text, [regex]::Escape($name), $replacement, "IgnoreCase")
    }
    return @{ Count = $count; Data = $latin1.GetBytes($text) }
}

$names = getNames
if ($names.Count -eq 0) {
    Write-Host "置き換える名前が見つかりません（形で消せるものだけを消します）。" -ForegroundColor Yellow
} else {
    Write-Host "置き換える名前: $($names.Count) 件（画面には出しません）"
}

$changed = 0
foreach ($root in $Path) {
    if (!(Test-Path -LiteralPath $root)) { continue }
    foreach ($item in @(Get-ChildItem -LiteralPath $root -Recurse -File -Force)) {
        $file = $item.FullName
        # 読み取り専用のテストデータもあるため、一時的に外して書き戻す
        $readOnly = $item.IsReadOnly
        try {
            if ($item.Length -eq 0) { continue }
            if (isZip $file) {
                try {
                    if ($readOnly) { $item.IsReadOnly = $false }
                    $count = scrubZip $file $names
                } catch {
                    # 壊れた ZIP（異常系のテストデータ）はバイト列として扱う
                    $count = -1
                }
                if ($count -ge 0) {
                    if ($count -gt 0) { $changed++; Write-Host "  $file（$count か所）" }
                    continue
                }
            }
            $data = [System.IO.File]::ReadAllBytes($file)
            if ($textExtensions -contains $item.Extension.ToLowerInvariant()) {
                $result = scrubText $data $names
                $count = $result.Count
                $data = $result.Data
            } else {
                $count = scrubBytesInPlace $data $names
            }
            if ($count -gt 0) {
                if ($readOnly) { $item.IsReadOnly = $false }
                # WriteAllBytes は隠し属性のファイルに書けない（作り直そうとする）ため、開いて書き直す
                $stream = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
                try {
                    $stream.SetLength(0)
                    $stream.Write($data, 0, $data.Length)
                } finally {
                    $stream.Dispose()
                }
                $changed++
                Write-Host "  $file（$count か所）"
            }
        } finally {
            if ($readOnly) { $item.IsReadOnly = $true }
        }
    }
}
Write-Host "個人情報を取り除きました（$changed ファイル）。"
