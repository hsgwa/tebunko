# 単一 .ps1 版に同梱する、第三者の部品のライセンス文（scripts\shared\licenses\LICENSE-*.txt）の扱い。
# tools\new_single_script.ps1 が読み込んで使う。テストは tests\tools\single_script_license.Tests.ps1。設計は docs/design/structure/single-script.md。
#
# 文面は元のファイルから読み、書き換えない。行ごとに、空行は # だけ、それ以外は # と 1 つの空白を行の頭に足すだけにする。
# 元に戻す関数（restoreLicenseFromComment）は、その逆（# だけの行は空行、# と空白で始まる行は頭の 2 文字を外す。どちらでもない行があれば $null）で、
# 道具の検査とテストが、同じ関数で「元のファイルと同じ中身か」を確かめる。

# 改行を CRLF にそろえ、末尾が改行で終わるようにする（取り出した文字列と比べるときの、元のファイルの形）
function normalizeLicenseText {
    param ([string]$text)

    $normalized = $text -replace "\r?\n", "`r`n"
    if (!$normalized.EndsWith("`r`n", [System.StringComparison]::Ordinal)) {
        $normalized += "`r`n"
    }
    return $normalized
}

# コメントにできる文面か確かめる。空（空白だけ）・改行を伴わない CR だけの行区切りがあるときは止める
# （PowerShell は CR だけでも改行と読むため、その後ろがコメントの外に出てしまう）
function assertLicenseText {
    param (
        [string]$fileName,
        [string]$text
    )

    if (($text -replace "\s", "").Length -eq 0) {
        throw "ライセンス文が空です: $fileName"
    }
    if ($text -match "\r(?!\n)") {
        throw "ライセンス文に、改行（CRLF・LF）を伴わない CR だけの行区切りがあります: $fileName"
    }
}

# ライセンス文の文字列を、「ここから」〜「ここまで」の印を付けた # のコメントにする。空・CR だけの行区切りがあれば止める
function convertToLicenseComment {
    param (
        [string]$fileName,
        [string]$text
    )

    assertLicenseText $fileName $text
    $block = New-Object System.Collections.Generic.List[string]
    [void]$block.Add("# ---- 第三者の部品のライセンス文: $fileName（ここから）----")
    $lines = @($text -split "\r?\n")
    if ($lines[$lines.Count - 1] -eq "") {
        $lines = @($lines[0..($lines.Count - 2)])
    }
    foreach ($line in $lines) {
        [void]$block.Add($(if ($line -eq "") { "#" } else { "# $line" }))
    }
    [void]$block.Add("# ---- 第三者の部品のライセンス文: $fileName（ここまで）----")
    return ($block -join "`r`n")
}

# コメントにしたライセンス文を、# を外して元の形（CRLF・末尾は改行）の文字列に戻す。印が見つからないとき、印の間にコメントでない行（# だけでも # と空白で始まる行でもない行）が混ざるときは $null
function restoreLicenseFromComment {
    param (
        [string]$scriptText,
        [string]$fileName
    )

    $lines = @($scriptText -split "\r?\n")
    $from = [array]::IndexOf($lines, "# ---- 第三者の部品のライセンス文: $fileName（ここから）----")
    $to = [array]::IndexOf($lines, "# ---- 第三者の部品のライセンス文: $fileName（ここまで）----")
    if ($from -lt 0 -or $to -le $from) {
        return $null
    }
    $restored = New-Object System.Collections.Generic.List[string]
    for ($i = $from + 1; $i -lt $to; $i++) {
        if ([string]::Equals($lines[$i], "#", [System.StringComparison]::Ordinal)) {
            [void]$restored.Add("")
        } elseif ($lines[$i].StartsWith("# ", [System.StringComparison]::Ordinal)) {
            [void]$restored.Add($lines[$i].Substring(2))
        } else {
            # コメントでない行（印の間の空行など）が混ざっている。元の形には戻せない
            return $null
        }
    }
    return (($restored -join "`r`n") + "`r`n")
}

# フォルダの LICENSE-*.txt を、名前の順（言語・地域に依らない文字コード順）で返す。
# フォルダが無い・1 件も無い・空（空白だけ）のファイル・CR だけの行区切りがあるファイルがあるときは止める（配布物の検査 tools\check_release_package.ps1 と同じく、空のまま作らないため）
function getLicenseFiles {
    param ([string]$dir)

    if (!(Test-Path -LiteralPath $dir -PathType Container)) {
        throw "同梱するライセンス文のフォルダがありません: $dir"
    }
    $list = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($file in @(Get-ChildItem -LiteralPath $dir -Filter "LICENSE-*.txt" -File)) {
        $list.Add($file)
    }
    if ($list.Count -eq 0) {
        throw "同梱するライセンス文がありません: $dir の LICENSE-*.txt"
    }
    $list.Sort([Comparison[System.IO.FileInfo]] { param ($a, $b) [string]::CompareOrdinal($a.Name, $b.Name) })
    foreach ($file in $list) {
        assertLicenseText $file.Name ([System.IO.File]::ReadAllText($file.FullName))
    }
    return $list.ToArray()
}
