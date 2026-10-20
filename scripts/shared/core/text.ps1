# TSV とセルの文字列・ファイル名の確かめを扱う（判断層。ファイルに触らない）。

# インデックスのTSVで、セル内改行の代わりに使う文字（U+2028 LINE SEPARATOR）。
# TSVの1行 = Excelの1行を保つため、セル内改行はこの文字に置き換えて保存し、検索結果の出力時に改行へ戻す
${cellNewLine} = [string][char]0x2028

function replaceCellNewLine {
    # ダブルクォートで囲まれた1セル内の改行（LF・CR・CRLF）を、セル内改行を表す文字に置き換える
    param (
        [string]$inputString
    )

    # "" はクォート内のダブルクォートのエスケープだが、"…" "…" の2つに分けてマッチしても結果は同じ
    return [regex]::Replace($inputString, '"[^"]*"', { param($m) $m.Value -replace "\r\n|\r|\n", ${cellNewLine} })
}

function formatTsv {
    # Excelが出力したTSVを、1行目 = Excelの1行目、1列目 = A列 となるように整形する。
    # Excelは使用範囲（UsedRange）の左上のセルから出力するため、その行・列番号を firstRow・firstColumn に渡す。
    # セル内改行は置き換え、行末の空セルと末尾の空行は取り除く（途中の空行は行番号を保つため残す）。
    # 空白以外の文字が無ければ空文字を返す
    param (
        [string]$content,
        [int]$firstRow = 1,
        [int]$firstColumn = 1
    )

    $content = replaceCellNewLine $content
    $columnPadding = "`t" * ($firstColumn - 1)

    $lines = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -lt $firstRow; $i++) {
        $lines.Add("")
    }
    foreach ($line in ($content -split "\r?\n")) {
        $lines.Add(($columnPadding + $line).TrimEnd("`t"))
    }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1].Trim() -eq "") {
        $lines.RemoveAt($lines.Count - 1)
    }

    return ($lines -join "`r`n")
}

function countTsvFields {
    # TSVの1行をExcelに貼り付けたときのセル数を返す。
    # Excelと同じく、" で始まるセルは閉じる " までを1セルとする（中のタブ・改行は区切りとしない）
    param (
        [string]$line
    )

    # 先頭にタブを足し、どのセルも「タブ＋中身」で数える（^ を使うと、先頭の空のセルの後のタブを読み飛ばして1セル少なくなる）
    return [regex]::Matches("`t${line}", '\t(?:"(?:[^"]|"")*"[^\t]*|[^\t]*)').Count
}

function toColumnName {
    # 列番号を列名に変換する（1 → A、27 → AA）
    param (
        [int]$number
    )

    $name = ""
    while ($number -gt 0) {
        $number--
        $name = [string][char](65 + ($number % 26)) + $name
        $number = [math]::Floor($number / 26)
    }
    return $name
}


function splitTsvCells {
    # TSVの1行をセルに分ける。" で始まるセルは閉じる " までを1セルとし、囲みの " を外して "" を " に戻す（countTsvFields と同じ区切り方）
    param (
        [string]$line
    )

    $cells = New-Object System.Collections.Generic.List[string]
    foreach ($match in [regex]::Matches("`t${line}", '\t(?:"(?:[^"]|"")*"[^\t]*|[^\t]*)')) {
        $cell = $match.Value.Substring(1)
        if ($cell -match '^"(?<inner>(?:[^"]|"")*)"(?<rest>.*)$') {
            $cell = $Matches.inner.Replace('""', '"') + $Matches.rest
        }
        $cells.Add($cell)
    }
    return , $cells.ToArray()
}

function formatTime {
    # 当日なら HH:mm、それ以前は M/d HH:mm
    param (
        $time
    )

    if ($null -eq $time) {
        return ""
    }
    if ($time.Date -eq (Get-Date).Date) {
        return $time.ToString("H:mm")
    }
    return $time.ToString("M/d H:mm")
}

function testTextTrimmed {
    # 文字が枠に入りきらず、省略（…）で切れているか。実際の幅（枠の ActualWidth）より、要る幅（文字を 1 行で出すのに要る幅）が
    # 誤差より大きいときだけ $true。幅 0 以下（まだ配置されていない・隠れている）は切れていないものとして $false。
    # 測る側（画面層）は幅を渡すだけで、切れているかの判定はここで決める
    param (
        [double]$actualWidth,
        [double]$requiredWidth,
        [double]$tolerance = 0.5   # 測り方による小数の誤差（これ以下の超過は切れていないとみなす）
    )

    if ($actualWidth -le 0) {
        return $false
    }
    return (($requiredWidth - $actualWidth) -gt $tolerance)
}

function toSafeFileName {
    # ファイル名に使えない文字を全角に変換
    param (
        [string]$name
    )

    $name = $name.Replace(">", "＞")
    $name = $name.Replace("<", "＜")
    $name = $name.Replace("\", "￥")
    $name = $name.Replace("*", "＊")
    $name = $name.Replace('"', '”')
    $name = $name.Replace(":", "：")
    $name = $name.Replace("?", "？")
    $name = $name.Replace("|", "｜")
    $name = $name.Replace("/", "／")

    return $name
}


${maxFileNameLength} = 255  # Windows のファイル名（パスの区切りの間の1つ）の上限

# Windows で使えないファイル名（フォルダ名として使えるかの判定に使う）
${reservedFileNames} = @("CON", "PRN", "AUX", "NUL") +
    @(1..9 | ForEach-Object { "COM$_" }) + @(1..9 | ForEach-Object { "LPT$_" })

# ファイル名に使えない文字（getFolderNameProblemText とパスの区切りごとの確かめ（testPathSegments）が使う。呼ぶたびに作り直さない）
${invalidFileNameChars} = [System.IO.Path]::GetInvalidFileNameChars()

function getFolderNameProblemText {
    # 作業フォルダの下に作る 1 つのフォルダ名として使えるかを調べ、使えない理由を返す（使えれば空文字列）。
    # 空・前後の空白・長すぎる・使えない文字（\ / : など。.. のような外へ出る名前を含む）・最後の .・Windows の予約名を拒む。
    # 設定ファイルや状態ファイルから読んだ名前を使う前の確かめ（ほかの名前との重複は調べない）。
    #   allowReserved: $true なら Windows の予約名（aux.backup など）は拒まない（既にある実在のフォルダの名前を空にせず、そのまま使い続けるため。
    #                  新しく付ける名前には使わない）
    param (
        [string]$name,
        [bool]$allowReserved = $false
    )

    if ($name -eq "") {
        return "名前を入力してください。"
    }
    if ($name -ne $name.Trim()) {
        return "名前の前後に空白は使えません。"
    }
    if ($name.Length -gt ${maxFileNameLength}) {
        return "名前が長すぎます（${maxFileNameLength} 文字まで）。"
    }
    if ($name.IndexOfAny(${invalidFileNameChars}) -ge 0) {
        return "名前に使えない文字が含まれています（\ / : * ? " + [char]34 + " < > | と制御文字）。"
    }
    if ($name.EndsWith(".")) {
        return "名前の最後に . は使えません。"
    }
    if (!$allowReserved -and ${reservedFileNames} -contains $name.Split(".")[0].ToUpperInvariant()) {
        return "「${name}」は Windows で使えない名前です。"
    }
    return ""
}

function testPlainFolderName {
    # 1 つのフォルダ名として、置き場の外へ出ず、名前のとおりに作れるかだけを調べる（予約名・最後の . や空白は問わない）。
    # 空・.・..・長すぎる名前・使えない文字（\ / : など）を拒む。getFolderNameProblemText より緩く、
    # 前の版で作られた名前を、名前でパスを組まずに扱う（列挙して消す）ときの、読み飛ばすかどうかの判断に使う
    param (
        [string]$name
    )

    return !($name -eq "" -or $name -eq "." -or $name -eq ".." -or $name.Length -gt ${maxFileNameLength} -or
        $name.IndexOfAny(${invalidFileNameChars}) -ge 0)
}

# 画面に出す文字列で置き換える制御文字・向きを変える文字の並び（正規表現の [] の中。C1・DEL・U+061C・U+200E/F・U+202A-202E・U+2066-2069）。
# C0（\x00-\x1F）は使う側が前に足す。formatDisplayText・testSourceBookName・HitRow の表示が同じ並びを使う
${displayControlChars} = '\x7F-\x9F\u061C\u200E\u200F\u202A-\u202E\u2066-\u2069'


function formatDisplayText {
    # インデックスから来た文字列を画面に出すときの表示用の文字列。制御文字（C0・C1）と向きを変える文字（U+202E など）を
    # 置き換え文字（U+FFFD）にする。拡張子を偽って見せる・表示を崩すのを防ぐ。コピー・開くには元の値を使う。
    # keepLineBreaks: タブ・改行は残す（本文の行の表示用）
    param (
        $text,
        [switch]$keepLineBreaks
    )

    $pattern = if ($keepLineBreaks) {
        '[\x00-\x08\x0B\x0C\x0E-\x1F' + ${displayControlChars} + ']'
    } else {
        '[\x00-\x1F' + ${displayControlChars} + ']'
    }
    return [regex]::Replace([string]$text, $pattern, [string][char]0xFFFD)
}
