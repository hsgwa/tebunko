# CSV/TSV の書き出し（RFC 4180 のセルの囲み・見出し行つきの本文ファイルの作成）。ツールを知らない。

function toCsvLine {
    # セルを区切り文字でつないだ 1 行（行末の改行は付けない）を返す。
    # 区切り文字・"・CR・LF を含むセルだけ " で囲み、中の " は "" にする。値は加工しない。$null は空
    param (
        [string[]]$cells,
        [string]$delimiter = ","
    )

    if ($null -eq $cells -or $cells.Count -eq 0) {
        return ""
    }
    $parts = foreach ($cell in $cells) {
        if ($null -eq $cell) {
            ""
        } elseif ($cell.IndexOfAny([char[]]@('"', "`r", "`n")) -ge 0 -or $cell.Contains($delimiter)) {
            '"' + $cell.Replace('"', '""') + '"'
        } else {
            $cell
        }
    }
    return ($parts -join $delimiter)
}

function writeCsvFile {
    # 本文ファイル（toCsvLine の行を CRLF 区切りで書いた UTF-8（BOM なし））の前に見出し行を付け、
    # BOM 付き UTF-8・CRLF の newPath に書く。本文は流し込む（全体をメモリに読まない）。
    # 例外（IOException・UnauthorizedAccessException）は呼ぶ側が分類できるよう、そのまま投げる
    param (
        [string]$bodyPath,
        [string]$newPath,
        [string[]]$headers,
        [string]$delimiter = ","
    )

    $encoding = New-Object System.Text.UTF8Encoding($true)
    $head = $encoding.GetPreamble() + $encoding.GetBytes((toCsvLine $headers $delimiter) + "`r`n")
    $in = $null
    $out = $null
    try {
        $in = [System.IO.FileStream]::new((toLongPath $bodyPath), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $out = [System.IO.FileStream]::new((toLongPath $newPath), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $out.Write($head, 0, $head.Length)
        $in.CopyTo($out)
    } finally {
        if ($null -ne $out) { $out.Dispose() }
        if ($null -ne $in) { $in.Dispose() }
    }
}
