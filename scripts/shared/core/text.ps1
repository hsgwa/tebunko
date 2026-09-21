# TSV とセルの文字列を扱う（判断層。ファイルに触らない）。

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

function prettyTsv {
    # Excelが出力したTSV（UTF-16）を整形してUTF-8で保存する。内容が空なら保存せず $false を返す
    #   firstRow, firstColumn: Excelが出力した範囲の左上のセルの行・列番号
    param (
        [string]$inputFilePath,
        [string]$outputFilePath,
        [int]$firstRow = 1,
        [int]$firstColumn = 1
    )

    $content = formatTsv ([System.IO.File]::ReadAllText((toLongPath $inputFilePath))) $firstRow $firstColumn
    if ($content -eq "") {
        return $false
    }

    [System.IO.File]::WriteAllText((toLongPath $outputFilePath), "${content}`r`n", ${utf8Bom})
    return $true
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

function readTsvContext {
    # インデックスのTSVの lineNumber 行目と、その前後 before 行・after 行を @{ LineNumber; Line } の配列で返す（画面の選択行のプレビュー）。
    # 行の数え方は検索（searchTsvFiles）と同じ（どちらも StreamReader.ReadLine で数えるため一致する。Excel のTSVでは行番号 = シートの行番号）。
    # ファイルが無い・読めない場合は空。インデックス作成中のTSVも読めるよう共有モードは ReadWrite|Delete。
    # ※以前は C#（TsvContextReader）で行の位置を覚えて速くしていたが、実行時コンパイル（csc.exe）を無くすため PowerShell で読む
    #   （プレビューは選択行の前後だけで、TSV は元のファイル1つ分＝通常は数千行までのため、先頭から目的行までの読み込みで十分）。
    param (
        [string]$path,
        [int]$lineNumber,
        [int]$before = 3,
        [int]$after = 3
    )

    $rows = New-Object System.Collections.Generic.List[psobject]
    $first = [Math]::Max(1, $lineNumber - $before)
    $last = $lineNumber + $after
    if ($last -lt $first) { return @() }

    $reader = $null
    try {
        $stream = New-Object System.IO.FileStream((toLongPath $path), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        $number = 0
        while ($null -ne ($line = $reader.ReadLine())) {
            $number++
            if ($number -lt $first) { continue }
            if ($number -gt $last) { break }
            $rows.Add([pscustomobject]@{ LineNumber = $number; Line = $line })
        }
    } catch [System.IO.IOException] {
        $rows.Clear()
    } catch [System.UnauthorizedAccessException] {
        $rows.Clear()
    } finally {
        if ($reader) { $reader.Dispose() }
    }
    # 呼び出し側で @() にして使う
    return $rows.ToArray()
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
