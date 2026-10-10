# TSV のファイルの読み書き（状態層）。整形の中身（formatTsv）は text.ps1（判断層）。

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
