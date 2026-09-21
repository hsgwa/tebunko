# 1ファイルを TSV に変換する（Excel は COM、Word・PowerPoint はファイルを直接読む）。

$excelMaxPath = 218       # Excelで開けるパスの長さの目安（古い版の上限）。作業フォルダのコピーのパスがこれ以上なら短い名前にする
$excelExtraCells = 1000000  # 使用範囲がデータの範囲よりこのセル数以上広いシートは、データの範囲だけを一時シートにコピーしてから書き出す

# ----------------------------------------------------------------------------
# 変換
# ----------------------------------------------------------------------------

function copyDataRangeToTempSheet {
    # 使用範囲（UsedRange）が実際のデータよりずっと広いシートのために、データのある範囲だけを
    # 同じブックの一時シートへ同じ位置でコピーし、その一時シートを返す。縮める必要が無い・できない場合は $null を返す。
    #
    # 最終行・右端のセルに書式だけが残っていると使用範囲がシート全体になり、Excelのテキスト保存が
    # 空セルのタブだけで数GBを書き出して制限時間も超える（実測: 10分で6GB、終わらない）。
    # 元のシートの行・列は削除しない（そこを参照する数式が #REF! になり、ほかのシートの表示値まで変わるため）。
    param (
        $workbook,
        $worksheet
    )

    $usedRange = $worksheet.UsedRange
    try {
        $firstRow = $usedRange.Row
        $firstColumn = $usedRange.Column
        $usedLastRow = $firstRow + $usedRange.Rows.Count - 1
        $usedLastColumn = $firstColumn + $usedRange.Columns.Count - 1
    } finally {
        releaseComObject $usedRange
    }

    # 使用範囲そのものが縮める基準より狭ければ、データの範囲を探すまでもない（ほとんどのシート）。
    # 大きいシートでは Find に時間がかかるため、シートごとに呼ばないようにする
    $usedCells = [double]($usedLastRow - $firstRow + 1) * ($usedLastColumn - $firstColumn + 1)
    if ($usedCells -lt $excelExtraCells) {
        return $null
    }

    # 値・数式のある最後の行・列を探す（書式だけのセルには当たらない）。
    # 引数: What, After, LookIn（-4123 = xlFormulas）, LookAt, SearchOrder（1 = xlByRows / 2 = xlByColumns）, SearchDirection（2 = xlPrevious）
    $topLeft = $worksheet.Range("A1")
    $cells = $worksheet.Cells
    try {
        $found = $cells.Find("*", $topLeft, -4123, [Type]::Missing, 1, 2)
        $dataLastRow = if ($found) { $found.Row } else { 0 }
        $found = $cells.Find("*", $topLeft, -4123, [Type]::Missing, 2, 2)
        $dataLastColumn = if ($found) { $found.Column } else { 0 }
    } finally {
        releaseComObject $cells
        releaseComObject $topLeft
    }

    # 値のあるセルが無い（書式だけのシート）場合は、そのまま書き出して空のTSVとして扱う
    if ($dataLastRow -lt $firstRow -or $dataLastColumn -lt $firstColumn) {
        return $null
    }

    # セル数は int を超えるため double で数える（シート全体は約 172 億セル）
    $dataCells = [double]($dataLastRow - $firstRow + 1) * ($dataLastColumn - $firstColumn + 1)
    if (($usedCells - $dataCells) -lt $excelExtraCells) {
        return $null
    }

    # 数式の参照先がずれないよう、コピー先は元と同じ位置（行・列）にする
    $temp = $null
    try {
        $temp = $workbook.Worksheets.Add()
        $source = $worksheet.Range($worksheet.Cells.Item($firstRow, $firstColumn), $worksheet.Cells.Item($dataLastRow, $dataLastColumn))
        try {
            [void]$source.Copy($temp.Cells.Item($firstRow, $firstColumn))
        } finally {
            releaseComObject $source
        }
        return $temp
    } catch {
        # ブックの構成が保護されている場合など。元のシートをそのまま書き出す（時間切れで失敗することがある）
        Write-Host "    $($worksheet.Name) の使用範囲を縮められませんでした: $($_.Exception.Message)" -ForegroundColor Yellow
        if ($temp) {
            $temp.Delete()
            releaseComObject $temp
        }
        return $null
    }
}

function extractWorkbook {
    # Excelファイルをシートごとに作業フォルダへTSV出力し、出力したシート数を返す
    param (
        [string]$sourcePath
    )

    $bookName = [System.IO.Path]::GetFileName($sourcePath)

    # 元のファイルを占有しないよう、作業フォルダにコピーしてからコピーを開く
    # （Excelで開いている間、元のファイルを利用者が上書き保存・移動できなくなるのを防ぐ。長いパスのファイルも開ける）。
    # ファイル名を参照する数式（CELL("filename") 等）の表示値が変わらないよう、コピーは元と同じファイル名にする。
    # 作業フォルダ＋ファイル名が長すぎてExcelで開けない場合だけ、短い名前にする。
    # コピーはブックを閉じた後に削除する（開けずに例外になった場合は、次のファイルの変換前・終了時に作業フォルダごと空にする）
    $copyPath = Join-Path $tmpDir $bookName
    if ($copyPath.Length -ge $excelMaxPath) {
        $copyPath = Join-Path $tmpDir ("source" + [System.IO.Path]::GetExtension($sourcePath))
    }
    copyFileShared $sourcePath $copyPath
    $openPath = $copyPath

    # 読み取り専用・リンク更新なしで開く。
    # パスワード付きのファイルは、ダイアログを出さずにエラーとするためダミーのパスワードを渡す
    $workbooks = (getApp "Excel").Workbooks
    try {
        $wb = $workbooks.Open($openPath, 0, $true, [Type]::Missing, "dummy", "dummy", $true)
    } finally {
        releaseComObject $workbooks
    }

    $sheets = New-Object System.Collections.Generic.List[object]
    try {
        $worksheets = $wb.Worksheets
        foreach ($ws in @($worksheets)) {
            try {
                # 非表示シートは除外（-1 = xlSheetVisible。0 = 非表示、2 = 完全に非表示）
                if ($ws.Visible -ne -1) {
                    continue
                }

                # Excelは [ ] を含むパスに保存できないため、一時ファイル名で保存し、整形時にリネームする
                $tmpPath = Join-Path $tmpDir ("sheet{0}.tmp" -f $ws.Index)
                $tsvPath = Join-Path $tmpDir (toIndexFileName $ws.Name)

                # 使用範囲が実際のデータよりずっと広いシートは、データの範囲だけを一時シートにコピーしてから書き出す
                $temp = copyDataRangeToTempSheet $wb $ws
                $target = if ($temp) { $temp } else { $ws }
                try {
                    # Excelは使用範囲（UsedRange）の左上のセルから出力するため、整形時にA1からの位置に戻せるよう控えておく。
                    # 一時シートは書式だけのセルが無く、元のシートより使用範囲が狭いことがあるため、書き出すシートから取る
                    $usedRange = $target.UsedRange
                    try {
                        $firstRow = $usedRange.Row
                        $firstColumn = $usedRange.Column
                    } finally {
                        releaseComObject $usedRange
                    }

                    [void]$target.Activate()
                    [void]$target.SaveAs($tmpPath, 42)  # 42 = xlUnicodeText
                } finally {
                    if ($temp) {
                        # 一時シートは、次のシートの Index がずれないようすぐ削除する（ブックは保存せずに閉じる）
                        $temp.Delete()
                        releaseComObject $temp
                    }
                }
                $sheets.Add(@($tmpPath, $tsvPath, $firstRow, $firstColumn))
            } finally {
                releaseComObject $ws
            }
        }
        releaseComObject $worksheets
    } finally {
        # 保存したファイルはブックを閉じるまでロックされている
        $wb.Close($false)
        releaseComObject $wb
        Remove-Item -LiteralPath $copyPath -Force
    }

    $count = 0
    foreach ($sheet in $sheets) {
        if (prettyTsv $sheet[0] $sheet[1] $sheet[2] $sheet[3]) {
            $count++
        }
        Remove-Item -LiteralPath $sheet[0] -Force
    }
    return $count
}

function extractWithWord {
    # Wordで開き、.docx 形式で保存する（旧形式 .doc 等を読めるようにするため）
    param (
        [string]$sourcePath,
        [string]$destPath
    )

    # 読み取り専用で開く。パスワード付きのファイルは、ダイアログを出さずにエラーとするためダミーのパスワードを渡す
    #   引数: FileName, ConfirmIndexings, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
    #         Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible
    $documents = (getApp "Word").Documents
    try {
        $doc = $documents.Open($sourcePath, $false, $true, $false, "dummy", "dummy", $false, "dummy", "dummy", [Type]::Missing, [Type]::Missing, $false)
    } finally {
        releaseComObject $documents
    }

    try {
        # 保存時のページ区切り（ページ番号の目安）を正しく記録させるため、ページ割りを確定させてから保存する
        $doc.Repaginate()
        [void]$doc.SaveAs2($destPath, 12)  # 12 = wdFormatXMLDocument（.docx）
    } finally {
        $doc.Close(0)
        releaseComObject $doc
    }
}

function extractWithPowerPoint {
    # PowerPointで開き、.pptx 形式で保存する（旧形式 .ppt 等を読めるようにするため）
    param (
        [string]$sourcePath,
        [string]$destPath
    )

    # 読み取り専用・ウィンドウ無しで開く。
    # ファイル名の後ろに "::<パスワード>::" を付けると、パスワード付きのファイルはダイアログを出さずにエラーになる
    $presentations = (getApp "PowerPoint").Presentations
    try {
        $pres = $presentations.Open("${sourcePath}::dummy::", -1, 0, 0)  # ReadOnly, Untitled = False, WithWindow = False
    } finally {
        releaseComObject $presentations
    }

    try {
        [void]$pres.SaveAs($destPath, 24)  # 24 = ppSaveAsOpenXMLPresentation（.pptx）
    } finally {
        $pres.Close()
        releaseComObject $pres
    }
}

function extractDocument {
    # Word・PowerPointのファイルを場所（ページ・スライド）ごとに作業フォルダへTSV出力し、出力した数を返す
    param (
        [string]$sourcePath
    )

    $isWord = ((getAppName $sourcePath) -eq "Word")

    # 元のファイルを占有しないよう、作業フォルダにコピーしてからコピーを読む（読んでいる間も、利用者が上書き保存・移動できる）。
    # 新形式（ZIP）はコピーをそのまま読む。
    # 旧形式・パスワード付き・拡張子と中身が異なるファイルは、Word・PowerPointで新形式に変換してから読む
    $copyPath = Join-Path $tmpDir ("source" + [System.IO.Path]::GetExtension($sourcePath))
    $readPath = $copyPath
    $workFiles = @($copyPath)
    try {
        copyFileShared $sourcePath $copyPath
        if (!(isZipFile $copyPath)) {
            # PowerPointは、プレゼンテーションではないファイル（中身がテキスト等）もアウトラインとして開き、
            # 文字化けした内容になるため、旧形式（複合ドキュメント形式）でなければ変換しない。
            # Wordはテキスト・HTML・RTFも正しく読めるため、そのまま Word で開く
            if (!$isWord -and !(isCompoundFile $copyPath)) {
                throw "ファイルが壊れているか、PowerPointのファイルではありません（新形式（ZIP）でも旧形式でもない内容です）。"
            }

            # Word・PowerPointは拡張子と中身が異なるファイル（中身が .doc の .docx 等）を開けないため、
            # コピーに旧形式の拡張子を付け直してから開く
            if ($isWord) {
                $legacyPath = Join-Path $tmpDir "source.doc"
                $readPath = Join-Path $tmpDir "ingested.docx"
            } else {
                $legacyPath = Join-Path $tmpDir "source.ppt"
                $readPath = Join-Path $tmpDir "ingested.pptx"
            }
            if ($legacyPath -ne $copyPath) {
                [System.IO.File]::Move($copyPath, $legacyPath)
                $copyPath = $legacyPath
            }
            $workFiles = @($copyPath, $readPath)

            if ($isWord) {
                extractWithWord $copyPath $readPath
            } else {
                extractWithPowerPoint $copyPath $readPath
            }
        }

        if ($isWord) {
            $units = readDocxUnits $readPath
        } else {
            $units = readPptxUnits $readPath
        }
    } finally {
        foreach ($path in $workFiles) {
            if (Test-Path -LiteralPath $path) {
                Remove-Item -LiteralPath $path -Force
            }
        }
    }

    return (writeUnits $units $tmpDir)
}

function ingestFile {
    # 1ファイルを変換し、作成したTSVの数を返す
    param (
        [string]$sourcePath
    )

    if ((getAppName $sourcePath) -eq "Excel") {
        return (extractWorkbook $sourcePath)
    }
    return (extractDocument $sourcePath)
}
