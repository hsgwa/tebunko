# 1ファイルから文字を抽出して TSV に書き出す（Excel のセルは COM、Excel の図形・コメントと Word・PowerPoint はファイルを直接読む）。
# $excelMaxPath（Excelで開けるパスの長さの目安）は paths.ps1 で定義する（selectTmpDir でも使うため）

$excelExtraCells = 1000000  # 使用範囲がデータの範囲よりこのセル数以上広いシートは、データの範囲だけを一時シートにコピーしてから書き出す

# ----------------------------------------------------------------------------
# 暗号化されたファイルの扱い（office_protection.ps1・office_protection_view.ps1）
# ----------------------------------------------------------------------------

# 「形式の分からないバイナリ」（透過暗号化の製品の暗号文の見込み）をOfficeで開く予備を、アプリごとに使うかどうか。
# 実機で数秒以内に例外にならない・ダイアログが出るアプリは $false にする（docs/design/indexing/known-issues.md）
${officeFallbackEnabled} = @{ Word = $true; Excel = $true; PowerPoint = $true }

function testOfficeFallbackEnabled {
    param (
        [string]$appName
    )

    return [bool]${officeFallbackEnabled}[$appName]
}

function throwProtectionFailure {
    # 暗号化の判定で、Officeにまだ一度も触れていない状態の失敗にする。
    # invokeIngestTask（indexer_run.ps1）が、この印（Data の OfficeUntouched）を見て stopApp を呼ばずに済ませる
    # （IRM・パスワード付きのファイルが並んでも、そのたびにOfficeを起動し直さないため）
    param (
        [string]$message
    )

    $exception = New-Object System.Management.Automation.RuntimeException($message)
    $exception.Data["OfficeUntouched"] = $true
    throw $exception
}

function isOfficeRequiredException {
    # invokeIngestTask（indexer_run.ps1）と見分け方をそろえる（重複させない）
    param (
        [System.Exception]$exception
    )

    $base = $exception.GetBaseException()
    return ($base -is [System.OperationCanceledException] -and $base.Message -eq ${officeRequiredMessage})
}

function isOfficeAppInUseException {
    # invokeIngestTask（indexer_run.ps1）と見分け方をそろえる（重複させない）
    param (
        [System.Exception]$exception
    )

    $base = $exception.GetBaseException()
    return ($base -is [System.InvalidOperationException] -and $base.Message.EndsWith(${officeAppInUseMessage}))
}

function isPassthroughIngestException {
    # 見張りの時間切れ・「Officeが要る」・「利用者のOfficeが使用中」の例外は、暗号化の失敗として包み直さずそのまま通す
    param (
        [System.Exception]$exception
    )

    if ($script:watchdog.TimedOut) {
        return $true
    }
    if (isOfficeRequiredException $exception) {
        return $true
    }
    if (isOfficeAppInUseException $exception) {
        return $true
    }
    return $false
}

function failProtectionCheck {
    # 暗号化の判定で、Officeを開く前に作業フォルダのコピーを消してから失敗にする（throwProtectionFailure）
    param (
        [string]$copyPath,
        [string]$message
    )

    Remove-Item -LiteralPath $copyPath -Force
    throwProtectionFailure $message
}

function logProtectionKind {
    # 暗号化の判定の結果をインデックス作成ログに書く（Password・Rights・Unknownだけ。旧形式・ふつうのテキストは書かない）
    param (
        [string]$kind
    )

    if ($kind -eq "Password" -or $kind -eq "Rights" -or $kind -eq "Unknown") {
        writeIndexerLog "    暗号化の判定: $kind" "Yellow"
    }
}

# ----------------------------------------------------------------------------
# 抽出
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
        writeIndexerLog "    $($worksheet.Name) の使用範囲を縮められませんでした: $($_.Exception.Message)" "Yellow"
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
    # コピーはブックを閉じた後に削除する（開けずに例外になった場合は、次のファイルの取り込み前・終了時に作業フォルダごと空にする）
    $copyPath = Join-Path $tmpDir $bookName
    if ($copyPath.Length -ge $excelMaxPath) {
        $copyPath = Join-Path $tmpDir ("source" + [System.IO.Path]::GetExtension($sourcePath))
    }
    copyFileShared $sourcePath $copyPath
    $openPath = $copyPath

    # 図形・コメントの文字は、テキスト保存には出ないため、新形式（ZIP）のブックを直接読む（Excel より速い）。
    # 旧形式（.xls）・パスワード付きのブックは ZIP ではないため読まない。
    # 読めなくてもセルの値は取り込めるため、インデックス作成ログに記録して続ける
    $objectUnits = $null
    $protection = $null
    if (isZipFile $copyPath) {
        $chartFailures = New-Object System.Collections.Generic.List[string]
        try {
            $objectUnits = readXlsxObjectUnits $copyPath $chartFailures
        } catch {
            writeIndexerLog "    図形・コメントを読み取れませんでした: $($_.Exception.Message)" "Yellow"
        }
        # 1つのグラフ・SmartArtが読めなくても、そこだけを空にしてほかの図形・コメントは読む（readXlsxObjectUnits）。
        # shared/ はツールを知らないため、読めなかった部品の名前をここでログに書く
        foreach ($failure in $chartFailures) {
            writeIndexerLog "    グラフ・SmartArt を読み取れませんでした: $failure" "Yellow"
        }
    } else {
        # IRM・秘密度ラベルの暗号化は、Excelを起動せずに失敗にする（サインイン画面を防ぐ）。
        # パスワード付き（既定のパスワードで暗号化されたブックを含む）は、今までどおりExcelに任せる。
        # 「形式の分からないバイナリ」（透過暗号化の製品の暗号文の見込み）は、アプリごとの切り替えが $false なら
        # Excelを起動せずに失敗にする
        $protection = getOfficeFileProtection $copyPath
        logProtectionKind $protection
        if ($protection -eq "Rights") {
            failProtectionCheck $copyPath (getProtectionFailureText "Rights")
        }
        if ($protection -eq "Unknown" -and !(testOfficeFallbackEnabled "Excel")) {
            failProtectionCheck $copyPath (getProtectionFailureText "Unknown")
        }
    }

    # 読み取り専用・リンク更新なしで開く。
    # パスワード付きのファイルは、ダイアログを出さずにエラーとするためダミーのパスワードを渡す。
    # 「形式の分からないバイナリ」で開けなかったときは、元の例外をログに書いて文言を言い換える
    # （見張りの時間切れ・「Officeが要る」・「利用者のOfficeが使用中」の例外はそのまま通す）
    $workbooks = (getApp "Excel").Workbooks
    try {
        try {
            $wb = $workbooks.Open($openPath, 0, $true, [Type]::Missing, "dummy", "dummy", $true)
        } catch {
            if ($protection -ne "Unknown" -or (isPassthroughIngestException $_.Exception)) {
                throw
            }
            writeIndexerLog "    予備の読み取りに失敗しました: $(describeIngestError $_.Exception)" "Yellow"
            throw (getProtectionFailureText "Unknown")
        }
    } finally {
        releaseComObject $workbooks
    }

    # 「形式の分からないバイナリ」は、開けた後のブックの形式（テキスト・HTML・CSVでない）も確かめる。
    # ほかの種類（HTML の .xls など）には当てない（本来読めているものまで失敗にしないため）
    if ($protection -eq "Unknown" -and !(testWorkbookFormat $wb.FileFormat)) {
        $wb.Close($false)
        releaseComObject $wb
        Remove-Item -LiteralPath $copyPath -Force
        throw (getProtectionFailureText "Unknown")
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
        # 透過暗号化の製品が、Excelの保存したこの一時ファイルまで暗号化していないかを確かめる。
        # ブックを閉じた後（このループ）でないと、保存したファイルはExcelがロックしていて読めない
        if (!(testOfficeOutput (readFileHead $sheet[0]) "UnicodeText")) {
            throw ${officeOutputEncryptedMessage}
        }
        if (prettyTsv $sheet[0] $sheet[1] $sheet[2] $sheet[3]) {
            $count++
        }
        Remove-Item -LiteralPath $sheet[0] -Force
    }
    if ($null -ne $objectUnits) {
        $count += writeUnits $objectUnits $tmpDir
    }
    return $count
}

function extractWithWord {
    # Wordで開き、.docx 形式で保存する（旧形式 .doc 等を読めるようにするため）。
    # $format を渡すと、自動判定に任せずその形式で開く（「形式の分からないバイナリ」を開くときだけ使う。
    # 文字コードを選ぶダイアログや、暗号文をテキストとして読むことを防ぐ）
    param (
        [string]$sourcePath,
        [string]$destPath,
        $format = [Type]::Missing
    )

    # 読み取り専用で開く。パスワード付きのファイルは、ダイアログを出さずにエラーとするためダミーのパスワードを渡す
    #   引数: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
    #         Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible
    $documents = (getApp "Word").Documents
    try {
        $doc = $documents.Open($sourcePath, $false, $true, $false, "dummy", "dummy", $false, "dummy", "dummy", $format, [Type]::Missing, $false)
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

# 読み取りのスレッド（Office を持たない）なら $true。Office が要るファイルは「Office が要る」の例外にする
$script:officeUnavailable = $false
# Office が要るときの例外の文言（invokeIngestTask が見分けて、司令に回し直しを頼む）
${officeRequiredMessage} = "このファイルの取り込みには Word・PowerPoint が要ります。"

function extractWithOffice {
    # Word・PowerPointどちらかで新形式に変換する（呼び分けをまとめる）。
    # $wordFormat を渡すと、Wordはその形式に固定して開く（「形式の分からないバイナリ」のときだけ渡す）
    param (
        [bool]$isWord,
        [string]$copyPath,
        [string]$readPath,
        $wordFormat = [Type]::Missing
    )

    if ($isWord) {
        extractWithWord $copyPath $readPath $wordFormat
    } else {
        extractWithPowerPoint $copyPath $readPath
    }
}

function verifyOfficeOutputIsZip {
    # Officeが保存した出力（$readPath）の先頭がZIPかを確かめる。
    # 透過暗号化の製品がこの一時ファイルまで暗号化していないか（ふつうのファイルでも起こりうる）
    param (
        [string]$readPath
    )

    if (!(testOfficeOutput (readFileHead $readPath) "Zip")) {
        throw ${officeOutputEncryptedMessage}
    }
}

function extractDocument {
    # Word・PowerPointのファイルを場所（ページ・スライド）ごとに作業フォルダへTSV出力し、出力した数を返す
    param (
        [string]$sourcePath
    )

    $isWord = ((getAppName $sourcePath) -eq "Word")
    $appName = $(if ($isWord) { "Word" } else { "PowerPoint" })

    # 元のファイルを占有しないよう、作業フォルダにコピーしてからコピーを読む（読んでいる間も、利用者が上書き保存・移動できる）。
    # 新形式（ZIP）はコピーをそのまま読む。
    # 旧形式・パスワード付き・拡張子と中身が異なるファイルは、Word・PowerPointで新形式に変換してから読む
    $copyPath = Join-Path $tmpDir ("source" + [System.IO.Path]::GetExtension($sourcePath))
    $readPath = $copyPath
    $workFiles = @($copyPath)
    try {
        copyFileShared $sourcePath $copyPath
        if (!(isZipFile $copyPath)) {
            $protection = getOfficeFileProtection $copyPath
            logProtectionKind $protection

            # パスワード付き（新形式）・IRM・秘密度ラベルの暗号化は、Office を使わずに失敗にする
            # （Word・PowerPointはサインイン画面が出ないようにするため。Excelは別処理で今のまま）
            if ($protection -eq "Password" -or $protection -eq "Rights") {
                throwProtectionFailure (getProtectionFailureText $protection)
            }

            if ($protection -eq "Unknown") {
                # 「形式の分からないバイナリ」（透過暗号化の製品の暗号文の見込み）。
                # アプリごとの切り替えが $false なら、Officeを使わずに失敗にする
                if (!(testOfficeFallbackEnabled $appName)) {
                    throwProtectionFailure (getProtectionFailureText "Unknown")
                }
                # 読み取りのスレッド（Office を持たない）では、Office のレーンに回す
                if ($script:officeUnavailable) {
                    throw (New-Object System.OperationCanceledException ${officeRequiredMessage})
                }

                # 拡張子は元のまま変えない（付け替えると、Wordが文字コードを尋ねる・PowerPointがアウトラインとして
                # 開くことがあるため）。Wordは形式を拡張子に合わせて固定し、自動判定に任せない
                $readPath = Join-Path $tmpDir $(if ($isWord) { "converted.docx" } else { "converted.pptx" })
                $workFiles = @($copyPath, $readPath)
                $wordFormat = getWordOpenFormat ([System.IO.Path]::GetExtension($copyPath))
                try {
                    extractWithOffice $isWord $copyPath $readPath $wordFormat
                } catch {
                    if (isPassthroughIngestException $_.Exception) {
                        throw
                    }
                    writeIndexerLog "    予備の読み取りに失敗しました: $(describeIngestError $_.Exception)" "Yellow"
                    throw (getProtectionFailureText "Unknown")
                }
                verifyOfficeOutputIsZip $readPath
            } else {
                # Legacy（旧形式・権限保護と分からないCFB）・Text（空・RTF・HTML・UTF-8・UTF-16等）は今と同じ

                # PowerPointは、プレゼンテーションではないファイル（中身がテキスト等）もアウトラインとして開き、
                # 文字化けした内容になるため、旧形式（複合ドキュメント形式）でなければ取り込まない。
                # Wordはテキスト・HTML・RTFも正しく読めるため、そのまま Word で開く
                if (!$isWord -and !(isCompoundFile $copyPath)) {
                    throw "ファイルが壊れているか、PowerPointのファイルではありません（新形式（ZIP）でも旧形式でもない内容です）。"
                }

                # 読み取りのスレッド（Office を持たない）では、Office のレーンに回す（indexer_run.ps1 の runIngestWorker・invokeIngestTask）
                if ($script:officeUnavailable) {
                    throw (New-Object System.OperationCanceledException ${officeRequiredMessage})
                }

                # Word・PowerPointは拡張子と中身が異なるファイル（中身が .doc の .docx 等）を開けないため、
                # コピーに旧形式の拡張子を付け直してから開く
                if ($isWord) {
                    $legacyPath = Join-Path $tmpDir "source.doc"
                    $readPath = Join-Path $tmpDir "converted.docx"
                } else {
                    $legacyPath = Join-Path $tmpDir "source.ppt"
                    $readPath = Join-Path $tmpDir "converted.pptx"
                }
                if ($legacyPath -ne $copyPath) {
                    [System.IO.File]::Move($copyPath, $legacyPath)
                    $copyPath = $legacyPath
                }
                $workFiles = @($copyPath, $readPath)

                extractWithOffice $isWord $copyPath $readPath
                verifyOfficeOutputIsZip $readPath
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
    # 1ファイルを取り込み、作成したTSVの数を返す。
    # テキストの拡張子（textExtensions）は Office を使わず読み取りのレーンで取り込む（extract_text.ps1）
    param (
        [string]$sourcePath
    )

    if (testTextExtension $sourcePath) {
        return (extractTextFile $sourcePath $tmpDir)
    }
    if ((getAppName $sourcePath) -eq "Excel") {
        return (extractWorkbook $sourcePath)
    }
    return (extractDocument $sourcePath)
}
