# Excel → TSV 変換
#
# config\変換対象フォルダパス.txt に記載したフォルダ配下のExcelファイルを
# シートごとにTSVへ変換し、work\index に保存する。
#   work\index\<相対フォルダ>\<ファイル名>.<拡張子>_<シート名>.tsv
#
# ・変換済みで更新の無いファイルはスキップする
# ・Ctrl+C で中断した場合、次回実行時は未処理のファイルから再開する
# ・変換に失敗したファイルは work\変換失敗一覧.txt に記録する

. "$PSScriptRoot\common.ps1"

$ErrorActionPreference = "Stop"

trap {
    Write-Host "＜エラー＞"
    Write-Host $_.Exception.Message -ForegroundColor Red
    pause
    exit 1
}

$targetExtensions = @(".xlsx", ".xlsm", ".xls", ".xlsb")
$restartInterval = 50  # Excelを再起動する間隔（ファイル数）。メモリ肥大化対策

if (-not ("WinGrep.User32" -as [type])) {
    Add-Type -Namespace WinGrep -Name User32 -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
'@
}

# ----------------------------------------------------------------------------
# 変換対象
# ----------------------------------------------------------------------------

function getOutputDir {
    # Excelファイルの相対パスから、変換結果の出力先フォルダを返す
    param (
        [string]$relPath
    )

    $relDir = [System.IO.Path]::GetDirectoryName($relPath)
    if ($relDir) {
        return (Join-Path $indexDir $relDir)
    }
    return $indexDir
}

function getIndexFiles {
    # 出力先フォルダにある、指定したExcelファイルの変換結果を返す
    param (
        [string]$outDir,
        [string]$bookName
    )

    if (!(Test-Path -LiteralPath $outDir)) {
        return @()
    }
    # -Filter は [ ] をワイルドカードとして扱わない
    return @(Get-ChildItem -LiteralPath $outDir -Filter "${bookName}_*.tsv" -File)
}

function createTargetList {
    # 変換対象フォルダ配下のExcelファイルのうち、未変換・更新ありのものの相対パスを返す
    param (
        [string]$targetFolder
    )

    $files = @(Get-ChildItem -LiteralPath $targetFolder -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { ($targetExtensions -contains $_.Extension.ToLower()) -and -not $_.Name.StartsWith('~$') })

    $targets = New-Object System.Collections.Generic.List[string]
    foreach ($file in $files) {
        $relPath = $file.FullName.Substring($targetFolder.Length).TrimStart("\")
        $latest = getIndexFiles (getOutputDir $relPath) $file.Name | Sort-Object LastWriteTime | Select-Object -Last 1
        if ($latest -and $latest.LastWriteTime -ge $file.LastWriteTime) {
            continue
        }
        $targets.Add($relPath)
    }

    Write-Host ("Excelファイル {0} 件中、変換済み {1} 件をスキップします。" -f $files.Count, ($files.Count - $targets.Count))
    return $targets.ToArray()
}

# ----------------------------------------------------------------------------
# Excel操作
# ----------------------------------------------------------------------------

$script:excel = $null
$script:excelPid = 0

function startExcel {
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $false
    $script:excel.DisplayAlerts = $false
    $script:excel.EnableEvents = $false
    $script:excel.ScreenUpdating = $false
    $script:excel.AskToUpdateLinks = $false
    $script:excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable（マクロ無効）

    # 終了できなかった場合に強制終了するため、プロセスIDを控えておく
    $processId = [uint32]0
    [void][WinGrep.User32]::GetWindowThreadProcessId([IntPtr]$script:excel.Hwnd, [ref]$processId)
    $script:excelPid = $processId
}

function stopExcel {
    if ($script:excel) {
        try { $script:excel.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:excel) } catch {}
        $script:excel = $null
    }

    # GC::WaitForPendingFinalizers() はCOMの解放待ちで長時間（約60秒）止まることがあるため使わず、
    # 終了しなかったExcelはプロセスIDを指定して強制終了する
    if ($script:excelPid) {
        $process = Get-Process -Id $script:excelPid -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit(5000)) {
            $process.Kill()
        }
        $script:excelPid = 0
    }
}

function releaseComObject($object) {
    if ($null -ne $object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object)
    }
}

function convertWorkbook {
    # Excelファイルをシートごとに作業フォルダへTSV出力し、出力したシート数を返す
    param (
        [string]$sourcePath
    )

    $bookName = [System.IO.Path]::GetFileName($sourcePath)

    # 読み取り専用・リンク更新なしで開く。
    # パスワード付きのファイルは、ダイアログを出さずにエラーとするためダミーのパスワードを渡す
    $workbooks = $script:excel.Workbooks
    try {
        $wb = $workbooks.Open($sourcePath, 0, $true, [Type]::Missing, "dummy", "dummy", $true)
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
                $tsvPath = Join-Path $tmpDir ("{0}_{1}.tsv" -f $bookName, (toSafeFileName $ws.Name))

                [void]$ws.Activate()
                [void]$ws.SaveAs($tmpPath, 42)  # 42 = xlUnicodeText
                $sheets.Add(@($tmpPath, $tsvPath))
            } finally {
                releaseComObject $ws
            }
        }
        releaseComObject $worksheets
    } finally {
        # 保存したファイルはブックを閉じるまでロックされている
        $wb.Close($false)
        releaseComObject $wb
    }

    $count = 0
    foreach ($sheet in $sheets) {
        if (prettyTsv $sheet[0] $sheet[1]) {
            $count++
        }
        Remove-Item -LiteralPath $sheet[0] -Force
    }
    return $count
}

function publishTsv {
    # 作業フォルダのTSVを出力先フォルダへ移動する
    param (
        [string]$bookName,
        [string]$outDir
    )

    [System.IO.Directory]::CreateDirectory($outDir) | Out-Null

    # 以前の変換結果を削除（シートの削除・名前変更に追従するため）
    getIndexFiles $outDir $bookName | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }

    foreach ($file in @(Get-ChildItem -LiteralPath $tmpDir -Filter "*.tsv" -File)) {
        [System.IO.File]::Move($file.FullName, (Join-Path $outDir $file.Name))
    }
}

function clearTmpDir {
    Get-ChildItem -LiteralPath $tmpDir -File | Remove-Item -Force
}

# ----------------------------------------------------------------------------
# メイン処理
# ----------------------------------------------------------------------------

$targetFolder = getTargetFolder
if (!(Test-Path -LiteralPath $targetFolder -PathType Container)) {
    throw "変換対象フォルダが見つかりません: ${targetFolder}"
}
$targetFolder = (Resolve-Path -LiteralPath $targetFolder).ProviderPath.TrimEnd("\")

[System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
[System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null

Write-Host "変換対象フォルダ: ${targetFolder}"
Write-Host "出力先フォルダ　: ${indexDir}"
Write-Host ""

$targetFiles = @(readListFile $listFile)
if ($targetFiles.Count -gt 0) {
    Write-Host "前回中断した変換を再開します。（残り $($targetFiles.Count) 件）"
} else {
    $retryFailed = $false
    $failedFiles = @(readListFile $errorFile)
    if ($failedFiles.Count -gt 0) {
        Write-Host "前回変換に失敗したファイルが $($failedFiles.Count) 件あります。"
        $answer = Read-Host "失敗したファイルのみ再変換しますか？（y: 失敗分のみ / N: 全ファイルを対象に変換）"
        $retryFailed = ($answer -eq "y")
    }
    if (Test-Path -LiteralPath $errorFile) {
        Remove-Item -LiteralPath $errorFile -Force
    }

    if ($retryFailed) {
        $targetFiles = $failedFiles
    } else {
        Write-Host "変換対象のファイルを検索しています..."
        $targetFiles = @(createTargetList $targetFolder)
    }
    writeListFile $listFile $targetFiles
}

if ($targetFiles.Count -eq 0) {
    Remove-Item -LiteralPath $listFile -Force
    Write-Host ""
    Write-Host "変換が必要なファイルはありません。" -ForegroundColor Green
    pause
    exit
}

Write-Host ""
Write-Host "＜＜注意＞＞" -ForegroundColor Yellow
Write-Host "・変換を中断する場合は Ctrl+C を押してください。（次回実行時に続きから再開します）"
Write-Host "・ウィンドウの×ボタンで閉じると、Excelプロセスが残る場合があります。"
Write-Host "  その場合は 9_Excel強制終了.bat を実行してください。"
Write-Host ""
Write-Host "$($targetFiles.Count) 件のファイルを変換します。"
pause

$total = $targetFiles.Count
$successCount = 0
$failedCount = 0

try {
    startExcel

    for ($i = 0; $i -lt $total; $i++) {
        $relPath = $targetFiles[$i]
        Write-Host ("[{0}/{1}] {2}" -f ($i + 1), $total, $relPath)

        try {
            clearTmpDir
            $sheetCount = convertWorkbook (Join-Path $targetFolder $relPath)
            publishTsv ([System.IO.Path]::GetFileName($relPath)) (getOutputDir $relPath)
            Write-Host "    ${sheetCount} シートを変換しました。"
            $successCount++
        } catch {
            Write-Host "    変換に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            [System.IO.File]::AppendAllText($errorFile, "${relPath}`r`n", ${utf8Bom})
            $failedCount++

            # Excelが不安定になっている可能性があるため再起動
            stopExcel
            startExcel
        }

        # 変換対象一覧から処理済みのファイルを削除
        if ($i + 1 -lt $total) {
            writeListFile $listFile ($targetFiles[($i + 1)..($total - 1)])
        } else {
            writeListFile $listFile @()
        }

        if ((($i + 1) % $restartInterval) -eq 0) {
            stopExcel
            startExcel
        }
    }
} finally {
    # Ctrl+C で中断された場合もここは実行される
    stopExcel
    clearTmpDir
}

if (@(readListFile $listFile).Count -eq 0) {
    Remove-Item -LiteralPath $listFile -Force
}

Write-Host ""
Write-Host "ExcelTsv変換が完了しました。（成功: ${successCount} 件 / 失敗: ${failedCount} 件）" -ForegroundColor Green
if (Test-Path -LiteralPath $errorFile) {
    Write-Host "変換に失敗したファイルは $(Split-Path $errorFile -Leaf) に記録しました。" -ForegroundColor Yellow
    Write-Host "1_変換.bat を再度実行すると、失敗したファイルのみを再変換できます。" -ForegroundColor Yellow
}
pause
