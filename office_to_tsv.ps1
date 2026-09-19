pause

# 実行パス取得
${execDir} = Split-Path $MyInvocation.MyCommand.Path -Parent
${indexDir} = "$execDir\index"
${workDir} = "$execDir\tmp"

[System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
[System.IO.Directory]::CreateDirectory($workDir) | Out-Null

function getTargetFolder {
    $targetFolder = Get-Content "${execDir}\01__変換対象フォルダパス.txt" -Raw -Encoding UTF8

    if (($targetFolder | Measure-Object -Line).Lines -ne 1) {
        Write-Host "＜エラー＞"
        Write-Host $targetFolder.Split("\n").Count
        Write-Host "01__変換対象フォルダパス.txt は１行だけ記載してください。" -ForegroundColor Red
        pause
        exit
    }

    return $targetFolder
}

$targetFolder = getTargetFolder

$listFile = "${execDir}\01__変換対象一覧（一時ファイル）.txt"
$errorFile = "${execDir}\01__変換失敗一覧（一時ファイル）.txt"

echo "＜＜注意＞＞"
echo "本プログラムを停止する際はCtrl+Cで停止させてください。"
echo "ウィンドウの閉じるボタンを押した場合、Excelプロセスがゾンビ化します。"
echo "バックグラウンドプロセスを停止させる際は、killExcellProcess.ps1を実行してください。"
echo ""
echo ""

if (!(Test-Path $listFile) -or ((Get-Content $listFile).Count -eq 0)) {
    if (!(Test-Path $errorFile)) {
        echo "変換対象一覧が見つかりません。変換対象の一覧を作成します。"
        pause
        Get-ChildItem "${targetFolder}" -Include "*.xlsx", "*.xls", "*.xlsm" -Recurse -Name | Set-Content -Path $listFile
        # Get-ChildItem "${targetFolder}" -Include "*.doc" -Recurse -Name | Set-Content -Path $listFile
        Remove-Item "$errorFile" -Force 2>${NULL}
    } else {
        echo "変換失敗一覧（一時ファイル）が見つかりました。再度取込を行います。"
        pause
        Move-Item $errorFile $listFile
    }
} else {
    echo "変換対象一覧が見つかりました。01__変換対象一覧（一時ファイル）.txtを読み込みます。"
    pause
}

function replaceNewLineToSpace {
    # 1セル内の改行をスペースに変換
    param (
        [string]$inputString
    )

    $output = New-Object -TypeName System.Text.StringBuilder
    $insideQuotes = $false
    $chars = $content.ToCharArray()

    foreach ($char in $chars) {
        if ($char -eq '"') {
            $insideQuotes = -not $insideQuotes
        }
        if (-not ($insideQuotes -and ($char -eq "`n" -or $char -eq "`r"))) {
            [void]$output.Append($char)
        }
    }

    return $output.ToString()
}

function prettyTsv {
    param (
        [string]$inputFilePath,
        [string]$outputFilePath
    )

    $content = Get-Content $inputFilePath -Raw

    $content = replaceNewLineToSpace $content
    $content = $content -replace "(?m)^[\s\t]+\r\n", ""  # 空行を削除
    $content = $content -replace "(?m)[\s\t]+$", ""      # 末尾のタブスペースを削除

    $content | Out-File -FilePath $outputFilePath -Encoding utf8 -Force
}

function cleanupScript ([ref]$excelInstance, [ref]$wbInstance) {
    if (${wbInstance}.Value -ne $null) {
        ${wbInstance}.Value.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbInstance.Value) | Out-Null
        $wbInstance.Value = $null
    }

    if (${excelInstance}.Value -ne $null) {
        ${excelInstance}.Value.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelInstance.Value) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        $excelInstance.Value = $null
    }
}

${excel} = $null
${wb} = $null

# SIGINT(Ctrl+C) のキャッチ
$event = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    cleanupScript ([ref]$excel) ([ref]$wb)
}

try {
    # 検索起点フォルダ配下のEXCELファイルを順次処理する
    $targetFiles = Get-Content $listFile

    foreach ($excelFile in $targetFiles) {
        try {
            ${excel} = New-Object -ComObject Excel.Application
            ${excel}.DisplayAlerts = $false
            ${excel}.EnableEvents = $false
            ${excel}.ScreenUpdating = $false

            # 検索起点フォルダからの相対パスを取得
            echo "Opening.. ${excelFile}"

            $excelFileDir = [System.IO.Path]::GetDirectoryName($excelFile)
            if ([string]::IsNullOrEmpty($excelFileDir)) {
                $excelFileDir = "."
            }

            $ext = "." + ($excelFile -split "\.")[-1]

            # EXCELファイルを開く
            ${wb} = ${excel}.Workbooks.Open("${targetFolder}\${excelFile}", $null, $true)

            # シートに対して順次検索処理を行う。非表示シートは除外。
            ${wb}.Worksheets | Where-Object { $_.Visible -ne 0 } | % {
                ${ws} = $_

                # 対象のシートのみのExcelブックを新規作成。
                ${book} = ${excel}.Workbooks.add()
                ${ws}.copy(${book}.worksheets.item(1))
                ${book}.Worksheets.item(2).delete()  # 新規で作成したExcelファイルには"sheet1"が含まれるため削除

                # Unicode(第２引数の"42"で指定)ファイルとして別名保存
                ${excelFileName} = [System.IO.Path]::GetFileNameWithoutExtension(${excelFile}.ToString())

                $file_name = ${ws}.Name
                # 保存字のNG文字を全角に変換
                $file_name = $file_name.Replace(">", "＞")
                $file_name = $file_name.Replace("<", "＜")
                $file_name = $file_name.Replace("\", "￥")
                $file_name = $file_name.Replace("*", "＊")
                $file_name = $file_name.Replace("""", "””")
                $file_name = $file_name.Replace(":", "：")
                $file_name = $file_name.Replace("?", "？")
                $file_name = $file_name.Replace("|", "｜")
                $file_path = "${workDir}\${excelFileName}${ext}_${file_name}.tsv"
                $file_path_tmp = "${file_path}.tmp"

                ${book}.saveAs($file_path_tmp, 42)

                # tsvファイルを整形
                prettyTsv $file_path_tmp $file_path

                echo "Indexed... ${file_path}"
            }

            cleanupScript ([ref]$excel) ([ref]$wb)

            # フォルダ作成。
            $indexFileDir = "${indexDir}\${excelFileDir}"
            [System.IO.Directory]::CreateDirectory($indexFileDir) | Out-Null

            Move-Item -Path "${workDir}\*.tsv" -Destination "${indexFileDir}" -Force
            Remove-Item "${workDir}\*.tmp" -Force

            # 変換対象一覧（一時ファイル）から対象ファイルの行削除
            $targetFiles = $targetFiles | Where-Object { $_ -ne $excelFile } | Select-Object -Unique
            $targetFiles | Set-Content -Path $listFile
        } catch {
            Write-Error "$_.Exceltion.Message}"
            # 変換失敗一覧（一時ファイル）に対象ファイルの行追加
            Add-Content -Path $errorFile -Value $excelFile

            # 変換対象一覧（一時ファイル）から対象ファイルの行削除
            $targetFiles = $targetFiles | Where-Object { $_ -ne $excelFile } | Select-Object -Unique
            $targetFiles | Set-Content -Path $listFile

            Write-Host "${excelFile}を${errorFile}に追記しました。"
        } finally {
        }
    }
} finally {
    Unregister-Event -SourceIdentifier PowerShell.Exiting
}

if ((Get-Content $listFile -Raw | Measure-Object -Line).Lines -le 1) {
    Remove-Item $listFile
    Write-Host "変換対象一覧（一時ファイル）を削除しました。"
    Write-Host "変換失敗一覧（一時ファイル）を処理する場合は、ExcelTsv変換を再度実行してください。" -ForegroundColor Yellow
}

echo "ExcelTsv変換が完了しました。"
echo ""
pause
