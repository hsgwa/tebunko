# 実行パス取得
${dir} = Split-Path $MyInvocation.MyCommand.Path -Parent

${execDir} = Split-Path $MyInvocation.MyCommand.Path -Parent

# 検索結果格納ファイル生成
${fullPath} = "${dir}\03_検索結果.txt"
${file} = New-Object System.IO.StreamWriter($fullPath, $false)

$indexDir = Get-Content "${dir}\02__検索対象インデックスパス.txt" -Raw -Encoding UTF8

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

# 読み込んだファイルをstring型の配列として順次処理する
(Get-Content .\02__検索ワード.txt -Encoding UTF8) -as [string[]] | % {
    # 検索対象文字列
    ${word} = $_
    echo "検索文字列： $word"
    # ファイル書き込み
    ${file}.writeline("【検索文字列　${word}】")
    ${file}.writeline("ファイル名	シート名")

    # 検索対象フォルダのファイルに対して順次処理を行う
    # "Select-String"を使用してファイルの中身を検索する

    # TODO: dirの中身が無い時にメッセージを出す。
    $indexDir -as [string[]] | % {
        $dir = $_
        echo "検索対象フォルダ：${dir}"
        # Select-String ${word}
        # Select-String
        Get-ChildItem "${dir}\" -Include "*.tsv" -Recurse | Select-String -Pattern ${word} | % {
            # 検索結果は":"区切りの文字列のため、必要情報を分割して取得する
            # ファイル名はパス付きで1つ目の配列に格納されている
            ${resFile} = $($_ -split ":")[0..1] -join ":"
            ${resFile} = ${resFile}.ToString() -replace " ", ""

            ${offsetDir} = $resFile.Replace($indexDir, "")

            # "\"でパスを配列化し、配列の末尾を取得(\は特殊文字のためエスケープ)
            ${resFileName} = (${resFile}.ToString() -split "\\")
            ${resFileName} = (${resFileName} -split "\\")[$(${resFileName}.Length) - 1]

            $ext = $resFileName -replace ".tsv", ""
            $ext = "." + ($ext -split "\.")[-1]

            $sheetName = $resFileName.Replace(".tsv", "")
            $sheetName = $sheetName -replace "^.*\.xls[a-z]?_", ""
            ${resFileName} = $resFileName.Replace("_" + $sheetName + ".tsv", "")

            # 検索結果は配列の3つ目に格納されている
            ${resStr} = $($_ -split ":")[3]
            if ($resStr -eq $null) {
                $resStr = ""
            }
            # Excelのセルの数だけタブが設定されているため、取り除く
            ${resStr} = ${resStr}.ToString() -replace "'t", ""

            # 結果の書き込み
            ${file}.writeline("${resFileName}	${sheetName}	${resStr}")
        }
    }
}

# ファイルをクローズし、メモリの解放を行う
${file}.Close()
${file} = $null
[System.GC]::Collect([System.GC]::MaxGeneration)
echo "${fullPath} に結果を出力しました。"
echo ""
