# 各スクリプト共通のパス定義と関数。スクリプト・テストから dot-source して使う。

# フォルダ構成
${rootDir}   = Split-Path $PSScriptRoot -Parent
${configDir} = "${rootDir}\config"
${workDir}   = "${rootDir}\work"
${indexDir}  = "${workDir}\index"
${tmpDir}    = Join-Path ([System.IO.Path]::GetTempPath()) "win_grep"  # Excelは [ ] を含むパスに保存できないため TEMP を使う
${outputDir} = "${rootDir}\output"

${utf8Bom} = New-Object System.Text.UTF8Encoding($true)

# 設定ファイル（ユーザーが編集する）
${targetFolderFile} = "${configDir}\変換対象フォルダパス.txt"
${searchWordFile}   = "${configDir}\検索ワード.txt"
${indexPathFile}    = "${configDir}\検索対象インデックスパス.txt"

# 途中状態・出力（自動生成）
${listFile}   = "${workDir}\変換対象一覧.txt"
${errorFile}  = "${workDir}\変換失敗一覧.txt"
${resultFile} = "${outputDir}\検索結果.txt"

function readConfigLines {
    # 設定ファイルの空行以外の行を配列で返す。ファイルが無ければ空ファイルを作成して例外。
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) {
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
        [System.IO.File]::WriteAllText($path, "", ${utf8Bom})
        throw "$([System.IO.Path]::GetFileName($path)) が無いため作成しました。内容を記入して再実行してください。"
    }

    return @(Get-Content -LiteralPath $path -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
}

function readListFile {
    # 変換対象一覧などの一時ファイルを読み込む（空行を除く）。ファイルが無ければ空配列
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) {
        return @()
    }
    return @(Get-Content -LiteralPath $path -Encoding UTF8 | Where-Object { $_.Trim() -ne "" })
}

function writeListFile {
    param (
        [string]$path,
        [string[]]$lines
    )

    if ($null -eq $lines) {
        $lines = [string[]]@()
    }
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllLines($path, $lines, ${utf8Bom})
}

function getTargetFolder {
    # 変換対象フォルダパスを返す。1行でなければ例外。
    param (
        [string]$path = ${targetFolderFile}
    )

    $lines = @(readConfigLines $path)
    if ($lines.Count -ne 1) {
        throw "$(Split-Path $path -Leaf) は１行だけ記載してください。（現在 $($lines.Count) 行）"
    }

    return $lines[0].Trim('"')
}

function getIndexFolders {
    # 検索対象インデックスのフォルダ一覧を返す。設定が無い・空なら work\index を使う。
    param (
        [string]$path = ${indexPathFile}
    )

    if (Test-Path $path) {
        $lines = @(readConfigLines $path)
        if ($lines.Count -gt 0) {
            return $lines
        }
    }

    return @(${indexDir})
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

function replaceNewLineToSpace {
    # ダブルクォートで囲まれた1セル内の改行を取り除く
    param (
        [string]$inputString
    )

    # "" はクォート内のダブルクォートのエスケープだが、"…" "…" の2つに分けてマッチしても結果は同じ
    return [regex]::Replace($inputString, '"[^"]*"', { param($m) $m.Value -replace "[\r\n]", "" })
}

function formatTsv {
    # セル内改行・空行・行末の空白を取り除く
    param (
        [string]$content
    )

    $content = replaceNewLineToSpace $content

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($content -split "\r?\n")) {
        $line = $line.TrimEnd()
        if ($line -ne "") {
            $lines.Add($line)
        }
    }

    return ($lines -join "`r`n")
}

function prettyTsv {
    # Excelが出力したTSV（UTF-16）を整形してUTF-8で保存する。内容が空なら保存せず $false を返す
    param (
        [string]$inputFilePath,
        [string]$outputFilePath
    )

    $content = formatTsv ([System.IO.File]::ReadAllText($inputFilePath))
    if ($content -eq "") {
        return $false
    }

    [System.IO.File]::WriteAllText($outputFilePath, "${content}`r`n", ${utf8Bom})
    return $true
}

function splitIndexFileName {
    # "ブック名.xlsx_シート名.tsv" をブック名とシート名に分解する
    param (
        [string]$fileName
    )

    if ($fileName -match "^(?<book>.*?\.xls[a-z]?)_(?<sheet>.*)\.tsv$") {
        return @{ book = $Matches.book; sheet = $Matches.sheet }
    }

    return @{ book = $fileName; sheet = "" }
}

function toResultLine {
    # 検索結果1件を "ブック名<TAB>シート名<TAB>該当行" に整形する
    param (
        [string]$fileName,
        [string]$line
    )

    $name = splitIndexFileName $fileName
    return "$($name.book)`t$($name.sheet)`t${line}"
}
