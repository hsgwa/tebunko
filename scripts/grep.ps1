# TSV Grep
#
# config\検索ワード.txt に記載したワード（1行に1つ。正規表現可）で
# work\index（または config\検索対象インデックスパス.txt のフォルダ）配下のTSVを検索し、
# 結果を output\検索結果.txt に出力する。

. "$PSScriptRoot\common.ps1"

$ErrorActionPreference = "Stop"

trap {
    Write-Host "＜エラー＞"
    Write-Host $_.Exception.Message -ForegroundColor Red
    pause
    exit 1
}

$words = @(readConfigLines $searchWordFile)
if ($words.Count -eq 0) {
    throw "$(Split-Path $searchWordFile -Leaf) に検索ワードを1行に1つ記載してください。"
}

# 検索対象のTSVを列挙する（TSVのフルパス → インデックスフォルダからの相対パス）
$tsvFiles = @{}
foreach ($dir in @(getIndexFolders)) {
    if (!(Test-Path -LiteralPath $dir -PathType Container)) {
        Write-Host "${dir} が見つかりません。先に 1_変換.bat を実行してください。" -ForegroundColor Yellow
        continue
    }

    $root = (Resolve-Path -LiteralPath $dir).ProviderPath.TrimEnd("\")
    $files = @(Get-ChildItem -LiteralPath $root -Filter "*.tsv" -File -Recurse)
    Write-Host "検索対象フォルダ：${root}（TSV $($files.Count) 件）"

    foreach ($file in $files) {
        $tsvFiles[$file.FullName] = $file.FullName.Substring($root.Length).TrimStart("\")
    }
}

if ($tsvFiles.Count -eq 0) {
    throw "検索対象のTSVがありません。先に 1_変換.bat を実行してください。"
}

$tsvPaths = [string[]]@($tsvFiles.Keys | Sort-Object)

# 検索結果格納ファイル生成
[System.IO.Directory]::CreateDirectory($outputDir) | Out-Null
$writer = New-Object System.IO.StreamWriter($resultFile, $false, ${utf8Bom})

Write-Host ""
try {
    foreach ($word in $words) {
        $params = @{ LiteralPath = $tsvPaths; Pattern = $word; Encoding = "UTF8" }

        # 正規表現として不正なワードは、文字列そのままで検索する
        try {
            [void][regex]::new($word)
        } catch {
            $params.SimpleMatch = $true
        }

        $hits = @(Select-String @params)

        $writer.WriteLine("【検索文字列　${word}】 $($hits.Count) 件")
        $writer.WriteLine("ファイル名`tシート名`t該当行")
        foreach ($hit in $hits) {
            # ブック名の前に、変換対象フォルダからの相対フォルダを付ける
            $relPath = $tsvFiles[$hit.Path]
            $relDir = [System.IO.Path]::GetDirectoryName($relPath)
            $line = toResultLine ([System.IO.Path]::GetFileName($relPath)) $hit.Line
            if ($relDir) {
                $line = "${relDir}\${line}"
            }
            $writer.WriteLine($line)
        }
        $writer.WriteLine("")

        Write-Host ("検索文字列： {0} → {1} 件" -f $word, $hits.Count)
    }
} finally {
    $writer.Close()
}

Write-Host ""
Write-Host "${resultFile} に結果を出力しました。" -ForegroundColor Green
Invoke-Item -LiteralPath $resultFile
pause
