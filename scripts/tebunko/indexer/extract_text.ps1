# 1ファイル（テキスト）から行を読み、場所「本文」のTSVに書き出す。

function extractTextFile {
    # テキストファイル 1 つを読み、場所「本文」の TSV 1 つを outDir に書き出す（本文インデックスに 種類=テキスト・部分=本文・対象=本文 で載る）。
    # 出力した TSV の数（0 か 1）を返す。空のファイル・空の行だけのファイルは 0（次のインデックス作成で取り込み直さない）。
    # 大きさの上限・バイナリの判定は readTextFile（shared\core\text_file.ps1）が例外にする。
    # maxBytes はテストで上限を小さく差し替えるための引数（既定は textFileMaxBytes）
    param (
        [string]$sourcePath,
        [string]$outDir,
        [long]$maxBytes = ${textFileMaxBytes}
    )

    $lines = readTextFile $sourcePath $maxBytes
    if ($lines.Count -eq 0) {
        return 0
    }
    $path = Join-Path $outDir (toIndexFileName "本文")
    [System.IO.File]::WriteAllLines((toLongPath $path), $lines, ${utf8Bom})
    return 1
}
