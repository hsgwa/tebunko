# 1ファイル（テキスト）から行を読み、場所「本文」のTSVに書き出す。

function extractTextFile {
    # テキストファイル 1 つを読み、場所「本文」の TSV 1 つを outDir に書き出す（本文インデックスに 種類=テキスト・部分=本文・対象=本文 で載る）。
    # 出力した TSV の数（0 か 1）を返す。空のファイル・空の行だけのファイルは 0（次のインデックス作成で取り込み直さない）。
    # 元のファイルには影響を出さないため、読み取りだけ・共有モードで outDir に source<元の拡張子> へコピーし（copyFileShared）、コピーを読んで、
    # TSV を書く前に消す（outDir の中身は本文インデックスに回るので、コピーを残さない）。強制終了で残ったコピーは作業フォルダごと片付けられる。
    # 大きさの上限は、コピーの前に元の大きさ（メタデータ。元は開かない）で確かめる。コピーを読むときの確かめ（readTextFile）は、コピー中に大きくなった場合の保険。
    # バイナリの判定は readTextFile（shared\core	ext_file.ps1）が例外にする。
    # maxBytes はテストで上限を小さく差し替えるための引数（既定は textFileMaxBytes）
    param (
        [string]$sourcePath,
        [string]$outDir,
        [long]$maxBytes = ${textFileMaxBytes}
    )

    if ([System.IO.FileInfo]::new((toLongPath $sourcePath)).Length -gt $maxBytes) {
        throw "ファイルサイズが大きすぎるため更新できません。"
    }
    $copyPath = Join-Path $outDir ("source" + [System.IO.Path]::GetExtension($sourcePath))
    try {
        copyFileShared $sourcePath $copyPath
        $lines = readTextFile $copyPath $maxBytes
    } finally {
        if (Test-Path -LiteralPath (toLongPath $copyPath)) {
            Remove-Item -LiteralPath (toLongPath $copyPath) -Force -ErrorAction SilentlyContinue
        }
    }
    if ($lines.Count -eq 0) {
        return 0
    }
    $path = Join-Path $outDir (toIndexFileName "本文")
    [System.IO.File]::WriteAllLines((toLongPath $path), $lines, ${utf8Bom})
    return 1
}
