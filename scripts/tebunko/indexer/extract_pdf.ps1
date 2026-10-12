# 1ファイル（PDF）からページごとの文字を読み、場所「ページN」のTSVに書き出す。

function extractPdfFile {
    # PDF ファイル 1 つを読み、ページごとの TSV（"ページ001" など）を outDir に書き出す（本文インデックスに 種類=PDF で載る）。
    # 出力した TSV の数を返す。文字の無いページは出さず、文字の無い PDF（画像だけ）は 0（次のインデックス作成で取り込み直さない）。
    # 元のファイルには影響を出さないため、extractTextFile と同じく、読み取りだけ・共有モードで outDir に source.copy へコピーし（copyFileShared）、
    # コピーを読んで、TSV を書く前に消す（outDir の中身は本文インデックスに回るので、コピーを残さない）。
    # 暗号化・壊れている・上限を超える・文字の対応表が無いなど読めない PDF は、readPdfFile が例外にする。
    # 大きさの上限は、コピーの前に元の大きさ（メタデータ。元は開かない）で確かめる。
    # maxBytes・limits はテストで上限を小さく差し替えるための引数（既定は pdfFileMaxBytes・newPdfLimits）
    param (
        [string]$sourcePath,
        [string]$outDir,
        [long]$maxBytes = ${pdfFileMaxBytes},
        [hashtable]$limits = $null
    )

    if ([System.IO.FileInfo]::new((toLongPath $sourcePath)).Length -gt $maxBytes) {
        throw "ファイルサイズが大きすぎるため更新できません。"
    }
    $copyPath = Join-Path $outDir "source.copy"
    $removeError = $null
    try {
        copyFileShared $sourcePath $copyPath
        $pages = readPdfFile $copyPath $maxBytes $limits
    } finally {
        try {
            if (Test-Path -LiteralPath (toLongPath $copyPath)) {
                Remove-Item -LiteralPath (toLongPath $copyPath) -Force -ErrorAction Stop
            }
        } catch {
            $removeError = $_.Exception.Message
        }
    }
    if ($removeError) {
        throw "作業領域のコピーを消せませんでした: $removeError"
    }
    $units = [ordered]@{}
    for ($i = 0; $i -lt $pages.Count; $i++) {
        addUnitLines $units ("ページ{0:D3}" -f ($i + 1)) ($pages[$i] -split "`n")
    }
    return (writeUnits $units $outDir)
}