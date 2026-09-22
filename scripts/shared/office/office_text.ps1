# Office ファイル（Office Open XML）を読むときの、文字・パスの小さな部品。
# 制限言語モード（ConstrainedLanguage）でも動く書き方だけで書く（いつものインデクサの office_reader.ps1 と、
# 制限言語モード用の office_reader_clm.ps1 の両方から読み込む）。

function resolveZipPath {
    # リレーションシップの Target（相対パス）を、ZIP内のパスに変換する
    param (
        [string]$baseDir,   # 例: "ppt/slides"
        [string]$target     # 例: "../notesSlides/notesSlide1.xml"
    )

    if ($target.StartsWith("/")) {
        return $target.TrimStart("/")
    }

    # 制限言語モードでも動くよう、List ではなく配列で持つ
    $parts = @()
    foreach ($part in (("$baseDir/$target") -split "/")) {
        if ($part -eq "..") {
            if ($parts.Count -gt 0) { $parts = @($parts | Select-Object -First ($parts.Count - 1)) }
        } elseif ($part -ne "" -and $part -ne ".") {
            $parts += $part
        }
    }
    return ($parts -join "/")
}

function toObjectCellText {
    # 図形・コメントの文字を、Excel のテキスト保存と同じ形の 1 セルにする。
    # 改行はセル内改行（$cellNewLine）にし、改行・" ・タブを含むときは " で囲む（中の " は "" にする）
    param (
        [string[]]$lines
    )

    $text = ($lines -join "`n").Trim()
    if ($text.IndexOfAny([char[]]@('"', "`t", "`r", "`n")) -lt 0) {
        return $text
    }
    $text = ($text -replace "\r\n|\r|\n", ${cellNewLine}).Replace('"', '""')
    return "`"$text`""
}

function getCellPosition {
    # セル番地（例: "AB12"）を @(行, 列) にする。読めなければ @(0, 0)
    param (
        [string]$ref
    )

    $m = [regex]::Match($ref, '^\$?([A-Za-z]+)\$?(\d+)$')
    if (-not $m.Success) {
        return @(0, 0)
    }
    $column = 0
    foreach ($ch in $m.Groups[1].Value.ToUpperInvariant().ToCharArray()) {
        $column = $column * 26 + ([int]$ch - [int][char]"A" + 1)
    }
    return @([int]$m.Groups[2].Value, $column)
}
