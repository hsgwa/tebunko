# 1 ファイルを変換し直すかどうかの判断（差分変換の要）。
# ファイルにも画面にも触らないため、そのままテストできる（tests\tebunko_grep\convert\convert_decide.Tests.ps1）。

# 読み取る内容（抽出版）。読み取る場所を増やしたら、その形式の版を上げる。
# 前の版で変換したファイルは、更新が無くても変換し直す（変換一覧の「抽出版」。空は 1）
#   2: Excel（.xlsx / .xlsm）の図形・コメントを読む
${extractVersions} = @{ ".xlsx" = 2; ".xlsm" = 2 }

function getExtractVersion {
    # ファイルの形式（拡張子）の今の抽出版を返す
    param (
        [string]$path
    )

    $version = ${extractVersions}[[System.IO.Path]::GetExtension($path).ToLowerInvariant()]
    return $(if ($null -eq $version) { 1 } else { $version })
}

function testExtractOutdated {
    # 変換一覧の行が、今の抽出版より前の版で変換したものか
    param (
        $row
    )

    $version = 1
    if ($row.抽出版) {
        [void][int]::TryParse([string]$row.抽出版, [ref]$version)
    }
    return ($version -lt (getExtractVersion $row.相対パス))
}

function getConvertDecision {
    # 前回の変換一覧の行と、いまのファイルの更新日時・サイズから、変換するかどうかと、その理由を返す。
    #   Convert: 変換するか / Reason: done（変換済み）・failed（前回失敗。再変換するかは呼び出し元が決める）・
    #            new（一覧に無い）・updated（更新された）・pending（前回未完了）・lost（変換結果が無い・壊れている）・
    #            outdated（前の抽出版で変換した）
    param (
        $old,                 # 前回の変換一覧の行（無ければ $null）
        [string]$updated,     # いまのファイルの更新日時（formatFileTime）
        [string]$size,        # いまのファイルのサイズ（文字列）
        [bool]$indexComplete  # 変換結果（TSV）がそろっているか（状態が「済」のときだけ意味がある）
    )

    # 前回と更新日時・サイズが同じで、前回「未変換」で終わっていなければ、同じファイルとみなす
    $sameFile = ($old -and $old.更新日時 -eq $updated -and $old.サイズ -eq $size -and $old.状態 -ne ${stateNew})
    # 変換済みでも、インデックス（TSV）が無くなっていれば変換し直す。
    # 一覧だけを見ると「済」のままになり、検索しても出てこない状態が続くため
    $lostIndex = ($sameFile -and $old.状態 -eq ${stateDone} -and -not $indexComplete)

    if ($sameFile -and -not $lostIndex) {
        if ($old.状態 -eq ${stateFailed}) {
            return @{ Convert = $false; Reason = "failed" }
        }
        if (testExtractOutdated $old) {
            return @{ Convert = $true; Reason = "outdated" }
        }
        return @{ Convert = $false; Reason = "done" }
    }
    if ($lostIndex) {
        return @{ Convert = $true; Reason = "lost" }
    }
    if ($null -eq $old) {
        return @{ Convert = $true; Reason = "new" }
    }
    if ($old.状態 -eq ${stateNew}) {
        return @{ Convert = $true; Reason = "pending" }
    }
    return @{ Convert = $true; Reason = "updated" }
}