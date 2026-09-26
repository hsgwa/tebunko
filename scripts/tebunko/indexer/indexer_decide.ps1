# 1 ファイルを取り込み直すかどうかの判断（差分取り込みの要）。
# ファイルにも画面にも触らないため、そのままテストできる（tests\tebunko\indexer\indexer_decide.Tests.ps1）。

# 読み取る内容（抽出版）。読み取る場所を増やしたら、その形式の版を上げる。
# 前の版で取り込んだファイルは、更新が無くても取り込み直す（取り込み一覧の「抽出版」。空は 1）
#   2: Excel（.xlsx / .xlsm）の図形・コメントを読む。
#      Word・PowerPoint のコメント・SmartArt・グラフを読み、Word のテキストボックスを本文から図形に分ける
#      （Word・PowerPoint は旧形式も新形式に変換してから読むため、.doc / .ppt も上げる）
${extractVersions} = @{
    ".xlsx" = 2; ".xlsm" = 2
    ".docx" = 2; ".docm" = 2; ".doc" = 2
    ".pptx" = 2; ".pptm" = 2; ".ppt" = 2
}

function getExtractVersion {
    # ファイルの形式（拡張子）の今の抽出版を返す
    param (
        [string]$path
    )

    $version = ${extractVersions}[[System.IO.Path]::GetExtension($path).ToLowerInvariant()]
    return $(if ($null -eq $version) { 1 } else { $version })
}

function testExtractOutdated {
    # 取り込み一覧の行が、今の抽出版より前の版で取り込んだものか
    param (
        $row
    )

    $version = 1
    if ($row.抽出版) {
        [void][int]::TryParse([string]$row.抽出版, [ref]$version)
    }
    return ($version -lt (getExtractVersion $row.相対パス))
}

function getIngestDecision {
    # 前回の取り込み一覧の行と、いまのファイルの更新日時・サイズから、取り込むかどうかと、その理由を返す。
    #   Ingest: 取り込むか / Reason: done（取り込み済み）・failed（前回失敗。再取り込みするかは呼び出し元が決める）・
    #            new（一覧に無い）・updated（更新された）・pending（前回未完了）・lost（インデックスが無い・壊れている）・
    #            outdated（前の抽出版で取り込んだ）
    param (
        $old,                 # 前回の取り込み一覧の行（無ければ $null）
        [string]$updated,     # いまのファイルの更新日時（formatFileTime）
        [string]$size,        # いまのファイルのサイズ（文字列）
        [bool]$indexComplete  # インデックス（TSV）がそろっているか（状態が「済」のときだけ意味がある）
    )

    # 前回と更新日時・サイズが同じで、前回「未取り込み」で終わっていなければ、同じファイルとみなす
    $sameFile = ($old -and $old.更新日時 -eq $updated -and $old.サイズ -eq $size -and $old.状態 -ne ${stateNew})
    # 取り込み済みでも、インデックス（TSV）が無くなっていれば取り込み直す。
    # 一覧だけを見ると「済」のままになり、検索しても出てこない状態が続くため
    $lostIndex = ($sameFile -and $old.状態 -eq ${stateDone} -and -not $indexComplete)

    if ($sameFile -and -not $lostIndex) {
        if ($old.状態 -eq ${stateFailed}) {
            return @{ Ingest = $false; Reason = "failed" }
        }
        if (testExtractOutdated $old) {
            return @{ Ingest = $true; Reason = "outdated" }
        }
        return @{ Ingest = $false; Reason = "done" }
    }
    if ($lostIndex) {
        return @{ Ingest = $true; Reason = "lost" }
    }
    if ($null -eq $old) {
        return @{ Ingest = $true; Reason = "new" }
    }
    if ($old.状態 -eq ${stateNew}) {
        return @{ Ingest = $true; Reason = "pending" }
    }
    return @{ Ingest = $true; Reason = "updated" }
}

# 取り込みのレーン（docs/00_共通_4_プロセスとスレッド.md 7.4）。Office のレーンはアプリごとに 1 つのスレッドで、そのアプリを 1 つ持つ。
# 読み取りのレーンは Office を使わず、ZIP の中の XML を直接読む（複数のスレッド）
${laneExcel}      = "Excel"
${laneWord}       = "Word"
${lanePowerPoint} = "PowerPoint"
${laneReader}     = "Reader"

function getIngestLane {
    # 取り込むファイルのレーンを拡張子で決める。Excel はセルの表示値を読むため、すべて Excel のレーン。
    # Word・PowerPoint は、旧形式（.doc / .ppt）だけ Office のレーン、新形式は読み取りのレーン
    # （中身が旧形式・パスワード付きと分かったら、読み取りのスレッドが Office のレーンに回し直す。getOfficeLane）
    param (
        [string]$relPath
    )

    switch -Regex ([System.IO.Path]::GetExtension($relPath).ToLowerInvariant()) {
        "^\.xls"   { return ${laneExcel} }
        "^\.doc$"  { return ${laneWord} }
        "^\.ppt$"  { return ${lanePowerPoint} }
    }
    return ${laneReader}
}

function getOfficeLane {
    # 読み取りのスレッドが Office が要ると返したファイルの、回し先の Office のレーン
    param (
        [string]$relPath
    )

    if ([System.IO.Path]::GetExtension($relPath).ToLowerInvariant().StartsWith(".ppt")) {
        return ${lanePowerPoint}
    }
    return ${laneWord}
}