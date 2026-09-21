# 1 ファイルを取り込み直すかどうかの判断（差分取り込みの要）。
# ファイルにも画面にも触らないため、そのままテストできる（tests\tebunko_grep\indexer\indexer_decide.Tests.ps1）。

function getIngestDecision {
    # 前回の取り込み一覧の行と、いまのファイルの更新日時・サイズから、取り込むかどうかと、その理由を返す。
    #   Ingest: 取り込むか / Reason: done（取り込み済み）・failed（前回失敗。再取り込みするかは呼び出し元が決める）・
    #            new（一覧に無い）・updated（更新された）・pending（前回未完了）・lost（インデックスが無い・壊れている）
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