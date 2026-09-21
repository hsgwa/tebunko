# ［2 検索］タブのプレビューの判断（高さから読む行数を決める・短い文字列にする）。
# 画面に触らないため、そのままテストできる（tests\windox_grep\ui\preview_view.Tests.ps1）。

function getPreviewRowCounts {
    # プレビューの高さに収まる行数から、選択行の前後に読む行数を決める（前後同数。余りの 1 行は後ろに付ける）。
    # 低くすれば選択行だけ、高くすればその分だけ前後の行が見える
    param (
        [double]$height,    # プレビューに使える高さ
        [double]$rowHeight, # 1 行の高さの目安
        [int]$maxRows       # 出す行数の上限
    )

    $rows = [math]::Floor($height / $rowHeight)
    $rows = [math]::Min([math]::Max($rows, 1), $maxRows)
    $before = [math]::Floor(($rows - 1) / 2)
    return , @([int]$before, [int]($rows - 1 - $before))
}

function toStatusText {
    # ステータスに出す短い文字列（改行・タブはスペースにし、長ければ末尾を省略）
    param (
        [string]$text
    )

    $text = ($text -replace "[`r`n`t]+", " ")
    if ($text.Length -gt 40) {
        return $text.Substring(0, 40) + "…"
    }
    return $text
}