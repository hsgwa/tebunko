# ステータスバーの更新中の 1 行の文言（判断層）。画面に触らないため、そのままテストできる
# （tests/tebunko/ui/shell/status_bar_view.Tests.ps1）。

function getIndexingStatusLine {
    # 更新中にステータスバーの左に出す 1 行。件数・失敗は更新全体のもの（インデックスごとではない）
    #   例: 更新中: 営業部2025　205 / 455 件（失敗 1 件）・残り約 3 分
    param (
        [string]$name,      # 更新しているインデックスの名前（取れないときは空）
        [int]$processed,
        [int]$total,
        [int]$failed,
        [string]$etaText    # 残り時間の文言（「残り約 3 分」。まだ分からなければ空）
    )

    $head = if ($name) { "更新中: $name" } else { "更新中" }
    if ($total -le 0 -or $processed -le 0) {
        # 1 件も終わっていない間は、件数を出さない
        return "$head　準備しています"
    }
    $text = "$head　$($processed.ToString('N0')) / $($total.ToString('N0')) 件"
    if ($failed -gt 0) {
        $text += "（失敗 $failed 件）"
    }
    if ($etaText) {
        $text += "・$etaText"
    }
    return $text
}
