# ［検索］タブのプレビューの判断（高さから読む行数を決める・短い文字列にする）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\preview_view.Tests.ps1）。

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

function getPreviewFillWidth {
    # Excel 以外（Word・PowerPoint・テキスト）のプレビューは列が 1 つ。その幅を、枠の幅いっぱいにする
    # （行番号の列を除いた残り。中身から決めた初めの幅より狭くはしない。枠が狭まれば初めの幅まで戻る）。
    # 枠の幅がまだ分からないとき・列が収まる最小幅に満たないときは初めの幅。
    # 列見出しのドラッグで手で変えた幅（今の幅が、前に合わせた幅と違う）は変えない
    param (
        [double]$viewportWidth,  # プレビューのスクロールの見える幅
        [double]$numberWidth,    # 行番号の列の幅
        [double]$baseWidth,      # 中身から決めた初めの列の幅
        [double]$currentWidth,   # 今の列の幅
        [double]$lastFitWidth,   # 前に合わせた列の幅（初めは初めの幅）
        [double]$minWidth = 24   # これより狭くしない
    )

    if ($currentWidth -ne $lastFitWidth) {
        return $currentWidth
    }
    $fill = $viewportWidth - $numberWidth
    if ($viewportWidth -le 0 -or $fill -lt $minWidth) {
        return $baseWidth
    }
    return [math]::Max($baseWidth, $fill)
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

function getPreviewMenuItems {
    # プレビューの右クリックメニューの並び。@{ Id; Header; Bold } の並び（区切りは Id が "separator"）。
    # 先頭は、選んでいる結果の行の場所で元のファイルを開く（ダブルクリックと同じ処理）
    return @(
        @{ Id = "openHere"; Header = "元のファイルのこの場所を開く"; Bold = $true }
        @{ Id = "separator"; Header = ""; Bold = $false }
        @{ Id = "copyCell"; Header = "選んだセルをコピー"; Bold = $false }
        @{ Id = "copyRow"; Header = "この行をコピー"; Bold = $false }
    )
}
