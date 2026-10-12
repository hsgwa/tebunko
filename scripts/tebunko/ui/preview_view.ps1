# ［検索］タブのプレビューの判断（選択行を出す位置・列の幅・短い文字列にする）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\preview_view.Tests.ps1）。

function getPreviewTopRow {
    # 選択行をプレビューの上から 3 行目に出すために、一番上に出す行（表の行の番号。0 から）を決める。
    # 選択行の前が 2 行に満たなければ先頭から。選択行が表に無い（-1）ときも先頭から
    param (
        [int]$hitIndex  # 表の中の選択行の番号（無ければ -1）
    )

    return [math]::Max(0, $hitIndex - 2)
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
