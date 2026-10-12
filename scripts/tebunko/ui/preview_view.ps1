# ［検索］タブのプレビューの判断（選択行を出す位置・読み足し・持つ行の上限・列の幅・短い文字列にする）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\preview_view.Tests.ps1）。

function getPreviewTopRow {
    # 選択行をプレビューの上から 3 行目に出すために、一番上に出す行（表の行の番号。0 から）を決める。
    # 選択行の前が 2 行に満たなければ先頭から。選択行が表に無い（-1）ときも先頭から
    param (
        [int]$hitIndex  # 表の中の選択行の番号（無ければ -1）
    )

    return [math]::Max(0, $hitIndex - 2)
}

function getPreviewReadMore {
    # スクロールの位置から、どの向きにどの行を読み足すかを決める。読み足す向きごとに @{ Side; FirstLine; Count } を返す（要らなければ何も返さない。
    # Side は "up"（上。FirstLine から Count 行で、今ある最初の行の直前まで）・"down"（下。今ある最後の行の次から Count 行））。
    # 見えている行が、読んである範囲の端から ahead 行以内に近づいたら、その向きに chunk 行を読み足す。
    # その向きにもう行が無い（done）・その向きを読んでいる途中（loading）ときは頼まない（同じ向きは 1 つずつ）。
    # 上は、1 行目に届くところまでで止める（chunk より少なくなる）
    param (
        [int]$firstVisible,     # 見えている先頭の行の番号（読んである行の中の位置。0 から）
        [int]$lastVisible,      # 見えている最後の行の位置
        [int]$rowCount,         # 読んである行の数
        [int]$firstLine,        # 読んである最初の行の、元のファイルの場所の中の番号（1 から）
        [int]$lastLine,         # 読んである最後の行の番号
        [bool]$topDone,         # 上にもう行が無い（または読めなかった）
        [bool]$bottomDone,
        [bool]$upLoading,       # 上を読んでいる途中
        [bool]$downLoading,
        [int]$ahead = 20,
        [int]$chunk = 200
    )

    if ($rowCount -le 0) { return }
    if (!$topDone -and !$upLoading -and $firstLine -gt 1 -and $firstVisible -le $ahead) {
        $count = [Math]::Min($chunk, $firstLine - 1)
        @{ Side = "up"; FirstLine = $firstLine - $count; Count = $count }
    }
    if (!$bottomDone -and !$downLoading -and ($rowCount - 1 - $lastVisible) -le $ahead) {
        @{ Side = "down"; FirstLine = $lastLine + 1; Count = $chunk }
    }
}

function getPreviewTrim {
    # 持つ行が上限（maxRows）を超えたとき、見ている所から遠い側の端から、どれだけ捨てるかを決める。@{ Side; Count } を返す。
    # Side は "top"（上から捨てる）・"bottom"（下から捨てる）・"none"（捨てない）。1 回に chunk 行ずつ、見えている行は捨てない
    param (
        [int]$firstVisible,     # 見えている先頭の行の位置（読んである行の中。0 から）
        [int]$lastVisible,
        [int]$rowCount,
        [int]$maxRows = 3000,
        [int]$chunk = 200
    )

    if ($rowCount -le $maxRows) { return @{ Side = "none"; Count = 0 } }
    $above = [Math]::Max(0, $firstVisible)
    $below = [Math]::Max(0, $rowCount - 1 - $lastVisible)
    if ($above -gt $below) {
        return @{ Side = "top"; Count = [Math]::Min($chunk, $above) }
    }
    $count = [Math]::Min($chunk, $below)
    if ($count -le 0) { return @{ Side = "none"; Count = 0 } }
    return @{ Side = "bottom"; Count = $count }
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
