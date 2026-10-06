# 結果の一覧（件数の行・ファイルごとにまとめた表）の判断（判断層）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\search\result_list_view.Tests.ps1）。

function getSearchProgressText {
    # 検索中の要約欄
    param (
        [int]$hits
    )

    return "検索中…　該当 $($hits.ToString('N0')) 件"
}

function getSearchSummaryText {
    # 検索が終わったときの要約欄（ヒットがあるとき）
    param (
        [int]$hits,
        [int]$files,
        [double]$seconds
    )

    return "該当 $($hits.ToString('N0')) 件（$($files.ToString('N0')) ファイル） ・ $($seconds.ToString('0.0')) 秒"
}

function getAppKind {
    # 元のファイル名の拡張子から、アプリの種類（Excel / Word / PowerPoint / テキスト。どれでもなければ空）を返す
    param (
        [string]$book
    )

    $extension = [System.IO.Path]::GetExtension($book).ToLowerInvariant()
    if ($extension -match '^\.xls') {
        return "Excel"
    }
    if ($extension -match '^\.doc') {
        return "Word"
    }
    if ($extension -match '^\.ppt') {
        return "PowerPoint"
    }
    if (testTextExtension $book) {
        return "テキスト"
    }
    return ""
}

function describeFileLocations {
    # ファイルの中でヒットした場所（見つかった順・重複なし）を、見出しの右端に出す文字列にする。
    # 1 か所ならその場所、2 か所以上なら「[シート]4月 ほか 2 か所」（場所の表記は describePlace）
    param (
        [string[]]$labels
    )

    # 検索中にファイル・場所が増えるたびに呼ぶため、パイプライン（Where-Object）を使わない（1 回 1 ms を超えて検索が遅くなる）
    $first = ""
    $count = 0
    foreach ($label in $labels) {
        if ($label) {
            if ($count -eq 0) {
                $first = $label
            }
            $count++
        }
    }
    if ($count -le 1) {
        return $first
    }
    return "$first ほか $($count - 1) か所"
}

function selectShownRows {
    # 行のうち、絞り込み（空ならすべて）に合うものを、元の順のまま返す
    param (
        $rows,
        [string]$filterText
    )

    $shown = New-Object 'System.Collections.Generic.List[object]'
    foreach ($row in $rows) {
        if ($filterText -eq "" -or $row.Contains($filterText)) {
            $shown.Add($row)
        }
    }
    return , $shown
}

function getResultItems {
    # 結果の表に並べる項目。ファイルごとに見出しを 1 つ置き、開いているファイルだけ、その下に行を並べる。
    # 絞り込みで行が 1 つも残らないファイル（ShownCount が 0）は、見出しも出さない。開いているファイルの行は作ってあること
    param (
        $groups
    )

    $items = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $groups) {
        if ($group.ShownCount -eq 0) {
            continue
        }
        $items.Add($group)
        if ($group.IsExpanded) {
            $items.AddRange($group.ShownRows)
        }
    }
    return , $items
}

function getShownHitRows {
    # 絞り込みに合う行を、表の順（閉じているファイルの行も含む）に並べて返す（結果の出力に使う。行はすべて作ってあること）
    param (
        $groups
    )

    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $groups) {
        $rows.AddRange($group.ShownRows)
    }
    return , $rows
}

function prepareHitRow {
    # 画面に出る行の表示用の値（強調セグメント・セル番地・「場所」の列の表記）を作る。作り済みなら何もしない
    param (
        [HitRow]$row
    )

    $row.Prepare()
    if ($row.PlaceDisplay) { return }
    $row.SetPlaceDisplay((describeHitPlace $row.PlaceText $row.IsExcel $row.IsObjectPlace $row.MatchCell $row.MatchCount $row.LineNumber $row.IsText))
}

function sortFileGroups {
    # 列見出しのクリックでの並べ替え。各ファイルの中の行を property の順に並べ替え、
    # ファイルの順は、並べ替えた後の先頭の行の順にする。同じ値のときは見つかった順（Order）
    param (
        $groups,
        [string]$property,
        [bool]$descending
    )

    # 鍵は項目の名前で渡す（スクリプトブロックより速い）。「場所」（Location）は同じ場所の中を行番号の順にする
    # （セル番地の文字の順だと A10 が A9 の前に来るため、番地ではなく行番号で並べる）
    $byValue = @{ Expression = $property; Descending = $descending }
    $byOrder = @{ Expression = "Order"; Descending = $false }
    $rowKeys = if ($property -eq "Location") { @($byValue, @{ Expression = "LineNumber"; Descending = $descending }, $byOrder) } else { @($byValue, $byOrder) }
    foreach ($group in $groups) {
        $sorted = @($group.Rows | Sort-Object $rowKeys)
        $group.Rows.Clear()
        $group.Rows.AddRange([object[]]$sorted)
    }
    $byFirst = @{ Expression = { if ($_.Rows.Count -gt 0) { $_.Rows[0].$property } }; Descending = $descending }
    $sortedGroups = New-Object 'System.Collections.Generic.List[object]'
    $sortedGroups.AddRange([object[]]@($groups | Sort-Object $byFirst, $byOrder))
    return , $sortedGroups
}
