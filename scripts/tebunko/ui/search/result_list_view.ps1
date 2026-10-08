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
    # 検索が終わったときの要約欄（ヒットがあるとき。秒を渡さなければ時間は付けない）
    param (
        [int]$hits,
        [int]$files,
        [double]$seconds = -1
    )

    $text = "一致 $($hits.ToString('N0')) 件（$($files.ToString('N0')) ファイル）"
    if ($seconds -lt 0) {
        return $text
    }
    return "$text・$($seconds.ToString('0.0')) 秒"
}

function getSearchStatusText {
    # 検索が終わったときのステータスバーの文（0 件もこの形。結果欄には 0 件の文を出さない）。
    # 条件（describeSearchOption の結果。無ければ空）があれば、後ろに付ける
    param (
        [string]$word,
        [int]$hits,
        [string]$optionText = ""
    )

    $text = "検索しました：$word $($hits.ToString('N0')) 件"
    if ($optionText -ne "") {
        $text += "　条件：$optionText"
    }
    return $text
}

function getFilteredSummaryText {
    # 結果を絞り込んでいるときの要約欄（全部の件数のうち、いくつ見せているか）
    param (
        [int]$hits,
        [int]$shown
    )

    return "$($hits.ToString('N0')) 件中 $($shown.ToString('N0')) 件を表示"
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

function testResultMenuKeepSelection {
    # 結果の行を右クリックしたとき、今の選びを変えずにメニューを出すか。
    # ヒットした行を複数選んでいて、その中の行を右クリックしたときだけ $true（［選んだ行をコピー］のため）。
    # 右クリックした行が選ばれていないとき・見出しのとき・選びに見出しを含むときは $false（右クリックした行だけを選ぶ）。
    #   rowSelected: 右クリックした行が今選ばれているか ／ item: その行の項目 ／ selectedItems: 今の選び
    param (
        [bool]$rowSelected,
        $item,
        $selectedItems
    )

    if (!$rowSelected -or $item -is [FileGroup]) {
        return $false
    }
    return @($selectedItems | Where-Object { $_ -is [FileGroup] }).Count -eq 0
}

function getResultMenuContext {
    # メニューを出す項目（選んでいる項目）から、メニューの種類（"row"・"group"）と、見出しを開いているかを決める
    param (
        $item
    )

    if ($item -is [FileGroup]) {
        return @{ Target = "group"; Expanded = [bool]$item.IsExpanded }
    }
    return @{ Target = "row"; Expanded = $false }
}

function getResultMenuItems {
    # 検索結果の右クリックメニューの並び。
    #   target: "row"（ヒットした行）・"group"（ファイルの見出し）
    #   openMode: 設定の既定の開き方（行のメニューの先頭に太字で出す）
    #   expanded: 見出しの下の行を開いているか（見出しのメニューの開閉の項目の文言を決める）
    # 返すもの: @{ Id; Header; Bold } の並び。区切りは Id が "separator"。
    # 画面層は、この結果をそのままメニューに並べる（可否や並びを画面層で決めない）
    param (
        [string]$target,
        [string]$openMode = ${openModeNormal},
        [bool]$expanded = $false
    )

    $labels = [ordered]@{
        ${openModeNormal}   = @{ Id = "openNormal";   Header = "開く" }
        ${openModeNew}      = @{ Id = "openNew";      Header = "新規で開く" }
        ${openModeReadOnly} = @{ Id = "openReadOnly"; Header = "読み取り専用で開く" }
    }
    $separator = @{ Id = "separator"; Header = ""; Bold = $false }

    if ($target -eq "group") {
        return @(
            @{ Id = "openReadOnly"; Header = $labels[${openModeReadOnly}].Header; Bold = $true }
            @{ Id = "openFolder"; Header = "フォルダを開く"; Bold = $false }
            $separator
            @{ Id = "copyPath"; Header = "ファイルのパスをコピー"; Bold = $false }
            @{ Id = "toggleGroup"; Header = $(if ($expanded) { "この結果を折りたたむ" } else { "この結果を開く" }); Bold = $false }
        )
    }

    # 先頭は既定の開き方。下には、ほかの開き方を［開く ▾］と同じ順（開く・新規で開く・読み取り専用で開く）で並べる
    $default = if ($labels.Contains($openMode)) { $openMode } else { ${openModeNormal} }
    $items = @(@{ Id = $labels[$default].Id; Header = $labels[$default].Header; Bold = $true })
    foreach ($mode in $labels.Keys) {
        if ($mode -ne $default) {
            $items += @{ Id = $labels[$mode].Id; Header = $labels[$mode].Header; Bold = $false }
        }
    }
    $items += @{ Id = "openFolder"; Header = "フォルダを開く"; Bold = $false }
    $items += $separator
    $items += @{ Id = "copyRows"; Header = "選んだ行をコピー"; Bold = $false }
    $items += @{ Id = "copyPath"; Header = "ファイルのパスをコピー"; Bold = $false }
    return $items
}
