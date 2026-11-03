# 結果の一覧（件数の行・ファイルごとにまとめた表）の判断（判断層）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\search\result_list_view.Tests.ps1）。

# 1 つのファイルの下に、一度に出す行の数（続きを出すたびに、この数ずつ増やす）。1 万行を作るのに 1 秒を超えるので、
# 何万件もヒットするファイルを開いたときに画面が固まらないよう、先頭からこの数までにする
${groupRowLimit} = 1000

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

function getSearchCountingText {
    # 件数の上限を超えたあと、総数を数えている間の要約欄
    param (
        [int]$limit,
        [int]$counted
    )

    return "検索中…　該当 $($limit.ToString('N0')) 件以上（総数を数えています：$($counted.ToString('N0')) 件）"
}

function getSearchFinishSummaryText {
    # 検索が終わったときの要約欄（ヒットがあるとき）。$hits は表に出した件数。
    # 上限を超えた（$truncated）ときは、数え終えていれば総数（$hitTotal）を添え、数えている間に止まった（-1）ときは総数を数えていないと伝える。
    # 秒を渡さなければ時間は付けない
    param (
        [int]$hits,
        [int]$files,
        [double]$seconds = -1,
        [bool]$truncated = $false,
        [int]$hitTotal = -1
    )

    if (!$truncated) {
        return getSearchSummaryText $hits $files $seconds
    }
    if ($hitTotal -ge 0) {
        $text = "全 $($hitTotal.ToString('N0')) 件中 $($hits.ToString('N0')) 件を表示（$($files.ToString('N0')) ファイル）"
    } else {
        $text = "該当 $($hits.ToString('N0')) 件以上（総数は数えていません）（$($files.ToString('N0')) ファイル）"
    }
    if ($seconds -ge 0) {
        $text += "・$($seconds.ToString('0.0')) 秒"
    }
    return $text
}

function getSearchFinishStatusText {
    # 検索が終わった（中止・上限を超えたを含む）ときのステータスバーの文。条件（describeSearchOption の結果。無ければ空）は後ろに付ける
    param (
        [string]$word,
        [int]$hits,
        [string]$optionText = "",
        [bool]$truncated = $false,
        [int]$hitTotal = -1,
        [bool]$countTimeout = $false,
        [bool]$cancelled = $false
    )

    if ($truncated) {
        if ($hitTotal -ge 0) {
            $text = "検索しました：$word（全 $($hitTotal.ToString('N0')) 件のうち、先頭の $($hits.ToString('N0')) 件を表示）。フォルダ・ファイル名の順に先頭から表示しています。ワード・対象ファイル・検索対象で絞り込んでください。"
        } elseif ($countTimeout) {
            $text = "正規表現の照合に時間がかかりすぎるため、総数を数えられませんでした（先頭の $($hits.ToString('N0')) 件を表示）。正規表現を見直してください。"
        } else {
            $text = "中止しました（先頭の $($hits.ToString('N0')) 件を表示。総数は数えていません）"
        }
        if ($optionText -ne "") {
            $text += "　条件：$optionText"
        }
        return $text
    }
    if ($cancelled) {
        return "中止しました（$($hits.ToString('N0')) 件まで表示）"
    }
    return getSearchStatusText $word $hits $optionText
}

function getNoHitText {
    # 検索が 0 件で終わったときの、結果欄の要約と結果の表の空いた所の文（インデックス未作成・種類に合うファイルが無いときは別の文）
    return "一致する結果はありません"
}

function getSearchStatusText {
    # 検索が終わったときのステータスバーの文（0 件もこの形。結果欄の要約には getNoHitText を出す）。
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

function getSkippedContentIndexText {
    # 検索で読めずに飛ばした本文インデックスのファイルの件数の文（ステータスバーの文の後ろに付ける。0 件は空）
    param (
        [int]$count
    )

    if ($count -le 0) {
        return ""
    }
    return "　読めなかったインデックスのファイル：$($count.ToString('N0')) 件"
}

function getFilteredSummaryText {
    # 結果を絞り込んでいるときの要約欄（全部の件数のうち、いくつ見せているか）
    param (
        [int]$hits,
        [int]$shown
    )

    return "一致 $($hits.ToString('N0')) 件中 $($shown.ToString('N0')) 件を表示"
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

function getFileHeaderLocationPrefix {
    # 見出しの右端（「<場所> ・ N 件」）で、件数の前に付ける場所。場所が 1 か所だけなら、行の「場所」列と同じなので付けない（件数だけにする）。
    # 2 か所以上なら describeFileLocations の文字列に「 ・ 」を付ける
    param (
        [string[]]$labels
    )

    $count = 0
    foreach ($label in $labels) {
        if ($label) { $count++ }
    }
    if ($count -le 1) {
        return ""
    }
    return "$(describeFileLocations $labels) ・ "
}

function getPreviewRowHeight {
    # 結果の表とプレビューの境目。結果が少ないとき（表が内容の高さで足りるとき）は、表を内容の高さまで縮めて、プレビューを広げる。
    #   available: 表とプレビューを合わせた高さ
    #   content  : 表（件数・絞り込みなどの行を含む）が内容に要る高さ。多くて測らないときは $null
    # 表が内容より低くなる（スクロールが要る）ときは、今までどおりの高さ（defaultHeight）にする。表は minResultHeight より低くしない。
    # 表に行がまだ無いとき（起動直後・検索前・0 件）は、内容が空なので広げても使い道が無い。defaultHeight にして、結果の側を大きく保つ
    #   itemCount: 表の行の数（既定は -1 = 行の数を問わない）
    param (
        [double]$available,
        $content,
        [double]$defaultHeight = 180,
        [double]$minResultHeight = 60,
        [int]$itemCount = -1
    )

    if ($null -eq $content -or $itemCount -eq 0) {
        return $defaultHeight
    }
    $height = $available - [double]$content
    if ($height -le $defaultHeight) {
        return $defaultHeight
    }
    return [Math]::Min($height, [Math]::Max($available - $minResultHeight, $defaultHeight))
}

function getGroupRowTarget {
    # 開いたファイルの下に出す行の数（行を作る数）。ヒットの数が rowLimit（0 は既定の groupRowLimit）を超えるなら rowLimit
    param (
        [int]$hitCount,
        [int]$rowLimit = 0
    )

    $limit = $(if ($rowLimit -gt 0) { $rowLimit } else { ${groupRowLimit} })
    return [Math]::Min($hitCount, $limit)
}

function getGroupDisplayCount {
    # 開いたファイルの下に表に出す行の数。作ってある行（shownCount。出力・コピーで全件を作ることがある）のうち、
    # 先頭から getGroupRowTarget まで。ヒットを持たない（hitCount が 0）ときは、作ってある行すべて
    param (
        [int]$shownCount,
        [int]$hitCount,
        [int]$rowLimit = 0
    )

    if ($hitCount -le 0) {
        return $shownCount
    }
    return [Math]::Min($shownCount, (getGroupRowTarget $hitCount $rowLimit))
}

function getNextRowLimit {
    # 「続き」を出したあとの、行を作る件数の上限（今の上限 rowLimit に groupRowLimit を足す。0 は既定の groupRowLimit）
    param (
        [int]$rowLimit = 0
    )

    return $(if ($rowLimit -gt 0) { $rowLimit } else { ${groupRowLimit} }) + ${groupRowLimit}
}

function getSortRowTarget {
    # 並べ替えの前に、ファイルごとに作る行の数。ヒットが limit 以下なら全件（行を並べ替えるため）、超えるなら 1 件
    # （生のヒットを先に並べ替えてあるので、ファイルの順を決める先頭の行だけあればよい。getSortedHitOrder・sortFileGroups）
    param (
        [int]$hitCount,
        [int]$limit = ${groupRowLimit}
    )

    return $(if ($hitCount -gt $limit) { 1 } else { $hitCount })
}

function getSortedHitOrder {
    # 生のヒットを property（Location・Kind・LineNumber・Line）の順に並べ替えた、ヒットの番号（見つかった順の 0 始まり）の並びを返す。
    # 同じ値のときは見つかった順。行を作らずに並べ替えるので、ヒットが何万件あっても速い（10 万件で 1 秒ほど）。
    # 文字は Sort-Object と同じ（現在のカルチャーで大文字・小文字を区別しない）。「場所」は同じ場所の中を行番号の順にする。
    # 「種別」（Kind）は場所から決まるので、場所ごとの種別を kindByLocation（場所 → 種別）で渡す。
    # property が上のどれでもなければ $null（並べ替えない）
    param (
        $hits,
        [string]$property,
        [bool]$descending,
        $kindByLocation = @{}
    )

    if ($property -eq "Line") {
        return , (getHitOrderByLine $hits $descending)
    }
    if ($property -eq "Location" -or $property -eq "Kind" -or $property -eq "LineNumber") {
        return , (getHitOrderByNumber $hits $property $descending $kindByLocation)
    }
    return $null
}

function getOrderedHits {
    # ファイルの素のヒットを、表の順に並べて返す（結果の出力に使う。行の部品 HitRow を新しく作らない）。
    #   hitOrder があれば、その順（getSortedHitOrder。ヒットが多いファイル）
    #   無く、行が全部作ってあれば（rows の数 = ヒットの数）、行の Order（見つかった順の番号）の順 = 並べ替えた表の順
    #   それ以外（行が足りない・並べ替えない列）は、見つかった順
    param (
        $hits,
        $hitOrder,
        $rows
    )

    $result = New-Object 'System.Collections.Generic.List[object]' $hits.Count
    if ($null -ne $hitOrder -and $hitOrder.Length -eq $hits.Count) {
        foreach ($index in $hitOrder) {
            $result.Add($hits[$index])
        }
    } elseif ($null -ne $rows -and $rows.Count -eq $hits.Count -and $hits.Count -gt 0) {
        foreach ($row in $rows) {
            $result.Add($hits[$row.Order])
        }
    } else {
        $result.AddRange([object[]]@($hits))
    }
    return , $result
}

function getHitOrderByLine {
    # 該当行（Line）の順（getSortedHitOrder）
    param (
        $hits,
        [bool]$descending
    )

    $n = $hits.Count
    $keys = New-Object 'string[]' $n
    $items = New-Object 'int[]' $n
    $i = 0
    foreach ($hit in $hits) {
        $line = $hit.Line
        $keys[$i] = $(if ($null -eq $line) { "" } else { [string]$line })
        $items[$i] = $i
        $i++
    }
    $comparer = [System.StringComparer]::CurrentCultureIgnoreCase
    [Array]::Sort($keys, $items, $comparer)
    if ($descending) {
        [Array]::Reverse($keys)
        [Array]::Reverse($items)
    }
    # 同じ値の並びは、見つかった順に戻す（[Array]::Sort は同じ値の順を保たない）
    $start = 0
    for ($i = 1; $i -le $n; $i++) {
        if ($i -lt $n -and $comparer.Compare($keys[$i - 1], $keys[$i]) -eq 0) {
            continue
        }
        if ($i - $start -gt 1) {
            [Array]::Sort($items, $start, $i - $start)
        }
        $start = $i
    }
    return $items
}

function getHitOrderByNumber {
    # 場所（Location。同じ場所の中は行番号の順）・種別（Kind。同じ種別の中は見つかった順）・行番号（LineNumber）の順（getSortedHitOrder）。
    # 数だけの鍵にして並べ替える: （文字の順位 × (最大の行番号 + 1) + 行番号）× 件数 + 見つかった順。
    # ヒットの値は 1 回だけ読む（PowerShell のオブジェクトの項目の読み出しが、並べ替えの時間の大半を占めるため）
    param (
        $hits,
        [string]$property,
        [bool]$descending,
        $kindByLocation
    )

    $n = $hits.Count
    $byLocation = ($property -eq "Location")
    $byText = ($property -ne "LineNumber")
    $ranks = @{}
    $locations = New-Object 'System.Collections.Generic.List[string]'
    $locationOf = New-Object 'string[]' $n
    $lineOf = New-Object 'long[]' $n
    [long]$maxLine = 0
    $i = 0
    foreach ($hit in $hits) {
        if ($byLocation -or !$byText) {
            [long]$number = $hit.LineNumber
            $lineOf[$i] = $number
            if ($number -gt $maxLine) { $maxLine = $number }
        }
        if ($byText) {
            $location = [string]$hit.Location
            $text = $(if ($byLocation) { $location } else { [string]$kindByLocation[$location] })
            $locationOf[$i] = $text
            if (!$ranks.ContainsKey($text)) {
                $ranks[$text] = 0
                $locations.Add($text)
            }
        }
        $i++
    }
    $sortedLocations = @($locations | Sort-Object)
    for ($r = 0; $r -lt $sortedLocations.Count; $r++) {
        $ranks[$sortedLocations[$r]] = $r
    }
    [long]$span = $maxLine + 1
    $keys = New-Object 'long[]' $n
    $items = New-Object 'int[]' $n
    for ($i = 0; $i -lt $n; $i++) {
        [long]$rank = 0
        if ($byText) {
            $rank = $ranks[$locationOf[$i]]
            if ($descending) { $rank = $sortedLocations.Count - 1 - $rank }
        }
        [long]$line = $lineOf[$i]
        if ($descending) { $line = $maxLine - $line }
        $keys[$i] = ($rank * $span + $line) * $n + $i
        $items[$i] = $i
    }
    [Array]::Sort($keys, $items)
    return $items
}

function getMoreRowsText {
    # 上限を超えたファイルの末尾に出す「続き」の行の文言。remaining は、まだ行にしていないヒットの数
    param (
        [int]$remaining,
        [int]$step = ${groupRowLimit}
    )

    return "続きの $([Math]::Min($remaining, $step).ToString('N0')) 件を表示（残り $($remaining.ToString('N0')) 件）"
}

function getResultItems {
    # 結果の表に並べる項目。ファイルごとに見出しを 1 つ置き、開いているファイルだけ、その下に行を並べる
    # （まだ行にしていないヒットが残るファイルは、行の末尾に「続き」の行 MoreRow を置く）。
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
            $items.AddRange($group.ShownRows.GetRange(0, (getGroupDisplayCount $group.ShownRows.Count $group.Hits.Count $group.RowLimit)))
            if ($null -ne $group.MoreRow) {
                $items.Add($group.MoreRow)
            }
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
    $row.SetPlaceDisplay((formatDisplayText (describeHitPlace $row.PlaceText $row.IsExcel $row.IsObjectPlace $row.MatchCell $row.MatchCount $row.LineNumber $row.IsText)))
}

function sortFileGroups {
    # 列見出しのクリックでの並べ替え。各ファイルの中の行を property の順に並べ替え、
    # ファイルの順は、並べ替えた後の先頭の行の順にする。同じ値のときは見つかった順（Order）。
    # ヒットが limit を超えるファイルは、全件の行を作らないので、行は並べ替えない（生のヒットを getSortedHitOrder で先に並べ替え、
    # その順に作ってある）。ファイルの順は、どのファイルも並べ替えた後の先頭の行の値で決める
    param (
        $groups,
        [string]$property,
        [bool]$descending,
        [int]$limit = ${groupRowLimit}
    )

    # 鍵は項目の名前で渡す（スクリプトブロックより速い）。「場所」（Location）は同じ場所の中を行番号の順にする
    # （セル番地の文字の順だと A10 が A9 の前に来るため、番地ではなく行番号で並べる）
    $byValue = @{ Expression = $property; Descending = $descending }
    $byOrder = @{ Expression = "Order"; Descending = $false }
    $rowKeys = if ($property -eq "Location") { @($byValue, @{ Expression = "LineNumber"; Descending = $descending }, $byOrder) } else { @($byValue, $byOrder) }
    foreach ($group in $groups) {
        if ($group.Hits.Count -gt $limit) {
            continue
        }
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
    #   readOnlySupported: 読み取り専用で開けるファイルか（testReadOnlySupported。Office の形式だけ。偽なら［読み取り専用で開く］を出さない）
    # 返すもの: @{ Id; Header; Bold } の並び。区切りは Id が "separator"。
    # 画面層は、この結果をそのままメニューに並べる（可否や並びを画面層で決めない）
    param (
        [string]$target,
        [string]$openMode = ${openModeNormal},
        [bool]$expanded = $false,
        [bool]$readOnlySupported = $true
    )

    $labels = [ordered]@{
        ${openModeNormal}   = @{ Id = "openNormal";   Header = "開く" }
        ${openModeNew}      = @{ Id = "openNew";      Header = "新規で開く" }
    }
    if ($readOnlySupported) {
        $labels[${openModeReadOnly}] = @{ Id = "openReadOnly"; Header = "読み取り専用で開く" }
    }
    $separator = @{ Id = "separator"; Header = ""; Bold = $false }

    # 先頭は既定の開き方（見出しは読み取り専用。読み取り専用で開けないファイルは通常）。下には、ほかの開き方を
    # ［開く ▼］と同じ順（開く・新規で開く・読み取り専用で開く）で並べる
    $wanted = if ($target -eq "group") { ${openModeReadOnly} } else { $openMode }
    $default = if ($labels.Contains($wanted)) { $wanted } else { ${openModeNormal} }
    $items = @(@{ Id = $labels[$default].Id; Header = $labels[$default].Header; Bold = $true })
    foreach ($mode in $labels.Keys) {
        if ($mode -ne $default) {
            $items += @{ Id = $labels[$mode].Id; Header = $labels[$mode].Header; Bold = $false }
        }
    }
    $items += @{ Id = "openFolder"; Header = "フォルダを開く"; Bold = $false }
    $items += $separator

    if ($target -eq "group") {
        $items += @{ Id = "copyPath"; Header = "ファイルのパスをコピー"; Bold = $false }
        $items += @{ Id = "toggleGroup"; Header = $(if ($expanded) { "この結果を折りたたむ" } else { "この結果を開く" }); Bold = $false }
        return $items
    }
    $items += @{ Id = "copyRows"; Header = "選んだ行をコピー"; Bold = $false }
    $items += @{ Id = "copyPath"; Header = "ファイルのパスをコピー"; Bold = $false }
    return $items
}
