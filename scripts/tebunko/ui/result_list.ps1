# ［検索］タブの結果の表の中身（ファイルごとの見出しと、開いているファイルの行）。
# 検索のヒットは生のまま、元のファイルの見出し（FileGroup）の Hits に持つ。表の行（HitRow）は、開いたとき・絞り込み・
# 並べ替え・出力・コピーのときに、必要なファイルの分だけ作る（ensureRows）。表（ResultGrid）には見出しと、開いているファイルの行だけを入れる。
# 検索した直後はすべて閉じているので、検索中の処理はヒットをファイルごとに分けるだけで済み、ヒットが多くても速い。
# 1 つのファイルの行は、先頭から groupRowLimit 件まで（続きの行で次の分を足す。ヒットが何万件でも固まらない）。
# 並べる項目・絞り込み・並べ替えの判断は search\result_list_view.ps1（getResultItems・getGroupRowTarget・sortFileGroups）。

$script:hitCount = 0             # すべてのヒットの数
$script:fileGroups = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)  # 元のファイルのフルパス → FileGroup
$script:groupList = New-Object 'System.Collections.Generic.List[object]'    # FileGroup（表の順）
$script:dirtyGroups = New-Object 'System.Collections.Generic.HashSet[object]'  # 検索中にヒットが増えた FileGroup（flushResults で表に反映する）
$script:resultItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:traits = @{}             # 行の種類の判定（ファイルの種類・テキストか・セル番地つきの場所か）を、ファイルと場所ごとに覚えておく
$script:places = @{}             # 場所の表示（describePlace）を、種類（Excel か）と場所ごとに覚えておく
$script:rowWord = ""             # 表の行の強調に使う検索ワードと正規表現（startSearch が setRowPattern で決める）
$script:rowPattern = $null
$script:expandNew = $false       # 検索中に新しく見つかったファイルを開いておくか（［すべて開く］を押したら $true）
$ui.ResultGrid.ItemsSource = $script:resultItems

# 見出しの下の行をこの数より多く出し入れするときは、1 件ずつ入れずに表を作り直す（1 件ずつだと画面の更新が重い）
${rebuildThreshold} = 200

function clearResults {
    # 新しく検索する前に、結果を空にする
    param (
        [string]$word,
        [regex]$pattern
    )

    $script:hitCount = 0
    $script:fileGroups.Clear()
    $script:groupList.Clear()
    $script:dirtyGroups.Clear()
    $script:places = @{}
    $script:traits = @{}
    $script:rowWord = $word
    $script:rowPattern = $pattern
    $script:expandNew = $false
    foreach ($column in $ui.ResultGrid.Columns) {
        $column.SortDirection = $null
    }
    setResultItems (New-Object 'System.Collections.Generic.List[object]')
}

function newFileGroup {
    # 元のファイルの見出し（FileGroup）を作って覚える（key は元のファイルのフルパス）
    param (
        [string]$key,
        [string]$relDir,
        [string]$book
    )

    $group = [FileGroup]::new()
    $group.Order = $script:groupList.Count
    $group.Book = $book
    $group.RelDir = $relDir
    $group.FullPath = $key
    # 結果の表に出す値（インデックスから来た文字列の制御文字・向きを変える文字は置き換える。開く・コピーには元の値を使う）
    $group.BookDisplay = formatDisplayText $book
    $group.RelDirDisplay = formatDisplayText $relDir
    $group.FullPathDisplay = formatDisplayText $key
    $group.AppKind = getContentIndexFileKind $book
    $group.IsExpanded = $script:expandNew
    $script:fileGroups[$key] = $group
    $script:groupList.Add($group)
    return $group
}

function getPlace {
    # 場所の表示（describePlace の @{ Place; Kind }）。同じ種類のファイル・場所なら同じなので覚えておく
    param (
        [string]$book,
        [string]$location,
        [string]$fileKind    # getContentIndexFileKind $book（呼ぶ側が 1 回だけ求めた値）
    )

    $key = "$([int]($fileKind -eq "Excel"))|$location"
    $described = $script:places[$key]
    if ($null -eq $described) {
        $described = describePlace $book $location
        $script:places[$key] = $described
    }
    return $described
}

function getRowTraits {
    # 行（HitRow）に入れる判定（ファイルの種類 FileKind・テキストか IsText・セル番地つきの場所か IsCellPrefixed）。
    # ファイルと場所が同じなら同じなので覚えておく（判定は 1 回 0.5 ミリ秒ほどかかり、同じファイルの行ごとに繰り返すと、1,000 行で 1 秒ほど余計にかかる）
    param (
        [string]$book,
        [string]$location
    )

    $key = "$book|$location"
    $traits = $script:traits[$key]
    if ($null -eq $traits) {
        $traits = @{
            FileKind = getContentIndexFileKind $book
            IsText = testTextExtension $book
            IsCellPrefixed = testCellPrefixedPlace $book (splitObjectPlace $location).Kind
        }
        $script:traits[$key] = $traits
    }
    return $traits
}

function addFileGroupLocation {
    # 見出しに、そのファイルで初めてヒットした場所（表の「場所」と同じ表記。describePlace）を足し、右端の表記を作り直す
    param (
        [FileGroup]$group,
        [string]$book,
        [string]$location
    )

    if ($group.AddLabel((getPlace $book $location (getContentIndexFileKind $book)).Place)) {
        # 場所の表記にはシート名などインデックスの文字列が入るので、表に出す値は制御文字・向きを変える文字を置き換える
        $group.SetLocationText((formatDisplayText (describeFileLocations $group.GetLocations())), (formatDisplayText (getFileHeaderLocationPrefix $group.GetLocations())))
    }
}

function testGroupShown {
    # ファイルが絞り込み（ファイル名。空なら全部）に合うか。ファイルごとに決まるので、行を作らなくても分かる
    param (
        [FileGroup]$group
    )

    return ($script:filterText -eq "" -or (testHitFilter $group.Book $script:filterText))
}

function ensureRows {
    # ファイルのヒットのうち、まだ表の行（HitRow）にしていないものを、先頭から target 件目まで行にする（既定は全部。
    # 絞り込みに合うものは ShownRows にも入れる）。行を作るのは重い（1 万行で 1 秒を超える）ので、画面に出す分だけを渡す
    param (
        [FileGroup]$group,
        [int]$target = -1
    )

    $end = $(if ($target -lt 0 -or $target -gt $group.Hits.Count) { $group.Hits.Count } else { $target })
    $shown = testGroupShown $group
    $order = $group.HitOrder
    for ($i = $group.Rows.Count; $i -lt $end; $i++) {
        $hitIndex = $(if ($null -ne $order -and $i -lt $order.Length) { $order[$i] } else { $i })
        $hit = $group.Hits[$hitIndex]
        # インデックスのフォルダ（work\index）からの相対パスの先頭がインデックス名（splitIndexRelPath と同じ）
        $relDir = [string]$hit.RelDir
        $cut = $relDir.IndexOf("\")
        $indexName = if ($cut -lt 0) { $relDir } else { $relDir.Substring(0, $cut) }
        $row = [HitRow]::Create($indexName, $hit.Root, $hit.RelPath, $relDir, $hit.FileName,
                $hit.Book, $hit.Location, [int]$hit.LineNumber, $hit.Line, $script:rowWord, $script:rowPattern)
        $traits = getRowTraits $hit.Book $hit.Location
        $described = getPlace $hit.Book $hit.Location $traits.FileKind
        $row.PlaceText = $described.Place
        $row.Kind = $described.Kind
        $row.IsExcel = ($traits.FileKind -eq "Excel")
        $row.IsText = $traits.IsText
        $row.IsCellPrefixed = $traits.IsCellPrefixed
        $row.Order = $hitIndex
        $row.FileGroup = $group
        $group.Rows.Add($row)
        if ($shown) {
            $group.ShownRows.Add($row)
        }
    }
}

function ensureDisplayRows {
    # ファイルを開いたときに表に出す分（先頭から、そのファイルの行の上限まで）の行を作る
    param (
        [FileGroup]$group
    )

    ensureRows $group (getGroupRowTarget $group.Hits.Count $group.RowLimit)
}

function resetShownRows {
    # 作った行のうち絞り込みに合うもの（ShownRows）を作り直す。絞り込みはファイルごとに決まるので、合うファイルは全部の行、合わなければ 0 件
    param (
        [FileGroup]$group
    )

    $group.ShownRows = $(if (testGroupShown $group) {
        New-Object 'System.Collections.Generic.List[object]' (, $group.Rows)
    } else {
        New-Object 'System.Collections.Generic.List[object]'
    })
}

function ensureAllRows {
    # すべてのファイルのヒットを、全部表の行にする（出力の前。ヒットが多いファイルは重い）
    foreach ($group in $script:groupList) {
        ensureRows $group
    }
}

function updateGroupCount {
    # 見出しの件数。絞り込みに合うファイルはヒットの数（行を作っていなくても分かる）、合わないファイルは 0
    param (
        [FileGroup]$group
    )

    $group.SetCount($(if (testGroupShown $group) { $group.Hits.Count } else { 0 }))
}

function testResultSorted {
    # 結果の表を列見出しで並べ替えているか
    return [bool]($ui.ResultGrid.Columns | Where-Object { $null -ne $_.SortDirection } | Select-Object -First 1)
}

function updateMoreRow {
    # 見出しの下の行の末尾の「続き」の行（MoreRow）を、まだ行にしていないヒットに合わせる。
    # 開いていて、残りがあり、絞り込みに合うときだけ置き、そうでなければ（表に入っていれば出して）無くす
    param (
        [FileGroup]$group
    )

    $remaining = $group.Hits.Count - (getGroupRowTarget $group.Hits.Count $group.RowLimit)
    if ($group.IsExpanded -and $remaining -gt 0 -and $group.ShownCount -gt 0) {
        if ($null -eq $group.MoreRow) {
            $group.MoreRow = [MoreRowsItem]::new()
            $group.MoreRow.Group = $group
        }
        $group.MoreRow.SetText((getMoreRowsText $remaining))
    } elseif ($null -ne $group.MoreRow) {
        if ($group.MoreInView) {
            [void]$script:resultItems.Remove($group.MoreRow)
            $group.MoreInView = $false
        }
        $group.MoreRow = $null
    }
}

function flushResults {
    param (
        [switch]$SkipPaneHeight
    )

    # 検索中に増えたヒットを表に反映する（pumpSearch が 1 回ごとに呼ぶ）。増えたヒットが無い回は何もしない。
    #   SkipPaneHeight: 表とプレビューの境目を決め直さない（呼んだ側が、あとで 1 回決め直すとき）
    # 新しいファイルは見出しを末尾に足し、開いているファイルは行を作って見出しの下に足す
    if ($script:dirtyGroups.Count -eq 0) {
        return
    }
    foreach ($group in $script:dirtyGroups) {
        if ($group.IsExpanded) {
            ensureDisplayRows $group
        }
        updateGroupCount $group
        if ($group.ShownCount -eq 0) {
            continue
        }
        if (!$group.InView) {
            $script:resultItems.Add($group)
            $group.InView = $true
            $group.DisplayedCount = 0
        }
        if ($group.IsExpanded -and ($group.DisplayedCount -lt (getGroupDisplayCount $group.ShownRows.Count $group.Hits.Count $group.RowLimit) -or $group.MoreInView -or $group.Hits.Count -gt (getGroupRowTarget $group.Hits.Count $group.RowLimit))) {
            insertGroupRows $group
        }
    }
    $script:dirtyGroups.Clear()
    if (!$SkipPaneHeight) {
        updateResultPaneHeight $script:resultItems.Count
    }
}

function insertGroupRows {
    # 開いているファイルの、まだ表に入れていない行を、見出しの下（入れ済みの行の後ろ）に入れる
    param (
        [FileGroup]$group
    )

    $index = $script:resultItems.IndexOf($group) + 1 + $group.DisplayedCount
    $end = getGroupDisplayCount $group.ShownRows.Count $group.Hits.Count $group.RowLimit
    for ($i = $group.DisplayedCount; $i -lt $end; $i++) {
        $script:resultItems.Insert($index, $group.ShownRows[$i])
        $index++
    }
    $group.DisplayedCount = $end
    updateMoreRow $group
    if ($null -ne $group.MoreRow -and !$group.MoreInView) {
        $script:resultItems.Insert($index, $group.MoreRow)
        $group.MoreInView = $true
    }
}

function setResultItems {
    # 表の中身を items で置き換える（1 件ずつ入れずに、まとめて 1 回で入れ替える）。選んでいた項目が残っていれば選び直す
    param (
        $items
    )

    $selected = $ui.ResultGrid.SelectedItem
    foreach ($group in $script:groupList) {
        $group.InView = $false
        $group.DisplayedCount = 0
        $group.MoreInView = $false
    }
    foreach ($item in $items) {
        if ($item -is [FileGroup]) {
            $item.InView = $true
            if ($item.IsExpanded) {
                $item.DisplayedCount = getGroupDisplayCount $item.ShownRows.Count $item.Hits.Count $item.RowLimit
            }
        } elseif ($item -is [MoreRowsItem]) {
            $item.Group.MoreInView = $true
        }
    }
    $script:resultItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]' (, [System.Collections.Generic.List[object]]$items)
    $ui.ResultGrid.ItemsSource = $script:resultItems
    if ($null -ne $selected -and $script:resultItems.Contains($selected)) {
        $ui.ResultGrid.SelectedItem = $selected
        $ui.ResultGrid.ScrollIntoView($selected)
    }
}

function rebuildResults {
    # 見出しの開閉・絞り込み・並べ替えのあとに、表の中身を作り直す（開いているファイルの行は、ここで作る。上限を超えた分は「続き」の行にする）
    foreach ($group in $script:groupList) {
        if ($group.IsExpanded) {
            ensureDisplayRows $group
        }
        updateMoreRow $group
    }
    setResultItems (getResultItems $script:groupList)
    updateResultPaneHeight $script:resultItems.Count
}

function toggleFileGroup {
    # 見出しの下の行を閉じる・開く（開くときに、そのファイルの行を作る）
    param (
        [FileGroup]$group
    )

    if ($group.IsExpanded) {
        $group.SetExpanded($false)
        updateMoreRow $group
        if ($group.DisplayedCount -gt ${rebuildThreshold}) {
            rebuildResults
            return
        }
        $index = $script:resultItems.IndexOf($group) + 1
        for ($i = 0; $i -lt $group.DisplayedCount; $i++) {
            $script:resultItems.RemoveAt($index)
        }
        $group.DisplayedCount = 0
        updateResultPaneHeight $script:resultItems.Count
    } else {
        $group.SetExpanded($true)
        ensureDisplayRows $group
        if ((getGroupDisplayCount $group.ShownRows.Count $group.Hits.Count $group.RowLimit) -gt ${rebuildThreshold}) {
            rebuildResults
            return
        }
        insertGroupRows $group
        updateResultPaneHeight $script:resultItems.Count
    }
}

function showMoreGroupRows {
    # 「続き」の行を選んだとき、そのファイルの次の分（groupRowLimit 件）の行を作って、見出しの下に足す
    param (
        [FileGroup]$group
    )

    $group.RowLimit = getNextRowLimit $group.RowLimit
    $first = $group.DisplayedCount
    ensureDisplayRows $group
    if (!$group.InView -or !$group.IsExpanded) {
        return
    }
    $end = getGroupDisplayCount $group.ShownRows.Count $group.Hits.Count $group.RowLimit
    if ($end - $first -gt ${rebuildThreshold}) {
        rebuildResults
    } else {
        insertGroupRows $group
        updateResultPaneHeight $script:resultItems.Count
    }
    if ($first -lt $end) {
        $ui.ResultGrid.ScrollIntoView($group.ShownRows[$first])
    }
}

function activateResultItem {
    # Enter・クリックで、見出しなら開閉、「続き」の行なら次の分を出す。どちらでもなければ $false（行は元のファイルを開く）
    param (
        $item
    )

    if ($item -is [FileGroup]) {
        toggleFileGroup $item
        return $true
    }
    if ($item -is [MoreRowsItem]) {
        showMoreGroupRows $item.Group
        return $true
    }
    return $false
}

function setAllFileGroupsExpanded {
    # ［すべて開く］［すべて折りたたむ］。検索中なら、このあと見つかるファイルも同じにする
    param (
        [bool]$expanded
    )

    $script:expandNew = $expanded
    foreach ($group in $script:groupList) {
        $group.SetExpanded($expanded)
    }
    rebuildResults
}

function applyResultFilter {
    # 絞り込み（空なら解除）を当てて表を作り直す。絞り込みはファイル名に当てるのでファイルごとに決まり、行は作らない
    param (
        [string]$filterText
    )

    foreach ($group in $script:groupList) {
        resetShownRows $group
        updateGroupCount $group
    }
    rebuildResults
}

function sortResults {
    # 列見出しのクリックでの並べ替え（もう一度で逆順）。ファイルの中の行を並べ替え、ファイルの順もそれに合わせる
    param (
        [System.Windows.Controls.DataGridColumn]$column
    )

    $property = $column.SortMemberPath
    if (!$property) {
        return
    }
    $descending = $column.SortDirection -eq [System.ComponentModel.ListSortDirection]::Ascending
    foreach ($other in $ui.ResultGrid.Columns) {
        $other.SortDirection = $null
    }
    $column.SortDirection = if ($descending) { [System.ComponentModel.ListSortDirection]::Descending } else { [System.ComponentModel.ListSortDirection]::Ascending }
    applySort $property $descending
}

function getKindByLocation {
    # ファイルのヒットの場所ごとの「種別」（場所 → セル・図形・本文など。行の Kind と同じ値）。種別で並べ替えるときに使う
    param (
        [FileGroup]$group
    )

    $kinds = @{}
    foreach ($hit in $group.Hits) {
        $location = [string]$hit.Location
        if (!$kinds.ContainsKey($location)) {
            $kinds[$location] = (getPlace $hit.Book $location (getRowTraits $hit.Book $location).FileKind).Kind
        }
    }
    return $kinds
}

function applySort {
    # ファイルの中の行とファイルの順を並べ替え、絞り込みを当て直して表を作り直す。
    # 並べ替えには全件の行を作る。ただしヒットが多いファイル（groupRowLimit 件を超える）は、行を作らずに生のヒットを並べ替えて
    # （getSortedHitOrder）、その順に先頭の 1 件だけ作る（続きの分は、開くときと続きを出すときに同じ順で作る）
    param (
        [string]$property,
        [bool]$descending
    )

    foreach ($group in $script:groupList) {
        if ($group.Hits.Count -gt ${groupRowLimit}) {
            $group.HitOrder = getSortedHitOrder $group.Hits $property $descending $(if ($property -eq "Kind") { getKindByLocation $group } else { @{} })
            $group.Rows.Clear()
            $group.RowLimit = 0
        } else {
            $group.HitOrder = $null
        }
        ensureRows $group (getSortRowTarget $group.Hits.Count)
    }
    $sorted = sortFileGroups $script:groupList $property $descending
    $script:groupList.Clear()
    $script:groupList.AddRange($sorted)
    foreach ($group in $script:groupList) {
        resetShownRows $group
    }
    rebuildResults
}

function finishResults {
    # 検索の終わりに、表の順を整える。検索中に並べ替えていたら、あとから来た行も含めて並べ直す
    flushResults -SkipPaneHeight
    $column = $ui.ResultGrid.Columns | Where-Object { $null -ne $_.SortDirection } | Select-Object -First 1
    if ($column) {
        applySort $column.SortMemberPath ($column.SortDirection -eq [System.ComponentModel.ListSortDirection]::Descending)
    } else {
        updateResultPaneHeight $script:resultItems.Count
    }
}

function getCurrentHitRow {
    # 選んでいる行。見出しを選んでいるときは、そのファイルの先頭の行（プレビュー・元のファイルを開く・パスのコピーに使う）。
    # 「続き」の行を選んでいるときは、行が無い
    $item = $ui.ResultGrid.SelectedItem
    if ($item -is [MoreRowsItem]) {
        return $null
    }
    if ($item -is [FileGroup]) {
        ensureRows $item 1
        if ($item.ShownRows.Count -gt 0) {
            return $item.ShownRows[0]
        }
        return $null
    }
    return $item
}

function getViewRows {
    # 表示中（絞り込み・並べ替え後）の行。閉じているファイルの行も含む（結果をファイルに出力するときに使う）
    ensureAllRows
    return , (getShownHitRows $script:groupList).ToArray()
}

function getExportHits {
    # 結果をファイルに出力するときの、表に出ている分（RelDir・Book・Location・LineNumber・Line を持つ並び）。
    #   並べ替えていて上限を超えていない: 表の順（getViewRows。行の部品 HitRow を作る）
    #   それ以外: 検索の順（ファイルの見つかった順 → ファイルの中の見つかった順）。行の部品（HitRow）は作らず、ファイルの素の値を並べる。
    #            上限を超えているときは溢れの続きとして書くので並べ替えは当てない。絞り込みはファイル名で、ファイルごとに決まる
    param (
        [bool]$truncated,
        [string]$filterText,
        [bool]$sorted = $false
    )

    if ($sorted -and !$truncated) {
        return , (getViewRows)
    }
    $hits = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in ($script:groupList | Sort-Object Order)) {
        if (testHitFilter $group.Book $filterText) {
            $hits.AddRange($group.Hits)
        }
    }
    return , $hits.ToArray()
}

function getSelectedRows {
    # 選んでいる行を表の順に返す。見出しを選んでいるときは、そのファイルの行（絞り込みに合うもの）をすべて含める
    $selected = New-Object 'System.Collections.Generic.HashSet[object]'
    foreach ($item in $ui.ResultGrid.SelectedItems) {
        [void]$selected.Add($item)
    }
    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $script:groupList) {
        $whole = $selected.Contains($group)
        if ($whole) {
            ensureRows $group
        }
        foreach ($row in $group.ShownRows) {
            if ($whole -or $selected.Contains($row)) {
                $rows.Add($row)
            }
        }
    }
    return , $rows.ToArray()
}

function getShownHitCount {
    # 絞り込みに合う行の数
    $count = 0
    foreach ($group in $script:groupList) {
        $count += $group.ShownCount
    }
    return $count
}
# ---- イベント ----

# 表示用（強調セグメント・DisplayLine・セル番地・「場所」の列の表記）は、行が画面に出るときだけ作る（件数が多くても軽い）。
# HitRow.Prepare は1回だけ実行し、作った値は PropertyChanged で反映する
$ui.ResultGrid.Add_LoadingRow({
    param ($s, $e)
    if ($e.Row.Item -is [HitRow]) { prepareHitRow $e.Row.Item }
})

# 並べ替えは表（DataGrid）に任せず、ファイルごとに行う（任せると見出しと行が混ざって並ぶ）
$ui.ResultGrid.Add_Sorting({
    param ($sender, $e)
    $e.Handled = $true
    safe { sortResults $e.Column }
})

# 見出しの行をクリックすると閉じる・開く（Shift・Ctrl を押しながらのクリックは、複数を選ぶための操作なので切り替えない）。
# MouseLeftButtonUp は押された要素にしか届かない（外側の DataGrid では受け取れない）ため、外側から順に届く
# PreviewMouseUp で受ける。行を選ぶ処理が済んでから表の中身を入れ替えるよう、開閉は画面の処理のあとに行う
$ui.ResultGrid.AddHandler([System.Windows.UIElement]::PreviewMouseUpEvent, [System.Windows.Input.MouseButtonEventHandler] {
    param ($sender, $e)
    # 修飾キーは列挙型の値で比べる（文字列 "None" と比べると、押していなくても真になる）
    if ($e.ChangedButton -ne [System.Windows.Input.MouseButton]::Left -or
            [System.Windows.Input.Keyboard]::Modifiers -ne [System.Windows.Input.ModifierKeys]::None) {
        return
    }
    $element = $e.OriginalSource
    while ($element -and !($element -is [System.Windows.Controls.DataGridRow])) {
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    if ($element -and ($element.Item -is [FileGroup] -or $element.Item -is [MoreRowsItem])) {
        $script:clickedItem = $element.Item
        [void]$ui.ResultGrid.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Input, [Action]{
            safe { [void](activateResultItem $script:clickedItem) }
        })
    }
}, $true)

$ui.ExpandAllButton.Add_Click({ safe { setAllFileGroupsExpanded $true } })
$ui.CollapseAllButton.Add_Click({ safe { setAllFileGroupsExpanded $false } })
