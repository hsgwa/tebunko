# ［検索］タブの結果の表の中身（ファイルごとの見出しと、開いているファイルの行）。
# 検索のヒットは生のまま、元のファイルの見出し（FileGroup）の Hits に持つ。表の行（HitRow）は、開いたとき・絞り込み・
# 並べ替え・出力・コピーのときに、必要なファイルの分だけ作る（ensureRows）。表（ResultGrid）には見出しと、開いているファイルの行だけを入れる。
# 検索した直後はすべて閉じているので、検索中の処理はヒットをファイルごとに分けるだけで済み、ヒットが多くても速い。
# 並べる項目・絞り込み・並べ替えの判断は search_view.ps1（getResultItems・selectShownRows・sortFileGroups）。

$script:hitCount = 0             # すべてのヒットの数
$script:fileGroups = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)  # 元のファイルのフルパス → FileGroup
$script:groupList = New-Object 'System.Collections.Generic.List[object]'    # FileGroup（表の順）
$script:dirtyGroups = New-Object 'System.Collections.Generic.HashSet[object]'  # 検索中にヒットが増えた FileGroup（flushResults で表に反映する）
$script:resultItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:places = @{}             # 場所の表示（describePlace）を、種類（Excel か）と場所ごとに覚えておく
$script:rowWord = ""             # 表の行の強調に使う検索ワードと正規表現（startSearch が setRowPattern で決める）
$script:rowPattern = $null
$script:expandNew = $false       # 検索中に新しく見つかったファイルを開いておくか（［すべて展開］を押したら $true）
$script:needsRebuild = $false    # 検索中に、表の途中に見出しを差し込む必要ができた（検索の終わりに作り直す）
$script:lastInViewOrder = -1     # 表に出ている見出しのうち、いちばん後に見つかったものの Order（新しい見出しを末尾に足せるかの判定）
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
    $script:rowWord = $word
    $script:rowPattern = $pattern
    $script:expandNew = $false
    $script:needsRebuild = $false
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
    $group.AppKind = getAppKind $book
    $group.IsExpanded = $script:expandNew
    $script:fileGroups[$key] = $group
    $script:groupList.Add($group)
    return $group
}

function getPlace {
    # 場所の表示（describePlace の @{ Place; Kind }）。同じ種類のファイル・場所なら同じなので覚えておく
    param (
        [string]$book,
        [string]$location
    )

    $key = "$([int]($book -match '\.xls[a-z]?$'))|$location"
    $described = $script:places[$key]
    if ($null -eq $described) {
        $described = describePlace $book $location
        $script:places[$key] = $described
    }
    return $described
}

function addFileGroupLocation {
    # 見出しに、そのファイルで初めてヒットした場所（表の「場所」と同じ表記。describePlace）を足し、右端の表記を作り直す
    param (
        [FileGroup]$group,
        [string]$book,
        [string]$location
    )

    if ($group.AddLabel((getPlace $book $location).Place)) {
        $group.SetLocationText((describeFileLocations $group.GetLocations()))
    }
}

function ensureRows {
    # ファイルのヒットのうち、まだ表の行（HitRow）にしていないものを行にする（絞り込みに合うものは ShownRows にも入れる）
    param (
        [FileGroup]$group
    )

    for ($i = $group.Rows.Count; $i -lt $group.Hits.Count; $i++) {
        $hit = $group.Hits[$i]
        # インデックスのフォルダ（work\index）からの相対パスの先頭がインデックス名（splitIndexRelPath と同じ）
        $relDir = [string]$hit.RelDir
        $cut = $relDir.IndexOf("\")
        $indexName = if ($cut -lt 0) { $relDir } else { $relDir.Substring(0, $cut) }
        $row = [HitRow]::Create($indexName, $hit.Root, $hit.RelPath, $relDir, $hit.FileName,
                $hit.Book, $hit.Location, [int]$hit.LineNumber, $hit.Line, $script:rowWord, $script:rowPattern)
        $described = getPlace $hit.Book $hit.Location
        $row.PlaceText = $described.Place
        $row.Kind = $described.Kind
        $row.Order = $i
        $row.FileGroup = $group
        $group.Rows.Add($row)
        if ($script:filterText -eq "" -or $row.Contains($script:filterText)) {
            $group.ShownRows.Add($row)
        }
    }
}

function ensureAllRows {
    # すべてのファイルのヒットを表の行にする（絞り込み・並べ替え・出力の前）
    foreach ($group in $script:groupList) {
        ensureRows $group
    }
}

function updateGroupCount {
    # 見出しの件数。絞り込んでいなければヒットの数（行を作っていなくても分かる）、絞り込み中は合う行の数
    param (
        [FileGroup]$group
    )

    if ($script:filterText -eq "") {
        $group.SetCount($group.Hits.Count)
    } else {
        $group.SetCount($group.ShownRows.Count)
    }
}

function flushResults {
    # 検索中に増えたヒットを表に反映する（pumpSearch が 1 回ごとに呼ぶ）。
    # 新しいファイルは見出しを末尾に足し、開いているファイルは行を作って見出しの下に足す
    foreach ($group in $script:dirtyGroups) {
        if ($script:filterText -ne "" -or $group.IsExpanded) {
            ensureRows $group
        }
        updateGroupCount $group
        if ($group.ShownCount -eq 0) {
            continue
        }
        if (!$group.InView) {
            if ($group.Order -lt $script:lastInViewOrder) {
                # 絞り込みで隠れていたファイルに、あとから合う行が来た。末尾に足すと順が崩れるため、検索の終わりに作り直す
                $script:needsRebuild = $true
                continue
            }
            $script:resultItems.Add($group)
            $group.InView = $true
            $group.DisplayedCount = 0
            $script:lastInViewOrder = $group.Order
        }
        if ($group.IsExpanded -and $group.DisplayedCount -lt $group.ShownRows.Count) {
            insertGroupRows $group
        }
    }
    $script:dirtyGroups.Clear()
}

function insertGroupRows {
    # 開いているファイルの、まだ表に入れていない行を、見出しの下（入れ済みの行の後ろ）に入れる
    param (
        [FileGroup]$group
    )

    $index = $script:resultItems.IndexOf($group) + 1 + $group.DisplayedCount
    for ($i = $group.DisplayedCount; $i -lt $group.ShownRows.Count; $i++) {
        $script:resultItems.Insert($index, $group.ShownRows[$i])
        $index++
    }
    $group.DisplayedCount = $group.ShownRows.Count
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
    }
    $script:lastInViewOrder = -1
    foreach ($item in $items) {
        if ($item -is [FileGroup]) {
            $item.InView = $true
            if ($item.IsExpanded) {
                $item.DisplayedCount = $item.ShownRows.Count
            }
            $script:lastInViewOrder = [Math]::Max($script:lastInViewOrder, $item.Order)
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
    # 見出しの開閉・絞り込み・並べ替えのあとに、表の中身を作り直す（開いているファイルの行は、ここで作る）
    $script:needsRebuild = $false
    foreach ($group in $script:groupList) {
        if ($group.IsExpanded) {
            ensureRows $group
        }
    }
    setResultItems (getResultItems $script:groupList)
}

function toggleFileGroup {
    # 見出しの下の行を閉じる・開く（開くときに、そのファイルの行を作る）
    param (
        [FileGroup]$group
    )

    if ($group.IsExpanded) {
        $group.SetExpanded($false)
        if ($group.DisplayedCount -gt ${rebuildThreshold}) {
            rebuildResults
            return
        }
        $index = $script:resultItems.IndexOf($group) + 1
        for ($i = 0; $i -lt $group.DisplayedCount; $i++) {
            $script:resultItems.RemoveAt($index)
        }
        $group.DisplayedCount = 0
    } else {
        $group.SetExpanded($true)
        ensureRows $group
        if ($group.ShownRows.Count -gt ${rebuildThreshold}) {
            rebuildResults
            return
        }
        insertGroupRows $group
    }
}

function setAllFileGroupsExpanded {
    # ［すべて展開］［すべて折りたたむ］。検索中なら、このあと見つかるファイルも同じにする
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
    # 絞り込み（空なら解除）を当てて表を作り直す。絞り込むときは、合う行を数えるためにすべての行を作る
    param (
        [string]$filterText
    )

    if ($filterText -ne "") {
        ensureAllRows
    }
    foreach ($group in $script:groupList) {
        $group.ShownRows = selectShownRows $group.Rows $filterText
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

function applySort {
    # ファイルの中の行とファイルの順を並べ替え、絞り込みを当て直して表を作り直す（並べ替えには、すべての行を作る）
    param (
        [string]$property,
        [bool]$descending
    )

    ensureAllRows
    $sorted = sortFileGroups $script:groupList $property $descending
    $script:groupList.Clear()
    $script:groupList.AddRange($sorted)
    foreach ($group in $script:groupList) {
        $group.ShownRows = selectShownRows $group.Rows $script:filterText
    }
    rebuildResults
}

function finishResults {
    # 検索の終わりに、表の順を整える。検索中に並べ替えていたら、あとから来た行も含めて並べ直す。
    # 絞り込みで隠れていたファイルに、あとから合う行が来ていたら作り直す
    flushResults
    $column = $ui.ResultGrid.Columns | Where-Object { $null -ne $_.SortDirection } | Select-Object -First 1
    if ($column) {
        applySort $column.SortMemberPath ($column.SortDirection -eq [System.ComponentModel.ListSortDirection]::Descending)
    } elseif ($script:needsRebuild) {
        rebuildResults
    }
}

function getCurrentHitRow {
    # 選んでいる行。見出しを選んでいるときは、そのファイルの先頭の行（プレビュー・元のファイルを開く・パスのコピーに使う）
    $item = $ui.ResultGrid.SelectedItem
    if ($item -is [FileGroup]) {
        ensureRows $item
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

# 表示用（強調セグメント・DisplayLine・セル列）は、行が画面に出るときだけ作る（件数が多くても軽い）。
# HitRow.Prepare は1回だけ実行し、作った値は PropertyChanged で反映する
$ui.ResultGrid.Add_LoadingRow({
    param ($s, $e)
    if ($e.Row.Item -is [HitRow]) { $e.Row.Item.Prepare() }
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
    if ($element -and $element.Item -is [FileGroup]) {
        $script:clickedGroup = $element.Item
        [void]$ui.ResultGrid.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Input, [Action]{
            safe { toggleFileGroup $script:clickedGroup }
        })
    }
}, $true)

$ui.ExpandAllButton.Add_Click({ safe { setAllFileGroupsExpanded $true } })
$ui.CollapseAllButton.Add_Click({ safe { setAllFileGroupsExpanded $false } })
