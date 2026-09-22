# ［検索］の結果の表（tebunko_grep\ui\result_list.ps1）のテスト。
# 画面の部品（$ui.ResultGrid など）は偽物にして、見出し・行の出し入れと状態の変化を確かめる。
. "$PSScriptRoot\..\..\helpers\load.ps1"
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
. "${scriptsDir}\shared\ui\types.ps1"
. "${scriptsDir}\tebunko_grep\ui\types_grep.ps1"
. "${scriptsDir}\tebunko_grep\ui\search_view.ps1"

# ---- 画面の偽物 ----
# 読み込み時に登録されるイベントの処理は $handlers に取っておき、テストから呼ぶ
$handlers = @{}

function newFakeButton([string]$name) {
    $button = [pscustomobject]@{ Name = $name }
    $button | Add-Member ScriptMethod Add_Click { param ($block) $handlers["$($this.Name).Click"] = $block }
    return $button
}

$grid = [pscustomobject]@{
    ItemsSource   = $null
    SelectedItem  = $null
    SelectedItems = @()
    Columns       = New-Object 'System.Collections.Generic.List[object]'
    Scrolled      = New-Object 'System.Collections.Generic.List[object]'
}
$grid | Add-Member ScriptMethod Add_LoadingRow { param ($block) $handlers["LoadingRow"] = $block }
$grid | Add-Member ScriptMethod Add_Sorting { param ($block) $handlers["Sorting"] = $block }
$grid | Add-Member ScriptMethod AddHandler { param ($event, $handler, $handledToo) $handlers["PreviewMouseUp"] = $handler }
$grid | Add-Member ScriptMethod ScrollIntoView { param ($item) $this.Scrolled.Add($item) }

$ui = [pscustomobject]@{
    ResultGrid        = $grid
    ExpandAllButton   = newFakeButton "ExpandAll"
    CollapseAllButton = newFakeButton "CollapseAll"
}

# イベント処理が使う画面の共通部品（shared\ui\shell.ps1）の代わり。例外はそのまま出す
function safe {
    param ([scriptblock]$block)
    & $block
}

. "${scriptsDir}\tebunko_grep\ui\result_list.ps1"

# ---- テストの準備 ----

function resetResults {
    $script:filterText = ""
    $ui.ResultGrid.SelectedItem = $null
    $ui.ResultGrid.SelectedItems = @()
    $ui.ResultGrid.Columns.Clear()
    $ui.ResultGrid.Scrolled.Clear()
    clearResults "見積" ([regex]"見積")
}

function addHit {
    # 検索のヒットを 1 件、元のファイルの見出しに足す（検索の処理と同じく、見出しが無ければ作る）
    param (
        [string]$book,
        [string]$location,
        [string]$line,
        [int]$lineNumber = 1,
        [string]$relDir = "営業部\2024"
    )

    $key = "C:\共有\$relDir\$book"
    $group = $null
    if (!$script:fileGroups.TryGetValue($key, [ref]$group)) {
        $group = newFileGroup $key $relDir $book
    }
    $group.Hits.Add([pscustomobject]@{
        Root = "C:\tebunko\work\index"; RelPath = "$relDir\$book\$location.tsv"; RelDir = $relDir
        FileName = "$location.tsv"; Book = $book; Location = $location; LineNumber = $lineNumber; Line = $line
    })
    [void]$group.AddLocation($location)
    addFileGroupLocation $group $book $location
    [void]$script:dirtyGroups.Add($group)
    $script:hitCount++
    return $group
}

function getItemNames {
    # 表に並んでいる項目を、見出しは「#ファイル名」、行は「ファイル名:行番号」で返す
    return @($script:resultItems | ForEach-Object {
        if ($_ -is [FileGroup]) { "#$($_.Book)" } else { "$($_.Book):$($_.LineNumber)" }
    })
}

function newSortColumn([string]$path) {
    $column = New-Object System.Windows.Controls.DataGridTextColumn
    $column.SortMemberPath = $path
    return $column
}

Describe "読み込み" -Tag Unit {
    It "表に結果の一覧をつなぎ、イベントを登録する" {
        [object]::ReferenceEquals($ui.ResultGrid.ItemsSource, $script:resultItems) | Should Be $true
        $handlers.ContainsKey("LoadingRow") | Should Be $true
        $handlers.ContainsKey("Sorting") | Should Be $true
        $handlers["PreviewMouseUp"] -is [System.Windows.Input.MouseButtonEventHandler] | Should Be $true
        $handlers.ContainsKey("ExpandAll.Click") | Should Be $true
        $handlers.ContainsKey("CollapseAll.Click") | Should Be $true
    }
}

Describe "clearResults" -Tag Unit {
    BeforeEach { resetResults }

    It "結果・並べ替えの印を空にし、強調に使う検索ワードを覚える" {
        [void](addHit "見積.xlsx" "4月" "`t見積")
        $column = newSortColumn "Book"
        $column.SortDirection = [System.ComponentModel.ListSortDirection]::Ascending
        $ui.ResultGrid.Columns.Add($column)
        $script:expandNew = $true

        clearResults "請求" ([regex]"請求")

        $script:hitCount | Should Be 0
        $script:fileGroups.Count | Should Be 0
        $script:groupList.Count | Should Be 0
        $script:dirtyGroups.Count | Should Be 0
        $script:resultItems.Count | Should Be 0
        $script:expandNew | Should Be $false
        $script:rowWord | Should Be "請求"
        $script:rowPattern.ToString() | Should Be "請求"
        $null -eq $column.SortDirection | Should Be $true
    }
}

Describe "newFileGroup" -Tag Unit {
    BeforeEach { resetResults }

    It "見つかった順に番号を付け、アプリの種類を決める" {
        $first = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        $second = newFileGroup "C:\共有\議事録.docx" "" "議事録.docx"

        $first.Order | Should Be 0
        $second.Order | Should Be 1
        $first.AppKind | Should Be "Excel"
        $second.AppKind | Should Be "Word"
        $second.FullPath | Should Be "C:\共有\議事録.docx"
        $script:groupList.Count | Should Be 2
    }

    It "フルパスは大文字と小文字を区別せずに引ける" {
        $group = newFileGroup "C:\共有\Mitsumori.xlsx" "" "Mitsumori.xlsx"
        $script:fileGroups["c:\共有\MITSUMORI.XLSX"] | Should Be $group
    }

    It "［すべて展開］のあとに見つかったファイルは開いておく" {
        $script:expandNew = $true
        (newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx").IsExpanded | Should Be $true
    }
}

Describe "getPlace・addFileGroupLocation" -Tag Unit {
    BeforeEach { resetResults }

    It "同じ種類のファイル・同じ場所の表記は 1 回だけ作る" {
        $a = getPlace "見積.xlsx" "4月"
        $b = getPlace "請求.xlsm" "4月"
        $a.Place | Should Be "[シート] 4月"
        [object]::ReferenceEquals($a, $b) | Should Be $true
        $script:places.Count | Should Be 1
    }

    It "Excel とそれ以外は別の表記にする" {
        (getPlace "見積.xlsx" "ページ003").Place | Should Be "[シート] ページ003"
        (getPlace "議事録.docx" "ページ003").Place | Should Be "[ページ] 3（目安）"
        $script:places.Count | Should Be 2
    }

    It "見出しの右端に、ヒットした場所を足していく" {
        $group = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        addFileGroupLocation $group "見積.xlsx" "4月"
        $group.LocationText | Should Be "[シート] 4月"
        addFileGroupLocation $group "見積.xlsx" "5月"
        addFileGroupLocation $group "見積.xlsx" "6月"
        $group.LocationText | Should Be "[シート] 4月 ほか 2 か所"
    }

    It "図形の場所は元のシートと同じ表記なので増やさない" {
        $group = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        addFileGroupLocation $group "見積.xlsx" "4月"
        addFileGroupLocation $group "見積.xlsx" "4月[図形]"
        $group.LocationText | Should Be "[シート] 4月"
        @($group.GetLocations()).Count | Should Be 1
    }
}

Describe "ensureRows" -Tag Unit {
    BeforeEach { resetResults }

    It "ヒットを表の行にし、場所・種別・見出しを入れる" {
        $group = addHit "見積.xlsx" "4月" "`t見積書" 3
        [void](addHit "見積.xlsx" "4月[図形]" "B2`t見積の注記" 5)

        ensureRows $group

        $group.Rows.Count | Should Be 2
        $row = $group.Rows[0]
        $row.IndexName | Should Be "営業部"
        $row.PlaceText | Should Be "[シート] 4月"
        $row.Kind | Should Be "セル"
        $row.LineNumber | Should Be 3
        $row.Order | Should Be 0
        $row.FileGroup | Should Be $group
        $group.Rows[1].Kind | Should Be "図形"
        $group.Rows[1].Order | Should Be 1
        $group.ShownRows.Count | Should Be 2
    }

    It "相対フォルダが空・無いときは、インデックス名も空にする" {
        $group = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        $group.Hits.Add([pscustomobject]@{ Root = "C:\tebunko\work\index"; RelPath = "見積.xlsx\4月.tsv"; RelDir = $null
            FileName = "4月.tsv"; Book = "見積.xlsx"; Location = "4月"; LineNumber = 1; Line = "`t見積" })

        ensureRows $group

        $group.Rows[0].IndexName | Should Be ""
        $group.Rows[0].RelDir | Should Be ""
    }

    It "インデックス直下のファイルは、相対パスそのものがインデックス名" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1 "営業部"
        ensureRows $group
        $group.Rows[0].IndexName | Should Be "営業部"
    }

    It "作り済みの行は作り直さず、増えた分だけ足す" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        ensureRows $group
        $first = $group.Rows[0]
        [void](addHit "見積.xlsx" "4月" "`t見積2" 2)

        ensureRows $group

        $group.Rows.Count | Should Be 2
        [object]::ReferenceEquals($group.Rows[0], $first) | Should Be $true
    }

    It "絞り込み中は、合う行だけを ShownRows に入れる" {
        $group = addHit "見積.xlsx" "4月" "`t見積 山田" 1
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)
        $script:filterText = "佐藤"

        ensureRows $group

        $group.Rows.Count | Should Be 2
        $group.ShownRows.Count | Should Be 1
        $group.ShownRows[0].LineNumber | Should Be 2
    }
}

Describe "updateGroupCount" -Tag Unit {
    BeforeEach { resetResults }

    It "絞り込んでいなければ、行を作っていなくてもヒットの数" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        updateGroupCount $group
        $group.ShownCount | Should Be 2
        $group.Rows.Count | Should Be 0
    }

    It "絞り込み中は合う行の数" {
        $group = addHit "見積.xlsx" "4月" "`t見積 山田" 1
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)
        $script:filterText = "山田"
        ensureRows $group
        updateGroupCount $group
        $group.ShownCount | Should Be 1
    }
}

Describe "flushResults" -Tag Unit {
    BeforeEach { resetResults }

    It "検索直後は見出しだけを末尾に足す" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "見積.xlsx" "5月" "`t見積" 2)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)

        flushResults

        (getItemNames) -join "," | Should Be "#見積.xlsx,#議事録.docx"
        $script:fileGroups["C:\共有\営業部\2024\見積.xlsx"].ShownCount | Should Be 2
        $script:dirtyGroups.Count | Should Be 0
        $script:lastInViewOrder | Should Be 1
    }

    It "開いているファイルは、あとから来た行も見出しの下に足す" {
        $script:expandNew = $true
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)
        flushResults
        [void](addHit "見積.xlsx" "4月" "`t見積" 7)

        flushResults

        (getItemNames) -join "," | Should Be "#見積.xlsx,見積.xlsx:1,見積.xlsx:7,#議事録.docx,議事録.docx:1"
        $script:fileGroups["C:\共有\営業部\2024\見積.xlsx"].DisplayedCount | Should Be 2
    }

    It "絞り込みに合う行が無いファイルは見出しも出さない" {
        $script:filterText = "佐藤"
        [void](addHit "見積.xlsx" "4月" "`t見積 山田" 1)
        [void](addHit "議事録.docx" "ページ001" "見積 佐藤" 1)

        flushResults

        (getItemNames) -join "," | Should Be "#議事録.docx"
    }

    It "隠れていたファイルに、あとから合う行が来たら作り直しを予約する" {
        $script:filterText = "佐藤"
        [void](addHit "見積.xlsx" "4月" "`t見積 山田" 1)
        [void](addHit "議事録.docx" "ページ001" "見積 佐藤" 1)
        flushResults
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)

        flushResults

        $script:needsRebuild | Should Be $true
        (getItemNames) -join "," | Should Be "#議事録.docx"
    }
}

Describe "setResultItems" -Tag Unit {
    BeforeEach { resetResults }

    It "選んでいた項目が残っていれば選び直して見えるようにする" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        flushResults
        $ui.ResultGrid.SelectedItem = $group

        setResultItems ([System.Collections.Generic.List[object]]@($group))

        [object]::ReferenceEquals($ui.ResultGrid.ItemsSource, $script:resultItems) | Should Be $true
        $ui.ResultGrid.SelectedItem | Should Be $group
        $ui.ResultGrid.Scrolled.Count | Should Be 1
    }

    It "選んでいた項目が無くなったら選び直さない" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        flushResults
        $ui.ResultGrid.SelectedItem = $group

        setResultItems (New-Object 'System.Collections.Generic.List[object]')

        $group.InView | Should Be $false
        $script:lastInViewOrder | Should Be -1
        $ui.ResultGrid.Scrolled.Count | Should Be 0
    }

    It "開いている見出しは、下に入れた行の数を覚える" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        $group.SetExpanded($true)
        ensureRows $group
        updateGroupCount $group

        setResultItems (getResultItems $script:groupList)

        $group.InView | Should Be $true
        $group.DisplayedCount | Should Be 2
        $script:lastInViewOrder | Should Be 0
    }
}

Describe "toggleFileGroup" -Tag Unit {
    BeforeEach { resetResults }

    It "開くと見出しの下に行を入れ、閉じると取り除く" {
        $mitsumori = addHit "見積.xlsx" "4月" "`t見積" 1
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)
        flushResults

        toggleFileGroup $mitsumori
        $mitsumori.IsExpanded | Should Be $true
        (getItemNames) -join "," | Should Be "#見積.xlsx,見積.xlsx:1,見積.xlsx:2,#議事録.docx"

        toggleFileGroup $mitsumori
        $mitsumori.IsExpanded | Should Be $false
        $mitsumori.DisplayedCount | Should Be 0
        (getItemNames) -join "," | Should Be "#見積.xlsx,#議事録.docx"
    }

    It "行が多いときは、1 件ずつ入れずに表を作り直す" {
        $group = $null
        for ($i = 1; $i -le ${rebuildThreshold} + 1; $i++) {
            $group = addHit "見積.xlsx" "4月" "`t見積" $i
        }
        flushResults

        toggleFileGroup $group
        $script:resultItems.Count | Should Be (${rebuildThreshold} + 2)
        $group.DisplayedCount | Should Be (${rebuildThreshold} + 1)

        toggleFileGroup $group
        $script:resultItems.Count | Should Be 1
        $group.IsExpanded | Should Be $false
    }
}

Describe "setAllFileGroupsExpanded" -Tag Unit {
    BeforeEach { resetResults }

    It "すべて開き、このあと見つかるファイルも開いておく" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)
        flushResults

        setAllFileGroupsExpanded $true

        $script:expandNew | Should Be $true
        (getItemNames) -join "," | Should Be "#見積.xlsx,見積.xlsx:1,#議事録.docx,議事録.docx:1"

        setAllFileGroupsExpanded $false
        $script:expandNew | Should Be $false
        (getItemNames) -join "," | Should Be "#見積.xlsx,#議事録.docx"
    }

    It "［すべて展開］［すべて折りたたむ］のボタンから切り替える" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        flushResults

        & $handlers["ExpandAll.Click"]
        $script:resultItems.Count | Should Be 2

        & $handlers["CollapseAll.Click"]
        $script:resultItems.Count | Should Be 1
    }
}

Describe "applyResultFilter" -Tag Unit {
    BeforeEach { resetResults }

    It "絞り込むと合う行のあるファイルだけを出し、件数も合う行の数にする" {
        $mitsumori = addHit "見積.xlsx" "4月" "`t見積 山田" 1
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)
        [void](addHit "議事録.docx" "ページ001" "見積 山田" 1)
        flushResults

        $script:filterText = "佐藤"
        applyResultFilter "佐藤"

        (getItemNames) -join "," | Should Be "#見積.xlsx"
        $mitsumori.ShownCount | Should Be 1
        (getShownHitCount) | Should Be 1

        $script:filterText = ""
        applyResultFilter ""

        (getItemNames) -join "," | Should Be "#見積.xlsx,#議事録.docx"
        $mitsumori.ShownCount | Should Be 2
        (getShownHitCount) | Should Be 3
    }
}

Describe "applyResultFilter（絞り込みの文字）" -Tag Unit {
    BeforeEach { resetResults }

    It "大文字と小文字を区別しない" {
        [void](addHit "Mitsumori.xlsx" "4月" "`tEstimate" 1)
        [void](addHit "議事録.docx" "ページ001" "見積" 1)
        flushResults

        $script:filterText = "ESTIMATE"
        applyResultFilter "ESTIMATE"

        (getItemNames) -join "," | Should Be "#Mitsumori.xlsx"
    }

    It "ワイルドカード・正規表現の記号も文字どおりに探す" {
        [void](addHit "見積.xlsx" "4月[図形]" "B2`t見積*" 1)
        [void](addHit "見積.xlsx" "4月" "`t見積A" 2)
        flushResults

        $script:filterText = "[図形]"
        applyResultFilter "[図形]"
        (getShownHitCount) | Should Be 1

        $script:filterText = "*"
        applyResultFilter "*"
        (getShownHitCount) | Should Be 1

        $script:filterText = "?"
        applyResultFilter "?"
        (getShownHitCount) | Should Be 0
    }

    It "場所・種別の表示でも絞り込める" {
        [void](addHit "見積.xlsx" "4月[コメント]" "B2`t見積" 1)
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        flushResults

        $script:filterText = "コメント"
        applyResultFilter "コメント"

        (getShownHitCount) | Should Be 1
    }
}

Describe "sortResults・applySort" -Tag Unit {
    BeforeEach { resetResults }

    It "並べ替えの項目が無い列では何もしない" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        flushResults
        $column = newSortColumn ""
        $ui.ResultGrid.Columns.Add($column)

        sortResults $column

        $null -eq $column.SortDirection | Should Be $true
        $script:groupList[0].Rows.Count | Should Be 0
    }

    It "1 回目は昇順、もう一度で降順にし、ほかの列の印を消す" {
        [void](addHit "b.docx" "ページ001" "見積" 1)
        [void](addHit "a.docx" "ページ001" "見積" 1)
        [void](addHit "c.docx" "ページ001" "見積" 1)
        flushResults
        $other = newSortColumn "LineNumber"
        $other.SortDirection = [System.ComponentModel.ListSortDirection]::Ascending
        $column = newSortColumn "Book"
        $ui.ResultGrid.Columns.Add($other)
        $ui.ResultGrid.Columns.Add($column)

        sortResults $column
        $column.SortDirection | Should Be ([System.ComponentModel.ListSortDirection]::Ascending)
        $null -eq $other.SortDirection | Should Be $true
        (getItemNames) -join "," | Should Be "#a.docx,#b.docx,#c.docx"

        sortResults $column
        $column.SortDirection | Should Be ([System.ComponentModel.ListSortDirection]::Descending)
        (getItemNames) -join "," | Should Be "#c.docx,#b.docx,#a.docx"

        sortResults $column
        $column.SortDirection | Should Be ([System.ComponentModel.ListSortDirection]::Ascending)
        (getItemNames) -join "," | Should Be "#a.docx,#b.docx,#c.docx"
    }

    It "同じ値のファイル・行は、見つかった順のまま並べる" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 5
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        [void](addHit "議事録.docx" "ページ001" "見積" 1)
        [void](addHit "規程.docx" "ページ001" "見積" 1)
        $group.SetExpanded($true)
        flushResults

        applySort "IndexName" $false
        (getItemNames) -join "," | Should Be "#見積.xlsx,見積.xlsx:5,見積.xlsx:2,#議事録.docx,#規程.docx"

        applySort "IndexName" $true
        (getItemNames) -join "," | Should Be "#見積.xlsx,見積.xlsx:5,見積.xlsx:2,#議事録.docx,#規程.docx"
    }

    It "ファイルの中の行も並べ替え、絞り込みを当て直す" {
        $group = addHit "見積.xlsx" "4月" "`t見積 山田" 1
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)
        [void](addHit "見積.xlsx" "4月" "`t見積 山田" 3)
        $group.SetExpanded($true)
        $script:filterText = "山田"
        flushResults

        applySort "LineNumber" $true

        (getItemNames) -join "," | Should Be "#見積.xlsx,見積.xlsx:3,見積.xlsx:1"
    }

    It "表の並べ替えのイベントは、表に任せずファイルごとに並べ替える" {
        [void](addHit "b.docx" "ページ001" "見積" 1)
        [void](addHit "a.docx" "ページ001" "見積" 1)
        flushResults
        $column = newSortColumn "Book"
        $ui.ResultGrid.Columns.Add($column)
        $e = [pscustomobject]@{ Column = $column; Handled = $false }

        & $handlers["Sorting"] $ui.ResultGrid $e

        $e.Handled | Should Be $true
        (getItemNames) -join "," | Should Be "#a.docx,#b.docx"
    }
}

Describe "finishResults" -Tag Unit {
    BeforeEach { resetResults }

    It "検索中に並べ替えていたら、あとから来た行も含めて並べ直す" {
        [void](addHit "b.docx" "ページ001" "見積" 1)
        flushResults
        $column = newSortColumn "Book"
        $column.SortDirection = [System.ComponentModel.ListSortDirection]::Ascending
        $ui.ResultGrid.Columns.Add($column)
        [void](addHit "a.docx" "ページ001" "見積" 1)

        finishResults

        (getItemNames) -join "," | Should Be "#a.docx,#b.docx"
    }

    It "隠れていたファイルに合う行が来ていたら作り直す" {
        $script:filterText = "佐藤"
        [void](addHit "見積.xlsx" "4月" "`t見積 山田" 1)
        [void](addHit "議事録.docx" "ページ001" "見積 佐藤" 1)
        flushResults
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)

        finishResults

        $script:needsRebuild | Should Be $false
        (getItemNames) -join "," | Should Be "#見積.xlsx,#議事録.docx"
    }

    It "どちらでもなければ表はそのまま" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        flushResults
        $items = $script:resultItems

        finishResults

        [object]::ReferenceEquals($script:resultItems, $items) | Should Be $true
    }
}

Describe "getCurrentHitRow" -Tag Unit {
    BeforeEach { resetResults }

    It "見出しを選んでいるときは、そのファイルの先頭の行" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 4
        flushResults
        $ui.ResultGrid.SelectedItem = $group

        (getCurrentHitRow).LineNumber | Should Be 4
    }

    It "見出しの行がすべて絞り込みで隠れていれば無し" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 4
        $script:filterText = "佐藤"
        $ui.ResultGrid.SelectedItem = $group

        $null -eq (getCurrentHitRow) | Should Be $true
    }

    It "行を選んでいるときはその行、何も選んでいなければ無し" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 4
        ensureRows $group
        $ui.ResultGrid.SelectedItem = $group.Rows[0]
        (getCurrentHitRow) | Should Be $group.Rows[0]

        $ui.ResultGrid.SelectedItem = $null
        $null -eq (getCurrentHitRow) | Should Be $true
    }
}

Describe "getViewRows・getSelectedRows" -Tag Unit {
    BeforeEach { resetResults }

    It "表示中の行は、閉じているファイルの行も含めて表の順に返す" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "議事録.docx" "ページ001" "見積" 2)
        flushResults

        $rows = getViewRows

        $rows.Count | Should Be 2
        $rows[0].Book | Should Be "見積.xlsx"
        $rows[1].Book | Should Be "議事録.docx"
    }

    It "結果が無い・何も選んでいないときは空の配列を返す" {
        $view = getViewRows
        $selected = getSelectedRows

        , $view | Should BeOfType [object[]]
        $view.Count | Should Be 0
        , $selected | Should BeOfType [object[]]
        $selected.Count | Should Be 0
        (getShownHitCount) | Should Be 0
    }

    It "見出しを選ぶとそのファイルの行をすべて、行を選ぶとその行を、表の順に返す" {
        $mitsumori = addHit "見積.xlsx" "4月" "`t見積" 1
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        $gijiroku = addHit "議事録.docx" "ページ001" "見積" 3
        [void](addHit "議事録.docx" "ページ001" "見積" 4)
        ensureRows $gijiroku
        $ui.ResultGrid.SelectedItems = @($gijiroku.Rows[1], $mitsumori)

        $rows = getSelectedRows

        (@($rows | ForEach-Object { $_.LineNumber }) -join ",") | Should Be "1,2,4"
    }
}

Describe "イベント" -Tag Unit {
    BeforeEach { resetResults }

    It "行が画面に出るときに、表示用の中身を作る" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        ensureRows $group
        $row = $group.Rows[0]
        $e = [pscustomobject]@{ Row = [pscustomobject]@{ Item = $row } }

        & $handlers["LoadingRow"] $ui.ResultGrid $e

        $row.Prepared | Should Be $true
        $row.MatchCell | Should Be "B1"
    }

    It "見出しを出すときは何もしない" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        $e = [pscustomobject]@{ Row = [pscustomobject]@{ Item = $group } }

        { & $handlers["LoadingRow"] $ui.ResultGrid $e } | Should Not Throw
    }
}
