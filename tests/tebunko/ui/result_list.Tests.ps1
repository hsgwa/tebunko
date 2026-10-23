# ［検索］の結果の表（tebunko\ui\result_list.ps1）のテスト。
# 画面の部品（$ui.ResultGrid など）は偽物にして、見出し・行の出し入れと状態の変化を確かめる。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_types.ps1"
    . "${scriptsDir}\tebunko\ui\search\search_bar_view.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\search\open_source_view.ps1"

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

    # 結果の表とプレビューの境目を決める処理（preview.ps1）の代わり。呼ばれたときの行数を控える
    $script:paneHeightCalls = @()
    function updateResultPaneHeight {
        param ([int]$itemCount = 0)
        $script:paneHeightCalls += $itemCount
    }

    . "${scriptsDir}\tebunko\ui\result_list.ps1"

    # ---- テストの準備 ----

    function resetResults {
        $script:paneHeightCalls = @()
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

        $script:hitCount | Should -Be 0
        $script:fileGroups.Count | Should -Be 0
        $script:groupList.Count | Should -Be 0
        $script:dirtyGroups.Count | Should -Be 0
        $script:resultItems.Count | Should -Be 0
        $script:expandNew | Should -Be $false
        $script:rowWord | Should -Be "請求"
        $script:rowPattern.ToString() | Should -Be "請求"
        $null -eq $column.SortDirection | Should -Be $true
    }
}

Describe "newFileGroup" -Tag Unit {
    BeforeEach { resetResults }

    It "見つかった順に番号を付け、アプリの種類を決める" {
        $first = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        $second = newFileGroup "C:\共有\議事録.docx" "" "議事録.docx"

        $first.Order | Should -Be 0
        $second.Order | Should -Be 1
        $first.AppKind | Should -Be "Excel"
        $second.AppKind | Should -Be "Word"
        $second.FullPath | Should -Be "C:\共有\議事録.docx"
        $script:groupList.Count | Should -Be 2
    }

    It "フルパスは大文字と小文字を区別せずに引ける" {
        $group = newFileGroup "C:\共有\Mitsumori.xlsx" "" "Mitsumori.xlsx"
        $script:fileGroups["c:\共有\MITSUMORI.XLSX"] | Should -Be $group
    }

    It "表に出す値（BookDisplay・RelDirDisplay・FullPathDisplay）は、向きを変える文字・制御文字を置き換え、開く・コピー用の値は元のまま" {
        $rlo = [string][char]0x202E
        $group = newFileGroup "C:\共有\営業${rlo}\見積${rlo}slx.xlsx" "営業${rlo}" "見積${rlo}slx.xlsx"

        $group.BookDisplay | Should -Be "見積$([char]0xFFFD)slx.xlsx"
        $group.RelDirDisplay | Should -Be "営業$([char]0xFFFD)"
        $group.FullPathDisplay | Should -Be "C:\共有\営業$([char]0xFFFD)\見積$([char]0xFFFD)slx.xlsx"
        $group.Book | Should -Be "見積${rlo}slx.xlsx"
        $group.FullPath | Should -Be "C:\共有\営業${rlo}\見積${rlo}slx.xlsx"
    }

    It "［すべて開く］のあとに見つかったファイルは開いておく" {
        $script:expandNew = $true
        (newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx").IsExpanded | Should -Be $true
    }
}

Describe "getPlace・addFileGroupLocation" -Tag Unit {
    BeforeEach { resetResults }

    It "Excel とそれ以外は別の表記にする" {
        (getPlace "見積.xlsx" "ページ003" "Excel").Place | Should -Be "[シート]ページ003"
        (getPlace "議事録.docx" "ページ003" "Word").Place | Should -Be "3 ページ（目安）"
        $script:places.Count | Should -Be 2
    }

    It "見出しの右端に、ヒットした場所を足していく" {
        $group = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        addFileGroupLocation $group "見積.xlsx" "4月"
        $group.LocationText | Should -Be "[シート]4月"
        $group.LocationPrefix | Should -Be ""
        addFileGroupLocation $group "見積.xlsx" "5月"
        addFileGroupLocation $group "見積.xlsx" "6月"
        $group.LocationText | Should -Be "[シート]4月 ほか 2 か所"
        $group.LocationPrefix | Should -Be "[シート]4月 ほか 2 か所 ・ "
    }

    It "見出しの右端の場所は、シート名の向きを変える文字を置き換える" {
        $rlo = [string][char]0x202E
        $group = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        addFileGroupLocation $group "見積.xlsx" "4月${rlo}"
        addFileGroupLocation $group "見積.xlsx" "5月"
        $group.LocationText | Should -Be "[シート]4月$([char]0xFFFD) ほか 1 か所"
        $group.LocationPrefix | Should -Be "[シート]4月$([char]0xFFFD) ほか 1 か所 ・ "
        $group.LocationText | Should -Not -Match $rlo
    }

    It "図形の場所は元のシートと同じ表記なので増やさない" {
        $group = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        addFileGroupLocation $group "見積.xlsx" "4月"
        addFileGroupLocation $group "見積.xlsx" "4月[図形]"
        $group.LocationText | Should -Be "[シート]4月"
        @($group.GetLocations()).Count | Should -Be 1
    }
}

Describe "ensureRows" -Tag Unit {
    BeforeEach { resetResults }

    It "ヒットを表の行にし、場所・種別・見出しを入れる" {
        $group = addHit "見積.xlsx" "4月" "`t見積書" 3
        [void](addHit "見積.xlsx" "4月[図形]" "B2`t見積の注記" 5)
        [void](addHit "見積.xlsx" "4月[ヘッダー・フッター]" "見積の件" 1)

        ensureRows $group

        $group.Rows.Count | Should -Be 3
        $row = $group.Rows[0]
        $row.IndexName | Should -Be "営業部"
        $row.PlaceText | Should -Be "[シート]4月"
        $row.Kind | Should -Be "セル"
        $row.LineNumber | Should -Be 3
        $row.Order | Should -Be 0
        $row.FileGroup | Should -Be $group
        $group.Rows[1].Kind | Should -Be "図形"
        $group.Rows[1].Order | Should -Be 1
        $group.Rows[2].Kind | Should -Be "ヘッダー・フッター"
        $group.Rows[2].Order | Should -Be 2
        $group.ShownRows.Count | Should -Be 3
        # Excel のファイルの行か（行を作る側が決める）
        $group.Rows[0].IsExcel | Should -Be $true
        # 行の先頭がセル番地か（Excel の図形・コメントだけ。ヘッダー・フッターとふつうのセルは違う）
        $group.Rows[0].IsCellPrefixed | Should -Be $false
        $group.Rows[1].IsCellPrefixed | Should -Be $true
        $group.Rows[2].IsCellPrefixed | Should -Be $false
    }

    It "Excel 以外のファイルの行は、IsExcel も IsCellPrefixed も偽" {
        $group = addHit "議事録.docx" "ページ003[図形]" "見積の件" 1
        ensureRows $group
        $group.Rows[0].IsExcel | Should -Be $false
        $group.Rows[0].IsCellPrefixed | Should -Be $false
    }

    It "Excel の <Location> の行は、Prepare 後に一致したセルの番地 = 「<Cell>」、該当行 = 「<Shown>」になる" -ForEach @(
        @{ Location = "4月[図形]"; Line = "B2`t見積の注記"; Cell = "B2"; Shown = "見積の注記" }
        @{ Location = "4月[コメント]"; Line = "C3`t見積の確認"; Cell = "C3"; Shown = "見積の確認" }
        @{ Location = "4月[ヘッダー・フッター]"; Line = "見積の件"; Cell = ""; Shown = "見積の件" }
    ) {
        $group = addHit "見積.xlsx" $Location $Line 1
        ensureRows $group

        $group.Rows[0].Prepare()

        $group.Rows[0].MatchCell | Should -Be $Cell
        $group.Rows[0].DisplayLine | Should -Be $Shown
    }

    It "相対フォルダが空・無いときは、インデックス名も空にする" {
        $group = newFileGroup "C:\共有\見積.xlsx" "" "見積.xlsx"
        $group.Hits.Add([pscustomobject]@{ Root = "C:\tebunko\work\index"; RelPath = "見積.xlsx\4月.tsv"; RelDir = $null
            FileName = "4月.tsv"; Book = "見積.xlsx"; Location = "4月"; LineNumber = 1; Line = "`t見積" })

        ensureRows $group

        $group.Rows[0].IndexName | Should -Be ""
        $group.Rows[0].RelDir | Should -Be ""
    }

    It "インデックス直下のファイルは、相対パスそのものがインデックス名" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1 "営業部"
        ensureRows $group
        $group.Rows[0].IndexName | Should -Be "営業部"
    }

    It "作り済みの行は作り直さず、増えた分だけ足す" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        ensureRows $group
        $first = $group.Rows[0]
        [void](addHit "見積.xlsx" "4月" "`t見積2" 2)

        ensureRows $group

        $group.Rows.Count | Should -Be 2
        [object]::ReferenceEquals($group.Rows[0], $first) | Should -Be $true
    }

    It "絞り込み中は、ファイル名が合えばそのファイルの行をすべて ShownRows に入れ、合わなければ入れない" {
        $group = addHit "見積.xlsx" "4月" "`t見積 山田" 1
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)
        $script:filterText = "見積"

        ensureRows $group

        $group.Rows.Count | Should -Be 2
        $group.ShownRows.Count | Should -Be 2

        $other = addHit "議事録.docx" "ページ001" "見積 佐藤" 1
        ensureRows $other

        $other.Rows.Count | Should -Be 1
        $other.ShownRows.Count | Should -Be 0
    }
}

Describe "updateGroupCount" -Tag Unit {
    BeforeEach { resetResults }

    It "絞り込んでいなければ、行を作っていなくてもヒットの数" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        updateGroupCount $group
        $group.ShownCount | Should -Be 2
        $group.Rows.Count | Should -Be 0
    }

    It "絞り込み中は、ファイル名が合えば<expected>（絞り込み: <text>）" -TestCases @(
        @{ text = "見積"; expected = 2 }
        @{ text = "山田"; expected = 0 }
    ) {
        param ($text, $expected)
        $group = addHit "見積.xlsx" "4月" "`t見積 山田" 1
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)
        $script:filterText = $text
        ensureRows $group
        updateGroupCount $group
        $group.ShownCount | Should -Be $expected
    }
}

Describe "flushResults" -Tag Unit {
    BeforeEach { resetResults }

    It "検索直後は見出しだけを末尾に足す" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "見積.xlsx" "5月" "`t見積" 2)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)

        flushResults

        (getItemNames) -join "," | Should -Be "#見積.xlsx,#議事録.docx"
        $script:fileGroups["C:\共有\営業部\2024\見積.xlsx"].ShownCount | Should -Be 2
        $script:dirtyGroups.Count | Should -Be 0
    }

    It "増えたヒットがある回だけ、そのときの行数で、表とプレビューの境目を決め直す" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)
        flushResults
        $script:paneHeightCalls | Should -Be @(2)

        flushResults
        $script:paneHeightCalls | Should -Be @(2) -Because "増えたヒットが無い回は決め直さない"
    }

    It "開いているファイルは、あとから来た行も見出しの下に足す" {
        $script:expandNew = $true
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)
        flushResults
        [void](addHit "見積.xlsx" "4月" "`t見積" 7)

        flushResults

        (getItemNames) -join "," | Should -Be "#見積.xlsx,見積.xlsx:1,見積.xlsx:7,#議事録.docx,議事録.docx:1"
        $script:fileGroups["C:\共有\営業部\2024\見積.xlsx"].DisplayedCount | Should -Be 2
    }

    It "ファイル名が合わないファイルは見出しも出さない" {
        $script:filterText = "議事録"
        [void](addHit "見積.xlsx" "4月" "`t議事録 山田" 1)
        [void](addHit "議事録.docx" "ページ001" "見積 佐藤" 1)

        flushResults

        (getItemNames) -join "," | Should -Be "#議事録.docx"
    }

    It "隠れているファイルに、あとから該当行に絞り込みの語がある行が来ても、隠れたまま" {
        $script:filterText = "議事録"
        [void](addHit "見積.xlsx" "4月" "`t見積 山田" 1)
        [void](addHit "議事録.docx" "ページ001" "見積 佐藤" 1)
        flushResults
        [void](addHit "見積.xlsx" "4月" "`t議事録 佐藤" 2)

        flushResults

        (getItemNames) -join "," | Should -Be "#議事録.docx"
    }
}

Describe "setResultItems" -Tag Unit {
    BeforeEach { resetResults }

    It "選んでいた項目が残っていれば選び直して見えるようにする" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        flushResults
        $ui.ResultGrid.SelectedItem = $group

        setResultItems ([System.Collections.Generic.List[object]]@($group))

        [object]::ReferenceEquals($ui.ResultGrid.ItemsSource, $script:resultItems) | Should -Be $true
        $ui.ResultGrid.SelectedItem | Should -Be $group
        $ui.ResultGrid.Scrolled.Count | Should -Be 1
    }

    It "選んでいた項目が無くなったら選び直さない" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        flushResults
        $ui.ResultGrid.SelectedItem = $group

        setResultItems (New-Object 'System.Collections.Generic.List[object]')

        $group.InView | Should -Be $false
        $ui.ResultGrid.Scrolled.Count | Should -Be 0
    }

    It "開いている見出しは、下に入れた行の数を覚える" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 1
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        $group.SetExpanded($true)
        ensureRows $group
        updateGroupCount $group

        setResultItems (getResultItems $script:groupList)

        $group.InView | Should -Be $true
        $group.DisplayedCount | Should -Be 2
    }
}

Describe "toggleFileGroup" -Tag Unit {
    BeforeEach { resetResults }

    It "開くと見出しの下に行を入れ、閉じると取り除く" {
        $mitsumori = addHit "見積.xlsx" "4月" "`t見積" 1
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)
        flushResults

        $script:paneHeightCalls = @()
        toggleFileGroup $mitsumori
        $mitsumori.IsExpanded | Should -Be $true
        (getItemNames) -join "," | Should -Be "#見積.xlsx,見積.xlsx:1,見積.xlsx:2,#議事録.docx"
        $script:paneHeightCalls | Should -Be @(4) -Because "開いたあとの行数で、表とプレビューの境目を決め直す"

        toggleFileGroup $mitsumori
        $mitsumori.IsExpanded | Should -Be $false
        $mitsumori.DisplayedCount | Should -Be 0
        (getItemNames) -join "," | Should -Be "#見積.xlsx,#議事録.docx"
        $script:paneHeightCalls | Should -Be @(4, 2) -Because "閉じたあとの行数で、決め直す"
    }

    It "行が多いときは、1 件ずつ入れずに表を作り直す" {
        $group = $null
        for ($i = 1; $i -le ${rebuildThreshold} + 1; $i++) {
            $group = addHit "見積.xlsx" "4月" "`t見積" $i
        }
        flushResults

        toggleFileGroup $group
        $script:resultItems.Count | Should -Be (${rebuildThreshold} + 2)
        $group.DisplayedCount | Should -Be (${rebuildThreshold} + 1)

        toggleFileGroup $group
        $script:resultItems.Count | Should -Be 1
        $group.IsExpanded | Should -Be $false
    }
}

Describe "setAllFileGroupsExpanded" -Tag Unit {
    BeforeEach { resetResults }

    It "すべて開き、このあと見つかるファイルも開いておく" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "議事録.docx" "ページ001" "見積の件" 1)
        flushResults

        setAllFileGroupsExpanded $true

        $script:expandNew | Should -Be $true
        (getItemNames) -join "," | Should -Be "#見積.xlsx,見積.xlsx:1,#議事録.docx,議事録.docx:1"

        setAllFileGroupsExpanded $false
        $script:expandNew | Should -Be $false
        (getItemNames) -join "," | Should -Be "#見積.xlsx,#議事録.docx"
    }

    It "［すべて開く］［すべて折りたたむ］のボタンから切り替える" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        flushResults

        & $handlers["ExpandAll.Click"]
        $script:resultItems.Count | Should -Be 2

        & $handlers["CollapseAll.Click"]
        $script:resultItems.Count | Should -Be 1
    }
}

Describe "applyResultFilter" -Tag Unit {
    BeforeEach { resetResults }

    It "絞り込むとファイル名が合うファイルだけを出し、そのファイルの行は全部残る" {
        $mitsumori = addHit "見積.xlsx" "4月" "`t見積 山田" 1
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)
        [void](addHit "議事録.docx" "ページ001" "見積 山田" 1)
        flushResults

        $script:filterText = "見積.x"
        applyResultFilter "見積.x"

        (getItemNames) -join "," | Should -Be "#見積.xlsx"
        $mitsumori.ShownCount | Should -Be 2
        (getShownHitCount) | Should -Be 2

        $script:filterText = "佐藤"
        applyResultFilter "佐藤"

        (getItemNames) -join "," | Should -Be "" -Because "該当行にだけある語では 0 件"
        (getShownHitCount) | Should -Be 0

        $script:filterText = ""
        applyResultFilter ""

        (getItemNames) -join "," | Should -Be "#見積.xlsx,#議事録.docx"
        $mitsumori.ShownCount | Should -Be 2
        (getShownHitCount) | Should -Be 3
    }
}

Describe "applyResultFilter（絞り込みの文字）" -Tag Unit {
    BeforeEach { resetResults }

    It "大文字と小文字を区別しない" {
        [void](addHit "Mitsumori.xlsx" "4月" "`tEstimate" 1)
        [void](addHit "議事録.docx" "ページ001" "見積" 1)
        flushResults

        $script:filterText = "MITSUMORI"
        applyResultFilter "MITSUMORI"

        (getItemNames) -join "," | Should -Be "#Mitsumori.xlsx"
    }

    It "ワイルドカード・正規表現の記号も文字どおりに探す（ファイル名に無ければ 0 件）" {
        [void](addHit "見積[1].xlsx" "4月" "B2`t見積*" 1)
        [void](addHit "見積.xlsx" "4月" "`t見積A" 2)
        flushResults

        $script:filterText = "[1]"
        applyResultFilter "[1]"
        (getShownHitCount) | Should -Be 1

        $script:filterText = "*"
        applyResultFilter "*"
        (getShownHitCount) | Should -Be 0

        $script:filterText = "?"
        applyResultFilter "?"
        (getShownHitCount) | Should -Be 0
    }

    It "出力に使う行（getShownHitRows）は、絞り込み後の行だけになる" {
        [void](addHit "見積.xlsx" "4月" "`t見積の行" 1)
        [void](addHit "議事録.docx" "ページ001" "議事録の行" 1)
        flushResults

        $script:filterText = "議事録"
        applyResultFilter "議事録"
        @(getViewRows | ForEach-Object { $_.Book }) -join "," | Should -Be "議事録.docx"
    }

    It "<name>では絞り込めない" -TestCases @(
        @{ name = "場所の種類"; text = "コメント" }
        @{ name = "場所"; text = "4月" }
        @{ name = "フォルダ名"; text = "営業部" }
    ) {
        param ($name, $text)
        [void](addHit "見積.xlsx" "4月[コメント]" "B2`t見積" 1)
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        flushResults

        $script:filterText = $text
        applyResultFilter $text
        (getShownHitCount) | Should -Be 0
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

        $null -eq $column.SortDirection | Should -Be $true
        $script:groupList[0].Rows.Count | Should -Be 0
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
        $column.SortDirection | Should -Be ([System.ComponentModel.ListSortDirection]::Ascending)
        $null -eq $other.SortDirection | Should -Be $true
        (getItemNames) -join "," | Should -Be "#a.docx,#b.docx,#c.docx"

        sortResults $column
        $column.SortDirection | Should -Be ([System.ComponentModel.ListSortDirection]::Descending)
        (getItemNames) -join "," | Should -Be "#c.docx,#b.docx,#a.docx"

        sortResults $column
        $column.SortDirection | Should -Be ([System.ComponentModel.ListSortDirection]::Ascending)
        (getItemNames) -join "," | Should -Be "#a.docx,#b.docx,#c.docx"
    }

    It "同じ値のファイル・行は、見つかった順のまま並べる" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 5
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        [void](addHit "議事録.docx" "ページ001" "見積" 1)
        [void](addHit "規程.docx" "ページ001" "見積" 1)
        $group.SetExpanded($true)
        flushResults

        applySort "IndexName" $false
        (getItemNames) -join "," | Should -Be "#見積.xlsx,見積.xlsx:5,見積.xlsx:2,#議事録.docx,#規程.docx"

        applySort "IndexName" $true
        (getItemNames) -join "," | Should -Be "#見積.xlsx,見積.xlsx:5,見積.xlsx:2,#議事録.docx,#規程.docx"
    }

    It "ファイルの中の行も並べ替え、絞り込みを当て直す" {
        $group = addHit "見積.xlsx" "4月" "`t見積 山田" 1
        [void](addHit "見積.xlsx" "4月" "`t見積 佐藤" 2)
        [void](addHit "見積.xlsx" "4月" "`t見積 山田" 3)
        [void](addHit "議事録.docx" "ページ001" "見積" 1)
        $group.SetExpanded($true)
        $script:filterText = "見積"
        flushResults

        applySort "LineNumber" $true

        (getItemNames) -join "," | Should -Be "#見積.xlsx,見積.xlsx:3,見積.xlsx:2,見積.xlsx:1"
    }

    It "表の並べ替えのイベントは、表に任せずファイルごとに並べ替える" {
        [void](addHit "b.docx" "ページ001" "見積" 1)
        [void](addHit "a.docx" "ページ001" "見積" 1)
        flushResults
        $column = newSortColumn "Book"
        $ui.ResultGrid.Columns.Add($column)
        $e = [pscustomobject]@{ Column = $column; Handled = $false }

        & $handlers["Sorting"] $ui.ResultGrid $e

        $e.Handled | Should -Be $true
        (getItemNames) -join "," | Should -Be "#a.docx,#b.docx"
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

        (getItemNames) -join "," | Should -Be "#a.docx,#b.docx"
    }

    It "終わりに、表とプレビューの境目を 1 回だけ決め直す" {
        [void](addHit "a.docx" "ページ001" "見積" 1)
        [void](addHit "b.docx" "ページ001" "見積" 1)

        finishResults

        $script:paneHeightCalls | Should -Be @(2)
    }

    It "隠れていたファイルに該当行だけ絞り込みの語がある行が来ても、隠れたままにして境目を 1 回決め直す" {
        $script:filterText = "議事録"
        [void](addHit "見積.xlsx" "4月" "`t見積 山田" 1)
        [void](addHit "議事録.docx" "ページ001" "見積 佐藤" 1)
        flushResults
        [void](addHit "見積.xlsx" "4月" "`t議事録 佐藤" 2)
        $script:paneHeightCalls = @()

        finishResults

        (getItemNames) -join "," | Should -Be "#議事録.docx"
        $script:paneHeightCalls | Should -Be @(1)
    }

    It "どちらでもなければ表はそのまま" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        flushResults
        $items = $script:resultItems

        finishResults

        [object]::ReferenceEquals($script:resultItems, $items) | Should -Be $true
    }
}

Describe "getCurrentHitRow" -Tag Unit {
    BeforeEach { resetResults }

    It "見出しを選んでいるときは、そのファイルの先頭の行" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 4
        flushResults
        $ui.ResultGrid.SelectedItem = $group

        (getCurrentHitRow).LineNumber | Should -Be 4
    }

    It "見出しの行がすべて絞り込みで隠れていれば無し" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 4
        $script:filterText = "佐藤"
        $ui.ResultGrid.SelectedItem = $group

        $null -eq (getCurrentHitRow) | Should -Be $true
    }

    It "行を選んでいるときはその行、何も選んでいなければ無し" {
        $group = addHit "見積.xlsx" "4月" "`t見積" 4
        ensureRows $group
        $ui.ResultGrid.SelectedItem = $group.Rows[0]
        (getCurrentHitRow) | Should -Be $group.Rows[0]

        $ui.ResultGrid.SelectedItem = $null
        $null -eq (getCurrentHitRow) | Should -Be $true
    }
}

Describe "getViewRows・getSelectedRows" -Tag Unit {
    BeforeEach { resetResults }

    It "表示中の行は、閉じているファイルの行も含めて表の順に返す" {
        [void](addHit "見積.xlsx" "4月" "`t見積" 1)
        [void](addHit "議事録.docx" "ページ001" "見積" 2)
        flushResults

        $rows = getViewRows

        $rows.Count | Should -Be 2
        $rows[0].Book | Should -Be "見積.xlsx"
        $rows[1].Book | Should -Be "議事録.docx"
    }

    It "結果が無い・何も選んでいないときは空の配列を返す" {
        $view = getViewRows
        $selected = getSelectedRows

        , $view | Should -BeOfType [object[]]
        $view.Count | Should -Be 0
        , $selected | Should -BeOfType [object[]]
        $selected.Count | Should -Be 0
        (getShownHitCount) | Should -Be 0
    }

    It "見出しを選ぶとそのファイルの行をすべて、行を選ぶとその行を、表の順に返す" {
        $mitsumori = addHit "見積.xlsx" "4月" "`t見積" 1
        [void](addHit "見積.xlsx" "4月" "`t見積" 2)
        $gijiroku = addHit "議事録.docx" "ページ001" "見積" 3
        [void](addHit "議事録.docx" "ページ001" "見積" 4)
        ensureRows $gijiroku
        $ui.ResultGrid.SelectedItems = @($gijiroku.Rows[1], $mitsumori)

        $rows = getSelectedRows

        (@($rows | ForEach-Object { $_.LineNumber }) -join ",") | Should -Be "1,2,4"
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

        $row.Prepared | Should -Be $true
        $row.MatchCell | Should -Be "B1"
        $row.PlaceDisplay | Should -Be "[シート]4月!B1"
    }
}

Describe "prepareHitRow" -Tag Unit {
    BeforeEach { resetResults }

    It "<name>" -TestCases @(
        @{ name = "Excel のセルはセル番地を足す"; book = "見積.xlsx"; location = "4月"; line = "`t見積書"; expected = "[シート]4月!B1" }
        @{ name = "Excel の図形は左上のセル番地を足す"; book = "見積.xlsx"; location = "4月[図形]"; line = "D5`t見積の注記"; expected = "[シート]4月!D5" }
        @{ name = "Excel のヘッダー・フッターはセル番地が無く、シート名だけ"; book = "見積.xlsx"; location = "4月[ヘッダー・フッター]"; line = "社外秘"; expected = "[シート]4月" }
        @{ name = "Word は場所ごとの表記のまま"; book = "議事録.docx"; location = "ページ003"; line = "見積の件"; expected = "3 ページ（目安）" }
    ) {
        param ($name, $book, $location, $line, $expected)
        $group = addHit $book $location $line 1
        ensureRows $group
        $row = $group.Rows[0]

        prepareHitRow $row

        $row.PlaceDisplay | Should -Be $expected
    }

    It "作り済みなら作り直さない" {
        $group = addHit "見積.xlsx" "4月" "`t見積書" 1
        ensureRows $group
        $row = $group.Rows[0]
        prepareHitRow $row
        $row.PlaceDisplay = "変えた"

        prepareHitRow $row

        $row.PlaceDisplay | Should -Be "変えた"
    }
}

Describe "結果の表の定義（search\result_list.xaml）" -Tag Unit {
    BeforeAll {
        $xaml = [xml](Get-Content -LiteralPath "${scriptsDir}\tebunko\xaml\search\result_list.xaml" -Raw -Encoding UTF8)
        $ns = New-Object System.Xml.XmlNamespaceManager($xaml.NameTable)
        $ns.AddNamespace("p", "http://schemas.microsoft.com/winfx/2006/xaml/presentation")
        $resultGrid = $xaml.SelectSingleNode("//p:DataGrid[@*[local-name()='Name']='ResultGrid']", $ns)
    }

    It "列は「場所」「種別」「該当行」の 3 つだけにする" {
        @($resultGrid.SelectSingleNode("p:DataGrid.Columns", $ns).ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] } | ForEach-Object { $_.GetAttribute("Header") }) -join "," | Should -Be "場所,種別,該当行"
    }

    It "「場所」の列は、ツールチップで全文を出す" {
        $column = $resultGrid.SelectSingleNode("p:DataGrid.Columns/p:DataGridTextColumn[@Header='場所']", $ns)
        $column.GetAttribute("SortMemberPath") | Should -Be "Location"
        $tip = $column.SelectSingleNode("p:DataGridTextColumn.ElementStyle/p:Style/p:Setter[@Property='ToolTip']", $ns)
        $tip.GetAttribute("Value") | Should -Be "{Binding PlaceDisplay}"
    }

    It "行の名前を付ける（見出しはファイル名、ヒットの行は場所）" {
        $names = @($resultGrid.SelectNodes("p:DataGrid.RowStyle//p:Setter[@Property='AutomationProperties.Name']", $ns) | ForEach-Object { $_.GetAttribute("Value") })
        $names | Should -Be @("{Binding BookDisplay}", "{Binding PlaceDisplay}", "{Binding AutomationName}")
        $resultGrid.SelectNodes("p:DataGrid.RowStyle//p:Setter[@Property='AutomationProperties.HelpText']", $ns).Count | Should -Be 1
    }

    It "使わなくなった列（IndexColumn）を参照するコードが無い" {
        @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -Include *.ps1, *.xaml | Select-String -SimpleMatch "IndexColumn").Count | Should -Be 0
    }
}

Describe "1 ファイルのヒットが多いとき" -Tag Unit {
    BeforeEach {
        resetResults
        # 1 ファイルに 5,000 件（上限の 1,000 件を超える）。LineNumber は見つかった順に 5000 → 1
        $script:big = $null
        foreach ($n in 5000..1) {
            $script:big = addHit "多.txt" "本文" "見積 $n" $n
        }
        flushResults
    }

    It "開くと行は上限（1,000）までで、表の項目は見出し・行・続きの行" {
        toggleFileGroup $script:big

        $script:big.Rows.Count | Should -Be 1000
        $script:resultItems.Count | Should -Be 1002
        $script:resultItems[1001] -is [MoreRowsItem] | Should -Be $true
        $script:resultItems[1001].Text | Should -Be "続きの 1,000 件を表示（残り 4,000 件）"
        $script:big.ShownCount | Should -Be 5000
    }

    It "続きを出すと 1,000 件ずつ増え、最後まで出すと続きの行が消える" {
        toggleFileGroup $script:big

        showMoreGroupRows $script:big
        $script:big.Rows.Count | Should -Be 2000
        $script:resultItems.Count | Should -Be 2002
        $script:resultItems[2001].Text | Should -Be "続きの 1,000 件を表示（残り 3,000 件）"
        $ui.ResultGrid.Scrolled[$ui.ResultGrid.Scrolled.Count - 1] | Should -Be $script:big.Rows[1000]

        foreach ($i in 1..3) { showMoreGroupRows $script:big }
        $script:big.Rows.Count | Should -Be 5000
        $script:resultItems.Count | Should -Be 5001
        $null -eq $script:big.MoreRow | Should -Be $true
        ($script:resultItems[5000] -is [HitRow]) | Should -Be $true
    }

    It "閉じると続きの行も無くなる" {
        toggleFileGroup $script:big
        toggleFileGroup $script:big

        $script:resultItems.Count | Should -Be 1
        $script:big.MoreInView | Should -Be $false
    }

    It "［すべて開く］でも、作る行は上限まで" {
        setAllFileGroupsExpanded $true

        $script:big.Rows.Count | Should -Be 1000
        $script:resultItems.Count | Should -Be 1002
    }

    It "絞り込みは行を作らず、ファイルごとに決める" {
        $script:filterText = "存在しない"
        applyResultFilter "存在しない"
        $script:big.Rows.Count | Should -Be 0
        $script:big.ShownCount | Should -Be 0
        $script:resultItems.Count | Should -Be 0

        $script:filterText = "多"
        applyResultFilter "多"
        $script:big.Rows.Count | Should -Be 0
        $script:big.ShownCount | Should -Be 5000
        $script:resultItems.Count | Should -Be 1
    }

    It "並べ替えると、上限を超えたファイルは見つかった順のまま、上限以下のファイルは並べ替える" {
        [void](addHit "少.txt" "本文" "見積" 20)
        [void](addHit "少.txt" "本文" "見積" 10)
        flushResults
        $script:fileGroups.Values | Where-Object { $_.Book -eq "少.txt" } | ForEach-Object { $small = $_ }
        $script:big.SetExpanded($true)
        $column = newSortColumn "LineNumber"
        $ui.ResultGrid.Columns.Add($column)

        sortResults $column

        $script:big.Rows.Count | Should -Be 1000
        @($script:big.Rows[0..2] | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "5000,4999,4998"
        @($small.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "10,20"
        # ファイルの順は先頭の行の値（少: 10、多: 見つかった順の先頭の 5000）
        @($script:groupList | ForEach-Object { $_.Book }) -join "," | Should -Be "少.txt,多.txt"
        $script:big.MoreRow.Note | Should -Be "・件数が多いため並べ替えていません"
    }

    It "続きの行を選んで Enter / クリックすると続きが出る。見出しは開閉する。行は対象外" {
        toggleFileGroup $script:big
        activateResultItem $script:big.MoreRow | Should -Be $true
        $script:big.Rows.Count | Should -Be 2000

        activateResultItem $script:big | Should -Be $true
        $script:big.IsExpanded | Should -Be $false
        activateResultItem $script:big.Rows[0] | Should -Be $false
    }

    It "見出しを選ぶと先頭の 1 件だけ作り、続きの行を選んでも行は無い" {
        $ui.ResultGrid.SelectedItem = $script:big
        (getCurrentHitRow).LineNumber | Should -Be 5000
        $script:big.Rows.Count | Should -Be 1

        toggleFileGroup $script:big
        $ui.ResultGrid.SelectedItem = $script:big.MoreRow
        $null -eq (getCurrentHitRow) | Should -Be $true
    }

    It "見出しを選んでコピーするときは、全件の行になる" {
        $ui.ResultGrid.SelectedItems = @($script:big)
        (getSelectedRows).Count | Should -Be 5000
    }

    It "全件の行を作った（コピー・並べ替えての出力）あとも、表に出す行は上限と続きの行のまま" {
        $ui.ResultGrid.SelectedItems = @($script:big)
        (getSelectedRows).Count | Should -Be 5000
        $script:big.Rows.Count | Should -Be 5000

        toggleFileGroup $script:big
        $script:resultItems.Count | Should -Be 1002
        $script:resultItems[1001].Text | Should -Be "続きの 1,000 件を表示（残り 4,000 件）"

        showMoreGroupRows $script:big
        $script:resultItems.Count | Should -Be 2002
        $script:resultItems[2001].Text | Should -Be "続きの 1,000 件を表示（残り 3,000 件）"

        $column = newSortColumn "LineNumber"
        $ui.ResultGrid.Columns.Add($column)
        sortResults $column
        $script:resultItems.Count | Should -Be 2002
        @($script:big.Rows[0..2] | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "5000,4999,4998"
    }

    It "開いたあとにコピーしても、続きの行の残りの件数は変わらない" {
        toggleFileGroup $script:big
        $ui.ResultGrid.SelectedItems = @($script:big)
        (getSelectedRows).Count | Should -Be 5000

        updateMoreRow $script:big
        $script:big.MoreRow.Text | Should -Be "続きの 1,000 件を表示（残り 4,000 件）"
        $script:resultItems.Count | Should -Be 1002
    }

    It "検索中に、開いているファイルにヒットが届くと、続きの行の残りの件数を更新する" {
        toggleFileGroup $script:big
        foreach ($n in 1..10) {
            [void](addHit "多.txt" "本文" "見積 追加$n" (6000 + $n))
        }

        flushResults

        $script:big.Rows.Count | Should -Be 1000
        $script:resultItems.Count | Should -Be 1002
        $script:resultItems[1001].Text | Should -Be "続きの 1,000 件を表示（残り 4,010 件）"
        $script:big.ShownCount | Should -Be 5010
    }

    It "最後の続きが少ないとき（200 件以下）は、表に 1 件ずつ足し、続きの行が消える" {
        foreach ($n in 1..1150) {
            $script:mid = addHit "中.txt" "本文" "見積 $n" $n
        }
        flushResults
        toggleFileGroup $script:mid
        $script:mid.Rows.Count | Should -Be 1000
        $before = $script:resultItems.Count

        showMoreGroupRows $script:mid

        $script:mid.Rows.Count | Should -Be 1150
        $script:resultItems.Count | Should -Be ($before + 150 - 1)
        $null -eq $script:mid.MoreRow | Should -Be $true
        $script:mid.MoreInView | Should -Be $false
        ($script:resultItems[$script:resultItems.Count - 1] -is [HitRow]) | Should -Be $true
    }
}

Describe "getExportHits" -Tag Unit {
    BeforeEach {
        resetResults
        [void](addHit "b.docx" "ページ001" "見積" 1)
        [void](addHit "b.docx" "ページ001" "見積" 2)
        [void](addHit "a.docx" "ページ001" "見積" 3)
        [void](addHit "c.xlsx" "4月" "`t見積" 4)
        flushResults
        $column = newSortColumn "Book"
        $ui.ResultGrid.Columns.Add($column)
        sortResults $column
    }

    It "<Name>" -ForEach @(
        @{ Name = "上限を超えていない・並べ替えている: 表の順"; Truncated = $false; Filter = ""; Sorted = $true; Expected = "3,1,2,4" }
        @{ Name = "上限を超えていない・並べ替えていない: 検索の順（行の部品を作らない）"; Truncated = $false; Filter = ""; Sorted = $false; Expected = "1,2,3,4" }
        @{ Name = "上限を超えていない・絞り込みあり: 検索の順で、合うファイルだけ"; Truncated = $false; Filter = "docx"; Sorted = $false; Expected = "1,2,3" }
        @{ Name = "上限を超えている: 並べ替えは当てず、検索の順（ファイルの見つかった順 → ファイルの中の順）"; Truncated = $true; Filter = ""; Expected = "1,2,3,4" }
        @{ Name = "上限を超えている: ファイル名の絞り込み（大文字・小文字を区別しない）"; Truncated = $true; Filter = "DOCX"; Expected = "1,2,3" }
        @{ Name = "上限を超えている: 合うファイルが無い"; Truncated = $true; Filter = "存在しない"; Expected = "" }
    ) {
        $hits = getExportHits $Truncated $Filter ([bool]$Sorted)

        (@($hits | ForEach-Object { $_.LineNumber }) -join ",") | Should -Be $Expected
    }

    It "上限を超えているときは、行の部品（HitRow）を作らず、検索の素の値を返す" {
        $hits = getExportHits $true ""

        $hits[0].GetType().Name | Should -Not -Be "HitRow"
        $hits[0].Book | Should -Be "b.docx"
    }
}
