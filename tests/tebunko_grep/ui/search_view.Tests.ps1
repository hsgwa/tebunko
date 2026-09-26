# ［2 検索］の判断（tebunko_grep\ui\search_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\ui\search_view.ps1"

Describe "describeSearchOption" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "既定のままなら空"; option = @{ CaseSensitive = $false; FileFilter = "" }; expected = "" }
        @{ name = "大文字と小文字の区別を出す"; option = @{ CaseSensitive = $true; FileFilter = "" }; expected = "大文字と小文字を区別" }
        @{ name = "対象ファイルを出す"; option = @{ CaseSensitive = $false; FileFilter = "*.xlsx" }; expected = "対象ファイル：*.xlsx" }
        @{ name = "両方あれば中黒でつなぐ"; option = @{ CaseSensitive = $true; FileFilter = "*.xlsx" }; expected = "大文字と小文字を区別・対象ファイル：*.xlsx" }
        @{ name = "図形・コメントを含めるなら出さない"; option = @{ CaseSensitive = $false; FileFilter = ""; IncludeShapes = $true; IncludeComments = $true }; expected = "" }
        @{ name = "図形・コメントを外したら出す"; option = @{ CaseSensitive = $false; FileFilter = ""; IncludeShapes = $false; IncludeComments = $false }; expected = "図形を除く・コメントを除く" }
        @{ name = "コメントだけ外したらコメントだけ出す"; option = @{ CaseSensitive = $false; FileFilter = ""; IncludeComments = $false }; expected = "コメントを除く" }
    ) {
        param ($name, $option, $expected)
        describeSearchOption $option | Should Be $expected
    }
}

Describe "getFastSearchView" -Tag Unit {
    It "Windows Search が使え（またはまだ確かめていない）、正規表現がオフで、2 文字以上の部分があれば使用可" {
        (getFastSearchView $true $false "見積").Text | Should Be "高速検索：使用可"
        (getFastSearchView $null $false "見積").Usable | Should Be $true
    }

    It "正規表現をオンにした・1 文字・Windows Search が使えないときは使用不可" {
        (getFastSearchView $true $true "見積").Text | Should Be "高速検索：使用不可"
        (getFastSearchView $true $false "見").Usable | Should Be $false
        (getFastSearchView $false $false "見積").Usable | Should Be $false
    }
}

Describe "getSearchProgressText / getSearchSummaryText" -Tag Unit {
    It "検索中は該当件数を出す" {
        getSearchProgressText 1234 | Should Be "検索中…　該当 1,234 件"
    }

    It "終わったら該当件数・ファイル数・秒数を出す" {
        getSearchSummaryText 1234 5 1.25 | Should Match "^該当 1,234 件（5 ファイル） ・ 1\.[23] 秒$"
    }
}

Describe "getWordNotice" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "正規表現でなければ出さない"; word = "("; useRegex = $false; expected = "" }
        @{ name = "正規表現として正しければ出さない"; word = "見積.*確定"; useRegex = $true; expected = "" }
        @{ name = "空のワードでは出さない"; word = ""; useRegex = $true; expected = "" }
        @{ name = "正規表現として不正なら、文字どおり検索すると伝える"; word = "("; useRegex = $true; expected = "正規表現として不正なため、文字どおり検索します。" }
    ) {
        param ($name, $word, $useRegex, $expected)
        getWordNotice $word $useRegex | Should Be $expected
    }
}

Describe "newSearchButtonState" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "検索中は［中止］にする"; searching = $true; stopping = $false; word = "見積"; hasIndex = $true; targetCount = 1; content = "中止"; enabled = $true }
        @{ name = "中止を頼んだ後は押せない"; searching = $true; stopping = $true; word = "見積"; hasIndex = $true; targetCount = 1; content = "中止"; enabled = $false }
        @{ name = "ワード・インデックス・検索対象がそろえば押せる"; searching = $false; stopping = $false; word = "見積"; hasIndex = $true; targetCount = 2; content = "検索"; enabled = $true }
        @{ name = "ワードが空なら押せない"; searching = $false; stopping = $false; word = ""; hasIndex = $true; targetCount = 2; content = "検索"; enabled = $false }
        @{ name = "インデックスが無ければ押せない"; searching = $false; stopping = $false; word = "見積"; hasIndex = $false; targetCount = 2; content = "検索"; enabled = $false }
        @{ name = "検索対象が選ばれていなければ押せない"; searching = $false; stopping = $false; word = "見積"; hasIndex = $true; targetCount = 0; content = "検索"; enabled = $false }
    ) {
        param ($name, $searching, $stopping, $word, $hasIndex, $targetCount, $content, $enabled)
        $state = newSearchButtonState $searching $stopping $word $hasIndex $targetCount
        $state.Content | Should Be $content
        $state.Enabled | Should Be $enabled
    }
}

Describe "getAppKind" -Tag Unit {
    It "拡張子からアプリの種類を返す（大文字・小文字は問わない）" {
        getAppKind "見積.xlsx" | Should Be "Excel"
        getAppKind "古い見積.XLS" | Should Be "Excel"
        getAppKind "マクロ.xlsm" | Should Be "Excel"
        getAppKind "報告書.docx" | Should Be "Word"
        getAppKind "報告書.doc" | Should Be "Word"
        getAppKind "提案.pptx" | Should Be "PowerPoint"
    }

    It "Office のファイルでなければ空" {
        getAppKind "メモ.txt" | Should Be ""
        getAppKind "" | Should Be ""
    }
}

Describe "describeFileLocations" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "場所が無ければ空"; locations = @(); expected = "" }
        @{ name = "1 か所ならその場所"; locations = @("[シート] 4月"); expected = "[シート] 4月" }
        @{ name = "2 か所以上なら先頭と、ほかの数"; locations = @("[シート] 4月", "[シート] 5月", "[シート] 6月"); expected = "[シート] 4月 ほか 2 か所" }
    ) {
        param ($name, $locations, $expected)
        describeFileLocations $locations | Should Be $expected
    }
}

# 結果の表の項目のテストに使う、HitRow・FileGroup と同じ項目を持つもの
function newTestRow {
    param ([int]$order, [string]$line, [int]$lineNumber = 1)

    $row = [pscustomobject]@{ Order = $order; Line = $line; LineNumber = $lineNumber }
    $row | Add-Member -MemberType ScriptMethod -Name Contains -Value { param ($text) $this.Line.Contains($text) }
    return $row
}

function newTestGroup {
    param ([int]$order, [object[]]$rows, [bool]$expanded = $false)

    $group = [pscustomobject]@{
        Order = $order; IsExpanded = $expanded; ShownCount = $rows.Count
        Rows = New-Object 'System.Collections.Generic.List[object]'
        ShownRows = New-Object 'System.Collections.Generic.List[object]'
    }
    $group.Rows.AddRange($rows)
    $group.ShownRows.AddRange($rows)
    return $group
}

Describe "selectShownRows" -Tag Unit {
    It "絞り込みが空ならすべて、あれば合う行だけを元の順で返す" {
        $rows = @((newTestRow 1 "見積 A"), (newTestRow 2 "請求 B"), (newTestRow 3 "見積 C"))
        (selectShownRows $rows "").Count | Should Be 3
        $shown = selectShownRows $rows "見積"
        $shown.Count | Should Be 2
        $shown[0].Order | Should Be 1
        $shown[1].Order | Should Be 3
    }

    It "合う行が無ければ空の一覧" {
        (selectShownRows @((newTestRow 1 "見積")) "請求").Count | Should Be 0
    }
}

Describe "getResultItems" -Tag Unit {
    It "閉じているファイルは見出しだけ、開いているファイルは見出しと行" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"), (newTestRow 2 "a2"))
        $b = newTestGroup 2 @((newTestRow 3 "b1")) $true
        $items = getResultItems @($a, $b)
        $items.Count | Should Be 3
        [object]::ReferenceEquals($items[0], $a) | Should Be $true
        [object]::ReferenceEquals($items[1], $b) | Should Be $true
        $items[2].Line | Should Be "b1"
    }

    It "絞り込みで行が残らないファイルは見出しも出さない" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"))
        $a.ShownRows.Clear()
        $a.ShownCount = 0
        $b = newTestGroup 2 @((newTestRow 2 "b1"))
        $items = getResultItems @($a, $b)
        $items.Count | Should Be 1
        [object]::ReferenceEquals($items[0], $b) | Should Be $true
    }
}

Describe "getShownHitRows" -Tag Unit {
    It "閉じているファイルの行も含め、表の順に返す" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"), (newTestRow 2 "a2"))
        $b = newTestGroup 2 @((newTestRow 3 "b1")) $true
        $rows = getShownHitRows @($a, $b)
        @($rows | ForEach-Object { $_.Line }) -join "," | Should Be "a1,a2,b1"
    }
}

Describe "sortFileGroups" -Tag Unit {
    It "ファイルの中の行を並べ替え、ファイルは先頭の行の順にする" {
        $a = newTestGroup 1 @((newTestRow 1 "a" 5), (newTestRow 2 "a" 9))
        $b = newTestGroup 2 @((newTestRow 3 "b" 7), (newTestRow 4 "b" 1))
        $sorted = sortFileGroups @($a, $b) "LineNumber" $false
        [object]::ReferenceEquals($sorted[0], $b) | Should Be $true
        @($b.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should Be "1,7"
        @($a.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should Be "5,9"
    }

    It "逆順にもできる" {
        $a = newTestGroup 1 @((newTestRow 1 "a" 5), (newTestRow 2 "a" 9))
        $b = newTestGroup 2 @((newTestRow 3 "b" 7), (newTestRow 4 "b" 1))
        $sorted = sortFileGroups @($a, $b) "LineNumber" $true
        [object]::ReferenceEquals($sorted[0], $a) | Should Be $true
        @($a.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should Be "9,5"
    }

    It "同じ値のときは見つかった順" {
        $a = newTestGroup 1 @((newTestRow 2 "x" 1), (newTestRow 1 "x" 1))
        $b = newTestGroup 2 @((newTestRow 3 "x" 1))
        $sorted = sortFileGroups @($b, $a) "LineNumber" $false
        [object]::ReferenceEquals($sorted[0], $a) | Should Be $true
        @($a.Rows | ForEach-Object { $_.Order }) -join "," | Should Be "1,2"
    }
}