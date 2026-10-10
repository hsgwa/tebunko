# 結果の一覧の判断（tebunko\ui\search\result_list_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_list_view.ps1"

    # 結果の表の項目のテストに使う、HitRow・FileGroup と同じ項目を持つもの
    function newTestRow {
        param ([int]$order, [string]$line, [int]$lineNumber = 1, [string]$location = "")

        $row = [pscustomobject]@{ Order = $order; Line = $line; LineNumber = $lineNumber; Location = $location }
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
}

Describe "getSearchProgressText / getSearchSummaryText" -Tag Unit {
    It "検索中は該当件数を出す" {
        getSearchProgressText 1234 | Should -Be "検索中…　該当 1,234 件"
    }

    It "<name>" -TestCases @(
        @{ name = "終わったら該当件数・ファイル数・秒数を出す"; seconds = @(1.25); expected = "^一致 1,234 件（5 ファイル）・1\.[23] 秒$" }
        @{ name = "秒を渡さなければ時間は付けない"; seconds = @(); expected = "^一致 1,234 件（5 ファイル）$" }
    ) {
        param ($name, $seconds, $expected)
        getSearchSummaryText 1234 5 @seconds | Should -Match $expected
    }
}

Describe "getSearchStatusText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "件数を 3 桁区切りで出す"; word = "見積"; hits = 1234; option = ""; expected = "検索しました：見積 1,234 件" }
        @{ name = "0 件もステータスバーの文にする"; word = "見積"; hits = 0; option = ""; expected = "検索しました：見積 0 件" }
        @{ name = "条件があれば後ろに付ける"; word = "見積"; hits = 0; option = "大文字・小文字を区別"; expected = "検索しました：見積 0 件　条件：大文字・小文字を区別" }
    ) {
        param ($name, $word, $hits, $option, $expected)
        getSearchStatusText $word $hits $option | Should -Be $expected
    }
}

Describe "getFilteredSummaryText" -Tag Unit {
    It "全部の件数のうち、見せている件数を出す" {
        getFilteredSummaryText 1234 56 | Should -Be "1,234 件中 56 件を表示"
    }
}

Describe "getAppKind" -Tag Unit {
    It "拡張子からアプリの種類を返す（大文字・小文字は問わない）" {
        getAppKind "見積.xlsx" | Should -Be "Excel"
        getAppKind "古い見積.XLS" | Should -Be "Excel"
        getAppKind "マクロ.xlsm" | Should -Be "Excel"
        getAppKind "報告書.docx" | Should -Be "Word"
        getAppKind "報告書.doc" | Should -Be "Word"
        getAppKind "提案.pptx" | Should -Be "PowerPoint"
        getAppKind "議事メモ.txt" | Should -Be "テキスト"
    }

    It "Office・テキストのファイルでなければ空" {
        getAppKind "資料.pdf" | Should -Be ""
        getAppKind "" | Should -Be ""
    }
}

Describe "describeFileLocations" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "場所が無ければ空"; locations = @(); expected = "" }
        @{ name = "1 か所ならその場所"; locations = @("[シート]4月"); expected = "[シート]4月" }
        @{ name = "2 か所以上なら先頭と、ほかの数"; locations = @("[シート]4月", "[シート]5月", "[シート]6月"); expected = "[シート]4月 ほか 2 か所" }
    ) {
        param ($name, $locations, $expected)
        describeFileLocations $locations | Should -Be $expected
    }
}

Describe "selectShownRows" -Tag Unit {
    It "絞り込みが空ならすべて、あれば合う行だけを元の順で返す" {
        $rows = @((newTestRow 1 "見積 A"), (newTestRow 2 "請求 B"), (newTestRow 3 "見積 C"))
        (selectShownRows $rows "").Count | Should -Be 3
        $shown = selectShownRows $rows "見積"
        $shown.Count | Should -Be 2
        $shown[0].Order | Should -Be 1
        $shown[1].Order | Should -Be 3
    }

    It "合う行が無ければ空の一覧" {
        (selectShownRows @((newTestRow 1 "見積")) "請求").Count | Should -Be 0
    }
}

Describe "getResultItems" -Tag Unit {
    It "閉じているファイルは見出しだけ、開いているファイルは見出しと行" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"), (newTestRow 2 "a2"))
        $b = newTestGroup 2 @((newTestRow 3 "b1")) $true
        $items = getResultItems @($a, $b)
        $items.Count | Should -Be 3
        [object]::ReferenceEquals($items[0], $a) | Should -Be $true
        [object]::ReferenceEquals($items[1], $b) | Should -Be $true
        $items[2].Line | Should -Be "b1"
    }

    It "絞り込みで行が残らないファイルは見出しも出さない" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"))
        $a.ShownRows.Clear()
        $a.ShownCount = 0
        $b = newTestGroup 2 @((newTestRow 2 "b1"))
        $items = getResultItems @($a, $b)
        $items.Count | Should -Be 1
        [object]::ReferenceEquals($items[0], $b) | Should -Be $true
    }
}

Describe "getShownHitRows" -Tag Unit {
    It "閉じているファイルの行も含め、表の順に返す" {
        $a = newTestGroup 1 @((newTestRow 1 "a1"), (newTestRow 2 "a2"))
        $b = newTestGroup 2 @((newTestRow 3 "b1")) $true
        $rows = getShownHitRows @($a, $b)
        @($rows | ForEach-Object { $_.Line }) -join "," | Should -Be "a1,a2,b1"
    }
}

Describe "sortFileGroups" -Tag Unit {
    It "ファイルの中の行を並べ替え、ファイルは先頭の行の順にする" {
        $a = newTestGroup 1 @((newTestRow 1 "a" 5), (newTestRow 2 "a" 9))
        $b = newTestGroup 2 @((newTestRow 3 "b" 7), (newTestRow 4 "b" 1))
        $sorted = sortFileGroups @($a, $b) "LineNumber" $false
        [object]::ReferenceEquals($sorted[0], $b) | Should -Be $true
        @($b.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "1,7"
        @($a.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "5,9"
    }

    It "逆順にもできる" {
        $a = newTestGroup 1 @((newTestRow 1 "a" 5), (newTestRow 2 "a" 9))
        $b = newTestGroup 2 @((newTestRow 3 "b" 7), (newTestRow 4 "b" 1))
        $sorted = sortFileGroups @($a, $b) "LineNumber" $true
        [object]::ReferenceEquals($sorted[0], $a) | Should -Be $true
        @($a.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "9,5"
    }

    It "「場所」（Location）では、同じ場所の中を行番号の順にする（見つかった順と食い違っていても）" {
        $a = newTestGroup 1 @((newTestRow 0 "a" 10 "売上"), (newTestRow 1 "a" 9 "売上"), (newTestRow 2 "a" 1 "仕入"))
        $sorted = sortFileGroups @($a) "Location" $false
        @($a.Rows | ForEach-Object { "$($_.Location)$($_.LineNumber)" }) -join "," | Should -Be "仕入1,売上9,売上10"
        $sorted = sortFileGroups @($a) "Location" $true
        @($a.Rows | ForEach-Object { "$($_.Location)$($_.LineNumber)" }) -join "," | Should -Be "売上10,売上9,仕入1"
    }

    It "同じ値のときは見つかった順" {
        $a = newTestGroup 1 @((newTestRow 2 "x" 1), (newTestRow 1 "x" 1))
        $b = newTestGroup 2 @((newTestRow 3 "x" 1))
        $sorted = sortFileGroups @($b, $a) "LineNumber" $false
        [object]::ReferenceEquals($sorted[0], $a) | Should -Be $true
        @($a.Rows | ForEach-Object { $_.Order }) -join "," | Should -Be "1,2"
    }
}

Describe "getResultMenuItems" -Tag Unit {
    BeforeAll {
        function menuText {
            param ($items)
            (@($items | ForEach-Object { $(if ($_.Bold) { "*" } else { "" }) + $(if ($_.Id -eq "separator") { "-" } else { $_.Header }) }) -join "/")
        }
    }

    It "行のメニューは、先頭に既定の開き方を太字で出し、下にほかの開き方を［開く ▾］の順で並べる（既定: <mode>）" -TestCases @(
        @{ mode = "normal";   expected = "*開く/新規で開く/読み取り専用で開く/フォルダを開く/-/選んだ行をコピー/ファイルのパスをコピー" }
        @{ mode = "readOnly"; expected = "*読み取り専用で開く/開く/新規で開く/フォルダを開く/-/選んだ行をコピー/ファイルのパスをコピー" }
        @{ mode = "new";      expected = "*新規で開く/開く/読み取り専用で開く/フォルダを開く/-/選んだ行をコピー/ファイルのパスをコピー" }
        @{ mode = "";         expected = "*開く/新規で開く/読み取り専用で開く/フォルダを開く/-/選んだ行をコピー/ファイルのパスをコピー" }
        @{ mode = "unknown";  expected = "*開く/新規で開く/読み取り専用で開く/フォルダを開く/-/選んだ行をコピー/ファイルのパスをコピー" }
    ) {
        param ($mode, $expected)
        menuText (getResultMenuItems "row" $mode) | Should -Be $expected
    }

    It "見出しのメニューは、読み取り専用で開くを太字で先頭に出し、開閉の項目は開き具合で文言を変える（expanded: <expanded>）" -TestCases @(
        @{ expanded = $true;  expected = "*読み取り専用で開く/フォルダを開く/-/ファイルのパスをコピー/この結果を折りたたむ" }
        @{ expanded = $false; expected = "*読み取り専用で開く/フォルダを開く/-/ファイルのパスをコピー/この結果を開く" }
    ) {
        param ($expanded, $expected)
        # 既定の開き方が別でも、見出しのメニューは変わらない
        menuText (getResultMenuItems "group" "new" $expanded) | Should -Be $expected
    }

    It "太字は 1 つだけで、Id は項目ごとに決まり、「…」・キー操作・（既定）は付けない" {
        foreach ($target in "row", "group") {
            $items = @(getResultMenuItems $target "readOnly" $true)
            @($items | Where-Object { $_.Bold }).Count | Should -Be 1
            $ids = @($items | Where-Object { $_.Id -ne "separator" } | ForEach-Object { $_.Id })
            $ids.Count | Should -Be @($ids | Select-Object -Unique).Count
            foreach ($item in $items) {
                $item.Header | Should -Not -Match "…|Ctrl|Enter|（既定）"
            }
        }
    }
}

Describe "testResultMenuKeepSelection" -Tag Unit {
    # 右クリックした行が選ばれていて、見出しでなく、選びに見出しを含まないときだけ、選びを変えない
    It "<name>" -TestCases @(
        @{ name = "選ばれていない行は、その行だけを選ぶ"; rowSelected = $false; item = "row"; selected = @("row"); expected = $false }
        @{ name = "選ばれている行を複数選んでいるなら、選びを変えない"; rowSelected = $true; item = "row"; selected = @("row", "row2"); expected = $true }
        @{ name = "選ばれている行 1 つなら、選びを変えない"; rowSelected = $true; item = "row"; selected = @("row"); expected = $true }
        @{ name = "見出しは、見出しだけを選ぶ"; rowSelected = $true; item = "group"; selected = @("group"); expected = $false }
        @{ name = "選びに見出しを含むなら、右クリックした行だけを選ぶ"; rowSelected = $true; item = "row"; selected = @("row", "group"); expected = $false }
        @{ name = "選びが空でも、選ばれている行なら変えない"; rowSelected = $true; item = "row"; selected = @(); expected = $true }
    ) {
        $items = @{ row = "行"; row2 = "行 2"; group = [FileGroup]::new() }
        $selectedItems = @($selected | ForEach-Object { $items[$_] })
        testResultMenuKeepSelection $rowSelected $items[$item] $selectedItems | Should -Be $expected
    }
}

Describe "getResultMenuContext" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "行は row・開いていない扱い"; kind = "row"; expanded = $false; target = "row"; expect = $false }
        @{ name = "開いている見出しは group・開いている"; kind = "group"; expanded = $true; target = "group"; expect = $true }
        @{ name = "閉じている見出しは group・閉じている"; kind = "group"; expanded = $false; target = "group"; expect = $false }
    ) {
        $item = $(if ($kind -eq "group") { $g = [FileGroup]::new(); $g.IsExpanded = $expanded; $g } else { "行" })
        $context = getResultMenuContext $item
        $context.Target | Should -Be $target
        $context.Expanded | Should -Be $expect
    }
}
