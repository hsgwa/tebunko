# ［2 検索］の判断（tebunko_grep\ui\search_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\ui\search_view.ps1"

Describe "describeSearchOption" -Tag Unit {
    It "既定のままなら空" {
        describeSearchOption @{ CaseSensitive = $false; FileFilter = "" } | Should Be ""
    }

    It "大文字と小文字の区別を出す" {
        describeSearchOption @{ CaseSensitive = $true; FileFilter = "" } | Should Be "大文字と小文字を区別"
    }

    It "対象ファイルを出す" {
        describeSearchOption @{ CaseSensitive = $false; FileFilter = "*.xlsx" } | Should Be "対象ファイル：*.xlsx"
    }

    It "両方あれば中黒でつなぐ" {
        describeSearchOption @{ CaseSensitive = $true; FileFilter = "*.xlsx" } | Should Be "大文字と小文字を区別・対象ファイル：*.xlsx"
    }
}

Describe "getWordNotice" -Tag Unit {
    It "正規表現でなければ出さない" {
        getWordNotice "(" $false | Should Be ""
    }

    It "正規表現として正しければ出さない" {
        getWordNotice "見積.*確定" $true | Should Be ""
    }

    It "空のワードでは出さない" {
        getWordNotice "" $true | Should Be ""
    }

    It "正規表現として不正なら、文字どおり検索すると伝える" {
        getWordNotice "(" $true | Should Be "正規表現として不正なため、文字どおり検索します。"
    }
}

Describe "newSearchButtonState" -Tag Unit {
    It "検索中は［中止］にする" {
        $state = newSearchButtonState $true $false "見積" $true 1
        $state.Content | Should Be "中止"
        $state.Enabled | Should Be $true
    }

    It "中止を頼んだ後は押せない" {
        (newSearchButtonState $true $true "見積" $true 1).Enabled | Should Be $false
    }

    It "ワード・インデックス・検索対象がそろえば押せる" {
        $state = newSearchButtonState $false $false "見積" $true 2
        $state.Content | Should Be "検索"
        $state.Enabled | Should Be $true
    }

    It "ワードが空なら押せない" {
        (newSearchButtonState $false $false "" $true 2).Enabled | Should Be $false
    }

    It "インデックスが無ければ押せない" {
        (newSearchButtonState $false $false "見積" $false 2).Enabled | Should Be $false
    }

    It "検索対象が選ばれていなければ押せない" {
        (newSearchButtonState $false $false "見積" $true 0).Enabled | Should Be $false
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

Describe "formatLocationLabel" -Tag Unit {
    It "Excel はシート名に「シート」を付ける" {
        formatLocationLabel "見積.xlsx" "4月" | Should Be "シート 4月"
    }

    It "Word のページは「N ページ」にする" {
        formatLocationLabel "報告書.docx" "ページ003" | Should Be "3 ページ"
    }

    It "PowerPoint のスライドは「スライド N」にし、非表示・ノートの印を残す" {
        formatLocationLabel "提案.pptx" "スライド007" | Should Be "スライド 7"
        formatLocationLabel "提案.pptx" "スライド003（非表示）" | Should Be "スライド 3（非表示）"
        formatLocationLabel "提案.pptx" "スライド003_ノート" | Should Be "スライド 3 ノート"
    }

    It "ページ・スライドでない場所はそのまま" {
        formatLocationLabel "報告書.docx" "ヘッダー・フッター" | Should Be "ヘッダー・フッター"
    }
}

Describe "describeFileLocations" -Tag Unit {
    It "場所が無ければ空" {
        describeFileLocations @() | Should Be ""
    }

    It "1 か所ならその場所" {
        describeFileLocations @("シート 4月") | Should Be "シート 4月"
    }

    It "2 か所以上なら先頭と、ほかの数" {
        describeFileLocations @("シート 4月", "シート 5月", "シート 6月") | Should Be "シート 4月 ほか 2 か所"
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
        Order = $order; IsExpanded = $expanded
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