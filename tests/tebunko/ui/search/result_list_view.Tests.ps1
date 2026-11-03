# 結果の一覧の判断（tebunko\ui\search\result_list_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_list_view.ps1"

    # 結果の表の項目のテストに使う、HitRow・FileGroup と同じ項目を持つもの
    function newTestRow {
        param ([int]$order, [string]$line, [int]$lineNumber = 1, [string]$location = "", [string]$book = "")

        $row = [pscustomobject]@{ Order = $order; Line = $line; LineNumber = $lineNumber; Location = $location; Book = $book }
        $row | Add-Member -MemberType ScriptMethod -Name MatchesFileName -Value { param ($text) $this.Book.Contains($text) }
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

Describe "getSearchCountingText" -Tag Unit {
    It "上限と、数えた件数を 3 桁区切りで出す" {
        getSearchCountingText 10000 45678 | Should -Be "検索中…　該当 10,000 件以上（総数を数えています：45,678 件）"
    }
}

Describe "getSearchFinishSummaryText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "上限以下は今の文（0 件でも上限を超えていなければ）"; a = @(9999, 3, 2.3, $false, -1); expected = "一致 9,999 件（3 ファイル）・2.3 秒" }
        @{ name = "ちょうど上限は超えていない"; a = @(10000, 3, 2.3, $false, 10000); expected = "一致 10,000 件（3 ファイル）・2.3 秒" }
        @{ name = "上限を超えて数え終えた（10,001 件）"; a = @(10000, 1234, 2.3, $true, 10001); expected = "全 10,001 件中 10,000 件を表示（1,234 ファイル）・2.3 秒" }
        @{ name = "上限を超えて数え終えた（234,567 件）"; a = @(10000, 1234, 2.3, $true, 234567); expected = "全 234,567 件中 10,000 件を表示（1,234 ファイル）・2.3 秒" }
        @{ name = "絞り込みを消したときは秒を付けない"; a = @(10000, 1234, -1, $true, 234567); expected = "全 234,567 件中 10,000 件を表示（1,234 ファイル）" }
        @{ name = "数えている間に止まった（総数 -1）"; a = @(10000, 1234, 2.3, $true, -1); expected = "該当 10,000 件以上（総数は数えていません）（1,234 ファイル）・2.3 秒" }
        @{ name = "数えている間に止まった・秒なし"; a = @(10000, 1234, -1, $true, -1); expected = "該当 10,000 件以上（総数は数えていません）（1,234 ファイル）" }
    ) {
        param ($name, $a, $expected)
        getSearchFinishSummaryText $a[0] $a[1] $a[2] $a[3] $a[4] | Should -Be $expected
    }
}

Describe "getSearchFinishStatusText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "上限以下は今の文"; a = @("見積", 0, "", $false, -1, $false, $false); expected = "検索しました：見積 0 件" }
        @{ name = "上限以下は条件を後ろに付ける"; a = @("見積", 9999, "大文字・小文字を区別", $false, 9999, $false, $false); expected = "検索しました：見積 9,999 件　条件：大文字・小文字を区別" }
        @{ name = "上限に達する前に中止した"; a = @("見積", 1234, "", $false, -1, $false, $true); expected = "中止しました（1,234 件まで表示）" }
        @{ name = "上限を超えて数え終えた（10,001 件）"; a = @("見積", 10000, "", $true, 10001, $false, $false); expected = "検索しました：見積（全 10,001 件のうち、先頭の 10,000 件を表示）。フォルダ・ファイル名の順に先頭から表示しています。ワード・対象ファイル・検索対象で絞り込んでください。" }
        @{ name = "上限を超えて数え終えた・条件つき（234,567 件）"; a = @("見積", 10000, "図形を除く", $true, 234567, $false, $false); expected = "検索しました：見積（全 234,567 件のうち、先頭の 10,000 件を表示）。フォルダ・ファイル名の順に先頭から表示しています。ワード・対象ファイル・検索対象で絞り込んでください。　条件：図形を除く" }
        @{ name = "数えている間に中止した"; a = @("見積", 10000, "", $true, -1, $false, $true); expected = "中止しました（先頭の 10,000 件を表示。総数は数えていません）" }
        @{ name = "数えている間に時間切れになった"; a = @("見積", 10000, "", $true, -1, $true, $false); expected = "正規表現の照合に時間がかかりすぎるため、総数を数えられませんでした（先頭の 10,000 件を表示）。正規表現を見直してください。" }
    ) {
        param ($name, $a, $expected)
        getSearchFinishStatusText $a[0] $a[1] $a[2] $a[3] $a[4] $a[5] $a[6] | Should -Be $expected
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

Describe "getSkippedContentIndexText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "0 件は空"; count = 0; expected = "" }
        @{ name = "負の件数も空"; count = -1; expected = "" }
        @{ name = "1 件"; count = 1; expected = "　読めなかったインデックスのファイル：1 件" }
        @{ name = "件数を 3 桁区切りで出す"; count = 1234; expected = "　読めなかったインデックスのファイル：1,234 件" }
    ) {
        param ($name, $count, $expected)
        getSkippedContentIndexText $count | Should -Be $expected
    }
}

Describe "getFilteredSummaryText" -Tag Unit {
    It "全部の件数のうち、見せている件数を出す" {
        getFilteredSummaryText 1234 56 | Should -Be "一致 1,234 件中 56 件を表示"
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

Describe "getFileHeaderLocationPrefix" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "場所が無ければ空"; locations = @(); expected = "" }
        @{ name = "1 か所だけなら行の場所列と重なるので空"; locations = @("[シート]4月"); expected = "" }
        @{ name = "2 か所以上なら先頭と、ほかの数に区切りを付ける"; locations = @("[シート]4月", "[シート]5月", "[シート]6月"); expected = "[シート]4月 ほか 2 か所 ・ " }
    ) {
        param ($name, $locations, $expected)
        getFileHeaderLocationPrefix $locations | Should -Be $expected
    }
}

Describe "getPreviewRowHeight" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "内容を測らない（多い）ときは今までの高さ"; available = 600; content = $null; expected = 180 }
        @{ name = "表が内容より低くなるときは今までの高さ"; available = 600; content = 500; expected = 180 }
        @{ name = "結果が少なければ、表を内容の高さにしてプレビューを広げる"; available = 600; content = 150; expected = 450 }
        @{ name = "表は最低の高さより低くしない"; available = 600; content = 10; expected = 540 }
        @{ name = "全体が低くてもプレビューは今までの高さより低くしない"; available = 200; content = 20; expected = 180 }
        @{ name = "行がまだ無い（起動直後・検索前・0 件）ときは、内容が小さくても今までの高さ"; available = 600; content = 40; itemCount = 0; expected = 180 }
        @{ name = "行が 1 つでもあれば、広げる"; available = 600; content = 150; itemCount = 1; expected = 450 }
    ) {
        param ($name, $available, $content, $expected, $itemCount = -1)
        getPreviewRowHeight $available $content -itemCount $itemCount | Should -Be $expected
    }
}

Describe "getGroupRowTarget" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "ヒットが上限以下なら全件"; hits = 800; limit = 0; expected = 800 }
        @{ name = "ヒットが多いなら既定の上限（1,000）まで"; hits = 100000; limit = 0; expected = 1000 }
        @{ name = "続きを出したあとは、増えた上限まで"; hits = 100000; limit = 3000; expected = 3000 }
        @{ name = "増えた上限よりヒットが少なければ全件"; hits = 2500; limit = 3000; expected = 2500 }
    ) {
        param ($name, $hits, $limit, $expected)
        getGroupRowTarget $hits $limit | Should -Be $expected
    }
}

Describe "getNextRowLimit" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "最初（0）は上限の次"; limit = 0; expected = 2000 }
        @{ name = "続きを出すたびに、上限の分ずつ増える"; limit = 2000; expected = 3000 }
    ) {
        param ($name, $limit, $expected)
        getNextRowLimit $limit | Should -Be $expected
    }
}

Describe "getSortRowTarget" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "上限以下なら並べ替えるので全件"; hits = 1000; expected = 1000 }
        @{ name = "上限を超えるなら並べ替えないので 1 件"; hits = 1001; expected = 1 }
    ) {
        param ($name, $hits, $expected)
        getSortRowTarget $hits | Should -Be $expected
    }
}

Describe "getMoreRowsText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "残りが多いときは上限の件数ずつ"; remaining = 99000; expected = "続きの 1,000 件を表示（残り 99,000 件）" }
        @{ name = "残りが上限より少ないときは残りの件数"; remaining = 250; expected = "続きの 250 件を表示（残り 250 件）" }
    ) {
        param ($name, $remaining, $expected)
        getMoreRowsText $remaining | Should -Be $expected
    }
}

Describe "getSortedHitOrder（getHitOrderByLine・getHitOrderByNumber）" -Tag Unit {
    BeforeAll {
        function newHit([string]$location, [int]$lineNumber, [string]$line) {
            return [pscustomobject]@{ Location = $location; LineNumber = $lineNumber; Line = $line }
        }
    }

    It "<name>" -TestCases @(
        @{ name = "行（該当行）の昇順。大文字・小文字は区別せず、同じ値は見つかった順"; property = "Line"; descending = $false; expected = "1,3,0,2" }
        @{ name = "行（該当行）の降順。同じ値は見つかった順のまま"; property = "Line"; descending = $true; expected = "0,2,1,3" }
        @{ name = "行番号の昇順。同じ値は見つかった順"; property = "LineNumber"; descending = $false; expected = "1,3,0,2" }
        @{ name = "行番号の降順。同じ値は見つかった順のまま"; property = "LineNumber"; descending = $true; expected = "0,2,1,3" }
    ) {
        param ($name, $property, $descending, $expected)
        # 該当行と行番号の順位が同じになるように作る（0: b・5 / 1: A・1 / 2: B・5 / 3: a・1）
        $hits = @((newHit "シート" 5 "b"), (newHit "シート" 1 "A"), (newHit "シート" 5 "B"), (newHit "シート" 1 "a"))
        (getSortedHitOrder $hits $property $descending) -join "," | Should -Be $expected
    }

    It "場所の順 → 行番号の順（A10 が A9 の前に来ない）。降順はその逆で、同じ値は見つかった順" {
        $hits = @((newHit "売上" 10 "x"), (newHit "売上" 9 "x"), (newHit "仕入" 1 "x"), (newHit "売上" 9 "x"))
        (getSortedHitOrder $hits "Location" $false) -join "," | Should -Be "2,1,3,0"
        (getSortedHitOrder $hits "Location" $true) -join "," | Should -Be "0,1,3,2"
    }

    It "該当行が無い（null）ヒットは空として扱う" {
        $hits = @((newHit "シート" 1 "b"), (newHit "シート" 2 $null), (newHit "シート" 3 "a"))
        (getSortedHitOrder $hits "Line" $false) -join "," | Should -Be "1,2,0"
    }

    It "種別（Kind）は、場所ごとの種別の順。同じ種別の中は見つかった順（降順も同じ）" {
        $hits = @((newHit "売上" 1 "x"), (newHit "売上[図形]" 1 "x"), (newHit "仕入" 1 "x"), (newHit "売上[コメント]" 1 "x"))
        $kinds = @{ "売上" = "セル"; "仕入" = "セル"; "売上[図形]" = "図形"; "売上[コメント]" = "コメント" }
        $asc = getSortedHitOrder $hits "Kind" $false $kinds
        # カルチャーの順は Sort-Object と同じ
        $expected = @($hits | ForEach-Object { $kinds[$_.Location] } | Sort-Object -Unique)
        (@($asc | ForEach-Object { $kinds[$hits[$_].Location] } | Select-Object -Unique)) -join "," | Should -Be ($expected -join ",")
        $ties = @($asc | Where-Object { $kinds[$hits[$_].Location] -eq "セル" })
        $ties -join "," | Should -Be "0,2"
        $desc = getSortedHitOrder $hits "Kind" $true $kinds
        (@($desc | ForEach-Object { $kinds[$hits[$_].Location] } | Select-Object -Unique)) -join "," | Should -Be ((@($expected)[($expected.Count - 1)..0]) -join ",")
        @($desc | Where-Object { $kinds[$hits[$_].Location] -eq "セル" }) -join "," | Should -Be "0,2"
    }

    It "並べ替えない項目は null" {
        $null -eq (getSortedHitOrder @((newHit "シート" 1 "a")) "Order" $false) | Should -Be $true
    }

    It "ヒットが 0 件でも落ちない" {
        (getSortedHitOrder @() "Line" $false).Count | Should -Be 0
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

Describe "getGroupDisplayCount" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "ヒットを持たなければ、作ってある行すべて"; shown = 7; hits = 0; limit = 0; expected = 7 }
        @{ name = "上限を超えるファイルは、全件の行を作ってあっても上限まで"; shown = 5000; hits = 5000; limit = 0; expected = 1000 }
        @{ name = "続きを出したあとは、増えた上限まで"; shown = 5000; hits = 5000; limit = 2000; expected = 2000 }
        @{ name = "作ってある行が上限より少なければ、その数"; shown = 300; hits = 5000; limit = 0; expected = 300 }
    ) {
        param ($name, $shown, $hits, $limit, $expected)
        getGroupDisplayCount $shown $hits $limit | Should -Be $expected
    }
}

Describe "getResultItems（続きの行）" -Tag Unit {
    It "全件の行を作ってあっても、上限までの行だけ並べる" {
        $rows = 1..5 | ForEach-Object { newTestRow $_ "a$_" }
        $a = newTestGroup 1 $rows $true
        $a | Add-Member -NotePropertyName HitCount -NotePropertyValue 2000
        $a | Add-Member -NotePropertyName RowLimit -NotePropertyValue 3
        $items = getResultItems @($a)
        $items.Count | Should -Be 4
        $items[3].Line | Should -Be "a3"
    }

    It "開いているファイルだけ、行の末尾に続きの行を並べる" {
        $a = newTestGroup 1 @((newTestRow 1 "a1")) $true
        $a | Add-Member -NotePropertyName MoreRow -NotePropertyValue "続き A"
        $b = newTestGroup 2 @((newTestRow 2 "b1"))
        $b | Add-Member -NotePropertyName MoreRow -NotePropertyValue "続き B"
        $items = getResultItems @($a, $b)
        $items.Count | Should -Be 4
        $items[2] | Should -Be "続き A"
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

Describe "sortFileGroups（ヒットが多いファイル）" -Tag Unit {
    It "上限を超えるファイルの行は並べ替えず（先に並べ替えたヒットの順に作ってある）、ファイルの順は先頭の行の値で決める" {
        # big は先頭の行が最小（1）になるよう、並べ替えたヒットの順に作ってある行
        $big = newTestGroup 1 @((newTestRow 9 "x" 1), (newTestRow 0 "x" 30), (newTestRow 3 "x" 4))
        $big | Add-Member -NotePropertyName HitCount -NotePropertyValue 5
        $small = newTestGroup 2 @((newTestRow 5 "y" 20), (newTestRow 6 "y" 10))
        $sorted = sortFileGroups @($small, $big) "LineNumber" $false 3
        @($big.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "1,30,4"
        @($small.Rows | ForEach-Object { $_.LineNumber }) -join "," | Should -Be "10,20"
        [object]::ReferenceEquals($sorted[0], $big) | Should -Be $true
        [object]::ReferenceEquals($sorted[1], $small) | Should -Be $true
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
        @{ expanded = $true;  expected = "*読み取り専用で開く/開く/新規で開く/フォルダを開く/-/ファイルのパスをコピー/この結果を折りたたむ" }
        @{ expanded = $false; expected = "*読み取り専用で開く/開く/新規で開く/フォルダを開く/-/ファイルのパスをコピー/この結果を開く" }
    ) {
        param ($expanded, $expected)
        # 既定の開き方が別でも、見出しのメニューは変わらない
        menuText (getResultMenuItems "group" "new" $expanded) | Should -Be $expected
    }

    It "読み取り専用で開けないファイルは、［読み取り専用で開く］を出さず、太字は［開く］にする（<target>・既定: <mode>）" -TestCases @(
        @{ target = "row";   mode = "normal";   expected = "*開く/新規で開く/フォルダを開く/-/選んだ行をコピー/ファイルのパスをコピー" }
        @{ target = "row";   mode = "readOnly"; expected = "*開く/新規で開く/フォルダを開く/-/選んだ行をコピー/ファイルのパスをコピー" }
        @{ target = "row";   mode = "new";      expected = "*新規で開く/開く/フォルダを開く/-/選んだ行をコピー/ファイルのパスをコピー" }
        @{ target = "group"; mode = "readOnly"; expected = "*開く/新規で開く/フォルダを開く/-/ファイルのパスをコピー/この結果を開く" }
    ) {
        param ($target, $mode, $expected)
        menuText (getResultMenuItems $target $mode $false $false) | Should -Be $expected
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

Describe "getNoHitText" -Tag Unit {
    It "0 件のときの文を返す" {
        getNoHitText | Should -Be "一致する結果はありません"
    }
}