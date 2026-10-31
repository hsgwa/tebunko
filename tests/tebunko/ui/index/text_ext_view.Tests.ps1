# インデックスごとの「取り込む拡張子」の文言と判断（tebunko\ui\index\text_ext_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\text_ext_view.ps1"
}

Describe "getIndexExtensionGroups" -Tag Unit {
    It "種類は Excel・Word・PowerPoint・PDF・テキストの順で、拡張子の数は 4・3・3・1・75" {
        $groups = @(getIndexExtensionGroups)
        (($groups | ForEach-Object { $_.Name }) -join ",") | Should -Be "Excel,Word,PowerPoint,PDF,テキスト"
        (($groups | ForEach-Object { @($_.Extensions).Count }) -join ",") | Should -Be "4,3,3,1,75"
        (@($groups[0].Extensions) | Sort-Object) -join "," | Should -Be ".xls,.xlsb,.xlsm,.xlsx"
        @($groups[3].Extensions)[0] | Should -Be ".pdf"
    }

    It "種類の拡張子を合わせると、選べる 86 個と同じ" {
        $all = @(getIndexExtensionGroups | ForEach-Object { $_.Extensions })
        (($all | Sort-Object) -join ",") | Should -Be ((@(getIndexSelectableExtensions) | Sort-Object) -join ",")
    }
}

Describe "getIndexExtensionGroupState / getIndexExtensionGroupToggle / setIndexExtensionGroup" -Tag Unit {
    It "状態: <name>" -TestCases @(
        @{ name = "何も外していない = All"; excluded = @(); expected = "All" }
        @{ name = ".docm だけ外した = Some"; excluded = @(".docm"); expected = "Some" }
        @{ name = "3 つとも外した = None"; excluded = @(".docx", ".docm", ".doc"); expected = "None" }
        @{ name = "ほかの種類の対象外は関係しない"; excluded = @(".xlsx", ".pdf", ".txt"); expected = "All" }
        @{ name = "大文字小文字を区別しない"; excluded = @(".DOCX", ".Docm", ".DOC"); expected = "None" }
    ) {
        param ($name, $excluded, $expected)
        getIndexExtensionGroupState @(".docx", ".docm", ".doc") $excluded | Should -Be $expected
    }

    It "押したあとに取り込み対象にするか: <state> は <expected>" -TestCases @(
        @{ state = "All"; expected = $false }
        @{ state = "Some"; expected = $true }
        @{ state = "None"; expected = $true }
    ) {
        param ($state, $expected)
        getIndexExtensionGroupToggle $state | Should -Be $expected
    }

    It "設定: <name>" -TestCases @(
        @{ name = "一部の Word を全部対象にすると、Word の対象外が消え、テキストの対象外は残る"; excluded = @(".docm", ".log", ".XLSX"); include = $true; expected = ".log,.XLSX" }
        @{ name = "全部外すと、その種類の全部が入る。ほかの種類は変えない"; excluded = @(".log"); include = $false; expected = ".log,.docx,.docm,.doc" }
        @{ name = "全部外すとき、すでに外していた分と重ならない"; excluded = @(".docm", ".DOC"); include = $false; expected = ".docx,.docm,.doc" }
        @{ name = "何も外していない種類を対象にしても変わらない"; excluded = @(".log"); include = $true; expected = ".log" }
    ) {
        param ($name, $excluded, $include, $expected)
        (@(setIndexExtensionGroup $excluded @(".docx", ".docm", ".doc") $include) -join ",") | Should -Be $expected
    }
}

Describe "getIndexTextExtensionRows" -Tag Unit {
    It "件数が無いときは、見出しを挟んで種類ごとに名前順。86 個の行と 5 つの見出し、件数は null" {
        $rows = @(getIndexTextExtensionRows $null @())
        $rows.Count | Should -Be 91
        @($rows | Where-Object { $_.Heading }).Count | Should -Be 5
        (($rows | Where-Object { $_.Heading } | ForEach-Object { $_.Group }) -join ",") | Should -Be "Excel,Word,PowerPoint,PDF,テキスト"
        $rows[0].Heading | Should -BeTrue
        $rows[0].Group | Should -Be "Excel"
        (($rows[1..4] | ForEach-Object { $_.Extension }) -join ",") | Should -Be ".xls,.xlsb,.xlsm,.xlsx"
        $textRows = @($rows | Where-Object { -not $_.Heading -and $_.Group -eq "テキスト" })
        (($textRows | ForEach-Object { $_.Extension }) -join ",") | Should -Be ((@(${textExtensions}) | Sort-Object) -join ",")
        @($rows | Where-Object { -not $_.Heading -and $null -ne $_.Count }).Count | Should -Be 0
        @($rows | Where-Object { -not $_.Checked }).Count | Should -Be 0
    }

    It "件数の多い順（同じ件数は拡張子の名前順）で、0 件は種類の中の末尾" {
        $counts = @{ ".txt" = 100; ".csv" = 100; ".log" = 5; ".md" = 1; ".xlsx" = 30; ".xls" = 30; ".xlsm" = 2; ".pdf" = 7 }
        $rows = @(getIndexTextExtensionRows $counts @())
        # Excel: 件数の多い順で、0 件の .xlsb が最後
        (($rows[1..4] | ForEach-Object { $_.Extension }) -join ",") | Should -Be ".xls,.xlsx,.xlsm,.xlsb"
        $text = @($rows | Where-Object { -not $_.Heading -and $_.Group -eq "テキスト" })
        (($text[0..3] | ForEach-Object { $_.Extension }) -join ",") | Should -Be ".csv,.txt,.log,.md"
        @($text | Where-Object { $_.Count -eq 0 }).Count | Should -Be 71
        @($text[4..74] | Where-Object { $_.Count -ne 0 }).Count | Should -Be 0
        (($text[4..74] | ForEach-Object { $_.Extension }) -join ",") | Should -Be ((($text[4..74] | ForEach-Object { $_.Extension }) | Sort-Object) -join ",")
        # 見出しは、種類の先頭
        $rows[$rows.Count - 76].Heading | Should -BeTrue
        $rows[$rows.Count - 76].Group | Should -Be "テキスト"
    }

    It "取り込まない拡張子は外した行（チェックなし・変えた）になる。大文字の設定も当てはまる" {
        $rows = @(getIndexTextExtensionRows @{ ".log" = 3 } @(".LOG", ".md", ".DOCM"))
        $log = $rows | Where-Object { $_.Extension -eq ".log" }
        $log.Checked | Should -BeFalse
        $log.Changed | Should -BeTrue
        $log.Count | Should -Be 3
        (@($rows | Where-Object { -not $_.Heading -and -not $_.Checked } | ForEach-Object { $_.Extension }) | Sort-Object) -join "," | Should -Be ".docm,.log,.md"
    }

    It "見出しの行の状態は、その種類の取り込み対象の数から決まる" {
        $rows = @(getIndexTextExtensionRows $null @(".docm", ".pptx", ".pptm", ".ppt"))
        $states = @{}
        foreach ($row in @($rows | Where-Object { $_.Heading })) { $states[$row.Group] = $row.State }
        $states["Excel"] | Should -Be "All"
        $states["Word"] | Should -Be "Some"
        $states["PowerPoint"] | Should -Be "None"
        $states["PDF"] | Should -Be "All"
        $states["テキスト"] | Should -Be "All"
        ($rows | Where-Object { $_.Heading -and $_.Group -eq "Excel" }).Checked | Should -BeTrue
        ($rows | Where-Object { $_.Heading -and $_.Group -eq "Word" }).Checked | Should -BeFalse
    }
}

Describe "convertIndexTextExtensionRows" -Tag Unit {
    It "チェックを外した行を、選べる一覧の順で返す（Office・PDF も含む）" {
        $rows = @(
            @{ Extension = ".md"; Checked = $false }
            @{ Extension = ".txt"; Checked = $true }
            @{ Extension = ".pdf"; Checked = $false }
            @{ Extension = ".XLSX"; Checked = $false }
            @{ Extension = ".log"; Checked = $false }
        )
        (@(convertIndexTextExtensionRows $rows) -join ",") | Should -Be ".xlsx,.pdf,.md,.log"
    }

    It "固定の一覧に無い拡張子と、見出しの行は捨てる。全部チェックなら空" {
        @(convertIndexTextExtensionRows @(@{ Extension = ".exe"; Checked = $false })).Count | Should -Be 0
        @(convertIndexTextExtensionRows @(@{ Heading = $true; Group = "Excel"; Extension = ""; Checked = $false })).Count | Should -Be 0
        @(convertIndexTextExtensionRows @(getIndexTextExtensionRows @{} @())).Count | Should -Be 0
    }

    It "行を作って戻すと、設定の値と同じになる（往復。選べる一覧の順）" {
        (@(convertIndexTextExtensionRows @(getIndexTextExtensionRows @{ ".log" = 2 } @(".log", ".md", ".pdf", ".docm"))) -join ",") | Should -Be ".docm,.pdf,.md,.log"
    }
}

Describe "getIndexTextExtensionSummary / formatTextExtensionCount" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "件数なしで何も外していない"; counts = $null; excluded = @(); expected = "更新すると、件数が出ます" }
        @{ name = "件数なしで外している"; counts = $null; excluded = @(".log", ".md"); expected = "更新すると、件数が出ます" }
        @{ name = "何も外していない"; counts = @{ ".txt" = 13000; ".log" = 150 }; excluded = @(); expected = "取り込み対象 13,150 件（対象外 0 件）" }
        @{ name = "外した"; counts = @{ ".txt" = 10350; ".log" = 2800 }; excluded = @(".log"); expected = "取り込み対象 10,350 件（対象外 2,800 件）" }
        @{ name = "Office・PDF の件数も含む"; counts = @{ ".txt" = 100; ".xlsx" = 50; ".pdf" = 30 }; excluded = @(".pdf"); expected = "取り込み対象 150 件（対象外 30 件）" }
        @{ name = "件数 0 の拡張子を外しても、外したことは件数に出る"; counts = @{ ".txt" = 5 }; excluded = @(".log"); expected = "取り込み対象 5 件（対象外 0 件）" }
        @{ name = "どれも 0 件"; counts = @{}; excluded = @(); expected = "取り込み対象 0 件（対象外 0 件）" }
    ) {
        param ($name, $counts, $excluded, $expected)
        getIndexTextExtensionSummary @(getIndexTextExtensionRows $counts $excluded) | Should -Be $expected
    }

    It "件数の文字は 3 桁区切り、数え中は …" {
        formatTextExtensionCount 1234567 | Should -Be "1,234,567"
        formatTextExtensionCount 0 | Should -Be "0"
        formatTextExtensionCount $null | Should -Be "…"
    }
}

Describe "getIndexTextExtensionBriefSummary" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "何も外していない"; excluded = @(); expected = "すべて取り込み" }
        @{ name = "テキストを外した"; excluded = @(".log"); expected = "対象外あり" }
        @{ name = "Office を外した"; excluded = @(".XLSX"); expected = "対象外あり" }
        @{ name = "PDF を外した"; excluded = @(".pdf"); expected = "対象外あり" }
        @{ name = "選べる一覧の外だけなら外していないのと同じ"; excluded = @(".exe"); expected = "すべて取り込み" }
    ) {
        param ($name, $excluded, $expected)
        getIndexTextExtensionBriefSummary $excluded | Should -Be $expected
    }
}

Describe "getIndexTextExtensionListView" -Tag Unit {
    BeforeAll {
        $script:rows = @(getIndexTextExtensionRows @{ ".txt" = 100; ".log" = 20; ".csv" = 5; ".xlsx" = 9 } @(".log"))
        $script:wordOff = @(getIndexTextExtensionRows @{ ".txt" = 100 } @(".docx", ".docm", ".doc"))
    }

    It "<name>" -TestCases @(
        @{ name = "絞らなければ 86 個すべてと、取り込む件数の要約"; filter = ""; keptOnly = $false; count = 86; summary = "取り込み対象 114 件（対象外 20 件）"; empty = $false }
        @{ name = "［取り込み対象のみ表示］は取り込む 85 個"; filter = ""; keptOnly = $true; count = 85; summary = "取り込み対象 114 件（対象外 20 件）"; empty = $false }
        @{ name = "文字で絞る（点なし・大文字も同じ）"; filter = "LOG"; keptOnly = $false; count = 1; summary = "86 個中 1 個を表示"; empty = $false }
        @{ name = "文字と［取り込み対象のみ表示］は両方に合うものだけ"; filter = "txt"; keptOnly = $true; count = 1; summary = "86 個中 1 個を表示"; empty = $false }
        @{ name = "Office の拡張子でも絞れる"; filter = "xls"; keptOnly = $false; count = 4; summary = "86 個中 4 個を表示"; empty = $false }
        @{ name = "合うものが無ければ一致なし"; filter = "zzz"; keptOnly = $false; count = 0; summary = "86 個中 0 個を表示"; empty = $true }
    ) {
        param ($name, $filter, $keptOnly, $count, $summary, $empty)
        $view = getIndexTextExtensionListView $script:rows $filter $keptOnly
        @($view.Visible).Count | Should -Be $count
        $view.Summary | Should -Be $summary
        $view.Empty | Should -Be $empty
    }

    It "見出しは、見せる行が 1 つ以上ある種類だけ残す（［取り込み対象のみ表示］では、全部外した種類の見出しが消える）" {
        (getIndexTextExtensionListView $script:rows "" $false).VisibleGroups -join "," | Should -Be "Excel,Word,PowerPoint,PDF,テキスト"
        (getIndexTextExtensionListView $script:wordOff "" $true).VisibleGroups -join "," | Should -Be "Excel,PowerPoint,PDF,テキスト"
        (getIndexTextExtensionListView $script:rows "xls" $false).VisibleGroups -join "," | Should -Be "Excel"
    }

    It "見出しの行を渡しても、拡張子の行だけを数える" {
        $view = getIndexTextExtensionListView @(getIndexTextExtensionRows @{ ".txt" = 1 } @()) "" $false
        @($view.Visible).Count | Should -Be 86
    }

    It "件数 0 の拡張子の数を返す（折りたたむのはテキストの中だけ）" {
        (getIndexTextExtensionListView $script:rows "" $false).ZeroCount | Should -Be 72
        (getIndexTextExtensionListView @(getIndexTextExtensionRows $null @()) "" $false).ZeroCount | Should -Be 0
    }
}

Describe "testIndexTextExcludesChanged" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "同じ"; current = @(".log"); applied = @(".log"); expected = $false }
        @{ name = "違う"; current = @(".log"); applied = @(".csv"); expected = $true }
        @{ name = "順だけ違う"; current = @(".log", ".csv"); applied = @(".csv", ".log"); expected = $false }
        @{ name = "大文字小文字だけ違う"; current = @(".LOG"); applied = @(".log"); expected = $false }
        @{ name = "変えて戻した（差なし）"; current = @(); applied = @(); expected = $false }
        @{ name = "前回が控え無し（null）で今が空 = 差なし"; current = @(); applied = $null; expected = $false }
        @{ name = "前回が空で今が空でない = 差あり"; current = @(".log"); applied = @(); expected = $true }
        @{ name = "前回が控え無しで今が空でない = 差あり"; current = @(".log"); applied = $null; expected = $true }
        @{ name = "今が空で前回が空でない = 差あり"; current = $null; applied = @(".log"); expected = $true }
    ) {
        param ($name, $current, $applied, $expected)
        testIndexTextExcludesChanged $current $applied | Should -Be $expected
    }
}

Describe "getIndexTextExtensionDetailRow" -Tag Unit {
    It "詳細の行: 要約は件数を言わない。控えの日時と件数は、ダイアログに渡すために持つ" {
        $row = getIndexTextExtensionDetailRow @(".log") @{ Time = "2024/10/14 15:30"; Counts = @{ ".txt" = 3 } }
        $row.Summary | Should -Be "対象外あり"
        $row.Caption | Should -Be ""
        $row.Time | Should -Be "2024/10/14 15:30"
        $row.Source | Should -Be "previous"
        $row.Counts[".txt"] | Should -Be 3
        $none = getIndexTextExtensionDetailRow @() $null
        $none.Summary | Should -Be "すべて取り込み"
        $none.Caption | Should -Be ""
        $none.Source | Should -Be "none"
        $none.Counts | Should -BeNullOrEmpty
    }
}

Describe "getIndexTextExtensionZeroHeader" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "折りたたみ中"; expanded = $false; expected = "$([char]0x25B8) 件数 0 の拡張子 62 個（折りたたみ中。開くと選べます）" }
        @{ name = "開いているとき"; expanded = $true; expected = "$([char]0x25BE) 件数 0 の拡張子 62 個（開いて選ぶ。見つかったときに効く設定）" }
    ) {
        param ($name, $expanded, $expected)
        getIndexTextExtensionZeroHeader 62 $expanded | Should -Be $expected
    }
}

Describe "getIndexTextExtensionDialogTexts" -Tag Unit {
    It "題と注にインデックスの名前が入り、文言がそろっている" {
        $texts = getIndexTextExtensionDialogTexts "営業部"
        $texts.Title | Should -Be "取り込み対象の拡張子 - 営業部"
        $texts.Note | Should -BeLike "*このインデックス（営業部）から外します*"
        foreach ($key in @("Title", "Note", "Extension", "Count", "Reset", "UncheckAll", "Ok", "Cancel", "KeptOnly", "FilterPlaceholder", "Empty")) {
            $texts[$key] | Should -Not -BeNullOrEmpty -Because $key
        }
    }

    It "注の件数の文は、件数の出どころで 3 通りに切り替わる" -TestCases @(
        @{ source = "previous"; time = "2024/10/14 15:30"; expected = "件数は、前回の更新（2024/10/14 15:30）で数えた結果です。" }
        @{ source = "now"; time = ""; expected = "件数は、いま数えた結果です。" }
        @{ source = "none"; time = ""; expected = "まだ件数がありません。更新すると、件数が出ます。" }
        @{ source = "previous"; time = ""; expected = "まだ件数がありません。更新すると、件数が出ます。" }
    ) {
        param ($source, $time, $expected)
        $note = (getIndexTextExtensionDialogTexts "営業部" $source $time).Note
        $note | Should -BeLike "*このインデックス（営業部）から外します。${expected}"
        $note | Should -Not -BeLike "*ここでは選べません*"
    }
}

Describe "選べる 86 個の外の値を保つ（addOutOfListExtensions と書き込み）" -Tag Io {
    It "86 個の外を持つ設定で、選びを変えない OK は「変わっていない」になり、書いても 86 個の外が残る" {
        $path = Join-Path $TestDrive "outside_keep.config"
        writeIndexTextExcludes "営業" @(".log", ".exe") $path
        $current = @((readTextExtensionSettings $path).IndexExcludes["営業"])
        $result = @(addOutOfListExtensions @(".log") $current @(getIndexSelectableExtensions))
        testIndexTextExcludesChanged $result $current | Should -BeFalse
        writeIndexTextExcludes "営業" $result $path
        ((readTextExtensionSettings $path).IndexExcludes["営業"] | Sort-Object) -join "," | Should -Be ".exe,.log"
    }
}

Describe "getTextExtCountsFromPlan" -Tag Unit {
    It "取り込んだ行（kind の行）だけを、インデックス名 → @{ Time; Counts } にする。取り込まない行・null は入れない" {
        $plan = @(
            [pscustomobject]@{ 区分 = "取り込む"; インデックス名 = "営業"; TextExtensionCounts = @{ ".txt" = 3; ".log" = 1 } },
            [pscustomobject]@{ 区分 = "対象外"; インデックス名 = "経理"; TextExtensionCounts = @{ ".md" = 1 } },
            $null
        )
        $result = getTextExtCountsFromPlan $plan "取り込む" "2024/10/14 15:30"
        @($result.Keys) | Should -Be @("営業")
        $result["営業"].Time | Should -Be "2024/10/14 15:30"
        $result["営業"].Counts[".txt"] | Should -Be 3
        $result["営業"].Counts[".log"] | Should -Be 1
    }

    It "計画が空なら空" {
        (getTextExtCountsFromPlan @() "取り込む" "t").Count | Should -Be 0
        (getTextExtCountsFromPlan $null "取り込む" "t").Count | Should -Be 0
    }
}
