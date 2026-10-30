# インデックスごとの「取り込む拡張子」の文言と判断（tebunko\ui\index\text_ext_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\text_ext_view.ps1"
}

Describe "getIndexTextExtensionRows" -Tag Unit {
    It "件数が無いときは、名前順で、件数は null" {
        $rows = @(getIndexTextExtensionRows $null @())
        $rows.Count | Should -Be 75
        (($rows | ForEach-Object { $_.Extension }) -join ",") | Should -Be ((@(${textExtensions}) | Sort-Object) -join ",")
        @($rows | Where-Object { $null -ne $_.Count }).Count | Should -Be 0
        @($rows | Where-Object { -not $_.Checked }).Count | Should -Be 0
    }

    It "件数の多い順（同じ件数は拡張子の名前順）で、0 件は末尾にまとめる" {
        $counts = @{ ".txt" = 100; ".csv" = 100; ".log" = 5; ".md" = 1 }
        $rows = @(getIndexTextExtensionRows $counts @())
        (($rows[0..3] | ForEach-Object { $_.Extension }) -join ",") | Should -Be ".csv,.txt,.log,.md"
        @($rows | Where-Object { $_.Count -eq 0 }).Count | Should -Be 71
        @($rows[4..74] | Where-Object { $_.Count -ne 0 }).Count | Should -Be 0
        # 0 件の群も名前順
        (($rows[4..74] | ForEach-Object { $_.Extension }) -join ",") | Should -Be ((($rows[4..74] | ForEach-Object { $_.Extension }) | Sort-Object) -join ",")
    }

    It "取り込まない拡張子は外した行（チェックなし・変えた）になる。大文字の設定も当てはまる" {
        $rows = @(getIndexTextExtensionRows @{ ".log" = 3 } @(".LOG", ".md"))
        $log = $rows | Where-Object { $_.Extension -eq ".log" }
        $log.Checked | Should -BeFalse
        $log.Changed | Should -BeTrue
        $log.Count | Should -Be 3
        (@($rows | Where-Object { -not $_.Checked } | ForEach-Object { $_.Extension }) | Sort-Object) -join "," | Should -Be ".log,.md"
    }
}

Describe "convertIndexTextExtensionRows" -Tag Unit {
    It "チェックを外した行を、固定の一覧の順で返す" {
        $rows = @(
            @{ Extension = ".md"; Checked = $false }
            @{ Extension = ".txt"; Checked = $true }
            @{ Extension = ".log"; Checked = $false }
        )
        (@(convertIndexTextExtensionRows $rows) -join ",") | Should -Be ".md,.log"
    }

    It "固定の一覧に無い拡張子は捨てる。全部チェックなら空" {
        @(convertIndexTextExtensionRows @(@{ Extension = ".exe"; Checked = $false })).Count | Should -Be 0
        @(convertIndexTextExtensionRows @(getIndexTextExtensionRows @{} @())).Count | Should -Be 0
    }

    It "行を作って戻すと、設定の値と同じになる（往復。固定の一覧の順）" {
        (@(convertIndexTextExtensionRows @(getIndexTextExtensionRows @{ ".log" = 2 } @(".log", ".md"))) -join ",") | Should -Be ".md,.log"
    }
}

Describe "getIndexTextExtensionSummary / formatTextExtensionCount" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "件数なしで何も外していない"; counts = $null; excluded = @(); expected = "更新すると、件数が出ます" }
        @{ name = "件数なしで外している"; counts = $null; excluded = @(".log", ".md"); expected = "更新すると、件数が出ます" }
        @{ name = "何も外していない"; counts = @{ ".txt" = 13000; ".log" = 150 }; excluded = @(); expected = "取り込み対象 13,150 件（対象外 0 件）" }
        @{ name = "外した"; counts = @{ ".txt" = 10350; ".log" = 2800 }; excluded = @(".log"); expected = "取り込み対象 10,350 件（対象外 2,800 件）" }
        @{ name = "件数 0 の拡張子を外しても、外したことは件数に出る"; counts = @{ ".txt" = 5 }; excluded = @(".log"); expected = "取り込み対象 5 件（対象外 0 件）" }
        @{ name = "テキストが 0 件"; counts = @{}; excluded = @(); expected = "取り込み対象 0 件（対象外 0 件）" }
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

Describe "getIndexTextExtensionListView" -Tag Unit {
    BeforeAll {
        $script:rows = @(getIndexTextExtensionRows @{ ".txt" = 100; ".log" = 20; ".csv" = 5 } @(".log"))
    }

    It "<name>" -TestCases @(
        @{ name = "絞らなければ 75 個すべてと、取り込む件数の要約"; filter = ""; keptOnly = $false; count = 75; summary = "取り込み対象 105 件（対象外 20 件）"; empty = $false }
        @{ name = "［取り込み対象のみ表示］は取り込む 74 個"; filter = ""; keptOnly = $true; count = 74; summary = "取り込み対象 105 件（対象外 20 件）"; empty = $false }
        @{ name = "文字で絞る（点なし・大文字も同じ）"; filter = "LOG"; keptOnly = $false; count = 1; summary = "75 個中 1 個を表示"; empty = $false }
        @{ name = "文字と［取り込み対象のみ表示］は両方に合うものだけ"; filter = "txt"; keptOnly = $true; count = 1; summary = "75 個中 1 個を表示"; empty = $false }
        @{ name = "合うものが無ければ一致なし"; filter = "zzz"; keptOnly = $false; count = 0; summary = "75 個中 0 個を表示"; empty = $true }
    ) {
        param ($name, $filter, $keptOnly, $count, $summary, $empty)
        $view = getIndexTextExtensionListView $script:rows $filter $keptOnly
        @($view.Visible).Count | Should -Be $count
        $view.Summary | Should -Be $summary
        $view.Empty | Should -Be $empty
    }

    It "件数 0 の拡張子の数を返す" {
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

Describe "getIndexTextExtensionCountNote / getIndexTextExtensionDetailRow" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "前回の控え"; source = "previous"; time = "2024/10/14 15:30"; expected = "前回の更新（2024/10/14 15:30）の件数です" }
        @{ name = "控え無し"; source = "none"; time = ""; expected = "更新すると、件数が出ます" }
        @{ name = "日時が空の前回は控え無しと同じ"; source = "previous"; time = ""; expected = "更新すると、件数が出ます" }
        @{ name = "確認ダイアログから開いたときは空"; source = "now"; time = ""; expected = "" }
    ) {
        param ($name, $source, $time, $expected)
        getIndexTextExtensionCountNote $source $time | Should -Be $expected
    }

    It "詳細の行: 控えがあれば日時つきの見出しと件数、無ければ初回の見出し" {
        $row = getIndexTextExtensionDetailRow @(".log") @{ Time = "2024/10/14 15:30"; Counts = @{ ".txt" = 3 } }
        $row.Summary | Should -Be "すべて取り込み（3 件）"
        $row.Caption | Should -Be "前回の更新（2024/10/14 15:30）の件数です"
        $row.Source | Should -Be "previous"
        $row.Counts[".txt"] | Should -Be 3
        $none = getIndexTextExtensionDetailRow @() $null
        $none.Summary | Should -Be "更新すると、件数が出ます"
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
        $note | Should -BeLike "*このインデックス（営業部）から外します。${expected}Excel・Word・PowerPoint は、ここでは選べません。"
    }
}

Describe "75 個の外の値を保つ（addOutOfListExtensions と書き込み）" -Tag Io {
    It "75 個の外を持つ設定で、選びを変えない OK は「変わっていない」になり、書いても 75 個の外が残る" {
        $path = Join-Path $TestDrive "outside_keep.config"
        writeIndexTextExcludes "営業" @(".log", ".exe") $path
        $current = @((readTextExtensionSettings $path).IndexExcludes["営業"])
        $result = @(addOutOfListExtensions @(".log") $current)
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
