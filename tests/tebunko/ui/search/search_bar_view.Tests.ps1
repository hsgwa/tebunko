# 検索バーの判断（tebunko\ui\search\search_bar_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\search\search_bar_view.ps1"
}

Describe "describeSearchOption" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "既定のままなら空"; option = @{ CaseSensitive = $false; FileKinds = @() }; expected = "" }
        @{ name = "大文字・小文字の区別を出す"; option = @{ CaseSensitive = $true; FileKinds = @() }; expected = "大文字・小文字を区別" }
        @{ name = "種類を絞っていれば出す"; option = @{ CaseSensitive = $false; FileKinds = @("excel", "text") }; expected = "種類：Excel・テキスト" }
        @{ name = "すべての種類なら出さない"; option = @{ CaseSensitive = $false; FileKinds = @("excel", "word", "powerpoint", "text") }; expected = "" }
        @{ name = "両方あれば中黒でつなぐ"; option = @{ CaseSensitive = $true; FileKinds = @("word") }; expected = "大文字・小文字を区別・種類：Word" }
        @{ name = "図形・コメントを含めるなら出さない"; option = @{ CaseSensitive = $false; IncludeShapes = $true; IncludeComments = $true }; expected = "" }
        @{ name = "図形・コメントを外したら出す"; option = @{ CaseSensitive = $false; IncludeShapes = $false; IncludeComments = $false }; expected = "図形を除く・コメントを除く" }
        @{ name = "コメントだけ外したらコメントだけ出す"; option = @{ CaseSensitive = $false; IncludeComments = $false }; expected = "コメントを除く" }
    ) {
        param ($name, $option, $expected)
        describeSearchOption $option | Should -Be $expected
    }
}

Describe "種類のチップ" -Tag Unit {
    It "getFileKindLabel: <kind> は <expected>" -TestCases @(
        @{ kind = "excel"; expected = "Excel" }
        @{ kind = "word"; expected = "Word" }
        @{ kind = "powerpoint"; expected = "PowerPoint" }
        @{ kind = "text"; expected = "テキスト" }
        @{ kind = "pdf"; expected = "" }
    ) {
        param ($kind, $expected)
        getFileKindLabel $kind | Should -Be $expected
    }

    It "toggleFileKind: <name>" -TestCases @(
        @{ name = "外す"; kinds = @("excel", "word", "powerpoint", "text"); kind = "word"; expected = "excel,powerpoint,text" }
        @{ name = "足す（決まった順に並ぶ）"; kinds = @("text", "excel"); kind = "word"; expected = "excel,word,text" }
        @{ name = "最後の 1 つも外せる（1 つも選ばない）"; kinds = @("excel"); kind = "excel"; expected = "" }
        @{ name = "空（1 つも選んでいない）から 1 つ選べる"; kinds = @(); kind = "text"; expected = "text" }
        @{ name = "知らない種類は変えない"; kinds = @("excel", "text"); kind = "pdf"; expected = "excel,text" }
    ) {
        param ($name, $kinds, $kind, $expected)
        (toggleFileKind $kinds $kind) -join "," | Should -Be $expected
    }

    It "getSearchKindError: <name>" -TestCases @(
        @{ name = "1 つも選んでいなければ知らせる"; kinds = @(); expected = "検索する種類を 1 つ以上選んでください。" }
        @{ name = "知らない種類だけでも 1 つも選んでいない"; kinds = @("pdf"); expected = "検索する種類を 1 つ以上選んでください。" }
        @{ name = "1 つ選んでいれば空"; kinds = @("word"); expected = "" }
        @{ name = "すべて選んでいれば空"; kinds = @("excel", "word", "powerpoint", "text"); expected = "" }
    ) {
        param ($name, $kinds, $expected)
        getSearchKindError $kinds | Should -Be $expected
    }

    It "getNoKindMatchText: <name>" -TestCases @(
        @{ name = "絞っていれば種類を示す"; kinds = @("word", "text"); expected = "種類（Word・テキスト）に合うファイルがありません。" }
        @{ name = "絞っていなければ種類を示さない"; kinds = @(); expected = "検索できるファイルがありません。" }
    ) {
        param ($name, $kinds, $expected)
        getNoKindMatchText $kinds | Should -Be $expected
    }
}

Describe "getFastSearchView" -Tag Unit {
    It "Windows Search が使え（またはまだ確かめていない）、正規表現がオフで、2 文字以上の部分があれば使用可" {
        (getFastSearchView $true $false "見積").Text | Should -Be "高速検索：使用可"
        (getFastSearchView $null $false "見積").Usable | Should -Be $true
    }

    It "正規表現をオンにした・1 文字・Windows Search が使えないときは使用不可" {
        (getFastSearchView $true $true "見積").Text | Should -Be "高速検索：使用不可"
        (getFastSearchView $true $false "見").Usable | Should -Be $false
        (getFastSearchView $false $false "見積").Usable | Should -Be $false
    }

    It "使えない理由は、上から順に最初に当てはまるものをツールチップに出す" {
        (getFastSearchView $false $true "見積").Tip | Should -Be "検索はできますが時間がかかります　［インデックス管理で確認］"
        (getFastSearchView $false $false "見積").Tip | Should -Be "検索はできますが時間がかかります　［インデックス管理で確認］"
        (getFastSearchView $true $true "見積").Tip | Should -Be "正規表現をオフにすると速く検索できます"
    }

    It "使えるとき・ワードが短いだけのとき・まだ確かめていないときは、理由を出さない" {
        (getFastSearchView $true $false "見積").Tip | Should -Be ""
        (getFastSearchView $true $false "見").Tip | Should -Be ""
        (getFastSearchView $true $false "").Tip | Should -Be ""
        (getFastSearchView $null $false "見積").Usable | Should -Be $true
    }
}

Describe "getTargetCountText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "選んだ数 / 全部の数"; checked = 3; total = 4; expected = "検索対象 3 / 4" }
        @{ name = "1 つも選んでいなくても数を出す"; checked = 0; total = 4; expected = "検索対象 0 / 4" }
        @{ name = "インデックスが無ければ見出しだけ"; checked = 0; total = 0; expected = "検索対象" }
    ) {
        param ($name, $checked, $total, $expected)
        getTargetCountText $checked $total | Should -Be $expected
    }
}

Describe "getTargetHintText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "インデックスがあり、対象が無ければ案内する"; total = 4; targetCount = 0; expected = "検索するフォルダを選んでください" }
        @{ name = "対象があれば出さない"; total = 4; targetCount = 2; expected = "" }
        @{ name = "インデックスが無ければ出さない（別の案内が出る）"; total = 0; targetCount = 0; expected = "" }
    ) {
        param ($name, $total, $targetCount, $expected)
        getTargetHintText $total $targetCount | Should -Be $expected
    }
}

Describe "getScopeButtonText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "既定のまま（図形もコメントも検索）なら変更の数を付けない"; shapes = $true; comments = $true; expected = "ファイル内の対象"; changed = $false }
        @{ name = "どちらかを外したら 1 件"; shapes = $true; comments = $false; expected = "ファイル内の対象・1 件変更"; changed = $true }
        @{ name = "両方外したら 2 件"; shapes = $false; comments = $false; expected = "ファイル内の対象・2 件変更"; changed = $true }
    ) {
        param ($name, $shapes, $comments, $expected, $changed)
        $view = getScopeButtonText $shapes $comments
        $view.Text | Should -Be $expected
        $view.Changed | Should -Be $changed
    }
}

Describe "getWordNotice" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "正規表現でなければ出さない"; word = "("; useRegex = $false; expected = "" }
        @{ name = "正規表現として正しければ出さない"; word = "見積.*確定"; useRegex = $true; expected = "" }
        @{ name = "空のワードでは出さない"; word = ""; useRegex = $true; expected = "" }
        @{ name = "正規表現として不正なら、理由を書かずに正しくないと伝える"; word = "("; useRegex = $true; expected = "正規表現が正しくありません" }
    ) {
        param ($name, $word, $useRegex, $expected)
        getWordNotice $word $useRegex | Should -Be $expected
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
        $state.Content | Should -Be $content
        $state.Enabled | Should -Be $enabled
    }
}

Describe "getConditionFlow（検索条件の行の折り返し）" -Tag Unit {
    # 項目は 0 種類・1〜4 チップ・5 ファイル内の対象・（空き）・6 大文字小文字・7 正規表現・8 高速検索の印。幅は右の間 8 を含む
    BeforeAll {
        $script:widths = @(38, 70, 62, 100, 84, 184, 150, 70, 178)
    }

    It "広い幅では 1 行になり、空きが余りをすべて取る（右の組が右端に寄る）" {
        $flow = getConditionFlow $widths 6 1100
        $flow.Lines.Count | Should -Be 1
        $flow.SpacerWidth | Should -Be (1100 - 936 - 0.5)
    }

    It "狭くなると、高速検索の印だけが次の行の左端に落ち、ファイル内の対象は 1 行目に残る" {
        $flow = getConditionFlow $widths 6 780
        $flow.Lines.Count | Should -Be 2
        @($flow.Lines[0]) | Should -Be @(0, 1, 2, 3, 4, 5, 6, 7)
        @($flow.Lines[1]) | Should -Be @(8)
        $flow.LineStarts[8] | Should -BeTrue
        $flow.LineStarts[5] | Should -BeFalse
        # 空きは 1 行目の余りを取る
        $flow.SpacerWidth | Should -Be (780 - 758 - 0.5)
    }

    It "もっと狭くなると、後ろの項目から順に落ちる（高速検索の印・正規表現・大文字小文字の順）" {
        $flow = getConditionFlow $widths 6 600
        @($flow.Lines[0]) | Should -Be @(0, 1, 2, 3, 4, 5)
        @($flow.Lines[1]) | Should -Be @(6, 7, 8)
        $flow.LineStarts[6] | Should -BeTrue
        # 空きの前の最後の項目（ファイル内の対象）は 1 行目にあるので、空きは 1 行目の余りを取る
        $flow.SpacerWidth | Should -Be (600 - 538 - 0.5)
    }

    It "出していない項目（幅 0）は数えず、行の先頭にもならない" {
        $flow = getConditionFlow @(38, 70, 62, 100, 84, 184, 150, 70, 0) 6 780
        $flow.Lines.Count | Should -Be 1
        $flow.LineStarts[8] | Should -BeFalse
    }

    It "余りが 0.5 に満たないときの空きは 0" {
        $flow = getConditionFlow $widths 6 758.2
        $flow.SpacerWidth | Should -Be 0
    }
}
