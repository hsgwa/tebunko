# 検索バーの判断（tebunko\ui\search\search_bar_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\search\search_bar_view.ps1"
}

Describe "describeSearchOption" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "既定のままなら空"; option = @{ CaseSensitive = $false; FileKinds = @() }; expected = "" }
        @{ name = "大文字と小文字の区別を出す"; option = @{ CaseSensitive = $true; FileKinds = @() }; expected = "大文字と小文字を区別" }
        @{ name = "種類を絞っていれば出す"; option = @{ CaseSensitive = $false; FileKinds = @("excel", "text") }; expected = "種類：Excel・テキスト" }
        @{ name = "すべての種類なら出さない"; option = @{ CaseSensitive = $false; FileKinds = @("excel", "word", "powerpoint", "text") }; expected = "" }
        @{ name = "両方あれば中黒でつなぐ"; option = @{ CaseSensitive = $true; FileKinds = @("word") }; expected = "大文字と小文字を区別・種類：Word" }
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
        @{ name = "最後の 1 つは外せない"; kinds = @("excel"); kind = "excel"; expected = "excel" }
        @{ name = "空はすべて選んでいるものとして扱う"; kinds = @(); kind = "text"; expected = "excel,word,powerpoint" }
        @{ name = "知らない種類は変えない"; kinds = @("excel", "text"); kind = "pdf"; expected = "excel,text" }
    ) {
        param ($name, $kinds, $kind, $expected)
        (toggleFileKind $kinds $kind) -join "," | Should -Be $expected
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
}

Describe "getNoIndexTargetText" -Tag Unit {
    It "ナビの名前（［インデックス管理］）で案内し、古い呼び名を使わない" {
        $text = getNoIndexTargetText
        $text | Should -Be "検索対象：なし（インデックスがありません。先に［インデックス管理］で作成してください）"
        $text | Should -Not -Match "［1 "
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
