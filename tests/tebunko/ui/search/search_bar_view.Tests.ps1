# 検索バーの判断（tebunko\ui\search\search_bar_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\search\search_bar_view.ps1"
}

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
        describeSearchOption $option | Should -Be $expected
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
