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
        @{ name = "すべての種類なら出さない"; option = @{ CaseSensitive = $false; FileKinds = @("excel", "word", "powerpoint", "text", "pdf") }; expected = "" }
        @{ name = "PDF だけ外した種類は出す"; option = @{ CaseSensitive = $false; FileKinds = @("excel", "word", "powerpoint", "text") }; expected = "種類：Excel・Word・PowerPoint・テキスト" }
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
        @{ kind = "pdf"; expected = "PDF" }
        @{ kind = "zip"; expected = "" }
    ) {
        param ($kind, $expected)
        getFileKindLabel $kind | Should -Be $expected
    }

    It "toggleFileKind: <name>" -TestCases @(
        @{ name = "外す"; kinds = @("excel", "word", "powerpoint", "text", "pdf"); kind = "word"; expected = "excel,powerpoint,text,pdf" }
        @{ name = "pdf を足す（最後に並ぶ）"; kinds = @("text", "excel"); kind = "pdf"; expected = "excel,text,pdf" }
        @{ name = "足す（決まった順に並ぶ）"; kinds = @("text", "excel"); kind = "word"; expected = "excel,word,text" }
        @{ name = "最後の 1 つも外せる（1 つも選ばない）"; kinds = @("excel"); kind = "excel"; expected = "" }
        @{ name = "空（1 つも選んでいない）から 1 つ選べる"; kinds = @(); kind = "text"; expected = "text" }
        @{ name = "知らない種類は変えない"; kinds = @("excel", "text"); kind = "zip"; expected = "excel,text" }
    ) {
        param ($name, $kinds, $kind, $expected)
        (toggleFileKind $kinds $kind) -join "," | Should -Be $expected
    }

    It "getSearchKindBalloonText: <name>" -TestCases @(
        @{ name = "1 つも選んでいなければ吹き出しを出す"; kinds = @(); show = $true }
        @{ name = "知らない種類だけでも 1 つも選んでいない"; kinds = @("zip"); show = $true }
        @{ name = "1 つ選んでいれば出さない"; kinds = @("word"); show = $false }
        @{ name = "すべて選んでいれば出さない"; kinds = @("excel", "word", "powerpoint", "text", "pdf"); show = $false }
    ) {
        param ($name, $kinds, $show)
        $balloon = getSearchKindBalloonText $kinds
        if ($show) {
            $balloon | Should -Be "種類を 1 つ以上選んでください"
        } else {
            $balloon | Should -BeNullOrEmpty
        }
    }

    It "getNoKindMatchText: <name>" -TestCases @(
        @{ name = "絞っていれば種類を示す"; kinds = @("word", "text"); expected = "種類（Word・テキスト）に合うファイルがありません。" }
        @{ name = "絞っていなければ種類を示さない"; kinds = @(); expected = "検索できるファイルがありません。" }
    ) {
        param ($name, $kinds, $expected)
        getNoKindMatchText $kinds | Should -Be $expected
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

Describe "getBalloonLeft" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "収まるときは基準の左端にそろえる"; anchor = 100; width = 200; available = 800; expected = 100 }
        @{ name = "右端に収まる限界の位置はそのまま"; anchor = 592; width = 200; available = 800; expected = 592 }
        @{ name = "右端からはみ出すときは右に間 8 を残して左へ寄せる"; anchor = 700; width = 200; available = 800; expected = 592 }
        @{ name = "幅より広くても左は 0 より左へ出さない"; anchor = 50; width = 900; available = 800; expected = 0 }
        @{ name = "間を変えられる"; anchor = 700; width = 200; available = 800; margin = 0; expected = 600 }
    ) {
        param ($name, $anchor, $width, $available, $expected, $margin = 8)
        getBalloonLeft $anchor $width $available $margin | Should -Be $expected
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
        @{ name = "正規表現が正しくなければ押せない"; searching = $false; stopping = $false; word = "("; hasIndex = $true; targetCount = 2; content = "検索"; enabled = $false; invalid = $true }
        @{ name = "検索中は正規表現の誤りに関わらず［中止］を押せる"; searching = $true; stopping = $false; word = "("; hasIndex = $true; targetCount = 2; content = "中止"; enabled = $true; invalid = $true }
        @{ name = "検索対象のツリーを読み込んでいる間は押せない"; searching = $false; stopping = $false; word = "見積"; hasIndex = $true; targetCount = 2; content = "検索"; enabled = $false; loading = $true }
    ) {
        param ($name, $searching, $stopping, $word, $hasIndex, $targetCount, $content, $enabled, $invalid = $false, $loading = $false)
        $state = newSearchButtonState $searching $stopping $word $hasIndex $targetCount $invalid $loading
        $state.Content | Should -Be $content
        $state.Enabled | Should -Be $enabled
    }
}

Describe "ワークスペースの読み込みの文言" -Tag Unit {
    It "接続できないときの知らせにワークスペースを入れる" {
        getWorkspaceUnreachableText "\\fileserver\共有\ws" | Should -Be "ワークスペースに接続できません：\\fileserver\共有\ws"
    }

    It "確かめている間のステータス" {
        getWorkspaceCheckingText | Should -Be "ワークスペースを確かめています…"
    }

    It "フォルダを読み込めなかったときのステータスにフォルダと理由を入れる" {
        getTreeFolderFailedText "営業部" "アクセスできません" | Should -Be "フォルダを読み込めませんでした：営業部（アクセスできません）"
    }
}

Describe "getSearchTargetText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "読み込み中は、ほかの状態より先に読み込み中と出す"; loading = $true; error = "\\fileserver\共有\ws"; total = 0; count = 0; all = $true; summary = $null; expected = "検索対象：読み込んでいます…" }
        @{ name = "接続できないときは、ワークスペースを入れて知らせる"; loading = $false; error = "\\fileserver\共有\ws"; total = 0; count = 0; all = $true; summary = $null; expected = "検索対象：なし（ワークスペースに接続できません：\\fileserver\共有\ws）" }
        @{ name = "インデックスが無ければ案内を出す"; loading = $false; error = ""; total = 0; count = 0; all = $true; summary = $null; expected = "［インデックス管理］で作ったインデックスの一覧です" }
        @{ name = "チェックが 1 つも無ければ「なし」"; loading = $false; error = ""; total = 2; count = 0; all = $false; summary = $null; expected = "検索対象：なし" }
        @{ name = "一部だけなら選んだ対象を出す"; loading = $false; error = ""; total = 2; count = 1; all = $false; summary = $null; targets = "営業部"; expected = "検索対象：営業部" }
        @{ name = "すべてで、集計がまだなら確認中"; loading = $false; error = ""; total = 2; count = 2; all = $true; summary = $null; expected = "検索対象：すべて（確認中…）" }
    ) {
        param ($name, $loading, $error, $total, $count, $all, $summary, $expected, $targets = "")
        getSearchTargetText $loading $error $total $count $all $targets $summary | Should -Be $expected
    }

    It "すべてで、集計があれば件数と最終更新を出す" {
        $summary = @{ Count = 12345; LastWrite = [datetime]"2026-09-20 10:30" }

        getSearchTargetText $false "" 2 2 $true "" $summary | Should -Be "検索対象：すべて（集約ファイル 12,345 件 ・ 最終更新 9/20 10:30）"
    }
}

Describe "getOldIndexNoticeText と古い形式の 1 行" -Tag Unit {
    It "古い形式のインデックスが <count> 件のとき、1 行は「<expected>」" -TestCases @(
        @{ count = 0; expected = "" }
        @{ count = 1; expected = "古い形式のインデックスは検索していません" }
        @{ count = 5; expected = "古い形式のインデックスは検索していません" }
    ) {
        param ($count, $expected)
        getOldIndexNoticeText $count | Should -Be $expected
    }

    It "検索対象の説明の最後に、古い形式の 1 行を足す。読み込み中・接続できないときは足さない" {
        getSearchTargetText $false "" 2 1 $false "営業部" $null 1 | Should -Be "検索対象：営業部`r`n古い形式のインデックスは検索していません"
        getSearchTargetText $false "" 0 0 $true "" $null 1 | Should -Be "［インデックス管理］で作ったインデックスの一覧です`r`n古い形式のインデックスは検索していません"
        getSearchTargetText $true "" 2 1 $false "営業部" $null 1 | Should -Be "検索対象：読み込んでいます…"
        getSearchTargetText $false "" 2 1 $false "営業部" $null 0 | Should -Be "検索対象：営業部"
    }
}

Describe "getConditionFlow（検索条件の行の折り返し）" -Tag Unit {
    # 項目は 0 種類・1〜4 チップ・5 ファイル内の対象・（空き）・6 大文字小文字・7 正規表現。幅は右の間 8 を含む
    BeforeAll {
        $script:widths = @(38, 70, 62, 100, 84, 184, 150, 70)
    }

    # 行は項目の番号を , でつなぎ、行の区切りを / で書く。758 は全部の幅の合計（ちょうど収まる）
    It "<Case>: 使える幅 <Available> のとき <Rows>（空き <Spacer>）" -TestCases @(
        @{ Case = "広い幅"; Available = 1100; Rows = "0,1,2,3,4,5,6,7"; Spacer = (1100 - 758 - 0.5) }
        @{ Case = "ちょうど収まる幅"; Available = 758; Rows = "0,1,2,3,4,5,6,7"; Spacer = 0 }
        @{ Case = "1 足りない幅（正規表現だけが落ちる）"; Available = 757; Rows = "0,1,2,3,4,5,6/7"; Spacer = (757 - 688 - 0.5) }
        @{ Case = "狭い幅（後ろから、正規表現・大文字小文字の順に落ち、ファイル内の対象は残る）"; Available = 600; Rows = "0,1,2,3,4,5/6,7"; Spacer = (600 - 538 - 0.5) }
    ) {
        param ($Case, $Available, $Rows, $Spacer)
        $flow = getConditionFlow $widths 6 $Available
        (@($flow.Lines | ForEach-Object { @($_) -join "," }) -join "/") | Should -Be $Rows
        $flow.SpacerWidth | Should -Be $Spacer
        # 落ちた行の先頭は行の左端（LineStarts）
        $first = @($flow.Lines | ForEach-Object { @($_)[0] })
        for ($i = 0; $i -lt 8; $i++) { $flow.LineStarts[$i] | Should -Be ($first -contains $i) }
    }

    It "出していない項目（幅 0）は数えず、行の先頭にもならない" {
        $flow = getConditionFlow @(38, 70, 62, 100, 84, 184, 150, 0) 6 780
        $flow.Lines.Count | Should -Be 1
        $flow.LineStarts[7] | Should -BeFalse
    }

    It "余りが 0.5 に満たないときの空きは 0" {
        $flow = getConditionFlow $widths 6 758.2
        $flow.SpacerWidth | Should -Be 0
    }
}
