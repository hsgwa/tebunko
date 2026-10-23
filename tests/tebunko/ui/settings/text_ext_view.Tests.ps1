# テキストファイルの拡張子の扱いを選ぶ画面の文言と判断（tebunko\ui\settings\text_ext_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\settings\text_ext_view.ps1"
}

Describe "getTextExtensionRows" -Tag Unit {
    It "固定の一覧と同じ 75 行を同じ順で返し、既定ではメモ帳で開く拡張子だけがメモ帳" {
        $rows = @(getTextExtensionRows)
        $rows.Count | Should -Be 75
        (($rows | ForEach-Object { $_.Extension }) -join ",") | Should -Be (${textExtensions} -join ",")
        @($rows | Where-Object { $_.Way -eq "exclude" }).Count | Should -Be 0
        (@($rows | Where-Object { $_.Way -eq "notepad" } | ForEach-Object { $_.Extension }) -join ",") | Should -Be (${textNotepadExtensions} -join ",")
    }

    It "実行・登録になる拡張子は選べる扱いが「メモ帳」「取り込まない」の 2 つだけ、そのほかは 3 つ" {
        foreach ($row in @(getTextExtensionRows)) {
            $values = @($row.Options | ForEach-Object { $_.Value })
            if (${textNotepadExtensions} -contains $row.Extension) {
                ($values -join ",") | Should -Be "notepad,exclude" -Because $row.Extension
            } else {
                ($values -join ",") | Should -Be "default,notepad,exclude" -Because $row.Extension
            }
            @($row.Options | Where-Object { $_.Text -eq "" }).Count | Should -Be 0
        }
    }

    It "<name>" -TestCases @(
        @{ name = "取り込まない設定の行は「取り込まない」"; excluded = @(".log"); notepad = @(); ext = ".log"; way = "exclude" }
        @{ name = "メモ帳の設定の行は「メモ帳」"; excluded = @(); notepad = @(".html"); ext = ".html"; way = "notepad" }
        @{ name = "取り込まないを優先する"; excluded = @(".html"); notepad = @(".html"); ext = ".html"; way = "exclude" }
        @{ name = "実行・登録になる拡張子を取り込まないにした行"; excluded = @(".bat"); notepad = @(); ext = ".bat"; way = "exclude" }
        @{ name = "設定が大文字でも当てはまる"; excluded = @(".LOG"); notepad = @(); ext = ".log"; way = "exclude" }
        @{ name = "設定に無い行は既定のアプリ"; excluded = @(".log"); notepad = @(".html"); ext = ".txt"; way = "default" }
    ) {
        param ($name, $excluded, $notepad, $ext, $way)
        $row = @(getTextExtensionRows $excluded $notepad) | Where-Object { $_.Extension -eq $ext }
        $row.Way | Should -Be $way
    }
}

Describe "convertTextExtensionRows" -Tag Unit {
    It "選んだ行から設定の 2 つの配列を作る。実行・登録になる拡張子の「メモ帳」は Notepad に書かない" {
        $rows = @(
            @{ Extension = ".log"; Way = "exclude" }
            @{ Extension = ".html"; Way = "notepad" }
            @{ Extension = ".txt"; Way = "default" }
            @{ Extension = ".bat"; Way = "notepad" }
            @{ Extension = ".ps1"; Way = "exclude" }
        )
        $result = convertTextExtensionRows $rows
        ($result.Excluded -join ",") | Should -Be ".log,.ps1"
        ($result.Notepad -join ",") | Should -Be ".html"
    }

    It "固定の一覧に無い拡張子・知らない扱いは捨てる" {
        $result = convertTextExtensionRows @(
            @{ Extension = ".exe"; Way = "exclude" }
            @{ Extension = ".txt"; Way = "ほか" }
        )
        @($result.Excluded).Count | Should -Be 0
        @($result.Notepad).Count | Should -Be 0
    }

    It "何も選ばなければ空" {
        $result = convertTextExtensionRows @()
        @($result.Excluded).Count | Should -Be 0
        @($result.Notepad).Count | Should -Be 0
    }

    It "行を作って戻すと、設定の値と同じになる（往復）" {
        $rows = @(getTextExtensionRows @(".log", ".bat") @(".html"))
        $result = convertTextExtensionRows $rows
        ($result.Excluded -join ",") | Should -Be ".log,.bat"
        ($result.Notepad -join ",") | Should -Be ".html"
    }
}

Describe "getTextExtensionSummary" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "既定のまま"; excluded = @(); notepad = @(); expected = "既定のままです" }
        @{ name = ".txt を取り込まず、.csv をメモ帳にする（メモ帳の数は 15 個＋1 個）"; excluded = @(".txt"); notepad = @(".csv"); expected = "取り込む 75 個のうち 74 個 ・ メモ帳で開く 16 個" }
        @{ name = "メモ帳にした拡張子を取り込まないにしたら、メモ帳の数から引く"; excluded = @(".bat"); notepad = @(); expected = "取り込む 75 個のうち 74 個 ・ メモ帳で開く 14 個" }
        @{ name = "取り込まないだけ"; excluded = @(".log"); notepad = @(); expected = "取り込む 75 個のうち 74 個 ・ メモ帳で開く 15 個" }
        @{ name = "メモ帳にするだけ"; excluded = @(); notepad = @(".html"); expected = "取り込む 75 個のうち 75 個 ・ メモ帳で開く 16 個" }
    ) {
        param ($name, $excluded, $notepad, $expected)
        getTextExtensionSummary $excluded $notepad | Should -Be $expected
    }
}

Describe "getTextExtensionWayText / getTextExtensionDialogTexts" -Tag Unit {
    It "<way> の文言" -TestCases @(
        @{ way = "default"; expected = "既定のアプリで開く" }
        @{ way = "notepad"; expected = "メモ帳で開く" }
        @{ way = "exclude"; expected = "取り込まない" }
        @{ way = "ほか"; expected = "" }
    ) {
        param ($way, $expected)
        getTextExtensionWayText $way | Should -Be $expected
    }

    It "ダイアログの文言がそろっている" {
        $texts = getTextExtensionDialogTexts
        foreach ($key in @("Title", "Note", "Extension", "Way", "Reset", "Ok", "Cancel")) {
            $texts[$key] | Should -Not -BeNullOrEmpty -Because $key
        }
        $texts.Note | Should -BeLike "*次にインデックス作成で件数を数えたときにインデックスから外します*"
    }
}
