# テキストの開き方を選ぶ画面の文言と判断（tebunko\ui\settings\text_open_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\settings\text_open_view.ps1"
}

Describe "getTextOpenRows / getTextOpenOptions" -Tag Unit {
    It "固定の一覧と同じ 75 行を同じ順で返し、既定ではエディタで開くのは実行・登録になる拡張子だけ" {
        $rows = @(getTextOpenRows)
        $rows.Count | Should -Be 75
        (($rows | ForEach-Object { $_.Extension }) -join ",") | Should -Be (${textExtensions} -join ",")
        (@($rows | Where-Object { $_.Way -eq "editor" } | ForEach-Object { $_.Extension }) -join ",") | Should -Be (${textNotepadExtensions} -join ",")
    }

    It "実行・登録になる拡張子は選べる開き方が「エディタで開く」だけ、そのほかは 2 つ" {
        foreach ($row in @(getTextOpenRows)) {
            $values = @($row.Options | ForEach-Object { $_.Value })
            if (${textNotepadExtensions} -contains $row.Extension) {
                ($values -join ",") | Should -Be "editor" -Because $row.Extension
            } else {
                ($values -join ",") | Should -Be "default,editor" -Because $row.Extension
            }
            @($row.Options | Where-Object { $_.Text -eq "" }).Count | Should -Be 0
        }
    }

    It "<name>" -TestCases @(
        @{ name = "設定でエディタにした行はエディタ"; notepad = @(".html"); ext = ".html"; way = "editor" }
        @{ name = "設定に無い行は既定のアプリ"; notepad = @(".html"); ext = ".txt"; way = "default" }
        @{ name = "設定が大文字でも当てはまる"; notepad = @(".HTML"); ext = ".html"; way = "editor" }
        @{ name = "実行・登録になる拡張子は設定が空でもエディタ"; notepad = @(); ext = ".bat"; way = "editor" }
    ) {
        param ($name, $notepad, $ext, $way)
        $row = @(getTextOpenRows $notepad) | Where-Object { $_.Extension -eq $ext }
        $row.Way | Should -Be $way
    }
}

Describe "convertTextOpenRows" -Tag Unit {
    It "選んだ行から、設定に書く配列を作る。実行・登録になる拡張子（常にエディタ）は書かない" {
        $rows = @(
            @{ Extension = ".html"; Way = "editor" }
            @{ Extension = ".txt"; Way = "default" }
            @{ Extension = ".bat"; Way = "editor" }
            @{ Extension = ".css"; Way = "editor" }
        )
        (@(convertTextOpenRows $rows) -join ",") | Should -Be ".html,.css"
    }

    It "固定の一覧に無い拡張子・知らない開き方は捨てる" {
        @(convertTextOpenRows @(@{ Extension = ".exe"; Way = "editor" }, @{ Extension = ".txt"; Way = "ほか" })).Count | Should -Be 0
    }

    It "何も選ばなければ空" {
        @(convertTextOpenRows @()).Count | Should -Be 0
    }

    It "行を作って戻すと、設定の値と同じになる（往復）" {
        (@(convertTextOpenRows @(getTextOpenRows @(".html", ".css"))) -join ",") | Should -Be ".html,.css"
    }
}

Describe "getTextOpenSummary / getTextOpenChangeSummary" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "既定のまま"; notepad = @(); expected = "すべて既定のままです（エディタで開くのは、実行・登録になる 15 個）" }
        @{ name = "2 つをエディタにする"; notepad = @(".html", ".css"); expected = "変更したもの 2 個（エディタで開く）" }
        @{ name = "実行・登録になる拡張子を書いても変えたものに数えない"; notepad = @(".bat"); expected = "すべて既定のままです（エディタで開くのは、実行・登録になる 15 個）" }
    ) {
        param ($name, $notepad, $expected)
        getTextOpenSummary $notepad | Should -Be $expected
    }
}

Describe "getTextOpenListView" -Tag Unit {
    BeforeAll {
        # .csv・.html をエディタにした状態
        $script:rows = @(getTextOpenRows @(".csv", ".html"))
    }

    It "<name>" -TestCases @(
        @{ name = "絞らなければ 75 個すべてと、変更したものの要約"; filter = ""; changedOnly = $false; count = 75; summary = "変更したもの 2 個（エディタで開く）"; empty = $false; header = $true }
        @{ name = "［変更したものだけ］は変えた 2 個だけ"; filter = ""; changedOnly = $true; count = 2; summary = "変更したもの 2 個を表示"; empty = $false; header = $false }
        @{ name = "文字で絞る（先頭の . は有っても無くてもよい。大文字も同じ）"; filter = ".CS"; changedOnly = $false; count = 4; summary = "75 個中 4 個を表示"; empty = $false; header = $false }
        @{ name = "点なしでも同じ"; filter = "html"; changedOnly = $false; count = 1; summary = "75 個中 1 個を表示"; empty = $false; header = $false }
        @{ name = "合うものが無ければ 0 件で一致なし"; filter = "zzz"; changedOnly = $false; count = 0; summary = "75 個中 0 個を表示"; empty = $true; header = $false }
        @{ name = "実行・登録になる形式が見えるときは見出しを出す"; filter = "bat"; changedOnly = $false; count = 1; summary = "75 個中 1 個を表示"; empty = $false; header = $true }
        @{ name = "前後の空白は無視する"; filter = "  html  "; changedOnly = $false; count = 1; summary = "75 個中 1 個を表示"; empty = $false; header = $false }
    ) {
        param ($name, $filter, $changedOnly, $count, $summary, $empty, $header)
        $view = getTextOpenListView $script:rows $filter $changedOnly
        @($view.Visible).Count | Should -Be $count
        $view.Summary | Should -Be $summary
        $view.Empty | Should -Be $empty
        $view.ShowRestrictedHeader | Should -Be $header
    }

    It "見せる拡張子は固定の一覧の順" {
        $view = getTextOpenListView $script:rows "" $false
        ($view.Visible -join ",") | Should -Be (${textExtensions} -join ",")
    }
}

Describe "testTextEditorPath" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "ふつうの .exe"; path = "C:\Program Files\Notepad++\notepad++.exe"; expected = $true }
        @{ name = "拡張子の大文字"; path = "D:\tools\Editor.EXE"; expected = $true }
        @{ name = "空"; path = ""; expected = $false }
        @{ name = "null"; path = $null; expected = $false }
        @{ name = "文字列でない"; path = 5; expected = $false }
        @{ name = "相対パス"; path = "notepad.exe"; expected = $false }
        @{ name = "ドライブ文字だけ"; path = "C:\"; expected = $false }
        @{ name = "UNC"; path = "\\server\share\editor.exe"; expected = $false }
        @{ name = "環境変数"; path = "%SystemRoot%\System32\notepad.exe"; expected = $false }
        @{ name = "引数付き"; path = "C:\tools\editor.exe -n"; expected = $false }
        @{ name = "引用符つき"; path = '"C:\tools\editor.exe"'; expected = $false }
        @{ name = "引用符が途中にある"; path = 'C:\tools\"x".exe'; expected = $false }
        @{ name = "[ ] を含む（Start-Process がワイルドカードとして扱う）"; path = "C:\tools[1]\editor.exe"; expected = $false }
        @{ name = "* を含む"; path = "C:\tools\ed*.exe"; expected = $false }
        @{ name = "? を含む"; path = "C:\tools\ed?.exe"; expected = $false }
        @{ name = ".. を含む"; path = "C:\tools\..\editor.exe"; expected = $false }
        @{ name = ".exe でない"; path = "C:\tools\editor.bat"; expected = $false }
        @{ name = ".exe でない（.com）"; path = "C:\tools\editor.com"; expected = $false }
        @{ name = ".exe で終わらない"; path = "C:\tools\editor.exe.txt"; expected = $false }
    ) {
        param ($name, $path, $expected)
        testTextEditorPath $path | Should -Be $expected
    }
}

Describe "getTextEditorChoice / getTextEditorLabel / addTextEditorNotice" -Tag Unit {
    BeforeAll {
        $script:notepad = "C:\Windows\System32\notepad.exe"
    }

    It "<name>" -TestCases @(
        @{ name = "空はメモ帳（知らせなし）"; configured = ""; exists = $false; path = "C:\Windows\System32\notepad.exe"; notice = $false }
        @{ name = "null もメモ帳"; configured = $null; exists = $false; path = "C:\Windows\System32\notepad.exe"; notice = $false }
        @{ name = "形が正しく、ファイルがあれば、そのプログラム"; configured = "D:\tools\editor.exe"; exists = $true; path = "D:\tools\editor.exe"; notice = $false }
        @{ name = "ファイルが無ければメモ帳（知らせる）"; configured = "D:\tools\editor.exe"; exists = $false; path = "C:\Windows\System32\notepad.exe"; notice = $true }
        @{ name = "形が正しくなければ、ファイルがあってもメモ帳（知らせる）"; configured = "D:\tools\editor.bat"; exists = $true; path = "C:\Windows\System32\notepad.exe"; notice = $true }
        @{ name = "引数付きはメモ帳（知らせる）"; configured = "D:\tools\editor.exe /a"; exists = $true; path = "C:\Windows\System32\notepad.exe"; notice = $true }
    ) {
        param ($name, $configured, $exists, $path, $notice)
        $choice = getTextEditorChoice $configured $exists $script:notepad
        $choice.Path | Should -Be $path
        ($choice.Notice -ne "") | Should -Be $notice
        if ($notice) {
            $choice.Notice | Should -Be "設定のエディタ（textEditorPath）が使えないため、メモ帳で開きます。"
        }
    }

    It "エディタの名前は、設定が空・使えない値なら「メモ帳（既定）」、使える値ならファイル名" -TestCases @(
        @{ configured = ""; expected = "メモ帳（既定）" }
        @{ configured = "D:\tools\editor.bat"; expected = "メモ帳（既定）" }
        @{ configured = "D:\tools\MyEditor.exe"; expected = "MyEditor" }
    ) {
        param ($configured, $expected)
        getTextEditorLabel $configured | Should -Be $expected
    }

    It "ステータスの文に、知らせがあれば後ろに足す" {
        addTextEditorNotice "エディタで開きました：a.bat" "" | Should -Be "エディタで開きました：a.bat"
        addTextEditorNotice "エディタで開きました：a.bat" "使えません。" | Should -Be "エディタで開きました：a.bat 使えません。"
    }
}

Describe "getTextOpenWayText / getTextOpenDialogTexts" -Tag Unit {
    It "<way> の文言" -TestCases @(
        @{ way = "default"; expected = "既定のアプリで開く" }
        @{ way = "editor"; expected = "エディタで開く" }
        @{ way = "ほか"; expected = "" }
    ) {
        param ($way, $expected)
        getTextOpenWayText $way | Should -Be $expected
    }

    It "エディタの説明にエディタの名前が入る" {
        ((getTextOpenDialogTexts).EditorSettingsHint -f (getTextEditorLabel "")) | Should -Be "「エディタで開く」に使うアプリです。今はメモ帳（既定）だけです。"
    }

    It "ダイアログの文言がそろっている" {
        $texts = getTextOpenDialogTexts
        foreach ($key in @("Title", "Note", "Extension", "Way", "Reset", "Ok", "Cancel", "ChangedOnly", "FilterPlaceholder", "Empty", "RestrictedHeader", "RestrictedNote", "EditorTitle", "EditorHint", "FutureMark", "Change", "SettingsHint", "EditorSettingsHint")) {
            $texts[$key] | Should -Not -BeNullOrEmpty -Because $key
        }
    }
}
