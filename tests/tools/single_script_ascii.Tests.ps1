# tools\single_script_ascii.ps1（単一 .ps1 を ASCII だけの文字に書き換える関数）のテスト。
# 書き換え前後のスクリプトを実行して、結果が同じになること、書き換え後が ASCII だけで構文エラーが 0 であることを確かめる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "${scriptsDir}\..\tools\single_script_ascii.ps1"

    # スクリプトの文字列を実行して、出力を 1 つの文字列にする
    function invokeSnippet {
        param ([string]$code)
        $values = @(& ([scriptblock]::Create($code)))
        return ($values | ForEach-Object { [string]$_ }) -join "|"
    }

    function assertAsciiAndSame {
        param ([string]$code)
        $converted = convertToAsciiScript $code "テスト"
        [regex]::IsMatch($converted, '[^\x00-\x7F]') | Should -Be $false
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($converted, [ref]$null, [ref]$errors)
        $errors.Count | Should -Be 0
        (invokeSnippet $converted) | Should -Be (invokeSnippet $code)
        return $converted
    }
}

Describe "convertToAsciiScript（書き換えても、実行した結果が変わらない）" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "単一引用符の文字列"; code = "'あいう'" }
        @{ name = "単一引用符の文字列（引用符を含む・ASCII と混在）"; code = "'日本''語 a'" }
        @{ name = "単一引用符の文字列（全角の二重引用符）"; code = "'" + [char]0x201D + "'" }
        @{ name = "単一引用符の文字列（全角の一重引用符）"; code = "'" + [char]0x2019 + [char]0x2018 + "x'" }
        @{ name = "単一引用符の文字列（バックスラッシュを含む）"; code = "'C:\共有\営業部\n\d+あ\'" }
        @{ name = "サロゲートペア（𠮷）は 1 文字として残る"; code = "'𠮷野家'.Length" }
        @{ name = "二重引用符の文字列と式の展開"; code = "`"あ`$(1 + 1)い`"" }
        @{ name = "二重引用符の文字列と変数"; code = "`$x = 'z'; `"日本`${x}語`"" }
        @{ name = "二重引用符の文字列（全角の二重引用符を 2 つ重ねた 1 文字）"; code = "`"a" + [char]0x201D + [char]0x201D + "b`"" }
        @{ name = "二重引用符の文字列（バッククォートで日本語を逃がしている）"; code = "`"``日本`"" }
        @{ name = "式の中の単一引用符の日本語"; code = "`$a = 'x', 'y'; `"`$(`$a -join '、')`"" }
        @{ name = "式の中の二重引用符と添字"; code = "`$m = @{ 'k' = 'v' }; `"`$([string]`$m[`"k`"])`"" }
        @{ name = "ハッシュテーブルの引用符の無いキー"; code = "`$h = @{ 相対パス = 1; サイズ = 2 }; `$h['相対パス']; `$h['サイズ']" }
        @{ name = "メンバー名の日本語"; code = "`$o = [pscustomobject]@{ 名前 = 'a' }; `$o.名前" }
        @{ name = "単一引用符のヒアストリング"; code = "`$s = @'`r`nあ'い`r`nう`r`n'@; `$s" }
        @{ name = "二重引用符のヒアストリング"; code = "`$x = 1; `$s = @`"`r`n日本`$x`r`n`"@; `$s" }
        @{ name = "行のコメント"; code = "# 日本語のコメント`r`n1 + 1 # 末尾の日本語`r`n" }
        @{ name = "ブロックコメント"; code = "<# 日本語`r`n複数行 #>`r`n'a'" }
        @{ name = "関数の引数に渡す文字列"; code = "function f { param (`$x) `$x.Length }; f '日本語'; f -x ('語' + 'a')" }
        @{ name = "-split・-replace に渡す文字列"; code = "'あ、い、う' -split '、'; 'あいう' -replace 'い', 'イ'" }
    ) {
        param ($name, $code)
        $null = assertAsciiAndSame $code
    }

    It "日本語を含まないスクリプトは、そのまま返す" {
        $code = "`$a = 'abc'`r`n# comment`r`n"
        (convertToAsciiScript $code "テスト") | Should -Be $code
    }

    It "日本語のコメントだけを消し、ASCII のコメントは残す" {
        $converted = assertAsciiAndSame "# ascii comment`r`n# 日本語`r`n1"
        $converted | Should -Match "# ascii comment"
    }

    It "ブロックコメントは、行数が変わらないよう改行だけ残す" {
        $converted = convertToAsciiScript "1`r`n<# あ`r`nい`r`nう #>`r`n2" "テスト"
        ($converted -split "`r`n").Count | Should -Be 5
    }

    It "<name>は、書き換えられないので、ファイル名と行番号を示して止める" -TestCases @(
        @{ name = "日本語の変数名"; code = "`$a = 1`r`n`$変数 = 2" }
        @{ name = "引用符の無い日本語の引数"; code = "Write-Output 日本語" }
        @{ name = "文字列の中の日本語の変数名"; code = "`$変数 = 1; `"`$変数`"" }
    ) {
        param ($name, $code)
        { convertToAsciiScript $code "元.ps1" } | Should -Throw "*元.ps1*"
    }

    It "構文エラーのあるスクリプトは止める" {
        { convertToAsciiScript "'あ" "元.ps1" } | Should -Throw "*元.ps1*"
    }
}

Describe "convertToAsciiXaml（XAML の日本語を文字参照にする）" -Tag Unit {
    It "日本語は &#x....; にし、日本語のコメントは消す。XML としての内容は変わらない" {
        $xaml = "<Grid xmlns=`"http://example.com/x`">`r`n  <!-- 日本語のコメント -->`r`n  <TextBlock Text=`"検索 𠮷`" />`r`n  <!-- ascii -->`r`n</Grid>"
        $converted = convertToAsciiXaml $xaml "テスト.xaml"
        [regex]::IsMatch($converted, '[^\x00-\x7F]') | Should -Be $false
        $converted | Should -Match "&#x691C;"
        $converted | Should -Match "&#x20BB7;"
        $converted | Should -Not -Match "日本語のコメント"
        $converted | Should -Match "<!-- ascii -->"
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($converted)
        $doc.DocumentElement.ChildNodes.Where({ $_.Name -eq "TextBlock" })[0].GetAttribute("Text") | Should -Be "検索 𠮷"
    }

    It "日本語を含まない XAML は、そのまま返す" {
        $xaml = "<Grid />"
        (convertToAsciiXaml $xaml "テスト.xaml") | Should -Be $xaml
    }
}
