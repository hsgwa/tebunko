# 検索条件の組み立て（tebunko\search\search_query.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "newSearchRegex" -Tag Unit {
    It "文字どおりなら記号をそのまま探し、既定は大文字と小文字を区別しない" {
        $regex = (newSearchRegex "C++ (株)").Regex
        $regex.IsMatch("c++ (株)") | Should -Be $true
        $regex.IsMatch("C (株)") | Should -Be $false
    }

    It "正規表現として不正なワードは文字どおりにする" {
        $result = newSearchRegex "(" $false
        $result.SimpleMatch | Should -Be $true
        $result.Regex.IsMatch("a(b") | Should -Be $true
    }

    It "大文字と小文字を区別できる" {
        (newSearchRegex "ID" $true $true).Regex.IsMatch("社員id") | Should -Be $false
        (newSearchRegex "ID" $true $true).Regex.IsMatch("社員ID") | Should -Be $true
    }
}

Describe "getRegexScanMode" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "改行に一致しえず、行の外を見ない正規表現は lines"; patterns = @([regex]::Escape("C++ (株) a`tb"), "見積.*確定", "^abc$", "[a-z0-9]+", "\d{3}-\w+", "(?:a|b)(?<n>c)\k<n>", "\bID\b", "[]a]"); expected = "lines" }
        @{ name = "改行に一致しうる・行の外を見る正規表現は filter"; patterns = @("見積\s確定", "[^,]+", "\W", "\x0A", "\p{L}", "a(?=b)", "(?<=a)b", "[\t-z]", "(a)\1"); expected = "filter" }
        @{ name = "全文では 1 行と結果が変わりうる正規表現は scan"; patterns = @("\Aabc", "abc\z", "abc\Z", "\Gabc", "a(?!b)", "(?<!a)b", "(?i)abc", "(?s)a.b", "(?(a)b|c)"); expected = "scan" }
        @{ name = "文字どおりのワードの改行は lines にしない"; patterns = @([regex]::Escape("a`nb")); expected = "filter" }
        @{ name = "改行・タブの文字をそのまま含む正規表現は filter"; patterns = @("a`tb", "[`n]"); expected = "filter" }
        @{ name = "末尾が \ だけで終わる（書きかけの）正規表現は scan（安全側）"; patterns = @("abc\"); expected = "scan" }
        @{ name = "アトミックグループ・' で囲む名前付きグループ・改行に一致しない \S は lines"; patterns = @("(?>ab+)c", "(?'n'a)\k'n'", "\S+", "[\]\-]"); expected = "lines" }
        @{ name = "コメント・インラインのオプションは scan"; patterns = @("a(?#コメント)b", "(?m)^a", "(?x) a b"); expected = "scan" }
        @{ name = "空の正規表現は lines（どの行にも一致する）"; patterns = @(""); expected = "lines" }
    ) {
        param ($name, $patterns, $expected)
        foreach ($pattern in $patterns) {
            getRegexScanMode $pattern | Should -Be $expected
        }
    }
}

Describe "newPlaceExclude" -Tag Unit {
    It "どちらも検索するなら `$null" {
        newPlaceExclude $true $true | Should -Be $null
    }

    It "外す種類の場所（名前の末尾）だけに一致する" {
        $shapes = newPlaceExclude $false $true
        $shapes.IsMatch("売上[図形]") | Should -Be $true
        $shapes.IsMatch("売上[コメント]") | Should -Be $false
        $shapes.IsMatch("売上") | Should -Be $false
        $both = newPlaceExclude $false $false
        $both.IsMatch("売上[図形]") | Should -Be $true
        $both.IsMatch("売上[コメント]") | Should -Be $true
        $both.IsMatch("ページ001") | Should -Be $false
    }
}

Describe "newFileKindFilter" -Tag Unit {
    # kinds: 選んだ種類、matches / misses: 作った条件に当たる・当たらない名前
    It "<name>" -TestCases @(
        @{ name = "すべての種類なら絞り込まない"; kinds = @("excel", "word", "powerpoint", "text"); matches = @(); misses = @(); empty = $true }
        @{ name = "空なら絞り込まない"; kinds = @(); matches = @(); misses = @(); empty = $true }
        @{ name = "知らない種類だけなら絞り込まない"; kinds = @("pdf"); matches = @(); misses = @(); empty = $true }
        @{ name = "excel は Excel の拡張子だけ"; kinds = @("excel"); empty = $false
           matches = @("a.xlsx", "a.XLSM", "a.xls", "a.xlsb"); misses = @("a.docx", "a.pptx", "a.txt", "a.xlsx.bak") }
        @{ name = "word は Word の拡張子だけ"; kinds = @("word"); empty = $false
           matches = @("a.docx", "a.docm", "a.doc"); misses = @("a.xlsx", "a.ppt", "a.md") }
        @{ name = "powerpoint は PowerPoint の拡張子だけ"; kinds = @("powerpoint"); empty = $false
           matches = @("a.pptx", "a.pptm", "a.ppt"); misses = @("a.docx", "a.xls", "a.csv") }
        @{ name = "text は拡張子の一覧のもの"; kinds = @("text"); empty = $false
           matches = @("a.txt", "a.CSV", "a.ps1"); misses = @("a.xlsx", "a.docx", "a.pptx", "a.pdf") }
        @{ name = "複数の種類は和集合"; kinds = @("excel", "text"); empty = $false
           matches = @("a.xlsx", "a.txt"); misses = @("a.docx", "a.ppt") }
    ) {
        param ($name, $kinds, $matches, $misses, $empty)
        $text = newFileKindFilter $kinds
        if ($empty) {
            $text | Should -Be ""
            return
        }
        $filter = newFileFilter $text
        $filter.Exclude | Should -Be $null
        foreach ($file in $matches) {
            $filter.Include.IsMatch($file) | Should -Be $true -Because $file
        }
        foreach ($file in $misses) {
            $filter.Include.IsMatch($file) | Should -Be $false -Because $file
        }
    }
}

Describe "newFileFilter" -Tag Unit {
    It "; で区切ったワイルドカードで含め、! で始まるもので除く" {
        $filter = newFileFilter "*.xlsx；見積 ; !*old*"
        $filter.Include.IsMatch("A社.XLSX") | Should -Be $true
        $filter.Include.IsMatch("2024見積書.docx") | Should -Be $true
        $filter.Include.IsMatch("報告書.docx") | Should -Be $false
        $filter.Exclude.IsMatch("A社_old.xlsx") | Should -Be $true
    }

    It "空なら条件なし" {
        $filter = newFileFilter "  "
        $filter.Include | Should -Be $null
        $filter.Exclude | Should -Be $null
    }

    It "? は任意の1文字、ほかの記号は文字どおり" {
        $filter = newFileFilter "v?.[確定].xlsx"
        $filter.Include.IsMatch("v1.[確定].xlsx") | Should -Be $true
        $filter.Include.IsMatch("v1x[確定].xlsx") | Should -Be $false
    }

    It "除外だけなら Include は `$null（除外に当たらないものはすべて対象）" {
        $filter = newFileFilter "！*old*"
        $filter.Include | Should -Be $null
        $filter.Exclude.IsMatch("A社_OLD.xlsx") | Should -Be $true
    }

    It "! や ; だけ・空の項目は無視する" {
        $filter = newFileFilter "!;;； ; ! "
        $filter.Include | Should -Be $null
        $filter.Exclude | Should -Be $null
    }

    It "ワイルドカードは名前全体に一致させる（*.xlsx は .xlsx.bak に当たらない）" {
        $filter = newFileFilter "*.xlsx"
        $filter.Include.IsMatch("a.xlsx") | Should -Be $true
        $filter.Include.IsMatch("a.xlsx.bak") | Should -Be $false
    }

    It "正規表現の記号（( ) + ^ $ など）を含む名前も文字どおりに部分一致させる" {
        $filter = newFileFilter "(1)+`$^"
        $filter.Include.IsMatch("見積(1)+`$^版.xlsx") | Should -Be $true
        $filter.Include.IsMatch("見積1.xlsx") | Should -Be $false
    }
}

Describe "truncateHitLine" -Tag Unit {
    BeforeAll {
        # 一致の位置を testIndex 文字目とする、長さ 3000 の行（先頭からの通し番号を並べて作り、切った範囲が分かるようにする）
        $longLine = (0..2999 | ForEach-Object { [string]([char](0x30 + ($_ % 10))) }) -join ""
    }

    It "hitLineMaxChars（1000文字）以下ならそのまま返す" {
        $line = "a" * 1000
        truncateHitLine $line 500 | Should -Be $line
        (truncateHitLine $line -1) | Should -Be $line
    }

    It "一致の位置が分からない（-1）ときは先頭から1000文字取り、末尾に … を付ける" {
        $result = truncateHitLine $longLine -1
        $result.Length | Should -Be 1001
        $result.Substring(0, 1000) | Should -Be $longLine.Substring(0, 1000)
        $result.Substring(1000) | Should -Be "…"
    }

    It "一致の位置が先頭付近なら、先頭から1000文字（先頭に … は付かず、末尾に付く）" {
        $result = truncateHitLine $longLine 10
        $result.Substring(0, 1000) | Should -Be $longLine.Substring(0, 1000)
        $result.EndsWith("…") | Should -Be $true
        $result.StartsWith("…") | Should -Be $false
    }

    It "一致の位置が中ほどなら、200文字前から1000文字取り、両端に … を付ける" {
        $result = truncateHitLine $longLine 1500
        $result.Length | Should -Be 1002
        $result.Substring(0, 1) | Should -Be "…"
        $result.Substring(1, 1000) | Should -Be $longLine.Substring(1300, 1000)
        $result.Substring(1001, 1) | Should -Be "…"
    }

    It "一致の位置が末尾付近なら、切り取った範囲が行の末尾に届き、末尾には … が付かない（届いた範囲は1000文字に満たないことがある）" {
        $result = truncateHitLine $longLine 2990
        $start = 2990 - 200  # 2790（200文字前）
        $expectedCore = $longLine.Substring($start, $longLine.Length - $start)  # 末尾まで（210文字）
        $result | Should -Be ("…" + $expectedCore)
    }
}
