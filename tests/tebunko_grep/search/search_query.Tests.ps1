# 検索条件の組み立て（tebunko_grep\search\search_query.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "isValidRegex" -Tag Unit {
    It "正しい正規表現は true" {
        isValidRegex "見積.*確定" | Should Be $true
    }

    It "不正な正規表現は false" {
        isValidRegex "(" | Should Be $false
    }
}

Describe "newSearchRegex" -Tag Unit {
    It "文字どおりなら記号をそのまま探し、既定は大文字と小文字を区別しない" {
        $regex = (newSearchRegex "C++ (株)").Regex
        $regex.IsMatch("c++ (株)") | Should Be $true
        $regex.IsMatch("C (株)") | Should Be $false
    }

    It "正規表現として不正なワードは文字どおりにする" {
        $result = newSearchRegex "(" $false
        $result.SimpleMatch | Should Be $true
        $result.Regex.IsMatch("a(b") | Should Be $true
    }

    It "大文字と小文字を区別できる" {
        (newSearchRegex "ID" $true $true).Regex.IsMatch("社員id") | Should Be $false
        (newSearchRegex "ID" $true $true).Regex.IsMatch("社員ID") | Should Be $true
    }
}

Describe "getRegexScanMode" -Tag Unit {
    It "改行に一致しえず、行の外を見ない正規表現は lines" {
        foreach ($pattern in @([regex]::Escape("C++ (株) a`tb"), "見積.*確定", "^abc$", "[a-z0-9]+", "\d{3}-\w+", "(?:a|b)(?<n>c)\k<n>", "\bID\b", "[]a]")) {
            getRegexScanMode $pattern | Should Be "lines"
        }
    }

    It "改行に一致しうる・行の外を見る正規表現は filter" {
        foreach ($pattern in @("見積\s確定", "[^,]+", "\W", "\x0A", "\p{L}", "a(?=b)", "(?<=a)b", "[\t-z]", "(a)\1")) {
            getRegexScanMode $pattern | Should Be "filter"
        }
    }

    It "全文では 1 行と結果が変わりうる正規表現は scan" {
        foreach ($pattern in @("\Aabc", "abc\z", "abc\Z", "\Gabc", "a(?!b)", "(?<!a)b", "(?i)abc", "(?s)a.b", "(?(a)b|c)")) {
            getRegexScanMode $pattern | Should Be "scan"
        }
    }

    It "文字どおりのワードの改行は lines にしない" {
        getRegexScanMode ([regex]::Escape("a`nb")) | Should Be "filter"
    }
}

Describe "newPlaceExclude" -Tag Unit {
    It "どちらも検索するなら `$null" {
        newPlaceExclude $true $true | Should Be $null
    }

    It "外す種類の場所（名前の末尾）だけに一致する" {
        $shapes = newPlaceExclude $false $true
        $shapes.IsMatch("売上[図形]") | Should Be $true
        $shapes.IsMatch("売上[コメント]") | Should Be $false
        $shapes.IsMatch("売上") | Should Be $false
        $both = newPlaceExclude $false $false
        $both.IsMatch("売上[図形]") | Should Be $true
        $both.IsMatch("売上[コメント]") | Should Be $true
        $both.IsMatch("ページ001") | Should Be $false
    }
}

Describe "newFileFilter" -Tag Unit {
    It "; で区切ったワイルドカードで含め、! で始まるもので除く" {
        $filter = newFileFilter "*.xlsx；見積 ; !*old*"
        $filter.Include.IsMatch("A社.XLSX") | Should Be $true
        $filter.Include.IsMatch("2024見積書.docx") | Should Be $true
        $filter.Include.IsMatch("報告書.docx") | Should Be $false
        $filter.Exclude.IsMatch("A社_old.xlsx") | Should Be $true
    }

    It "空なら条件なし" {
        $filter = newFileFilter "  "
        $filter.Include | Should Be $null
        $filter.Exclude | Should Be $null
    }

    It "? は任意の1文字、ほかの記号は文字どおり" {
        $filter = newFileFilter "v?.[確定].xlsx"
        $filter.Include.IsMatch("v1.[確定].xlsx") | Should Be $true
        $filter.Include.IsMatch("v1x[確定].xlsx") | Should Be $false
    }
}
