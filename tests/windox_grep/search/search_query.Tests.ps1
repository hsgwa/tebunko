# 検索条件の組み立て（windox_grep\search\search_query.ps1）のテスト
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
