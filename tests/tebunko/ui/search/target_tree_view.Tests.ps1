# 検索対象のツリーの判断（tebunko\ui\search\target_tree_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\search\target_tree_view.ps1"
}

Describe "matchesTreeFilter" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "空ならすべて合う"; text = "経理資料"; filter = ""; expected = $true }
        @{ name = "空白だけでもすべて合う"; text = "経理資料"; filter = "  　"; expected = $true }
        @{ name = "名前の一部に合う"; text = "経理資料"; filter = "理資"; expected = $true }
        @{ name = "大文字・小文字を区別しない"; text = "Sales2025"; filter = "sales"; expected = $true }
        @{ name = "空白で区切った語は、すべてを含むときだけ合う"; text = "営業 2025 見積"; filter = "見積 営業"; expected = $true }
        @{ name = "含まない語があれば合わない"; text = "営業 2025 見積"; filter = "営業 経理"; expected = $false }
        @{ name = "合わない"; text = "経理資料"; filter = "営業"; expected = $false }
    ) {
        param ($name, $text, $filter, $expected)
        matchesTreeFilter $text $filter | Should -Be $expected
    }
}
