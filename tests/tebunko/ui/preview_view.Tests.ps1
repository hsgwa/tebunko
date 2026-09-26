# プレビューの判断（tebunko\ui\preview_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\preview_view.ps1"
}

Describe "getPreviewRowCounts" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "高さに入る行数を前後に割り振る（余りは後ろ）"; height = 220; before = 4; after = 5 }
        @{ name = "前後同数になる高さでは同じ数にする"; height = 242; before = 5; after = 5 }
        @{ name = "高さが足りなくても選択行の 1 行は出す"; height = 0; before = 0; after = 0 }
        @{ name = "上限を超えて読まない"; height = 100000; before = 50; after = 50 }
    ) {
        param ($name, $height, $before, $after)
        $counts = getPreviewRowCounts $height 22 101
        $counts[0] | Should -Be $before
        $counts[1] | Should -Be $after
    }
}

Describe "toStatusText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "短い文字列はそのまま"; text = "見積書"; expected = "見積書" }
        @{ name = "改行とタブはスペースにする"; text = "A`r`nB`tC"; expected = "A B C" }
        @{ name = "40 文字を超えたら末尾を省略する"; text = ("あ" * 50); expected = (("あ" * 40) + "…") }
        @{ name = "ちょうど 40 文字なら省略しない"; text = ("あ" * 40); expected = ("あ" * 40) }
    ) {
        param ($name, $text, $expected)
        toStatusText $text | Should -Be $expected
    }
}
