# プレビューの判断（win_grep\ui\preview_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\win_grep\ui\preview_view.ps1"

Describe "getPreviewRowCounts" -Tag Unit {
    It "高さに入る行数を前後に割り振る（余りは後ろ）" {
        $counts = getPreviewRowCounts 220 22 101
        $counts[0] | Should Be 4
        $counts[1] | Should Be 5
    }

    It "前後同数になる高さでは同じ数にする" {
        $counts = getPreviewRowCounts 242 22 101
        $counts[0] | Should Be 5
        $counts[1] | Should Be 5
    }

    It "高さが足りなくても選択行の 1 行は出す" {
        $counts = getPreviewRowCounts 0 22 101
        $counts[0] | Should Be 0
        $counts[1] | Should Be 0
    }

    It "高さが負でも落ちない" {
        $counts = getPreviewRowCounts -50 22 101
        $counts[0] | Should Be 0
        $counts[1] | Should Be 0
    }

    It "上限を超えて読まない" {
        $counts = getPreviewRowCounts 100000 22 101
        ($counts[0] + $counts[1] + 1) | Should Be 101
    }
}

Describe "toStatusText" -Tag Unit {
    It "短い文字列はそのまま" {
        toStatusText "見積書" | Should Be "見積書"
    }

    It "改行とタブはスペースにする" {
        toStatusText "A`r`nB`tC" | Should Be "A B C"
    }

    It "40 文字を超えたら末尾を省略する" {
        $text = toStatusText ("あ" * 50)
        $text.Length | Should Be 41
        $text.EndsWith("…") | Should Be $true
    }

    It "ちょうど 40 文字なら省略しない" {
        toStatusText ("あ" * 40) | Should Be ("あ" * 40)
    }
}