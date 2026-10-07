# ボタンの幅の固定が、スタイルの MinWidth に負けないことのテスト。
# 普通のボタンのスタイル（Button.Base）は MinWidth を持ち、WPF は Width より MinWidth を優先する。
# 固定の幅（Width）を付けたボタンは、MinWidth も明示しないと、Width が効かず広いままになる
# （［検索］の Width="58" が MinWidth=88 に負けて 88 で出ていた）。
# XAML の要素の属性を見る（画面は開かない）。ボタンの Width は MinWidth 以上にそろえて書く。
BeforeAll {
    $script:xamlDir = (Resolve-Path "$PSScriptRoot\..\..\scripts").Path
}

Describe "ボタンの固定の幅" -Tag Meta {
    It "Width を付けたボタンは、それ以下の MinWidth を明示している" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in Get-ChildItem -LiteralPath $script:xamlDir -Recurse -Filter "*.xaml") {
            $text = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
            foreach ($m in [regex]::Matches($text, '<(Button|ToggleButton)[^>]*>')) {
                $tag = $m.Value
                if ($tag -notmatch '\sWidth="(\d+)"') { continue }
                $width = [int]$Matches[1]
                $min = if ($tag -match '\sMinWidth="(\d+)"') { [int]$Matches[1] } else { -1 }
                if ($min -lt 0 -or $min -gt $width) {
                    $bad.Add("$($f.Name)：$($tag.Substring(0, [Math]::Min(80, $tag.Length)))")
                }
            }
        }
        ($bad -join "`n") | Should -Be ""
    }
}
