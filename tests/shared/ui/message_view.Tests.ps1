# メッセージの画面の種類・ボタン・文の分け方（scripts/shared/ui/message_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\shared\ui\message_view.ps1"
}

Describe "getMessageLook" -Tag Unit {
    It "<icon> は <kind>（アイコン <glyph>・色 <brush>）" -TestCases @(
        @{ icon = "Information"; kind = "Info";     glyph = "Icon.Info";               brush = "Accent" }
        @{ icon = "Asterisk";    kind = "Info";     glyph = "Icon.Info";               brush = "Accent" }
        @{ icon = "Warning";     kind = "Warning";  glyph = "Icon.TriangleAlert";      brush = "Warn" }
        @{ icon = "Exclamation"; kind = "Warning";  glyph = "Icon.TriangleAlert";      brush = "Warn" }
        @{ icon = "Error";       kind = "Error";    glyph = "Icon.CircleAlert";        brush = "Danger.Text" }
        @{ icon = "Stop";        kind = "Error";    glyph = "Icon.CircleAlert";        brush = "Danger.Text" }
        @{ icon = "Question";    kind = "Question"; glyph = "Icon.CircleQuestionMark"; brush = "Accent" }
        @{ icon = "None";        kind = "None";     glyph = "";                        brush = "" }
        @{ icon = "";            kind = "None";     glyph = "";                        brush = "" }
    ) {
        $look = getMessageLook $icon
        $look.Kind | Should -Be $kind
        $look.Icon | Should -Be $glyph
        $look.Brush | Should -Be $brush
    }

    It "アイコンの名前は theme.xaml にあり、色の名前も theme.xaml にある" {
        $theme = Get-Content -LiteralPath "${scriptsDir}\shared\xaml\theme.xaml" -Raw -Encoding UTF8
        foreach ($icon in "Information", "Warning", "Error", "Question") {
            $look = getMessageLook $icon
            $theme | Should -Match ('x:Key="' + [regex]::Escape($look.Icon) + '"')
            $theme | Should -Match ('x:Key="' + [regex]::Escape($look.Brush) + '"')
        }
    }
}

Describe "getMessageButtons" -Tag Unit {
    It "<buttons> のボタンは左から <texts>（戻り値 <results>）" -TestCases @(
        @{ buttons = "OK";          texts = "OK";                 results = "OK" }
        @{ buttons = "OKCancel";    texts = "OK,キャンセル";       results = "OK,Cancel" }
        @{ buttons = "YesNo";       texts = "はい,いいえ";         results = "Yes,No" }
        @{ buttons = "YesNoCancel"; texts = "はい,いいえ,キャンセル"; results = "Yes,No,Cancel" }
    ) {
        $list = getMessageButtons $buttons
        (@($list | ForEach-Object { $_.Text }) -join ",") | Should -Be $texts
        (@($list | ForEach-Object { $_.Result }) -join ",") | Should -Be $results
    }

    It "既定のボタン（Enter）は、指定が無ければ最初のボタン。指定があればその戻り値のボタン" -TestCases @(
        @{ buttons = "OK";          default = "None"; expected = "OK" }
        @{ buttons = "YesNo";       default = "None"; expected = "Yes" }
        @{ buttons = "YesNo";       default = "No";   expected = "No" }
        @{ buttons = "YesNoCancel"; default = "Cancel"; expected = "Cancel" }
        @{ buttons = "OKCancel";    default = "Yes";  expected = "OK" }
    ) {
        $list = getMessageButtons $buttons $default
        @($list | Where-Object { $_.IsDefault }).Count | Should -Be 1
        (@($list | Where-Object { $_.IsDefault })[0]).Result | Should -Be $expected
        (@($list | Where-Object { $_.Primary })[0]).Result | Should -Be $expected
    }

    It "Esc で押すボタン: OK だけなら OK、取りやめがあればそれ、はい・いいえだけなら無い" -TestCases @(
        @{ buttons = "OK";          cancel = "OK" }
        @{ buttons = "OKCancel";    cancel = "Cancel" }
        @{ buttons = "YesNoCancel"; cancel = "Cancel" }
        @{ buttons = "YesNo";       cancel = "" }
    ) {
        $list = getMessageButtons $buttons
        $found = @($list | Where-Object { $_.IsCancel } | ForEach-Object { $_.Result }) -join ","
        $found | Should -Be $cancel
    }
}

Describe "getMessageCloseResult" -Tag Unit {
    It "<buttons> で閉じたときの戻り値は <expected>" -TestCases @(
        @{ buttons = "OK"; expected = "OK" }
        @{ buttons = "OKCancel"; expected = "Cancel" }
        @{ buttons = "YesNo"; expected = "No" }
        @{ buttons = "YesNoCancel"; expected = "Cancel" }
    ) {
        getMessageCloseResult $buttons | Should -Be $expected
    }
}

Describe "getMessageParts" -Tag Unit {
    It "空行が無ければ、全部が見出し" {
        $parts = getMessageParts "「A」は使えません。`nほかのフォルダを選んでください。"
        $parts.Heading | Should -Be "「A」は使えません。`nほかのフォルダを選んでください。"
        $parts.Hint | Should -Be ""
        $parts.Detail | Should -Be ""
    }

    It "最初の段落が見出し、残りが補足（短いとき）" {
        $parts = getMessageParts "見出しです。`n`n補足の 1 行目。`n補足の 2 行目。"
        $parts.Heading | Should -Be "見出しです。"
        $parts.Hint | Should -Be "補足の 1 行目。`n補足の 2 行目。"
        $parts.Detail | Should -Be ""
    }

    It "残りが長い（行が多い・字が多い）ときは、詳細の枠に回す" -TestCases @(
        @{ rest = (1..6 | ForEach-Object { "行 $_" }) -join "`n" }
        @{ rest = "あ" * 241 }
    ) {
        $parts = getMessageParts "見出し`n`n$rest"
        $parts.Heading | Should -Be "見出し"
        $parts.Hint | Should -Be ""
        $parts.Detail | Should -Be $rest
    }

    It "文言は変えない（CRLF は LF にそろえるだけ）" {
        $parts = getMessageParts "A`r`n`r`nB"
        $parts.Heading | Should -Be "A"
        $parts.Hint | Should -Be "B"
    }
}
