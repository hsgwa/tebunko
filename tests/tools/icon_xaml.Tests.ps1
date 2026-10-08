# アイコンの元データ（SVG）から画面のベクターの絵（XAML）を作る変換（tools\icon_xaml.ps1）のテスト
BeforeAll {
    $rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
    . "$rootDir\tools\icon_xaml.ps1"

    # 変換の入力に使う SVG（ルートの属性と中身を差し替えられる）
    function newSvg([string]$body, [string]$viewBox = "0 0 256 256") {
        return "<?xml version=`"1.0`"?><svg xmlns=`"http://www.w3.org/2000/svg`" width=`"256`" height=`"256`" viewBox=`"$viewBox`"><!-- 説明 -->$body</svg>"
    }
}

Describe "convertSvgToIconXaml" -Tag Unit {
    It "path の d と fill を、順番どおりに GeometryDrawing へ写す" {
        $xaml = convertSvgToIconXaml (newSvg '<path d="M0,0L10,0L10,10Z" fill="#7abdff"/><path d="M1,1L2,2Z" fill="#1F86EC"/>')
        $xaml | Should -Match 'Brush="#7ABDFF" Geometry="M0,0L10,0L10,10Z"'
        $xaml | Should -Match 'Brush="#1F86EC" Geometry="M1,1L2,2Z"'
        $xaml.IndexOf("#7ABDFF") | Should -BeLessThan $xaml.IndexOf("#1F86EC")
    }

    It "viewBox の大きさの透明な四角を最初に置き、絵の枠にする" {
        $xaml = convertSvgToIconXaml (newSvg '<path d="M0,0L1,1Z" fill="#000000"/>' "0 0 100 50")
        $xaml | Should -Match 'Brush="Transparent" Geometry="M0,0 L100,0 L100,50 L0,50 Z"'
        $xaml.IndexOf("Transparent") | Should -BeLessThan $xaml.IndexOf("#000000")
    }

    It "WPF が読める XAML（DrawingImage）になる" {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        $xaml = convertSvgToIconXaml (newSvg '<path d="M0,0L10,0Q20,5 10,10Z" fill="#3FA5FF"/>')
        $stream = New-Object System.IO.MemoryStream -ArgumentList (,[System.Text.Encoding]::UTF8.GetBytes($xaml.TrimStart([char]0xFEFF)))
        try {
            $image = [System.Windows.Markup.XamlReader]::Load($stream)
        } finally {
            $stream.Dispose()
        }
        $image | Should -BeOfType ([System.Windows.Media.DrawingImage])
        $image.Width | Should -Be 256
    }

    It "改行は CRLF で、末尾も改行" {
        $xaml = convertSvgToIconXaml (newSvg '<path d="M0,0L1,1Z" fill="#000000"/>')
        $xaml | Should -Match "\r\n\z"
        ($xaml -replace "`r`n", "") | Should -Not -Match "[`r`n]"
    }

    It "受け付けない形は、理由を出して止まる: <Name>" -ForEach @(
        @{ Name = "path の transform"; Body = '<path d="M0,0Z" fill="#000000" transform="scale(2)"/>'; Message = "*属性は d と fill だけ*transform*" }
        @{ Name = "path の stroke"; Body = '<path d="M0,0Z" fill="#000000" stroke="#FFFFFF"/>'; Message = "*stroke*" }
        @{ Name = "fill がグラデーションの参照"; Body = '<path d="M0,0Z" fill="url(#g)"/>'; Message = "*#RRGGBB*" }
        @{ Name = "fill が色の名前"; Body = '<path d="M0,0Z" fill="red"/>'; Message = "*#RRGGBB*" }
        @{ Name = "fill が無い"; Body = '<path d="M0,0Z"/>'; Message = "*#RRGGBB*" }
        @{ Name = "d に円弧（A）がある"; Body = '<path d="M0,0A5,5 0 0 1 10,10Z" fill="#000000"/>'; Message = "*d に使えない文字*" }
        @{ Name = "path 以外の図形"; Body = '<circle cx="1" cy="1" r="1" fill="#000000"/>'; Message = "*path だけ*circle*" }
        @{ Name = "グループ"; Body = '<g><path d="M0,0Z" fill="#000000"/></g>'; Message = "*path だけ*" }
        @{ Name = "path が 1 つも無い"; Body = ''; Message = "*path がありません*" }
    ) {
        { convertSvgToIconXaml (newSvg $Body) } | Should -Throw -ExpectedMessage $Message
    }

    It "viewBox が 0 0 <幅> <高さ> の形でなければ止まる" {
        { convertSvgToIconXaml (newSvg '<path d="M0,0Z" fill="#000000"/>' "0 -7 432 432") } | Should -Throw -ExpectedMessage "*viewBox*"
    }

    It "ルートが svg でなければ止まる" {
        { convertSvgToIconXaml '<html xmlns="http://www.w3.org/1999/xhtml"/>' } | Should -Throw -ExpectedMessage "*ルートが svg*"
    }

    It "XML として読めなければ止まる" {
        { convertSvgToIconXaml "<svg" } | Should -Throw -ExpectedMessage "*SVG を読めません*"
    }
}

Describe "画面のアイコン（app_icon.xaml）" -Tag Meta {
    It "元データ（docs\images\logo.svg）から作ったものと一致する（SVG だけ直して作り直し忘れると落ちる。tools\new_icon.ps1 で作り直す）" {
        $svg = [System.IO.File]::ReadAllText("$rootDir\docs\images\logo.svg", (New-Object System.Text.UTF8Encoding($false)))
        $expected = convertSvgToIconXaml $svg
        $actual = [System.IO.File]::ReadAllText("$rootDir\scripts\tebunko\xaml\app_icon.xaml", (New-Object System.Text.UTF8Encoding($true)))
        $actual | Should -BeExactly $expected
    }
}
