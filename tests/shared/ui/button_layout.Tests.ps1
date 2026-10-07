# ボタンの中身（字・アイコン）が、枠の真ん中に描かれていることのテスト。
# 本物の XAML を読み込んで描き、字・アイコンの画素（インク）の中心が枠の中心から 1 画素以内にあることを確かめる。
# 許す幅を 1 画素にした理由: 日本語の字は、字によって字面が行の枠の上か下に半画素ほど寄る（フォントの作りで、直せない）。
# 一方、枠の高さを超える Height の継承（Button.Base の Height=30）や余白の足し引きによるずれは 2〜3 画素になり、これで見つかる。
# 中身が左に寄せてあるもの（チップ・範囲ボタン）は、左右を見ず、上下だけを見る。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\app_host.ps1"
    . "$PSScriptRoot\..\..\helpers\button_ink.ps1"
    $script:fonts = "${scriptsDir}\tebunko\fonts"
    $script:roots = @{}
    function getRoot([string]$relative) {
        if (-not $script:roots.ContainsKey($relative)) {
            $script:roots[$relative] = loadXaml "${scriptsDir}\tebunko\xaml\$relative" $script:fonts
        }
        return $script:roots[$relative]
    }
}

Describe "ボタンの中身の位置（描いた画素で測る）" -Tag Unit {
    It "<Name>（<Kind>）: 中身の中心が枠の中心から 1 画素以内（上下<Horizontal>）" -TestCases @(
        @{ Kind = "主ボタン"; Name = "SearchButton"; File = "search\search_bar.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "メニューを開くボタン（字と▾）"; Name = "ScopeButton"; File = "search\search_bar.xaml"; Horizontal = ""; CheckX = $false }
        @{ Kind = "種類のチップ"; Name = "KindChipExcel"; File = "search\search_bar.xaml"; Horizontal = ""; CheckX = $false }
        @{ Kind = "アイコンと字のボタン"; Name = "ExportButton"; File = "search\result_list.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "字と▾のボタン"; Name = "ActionsButton"; File = "index\index_list.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "ふつうのボタン"; Name = "IndexingButton"; File = "index\index_list.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "アイコンと字のボタン"; Name = "NewIndexButton"; File = "index\index_list.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "分割ボタンの［開く］"; Name = "OpenButton"; File = "search\preview.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "分割ボタンの▾"; Name = "OpenMenuButton"; File = "search\preview.xaml"; Horizontal = "・左右"; CheckX = $true }
    ) {
        param ($Kind, $Name, $File, $Horizontal, $CheckX)
        $root = getRoot $File
        $button = $root.FindName($Name)
        $button.IsEnabled = $true
        $m = measureInk $root $button
        $m.Groups.Count | Should -BeGreaterThan 0
        [Math]::Abs($m.Dy) | Should -BeLessOrEqual 1
        if ($CheckX) { [Math]::Abs($m.Dx) | Should -BeLessOrEqual 1 }
    }

    It "<Name>: アイコンと字の縦の中心がそろっている（1 画素以内）" -TestCases @(
        @{ Name = "ExportButton"; File = "search\result_list.xaml" }
        @{ Name = "NewIndexButton"; File = "index\index_list.xaml" }
        @{ Name = "ActionsButton"; File = "index\index_list.xaml" }
        @{ Name = "ScopeButton"; File = "search\search_bar.xaml" }
    ) {
        param ($Name, $File)
        $root = getRoot $File
        $m = measureInk $root $root.FindName($Name)
        $m.Groups.Count | Should -Be 2
        [Math]::Abs($m.Groups[0].Dy - $m.Groups[1].Dy) | Should -BeLessOrEqual 1
    }
}
