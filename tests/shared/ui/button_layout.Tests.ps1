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
    # 名前で探す。行のボタン（RowUpdate / RowStop）は表の行の雛形の中にあるので、雛形を 1 つ作って Tag で探す。
    function findButton($root, [string]$name) {
        if ($name -like "Row*") {
            foreach ($c in $root.FindName("IndexGrid").Columns) {
                if ($c -isnot [System.Windows.Controls.DataGridTemplateColumn]) { continue }
                $e = $c.CellTemplate.LoadContent()
                $q = New-Object System.Collections.Queue; $q.Enqueue($e)
                while ($q.Count) {
                    $x = $q.Dequeue()
                    if ($x -is [System.Windows.Controls.Button] -and "$($x.Tag)" -eq $name) {
                        $p = $x
                        while ($p) { if ($p -is [System.Windows.UIElement]) { $p.Visibility = "Visible" }; $p = [System.Windows.LogicalTreeHelper]::GetParent($p) }
                        # 雛形の根は背景が透明で面の色が取れないので、白い台に載せて描く。
                        $stand = New-Object System.Windows.Controls.Grid
                        $stand.Background = [System.Windows.Media.Brushes]::White
                        $stand.Children.Add($e) | Out-Null
                        return @{ Root = $stand; Button = $x }
                    }
                    foreach ($ch in [System.Windows.LogicalTreeHelper]::GetChildren($x)) { if ($ch -is [System.Windows.DependencyObject]) { $q.Enqueue($ch) } }
                }
            }
            throw "ボタンが無い: $name"
        }
        $b = $root.FindName($name)
        # 折りたたまれた親の中（詳細など）は大きさが 0 になるので、測るあいだだけ表に出す。
        $x = $b
        while ($x) { if ($x -is [System.Windows.UIElement]) { $x.Visibility = "Visible" }; $x = [System.Windows.LogicalTreeHelper]::GetParent($x) }
        return @{ Root = $root; Button = $b }
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
        @{ Kind = "アイコンと字のボタン（空の表）"; Name = "IndexEmptyAddButton"; File = "index\index_list.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "行のアイコンと字のボタン"; Name = "RowUpdate"; File = "index\index_list.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "行のアイコンと字のボタン"; Name = "RowStop"; File = "index\index_list.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "詳細の［...］（点は図形）"; Name = "IndexDetailPathButton"; File = "index\index_detail.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "分割ボタンの［開く］"; Name = "OpenButton"; File = "search\preview.xaml"; Horizontal = "・左右"; CheckX = $true }
        @{ Kind = "分割ボタンの▾"; Name = "OpenMenuButton"; File = "search\preview.xaml"; Horizontal = "・左右"; CheckX = $true }
    ) {
        param ($Kind, $Name, $File, $Horizontal, $CheckX)
        $f = findButton (getRoot $File) $Name
        $f.Button.IsEnabled = $true
        $m = measureInk $f.Root $f.Button
        $m.Groups.Count | Should -BeGreaterThan 0
        [Math]::Abs($m.Dy) | Should -BeLessOrEqual 1
        if ($CheckX) { [Math]::Abs($m.Dx) | Should -BeLessOrEqual 1 }
    }

    It "<Name>: アイコンと字の縦の中心がそろっている（1 画素以内）" -TestCases @(
        @{ Name = "ExportButton"; File = "search\result_list.xaml" }
        @{ Name = "NewIndexButton"; File = "index\index_list.xaml" }
        @{ Name = "ActionsButton"; File = "index\index_list.xaml" }
        @{ Name = "ScopeButton"; File = "search\search_bar.xaml" }
        @{ Name = "RowUpdate"; File = "index\index_list.xaml" }
        @{ Name = "RowStop"; File = "index\index_list.xaml" }
    ) {
        param ($Name, $File)
        $f = findButton (getRoot $File) $Name
        $m = measureInk $f.Root $f.Button
        $m.Groups.Count | Should -Be 2
        [Math]::Abs($m.Groups[0].Dy - $m.Groups[1].Dy) | Should -BeLessOrEqual 1
    }

    # 無効の［検索］: 字が薄く、面の色の取り違えで大きくずれたように見えやすい。実際には有効のときと同じ位置にある。
    It "SearchButton（無効）: 中身の中心が枠の中心から 1 画素以内" {
        $root = getRoot "search\search_bar.xaml"
        $b = $root.FindName("SearchButton")
        $b.IsEnabled = $false
        try {
            $m = measureInk $root $b
            $m.Groups.Count | Should -BeGreaterThan 0
            [Math]::Abs($m.Dy) | Should -BeLessOrEqual 1
            [Math]::Abs($m.Dx) | Should -BeLessOrEqual 1
        } finally { $b.IsEnabled = $true }
    }
}

# 帯の右のボタン（［中止］［続きから再開］［ログを開く］）は、帯が折りたたまれた入れ子の中にあり、単独では描けない。
# 本物の画面（tools/capture_screens.ps1 の index-tab/running）で、中心のずれが 1 画素以内であることを確かめてある。
