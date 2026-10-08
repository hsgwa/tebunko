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

# 帯のボタンの見える枠（Bd）の高さ。Height（24）はボタン全体で、フォーカスの輪が内側から取ると見える枠が 22 に縮む（Margin=-1 で外へ出してある）。
Describe "帯のボタンの見える枠の大きさ" -Tag Unit {
    It "<Name>: 見える枠の高さが 24" -TestCases @(
        @{ Name = "IndexingSearchButton" }
        @{ Name = "IndexingResumeButton" }
        @{ Name = "IndexingStopButton" }
    ) {
        param ($Name)
        $f = findButton (getRoot "index\index.xaml") $Name
        $f.Root.Measure([System.Windows.Size]::new(1000, 600))
        $f.Root.Arrange([System.Windows.Rect]::new(0, 0, 1000, 600))
        $f.Root.UpdateLayout()
        $bd = $f.Button.Template.FindName("Bd", $f.Button)
        $bd.ActualHeight | Should -Be 24
        $f.Button.ActualHeight | Should -Be 24
    }

    # 帯の右に 2 つ以上並ぶ（中止のあとで［検索する］と［続きから再開］、更新中は見積もり時間と［中止］）とき、間は 10
    It "並んだボタンの間は 10（見える枠どうし）" {
        $root = getRoot "index\index.xaml"
        $names = "IndexingSearchButton", "IndexingResumeButton", "IndexingStopButton"
        $buttons = @($names | ForEach-Object { (findButton $root $_).Button })
        $root.Measure([System.Windows.Size]::new(1000, 600))
        $root.Arrange([System.Windows.Rect]::new(0, 0, 1000, 600))
        $root.UpdateLayout()
        $frames = @($buttons | ForEach-Object {
            $bd = $_.Template.FindName("Bd", $_)
            $bd.TransformToAncestor($root).TransformBounds([System.Windows.Rect]::new(0, 0, $bd.ActualWidth, $bd.ActualHeight))
        })
        ($frames[1].Left - $frames[0].Right) | Should -Be 10
        ($frames[2].Left - $frames[1].Right) | Should -Be 10
    }
}

# 帯の右のボタン（［中止］［続きから再開］［検索する］）は、帯が折りたたまれた入れ子の中にあり、単独では描けない。
# 本物の画面（tools/capture_screens.ps1 の index-tab/running）で、中心のずれが 1 画素以内であることを確かめてある。

# メッセージの画面（dialog_confirm.xaml）。窓は開かず、中身を白い台に移して描く（窓の外の余白が写らないように）。
Describe "メッセージの画面のボタン・見出しの位置（描いた画素で測る）" -Tag Unit {
    BeforeAll {
        function newMessageRoot {
            $w = loadXaml "${scriptsDir}\shared\xaml\dialog_confirm.xaml" "${scriptsDir}\shared\fonts"
            $content = $w.Content
            $w.Content = $null
            $root = New-Object System.Windows.Controls.Grid
            $root.Background = $w.Background
            $res = $w.Resources
            $w.Resources = New-Object System.Windows.ResourceDictionary
            $root.Resources = $res
            $root.Children.Add($content) | Out-Null
            $w.FindName("HeadingText").Text = "前回の更新で起動した Office が残ったまま動いています。"
            $w.FindName("HeadingIconHost").Visibility = "Visible"
            $panel = $w.FindName("ButtonPanel")
            foreach ($spec in @(@("キャンセル", "Default"), @("いいえ", "Default"), @("OK", "Primary"), @("終了する", "Danger.Filled"))) {
                $b = New-Object System.Windows.Controls.Button
                $b.Content = $spec[0]
                if ($spec[1] -ne "Default") { $b.Style = $root.FindResource($spec[1]) }
                $panel.Children.Add($b) | Out-Null
            }
            return @{ Window = $w; Root = $root; Panel = $panel }
        }
    }

    It "<Name>: ボタンの字の中心が枠の中心から 1 画素以内で、高さ・最小幅・間隔がそろう" -TestCases @(
        @{ Name = "キャンセル（ふつう）"; Index = 0 }
        @{ Name = "いいえ（ふつう）"; Index = 1 }
        @{ Name = "OK（主なボタン）"; Index = 2 }
        @{ Name = "終了する（取り消せない操作）"; Index = 3 }
    ) {
        param ($Name, $Index)
        $m0 = newMessageRoot
        $b = $m0.Panel.Children[$Index]
        $m = measureInk $m0.Root $b -Width 520 -Height 300
        $m.Groups.Count | Should -BeGreaterThan 0
        [Math]::Abs($m.Dy) | Should -BeLessOrEqual 1
        [Math]::Abs($m.Dx) | Should -BeLessOrEqual 1
        $b.ActualHeight | Should -Be 32
        $b.ActualWidth | Should -BeGreaterOrEqual 80
        $b.Margin.Left | Should -Be 8
    }

    It "見出しの左のアイコンと 1 行目の字の縦の中心がそろっている（1 画素以内）" {
        $m0 = newMessageRoot
        $icon = $m0.Window.FindName("HeadingIconHost")
        $m = measureInk $m0.Root $icon $icon.Parent -Width 520 -Height 300
        $m.Groups.Count | Should -BeGreaterThan 1
        $textDy = ($m.Groups | Select-Object -Skip 1 | ForEach-Object { $_.Dy } | Measure-Object -Average).Average
        [Math]::Abs($m.Groups[0].Dy - $textDy) | Should -BeLessOrEqual 1
    }
}
