# プレビューの表の左端の揃いのテスト（scripts\tebunko\xaml\search\preview.xaml）。
# 本物の XAML を読み込んで描き、表が見える幅より狭いときも、行（行番号とセル）が見出しと同じ左端から並ぶことを確かめる。
# 行の幅を固定している（RowWidth）ため、行を左に寄せていないと、枠の真ん中に出て見出しとずれる。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_types.ps1"
    . "${scriptsDir}\tebunko\ui\preview_view.ps1"
    . "${scriptsDir}\shared\ui\app_host.ps1"

    # 行の入れ物（x:Name が RowPanel の StackPanel）を、見た目の木から集める
    function findRowPanels($element) {
        $found = New-Object System.Collections.Generic.List[object]
        $queue = New-Object System.Collections.Queue
        $queue.Enqueue($element)
        while ($queue.Count -gt 0) {
            $e = $queue.Dequeue()
            if ($e -is [System.Windows.FrameworkElement] -and $e.Name -eq "RowPanel") { $found.Add($e) }
            for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($e); $i++) {
                $queue.Enqueue([System.Windows.Media.VisualTreeHelper]::GetChild($e, $i))
            }
        }
        , $found
    }
}

Describe "プレビューの表の左端" -Tag Unit {
    It "表が見える幅より狭いとき（<Columns> 列）も、行は左端から並ぶ" -TestCases @(
        @{ Columns = 2 }
        @{ Columns = 2000 }
    ) {
        param ($Columns)
        $cells = @(1..$Columns | ForEach-Object { "c$_" })
        $line = $cells -join "`t"
        $row = [HitRow]::Create("営業", "C:\index", "営業\a.xlsx", "営業", "a.tsv", "a.xlsx", "[シート]表", 2, $line, "c1", $null)
        $row.IsExcel = $true
        $table = $row.BuildPreview([int[]](1, 2), [string[]]($line, $line))
        $content = loadXaml "${scriptsDir}\tebunko\xaml\search\preview.xaml"
        $rowsControl = $content.FindName("PreviewRows")
        $content.FindName("PreviewHeaderPanel").DataContext = $table.Window
        $rowsControl.ItemsSource = $table.Rows
        $content.FindName("PreviewHeaderScroll").Visibility = "Visible"
        $content.Measure([System.Windows.Size]::new(1400, 700))
        $content.Arrange([System.Windows.Rect]::new(0, 0, 1400, 700))
        $content.UpdateLayout()
        $table.EnsureCells(0, 10)
        $content.UpdateLayout()

        $panels = findRowPanels $rowsControl
        $panels.Count | Should -BeGreaterThan 0
        $origin = [System.Windows.Point]::new(0, 0)
        $headerX = $content.FindName("PreviewHeaderPanel").TransformToAncestor($content).Transform($origin).X
        foreach ($panel in $panels) {
            $x = $panel.TransformToAncestor($content).Transform($origin).X
            [math]::Abs($x - $headerX) | Should -BeLessOrEqual 1 -Because "行の左端が、見出しの左端からずれない"
        }
    }
}
