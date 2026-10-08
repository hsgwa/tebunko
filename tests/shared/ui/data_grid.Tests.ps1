# DataGrid の共通の部品（shared\ui\data_grid.ps1）のテスト。WPF の要素は STA のスレッドでしか作れない。
BeforeDiscovery {
    $sta = [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq "STA"
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\data_grid.ps1"
}

Describe "getDataGridRowAt" -Tag Unit {
    It "行そのものなら、その行を返す" -Skip:(!$sta) {
        $row = New-Object System.Windows.Controls.DataGridRow
        [object]::ReferenceEquals((getDataGridRowAt $row), $row) | Should -Be $true
    }

    It "<name>では `$null" -TestCases @(
        @{ name = "未指定"; make = { $null } }
        @{ name = "列見出し"; make = { New-Object System.Windows.Controls.Primitives.DataGridColumnHeader } }
        @{ name = "スクロールバー"; make = { New-Object System.Windows.Controls.Primitives.ScrollBar } }
        @{ name = "行に着かない要素"; make = { New-Object System.Windows.Controls.TextBlock } }
        @{ name = "表示する要素でないもの"; make = { "文字列" } }
    ) -Skip:(!$sta) {
        getDataGridRowAt (& $make) | Should -BeNullOrEmpty
    }
}
