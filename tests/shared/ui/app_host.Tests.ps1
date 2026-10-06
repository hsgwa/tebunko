# 画面の土台（shared\ui\app_host.ps1）の XAML 読み込みと、theme.xaml への参照の差し替えのテスト。
# 設計は docs/design/structure/single-script.md。単一 .ps1 版（${bundledXaml}）では、
# 参照先のファイルが実在しなくても働くことを確かめる（Resolve-Path ではなく GetFullPath で正規化するため）。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\app_host.ps1"
}

Describe "getXamlText" -Tag Unit {
    It "`${bundledXaml} が無ければファイルから読む（zip 版の形）" {
        $path = "TestDrive:\real.xaml" | Resolve-Path -ErrorAction SilentlyContinue
        $file = "$TestDrive\real.xaml"
        [System.IO.File]::WriteAllText($file, "<ResourceDictionary/>", (New-Object System.Text.UTF8Encoding($true)))
        (getXamlText $file) | Should -Be "<ResourceDictionary/>"
    }

    It "`${bundledXaml} にあれば、ファイルが実在しなくてもそちらを返す（単一 .ps1 版の形）" {
        $missing = "$TestDrive\bundled\missing.xaml"
        $global:bundledXaml = @{ ([System.IO.Path]::GetFullPath($missing)) = "<ResourceDictionary><!-- 埋め込み --></ResourceDictionary>" }
        try {
            Test-Path -LiteralPath $missing | Should -Be $false
            (getXamlText $missing) | Should -Be "<ResourceDictionary><!-- 埋め込み --></ResourceDictionary>"
        } finally {
            Remove-Variable -Name bundledXaml -Scope Global -ErrorAction SilentlyContinue
        }
    }
}

Describe "loadXaml（theme.xaml への参照の差し替え）" -Tag Unit {
    It "`${bundledXaml} だけで、参照先（theme.xaml 相当）が実在しなくても読み込める" {
        # 実ファイルは一切作らない。tebunko.xaml が ../../shared/xaml/theme.xaml を指す今の形に合わせ、
        # $TestDrive\xaml\win.xaml から見た相対参照で $TestDrive\shared\xaml\theme.xaml を指す
        $mainPath = "$TestDrive\xaml\win.xaml"
        $themePath = "$TestDrive\shared\xaml\theme.xaml"
        $mainXml = @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <ResourceDictionary.MergedDictionaries>
        <ResourceDictionary Source="../shared/xaml/theme.xaml" />
    </ResourceDictionary.MergedDictionaries>
    <SolidColorBrush x:Key="Own" Color="#111111" />
</ResourceDictionary>
'@
        $themeXml = @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <SolidColorBrush x:Key="Accent" Color="#223344" />
</ResourceDictionary>
'@
        $global:bundledXaml = @{
            ([System.IO.Path]::GetFullPath($mainPath))  = $mainXml
            ([System.IO.Path]::GetFullPath($themePath)) = $themeXml
        }
        try {
            Test-Path -LiteralPath $mainPath | Should -Be $false
            Test-Path -LiteralPath $themePath | Should -Be $false
            $result = loadXaml $mainPath
            $result.Contains("Own") | Should -Be $true
            $result.Contains("Accent") | Should -Be $true
            $result["Accent"].Color.ToString() | Should -Be "#FF223344"
        } finally {
            Remove-Variable -Name bundledXaml -Scope Global -ErrorAction SilentlyContinue
        }
    }
}

Describe "newAppFontFamily（画面の既定のフォント）" -Tag Unit {
    BeforeAll {
        $bundledFonts = "${scriptsDir}\shared\fonts"
    }

    It "フォルダがあれば、そこを指し、Rethink Sans を先に Yu Gothic UI・Meiryo UI を足りない文字の代わりにする" {
        $family = newAppFontFamily $bundledFonts
        $family.BaseUri.LocalPath | Should -Be ("$bundledFonts\")
        $family.Source | Should -Be "./#Rethink Sans, Yu Gothic UI, Meiryo UI"
        # 同梱のフォルダから Rethink Sans が見つかる
        @([System.Windows.Media.Fonts]::GetFontFamilies($family.BaseUri) | ForEach-Object { $_.FamilyNames.Values }) | Should -Contain "Rethink Sans"
    }

    It "<name> は、Yu Gothic UI・Meiryo UI だけにする（単一 .ps1 版など、フォントを同梱しない形）" -ForEach @(
        @{ name = "フォルダが無い"; folder = "$TestDrive\no_such_fonts" }
        @{ name = "指定が空" ; folder = "" }
    ) {
        $family = newAppFontFamily $folder
        $family.Source | Should -Be "Yu Gothic UI, Meiryo UI"
    }

    It "同梱のフォントのファイルとライセンスの文面がある" {
        foreach ($name in "RethinkSans-wght.ttf", "RethinkSans-Italic-wght.ttf", "OFL.txt", "LICENSE-Lucide.txt") {
            Test-Path -LiteralPath "$bundledFonts\$name" | Should -Be $true
        }
    }
}

Describe "loadXaml（Font.Body）" -Tag Unit {
    It "FrameworkElement には Font.Body を入れ、ルートの DynamicResource で使える" {
        $path = "$TestDrive\win.xaml"
        [System.IO.File]::WriteAllText($path, '<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TextElement.FontFamily="{DynamicResource Font.Body}" />', (New-Object System.Text.UTF8Encoding($true)))
        $result = loadXaml $path
        $result.Resources.Contains("Font.Body") | Should -Be $true
        $result.Resources["Font.Body"] | Should -BeOfType [System.Windows.Media.FontFamily]
    }

    It "渡したフォントのフォルダから Font.Body を作る（呼び出し側の変数を暗黙に読まない）" {
        $path = "$TestDrive\win2.xaml"
        [System.IO.File]::WriteAllText($path, '<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" />', (New-Object System.Text.UTF8Encoding($true)))
        $fonts = "${scriptsDir}\shared\fonts"
        # 同じ名前の変数が呼び出し側にあっても、引数のほうを使う
        $fontsDir = "$TestDrive\no_such_fonts"
        $result = loadXaml $path $fonts
        $result.Resources["Font.Body"].BaseUri.LocalPath | Should -Be "$fonts\"
        # 引数が無いときは、変数があっても Yu Gothic UI・Meiryo UI だけ
        (loadXaml $path).Resources["Font.Body"].Source | Should -Be "Yu Gothic UI, Meiryo UI"
    }

    It "loadWindow は、フォントのフォルダを loadXaml に渡す" {
        $text = Get-Content -LiteralPath "${scriptsDir}\shared\ui\app_host.ps1" -Raw -Encoding UTF8
        $text | Should -Match 'function loadWindow \{[\s\S]*?\$loaded = loadXaml \$path \$fontsFolder'
    }
}
