# 画面の土台（shared\ui\app_host.ps1）の XAML 読み込みと、theme.xaml への参照の差し替えのテスト。
# 設計は docs/design/structure/single-script.md。単一 .ps1 版（${bundledXaml}）では、
# 参照先のファイルが実在しなくても働くことを確かめる（Resolve-Path ではなく GetFullPath で正規化するため）。
BeforeDiscovery {
    . "$PSScriptRoot\..\..\helpers\path_chars.ps1"
    $pathCharCases = getPathCharCases
}

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

Describe "loadAppIcon（画面のアイコン）" -Tag Unit {
    BeforeEach {
        ${script:appIcon} = $null
    }

    It "アイコンの XAML（DrawingImage）を読み、凍結して返す。2 回目は同じものを返す" {
        ${iconXamlFile} = "${scriptsDir}\tebunko\xaml\app_icon.xaml"
        $icon = loadAppIcon
        $icon | Should -BeOfType ([System.Windows.Media.DrawingImage])
        $icon.IsFrozen | Should -BeTrue
        [object]::ReferenceEquals((loadAppIcon), $icon) | Should -BeTrue
    }

    It "読めなければ `$null を返す（アイコンが無くても画面は開ける）" {
        ${iconXamlFile} = "$TestDrive\none\app_icon.xaml"
        loadAppIcon | Should -BeNullOrEmpty
    }

    It "単一 .ps1 版のように、XAML が文字列で埋め込まれていて実在しなくても読める" {
        $missing = "$TestDrive\bundled\app_icon.xaml"
        ${iconXamlFile} = $missing
        $text = [System.IO.File]::ReadAllText("${scriptsDir}\tebunko\xaml\app_icon.xaml", (New-Object System.Text.UTF8Encoding($true)))
        $bundledXaml = @{ ([System.IO.Path]::GetFullPath($missing)) = $text }
        loadAppIcon | Should -BeOfType ([System.Windows.Media.DrawingImage])
    }
}

Describe "writeErrorLog（記録先のフォルダ名に特別な文字が入っていても書ける）" -Tag Io {
    It "フォルダ名に <label> を含む記録先に、フォルダを作って記録する" -ForEach $pathCharCases {
        $base = Join-Path ([System.IO.Path]::GetTempPath()) ("errlog_" + [guid]::NewGuid().ToString("N"))
        $logFile = Join-Path $base "work_a${ch}b\画面エラー.txt"
        try {
            function getGuiErrorLogFile { return $logFile }
            try { throw "テスト用の例外" } catch { $record = $_ }
            writeErrorLog "確かめ" $record
            [System.IO.File]::Exists($logFile) | Should -Be $true
            [System.IO.File]::ReadAllText($logFile) | Should -Match "テスト用の例外"
        } finally {
            # Pester の TestDrive でなく TEMP に作ったので、ワイルドカードを解釈しない .NET で消す
            if ([System.IO.Directory]::Exists($base)) { [System.IO.Directory]::Delete($base, $true) }
        }
    }
}

Describe "writeErrorLog（版と環境の見出しと、大きさでの入れ替え）" -Tag Io {
    BeforeAll {
        function global:useTestHeader { function global:getGuiErrorLogHeader { return @("==== テストの診断情報 ====", "====") } }
        function writeOneError {
            try { throw "テスト用の例外" } catch { writeErrorLog "確かめ" $_ }
        }
        function countHeaders {
            param ([string]$path)
            return ([regex]::Matches([System.IO.File]::ReadAllText($path), "==== テストの診断情報 ====")).Count
        }
    }

    BeforeEach {
        $script:errorLogHeaderPath = $null
        $global:testErrorLog = Join-Path $TestDrive ("err_" + [guid]::NewGuid().ToString("N") + "\gui_error_log.txt")
        function global:getGuiErrorLogFile { return $global:testErrorLog }
        # 前の It が定義した見出しの関数を残さない（「無い」場合を本当に無い状態で確かめるため）
        Remove-Item function:global:getGuiErrorLogHeader -ErrorAction SilentlyContinue
    }

    AfterAll {
        Remove-Item function:global:getGuiErrorLogFile, function:global:getGuiErrorLogHeader, function:global:useTestHeader -ErrorAction SilentlyContinue
        Remove-Variable -Name testErrorLog -Scope Global -ErrorAction SilentlyContinue
    }

    It "1 回目の記録の前に見出しがあり、同じファイルへの 2 回目には無い" {
        useTestHeader
        writeOneError
        writeOneError
        $text = [System.IO.File]::ReadAllText($global:testErrorLog)
        $text | Should -Match "^﻿?==== テストの診断情報 ===="
        (countHeaders $global:testErrorLog) | Should -Be 1
        ([regex]::Matches($text, "==== [\d\- :]+ 確かめ ====")).Count | Should -Be 2
    }

    It "1 MB 以上なら _previous に移し、新しいファイルの先頭に見出しを書いてから記録する" {
        useTestHeader
        writeOneError
        [System.IO.File]::AppendAllText($global:testErrorLog, ("x" * (1MB + 1)))
        writeOneError
        $previous = getPreviousLogPath $global:testErrorLog
        [System.IO.File]::Exists($previous) | Should -Be $true
        [System.IO.File]::ReadAllText($previous) | Should -Match "x{1000}"
        $text = [System.IO.File]::ReadAllText($global:testErrorLog)
        $text | Should -Match "^﻿?==== テストの診断情報 ===="
        $text | Should -Not -Match "x{1000}"
        $text | Should -Match "テスト用の例外"
    }

    It "記録先のパスが変わったら、新しいファイルにも見出しを書く" {
        useTestHeader
        writeOneError
        $first = $global:testErrorLog
        $global:testErrorLog = Join-Path $TestDrive ("err2_" + [guid]::NewGuid().ToString("N") + "\gui_error_log.txt")
        writeOneError
        (countHeaders $first) | Should -Be 1
        (countHeaders $global:testErrorLog) | Should -Be 1
    }

    It "見出しの関数が <label> でも、記録そのものは書かれる" -ForEach @(
        @{ label = "例外を投げる"; define = { function global:getGuiErrorLogHeader { throw "見出しを作れない" } } }
        @{ label = "無い"; define = { } }
    ) {
        & $define
        if ($label -eq "無い") {
            # 読み込んだ本物の定義も外して、本当に無い状態にする（この It の中だけ）
            Remove-Item function:getGuiErrorLogHeader -ErrorAction SilentlyContinue
            Get-Command getGuiErrorLogHeader -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        }
        writeOneError
        $text = [System.IO.File]::ReadAllText($global:testErrorLog)
        $text | Should -Match "テスト用の例外"
        $text | Should -Not -Match "テストの診断情報"
    }
}