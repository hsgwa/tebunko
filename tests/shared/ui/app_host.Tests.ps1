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
    It "欧文は Segoe UI、足りない文字（日本語）は Yu Gothic UI・Meiryo UI で表す（Windows 標準のフォントだけ。同梱のフォルダを指さない）" {
        $family = newAppFontFamily
        $family.Source | Should -Be "Segoe UI, Yu Gothic UI, Meiryo UI"
        $family.BaseUri | Should -BeNullOrEmpty
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

    It "どの呼び出し元の変数にも頼らず、同じ Font.Body を入れる" {
        $path = "$TestDrive\win2.xaml"
        [System.IO.File]::WriteAllText($path, '<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" />', (New-Object System.Text.UTF8Encoding($true)))
        (loadXaml $path).Resources["Font.Body"].Source | Should -Be "Segoe UI, Yu Gothic UI, Meiryo UI"
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

Describe "setGuiErrorLogSource（使う側が登録した式から、記録先と見出しを得る）" -Tag Unit {
    AfterEach {
        setGuiErrorLogSource $null $null
    }

    It "登録が無いときは、記録先は `$null・見出しは例外" {
        setGuiErrorLogSource $null $null
        getGuiErrorLogFile | Should -Be $null
        { getGuiErrorLogHeader } | Should -Throw "*登録されていない*"
    }

    It "登録した式の戻り値を返す" {
        setGuiErrorLogSource { "C:\work\gui_error_log.txt" } { @("見出し 1", "見出し 2") }
        getGuiErrorLogFile | Should -Be "C:\work\gui_error_log.txt"
        @(getGuiErrorLogHeader) | Should -Be @("見出し 1", "見出し 2")
    }
}

Describe "writeErrorLog（記録先のフォルダ名に特別な文字が入っていても書ける）" -Tag Io {
    It "フォルダ名に <label> を含む記録先に、フォルダを作って記録する" -ForEach $pathCharCases {
        $base = Join-Path ([System.IO.Path]::GetTempPath()) ("errlog_" + [guid]::NewGuid().ToString("N"))
        $logFile = Join-Path $base "work_a${ch}b\画面エラー.txt"
        try {
            setGuiErrorLogSource { $logFile }.GetNewClosure() $null
            try { throw "テスト用の例外" } catch { $record = $_ }
            writeErrorLog "確かめ" $record
            [System.IO.File]::Exists($logFile) | Should -Be $true
            [System.IO.File]::ReadAllText($logFile) | Should -Match "テスト用の例外"
        } finally {
            setGuiErrorLogSource $null $null
            # Pester の TestDrive でなく TEMP に作ったので、ワイルドカードを解釈しない .NET で消す
            if ([System.IO.Directory]::Exists($base)) { [System.IO.Directory]::Delete($base, $true) }
        }
    }
}

Describe "writeErrorLog（版と環境の見出しと、大きさでの入れ替え）" -Tag Io {
    BeforeAll {
        function useTestHeader { setGuiErrorLogSource { $global:testErrorLog } { @("==== テストの診断情報 ====", "====") } }
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
        # 前の It が登録した見出しの式を残さない（「無い」場合を本当に無い状態で確かめるため）
        setGuiErrorLogSource { $global:testErrorLog } $null
    }

    AfterAll {
        setGuiErrorLogSource $null $null
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

    It "見出しの式が <label> でも、記録そのものは書かれる" -ForEach @(
        @{ label = "例外を投げる"; define = { setGuiErrorLogSource { $global:testErrorLog } { throw "見出しを作れない" } } }
        @{ label = "無い"; define = { } }
    ) {
        & $define
        writeOneError
        $text = [System.IO.File]::ReadAllText($global:testErrorLog)
        $text | Should -Match "テスト用の例外"
        $text | Should -Not -Match "テストの診断情報"
    }
}
