# スクリプトの構成（パス定義・構文・XAML・型の読み込み）のテスト
. "$PSScriptRoot\..\helpers\load.ps1"

Describe "パス定義" -Tag Meta {
    It "リポジトリ直下を基準にする" {
        $rootDir | Should Be (Resolve-Path "$here\..").Path
        $indexDir | Should Be "$rootDir\work\index"
        $publishDir | Should Be "$rootDir\work\変換出力\$PID"
        $resultFile | Should Be "$rootDir\work\検索結果.txt"
        $settingsFile | Should Be "$rootDir\setting.config"
    }
}

Describe "実行時コンパイル（csc.exe）を使わない" -Tag Meta {
    # 画面・共通・変換の各スクリプトが Add-Type -TypeDefinition（実行時コンパイル）を使わないこと。
    # 画面で使う型は PowerShell class に移した（csc.exe の親子関係・一時 DLL を出さないため）
    It "scripts に Add-Type -TypeDefinition が無い" {
        foreach ($file in (Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.ps1")) {
            $source = Get-Content $file.FullName -Raw -Encoding UTF8
            ($source -match "Add-Type\s+-TypeDefinition") | Should Be $false
        }
    }
}

Describe "スクリプトの構文" -Tag Meta {
    Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.ps1" | ForEach-Object {
        $script = $_

        It "$($script.Name) に構文エラーが無い" {
            $errors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$null, [ref]$errors) | Out-Null
            # 継承元の型が別ファイルにある場合、1 ファイルだけを読むと型が見つからない（TypeNotFound）。
            # 読み込む順で解決できることは、下の「型の読み込み」で実際に読み込んで確かめる
            @($errors | Where-Object { $_.ErrorId -ne "TypeNotFound" }).Count | Should Be 0
        }
    }
}

Describe "画面定義（XAML）" -Tag Meta {
    $xamlNs = "http://schemas.microsoft.com/winfx/2006/xaml"

    Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.xaml" | ForEach-Object {
        $file = $_

        It "$($file.Name) が XML として読める" {
            { [xml](Get-Content $file.FullName -Raw -Encoding UTF8) } | Should Not Throw
        }
    }

    It "フォルダ選択の画面に、config_gui.ps1 が使う x:Name がすべてある" {
        [xml]$xaml = Get-Content "$here\..\scripts\shared\xaml\dialog_folder_select.xaml" -Raw -Encoding UTF8
        $names = @($xaml.SelectNodes("//*") | ForEach-Object { $_.GetAttribute("Name", $xamlNs) } | Where-Object { $_ -ne "" })
        foreach ($name in @(
                "DescriptionText", "BackButton", "ForwardButton", "UpButton", "AddressBox",
                "FolderTree", "EntryList", "EntryPlaceholder", "StatusText", "FolderBox", "OkButton", "ErrorText")) {
            $names -contains $name | Should Be $true
        }
    }
}

Describe "型の読み込み" -Tag Meta {
    # 画面で使う型は shared と win_grep に分かれている。gui.ps1 と同じ順で読み込めば、
    # 継承（NotifyBase を継承する型）が解決できることを確かめる
    It "shared と win_grep の型を順に読み込める" {
        $probe = Join-Path $TestDrive "probe.ps1"
        $scripts = (Resolve-Path "$here\..\scripts").Path
        Set-Content -LiteralPath $probe -Encoding UTF8 -Value @(
            'Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase'
            ". `"$scripts\shared\ui\types.ps1`""
            ". `"$scripts\win_grep\ui\types_grep.ps1`""
            '([HitRow], [IndexNode], [FolderNode], [ConfirmFact], [PreviewTable]).Count'
        )
        $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $probe 2>&1
        ($output -join "") | Should Be "5"
    }
}
