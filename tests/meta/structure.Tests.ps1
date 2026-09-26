# スクリプトの構成（パス定義・構文・XAML・型の読み込み）のテスト
. "$PSScriptRoot\..\helpers\load.ps1"

Describe "パス定義" -Tag Meta {
    It "リポジトリ直下を基準にする（書き込めるため、設定ファイルもリポジトリ直下に置く）" {
        $rootDir | Should Be (Resolve-Path "$here\..").Path
        $dataDir | Should Be $rootDir
        $settingsFile | Should Be "$rootDir\setting.config"
    }

    It "work の中身は work の置き場所（既定はリポジトリ直下の work。setting.config の workspaceFolder で変わる）を基準にする" {
        $workspace.Dir | Should Be (getWorkDir $settingsFile)
        $workspace.IndexDir | Should Be "$($workspace.Dir)\index"
        $workspace.PublishDir | Should Be "$($workspace.Dir)\取り込み出力\$PID"
        $workspace.ResultFile | Should Be "$($workspace.Dir)\検索結果.txt"
    }
}

Describe "画面の部品でのパスの組み立て" -Tag Meta {
    # ui\ 配下のファイルは gui.ps1 から dot-source する部品。中で $PSScriptRoot を使うと ui\ を指すため、
    # "${PSScriptRoot}\tebunko\indexer.ps1" のように起動口からの相対パスを書くと存在しないパスになる
    # （［インデックス作成を開始］でインデクサが起動しなかった不具合）。パスは起動口（gui.ps1）で決めて変数で渡す
    It "ui 配下のスクリプトで `$PSScriptRoot を使っていない" {
        $found = @(Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.ps1" |
            Where-Object { $_.DirectoryName -match '\\ui$' } |
            Select-String -Pattern '\$\{?PSScriptRoot\}?' |
            ForEach-Object { "$($_.Filename):$($_.LineNumber)" })
        ($found -join ", ") | Should Be ""
    }

    It "gui.ps1 が指すインデクサのファイルがある" {
        $line =@(Select-String -Path "$here\..\scripts\tebunko\gui.ps1" -Pattern '^\$\{indexerScriptPath\}\s*=\s*"\$PSScriptRoot\\(.+)"')
        $line.Count | Should Be 1
        Test-Path -LiteralPath "$here\..\scripts\tebunko\$($line[0].Matches[0].Groups[1].Value)" | Should Be $true
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
    Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.xaml" | ForEach-Object {
        $file = $_

        It "$($file.Name) が XML として読める" {
            { [xml](Get-Content $file.FullName -Raw -Encoding UTF8) } | Should Not Throw
        }
    }
}

Describe "型の読み込み" -Tag Meta {
    # 画面で使う型は shared と tebunko に分かれている。gui.ps1 と同じ順で読み込めば、
    # 継承（NotifyBase を継承する型）が解決できることを確かめる
    It "shared と tebunko の型を順に読み込める" {
        $probe = Join-Path $TestDrive "probe.ps1"
        $scripts = (Resolve-Path "$here\..\scripts").Path
        Set-Content -LiteralPath $probe -Encoding UTF8 -Value @(
            'Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase'
            ". `"$scripts\shared\ui\types.ps1`""
            ". `"$scripts\tebunko\ui\types.ps1`""
            '([HitRow], [IndexNode], [ConfirmFact], [PreviewTable]).Count'
        )
        $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $probe 2>&1
        ($output -join "") | Should Be "4"
    }
}

Describe "画面の部品の名前" -Tag Meta {
    # gui.ps1 が FindName で取る名前が、XAML に実在すること。
    # タブの中身を別ファイルに分けているため、名前を足したり動かしたりすると気づきにくい
    $xamlNs = "http://schemas.microsoft.com/winfx/2006/xaml"
    $gui = [System.IO.File]::ReadAllText("$here\..\scripts\tebunko\gui.ps1")

    function getXamlNames {
        param ([string]$path)
        [xml]$xaml = Get-Content $path -Raw -Encoding UTF8
        return @($xaml.SelectNodes("//*") | ForEach-Object { $_.GetAttribute("Name", $xamlNs) } | Where-Object { $_ -ne "" })
    }

    It "ウィンドウの枠の名前がある" {
        $names = getXamlNames "$here\..\scripts\tebunko\xaml\tebunko.xaml"
        foreach ($name in @("Tabs", "IndexTab", "SearchTab", "SettingsTab", "KillTab", "IndexTabHeader", "KillTabHeader", "StatusText")) {
            $names -contains $name | Should Be $true
        }
    }

    foreach ($tab in @(
            @{ File = "tab_index.xaml"; Marker = 'Tab = "IndexTab"' }
            @{ File = "tab_search.xaml"; Marker = 'Tab = "SearchTab"' }
            @{ File = "tab_settings.xaml"; Marker = 'Tab = "SettingsTab"' }
            @{ File = "tab_kill.xaml"; Marker = 'Tab = "KillTab"' })) {
        $file = $tab.File
        $marker = $tab.Marker

        It "$file に、gui.ps1 が使う名前がすべてある" {
            # gui.ps1 の $tabs から、そのタブの名前の一覧を取り出す
            $start = $gui.IndexOf($marker)
            $start | Should Not Be -1
            $listStart = $gui.IndexOf("Names = @(", $start)
            $listEnd = $gui.IndexOf(") }", $listStart)
            $list = $gui.Substring($listStart, $listEnd - $listStart)
            $wanted = @([regex]::Matches($list, '"([A-Za-z]+)"') | ForEach-Object { $_.Groups[1].Value })
            $wanted.Count -gt 0 | Should Be $true

            $names = getXamlNames "$here\..\scripts\tebunko\xaml\$file"
            $missing = @($wanted | Where-Object { $names -notcontains $_ })
            ($missing -join ", ") | Should Be ""
        }
    }
}