# 設定ファイル（tebunko_grep\core\settings_grep.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "readSettings / writeSettings" -Tag Io {
    It "ファイルが無ければ既定値を返し、ファイルを作らない" {
        $path = "$TestDrive\空\setting.config"
        $settings = readSettings $path
        @($settings.targetFolders).Count | Should Be 0
        @($settings.indexSources).Count | Should Be 0
        $settings.useRegex | Should Be $false
        $settings.openMode | Should Be "normal"
        Test-Path -LiteralPath $path | Should Be $false
    }

    It "開き方（openMode）を保存・読み込みできる" {
        $path = "$TestDrive\開き方\setting.config"
        foreach ($mode in ${openModes}) {
            writeOpenMode $mode $path
            readOpenMode $path | Should Be $mode
        }
    }

    It "開き方が無い・知らない値なら「通常」とする" {
        $path = "$TestDrive\開き方2\setting.config"
        readOpenMode $path | Should Be ${openModeNormal}        # ファイルが無い
        updateSettings "openMode" "知らない値" $path
        readOpenMode $path | Should Be ${openModeNormal}
    }

    It "保存した設定をそのまま読み込める（1件だけの一覧も配列のまま）" {
        $path = "$TestDrive\往復 [1]\setting.config"
        $settings = newSettings
        $settings.targetFolders = @([pscustomobject]@{ path = "C:\a [1]"; enabled = $true })
        $settings.indexSources = @([pscustomobject]@{ name = "index<1>"; path = "D:\index<1>" })
        $settings.useRegex = $true
        writeSettings $settings $path
        $read = readSettings $path
        @($read.targetFolders).Count | Should Be 1
        $read.targetFolders[0].path | Should Be "C:\a [1]"
        $read.indexSources[0].path | Should Be "D:\index<1>"
        $read.useRegex | Should Be $true
        # 1つの項目だけを変えると、ほかの項目は保つ
        updateSettings "useRegex" $false $path
        $read = readSettings $path
        $read.useRegex | Should Be $false
        $read.indexSources[0].path | Should Be "D:\index<1>"
    }

    It "記載の無い項目は既定値、空のファイルは既定値" {
        $path = "$TestDrive\一部.json"
        [System.IO.File]::WriteAllText($path, '{ "useRegex": true }', ${utf8Bom})
        $settings = readSettings $path
        $settings.useRegex | Should Be $true
        @($settings.targetFolders).Count | Should Be 0
        [System.IO.File]::WriteAllText($path, "", ${utf8Bom})
        (readSettings $path).useRegex | Should Be $false
    }

    It "JSON として読めなければ例外を投げる" {
        $path = "$TestDrive\壊れ.json"
        [System.IO.File]::WriteAllText($path, "{ targetFolders: ", ${utf8Bom})
        { readSettings $path } | Should Throw "読み込めません"
    }

    It "設定ファイルが無く config フォルダに以前の設定ファイル（*.txt）があれば、移して保存する" {
        $dir = "$TestDrive\以前"
        writeListFile "$dir\config\変換対象フォルダパス.txt" @("C:\データ\Excel\", "", "# D:\old\報告書")
        writeListFile "$dir\config\検索オプション.txt" @("正規表現=オン")
        $settings = readSettings "$dir\setting.config"
        Test-Path -LiteralPath "$dir\setting.config" | Should Be $true
        $folders = @(getTargetFolders "$dir\setting.config")
        $folders.Count | Should Be 2
        $folders[0].Path | Should Be "C:\データ\Excel"
        $folders[0].Enabled | Should Be $true
        $folders[1].Path | Should Be "D:\old\報告書"
        $folders[1].Enabled | Should Be $false
        (readSearchOption "$dir\setting.config").UseRegex | Should Be $true
        # 移した後は以前の設定ファイルを使わない
        writeListFile "$dir\config\検索オプション.txt" @("正規表現=オフ")
        (readSearchOption "$dir\setting.config").UseRegex | Should Be $true
    }
}

Describe "getTargetFolders / writeTargetFolders" -Tag Io {
    It "記載順に返し、enabled が false はチェックなし、記載が無ければチェックあり" {
        $path = "$TestDrive\targets.json"
        [System.IO.File]::WriteAllText($path, @'
{ "targetFolders": [
    { "path": "C:\\データ\\Excel", "enabled": true },
    { "path": "D:\\old\\報告書", "enabled": false },
    { "path": "\"F:\\引用符付き\\\"" },
    { "path": "  " }
] }
'@, ${utf8Bom})
        $folders = @(getTargetFolders $path)
        $folders.Count | Should Be 3
        $folders[0].Path | Should Be "C:\データ\Excel"
        $folders[0].Enabled | Should Be $true
        $folders[1].Path | Should Be "D:\old\報告書"
        $folders[1].Enabled | Should Be $false
        $folders[2].Path | Should Be "F:\引用符付き"
        $folders[2].Enabled | Should Be $true
    }

    It "同じフォルダ（大文字・小文字、末尾の \ の違い）は最初のものだけ使う" {
        $path = "$TestDrive\targets_dup.json"
        writeTargetFolders @(
            [pscustomobject]@{ Path = "C:\Data"; Enabled = $true },
            [pscustomobject]@{ Path = "c:\data\"; Enabled = $false }
        ) $path
        $folders = @(getTargetFolders $path)
        $folders.Count | Should Be 1
        $folders[0].Enabled | Should Be $true
    }

    It "設定が無ければ空の配列を返す" {
        @(getTargetFolders "$TestDrive\none_targets.json").Count | Should Be 0
    }

    It "保存した一覧をそのまま読み込め、ほかの設定は保つ" {
        $path = "$TestDrive\targets_write.json"
        writeSearchOption @{ UseRegex = $true } $path
        writeTargetFolders @(
            [pscustomobject]@{ Path = "C:\a [1]"; Enabled = $true },
            [pscustomobject]@{ Path = "D:\b"; Enabled = $false }
        ) $path
        $folders = @(getTargetFolders $path)
        $folders.Count | Should Be 2
        $folders[0].Path | Should Be "C:\a [1]"
        $folders[1].Path | Should Be "D:\b"
        $folders[1].Enabled | Should Be $false
        (readSearchOption $path).UseRegex | Should Be $true
        writeTargetFolders @() $path
        @(getTargetFolders $path).Count | Should Be 0
    }
}

Describe "indexSources / setIndexSourceFolder" -Tag Io {
    It "インデックス名に対する元のフォルダを記録し、読み込める" {
        $path = "$TestDrive\sources[1].config"
        setIndexSourceFolder "営業" "\server\営業\" $path
        setIndexSourceFolder "見積" "E:\見積" $path
        $sources = @(readIndexSources $path)
        $sources.Count | Should Be 2
        $sources[0].Name | Should Be "営業"
        $sources[0].Path | Should Be "\server\営業"
        $sources[1].Path | Should Be "E:\見積"

        # 同じ名前をもう一度記録すると、場所を書き換える（1 つの名前につき 1 か所）
        setIndexSourceFolder "営業" "D:\新しい営業" $path
        $sources = @(readIndexSources $path)
        $sources.Count | Should Be 2
        @($sources | Where-Object { $_.Name -eq "営業" })[0].Path | Should Be "D:\新しい営業"
    }

    It "変換対象フォルダにある名前なら、そのフォルダの場所を書き換える" {
        $path = "$TestDrive\sources_target.config"
        writeTargetFolders @(
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $false }
        ) $path
        setIndexSourceFolder "見積" "\server\移動先\見積" $path

        $folders = @(getTargetFolders $path)
        $folders[0].Name | Should Be "見積"
        $folders[0].Path | Should Be "\server\移動先\見積"
        $folders[0].Enabled | Should Be $true
        $folders[1].Path | Should Be "C:\data\営業"
        # 変換対象フォルダを書き換えたので、indexSources には入れない
        @(readIndexSources $path).Count | Should Be 0
    }

    It "名前・フォルダが空なら何もしない" {
        $path = "$TestDrive\sources_empty.config"
        setIndexSourceFolder "" "C:\a" $path
        setIndexSourceFolder "営業" "  " $path
        @(readIndexSources $path).Count | Should Be 0
    }
}

Describe "readSearchExcludes / writeSearchExcludes" -Tag Io {
    It "設定が無ければ空（すべて検索する）" {
        @(readSearchExcludes "$TestDrive\none_excludes.json").Count | Should Be 0
    }

    It "チェックを外したフォルダを保存し、ほかの設定は変えない" {
        $path = "$TestDrive\excludes.json"
        writeSearchOption @{ UseRegex = $true } $path
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "D:\index1\見積\2024\"; Subfolders = $true },
            [pscustomobject]@{ Path = "D:\index1\見積"; Subfolders = $false },
            [pscustomobject]@{ Path = ""; Subfolders = $true }) $path
        $excludes = @(readSearchExcludes $path)
        $excludes.Count | Should Be 2
        $excludes[0].Path | Should Be "D:\index1\見積\2024"
        $excludes[0].Subfolders | Should Be $true
        $excludes[1].Subfolders | Should Be $false
        (readSearchOption $path).UseRegex | Should Be $true
    }

    It "空で保存すると空になる" {
        $path = "$TestDrive\excludes_empty.json"
        writeSearchExcludes @([pscustomobject]@{ Path = "D:\a"; Subfolders = $true }) $path
        writeSearchExcludes @() $path
        @(readSearchExcludes $path).Count | Should Be 0
    }
}

Describe "readSearchOption / writeSearchOption" -Tag Io {
    $path = Join-Path $TestDrive "setting.config"

    It "ファイルが無ければ、文字どおり・大文字と小文字を区別しない・対象ファイルはすべて" {
        $option = readSearchOption $path
        $option.UseRegex | Should Be $false
        $option.CaseSensitive | Should Be $false
        $option.FileFilter | Should Be ""
        # 図形・コメントも検索する
        $option.IncludeShapes | Should Be $true
        $option.IncludeComments | Should Be $true
    }

    It "図形・コメントを検索するかを保存・読み込みできる" {
        $optionPath = Join-Path $TestDrive "setting_object.config"
        writeSearchOption @{ IncludeShapes = $false } $optionPath
        $option = readSearchOption $optionPath
        $option.IncludeShapes | Should Be $false
        $option.IncludeComments | Should Be $true
    }

    It "保存した値を読み込む" {
        writeSearchOption @{ UseRegex = $true } $path
        (readSearchOption $path).UseRegex | Should Be $true
        writeSearchOption @{ UseRegex = $false } $path
        (readSearchOption $path).UseRegex | Should Be $false
    }

    It "指定した項目だけを変え、ほかの項目は保つ" {
        writeSearchOption @{ UseRegex = $true; CaseSensitive = $true; FileFilter = "*.xlsx;!*old*" } $path
        writeSearchOption @{ CaseSensitive = $false } $path
        $option = readSearchOption $path
        $option.UseRegex | Should Be $true
        $option.CaseSensitive | Should Be $false
        $option.FileFilter | Should Be "*.xlsx;!*old*"
    }
}
