# 比較の設定（tebunko_diff\core\settings_diff.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"

Describe "readDiffSettings / updateDiffSettings" -Tag Io {
    It "ファイルが無ければ既定値" {
        $settings = readDiffSettings "$TestDrive\無い\setting.config"
        $settings.diffMode | Should Be "file"
        $settings.diffSubfolders | Should Be $true
        $settings.diffOptions.includeShapes | Should Be $true
        $settings.diffTreeHeight | Should Be 260
    }

    It "比較の設定を保存しても、検索の設定を残す" {
        $path = "$TestDrive\keep\setting.config"
        writeSettings (newSettings) $path
        updateSettings "useRegex" $true $path
        updateDiffSettings ([ordered]@{ diffMode = "folder"; diffFolderLeft = "C:\共有\2024" }) $path

        $diff = readDiffSettings $path
        $diff.diffMode | Should Be "folder"
        $diff.diffFolderLeft | Should Be "C:\共有\2024"
        (readSettings $path).useRegex | Should Be $true
    }

    It "比べ方（入れ子）を保存して読み直せる" {
        $path = "$TestDrive\options\setting.config"
        updateDiffSettings ([ordered]@{ diffOptions = [ordered]@{ includeShapes = $false; includeComments = $true; ignoreWhitespace = $true; caseSensitive = $true } }) $path
        $options = (readDiffSettings $path).diffOptions
        $options.includeShapes | Should Be $false
        $options.ignoreWhitespace | Should Be $true
    }

    It "知らない値は既定値にし、ツリーの高さは範囲に収める" {
        $path = "$TestDrive\bad\setting.config"
        [System.IO.Directory]::CreateDirectory((Split-Path $path)) | Out-Null
        [System.IO.File]::WriteAllText($path, '{ "diffMode": "何か", "diffView": "x", "diffTreeHeight": 99999 }')
        $settings = readDiffSettings $path
        $settings.diffMode | Should Be "file"
        $settings.diffView | Should Be "side"
        $settings.diffTreeHeight | Should Be 2000
    }

    It "比較の設定に無いキーは保存しない" {
        { updateDiffSettings ([ordered]@{ useRegex = $true }) "$TestDrive\x\setting.config" } | Should Throw "比較の設定に無いキー"
    }
}

Describe "getDiffOptions" -Tag Unit {
    It "設定の小文字始まりのキーも、大文字始まりのキーも読む" {
        $options = getDiffOptions ([ordered]@{ includeShapes = $false; IgnoreWhitespace = $true })
        $options.IncludeShapes | Should Be $false
        $options.IgnoreWhitespace | Should Be $true
        $options.CaseSensitive | Should Be $true
    }
}
