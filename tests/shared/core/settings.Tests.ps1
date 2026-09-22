# 設定ファイルの共通の読み書き（shared\core\settings.ps1）のテスト。
# 1 つの setting.config を検索（tebunko_grep）と比較（tebunko_diff）で使うため、書き戻しでほかのツールのキーが消えないこと
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "writeSettingsFile" -Tag Io {
    It "書く側が知らないキーを残す" {
        $path = "$TestDrive\keep\setting.config"
        [System.IO.Directory]::CreateDirectory((Split-Path $path)) | Out-Null
        [System.IO.File]::WriteAllText($path, '{ "useRegex": true, "diffMode": "folder", "diffOptions": { "ignoreWhitespace": true } }')

        writeSettingsFile ([ordered]@{ useRegex = $false; caseSensitive = $true }) $path

        $data = readSettingsData $path
        $data.useRegex | Should Be $false
        $data.caseSensitive | Should Be $true
        $data.diffMode | Should Be "folder"
        $data.diffOptions.ignoreWhitespace | Should Be $true
    }

    It "検索の設定を保存しても、比較のキーが消えない（writeSettings）" {
        $path = "$TestDrive\grep\setting.config"
        [System.IO.Directory]::CreateDirectory((Split-Path $path)) | Out-Null
        [System.IO.File]::WriteAllText($path, '{ "diffFileLeft": "C:\\共有\\a.xlsx" }')

        updateSettings "useRegex" $true $path

        $data = readSettingsData $path
        $data.useRegex | Should Be $true
        $data.diffFileLeft | Should Be "C:\共有\a.xlsx"
    }

    It "今のファイルが JSON として読めなければ、書く内容だけで書き直す" {
        $path = "$TestDrive\broken\setting.config"
        [System.IO.Directory]::CreateDirectory((Split-Path $path)) | Out-Null
        [System.IO.File]::WriteAllText($path, '{ 壊れている')

        writeSettingsFile ([ordered]@{ useRegex = $true }) $path

        (readSettingsData $path).useRegex | Should Be $true
    }
}

Describe "readSettingsData" -Tag Io {
    It "ファイルが無い・空・null なら null" {
        readSettingsData "$TestDrive\無い.config" | Should BeNullOrEmpty
        $empty = "$TestDrive\empty.config"
        [System.IO.File]::WriteAllText($empty, "  ")
        readSettingsData $empty | Should BeNullOrEmpty
        $nullFile = "$TestDrive\null.config"
        [System.IO.File]::WriteAllText($nullFile, "null")
        readSettingsData $nullFile | Should BeNullOrEmpty
    }

    It "JSON として読めなければ、ファイル名の入ったメッセージで例外にする" {
        $path = "$TestDrive\bad.config"
        [System.IO.File]::WriteAllText($path, "{")
        { readSettingsData $path } | Should Throw "bad.config を読み込めません"
    }
}

Describe "mergeSettingValues" -Tag Unit {
    It "既定値と同じ型にそろえる（文字列・真偽・数値・配列）" {
        $data = '{ "text": 12, "flag": "false", "count": "300", "list": [1, null, 2] }' | ConvertFrom-Json
        $settings = mergeSettingValues ([ordered]@{ text = ""; flag = $true; count = 260; list = @() }) $data
        $settings.text | Should Be "12"
        $settings.flag | Should Be $false
        $settings.count | Should Be 300
        @($settings.list).Count | Should Be 2
    }

    It "読めない値は既定値のまま" {
        $data = '{ "flag": "はい", "count": "たくさん" }' | ConvertFrom-Json
        $settings = mergeSettingValues ([ordered]@{ flag = $true; count = 260 }) $data
        $settings.flag | Should Be $true
        $settings.count | Should Be 260
    }

    It "入れ子の設定は、既定値のキーだけを読み、無いキーは既定値にする" {
        $data = '{ "options": { "a": false, "unknown": 1 } }' | ConvertFrom-Json
        $settings = mergeSettingValues ([ordered]@{ options = [ordered]@{ a = $true; b = $true } }) $data
        $settings.options.a | Should Be $false
        $settings.options.b | Should Be $true
        $settings.options.Contains("unknown") | Should Be $false
    }

    It "中身が null なら既定値のまま返す" {
        $settings = mergeSettingValues ([ordered]@{ flag = $true }) $null
        $settings.flag | Should Be $true
    }
}
