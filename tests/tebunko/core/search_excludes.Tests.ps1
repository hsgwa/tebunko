# 検索対象ツリーでチェックを外したフォルダ（tebunko\core\search_excludes.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "testFolderOrUnder（フォルダとその下の比較。ここだけで行う）" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "同じフォルダ"; path = "D:\index\A"; folder = "D:\index\A"; expected = $true }
        @{ name = "フォルダの下"; path = "D:\index\A\見積"; folder = "D:\index\A"; expected = $true }
        @{ name = "A と AB は別（頭が同じ名前は下に入らない）"; path = "D:\index\AB"; folder = "D:\index\A"; expected = $false }
        @{ name = "A と AB の下も別"; path = "D:\index\AB\見積"; folder = "D:\index\A"; expected = $false }
        @{ name = "親のフォルダは含まれない"; path = "D:\index"; folder = "D:\index\A"; expected = $false }
        @{ name = "大文字・小文字は区別しない（同じフォルダ）"; path = "D:\INDEX\a"; folder = "d:\index\A"; expected = $true }
        @{ name = "大文字・小文字は区別しない（下）"; path = "D:\INDEX\a\見積"; folder = "d:\index\A"; expected = $true }
        @{ name = "folder の末尾の \ は無いものとして比べる（同じフォルダ）"; path = "D:\index\A"; folder = "D:\index\A\"; expected = $true }
        @{ name = "folder の末尾の \ は無いものとして比べる（下）"; path = "D:\index\A\見積"; folder = "D:\index\A\"; expected = $true }
        @{ name = "folder の末尾の \ があっても A と AB は別"; path = "D:\index\AB"; folder = "D:\index\A\"; expected = $false }
        @{ name = "\\?\ 付きと無しは別の書き方として比べる（付いているものどうしなら下に入る）"; path = "\\?\D:\index\A\見積"; folder = "\\?\D:\index\A"; expected = $true }
        @{ name = "\\?\ 付きの path は、付いていない folder の下に入らない"; path = "\\?\D:\index\A\見積"; folder = "D:\index\A"; expected = $false }
        @{ name = "folder が空ならどれも含まない"; path = "D:\index\A"; folder = ""; expected = $false }
        @{ name = "folder が \ だけでも、どれも含まない"; path = "D:\index\A"; folder = "\"; expected = $false }
        @{ name = "ß と ss は別のフォルダ（序数の比較。文化による同一視をしない）"; path = "D:\index\Straße"; folder = "D:\index\Strasse"; expected = $false }
        @{ name = "ß と ss は別のフォルダ（下）"; path = "D:\index\Straße\見積"; folder = "D:\index\Strasse"; expected = $false }
    ) {
        param ($name, $path, $folder, $expected)
        testFolderOrUnder $path $folder | Should -Be $expected
    }

    It "path が null なら含まない" {
        testFolderOrUnder $null "D:\index\A" | Should -Be $false
    }
}
Describe "readSearchExcludes / writeSearchExcludes" -Tag Io {
    It "チェックを外したフォルダを保存し、ほかの設定は変えない" {
        $path = "$TestDrive\excludes.json"
        writeSearchOption @{ UseRegex = $true } $path
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "D:\index1\見積\2024\"; Subfolders = $true },
            [pscustomobject]@{ Path = "D:\index1\見積"; Subfolders = $false },
            [pscustomobject]@{ Path = ""; Subfolders = $true }) $path
        $excludes = @(readSearchExcludes $path)
        $excludes.Count | Should -Be 2
        $excludes[0].Path | Should -Be "D:\index1\見積\2024"
        $excludes[0].Subfolders | Should -Be $true
        $excludes[1].Subfolders | Should -Be $false
        (readSearchOption $path).UseRegex | Should -Be $true
    }

    It "subfolders の記載が無ければフォルダ以下すべて、文字列の ""false"" は直下のファイルだけとして読む" {
        $path = "$TestDrive\excludes_hand.json"
        [System.IO.File]::WriteAllText($path, '{ "searchExcludes": [ { "path": "D:\\a" }, { "path": "D:\\b", "subfolders": "false" }, { "path": "  " } ] }', ${utf8Bom})
        $excludes = @(readSearchExcludes $path)
        $excludes.Count | Should -Be 2
        $excludes[0].Subfolders | Should -Be $true
        $excludes[1].Subfolders | Should -Be $false
    }

    It "設定が無い・空で保存したときは空（すべて検索する）" {
        @(readSearchExcludes "$TestDrive\none_excludes.json").Count | Should -Be 0
        $path = "$TestDrive\excludes_empty.json"
        writeSearchExcludes @([pscustomobject]@{ Path = "D:\a"; Subfolders = $true }) $path
        writeSearchExcludes @() $path
        @(readSearchExcludes $path).Count | Should -Be 0
    }
}

Describe "removeSearchExcludesUnder" -Tag Io {
    It "フォルダとその下だけを、大文字・小文字を区別せずに消す（頭が同じ名前は消さない）" {
        $path = "$TestDrive\under.json"
        writeSearchOption @{ UseRegex = $true } $path
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "D:\index\Sales"; Subfolders = $false },
            [pscustomobject]@{ Path = "D:\INDEX\sales\見積"; Subfolders = $true },
            [pscustomobject]@{ Path = "D:\index\Sales2"; Subfolders = $true },
            [pscustomobject]@{ Path = "D:\index\技術"; Subfolders = $true }) $path

        removeSearchExcludesUnder "D:\index\Sales\" $path

        @(readSearchExcludes $path | ForEach-Object { $_.Path }) | Should -Be @("D:\index\Sales2", "D:\index\技術")
        (readSearchOption $path).UseRegex | Should -Be $true
    }

    It "消す記録が無ければ設定ファイルを書き換えない・作らない" {
        $path = "$TestDrive\under_none.json"
        removeSearchExcludesUnder "D:\index\Sales" $path
        Test-Path -LiteralPath $path | Should -Be $false
        writeSearchExcludes @([pscustomobject]@{ Path = "D:\index\技術"; Subfolders = $true }) $path
        $before = (Get-Item -LiteralPath $path).LastWriteTimeUtc
        Start-Sleep -Milliseconds 50
        removeSearchExcludesUnder "D:\index\Sales" $path
        (Get-Item -LiteralPath $path).LastWriteTimeUtc | Should -Be $before
    }

    It "空のフォルダ名・読み書きの失敗では例外にしない" {
        { removeSearchExcludesUnder "" "$TestDrive\x.json" } | Should -Not -Throw
        New-Item -ItemType Directory -Path "$TestDrive\dirsetting" -Force | Out-Null
        { removeSearchExcludesUnder "D:\a" "$TestDrive\dirsetting" } | Should -Not -Throw
    }
}

Describe "moveSearchExcludes" -Tag Io {
    It "前のワークスペースのインデックスの下を、移した先のインデックスの下に付け替える（ほかは変えない）" {
        $settings = "$TestDrive\excludes\setting.config"
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "C:\old\content_index"; Subfolders = $true },
            [pscustomobject]@{ Path = "C:\OLD\content_index\営業\見積"; Subfolders = $false },
            [pscustomobject]@{ Path = "C:\old\content_index2\技術"; Subfolders = $true }
        ) $settings

        moveSearchExcludes "C:\old" "D:\新しい場所" $settings | Should -Be 2

        $excludes = @(readSearchExcludes $settings)
        @($excludes | ForEach-Object { $_.Path }) -join "|" | Should -Be "D:\新しい場所\content_index|D:\新しい場所\content_index\営業\見積|C:\old\content_index2\技術"
        @($excludes | ForEach-Object { $_.Subfolders }) -join "," | Should -Be "True,False,True"
    }
}