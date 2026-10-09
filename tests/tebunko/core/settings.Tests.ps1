# 設定ファイル（tebunko\core\settings.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "readSettings / writeSettings" -Tag Io {
    It "ファイルが無ければ既定値を返し、ファイルを作らない" {
        $path = "$TestDrive\空\setting.config"
        $settings = readSettings $path
        @($settings.targetFolders).Count | Should -Be 0
        @($settings.indexSources).Count | Should -Be 0
        $settings.useRegex | Should -Be $false
        $settings.openMode | Should -Be "normal"
        Test-Path -LiteralPath $path | Should -Be $false
    }

    It "開き方（openMode）を保存・読み込みできる" {
        $path = "$TestDrive\開き方\setting.config"
        foreach ($mode in ${openModes}) {
            writeOpenMode $mode $path
            readOpenMode $path | Should -Be $mode
        }
    }

    It "取り込みのスレッドの数（ingestThreads）は数値で読む。数値にできなければ既定値（0）" {
        $path = "$TestDrive\スレッド\setting.config"
        (readSettings $path).ingestThreads | Should -Be 0
        updateSettings "ingestThreads" 2 $path
        (readSettings $path).ingestThreads | Should -Be 2
        updateSettings "ingestThreads" "3" $path   # 手で書いた文字列
        (readSettings $path).ingestThreads | Should -Be 3
        updateSettings "ingestThreads" "たくさん" $path
        (readSettings $path).ingestThreads | Should -Be 0
    }

    It "開き方が無い・知らない値なら「通常」とする" {
        $path = "$TestDrive\開き方2\setting.config"
        readOpenMode $path | Should -Be ${openModeNormal}        # ファイルが無い
        updateSettings "openMode" "知らない値" $path
        readOpenMode $path | Should -Be ${openModeNormal}
    }

    It "保存した設定をそのまま読み込める（1件だけの一覧も配列のまま）" {
        $path = "$TestDrive\往復 [1]\setting.config"
        $settings = newSettings
        $settings.targetFolders = @([pscustomobject]@{ path = "C:\a [1]"; enabled = $true })
        $settings.indexSources = @([pscustomobject]@{ name = "index<1>"; path = "D:\index<1>" })
        $settings.useRegex = $true
        writeSettings $settings $path
        $read = readSettings $path
        @($read.targetFolders).Count | Should -Be 1
        $read.targetFolders[0].path | Should -Be "C:\a [1]"
        $read.indexSources[0].path | Should -Be "D:\index<1>"
        $read.useRegex | Should -Be $true
        # 1つの項目だけを変えると、ほかの項目は保つ
        updateSettings "useRegex" $false $path
        $read = readSettings $path
        $read.useRegex | Should -Be $false
        $read.indexSources[0].path | Should -Be "D:\index<1>"
    }

    It "記載の無い項目は既定値、空のファイルは既定値" {
        $path = "$TestDrive\一部.json"
        [System.IO.File]::WriteAllText($path, '{ "useRegex": true }', ${utf8Bom})
        $settings = readSettings $path
        $settings.useRegex | Should -Be $true
        @($settings.targetFolders).Count | Should -Be 0
        [System.IO.File]::WriteAllText($path, "", ${utf8Bom})
        (readSettings $path).useRegex | Should -Be $false
    }

    It "中身が JSON の null・配列・数値だけなら既定値（空のファイルと同じ）" {
        $path = "$TestDrive\値だけ.json"
        foreach ($json in @("null", " null ", "[]", "123", '"x"')) {
            [System.IO.File]::WriteAllText($path, $json, ${utf8Bom})
            $settings = readSettings $path
            $settings.useRegex | Should -Be $false
            @($settings.targetFolders).Count | Should -Be 0
            $settings.openMode | Should -Be ${openModeNormal}
        }
    }

    It "手で書いた真偽値の文字列（""false"" 等）は真偽値として読み、読めない値は既定値のまま" {
        $path = "$TestDrive\文字列の真偽値.json"
        [System.IO.File]::WriteAllText($path, '{ "useRegex": "false", "caseSensitive": "True", "includeShapes": "いいえ", "includeComments": 0 }', ${utf8Bom})
        $settings = readSettings $path
        $settings.useRegex | Should -Be $false
        $settings.caseSensitive | Should -Be $true
        $settings.includeShapes | Should -Be $true    # 読めない値は既定値
        $settings.includeComments | Should -Be $false # 数値の 0 は偽
    }

    It "文字列の項目に数値が書かれていても文字列として読む。一覧に 1 件だけ書かれていても配列にする" {
        $path = "$TestDrive\型違い.json"
        [System.IO.File]::WriteAllText($path, '{ "workspaceFolder": 123, "targetFolders": { "path": "C:\\a" }, "indexSources": null }', ${utf8Bom})
        $settings = readSettings $path
        $settings.workspaceFolder | Should -BeExactly "123"
        @($settings.targetFolders).Count | Should -Be 1
        @($settings.indexSources).Count | Should -Be 0
    }

    It "JSON として読めなければ「壊れている」と分かる例外（FormatException）を投げ、ファイルは動かさない" {
        $path = "$TestDrive\壊れ.json"
        [System.IO.File]::WriteAllText($path, "{ targetFolders: ", ${utf8Bom})
        { readSettings $path } | Should -Throw -ExpectedMessage "*読み込めません*" -ExceptionType ([System.FormatException])
        [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | Should -Be "{ targetFolders: "
        @(Get-ChildItem -LiteralPath $TestDrive -Filter "壊れ.json*").Count | Should -Be 1
    }

    It "ほかのプログラムが開いていて読めないときは、壊れているとはせず IOException のまま" {
        $path = "$TestDrive\ロック中.json"
        [System.IO.File]::WriteAllText($path, '{ "workspaceFolder": "a" }', ${utf8Bom})
        $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            # PowerShell は .NET のメソッドの例外を MethodInvocationException に包むため、中身の型を確かめる
            foreach ($action in @({ readSettings $path }, { repairBrokenSettings $path })) {
                $thrown = $null
                try { & $action } catch { $thrown = $_.Exception }
                $thrown | Should -Not -BeNullOrEmpty
                $thrown | Should -Not -BeOfType ([System.FormatException])
                $thrown.InnerException | Should -BeOfType ([System.IO.IOException])
            }
        } finally {
            $stream.Dispose()
        }
        @(Get-ChildItem -LiteralPath $TestDrive -Filter "ロック中.json*").Count | Should -Be 1
    }
}

Describe "repairBrokenSettings / getSettingsRecoveryMessage" -Tag Io {
    It "壊れたファイルを中身のまま setting.config.broken-* に移してパスを返し、移した後は既定値で読める" {
        $dir = "$TestDrive\壊れた設定"
        New-Item -ItemType Directory -Path $dir | Out-Null
        $path = "$dir\setting.config"
        [System.IO.File]::WriteAllText($path, '{ "targetFolders": [', ${utf8Bom})
        $broken = repairBrokenSettings $path
        $broken | Should -BeLike "$dir\setting.config.broken-*"
        Test-Path -LiteralPath $path | Should -Be $false
        [System.IO.File]::ReadAllText($broken, [System.Text.Encoding]::UTF8) | Should -Be '{ "targetFolders": ['
        @((readSettings $path).targetFolders).Count | Should -Be 0
    }

    It "退避の後、getWorkDir は既定のワークスペースを返す（本物の setting.config には触れない）" {
        $dir = "$TestDrive\退避後のワークスペース"
        New-Item -ItemType Directory -Path $dir | Out-Null
        $path = "$dir\setting.config"
        [System.IO.File]::WriteAllText($path, '{ "workspaceFolder": ', ${utf8Bom})
        { getWorkDir $path } | Should -Throw -ExceptionType ([System.FormatException])
        repairBrokenSettings $path | Should -Not -BeNullOrEmpty
        getWorkDir $path | Should -Be (getDefaultWorkDir)
    }

    It "<name> のファイルは何もせず空文字を返す" -TestCases @(
        @{ name = "壊れていない"; content = '{ "useRegex": true }' }
        @{ name = "空"; content = "" }
        @{ name = "空白だけ"; content = "  `r`n" }
        @{ name = "null だけ"; content = "null" }
    ) {
        param ($name, $content)
        $path = "$TestDrive\何もしない-$name.config"
        [System.IO.File]::WriteAllText($path, $content, ${utf8Bom})
        repairBrokenSettings $path | Should -BeNullOrEmpty
        [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | Should -Be $content
        @(Get-ChildItem -LiteralPath $TestDrive -Filter "何もしない-$name.config*").Count | Should -Be 1
    }

    It "ファイルが無ければ何もせず空文字を返す" {
        repairBrokenSettings "$TestDrive\無い\setting.config" | Should -BeNullOrEmpty
    }

    It "同じ秒に 2 回壊れても、前の退避を上書きしない（連番を付ける）" {
        $dir = "$TestDrive\連続退避"
        New-Item -ItemType Directory -Path $dir | Out-Null
        $path = "$dir\setting.config"
        $brokenPaths = @()
        foreach ($content in @("{ 1", "{ 2", "{ 3")) {
            [System.IO.File]::WriteAllText($path, $content, ${utf8Bom})
            $brokenPaths += repairBrokenSettings $path
        }
        @($brokenPaths | Select-Object -Unique).Count | Should -Be 3
        @($brokenPaths | ForEach-Object { [System.IO.File]::ReadAllText($_, [System.Text.Encoding]::UTF8) } | Sort-Object) | Should -Be @("{ 1", "{ 2", "{ 3")
    }

    It "移せなかったときは例外にし、壊れたファイルはそのまま残す" {
        $dir = "$TestDrive\移せない"
        New-Item -ItemType Directory -Path $dir | Out-Null
        $path = "$dir\setting.config"
        [System.IO.File]::WriteAllText($path, "{ 壊れ", ${utf8Bom})
        # 読めるが、移せない（削除を許さない共有で開いている）状態にする
        $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        try {
            { repairBrokenSettings $path } | Should -Throw -ExpectedMessage "*移せませんでした*"
        } finally {
            $stream.Dispose()
        }
        [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | Should -Be "{ 壊れ"
    }

    It "getSettingsRecoveryMessage は、既定に戻ったことと退避したパスを含む" {
        $message = getSettingsRecoveryMessage "C:\Users\test\setting.config.broken-20260101-000000"
        $message | Should -BeLike "設定ファイルが壊れていたため、既定の設定で起動しました。*"
        $message | Should -BeLike "*ワークスペースの場所・登録したフォルダなども既定に戻っています。*"
        $message | Should -BeLike "*tebunko を閉じてから、退避したファイル（C:\Users\test\setting.config.broken-20260101-000000）を直して setting.config に置き換え、開き直してください。"
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
        $folders.Count | Should -Be 3
        $folders[0].Path | Should -Be "C:\データ\Excel"
        $folders[0].Enabled | Should -Be $true
        $folders[1].Path | Should -Be "D:\old\報告書"
        $folders[1].Enabled | Should -Be $false
        $folders[2].Path | Should -Be "F:\引用符付き"
        $folders[2].Enabled | Should -Be $true
    }

    It "同じフォルダ（大文字・小文字、末尾の \ の違い）は最初のものだけ使う" {
        $path = "$TestDrive\targets_dup.json"
        writeTargetFolders @(
            [pscustomobject]@{ Path = "C:\Data"; Enabled = $true },
            [pscustomobject]@{ Path = "c:\data\"; Enabled = $false }
        ) $path
        $folders = @(getTargetFolders $path)
        $folders.Count | Should -Be 1
        $folders[0].Enabled | Should -Be $true
    }

    It "インデックス名が重なれば（大文字・小文字の違いも）2 つ目以降の名前を空にする（取り込み時に割り当て直す）" {
        $path = "$TestDrive\targets_names.json"
        [System.IO.File]::WriteAllText($path, @'
{ "targetFolders": [
    null,
    { "path": "C:\\a", "name": "見積" },
    { "path": "C:\\b", "name": " 見積 " },
    { "path": "C:\\c", "name": "Sales" },
    { "path": "C:\\d", "name": "sales" },
    { "path": "C:\\e", "name": "a/b" }
] }
'@, ${utf8Bom})
        $folders = @(getTargetFolders $path)
        ($folders | ForEach-Object { "$($_.Path)=$($_.Name)" }) -join "," |
            Should -Be "C:\a=見積,C:\b=,C:\c=Sales,C:\d=,C:\e=a／b"
    }

    It "enabled に文字列の ""false"" が書かれていてもチェックなしにする" {
        $path = "$TestDrive\targets_enabled.json"
        [System.IO.File]::WriteAllText($path, '{ "targetFolders": [ { "path": "C:\\a", "enabled": "false" } ] }', ${utf8Bom})
        @(getTargetFolders $path)[0].Enabled | Should -Be $false
    }

    It "保存した一覧をそのまま読み込め、ほかの設定は保つ" {
        $path = "$TestDrive\targets_write.json"
        writeSearchOption @{ UseRegex = $true } $path
        writeTargetFolders @(
            [pscustomobject]@{ Path = "C:\a [1]"; Enabled = $true },
            [pscustomobject]@{ Path = "D:\b"; Enabled = $false }
        ) $path
        $folders = @(getTargetFolders $path)
        $folders.Count | Should -Be 2
        $folders[0].Path | Should -Be "C:\a [1]"
        $folders[1].Path | Should -Be "D:\b"
        $folders[1].Enabled | Should -Be $false
        (readSearchOption $path).UseRegex | Should -Be $true
        writeTargetFolders @() $path
        @(getTargetFolders $path).Count | Should -Be 0
    }
}

Describe "indexSources / setIndexSourceFolder" -Tag Io {
    It "インデックス名に対する元のフォルダを記録し、読み込める" {
        $path = "$TestDrive\sources[1].config"
        setIndexSourceFolder "営業" "\server\営業\" $path
        setIndexSourceFolder "見積" "E:\見積" $path
        $sources = @(readIndexSources $path)
        $sources.Count | Should -Be 2
        $sources[0].Name | Should -Be "営業"
        $sources[0].Path | Should -Be "\server\営業"
        $sources[1].Path | Should -Be "E:\見積"

        # 同じ名前をもう一度記録すると、場所を書き換える（1 つの名前につき 1 か所）
        setIndexSourceFolder "営業" "D:\新しい営業" $path
        $sources = @(readIndexSources $path)
        $sources.Count | Should -Be 2
        @($sources | Where-Object { $_.Name -eq "営業" })[0].Path | Should -Be "D:\新しい営業"
    }

    It "クロール対象フォルダにある名前なら、そのフォルダの場所を書き換える" {
        $path = "$TestDrive\sources_target.config"
        writeTargetFolders @(
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $false }
        ) $path
        setIndexSourceFolder "見積" "\server\移動先\見積" $path

        $folders = @(getTargetFolders $path)
        $folders[0].Name | Should -Be "見積"
        $folders[0].Path | Should -Be "\server\移動先\見積"
        $folders[0].Enabled | Should -Be $true
        $folders[1].Path | Should -Be "C:\data\営業"
        # クロール対象フォルダを書き換えたので、indexSources には入れない
        @(readIndexSources $path).Count | Should -Be 0
    }

    It "名前・フォルダが空なら何もしない" {
        $path = "$TestDrive\sources_empty.config"
        setIndexSourceFolder "" "C:\a" $path
        setIndexSourceFolder "営業" "  " $path
        @(readIndexSources $path).Count | Should -Be 0
    }

    It "名前・フォルダが空の記録と、同じ名前（大文字・小文字の違いも）の 2 つ目以降は読まない" {
        $path = "$TestDrive\sources_hand.config"
        [System.IO.File]::WriteAllText($path, @'
{ "indexSources": [
    { "name": "", "path": "C:\\a" },
    { "name": "営業", "path": "" },
    { "name": "Sales", "path": "C:\\b\\" },
    { "name": "sales", "path": "C:\\c" }
] }
'@, ${utf8Bom})
        $sources = @(readIndexSources $path)
        $sources.Count | Should -Be 1
        $sources[0].Name | Should -Be "Sales"
        $sources[0].Path | Should -Be "C:\b"
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

Describe "readFileKinds / writeFileKinds" -Tag Io {
    # json: 設定ファイルの中身（$null ならファイルを作らない）、expected: readFileKinds の結果
    It "<name>" -TestCases @(
        @{ name = "ファイルが無ければすべての種類"; json = $null; expected = @("excel", "word", "powerpoint", "text") }
        @{ name = "キーが無ければすべての種類"; json = '{ "useRegex": true }'; expected = @("excel", "word", "powerpoint", "text") }
        @{ name = "空の配列ならすべての種類"; json = '{ "fileKinds": [] }'; expected = @("excel", "word", "powerpoint", "text") }
        @{ name = "excel だけなら Excel だけ"; json = '{ "fileKinds": ["excel"] }'; expected = @("excel") }
        @{ name = "順番は種類の並びにそろえ、大文字小文字は区別しない"; json = '{ "fileKinds": ["Text", "word"] }'; expected = @("word", "text") }
        @{ name = "知らない値は捨てる"; json = '{ "fileKinds": ["pdf", "excel"] }'; expected = @("excel") }
        @{ name = "知らない値だけならすべての種類"; json = '{ "fileKinds": ["pdf"] }'; expected = @("excel", "word", "powerpoint", "text") }
        @{ name = "前の版の fileFilter は種類に読み替えない"; json = '{ "fileFilter": "*.xlsx;!~$*" }'; expected = @("excel", "word", "powerpoint", "text") }
    ) {
        param ($name, $json, $expected)
        $path = Join-Path $TestDrive "kinds_$([guid]::NewGuid()).config"
        if ($null -ne $json) {
            [System.IO.File]::WriteAllText($path, $json, ${utf8Bom})
        }
        (@(readFileKinds $path) -join ",") | Should -Be ($expected -join ",")
    }

    It "書いた種類を読み返せ、ほかの設定は変えない" {
        $path = "$TestDrive\kinds_write.config"
        writeSearchOption @{ UseRegex = $true } $path
        writeFileKinds @("text", "excel", "pdf") $path
        (@(readFileKinds $path) -join ",") | Should -Be "excel,text"
        (readSearchOption $path).UseRegex | Should -Be $true
    }

    It "1 つも選んでいないときは保存しない（前に保存した種類のまま。保存の形は変えない）" {
        $path = "$TestDrive\kinds_none.config"
        writeFileKinds @("word") $path
        writeFileKinds @() $path
        (@(readFileKinds $path) -join ",") | Should -Be "word"
    }

    It "すべての種類を書いたときは、設定ファイルの fileKinds を空にする" {
        $path = "$TestDrive\kinds_all.config"
        writeFileKinds @("excel") $path
        writeFileKinds @("excel", "word", "powerpoint", "text") $path
        @((readSettings $path).fileKinds).Count | Should -Be 0
        @(readFileKinds $path).Count | Should -Be 4
    }
}

Describe "readSearchOption / writeSearchOption" -Tag Io {
    # writes: 順に保存する項目、expected: 読み込んだときの値
    It "<name>" -TestCases @(
        @{ name = "ファイルが無ければ、文字どおり・大文字と小文字を区別しない・図形とコメントも検索する"; writes = @()
           expected = @{ UseRegex = $false; CaseSensitive = $false; IncludeShapes = $true; IncludeComments = $true } }
        @{ name = "図形・コメントを検索するかを保存・読み込みできる"; writes = @(@{ IncludeShapes = $false })
           expected = @{ IncludeShapes = $false; IncludeComments = $true } }
        @{ name = "指定した項目だけを変え、ほかの項目は保つ"; writes = @(@{ UseRegex = $true; CaseSensitive = $true; IncludeShapes = $false }, @{ CaseSensitive = $false })
           expected = @{ UseRegex = $true; CaseSensitive = $false; IncludeShapes = $false } }
    ) {
        param ($name, $writes, $expected)
        $path = Join-Path $TestDrive "setting_$([guid]::NewGuid()).config"
        foreach ($option in $writes) {
            writeSearchOption $option $path
        }
        $read = readSearchOption $path
        foreach ($key in $expected.Keys) {
            $read.$key | Should -Be $expected[$key]
        }
    }
}

Describe "前の版の fileFilter" -Tag Io {
    It "読み込んだ検索オプションに FileFilter は無く、設定を保存し直しても fileFilter は残らない" {
        $path = Join-Path $TestDrive "old_$([guid]::NewGuid()).config"
        [System.IO.File]::WriteAllText($path, '{ "fileFilter": "*.xlsx;!~$*", "useRegex": true }', ${utf8Bom})
        (readSearchOption $path).ContainsKey("FileFilter") | Should -Be $false
        (newSettings).Contains("fileFilter") | Should -Be $false
        writeSearchOption @{ CaseSensitive = $true } $path
        (Get-Content -LiteralPath $path -Raw -Encoding UTF8) | Should -Not -Match "fileFilter"
    }
}

Describe "getWorkDir / writeWorkspaceFolder" -Tag Io {
    It "設定が無ければ、既定の場所（%USERPROFILE%\Documents\tebunko_ws。テストでは環境変数で差し替わる）" {
        getWorkDir "$TestDrive\既定\setting.config" | Should -Be (getDefaultWorkDir)
        getWorkDir "$TestDrive\既定\setting.config" | Should -Be $env:TEBUNKO_DEFAULT_WORKSPACE.TrimEnd("\", "/")
        getDefaultWorkDir ([System.Environment]::GetFolderPath("UserProfile")) | Should -Be (getRealDefaultWorkspace)
    }

    It "保存したフォルダを返し、ほかの設定は保つ" {
        $path = "$TestDrive\別の場所\setting.config"
        writeOpenMode ${openModeReadOnly} $path
        writeWorkspaceFolder "D:\データ\tebunko\" $path
        getWorkDir $path | Should -Be "D:\データ\tebunko"
        readOpenMode $path | Should -Be ${openModeReadOnly}
    }

    It "既定の場所を選んだときは空で保存する（ツールのフォルダを移しても既定のまま付いてくる）" {
        $path = "$TestDrive\既定に戻す\setting.config"
        writeWorkspaceFolder "D:\データ" $path
        writeWorkspaceFolder ((getDefaultWorkDir).ToUpperInvariant()) $path
        (readSettings $path).workspaceFolder | Should -Be ""
        getWorkDir $path | Should -Be (getDefaultWorkDir)
    }

    It "手で書いた相対パスは設定ファイルのフォルダから、環境変数は展開して読む" {
        $path = "$TestDrive\相対\setting.config"
        updateSettings "workspaceFolder" "..\データ" $path
        getWorkDir $path | Should -Be "$TestDrive\データ"
        updateSettings "workspaceFolder" "%TEMP%\tebunko" $path
        getWorkDir $path | Should -Be ([System.IO.Path]::GetFullPath("$env:TEMP\tebunko"))
    }
}

Describe "getSettingsFilePath" -Tag Io {
    It "<Case>" -TestCases @(
        @{ Case = "書き込めるツールのフォルダなら、そこの setting.config"; Writable = $true }
        @{ Case = "書き込めない（存在しない）ツールのフォルダなら、既定のワークスペースの setting.config"; Writable = $false }
    ) {
        $root = "$TestDrive\ツール-$Writable"
        if ($Writable) {
            [System.IO.Directory]::CreateDirectory($root) | Out-Null
        }
        $expected = if ($Writable) { "$root\setting.config" } else { "$TestDrive\既定のワークスペース\setting.config" }
        getSettingsFilePath $root "$TestDrive\既定のワークスペース" | Should -Be $expected
    }

    It "ツールのフォルダに書けず、既定のワークスペースがまだ無いとき、保存でフォルダができ、同じ場所から読める" {
        # 書く先は $TestDrive の下だけ（既定のワークスペースは $TestDrive に向ける。本物の Documents には書かない）
        $toolDir = "$TestDrive\無いツールのフォルダ"
        $defaultWork = "$TestDrive\まだ無い既定のワークスペース"
        $path = getSettingsFilePath $toolDir $defaultWork
        $path.StartsWith($TestDrive, [System.StringComparison]::OrdinalIgnoreCase) | Should -Be $true
        Test-Path -LiteralPath $defaultWork | Should -Be $false

        writeSettings @{ ingestThreads = 3 } $path

        Test-Path -LiteralPath $defaultWork -PathType Container | Should -Be $true
        $path | Should -Be "$defaultWork\setting.config"
        Test-Path -LiteralPath $path -PathType Leaf | Should -Be $true
        Test-Path -LiteralPath $toolDir | Should -Be $false
        (readSettings $path).ingestThreads | Should -Be 3
        # 同じ場所を、もう一度求めても同じ（保存したあとは書き込めるワークスペースになるが、ツールのフォルダは無いまま）
        getSettingsFilePath $toolDir $defaultWork | Should -Be $path
    }
}

Describe "testSettingsFileName / testDefaultWorkspace（設定ファイルとそれに付いてできるファイル）" -Tag Io {
    It "名前の判定: <Name> は <Expected>" -TestCases @(
        @{ Name = "setting.config"; Expected = $true }
        @{ Name = "SETTING.CONFIG"; Expected = $true }
        @{ Name = "setting.config.tmp"; Expected = $true }
        @{ Name = "setting.config.broken-20261003-120000"; Expected = $true }
        @{ Name = "setting.config.broken-20261003-120000-2"; Expected = $true }
        @{ Name = "setting.config.bak"; Expected = $false }
        @{ Name = "my_setting.config"; Expected = $false }
        @{ Name = "setting.config.broken-"; Expected = $false }
        @{ Name = "setting.config.tmp.old"; Expected = $false }
    ) {
        testSettingsFileName $Name | Should -Be $Expected
    }

    It "<Case>" -TestCases @(
        @{ Case = "setting.config だけなら使える"; Names = @("setting.config"); Usable = $true }
        @{ Case = ".tmp と .broken-<日時>（連番付きを含む）が付いていても使える"; Names = @("setting.config", "setting.config.tmp", "setting.config.broken-20261003-120000", "setting.config.broken-20261003-120000-2"); Usable = $true }
        @{ Case = "似た名前（.bak）があれば使えない"; Names = @("setting.config", "setting.config.bak"); Usable = $false }
        @{ Case = "似た名前（my_setting.config）があれば使えない"; Names = @("my_setting.config"); Usable = $false }
        @{ Case = "ほかのファイルが 1 つでもあれば使えない"; Names = @("setting.config", "README.md"); Usable = $false }
    ) {
        $dir = "$TestDrive\既定-" + [guid]::NewGuid().ToString("N")
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        foreach ($name in $Names) {
            [System.IO.File]::WriteAllText("$dir\$name", "")
        }
        (testDefaultWorkspace $dir).Usable | Should -Be $Usable
    }
}
Describe "getDefaultWorkDir の差し替え（環境変数 TEBUNKO_DEFAULT_WORKSPACE）" -Tag Io {
    # 場所を求めるだけで、そこに書かない。環境変数は必ず元に戻す（外したまま後のテストが書くと、既定のワークスペースに書く）
    BeforeEach {
        $script:savedWorkspace = $env:TEBUNKO_DEFAULT_WORKSPACE
    }
    AfterEach {
        $env:TEBUNKO_DEFAULT_WORKSPACE = $script:savedWorkspace
    }

    It "<Case>" -TestCases @(
        @{ Case = "未設定なら、プロファイルの Documents\tebunko_ws"; Value = $null; Expected = (Join-Path ([System.Environment]::GetFolderPath("UserProfile")) "Documents\tebunko_ws") }
        @{ Case = "空なら、プロファイルの Documents\tebunko_ws"; Value = ""; Expected = (Join-Path ([System.Environment]::GetFolderPath("UserProfile")) "Documents\tebunko_ws") }
        @{ Case = "空白だけなら、プロファイルの Documents\tebunko_ws"; Value = "  "; Expected = (Join-Path ([System.Environment]::GetFolderPath("UserProfile")) "Documents\tebunko_ws") }
        @{ Case = "絶対パスならそれ"; Value = "C:\Users\test\ws"; Expected = "C:\Users\test\ws" }
        @{ Case = "末尾の区切りは除く"; Value = "C:\Users\test\ws\"; Expected = "C:\Users\test\ws" }
        @{ Case = "ドライブ直下の区切りが / でもよい"; Value = "D:/work/ws"; Expected = "D:/work/ws" }
        @{ Case = "UNC のパスもよい"; Value = "\\server\share\ws"; Expected = "\\server\share\ws" }
    ) {
        $env:TEBUNKO_DEFAULT_WORKSPACE = $Value
        getDefaultWorkDir | Should -Be $Expected
    }

    It "絶対パスでない値は例外にする（黙って既定のワークスペースに戻らない）: <Value>" -TestCases @(
        @{ Value = "ws" }
        @{ Value = "..\ws" }
        @{ Value = "\ws" }
        @{ Value = "C:ws" }
        @{ Value = "\\server" }
        @{ Value = "C:\" }
        @{ Value = "C:\\" }
        @{ Value = "\\server\" }
        @{ Value = "\\server\\" }
    ) {
        $env:TEBUNKO_DEFAULT_WORKSPACE = $Value
        { getDefaultWorkDir } | Should -Throw "*絶対パス*"
    }

    It "引数でプロファイルを渡したときは、環境変数を見ない" {
        $env:TEBUNKO_DEFAULT_WORKSPACE = "C:\Users\test\ws"
        getDefaultWorkDir "C:\Users\test" | Should -Be "C:\Users\test\Documents\tebunko_ws"
    }

    It "getWorkDir（workspaceFolder が空）と、設定ファイルの逃げ先も、差し替えた場所になる" {
        $env:TEBUNKO_DEFAULT_WORKSPACE = "$TestDrive\差し替えた既定"
        getWorkDir "$TestDrive\無い設定\setting.config" | Should -Be "$TestDrive\差し替えた既定"
        # ツールのフォルダに書けない（存在しない）とき、設定ファイルは既定のワークスペースの直下
        getSettingsFilePath "$TestDrive\無いツール" (getDefaultWorkDir) | Should -Be "$TestDrive\差し替えた既定\setting.config"
    }
}

Describe "getDefaultWorkDir / testDefaultWorkspace / getWorkspaceBlockMessage" -Tag Io {
    It "既定はプロファイルの Documents\tebunko_ws（OneDrive のドキュメントではない）" {
        getDefaultWorkDir "C:\Users\test" | Should -Be "C:\Users\test\Documents\tebunko_ws"
    }

    It "無い・空・前から使っているワークスペースなら使える" {
        (testDefaultWorkspace "$TestDrive\無い").Usable | Should -Be $true
        [System.IO.Directory]::CreateDirectory("$TestDrive\空") | Out-Null
        (testDefaultWorkspace "$TestDrive\空").Usable | Should -Be $true
        [System.IO.Directory]::CreateDirectory("$TestDrive\前から\index") | Out-Null
        (testDefaultWorkspace "$TestDrive\前から").Usable | Should -Be $true
        [System.IO.Directory]::CreateDirectory("$TestDrive\前から2\content_index") | Out-Null
        (testDefaultWorkspace "$TestDrive\前から2").Usable | Should -Be $true
        [System.IO.Directory]::CreateDirectory("$TestDrive\一覧だけ") | Out-Null
        [System.IO.File]::WriteAllText("$TestDrive\一覧だけ\ingest_status.tsv", "")
        (testDefaultWorkspace "$TestDrive\一覧だけ").Usable | Should -Be $true
    }

    It "ほかのファイルが置いてあれば使えず、空のフォルダではないと伝える" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\ほか") | Out-Null
        [System.IO.File]::WriteAllText("$TestDrive\ほか\README.md", "")
        $check = testDefaultWorkspace "$TestDrive\ほか"
        $check.Usable | Should -Be $false
        $check.Message | Should -Be "「$TestDrive\ほか」は空のフォルダではありません。ワークスペースには別の空のフォルダを選んでください（［設定］の［変更…］）。"
    }

    It "今のワークスペースが既定の場所で、使えないときだけ文言を返す" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\既定ほか") | Out-Null
        [System.IO.File]::WriteAllText("$TestDrive\既定ほか\a.txt", "")
        getWorkspaceBlockMessage "$TestDrive\既定ほか" "$TestDrive\既定ほか" | Should -Match "空のフォルダではありません"
        getWorkspaceBlockMessage "$TestDrive\別" "$TestDrive\既定ほか" | Should -Be ""
        getWorkspaceBlockMessage "$TestDrive\空" "$TestDrive\空" | Should -Be ""
    }
}

Describe "writeSettings の書き込み（一時ファイルから置き換え）" -Tag Io {
    It "一時ファイルを残さず、BOM なし UTF-8 で書き、読み戻せる" {
        $path = "$TestDrive\原子的\setting.config"
        writeSettings ([ordered]@{ workspaceFolder = "C:"; useRegex = $true }) $path
        Test-Path -LiteralPath "${path}.tmp" | Should -Be $false
        $bytes = [System.IO.File]::ReadAllBytes($path)
        ($bytes[0] -eq 239 -and $bytes[1] -eq 187) | Should -Be $false
        $settings = readSettings $path
        $settings.workspaceFolder | Should -Be "C:"
        $settings.useRegex | Should -Be $true
    }
}

Describe "設定の同時の書き込み" -Tag Io {
    It "2 つのランスペースが別々のキーを交互に 50 回ずつ書いても、両方の値が残る" {
        $path = "$TestDrive\同時\setting.config"
        writeSettings (newSettings) $path
        $loader = (Resolve-Path "$PSScriptRoot\..\..\helpers\load.ps1").Path
        $script = {
            param ($loader, $path, $key, $prefix)
            . $loader
            for ($i = 1; $i -le 50; $i++) {
                updateSettings $key "$prefix$i" $path
            }
        }
        $jobs = @()
        foreach ($pair in @(@("openMode", "f"), @("workspaceFolder", "w"))) {
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
            $runspace.Open()
            $ps = [powershell]::Create()
            $ps.Runspace = $runspace
            [void]$ps.AddScript($script).AddArgument($loader).AddArgument($path).AddArgument($pair[0]).AddArgument($pair[1])
            $jobs += @{ Ps = $ps; Runspace = $runspace; Handle = $ps.BeginInvoke() }
        }
        foreach ($job in $jobs) {
            [void]$job.Ps.EndInvoke($job.Handle)
            $job.Ps.HadErrors | Should -Be $false
            $job.Ps.Dispose()
            $job.Runspace.Dispose()
        }
        { Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json } | Should -Not -Throw
        $settings = readSettings $path
        $settings.openMode | Should -Be "f50"
        $settings.workspaceFolder | Should -Be "w50"
    }
}

Describe "mergeAssignedIndexNames / saveAssignedIndexNames" -Tag Io {
    It "<name>" -TestCases @(
        @{ name = "名前が空の項目に、パスで突き合わせて名前を付ける（大文字と小文字は区別しない）"
           current = @(@{ Name = ""; Path = "C:\営業" }); assigned = @(@{ Name = "営業"; Path = "c:\営業" }); expected = "営業" }
        @{ name = "名前がある項目は、割り当てた名前で書き換えない"
           current = @(@{ Name = "旧"; Path = "C:\営業" }); assigned = @(@{ Name = "新"; Path = "C:\営業" }); expected = "旧" }
        @{ name = "同じ名前をほかの項目が使っていれば付けない"
           current = @(@{ Name = ""; Path = "C:\営業" }, @{ Name = "営業"; Path = "C:\経理" }); assigned = @(@{ Name = "営業"; Path = "C:\営業" }); expected = "" }
        @{ name = "割り当てに無いパス（画面で足したフォルダ）はそのまま残す"
           current = @(@{ Name = ""; Path = "C:\総務" }); assigned = @(@{ Name = "営業"; Path = "C:\営業" }); expected = "" }
    ) {
        param($name, $current, $assigned, $expected)
        $result = @(mergeAssignedIndexNames @($current | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Enabled = $true } }) @($assigned | ForEach-Object { [pscustomobject]$_ }))
        $result.Count | Should -Be $current.Count
        $result[0].Name | Should -Be $expected
    }

    It "インデクサが読んだ後に画面で足したクロール対象フォルダは、名前の書き戻しの後も残る" {
        $path = "$TestDrive\後勝ち\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = ""; Path = "C:\営業"; Enabled = $true }) $path
        $read = @(getTargetFolders $path)                               # インデクサが始めに読む
        $assigned = @(assignIndexNames $read @())
        writeTargetFolders @($read + [pscustomobject]@{ Name = "総務"; Path = "C:\総務"; Enabled = $false }) $path   # その間に画面で足す
        saveAssignedIndexNames $assigned $path                          # インデクサが名前を書き戻す
        $after = @(getTargetFolders $path)
        $after.Count | Should -Be 2
        $after[0].Name | Should -Be $assigned[0].Name
        $after[1].Path | Should -Be "C:\総務"
        $after[1].Enabled | Should -Be $false
    }
}

Describe "saveAssignedIndexNames の返す一覧" -Tag Io {
    It "保存した名前を返す。同じ名前をほかの項目が使っていれば付けず、画面が先に付けた名前はそのまま返す" {
        $path = "$TestDrive\競合\setting.config"
        writeTargetFolders @(
            [pscustomobject]@{ Name = ""; Path = "C:\営業"; Enabled = $true }
            [pscustomobject]@{ Name = "営業"; Path = "C:\経理"; Enabled = $true }
            [pscustomobject]@{ Name = "画面"; Path = "C:\総務"; Enabled = $true }
        ) $path
        $assigned = @(
            [pscustomobject]@{ Name = "営業"; Path = "C:\営業"; Enabled = $true }
            [pscustomobject]@{ Name = "インデクサ"; Path = "C:\総務"; Enabled = $true }
        )
        $saved = @(saveAssignedIndexNames $assigned $path)
        $saved.Count | Should -Be 3
        $saved[0].Name | Should -Be ""           # 同じ名前を経理が使っている
        $saved[2].Name | Should -Be "画面"       # 画面が先に付けた名前
        (@(getTargetFolders $path) | ForEach-Object { $_.Name }) -join "," | Should -Be ",営業,画面"
    }

    It "名前が付けば、その名前を返す" {
        $path = "$TestDrive\返す\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = ""; Path = "C:\営業"; Enabled = $true }) $path
        $saved = @(saveAssignedIndexNames @([pscustomobject]@{ Name = "営業"; Path = "C:\営業"; Enabled = $true }) $path)
        $saved.Count | Should -Be 1
        $saved[0].Name | Should -Be "営業"
    }
}
