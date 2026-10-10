# 検索対象ツリーのデータ集め（tebunko\index\index_tree_data.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "testIndexBookName" -Tag Unit {
    It "<name> は <expected>" -TestCases @(
        @{ name = "見積.xlsx"; expected = $true }
        @{ name = "見積.XLS"; expected = $true }
        @{ name = "議事録.docx"; expected = $true }
        @{ name = "資料.pptm"; expected = $true }
        @{ name = "営業部"; expected = $false }
        @{ name = "memo.txt"; expected = $true }
        @{ name = "memo.MD"; expected = $true }
        @{ name = "memo.json"; expected = $true }
        @{ name = "chart.js"; expected = $true }
        @{ name = "app.py"; expected = $true }
        @{ name = "a.xlsxx"; expected = $false }
        @{ name = "a.xl"; expected = $false }
        @{ name = "a.pdf"; expected = $false }
        @{ name = "a.exe"; expected = $false }
    ) {
        param ($name, $expected)
        testIndexBookName $name | Should -Be $expected
    }
}

Describe "インデックスのフォルダの読み取り" -Tag Io {
    # インデックスのフォルダ:
    #   営業\直下.xlsx\Sheet1.tsv          （根の直下のファイル）
    #   営業\営業部\見積.xlsx\4月.tsv
    #   営業\営業部\東京\
    #   営業\総務部\
    BeforeEach {
        $script:indexRoot = "$TestDrive\index\営業"
        Remove-Item -LiteralPath "$TestDrive\index" -Recurse -Force -ErrorAction SilentlyContinue
        newTsv "$script:indexRoot\直下.xlsx\Sheet1.tsv" @("1`ta")
        newTsv "$script:indexRoot\営業部\見積.xlsx\4月.tsv" @("1`ta")
        [System.IO.Directory]::CreateDirectory("$script:indexRoot\営業部\東京") | Out-Null
        [System.IO.Directory]::CreateDirectory("$script:indexRoot\総務部") | Out-Null
    }

    It "testIndexBookDirPath は、名前が .xlsx などで終わる本物のフォルダ（本文インデックスのファイル・サブフォルダがある）を見分ける" {
        $dir = "$TestDrive\bookdir_path"
        newTsv "$dir\資料.xlsx\content_index.docx.001.tsv" @("x")
        [void][System.IO.Directory]::CreateDirectory("$dir\親.xlsx\子")
        [void][System.IO.Directory]::CreateDirectory("$dir\空.xlsx")
        newTsv "$dir\B.xlsx\S.tsv" @("x")
        testIndexBookDirPath "$dir\資料.xlsx" | Should -Be $false
        testIndexBookDirPath "$dir\親.xlsx" | Should -Be $false
        testIndexBookDirPath "$dir\空.xlsx" | Should -Be $true
        testIndexBookDirPath "$dir\B.xlsx" | Should -Be $true
        testIndexBookDirPath "$dir\営業部" | Should -Be $false
        # 本物のフォルダはツリーに出し、元のファイルごとのフォルダは出さない
        testIndexFolderHasSubfolders $dir | Should -Be $true
    }

    It "testIndexBookDirPath は、chart.js のような名前の本物のフォルダを誤判定しない" {
        # .js は取り込み対象のテキストの拡張子だが、中にファイル・サブフォルダがある本物のフォルダは元のファイルごとのフォルダではない
        $dir = "$TestDrive\bookdir_path_js"
        [void][System.IO.Directory]::CreateDirectory("$dir\chart.js\lib")
        [System.IO.File]::WriteAllText("$dir\chart.js\index.js", "dummy")
        testIndexBookName "chart.js" | Should -Be $true
        testIndexBookDirPath "$dir\chart.js" | Should -Be $false
    }

    It "調べた直後に、そのフォルダを移動できる（ツリーを開いた後の上書きのインポート）" {
        # 調べたフォルダを掴んだまま残さないこと
        $dir = "$TestDrive\node_handle\営業"
        newTsv "$dir\content_index.xlsx.001.tsv" @("x")
        newTsv "$dir\見積\content_index.xlsx.001.tsv" @("x")
        newTsv "$dir\見積\B.xlsx\S.tsv" @("x")
        newTsv "$dir\見積\C.xlsx\S.tsv" @("x")
        [void][System.IO.Directory]::CreateDirectory("$dir\見積\親.xlsx\子")

        testIndexBookDirPath "$dir\見積\親.xlsx" | Should -Be $false
        testIndexFolderHasSubfolders $dir | Should -Be $true
        testIndexFolderHasFiles $dir | Should -Be $true
        testIndexFolderHasFiles "$dir\見積" | Should -Be $true
        $null = getIndexFolderChildren $dir

        { [System.IO.Directory]::Move($dir, "$TestDrive\node_handle\moved") } | Should -Not -Throw
    }

    It "直下の本文インデックスのファイルもファイルとして数える（ほかの .tsv は数えない）" {
        newTsv "$script:indexRoot\人事部\a.tsv" @("x")
        testIndexFolderHasFiles "$script:indexRoot\人事部" | Should -Be $false
        newTsv "$script:indexRoot\総務部\content_index.xlsx.001.tsv" @("x")
        testIndexFolderHasFiles "$script:indexRoot\総務部" | Should -Be $true
        testIndexFolderHasFiles "$script:indexRoot\営業部\東京" | Should -Be $false
        testIndexFolderHasFiles "$TestDrive\無い" | Should -Be $false
        testIndexFolderHasSubfolders "$TestDrive\無い" | Should -Be $false
    }

    It "getIndexFolderChildren は、サブフォルダを名前順に返し（Office のフォルダは除く）、直下のファイルの有無とサブフォルダの有無を付ける" {
        $result = getIndexFolderChildren $script:indexRoot

        $result.Error | Should -Be ""
        $result.HasFiles | Should -Be $true
        @($result.Folders | ForEach-Object { $_.Name }) | Should -Be @("営業部", "総務部")
        $result.Folders[0].HasSubfolders | Should -Be $true
        $result.Folders[1].HasSubfolders | Should -Be $false
    }

    It "getIndexFolderChildren は、直下にファイルが無ければ HasFiles を偽にする" {
        Remove-Item -LiteralPath "$script:indexRoot\直下.xlsx" -Recurse -Force

        (getIndexFolderChildren $script:indexRoot).HasFiles | Should -Be $false
    }

    It "getIndexFolderChildren は、フォルダが無ければ Error に文面を入れて空で返す" {
        $result = getIndexFolderChildren "$TestDrive\index\無い"

        $result.Error | Should -Not -BeNullOrEmpty
        @($result.Folders).Count | Should -Be 0
        $result.HasFiles | Should -Be $false
    }

    It "本文インデックスのファイルの型は contentIndexFilePattern と同じ（index_tree_data.ps1 のソースを読んで確かめる）" {
        $source = [System.IO.File]::ReadAllText("${scriptsDir}\tebunko\index\index_tree_data.ps1")
        # 直下のブックのフォルダの中を探す "*.tsv" は対象外
        $literals = @([regex]::Matches($source, 'GetFiles\([^,]+,\s*"([^"]*\.tsv)"\)') |
            ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -ne "*.tsv" })
        $literals.Count | Should -Be 0
        ([regex]::Matches($source, 'GetFiles\([^,]+,\s*\$\{contentIndexFilePattern\}\)')).Count | Should -Be 2
    }
}

Describe "getIndexTreeData" -Tag Io {
    BeforeAll {
        function newTreeFiles([string]$dir) {
            newTsv "$dir\営業部\直下.xlsx\Sheet1.tsv" @("`t見積")
            newTsv "$dir\営業部\2024\見積.xlsx\4月.tsv" @("`t見積")
            newTsv "$dir\営業部\2024\東京\議事録.docx\ページ001.tsv" @("見積")
            newTsv "$dir\総務部\規程.docx\ページ001.tsv" @("規程")
        }
    }

    BeforeEach {
        $script:dir = "$TestDrive\work\index"
        $script:statusPath = "$TestDrive\work\ingest_status.tsv"
        $script:settingsPath = "$TestDrive\setting.config"
        Remove-Item -LiteralPath "$TestDrive\work", $script:settingsPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    It "インデックスのフォルダが無ければ、Missing で空を返す" {
        $data = getIndexTreeData $script:dir $script:statusPath $script:settingsPath

        $data.State | Should -Be ${pathStateMissing}
        @($data.Indexes).Count | Should -Be 0
    }

    It "インデックスごとに、名前・元のフォルダ・フォルダの有無・サブフォルダの有無を返す" {
        newTreeFiles $script:dir
        [System.IO.File]::WriteAllText($script:settingsPath, '{ "targetFolders": [ { "name": "営業部", "path": "C:\\共有\\営業部", "enabled": true } ] }', (New-Object System.Text.UTF8Encoding $false))

        $data = getIndexTreeData $script:dir $script:statusPath $script:settingsPath

        $data.State | Should -Be ${pathStateFound}
        (@($data.Indexes | ForEach-Object { $_.Name }) -join ",") | Should -Be "営業部,総務部"
        $sales = $data.Indexes[0]
        $sales.SourcePath | Should -Be "C:\共有\営業部"
        $sales.Exists | Should -Be $true
        $sales.HasSubfolders | Should -Be $true
        $data.Indexes[1].SourcePath | Should -Be ""
        $data.Indexes[1].HasSubfolders | Should -Be $false
        $data.Sources["営業部"] | Should -Be "C:\共有\営業部"
    }

    It "渡したパスの先祖すべての子を返す（深い階層の展開・除外を戻すため）" {
        newTreeFiles $script:dir
        $root = (getIndexTreeData $script:dir $script:statusPath $script:settingsPath).Root

        $data = getIndexTreeData $script:dir $script:statusPath $script:settingsPath @("$root\営業部\2024\東京")

        $data.Children.ContainsKey("$root\営業部") | Should -Be $true
        $data.Children.ContainsKey("$root\営業部\2024") | Should -Be $true
        $data.Children.ContainsKey("$root\営業部\2024\東京") | Should -Be $true
        $data.Children.ContainsKey("$root\総務部") | Should -Be $false
        (@($data.Children["$root\営業部"].Folders | ForEach-Object { $_.Name }) -join ",") | Should -Be "2024"
    }

    It "インデックスの外のパス・無いフォルダのパスは無視する" {
        newTreeFiles $script:dir
        $root = (getIndexTreeData $script:dir $script:statusPath $script:settingsPath).Root

        $data = getIndexTreeData $script:dir $script:statusPath $script:settingsPath @("C:\ほか\営業部", "$root\営業部\無い\奥")

        $data.Children.ContainsKey("C:\ほか\営業部") | Should -Be $false
        $data.Children.ContainsKey("$root\営業部") | Should -Be $true
        $data.Children.ContainsKey("$root\営業部\無い") | Should -Be $false
    }
}

Describe "getIndexTreeData（届かないとき）" -Tag Unit {
    It "届かない場所は、列挙せずに State と文面だけを返す" {
        Mock getPathState { @{ State = ${pathStateUnreachable}; Message = "届きません"; IsDirectory = $false } }
        Mock getSearchIndexData { throw "届かない場所を列挙した" }

        $data = getIndexTreeData "\\fileserver\共有\index" "x" "y"

        $data.State | Should -Be ${pathStateUnreachable}
        $data.Message | Should -Be "届きません"
        @($data.Indexes).Count | Should -Be 0
        Should -Invoke getSearchIndexData -Times 0
    }
}
