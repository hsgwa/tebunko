# インデックスのエクスポート・インポートの判断（tebunko\index\index_archive_rules.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function newManifestFile([string]$path, $size = 10, [string]$sha256 = ("a" * 64)) {
        return [pscustomobject]@{ path = $path; size = $size; sha256 = $sha256 }
    }

    function newManifest([object[]]$files, [string]$indexName = "営業", [string]$sourceFolder = "C:\共有\営業部", [int]$formatVersion = ${indexArchiveFormatVersion}) {
        return [pscustomobject]@{
            format = ${indexArchiveFormat}; formatVersion = $formatVersion; appVersion = "v0.3.0"; exportedAt = "2026-09-27T10:00:00+09:00"
            indexName = $indexName; sourceFolder = $sourceFolder; files = $files
        }
    }
}

Describe "testIndexArchiveEntryPath" -Tag Unit {
    It "安全なパスは空文字列を返す" -TestCases @(
        @{ path = "取り込み一覧.tsv" }
        @{ path = "content_index/content_index.xlsx.001.tsv" }
        @{ path = "content_index/2024/見積/content_index.xlsx.001.tsv" }
        @{ path = "content_index/[かっこ]/content_index.xlsx.001.tsv" }
    ) {
        param ($path)
        testIndexArchiveEntryPath $path | Should -Be ""
    }

    It "危ないパスは理由を返す（zip slip・予約名など）" -TestCases @(
        @{ path = ""; label = "空" }
        @{ path = "content_index\x.tsv"; label = "円マーク区切り" }
        @{ path = "/x.tsv"; label = "/ 始まり" }
        @{ path = "C:/x.tsv"; label = "ドライブ文字" }
        @{ path = "../x.tsv"; label = ".." }
        @{ path = "content_index/../../x.tsv"; label = "途中の .." }
        @{ path = "content_index//x.tsv"; label = "空の区切り" }
        @{ path = "content_index/a:b.tsv"; label = "コロン（代替データストリーム）" }
        @{ path = "content_index/CON/x.tsv"; label = "予約名のフォルダ" }
        @{ path = "content_index/CON.tsv"; label = "拡張子付きの予約名" }
        @{ path = "content_index/x ./y.tsv"; label = "末尾が空白の区切り" }
        @{ path = "content_index/$('x' * 300)/y.tsv"; label = "長すぎる区切り" }
    ) {
        param ($path, $label)
        testIndexArchiveEntryPath $path | Should -Not -BeNullOrEmpty
    }
}

Describe "testIndexArchiveEntryLocation" -Tag Unit {
    It "許される場所は `$true" -TestCases @(
        @{ path = "取り込み一覧.tsv" }
        @{ path = "content_index/content_index.xlsx.001.tsv" }
        @{ path = "content_index/2024/見積/content_index.docx.012.tsv" }
    ) {
        param ($path)
        testIndexArchiveEntryLocation $path | Should -Be $true
    }

    It "許されない場所は `$false" -TestCases @(
        @{ path = "tebunko-index.json" }
        @{ path = "元のフォルダ.txt" }
        @{ path = "content_index/2024/元のフォルダ.txt" }
        @{ path = "content_index/content.xlsx.001.tsv"; label = "前の版の名前" }
        @{ path = "content_index/content_index.xlsx.tsv"; label = "番号が無い" }
        @{ path = "system_index/content_index.xlsx.001.tsv" }
    ) {
        param ($path)
        testIndexArchiveEntryLocation $path | Should -Be $false
    }
}

Describe "getManifestSourceFolder" -Tag Unit {
    It "絶対パス・UNC はそのまま、それ以外は空にする" -TestCases @(
        @{ folder = "C:\共有\営業部"; expected = "C:\共有\営業部" }
        @{ folder = "\\server\share\営業部"; expected = "\\server\share\営業部" }
        @{ folder = ""; expected = "" }
        @{ folder = "営業部"; expected = "" }
    ) {
        param ($folder, $expected)
        getManifestSourceFolder ([pscustomobject]@{ sourceFolder = $folder }) | Should -Be $expected
    }

    It "制御文字を含むパスは空にする" {
        # 制御文字は結果の XML に書けないので、表（テスト名に値が入る）ではなく本体で組み立てる
        $folder = "C:\" + [string][char]1

        getManifestSourceFolder ([pscustomobject]@{ sourceFolder = $folder }) | Should -Be ""
    }
}

Describe "testIndexArchiveManifest" -Tag Unit {
    It "正しい目録は空文字列を返す" {
        $files = @((newManifestFile "取り込み一覧.tsv"), (newManifestFile "content_index/content_index.xlsx.001.tsv"))
        $manifest = newManifest $files
        $entryNames = @("tebunko-index.json", "取り込み一覧.tsv", "content_index/content_index.xlsx.001.tsv")

        testIndexArchiveManifest $manifest $entryNames 100 | Should -Be ""
    }

    It "目録が大きすぎれば理由を返す（`$manifest を読む前に判定する）" {
        testIndexArchiveManifest $null @("tebunko-index.json") (65MB) | Should -Not -BeNullOrEmpty
    }

    It "目録が読めない（`$null）とき理由を返す" {
        testIndexArchiveManifest $null @("tebunko-index.json") 10 | Should -Not -BeNullOrEmpty
    }

    It "format が違うと理由を返す" {
        $files = @((newManifestFile "取り込み一覧.tsv"))
        $manifest = newManifest $files
        $manifest.format = "other"
        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "formatVersion が整数でない・1未満なら理由を返す" -TestCases @(
        @{ version = "abc" }
        @{ version = "0" }
        @{ version = "-1" }
    ) {
        param ($version)
        $files = @((newManifestFile "取り込み一覧.tsv"))
        $manifest = newManifest $files
        $manifest.formatVersion = $version
        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "formatVersion が今の版より大きければ「新しい版」の理由を返す" {
        $files = @((newManifestFile "取り込み一覧.tsv"))
        $manifest = newManifest $files -formatVersion 2

        $reason = testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv") 10
        $reason | Should -Match "新しい版"
    }

    It "インデックス名が testIndexName を通らなければ理由を返す" {
        $files = @((newManifestFile "取り込み一覧.tsv"))
        $manifest = newManifest $files -indexName "a\..\b"

        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "zip の中に同じ名前のエントリーが 2 つあれば理由を返す" {
        $files = @((newManifestFile "取り込み一覧.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv", "取り込み一覧.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "目録に危ない・許されない場所のパスがあれば理由を返す" -TestCases @(
        @{ path = "../x.tsv" }
        @{ path = "system_index/x.tsv" }
    ) {
        param ($path)
        $files = @((newManifestFile "取り込み一覧.tsv"), (newManifestFile $path))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv", $path) 10 | Should -Not -BeNullOrEmpty
    }

    It "目録に同じパスが2つあれば理由を返す" {
        $files = @((newManifestFile "取り込み一覧.tsv"), (newManifestFile "取り込み一覧.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "大きさが数値でない・SHA-256 が無ければ理由を返す" -TestCases @(
        @{ size = "abc"; sha = ("a" * 64) }
        @{ size = -1; sha = ("a" * 64) }
        @{ size = 10; sha = "" }
    ) {
        param ($size, $sha)
        $files = @((newManifestFile "取り込み一覧.tsv" $size $sha))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "取り込み一覧.tsv が目録に無ければ理由を返す" {
        $files = @((newManifestFile "content_index/content_index.xlsx.001.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "content_index/content_index.xlsx.001.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "目録に無いエントリーが zip にあれば理由を返す" {
        $files = @((newManifestFile "取り込み一覧.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv", "content_index/content_index.xlsx.001.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "目録にあるのに zip に無いエントリーがあれば理由を返す" {
        $files = @((newManifestFile "取り込み一覧.tsv"), (newManifestFile "content_index/content_index.xlsx.001.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "取り込み一覧.tsv") 10 | Should -Not -BeNullOrEmpty
    }
}

Describe "testImportedStatusRow・testImportedStatusLines" -Tag Unit {
    BeforeAll {
        # ${statusColumns}・${stateDone} は lib.ps1（load.ps1 の中で読み込む）が定義するため、BeforeAll（実行時）で使う
        ${script:header} = ${statusColumns} -join "`t"
        function newRow([string]$relPath = "見積\A社.xlsx", [string]$state = ${stateDone}) {
            return ([string]::Join("`t", @($relPath, "2024/01/01 00:00:00", "100", $state, "1", "2024/01/01 00:00:00", "", "1")))
        }
    }

    It "列の数が正しい行は空文字列を返す" {
        testImportedStatusRow (newRow) | Should -Be ""
    }

    It "列の数が違う・状態が知らない値・相対パスが危ないと理由を返す" -TestCases @(
        @{ line = "a`tb"; label = "列が足りない" }
        @{ line = "見積\A社.xlsx`t2024/01/01 00:00:00`t100`tそのた`t1`t2024/01/01 00:00:00`t`t1"; label = "知らない状態" }
        @{ line = "..\A社.xlsx`t2024/01/01 00:00:00`t100`t済`t1`t2024/01/01 00:00:00`t`t1"; label = "危ない相対パス" }
        @{ line = "`t2024/01/01 00:00:00`t100`t済`t1`t2024/01/01 00:00:00`t`t1"; label = "空の相対パス" }
    ) {
        param ($line, $label)
        testImportedStatusRow $line | Should -Not -BeNullOrEmpty
    }

    It "見出し行が正しく、行がすべて正しければ空文字列を返す" {
        $lines = @($script:header, (newRow "a.xlsx"), (newRow "b.xlsx"))
        testImportedStatusLines $lines | Should -Be ""
    }

    It "見出しが違うと理由を返す" {
        testImportedStatusLines @("違う見出し", (newRow)) | Should -Not -BeNullOrEmpty
    }

    It "行がおかしいと理由を返す" {
        testImportedStatusLines @($script:header, "壊れた行") | Should -Not -BeNullOrEmpty
    }

    It "相対パスが重なると理由を返す" {
        testImportedStatusLines @($script:header, (newRow "a.xlsx"), (newRow "a.xlsx")) | Should -Not -BeNullOrEmpty
    }

    It "空行は読み飛ばす" {
        $lines = @($script:header, (newRow "a.xlsx"), "")
        testImportedStatusLines $lines | Should -Be ""
    }
}

Describe "getImportIndexName" -Tag Unit {
    It "Cancel は `$null を返す" {
        getImportIndexName "営業" @("営業") ${importCollisionCancel} | Should -Be $null
    }

    It "Overwrite は提案した名前のまま返す" {
        getImportIndexName "営業" @("営業") ${importCollisionOverwrite} | Should -Be "営業"
    }

    It "Rename は重ならない名前にする（重ならなければそのまま）" -TestCases @(
        @{ usedNames = @(); expected = "営業" }
        @{ usedNames = @("営業"); expected = "営業(2)" }
        @{ usedNames = @("営業", "営業(2)"); expected = "営業(3)" }
    ) {
        param ($usedNames, $expected)
        getImportIndexName "営業" $usedNames ${importCollisionRename} | Should -Be $expected
    }
}

Describe "getExportFileName" -Tag Unit {
    It "既定の形式にし、重なれば (2) を付ける" {
        $now = [datetime]::new(2026, 9, 27)

        getExportFileName "営業" @() $now | Should -Be "営業_インデックス_20260927.zip"
        getExportFileName "営業" @("営業_インデックス_20260927.zip") $now | Should -Be "営業_インデックス_20260927(2).zip"
        getExportFileName "営業" @("営業_インデックス_20260927.zip", "営業_インデックス_20260927(2).zip") $now | Should -Be "営業_インデックス_20260927(3).zip"
    }

    It "ファイル名に使えない文字は全角にする" {
        $now = [datetime]::new(2026, 9, 27)
        getExportFileName "A/B" @() $now | Should -Be "A／B_インデックス_20260927.zip"
    }
}

Describe "getIndexFolderConflict" -Tag Unit {
    BeforeAll {
        $others = @([pscustomobject]@{ Name = "売上"; Path = "C:\data\売上" })
    }

    It "<label>: <expected>" -TestCases @(
        @{ label = "重ならない"; folder = "C:\data\見積"; expected = "" }
        @{ label = "同じフォルダ"; folder = "C:\data\売上"; expected = "既にあります" }
        @{ label = "中のフォルダ"; folder = "C:\data\売上\2024"; expected = "中のフォルダです" }
        @{ label = "含むフォルダ"; folder = "C:\data"; expected = "があります" }
    ) {
        param ($label, $folder, $expected)
        $message = getIndexFolderConflict $folder $others
        if ($expected -eq "") { $message | Should -Be "" } else { $message | Should -Match $expected }
    }

    It "比べる相手が無ければ空文字列" {
        getIndexFolderConflict "C:\data\売上" @() | Should -Be ""
    }
}