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
        @{ path = "ingest_status.tsv" }
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
        @{ path = "ingest_status.tsv" }
        @{ path = "content_index/content_index.xlsx.001.tsv" }
        @{ path = "content_index/2024/見積/content_index.docx.012.tsv" }
    ) {
        param ($path)
        testIndexArchiveEntryLocation $path | Should -Be $true
    }

    It "許されない場所は `$false" -TestCases @(
        @{ path = "tebunko-index.json" }
        @{ path = "source_folder.txt" }
        @{ path = "content_index/2024/source_folder.txt" }
        @{ path = "content_index/content.xlsx.001.tsv"; label = "前の版の名前" }
        @{ path = "content_index/content_index.xlsx.tsv"; label = "番号が無い" }
        @{ path = "content_index/content_index.xlsx.99999999999.tsv"; label = "番号が桁あふれ" }
        @{ path = "content_index/content_index.xlsx.００１.tsv"; label = "全角の数字" }
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
        $files = @((newManifestFile "ingest_status.tsv"), (newManifestFile "content_index/content_index.xlsx.001.tsv"))
        $manifest = newManifest $files
        $entryNames = @("tebunko-index.json", "ingest_status.tsv", "content_index/content_index.xlsx.001.tsv")

        testIndexArchiveManifest $manifest $entryNames 100 | Should -Be ""
    }

    It "目録が大きすぎれば理由を返す（`$manifest を読む前に判定する）" {
        testIndexArchiveManifest $null @("tebunko-index.json") (65MB) | Should -Not -BeNullOrEmpty
    }

    It "目録が読めない（`$null）とき理由を返す" {
        testIndexArchiveManifest $null @("tebunko-index.json") 10 | Should -Not -BeNullOrEmpty
    }

    It "上限: <label>は、ちょうどなら通り、1 超えると理由を返す" -TestCases @(
        @{ label = "エントリーの数"; kind = "entries" }
        @{ label = "目録の files の大きさの合計"; kind = "total" }
        @{ label = "ingest_status.tsv の大きさ"; kind = "status" }
    ) {
        param ($label, $kind)
        $statusSize = 10
        $contentSize = 10
        $extraNames = @()
        switch ($kind) {
            "entries" { $extraNames = @(1..(${indexArchiveMaxEntries} - 3) | ForEach-Object { "x$_" }) }
            "total" { $statusSize = 1; $contentSize = ${indexArchiveMaxTotalBytes} - 1 }
            "status" { $statusSize = ${indexArchiveMaxStatusBytes} }
        }
        $entryNames = @("tebunko-index.json", "ingest_status.tsv", "content_index/content_index.xlsx.001.tsv") + $extraNames
        $files = @((newManifestFile "ingest_status.tsv" $statusSize), (newManifestFile "content_index/content_index.xlsx.001.tsv" $contentSize))
        # 上限ちょうど（エントリーは目録に無い名前が混ざるので、数の確かめだけ見る）
        if ($kind -eq "entries") {
            (testIndexArchiveManifest (newManifest $files) $entryNames 100) | Should -Not -BeLike "*多すぎ*"
            (testIndexArchiveManifest (newManifest $files) ($entryNames + "x_over") 100) | Should -BeLike "*多すぎ*"
        } else {
            testIndexArchiveManifest (newManifest $files) $entryNames 100 | Should -Be ""
            $over = if ($kind -eq "total") { @((newManifestFile "ingest_status.tsv" 2), (newManifestFile "content_index/content_index.xlsx.001.tsv" $contentSize)) } else { @((newManifestFile "ingest_status.tsv" ($statusSize + 1)), (newManifestFile "content_index/content_index.xlsx.001.tsv" $contentSize)) }
            testIndexArchiveManifest (newManifest $over) $entryNames 100 | Should -BeLike "*大きすぎ*"
        }
    }

    It "format が違うと理由を返す" {
        $files = @((newManifestFile "ingest_status.tsv"))
        $manifest = newManifest $files
        $manifest.format = "other"
        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "formatVersion が整数でない・1未満なら理由を返す" -TestCases @(
        @{ version = "abc" }
        @{ version = "0" }
        @{ version = "-1" }
    ) {
        param ($version)
        $files = @((newManifestFile "ingest_status.tsv"))
        $manifest = newManifest $files
        $manifest.formatVersion = $version
        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "formatVersion が今の版より大きければ「新しい版」の理由を返す" {
        $files = @((newManifestFile "ingest_status.tsv"))
        $manifest = newManifest $files -formatVersion (${indexArchiveFormatVersion} + 1)

        $reason = testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv") 10
        $reason | Should -Match "新しい版"
    }

    It "formatVersion が今の版より小さければ、古い形式の理由（作り直しの案内つき）を返す" {
        $files = @((newManifestFile "ingest_status.tsv"))
        $manifest = newManifest $files -formatVersion 1

        $reason = testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv") 10
        $reason | Should -Match "古い形式"
        $reason | Should -Match ([regex]::Escape(${indexRebuildMessage}))
    }

    It "インデックス名が testIndexName を通らなければ理由を返す" {
        $files = @((newManifestFile "ingest_status.tsv"))
        $manifest = newManifest $files -indexName "a\..\b"

        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "zip の中に同じ名前のエントリーが 2 つあれば理由を返す" {
        $files = @((newManifestFile "ingest_status.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv", "ingest_status.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "目録に危ない・許されない場所のパスがあれば理由を返す" -TestCases @(
        @{ path = "../x.tsv" }
        @{ path = "system_index/x.tsv" }
    ) {
        param ($path)
        $files = @((newManifestFile "ingest_status.tsv"), (newManifestFile $path))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv", $path) 10 | Should -Not -BeNullOrEmpty
    }

    It "目録に同じパスが2つあれば理由を返す" {
        $files = @((newManifestFile "ingest_status.tsv"), (newManifestFile "ingest_status.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "大きさが数値でない・SHA-256 が無ければ理由を返す" -TestCases @(
        @{ size = "abc"; sha = ("a" * 64) }
        @{ size = -1; sha = ("a" * 64) }
        @{ size = 10; sha = "" }
    ) {
        param ($size, $sha)
        $files = @((newManifestFile "ingest_status.tsv" $size $sha))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "ingest_status.tsv が目録に無ければ理由を返す" {
        $files = @((newManifestFile "content_index/content_index.xlsx.001.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "content_index/content_index.xlsx.001.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "目録に無いエントリーが zip にあれば理由を返す" {
        $files = @((newManifestFile "ingest_status.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv", "content_index/content_index.xlsx.001.tsv") 10 | Should -Not -BeNullOrEmpty
    }

    It "目録にあるのに zip に無いエントリーがあれば理由を返す" {
        $files = @((newManifestFile "ingest_status.tsv"), (newManifestFile "content_index/content_index.xlsx.001.tsv"))
        $manifest = newManifest $files

        testIndexArchiveManifest $manifest @("tebunko-index.json", "ingest_status.tsv") 10 | Should -Not -BeNullOrEmpty
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

Describe "testImportedContentIndexFile" -Tag Unit {
    BeforeAll {
        $script:rs = [string]${contentIndexMark}
        $script:ver = "$($script:rs) 版=${contentIndexVersion}"
        $script:good = "$($script:ver)`n$($script:rs) ファイル名=a.xlsx`n$($script:rs) シート=Sheet1`n品名`t数量`n鉛筆`t10`n"
    }

    It "断る（<label>）: 理由を返す" -TestCases @(
        @{ label = "版が違う（検索・更新の両方が断る）"; text = { "$($script:rs) 版=99`n$($script:rs) ファイル名=a.xlsx`n" }; size = 100 }
        @{ label = "版の行が無い（検索・更新の両方が断る）"; text = { "$($script:rs) ファイル名=a.xlsx`n$($script:rs) シート=Sheet1`n鉛筆`n" }; size = 100 }
        @{ label = "2 行目以降に版の違う行がある（検索だけが断る）"; text = { "$($script:ver)`n$($script:rs) ファイル名=a.xlsx`n$($script:rs) 版=99`n" }; size = 100 }
        @{ label = "版の行が 1 行目に無い（更新だけが断る）"; text = { "品名`n$($script:ver)`n$($script:rs) ファイル名=a.xlsx`n" }; size = 100 }
        @{ label = "ファイル名の行で終わり改行が無い（更新だけが断る）"; text = { "$($script:ver)`n$($script:rs) ファイル名=a.xlsx" }; size = 100 }
        @{ label = "大きさが上限を超える（大きさで断る。中身は渡さない）"; text = { $null }; size = (64MB + 1) }
    ) {
        param ($label, $text, $size)
        testImportedContentIndexFile $size (& $text) | Should -Not -BeNullOrEmpty
    }

    It "受け入れる（<label>）: 空文字列を返す" -TestCases @(
        @{ label = "正しい中身"; text = { $script:good } }
        @{ label = "空の中身"; text = { "" } }
        @{ label = "メタ情報の行に = が無い"; text = { $script:good.Replace("シート=Sheet1", "シート") } }
        @{ label = "中身の行が途中で切れている"; text = { $script:good.Substring(0, $script:good.IndexOf("鉛筆") + 1) } }
        @{ label = "メタ情報の行だけで中身が無い"; text = { $script:good.Substring(0, $script:good.IndexOf("品名")) } }
    ) {
        param ($label, $text)
        testImportedContentIndexFile 100 (& $text) | Should -Be ""
    }

    It "大きさが上限ちょうどなら受け入れる" {
        testImportedContentIndexFile 64MB $script:good | Should -Be ""
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