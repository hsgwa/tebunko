# 前の版のファイル（見本・golden。tests/testdata/compat/index/<見本の名前>/）が、今のコードでもそのまま読め、
# 今のコードが作るものと同じになることを確かめる（互換テスト）。
# 見本の置き方・作り方は tests/testdata/README.md「前の版のファイル（compat\）」。
# Office は要らない（見本の Word・PowerPoint・Excel は、読み直すときも COM 無しで読める形式のため）。
# f だけ、ConfirmTargets で取りやめたときに Excel を起動していないことを確かめるため、プロセスの数を見る
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\indexer.ps1"

    $reporterPath = "${scriptsDir}\tebunko\indexer\indexing_reporter.ps1"

    ${compatRoot} = "${testDataDir}\compat\index"
    ${sampleName} = "v0.3.1+english-names"
    ${sampleDir}  = "${compatRoot}\${sampleName}"
    ${sampleWsDir} = "${sampleDir}\ws"
    ${expected}   = [System.IO.File]::ReadAllText("${sampleDir}\expected.json", [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    ${indexName}  = ${expected}.indexName

    function script:restoreFileTimes {
        # file_times.tsv の Ticks で、コピー先（source/ を写した先）のファイルの更新日時を見本の記録どおりに戻す
        # （git は取り出すときにファイルの更新日時を今の日時にしてしまうため。tools/make_index_golden.ps1 のコメントも参照）
        param ([string]$destSourceDir, [string]$timesPath)
        foreach ($line in @([System.IO.File]::ReadAllLines($timesPath, [System.Text.Encoding]::UTF8))) {
            if (!$line -or $line.StartsWith("#")) { continue }
            $cols = $line -split "`t"
            $path = Join-Path $destSourceDir ($cols[0] -replace "/", "\")
            [System.IO.File]::SetLastWriteTime($path, (New-Object datetime ([long]$cols[1])))
        }
    }

    function script:newCompatRoot {
        # 見本の source/・ws/ を TestDrive にコピーする。root\work は runIndexer・writeTestSettings がそのまま使う置き場所
        param ([string]$name)
        $root = Join-Path $TestDrive $name
        $sourceDir = Join-Path $root "source"
        $workDir = Join-Path $root "work"
        Copy-Item -LiteralPath "${sampleDir}\source" -Destination $sourceDir -Recurse -Force
        Copy-Item -LiteralPath ${sampleWsDir} -Destination $workDir -Recurse -Force
        restoreFileTimes $sourceDir "${sampleDir}\file_times.tsv"
        return @{ Root = $root; SourceDir = $sourceDir; WorkDir = $workDir }
    }

    function script:toRelativePath {
        # 検索ヒットの RelDir・Book（インデックス名を含む）から、source/ からの相対パス（/ 区切り）を作る
        param ([string]$relDir, [string]$book)
        $dir = $relDir.Substring(${indexName}.Length).Trim("\")
        $rel = if ($dir) { "$dir\$book" } else { $book }
        return $rel.Replace("\", "/")
    }

    function script:assertSearchHitsMatchExpected {
        # workDir（ワークスペース）を検索し、見本（expected.json の searches）とヒットが同じになることを確かめる
        param ([string]$workDir, [object[]]$searches)
        $ws = [Workspace]::new($workDir)
        foreach ($search in $searches) {
            $folders = @(@{ Root = $ws.IndexDir; RelPath = ""; Recurse = $true })
            $index = getIndexPackFiles $folders
            $found = searchPackIndex $search.word $index.Packs $true 0 -caseSensitive $false -includeShapes $true -includeComments $true
            $actualHits = @($found.Hits | ForEach-Object {
                [pscustomobject]@{ relativePath = (toRelativePath $_.RelDir $_.Book); location = $_.Location; line = $_.Line }
            } | Sort-Object relativePath, location, line)
            $expectedHits = @($search.hits | Sort-Object relativePath, location, line)
            $actualHits.Count | Should -Be $expectedHits.Count -Because "検索語 '$($search.word)' のヒット数"
            for ($i = 0; $i -lt $expectedHits.Count; $i++) {
                $actualHits[$i].relativePath | Should -Be $expectedHits[$i].relativePath -Because "検索語 '$($search.word)' の $i 件目"
                $actualHits[$i].location | Should -Be $expectedHits[$i].location -Because "検索語 '$($search.word)' の $i 件目"
                $actualHits[$i].line | Should -Be $expectedHits[$i].line -Because "検索語 '$($search.word)' の $i 件目"
            }
        }
    }

    function script:assertIngestListMatchesExpected {
        # readStatusFile の結果が、見本（expected.json の ingestList）と同じになることを確かめる
        param ($status)
        $status.Rows.Count | Should -Be ${expected}.ingestList.Count
        foreach ($row in ${expected}.ingestList) {
            $key = "${indexName}\$($row.relativePath.Replace('/', '\'))"
            $status.Rows.ContainsKey($key) | Should -Be $true -Because $key
            $status.Rows[$key].状態 | Should -Be $row.status -Because $key
            $status.Rows[$key].抽出版 | Should -Be ([string]$row.extractVersion) -Because $key
        }
    }
}

Describe "見本をそのまま読む（index_compat a）" -Tag Io {
    It "getSearchIndexes・getLegacyIndexState・readSystemIndexState・readStatusFile が、今のコードでそのまま読める" {
        $ws = [Workspace]::new(${sampleWsDir})

        $indexes = @(getSearchIndexes $ws.IndexDir $ws.StatusFile)
        @($indexes.Name) | Should -Contain ${indexName}

        $legacy = getLegacyIndexState $ws.Dir
        $legacy.HasLegacyIndex | Should -Be $false

        $state = readSystemIndexState $ws.SystemIndexStateFile
        $state.Covered.Contains(${indexName}) | Should -Be $true

        $status = readStatusFile $ws.StatusFile
        assertIngestListMatchesExpected $status
    }
}

Describe "システムインデックスの語（index_compat b）" -Tag Io {
    It "writeSystemIndexFolders が作る語が、見本の system_index と同じになる（getSearchGrams・convertToGramText・getPackContentText）" {
        $contentIndexRoot = "${sampleWsDir}\content_index"
        $outRoot = Join-Path $TestDrive "sysidx_out"
        $folders = @("$contentIndexRoot\${indexName}", "$contentIndexRoot\${indexName}\2024")

        [void](writeSystemIndexFolders $folders $contentIndexRoot $outRoot 1)

        foreach ($sub in @("", "\2024")) {
            $goldenPath = "${sampleWsDir}\system_index\${indexName}$sub\${systemIndexFileName}"
            $actualPath = "$outRoot\${indexName}$sub\${systemIndexFileName}"
            [System.IO.File]::Exists($actualPath) | Should -Be $true -Because $actualPath
            [System.IO.File]::ReadAllText($actualPath, [System.Text.Encoding]::ASCII) |
                Should -Be ([System.IO.File]::ReadAllText($goldenPath, [System.Text.Encoding]::ASCII)) -Because $sub
        }
    }
}

Describe "検索結果（index_compat c）" -Tag Io {
    It "newSearchRequest・invokeSearchRequest（searchPackIndex）の結果が、見本（expected.json）と同じになる" {
        assertSearchHitsMatchExpected ${sampleWsDir} ${expected}.searches
    }
}

Describe "取り込み直し（index_compat d・e・f・g）" -Tag Io {
    It "d: 更新の無い見本を取り込み直しても、何も取り込まない（抽出版・取り込み日時は変わらない）" {
        $compat = newCompatRoot "d"
        writeTestSettings $compat.Root @(@{ name = ${indexName}; path = $compat.SourceDir; enabled = $true })
        $before = readTestStatus $compat.Root

        runIndexer $compat.Root | Should -Be 0

        (readTestProgress).Detail | Should -Be "取り込みが必要なファイルはありませんでした"
        $after = readTestStatus $compat.Root
        foreach ($key in $before.Rows.Keys) {
            $after.Rows[$key].取り込み日時 | Should -Be $before.Rows[$key].取り込み日時 -Because $key
            $after.Rows[$key].抽出版 | Should -Be $before.Rows[$key].抽出版 -Because $key
        }
    }

    It "e: 1 つの docx の更新日時を進めると、その 1 件だけを取り込み直す" {
        $compat = newCompatRoot "e"
        writeTestSettings $compat.Root @(@{ name = ${indexName}; path = $compat.SourceDir; enabled = $true })
        $before = readTestStatus $compat.Root
        $touched = "基本.docx"
        (Get-Item -LiteralPath "$($compat.SourceDir)\$touched").LastWriteTime = (Get-Date).AddDays(1)

        runIndexer $compat.Root | Should -Be 0

        $after = readTestStatus $compat.Root
        foreach ($key in $before.Rows.Keys) {
            if ($key -eq "${indexName}\$touched") {
                $after.Rows[$key].取り込み日時 | Should -Not -Be $before.Rows[$key].取り込み日時
            } else {
                $after.Rows[$key].取り込み日時 | Should -Be $before.Rows[$key].取り込み日時 -Because $key
            }
        }
    }

    It "f: xlsx の抽出版を 1 つ下げると、ConfirmTargets の確認で取りやめた間は Excel を起動せず取り込まない" {
        $compat = newCompatRoot "f"
        $statusPath = "$($compat.WorkDir)\ingest_status.tsv"
        $status = readStatusFile $statusPath
        $key = "${indexName}\進捗.xlsx"
        $row = $status.Rows[$key]
        $downgraded = ([int]$row.抽出版) - 1
        $status.Rows[$key] = [pscustomobject]@{ 相対パス = $row.相対パス; 更新日時 = $row.更新日時; サイズ = $row.サイズ; 状態 = $row.状態
            TSV数 = $row.TSV数; 取り込み日時 = $row.取り込み日時; エラー = $row.エラー; 抽出版 = [string]$downgraded }
        writeStatusFile $status.Folders @($status.Rows.Values) $statusPath
        writeTestSettings $compat.Root @(@{ name = ${indexName}; path = $compat.SourceDir; enabled = $true })
        $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
        $cancel = @{ Script = $reporterPath; Pattern = '^\s+if \(!\$ch\.Answered\.WaitOne'; Action = { answerIndexingPlan $channel $null } }

        runIndexer $compat.Root @{ ConfirmTargets = $true } @($cancel) | Should -Be 2

        (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count) | Should -Be $before
        (readTestStatus $compat.Root).Rows[$key].抽出版 | Should -Be ([string]$downgraded)
    }

    It "g: pptx の抽出版を 1 つ下げると、Office 無しでその 1 件だけを取り込み直す" {
        $compat = newCompatRoot "g"
        $statusPath = "$($compat.WorkDir)\ingest_status.tsv"
        $status = readStatusFile $statusPath
        $key = "${indexName}\基本.pptx"
        $row = $status.Rows[$key]
        $downgraded = ([int]$row.抽出版) - 1
        $status.Rows[$key] = [pscustomobject]@{ 相対パス = $row.相対パス; 更新日時 = $row.更新日時; サイズ = $row.サイズ; 状態 = $row.状態
            TSV数 = $row.TSV数; 取り込み日時 = $row.取り込み日時; エラー = $row.エラー; 抽出版 = [string]$downgraded }
        writeStatusFile $status.Folders @($status.Rows.Values) $statusPath
        writeTestSettings $compat.Root @(@{ name = ${indexName}; path = $compat.SourceDir; enabled = $true })
        $before = readTestStatus $compat.Root

        runIndexer $compat.Root | Should -Be 0

        $after = readTestStatus $compat.Root
        foreach ($key2 in $before.Rows.Keys) {
            if ($key2 -eq $key) {
                $after.Rows[$key2].抽出版 | Should -Be $row.抽出版
                $after.Rows[$key2].取り込み日時 | Should -Not -Be $before.Rows[$key2].取り込み日時
            } else {
                $after.Rows[$key2].取り込み日時 | Should -Be $before.Rows[$key2].取り込み日時 -Because $key2
            }
        }
    }
}

Describe "zip のインポート（index_compat h）" -Tag Io {
    It "importIndex で export.zip を空のワークスペースに取り込むと、検索結果・取り込み一覧が見本と同じになる" {
        $root = Join-Path $TestDrive "h"
        $ws = [Workspace]::new("$root\work")
        [System.IO.Directory]::CreateDirectory($ws.Dir) | Out-Null
        $settingsPath = "$root\setting.config"
        writeSettings (newSettings) $settingsPath

        [void](importIndex "${sampleDir}\export.zip" "Overwrite" "" "" $ws $settingsPath)

        $status = readStatusFile $ws.StatusFile
        assertIngestListMatchesExpected $status
        assertSearchHitsMatchExpected $ws.Dir ${expected}.searches
    }
}

Describe "zip のエクスポート（index_compat i）" -Tag Io {
    It "すべての見本について、今の exportIndex で作り直した zip のファイル名の並びが、見本の export.zip と同じになる" {
        foreach ($name in @([System.IO.Directory]::GetDirectories(${compatRoot}) | ForEach-Object { [System.IO.Path]::GetFileName($_) })) {
            $dir = "${compatRoot}\$name"
            $thisExpected = [System.IO.File]::ReadAllText("$dir\expected.json", [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            $root = Join-Path $TestDrive "i_$name"
            Copy-Item -LiteralPath "$dir\ws" -Destination "$root\work" -Recurse -Force
            $ws = [Workspace]::new("$root\work")
            $destZip = "$root\export.zip"

            [void](exportIndex $thisExpected.indexName $destZip $ws "$root\setting.config")

            $actualNames = @()
            $archive = [System.IO.Compression.ZipFile]::OpenRead($destZip)
            try {
                $actualNames = @($archive.Entries | ForEach-Object { $_.FullName } | Sort-Object)
            } finally {
                $archive.Dispose()
            }
            $goldenNames = @()
            $goldenArchive = [System.IO.Compression.ZipFile]::OpenRead("$dir\export.zip")
            try {
                $goldenNames = @($goldenArchive.Entries | ForEach-Object { $_.FullName } | Sort-Object)
            } finally {
                $goldenArchive.Dispose()
            }
            ($actualNames -join "`n") | Should -Be ($goldenNames -join "`n") -Because $name
        }
    }

    It "見本の export.zip の目録（formatVersion）が、今の indexArchiveFormatVersion と同じになる" {
        foreach ($name in @([System.IO.Directory]::GetDirectories(${compatRoot}) | ForEach-Object { [System.IO.Path]::GetFileName($_) })) {
            $dir = "${compatRoot}\$name"
            $archive = [System.IO.Compression.ZipFile]::OpenRead("$dir\export.zip")
            try {
                $entry = $archive.Entries | Where-Object { $_.FullName -ceq ${indexArchiveManifestFileName} } | Select-Object -First 1
                $entry | Should -Not -BeNullOrEmpty -Because $name
                $stream = $entry.Open()
                try {
                    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
                    $manifest = $reader.ReadToEnd() | ConvertFrom-Json
                } finally {
                    $stream.Dispose()
                }
                [int]$manifest.formatVersion | Should -Be ${indexArchiveFormatVersion} -Because $name
            } finally {
                $archive.Dispose()
            }
        }
    }
}
