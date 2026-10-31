# 前の版のファイル（見本・golden。tests/testdata/compat/index/<見本の名前>/）が、今のコードでもそのまま読め、
# 今のコードが作るものと同じになることを確かめる（互換テスト）。見本の数だけ同じ確かめを繰り返す。
# 見本の置き方・作り方は tests/testdata/README.md「前の版のファイル（compat\）」。
# Office は要らない（見本の Word・PowerPoint・Excel は、読み直すときも COM 無しで読める形式のため）。
# f だけ、ConfirmTargets で取りやめたときに Excel を起動していないことを確かめるため、プロセスの数を見る
BeforeDiscovery {
    $compatIndexRoot = Join-Path (Resolve-Path "$PSScriptRoot\..\..").Path "testdata\compat\index"
    # 今の形式（ws\store\ のある見本）と、前の形式（ws\content_index\ だけの見本）に分ける。前の形式の見本は「古い形式」の入力（index_compat j）
    $allSamples = @([System.IO.Directory]::GetDirectories($compatIndexRoot))
    # 分けた塊の見本（v2+split）は、Excel・Word などの取り込み直しの確かめ（d〜g）に合わない中身のため、専用の確かめ（k）で読む
    $samples = @($allSamples | Where-Object { [System.IO.Directory]::Exists("$_\ws\store") -and [System.IO.Path]::GetFileName($_) -ne "v2+split" } | ForEach-Object { @{ Sample = [System.IO.Path]::GetFileName($_) } })
    $oldSamples = @($allSamples | Where-Object { ![System.IO.Directory]::Exists("$_\ws\store") } | ForEach-Object { @{ Sample = [System.IO.Path]::GetFileName($_) } })
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\indexer.ps1"
    . "$PSScriptRoot\..\..\helpers\compat.ps1"

    $reporterPath = "${scriptsDir}\tebunko\indexer\indexing_reporter.ps1"
    ${compatRoot} = "${testDataDir}\compat\index"

    function script:getSample {
        # 見本の場所と expected.json を読んだ結果をまとめる
        param ([string]$name)
        $dir = "${compatRoot}\$name"
        $expected = [System.IO.File]::ReadAllText("$dir\expected.json", [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        return @{ Name = $name; Dir = $dir; WsDir = "$dir\ws"; Expected = $expected; IndexName = $expected.indexName }
    }

    function script:restoreFileTimes {
        # file_times.tsv の記録で、コピー先のファイルの更新日時を見本の記録どおりに戻す
        # （git は取り出すときにファイルの更新日時を今の日時にしてしまうため。tools/make_index_golden.ps1 のコメントも参照）。
        #   source/<相対パス> … ローカルの Ticks（秒まで・DateTimeKind なし）。destRoot\source に戻す
        #   ws/<相対パス>     … UTC の Ticks。destRoot\work に戻す
        param ([string]$destRoot, [string]$timesPath)
        foreach ($line in @([System.IO.File]::ReadAllLines($timesPath, [System.Text.Encoding]::UTF8))) {
            if (!$line -or $line.StartsWith("#")) { continue }
            $cols = $line -split "`t"
            $relative = $cols[0] -replace "/", "\"
            if ($relative.StartsWith("source\")) {
                [System.IO.File]::SetLastWriteTime((Join-Path $destRoot $relative), (New-Object datetime ([long]$cols[1])))
            } elseif ($relative.StartsWith("ws\")) {
                [System.IO.File]::SetLastWriteTimeUtc((Join-Path $destRoot ("work\" + $relative.Substring(3))), (New-Object datetime ([long]$cols[1]), ([System.DateTimeKind]::Utc)))
            } else {
                throw "file_times.tsv の行が想定と違います: $line"
            }
        }
    }

    function script:newCompatRoot {
        # 見本の source/・ws/ を TestDrive にコピーし、更新日時を戻す。root\work は runIndexer・writeTestSettings がそのまま使う置き場所
        param ($s, [string]$name)
        $root = Join-Path $TestDrive "$($s.Name)_$name"
        $sourceDir = Join-Path $root "source"
        $workDir = Join-Path $root "work"
        Copy-Item -LiteralPath "$($s.Dir)\source" -Destination $sourceDir -Recurse -Force
        Copy-Item -LiteralPath $s.WsDir -Destination $workDir -Recurse -Force
        restoreFileTimes $root "$($s.Dir)\file_times.tsv"
        return @{ Root = $root; SourceDir = $sourceDir; WorkDir = $workDir }
    }

    function script:toRelativePath {
        # 検索ヒットの RelDir・Book（インデックス名を含む）から、source/ からの相対パス（/ 区切り）を作る
        param ($s, [string]$relDir, [string]$book)
        $dir = $relDir.Substring($s.IndexName.Length).Trim("\")
        $rel = if ($dir) { "$dir\$book" } else { $book }
        return $rel.Replace("\", "/")
    }

    function script:assertSearchHitsMatchExpected {
        # workDir（ワークスペース）を検索し、見本（expected.json の searches）とヒットが同じになることを確かめる
        param ($s, [string]$workDir)
        $ws = [Workspace]::new($workDir)
        foreach ($search in $s.Expected.searches) {
            $folders = @(@{ Root = $ws.IndexDir; RelPath = ""; Recurse = $true })
            $index = getContentIndexFiles $folders
            $found = searchContentIndex $search.word $index.ContentIndexFiles $true 0 -caseSensitive $false -includeShapes $true -includeComments $true
            $actualHits = @($found.Hits | ForEach-Object {
                [pscustomobject]@{ relativePath = (toRelativePath $s $_.RelDir $_.Book); location = $_.Location; line = $_.Line }
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
        param ($s, $status)
        $status.Rows.Count | Should -Be $s.Expected.ingestList.Count
        foreach ($row in $s.Expected.ingestList) {
            $key = "$($s.IndexName)\$($row.relativePath.Replace('/', '\'))"
            $status.Rows.ContainsKey($key) | Should -Be $true -Because $key
            $status.Rows[$key].状態 | Should -Be $row.status -Because $key
            $status.Rows[$key].抽出版 | Should -Be ([string]$row.extractVersion) -Because $key
        }
    }

    function script:getOutdatedKeys {
        # 見本を作ったときの抽出版が、今のコードの抽出版（getExtractVersion）より前のファイルのキー。
        # 抽出版を上げると、そのファイルは更新が無くても取り込み直しになる（これは仕様。docs/design/index-data/format.md「抽出版を上げたときの扱い」）。
        # 抽出版を上げても見本を作り直さずに済むよう、d の「取り込み直すはずのファイル」はここから求める
        param ($s)
        return @($s.Expected.ingestList | Where-Object { [int]$_.extractVersion -lt (getExtractVersion $_.relativePath) } |
            ForEach-Object { "$($s.IndexName)\$($_.relativePath.Replace('/', '\'))" })
    }

    function script:setTestExtractVersions {
        # 取り込み一覧（ingest_status.tsv）の指定した行の抽出版を書き換える。versions: @{ キー = 抽出版 }
        param ([string]$statusPath, [hashtable]$versions)
        $status = readStatusFile $statusPath
        foreach ($key in $versions.Keys) {
            $row = $status.Rows[$key]
            $status.Rows[$key] = [pscustomobject]@{ 相対パス = $row.相対パス; 更新日時 = $row.更新日時; サイズ = $row.サイズ; 状態 = $row.状態
                TSV数 = $row.TSV数; 取り込み日時 = $row.取り込み日時; エラー = $row.エラー; 抽出版 = [string]$versions[$key] }
        }
        writeStatusFile $status.Folders @($status.Rows.Values) $statusPath
    }

    function script:alignExtractVersions {
        # すべての行の抽出版を今のコードの抽出版にそろえる（1 つのファイルだけが取り込み直しになる場面にするため）
        param ([string]$statusPath)
        $status = readStatusFile $statusPath
        $versions = @{}
        foreach ($key in $status.Rows.Keys) { $versions[$key] = getExtractVersion $status.Rows[$key].相対パス }
        setTestExtractVersions $statusPath $versions
    }

    function script:getFolderOfKey {
        # 取り込み一覧のキー（<インデックス名>\<相対パス>）が入っているフォルダ（インデックス名からの相対）
        param ([string]$key)
        return (Split-Path -Parent $key)
    }

    function script:getFileHashes {
        # フォルダの下の本文インデックス（content_index.*.tsv）の内容の SHA256。@{ フォルダからの相対パス = ハッシュ }
        param ([string]$root)
        $hashes = @{}
        foreach ($file in @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter "${contentIndexFileNamePrefix}.*.tsv")) {
            $hashes[$file.FullName.Substring($root.Length + 1)] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }
        return $hashes
    }

    function script:assertContentIndexUnchanged {
        # 取り込み直したファイル（reingestKeys）の本文インデックス（content_index.<拡張子>.*.tsv）のほかは、内容が変わっていないことを確かめる
        param ([hashtable]$before, [string]$workDir, [string[]]$reingestKeys)
        $after = getFileHashes ([Workspace]::new($workDir)).IndexDir
        $before.Count | Should -BeGreaterThan 0
        $after.Count | Should -Be $before.Count
        $exempt = @($reingestKeys | ForEach-Object {
            $folder = Split-Path -Parent $_
            $ext = [System.IO.Path]::GetExtension($_).TrimStart(".").ToLowerInvariant()
            "$folder\${contentIndexFileNamePrefix}.$ext."
        })
        foreach ($path in $before.Keys) {
            if (@($exempt | Where-Object { $path.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0) { continue }
            $after[$path] | Should -Be $before[$path] -Because "$path の内容が変わってはいけない"
        }
    }

    function script:findIngestKey {
        # 見本の取り込み一覧から、指定した拡張子の最初のファイルのキーを返す（無ければ $null）
        param ($s, [string[]]$extensions)
        $row = $s.Expected.ingestList | Where-Object { $extensions -contains [System.IO.Path]::GetExtension($_.relativePath).ToLowerInvariant() } | Select-Object -First 1
        if (!$row) { return $null }
        return "$($s.IndexName)\$($row.relativePath.Replace('/', '\'))"
    }
}

Describe "見本をそのまま読む（index_compat a）" -Tag Io {
    It "<Sample>: getSearchIndexes・getLegacyIndexState・readStatusFile が、今のコードでそのまま読める" -TestCases $samples {
        $s = getSample $Sample
        $ws = [Workspace]::new($s.WsDir)

        $indexes = @(getSearchIndexes $ws.IndexDir $ws.StatusFile)
        @($indexes.Name) | Should -Contain $s.IndexName

        $legacy = getLegacyIndexState $ws.Dir
        $legacy.HasLegacyIndex | Should -Be $false

        # 今の形式の見本は、目録（store\format.tsv）どおりに「作り終えた」ものとして検索の対象になる
        $format = getWorkspaceIndexFormat $ws
        $format.State | Should -Be "Current"
        $format.Indexes[$s.IndexName].State | Should -Be "Current"
        $format.Indexes[$s.IndexName].Searchable | Should -Be $true

        $status = readStatusFile $ws.StatusFile
        assertIngestListMatchesExpected $s $status
    }
}

Describe "検索結果（index_compat c）" -Tag Io {
    It "<Sample>: getContentIndexFiles・searchContentIndex の結果が、見本（expected.json）と同じになる" -TestCases $samples {
        $s = getSample $Sample
        assertSearchHitsMatchExpected $s $s.WsDir
    }
}

Describe "取り込み直し（index_compat d・e・f・g）" -Tag Io {
    It "<Sample> d: 更新の無い見本を取り込み直しても、抽出版が今より前のファイルのほかは取り込まず" -TestCases $samples {
        $s = getSample $Sample
        $compat = newCompatRoot $s "d"
        writeTestSettings $compat.Root @(@{ name = $s.IndexName; path = $compat.SourceDir; enabled = $true })
        $before = readTestStatus $compat.Root
        $contentBefore = getFileHashes ([Workspace]::new($compat.WorkDir)).IndexDir
        $outdated = @(getOutdatedKeys $s)

        if ($outdated.Count -eq 0) {
            runIndexer $compat.Root | Should -Be 0
            (readTestProgress).Detail | Should -Be "更新が必要なファイルはありませんでした"
        } else {
            # 抽出版を上げた後は、取り込み直しに Office を使うことがある（xlsx など）。Office の無い環境でも流せるよう、
            # 取り込みの計画を確かめて取りやめる（実際に取り込み直す場面は e・g）
            $global:compatPlanSeen = $null
            $cancel = @{ Script = $reporterPath; Pattern = '^\s+if \(!\$ch\.Answered\.WaitOne'; Action = { $global:compatPlanSeen = @($channel.Plan); answerIndexingPlan $channel $null } }
            runIndexer $compat.Root @{ ConfirmTargets = $true } @($cancel) | Should -Be 2
            $plan = @($global:compatPlanSeen)
            $plan.Count | Should -Be 1
            $plan[0].更新あり | Should -Be $outdated.Count
            $plan[0].新規 | Should -Be 0
            $plan[0].前回未完了 | Should -Be 0
            $plan[0].インデックスなし | Should -Be 0
            $plan[0].取り込み対象 | Should -Be $outdated.Count
        }
        $after = readTestStatus $compat.Root
        $after.Rows.Count | Should -Be $before.Rows.Count
        foreach ($key in $before.Rows.Keys) {
            $after.Rows[$key].取り込み日時 | Should -Be $before.Rows[$key].取り込み日時 -Because $key
            $after.Rows[$key].抽出版 | Should -Be $before.Rows[$key].抽出版 -Because $key
        }
        assertContentIndexUnchanged $contentBefore $compat.WorkDir @()
    }

    It "<Sample> e: 1 つの docx の更新日時を進めると、その 1 件だけを取り込み直し" -TestCases $samples {
        $s = getSample $Sample
        $key = findIngestKey $s @(".docx")
        if (!$key) { Set-ItResult -Skipped -Because "この見本に docx が無い"; return }
        $compat = newCompatRoot $s "e"
        alignExtractVersions ([Workspace]::new($compat.WorkDir)).StatusFile
        writeTestSettings $compat.Root @(@{ name = $s.IndexName; path = $compat.SourceDir; enabled = $true })
        $before = readTestStatus $compat.Root
        $contentBefore = getFileHashes ([Workspace]::new($compat.WorkDir)).IndexDir
        $relative = $key.Substring($s.IndexName.Length + 1)
        (Get-Item -LiteralPath "$($compat.SourceDir)\$relative").LastWriteTime = (Get-Date).AddDays(1)

        runIndexer $compat.Root | Should -Be 0

        assertContentIndexUnchanged $contentBefore $compat.WorkDir @($key)
        assertSearchHitsMatchExpected $s $compat.WorkDir

        $after = readTestStatus $compat.Root
        foreach ($rowKey in $before.Rows.Keys) {
            if ($rowKey -eq $key) {
                $after.Rows[$rowKey].取り込み日時 | Should -Not -Be $before.Rows[$rowKey].取り込み日時
            } else {
                $after.Rows[$rowKey].取り込み日時 | Should -Be $before.Rows[$rowKey].取り込み日時 -Because $rowKey
            }
        }
    }

    It "<Sample> f: xlsx の抽出版を 1 つ下げると、ConfirmTargets の確認で取りやめた間は Excel を起動せず取り込まない" -TestCases $samples {
        $s = getSample $Sample
        $key = findIngestKey $s @(".xlsx", ".xlsm")
        if (!$key) { Set-ItResult -Skipped -Because "この見本に xlsx が無い"; return }
        $compat = newCompatRoot $s "f"
        $statusPath = ([Workspace]::new($compat.WorkDir)).StatusFile
        alignExtractVersions $statusPath
        $downgraded = (getExtractVersion $key) - 1
        setTestExtractVersions $statusPath @{ $key = $downgraded }
        writeTestSettings $compat.Root @(@{ name = $s.IndexName; path = $compat.SourceDir; enabled = $true })
        $beforeProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
        $contentBefore = getFileHashes ([Workspace]::new($compat.WorkDir)).IndexDir
        $cancel = @{ Script = $reporterPath; Pattern = '^\s+if \(!\$ch\.Answered\.WaitOne'; Action = { answerIndexingPlan $channel $null } }

        runIndexer $compat.Root @{ ConfirmTargets = $true } @($cancel) | Should -Be 2

        assertContentIndexUnchanged $contentBefore $compat.WorkDir @()

        (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count) | Should -Be $beforeProcesses
        (readTestStatus $compat.Root).Rows[$key].抽出版 | Should -Be ([string]$downgraded)
    }

    It "<Sample> g: pptx の抽出版を 1 つ下げると、Office 無しでその 1 件だけを取り込み直し" -TestCases $samples {
        $s = getSample $Sample
        $key = findIngestKey $s @(".pptx")
        if (!$key) { Set-ItResult -Skipped -Because "この見本に pptx が無い"; return }
        $compat = newCompatRoot $s "g"
        $statusPath = ([Workspace]::new($compat.WorkDir)).StatusFile
        alignExtractVersions $statusPath
        setTestExtractVersions $statusPath @{ $key = ((getExtractVersion $key) - 1) }
        writeTestSettings $compat.Root @(@{ name = $s.IndexName; path = $compat.SourceDir; enabled = $true })
        $before = readTestStatus $compat.Root
        $contentBefore = getFileHashes ([Workspace]::new($compat.WorkDir)).IndexDir

        runIndexer $compat.Root | Should -Be 0

        assertContentIndexUnchanged $contentBefore $compat.WorkDir @($key)
        $after = readTestStatus $compat.Root
        foreach ($rowKey in $before.Rows.Keys) {
            if ($rowKey -eq $key) {
                $after.Rows[$rowKey].抽出版 | Should -Be ([string](getExtractVersion $key))
                $after.Rows[$rowKey].取り込み日時 | Should -Not -Be $before.Rows[$rowKey].取り込み日時
            } else {
                $after.Rows[$rowKey].取り込み日時 | Should -Be $before.Rows[$rowKey].取り込み日時 -Because $rowKey
            }
        }
    }
}

Describe "zip のインポート（index_compat h）" -Tag Io {
    It "<Sample>: importIndex で export.zip を空のワークスペースに取り込むと、検索結果・取り込み一覧が見本と同じになる" -TestCases $samples {
        $s = getSample $Sample
        $root = Join-Path $TestDrive "h_$Sample"
        $ws = [Workspace]::new("$root\work")
        [System.IO.Directory]::CreateDirectory($ws.Dir) | Out-Null
        $settingsPath = "$root\setting.config"
        writeSettings (newSettings) $settingsPath

        [void](importIndex "$($s.Dir)\export.zip" "Overwrite" "" "" $ws $settingsPath)

        $status = readStatusFile $ws.StatusFile
        assertIngestListMatchesExpected $s $status
        assertSearchHitsMatchExpected $s $ws.Dir
        # 取り込んだインデックスは、目録に「作り終えた」で載る
        (readIndexFormat $ws.FormatFile).Indexes[$s.IndexName] | Should -Be ${catalogRowFinished}
    }
}

Describe "zip のエクスポート（index_compat i）" -Tag Io {
    It "<Sample>: 今の exportIndex で作り直した zip のファイル名の並びが、見本の export.zip と同じになる" -TestCases $samples {
        $s = getSample $Sample
        $root = Join-Path $TestDrive "i_$Sample"
        Copy-Item -LiteralPath $s.WsDir -Destination "$root\work" -Recurse -Force
        $ws = [Workspace]::new("$root\work")
        $destZip = "$root\export.zip"

        [void](exportIndex $s.IndexName $destZip $ws "$root\setting.config")

        (getZipEntryNames $destZip) -join "`n" | Should -Be ((getZipEntryNames "$($s.Dir)\export.zip") -join "`n")
    }

    It "<Sample>: 見本の export.zip の目録（formatVersion）が、今の indexArchiveFormatVersion と同じになる" -TestCases $samples {
        $s = getSample $Sample
        $text = readZipEntryText "$($s.Dir)\export.zip" ${indexArchiveManifestFileName}
        $text | Should -Not -BeNullOrEmpty
        [int]($text | ConvertFrom-Json).formatVersion | Should -Be ${indexArchiveFormatVersion}
    }
}

Describe "古い形式（index_compat j）" -Tag Io {
    It "<Sample>: 前の形式の見本は「古い形式」として扱い、検索の対象に入れない（読まず、消さない）" -TestCases $oldSamples {
        $s = getSample $Sample
        $compat = newCompatRoot $s "j1"
        $ws = [Workspace]::new($compat.WorkDir)

        $format = getWorkspaceIndexFormat $ws
        $format.State | Should -Be "Old"
        $format.Message | Should -Be (getIndexRebuildMessage)
        $format.Indexes[$s.IndexName].State | Should -Be "Old"
        $format.Indexes[$s.IndexName].Searchable | Should -Be $false
        $format.Indexes[$s.IndexName].Message | Should -Be (getIndexRebuildMessage)
        @(getSearchIndexes $ws.IndexDir $ws.StatusFile).Count | Should -Be 0

        # 前の形式のファイルは、そのまま残っている（作り直しが終わるまで消さない）
        [System.IO.Directory]::Exists($ws.OldIndexDir) | Should -Be $true
        [System.IO.File]::Exists($ws.OldStatusFile) | Should -Be $true
    }

    It "<Sample>: 前の形式の見本の本文インデックスは、1 行目が今の版ではない" -TestCases $oldSamples {
        $s = getSample $Sample
        $ws = [Workspace]::new($s.WsDir)
        $file = @(Get-ChildItem -LiteralPath $ws.OldIndexDir -Recurse -File -Filter ${contentIndexFilePattern} | Select-Object -First 1)
        $file.Count | Should -Be 1
        testIndexFileCurrentVersion $file[0].FullName | Should -Be $false
    }

    It "<Sample>: ［更新］で全部を取り込み直すと、新しい置き場に作り、古い置き場を消し、検索の結果が見本と同じになる" -TestCases $oldSamples {
        $s = getSample $Sample
        if ($null -ne (findIngestKey $s @(".xlsx", ".xlsm", ".xls", ".doc", ".ppt"))) { Set-ItResult -Skipped -Because "この見本の取り込みに Office が要る"; return }
        $compat = newCompatRoot $s "j2"
        writeTestSettings $compat.Root @(@{ name = $s.IndexName; path = $compat.SourceDir; enabled = $true })
        $ws = [Workspace]::new($compat.WorkDir)

        runIndexer $compat.Root | Should -Be 0

        (readIndexFormat $ws.FormatFile).Indexes[$s.IndexName] | Should -Be ${catalogRowFinished}
        [System.IO.Directory]::Exists($ws.OldIndexDir) | Should -Be $false
        [System.IO.File]::Exists($ws.OldStatusFile) | Should -Be $false
        (getWorkspaceIndexFormat $ws).State | Should -Be "Current"
        $status = readStatusFile $ws.StatusFile
        $status.Rows.Count | Should -Be $s.Expected.ingestList.Count
        assertSearchHitsMatchExpected $s $compat.WorkDir
    }
}

Describe "分けた塊（版 3）の見本（index_compat k）" -Tag Io {
    It "v2+split: 見本をそのまま読み、検索の結果が見本と同じで、行番号が場所の中の通しの番号になる" {
        $s = getSample "v2+split"
        $ws = [Workspace]::new($s.WsDir)

        (getWorkspaceIndexFormat $ws).Indexes[$s.IndexName].State | Should -Be "Current"
        assertIngestListMatchesExpected $s (readStatusFile $ws.StatusFile)
        assertSearchHitsMatchExpected $s $s.WsDir

        $files = @(Get-ChildItem -LiteralPath $ws.IndexDir -Recurse -File -Filter "content_index.txt.*.tsv" | Sort-Object Name)
        $files.Count | Should -BeGreaterThan 2
        foreach ($file in $files) { getContentIndexTextVersion (readContentIndexText $file.FullName) | Should -Be ${contentIndexSplitVersion} }
        getContentIndexTextVersion (readContentIndexText (Get-ChildItem -LiteralPath $ws.IndexDir -Recurse -File -Filter "content_index.csv.001.tsv").FullName) | Should -Be ${contentIndexVersion}

        $index = getContentIndexFiles @(@{ Root = $ws.IndexDir; RelPath = ""; Recurse = $true })
        $found = searchContentIndex "行100の説明" $index.ContentIndexFiles $true 0
        @($found.Hits).Count | Should -Be 1
        $found.Hits[0].LineNumber | Should -Be 100
        @($found.Skipped).Count | Should -Be 0
    }

    It "v2+split: 塊の切れ目をまたぐ行を、場所の中の通しの番号で読める（readContentIndexLinesAcrossParts）" {
        $s = getSample "v2+split"
        $ws = [Workspace]::new($s.WsDir)
        $files = @(Get-ChildItem -LiteralPath $ws.IndexDir -Recurse -File -Filter "content_index.txt.*.tsv" | Sort-Object Name)
        $second = $files[2].FullName
        $secondFirst = (readContentIndexPlaces (readContentIndexText $second))[0].FirstLine
        $secondFirst | Should -BeGreaterThan 1

        $result = readContentIndexLinesAcrossParts $second "ログ.txt" "本文" ($secondFirst - 2) 4

        @($result.Rows | ForEach-Object { $_.LineNumber }) | Should -Be @(($secondFirst - 2)..($secondFirst + 1))
        $result.Rows[0].Line | Should -BeLike "$($secondFirst - 2),分割の単価,*"
        $result.Rows[3].Line | Should -BeLike "$($secondFirst + 1),分割の単価,*"
    }

    It "v2+split: importIndex で export.zip を空のワークスペースに取り込むと、検索結果・取り込み一覧が見本と同じになる" {
        $s = getSample "v2+split"
        $root = Join-Path $TestDrive "k_split"
        $ws = [Workspace]::new("$root\work")
        [System.IO.Directory]::CreateDirectory($ws.Dir) | Out-Null
        $settingsPath = "$root\setting.config"
        writeSettings (newSettings) $settingsPath

        [void](importIndex "$($s.Dir)\export.zip" "Overwrite" "" "" $ws $settingsPath)

        assertIngestListMatchesExpected $s (readStatusFile $ws.StatusFile)
        assertSearchHitsMatchExpected $s $ws.Dir
        @(Get-ChildItem -LiteralPath $ws.IndexDir -Recurse -File -Filter "content_index.txt.*.tsv").Count | Should -BeGreaterThan 2
    }
}
