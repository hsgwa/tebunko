# インデックスのエクスポート・インポート（tebunko\index\index_archive.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"
    . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"

    function newIndexFixture {
        # ワークスペース dir にインデックス name を1つ作る（本文インデックスのファイル1つ・取り込み一覧の行・設定への登録）。
        # 返すもの: @{ Workspace; SettingsPath; RelPath; Text }
        param (
            [string]$dir,
            [string]$name,
            [string]$sourcePath,
            [string]$relPath = "見積\A社.xlsx",
            [string]$text = "品名`t数量`n鉛筆`t10`n",
            [string]$settingsPath = "$dir\setting.config"
        )

        $ws = [Workspace]::new($dir)
        $indexDir = Join-Path $ws.IndexDir $name
        $relDir = Split-Path $relPath -Parent
        $packDir = if ($relDir) { Join-Path $indexDir $relDir } else { $indexDir }
        $book = @{ Name = (Split-Path $relPath -Leaf); Places = @(@{ Place = "Sheet1"; Text = $text }) }
        $packText = convertToPackText @($book)
        $packPath = Join-Path $packDir (getPackFileName (getPackExtension $relPath))
        [System.IO.Directory]::CreateDirectory((toLongPath $packDir)) | Out-Null
        [System.IO.File]::WriteAllText((toLongPath $packPath), $packText, (New-Object System.Text.UTF8Encoding($true)))

        writeStatusFile @([pscustomobject]@{ Path = $sourcePath; Name = $name }) @(
            (newStatusRow "$name\$relPath" "2024/01/01 00:00:00" "100" ${stateDone} "1" "2024/01/01 00:00:00" "" ([string](getExtractVersion $relPath)))
        ) $ws.StatusFile
        writeTargetFolders @([pscustomobject]@{ Name = $name; Path = $sourcePath; Enabled = $true }) $settingsPath

        return @{ Workspace = $ws; SettingsPath = $settingsPath; RelPath = $relPath; Text = $text; PackPath = $packPath }
    }

    function getStatusRowArray {
        # 取り込み一覧の行（readStatusFile の Rows）を配列にする（値の集合を @() に渡すと型の不一致になる環境があるため、1 件ずつ足す）
        param ($status)

        $list = New-Object System.Collections.Generic.List[object]
        foreach ($row in $status.Rows.Values) {
            $list.Add($row)
        }
        return , $list.ToArray()
    }

    function readSearchWord {
        # ワークスペース ws のインデックスから word を検索し、当たった行を返す（pack_search.ps1）
        param ($ws, [string]$word)

        $found = getIndexPackFiles @($ws.IndexDir)
        $regex = New-Object regex ([regex]::Escape($word))
        return @(searchPackFiles $found.Packs 0 $found.Packs.Count $regex -1 $regex "lines")
    }
}

Describe "exportIndex" -Tag Io {
    It "止める: インデックスが無ければ例外にする" {
        $ws = newTestWorkspace @{} "$TestDrive\export_none"

        { exportIndex "無い名前" "$TestDrive\export_none\out.zip" $ws "$TestDrive\export_none\setting.config" } | Should -Throw "*見つかりません*"
        Test-Path -LiteralPath "$TestDrive\export_none\out.zip" | Should -Be $false
    }

    It "止める: 取り込み一覧に記録が無ければ例外にする（フォルダだけあるとき）" {
        $dir = "$TestDrive\export_norecord"
        $ws = [Workspace]::new($dir)
        [System.IO.Directory]::CreateDirectory((Join-Path $ws.IndexDir "営業")) | Out-Null

        { exportIndex "営業" "$dir\out.zip" $ws "$dir\setting.config" } | Should -Throw "*記録がありません*"
    }

    It "止める: 入れる前の TSV が残っていれば例外にし、保存先にも .tmp にもファイルを残さない" {
        $fixture = newIndexFixture "$TestDrive\export_pending" "営業" "C:\共有\営業部"
        # 元のファイルごとのフォルダ（入れる前の TSV）を残す
        New-Item -ItemType Directory -Path (Join-Path $fixture.Workspace.IndexDir "営業\見積\B社.xlsx") -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $fixture.Workspace.IndexDir "営業\見積\B社.xlsx\Sheet1.tsv") -Value "本文" -Encoding UTF8

        $dest = "$TestDrive\export_pending\out.zip"
        { exportIndex "営業" $dest $fixture.Workspace $fixture.SettingsPath } | Should -Throw "*取り込みの途中*"
        Test-Path -LiteralPath $dest | Should -Be $false
        Test-Path -LiteralPath "${dest}.tmp" | Should -Be $false
    }

    It "止める: インデックス作成のロックを取れなければ例外にする" {
        $fixture = newIndexFixture "$TestDrive\export_locked" "営業" "C:\共有\営業部"
        $lock = newAppMutex "indexer" $fixture.Workspace.Dir
        $lock.Acquired | Should -Be $true
        try {
            { exportIndex "営業" "$TestDrive\export_locked\out.zip" $fixture.Workspace $fixture.SettingsPath } | Should -Throw "*インデックス作成中*"
        } finally {
            $lock.Mutex.ReleaseMutex()
            $lock.Mutex.Dispose()
        }
    }

    It "zip に目録・取り込み一覧・本文インデックスのファイルを書き出す（相対パスの名前を外す）" {
        $fixture = newIndexFixture "$TestDrive\export_ok" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\export_ok\out.zip"

        $result = exportIndex "営業" $dest $fixture.Workspace $fixture.SettingsPath

        $result.Name | Should -Be "営業"
        $result.Path | Should -Be $dest
        $result.Files | Should -Be 2
        Test-Path -LiteralPath $dest | Should -Be $true
        Test-Path -LiteralPath "${dest}.tmp" | Should -Be $false

        $archive = [System.IO.Compression.ZipFile]::Open($dest, [System.IO.Compression.ZipArchiveMode]::Read, [System.Text.Encoding]::UTF8)
        try {
            @($archive.Entries | ForEach-Object { $_.FullName }) | Sort-Object | Should -Be @(
                "content_index/見積/content_index.xlsx.001.tsv", "ingest_status.tsv", "tebunko-index.json"
            )
            $manifestEntry = $archive.GetEntry("tebunko-index.json")
            $reader = New-Object System.IO.StreamReader($manifestEntry.Open())
            $manifest = ConvertFrom-Json $reader.ReadToEnd()
            $reader.Dispose()
            $manifest.indexName | Should -Be "営業"
            $manifest.sourceFolder | Should -Be "C:\共有\営業部"
            $manifest.formatVersion | Should -Be 1

            $statusEntry = $archive.GetEntry("ingest_status.tsv")
            $statusReader = New-Object System.IO.StreamReader($statusEntry.Open(), [System.Text.Encoding]::UTF8)
            $statusText = $statusReader.ReadToEnd()
            $statusReader.Dispose()
            $statusText | Should -Match ([regex]::Escape("見積\A社.xlsx"))
            $statusText | Should -Not -Match ([regex]::Escape("営業\"))
        } finally {
            $archive.Dispose()
        }
    }
}

Describe "readIndexArchiveInfo" -Tag Io {
    It "目録を読んで @{ IndexName; SourceFolder; Files; Bytes } を返す（インポートしない）" {
        $fixture = newIndexFixture "$TestDrive\info_ok" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\info_ok\out.zip"
        exportIndex "営業" $dest $fixture.Workspace $fixture.SettingsPath | Out-Null

        $info = readIndexArchiveInfo $dest

        $info.IndexName | Should -Be "営業"
        $info.SourceFolder | Should -Be "C:\共有\営業部"
        $info.Files | Should -Be 2
        $info.Bytes | Should -BeGreaterThan 0
        $info.FormatVersion | Should -Be 1
        Test-Path -LiteralPath (Join-Path $fixture.Workspace.IndexDir "営業") | Should -Be $true
    }

    It "壊れた・別の版の zip は例外にする" -TestCases @(
        @{ label = "zip でないファイル" }
        @{ label = "目録が無い" }
        @{ label = "formatVersion が新しい" }
    ) {
        param ($label)
        $dir = "$TestDrive\info_bad_$([Guid]::NewGuid().ToString('N'))"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $path = "$dir\bad.zip"
        switch ($label) {
            "zip でないファイル" {
                Set-Content -LiteralPath $path -Value "これは zip ではありません" -Encoding UTF8
            }
            "目録が無い" {
                $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
                $archive.Dispose()
            }
            "formatVersion が新しい" {
                $fixture = newIndexFixture $dir "営業" "C:\共有\営業部"
                exportIndex "営業" $path $fixture.Workspace $fixture.SettingsPath | Out-Null
                $newPath = "$dir\bad2.zip"
                Copy-Item -LiteralPath $path -Destination $newPath
                $archive = [System.IO.Compression.ZipFile]::Open($newPath, [System.IO.Compression.ZipArchiveMode]::Update, [System.Text.Encoding]::UTF8)
                try {
                    $entry = $archive.GetEntry("tebunko-index.json")
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    $manifest = ConvertFrom-Json $reader.ReadToEnd()
                    $reader.Dispose()
                    $entry.Delete()
                    $manifest.formatVersion = 2
                    $newEntry = $archive.CreateEntry("tebunko-index.json")
                    $writer = New-Object System.IO.StreamWriter($newEntry.Open())
                    $writer.Write((ConvertTo-Json $manifest -Depth 5))
                    $writer.Dispose()
                } finally {
                    $archive.Dispose()
                }
                $path = $newPath
            }
        }

        { readIndexArchiveInfo $path } | Should -Throw
    }
}

Describe "importIndex" -Tag Io {
    It "往復: 別のワークスペース・別の元のフォルダへインポートすると検索でき、場所が新しい元のフォルダになる" {
        $fixtureA = newIndexFixture "$TestDrive\roundtrip_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\roundtrip.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\roundtrip_b")
        $settingsB = "$TestDrive\roundtrip_b\setting.config"
        $result = importIndex $dest ${importCollisionRename} "" "D:\別の場所\営業部" $wsB $settingsB

        $result.Name | Should -Be "営業"
        $result.SourcePath | Should -Be "D:\別の場所\営業部"
        $result.Files | Should -Be 2

        $hits = readSearchWord $wsB "鉛筆"
        $hits.Count | Should -Be 1
        $location = getSourceLocation $hits[0]
        $location.Known | Should -Be $true
        $location.Folder | Should -Be "D:\別の場所\営業部"
        (Join-Path $location.Folder (Join-Path $location.Rest "A社.xlsx")) | Should -Be "D:\別の場所\営業部\見積\A社.xlsx"
    }

    It "インポートしたインデックスに NotContentIndexed が付く（Directory.Move は親の属性を継がず、根に付いたあとの -Recurse は下へ降りないため）" {
        $fixtureA = newIndexFixture "$TestDrive\attr_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\attr.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        # 一度インデックス作成を通った作業フォルダ（content_index の根に属性が付いている）へ入れる
        $wsB = [Workspace]::new("$TestDrive\attr_b")
        [System.IO.Directory]::CreateDirectory($wsB.IndexDir) | Out-Null
        (setNotContentIndexed $wsB.IndexDir -Recurse).Ok | Should -BeTrue
        $result = importIndex $dest ${importCollisionRename} "" "D:\別の場所\営業部" $wsB "$TestDrive\attr_b\setting.config"

        $flag = [System.IO.FileAttributes]::NotContentIndexed
        $importedDir = Join-Path $wsB.IndexDir $result.Name
        ([System.IO.File]::GetAttributes($importedDir) -band $flag) | Should -Not -Be 0
        $files = @(Get-ChildItem -LiteralPath $importedDir -Recurse -Force)
        $files.Count | Should -BeGreaterThan 0
        foreach ($file in $files) {
            ([System.IO.File]::GetAttributes($file.FullName) -band $flag) | Should -Not -Be 0
        }
    }

    It "取り込み一覧の形: <label>。見出しの行があり、クロール対象フォルダの行は見出しの前にあり、getIndexNameMap が名前を返す" -TestCases @(
        @{ label = "空のワークスペースへ"; others = @() }
        @{ label = "ほかのインデックスがあるワークスペースへ"; others = @("経理", "総務") }
    ) {
        param ($label, $others)
        $fixtureA = newIndexFixture "$TestDrive\shape_a_$($others.Count)" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\shape_$($others.Count).zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $dirB = "$TestDrive\shape_b_$($others.Count)"
        $wsB = [Workspace]::new($dirB)
        $settingsB = "$dirB\setting.config"
        if ($others.Count -gt 0) {
            $first = newIndexFixture $dirB $others[0] "C:\$($others[0])" "見積\B社.xlsx"
            $rows = @(getStatusRowArray (readStatusFile $wsB.StatusFile))
            $extra = newIndexFixture "$TestDrive\shape_x_$($others.Count)" $others[1] "C:\$($others[1])" "見積\C社.xlsx"
            $folders = @((readStatusFile $wsB.StatusFile).Folders.ToArray()) + @([pscustomobject]@{ Path = "C:\$($others[1])"; Name = $others[1] })
            $rows += @(getStatusRowArray (readStatusFile $extra.Workspace.StatusFile))
            writeStatusFile $folders $rows $wsB.StatusFile
        }

        importIndex $dest ${importCollisionRename} "" "D:\別の場所\営業部" $wsB $settingsB | Out-Null

        $lines = @(readStatusLines $wsB.StatusFile)
        $header = ${statusColumns} -join "`t"
        $headerIndex = [array]::IndexOf($lines, $header)
        $headerIndex | Should -BeGreaterOrEqual 0
        @($lines | Select-Object -First $headerIndex | Where-Object { $_.StartsWith(${statusFolderKey} + "`t") }).Count | Should -Be (1 + $others.Count)
        @($lines | Select-Object -Skip ($headerIndex + 1) | Where-Object { $_.StartsWith(${statusFolderKey} + "`t") }).Count | Should -Be 0

        $map = getIndexNameMap $wsB.StatusFile
        $map["営業"] | Should -Be "D:\別の場所\営業部"
        foreach ($other in $others) {
            $map[$other] | Should -Be "C:\$other"
        }
        (readStatusFile $wsB.StatusFile).Rows.ContainsKey("営業\見積\A社.xlsx") | Should -Be $true
    }

    It "取り込み一覧の形: 上書きでは前の行・前のクロール対象フォルダの行が残らず、ほかのインデックスの行は残る" {
        $fixtureA = newIndexFixture "$TestDrive\shape_ow_a" "営業" "C:\共有\営業部" "見積\A社.xlsx"
        $dest = "$TestDrive\shape_ow.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $fixtureB = newIndexFixture "$TestDrive\shape_ow_b" "営業" "C:\前の場所" "前の資料\旧.xlsx"
        $status = readStatusFile $fixtureB.Workspace.StatusFile
        $rows = @(getStatusRowArray $status) + @(newStatusRow "経理\月次.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} "1" "2024/01/01 00:00:00" "" "")
        $folders = @($status.Folders.ToArray()) + @([pscustomobject]@{ Path = "C:\経理"; Name = "経理" })
        writeStatusFile $folders $rows $fixtureB.Workspace.StatusFile

        importIndex $dest ${importCollisionOverwrite} "" "C:\新しい場所" $fixtureB.Workspace $fixtureB.SettingsPath | Out-Null

        $after = readStatusFile $fixtureB.Workspace.StatusFile
        @($after.Rows.Keys | Sort-Object) | Should -Be @("営業\見積\A社.xlsx", "経理\月次.xlsx")
        @($after.Folders | ForEach-Object { "$($_.Name)|$($_.Path)" } | Sort-Object) | Should -Be @("営業|C:\新しい場所", "経理|C:\経理")
    }

    It "取り込み一覧の形: 取り込み一覧のエラー列にタブ・改行があっても 1 行 1 件を保つ" {
        $fixtureA = newIndexFixture "$TestDrive\shape_err_a" "営業" "C:\共有\営業部"
        $status = readStatusFile $fixtureA.Workspace.StatusFile
        $status.Rows["営業\見積\A社.xlsx"].エラー = "開けません`tまたは`r`n壊れています"
        $status.Rows["営業\見積\A社.xlsx"].状態 = ${stateFailed}
        writeStatusFile $status.Folders.ToArray() (getStatusRowArray $status) $fixtureA.Workspace.StatusFile
        $dest = "$TestDrive\shape_err.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\shape_err_b")
        importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB "$($wsB.Dir)\setting.config" | Out-Null

        $row = (readStatusFile $wsB.StatusFile).Rows["営業\見積\A社.xlsx"]
        $row.状態 | Should -Be ${stateFailed}
        $row.エラー | Should -Be "開けません または 壊れています"
    }

    It "その後 B でインデクサが動かすシステムインデックスの更新で、system_index の txt ができ、対応済みになる" {
        $fixtureA = newIndexFixture "$TestDrive\sysidx_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\sysidx.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null
        $wsB = [Workspace]::new("$TestDrive\sysidx_b")
        importIndex $dest ${importCollisionRename} "" "D:\別の場所\営業部" $wsB "$($wsB.Dir)\setting.config" | Out-Null
        Test-Path -LiteralPath $wsB.SystemIndexDir | Should -Be $false

        $result = updateSystemIndexes $wsB.IndexDir $wsB.SystemIndexDir $wsB.SystemIndexStateFile { $false }

        $result.Built | Should -BeGreaterThan 0
        $result.Unfinished | Should -Be 0
        @(Get-ChildItem -LiteralPath $wsB.SystemIndexDir -Recurse -Filter "*.txt").Count | Should -BeGreaterThan 0
        (readSystemIndexState $wsB.SystemIndexStateFile).Covered.Contains("営業") | Should -Be $true
    }

    It "上書き: 前の行・前のシステムインデックスが残らず、設定の並びが変わらない" {
        $fixtureA = newIndexFixture "$TestDrive\ow_a" "営業" "C:\共有\営業部" "見積\A社.xlsx"
        $dest = "$TestDrive\ow.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $fixtureB = newIndexFixture "$TestDrive\ow_b" "営業" "C:\前の場所" "前の資料\旧.xlsx"
        $wsB = $fixtureB.Workspace
        writeTargetFolders @(
            [pscustomobject]@{ Name = "経理"; Path = "C:\経理"; Enabled = $true }
            [pscustomobject]@{ Name = "営業"; Path = "C:\前の場所"; Enabled = $true }
            [pscustomobject]@{ Name = "総務"; Path = "C:\総務"; Enabled = $true }
        ) $fixtureB.SettingsPath
        # 前のシステムインデックス（営業の txt と、対応済みの記録）
        [void](updateSystemIndexes $wsB.IndexDir $wsB.SystemIndexDir $wsB.SystemIndexStateFile { $false })
        Test-Path -LiteralPath (Join-Path $wsB.SystemIndexDir "営業") | Should -Be $true

        importIndex $dest ${importCollisionOverwrite} "" "C:\新しい場所" $wsB $fixtureB.SettingsPath | Out-Null

        (readStatusFile $wsB.StatusFile).Rows.ContainsKey("営業\前の資料\旧.xlsx") | Should -Be $false
        (readStatusFile $wsB.StatusFile).Rows.ContainsKey("営業\見積\A社.xlsx") | Should -Be $true
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業\前の資料") | Should -Be $false
        Test-Path -LiteralPath (Join-Path $wsB.SystemIndexDir "営業") | Should -Be $false
        (readSystemIndexState $wsB.SystemIndexStateFile).Covered.Contains("営業") | Should -Be $false
        @(getTargetFolders $fixtureB.SettingsPath | ForEach-Object { "$($_.Name)|$($_.Path)" }) | Should -Be @("経理|C:\経理", "営業|C:\新しい場所", "総務|C:\総務")
    }

    It "空き容量: 調べられないとき（UNC・例外・`$null）は確かめを飛ばしてインポートできる" -TestCases @(
        @{ label = "`$null を返す"; getFreeSpace = { param ($root) $null } }
        @{ label = "例外になる"; getFreeSpace = { param ($root) throw "調べられない" } }
    ) {
        param ($label, $getFreeSpace)
        $fixtureA = newIndexFixture "$TestDrive\freeunknown_a_$([Guid]::NewGuid().ToString('N'))" "営業" "C:\共有\営業部"
        $dest = "$($fixtureA.Workspace.Dir)\freeunknown.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null
        $wsB = [Workspace]::new("$TestDrive\freeunknown_b_$([Guid]::NewGuid().ToString('N'))")

        $result = importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB "$($wsB.Dir)\setting.config" $getFreeSpace

        $result.Name | Should -Be "営業"
    }

    It "差分だけ: 元のフォルダが同じファイル（更新日時・サイズが同じ）は取り込み対象が 0 件になる" {
        # 取り込み一覧の行（更新日時・サイズ）と同じファイルを、元のフォルダに実際に置く
        $srcFolder = "$TestDrive\diffonly_src"
        [System.IO.Directory]::CreateDirectory((Join-Path $srcFolder "見積")) | Out-Null
        $filePath = Join-Path $srcFolder "見積\A社.xlsx"
        [System.IO.File]::WriteAllBytes($filePath, (New-Object byte[] 100))
        [System.IO.File]::SetLastWriteTime($filePath, [datetime]::Parse("2024/01/01 00:00:00"))

        $fixtureA = newIndexFixture "$TestDrive\diffonly_a" "営業" $srcFolder
        $dest = "$TestDrive\diffonly.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\diffonly_b")
        $settingsB = "$TestDrive\diffonly_b\setting.config"
        importIndex $dest ${importCollisionRename} "" $srcFolder $wsB $settingsB | Out-Null

        $status = readStatusFile $wsB.StatusFile
        $plan = createTargetList @{ Path = $srcFolder; Name = "営業" } $status.Rows $null
        $plan.Targets.Count | Should -Be 0
        $plan.Rows.Count | Should -Be 1
    }

    It "登録: targetFolders に名前・場所が入り、元のフォルダが無ければ Enabled が `$false。取り込み一覧・source_folder.txt もできる" {
        $fixtureA = newIndexFixture "$TestDrive\register_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\register.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\register_b")
        $settingsB = "$TestDrive\register_b\setting.config"
        $result = importIndex $dest ${importCollisionRename} "" "D:\無いフォルダ" $wsB $settingsB

        $result.Enabled | Should -Be $false
        $targets = @(getTargetFolders $settingsB)
        $targets.Count | Should -Be 1
        $targets[0].Name | Should -Be "営業"
        $targets[0].Path | Should -Be "D:\無いフォルダ"
        $targets[0].Enabled | Should -Be $false

        $status = readStatusFile $wsB.StatusFile
        @($status.Folders | Where-Object { $_.Name -eq "営業" }).Count | Should -Be 1
        $status.Rows.ContainsKey("営業\見積\A社.xlsx") | Should -Be $true
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業\source_folder.txt") | Should -Be $true
        Test-Path -LiteralPath (Join-Path $wsB.SystemIndexDir "営業") | Should -Be $false
    }

    It "前の版のワークスペース: content_index が空で前の版のしるしがあれば片付けてからインポートする" {
        $fixtureA = newIndexFixture "$TestDrive\legacy_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\legacy.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\legacy_b")
        $settingsB = "$TestDrive\legacy_b\setting.config"
        # 前の版のしるし（index\<名前>\元のフォルダ.txt）と、前の名前のシステムインデックス txt・状態ファイルのキーを作る
        New-Item -ItemType Directory -Path (Join-Path $wsB.LegacyIndexDir "旧") -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $wsB.LegacyIndexDir "旧\元のフォルダ.txt") -Value "旧`tC:\旧" -Encoding UTF8
        New-Item -ItemType Directory -Path (Join-Path $wsB.SystemIndexDir "旧") -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $wsB.SystemIndexDir "旧\システムインデックス.txt") -Value "x00000000" -Encoding UTF8
        [void](updateSystemIndexState { param ($s) [void]$s.Covered.Add("旧") } $wsB.SystemIndexStateFile)

        importIndex $dest ${importCollisionRename} "" "C:\共有\営業部" $wsB $settingsB | Out-Null

        Test-Path -LiteralPath $wsB.SystemIndexDir | Should -Be $false
        (readSystemIndexState $wsB.SystemIndexStateFile).Covered.Count | Should -Be 0
        Test-Path -LiteralPath (Join-Path $wsB.LegacyIndexDir "旧\元のフォルダ.txt") | Should -Be $true
    }

    It "前の版のワークスペース: 片付けられなければ例外にし、content_index にも設定にも足さない" {
        $fixtureA = newIndexFixture "$TestDrive\legacy_fail_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\legacy_fail.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\legacy_fail_b")
        $settingsB = "$TestDrive\legacy_fail_b\setting.config"
        New-Item -ItemType Directory -Path (Join-Path $wsB.LegacyIndexDir "旧") -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $wsB.LegacyIndexDir "旧\元のフォルダ.txt") -Value "旧`tC:\旧" -Encoding UTF8
        New-Item -ItemType Directory -Path (Join-Path $wsB.SystemIndexDir "旧") -Force | Out-Null
        $txt = Join-Path $wsB.SystemIndexDir "旧\システムインデックス.txt"
        Set-Content -LiteralPath $txt -Value "x00000000" -Encoding UTF8

        $stream = [System.IO.File]::Open($txt, "Open", "Read", "None")
        try {
            { importIndex $dest ${importCollisionRename} "" "C:\共有\営業部" $wsB $settingsB } | Should -Throw
        } finally {
            $stream.Dispose()
        }
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業") | Should -Be $false
        @(getTargetFolders $settingsB).Count | Should -Be 0
    }

    It "同じ名前: <mode>" -TestCases @(
        @{ mode = "Rename"; expectedName = "営業(2)" }
        @{ mode = "Overwrite"; expectedName = "営業" }
        @{ mode = "Cancel"; expectedName = $null }
    ) {
        param ($mode, $expectedName)
        $fixtureA = newIndexFixture "$TestDrive\collision_a_$mode" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\collision_$mode.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $fixtureB = newIndexFixture "$TestDrive\collision_b_$mode" "営業" "C:\前の場所"
        $wsB = $fixtureB.Workspace
        $settingsB = $fixtureB.SettingsPath
        writeSearchExcludes @([pscustomobject]@{ Path = (Join-Path $wsB.IndexDir "営業"); Subfolders = $true }) $settingsB

        $result = importIndex $dest $mode "" "C:\新しい場所" $wsB $settingsB

        if ($null -eq $expectedName) {
            $result | Should -Be $null
            @(getTargetFolders $settingsB)[0].Path | Should -Be "C:\前の場所"
            return
        }
        $result.Name | Should -Be $expectedName
        $targets = @(getTargetFolders $settingsB)
        if ($mode -eq ${importCollisionOverwrite}) {
            $targets.Count | Should -Be 1
            $targets[0].Path | Should -Be "C:\新しい場所"
            @(readSearchExcludes $settingsB).Count | Should -Be 0
        } else {
            $targets.Count | Should -Be 2
        }
    }

    It "同じ元のフォルダが検索だけのインデックス（別の名前）にあれば Warnings に入る（止めない）" {
        $fixtureA = newIndexFixture "$TestDrive\warn_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\warn.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\warn_b")
        $settingsB = "$TestDrive\warn_b\setting.config"
        writeIndexSources @([pscustomobject]@{ Name = "既存"; Path = "C:\共有\営業部" }) $settingsB

        $result = importIndex $dest ${importCollisionRename} "" "C:\共有\営業部" $wsB $settingsB

        @($result.Warnings).Count | Should -BeGreaterThan 0
        $result.Warnings[0] | Should -Match "既存"
    }

    It "止める: インデックス作成のロックを取れなければ例外にする" {
        $fixtureA = newIndexFixture "$TestDrive\import_locked_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\import_locked.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\import_locked_b")
        $settingsB = "$TestDrive\import_locked_b\setting.config"
        $lock = newAppMutex "indexer" $wsB.Dir
        try {
            { importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB $settingsB } | Should -Throw "*インデックス作成中*"
        } finally {
            $lock.Mutex.ReleaseMutex()
            $lock.Mutex.Dispose()
        }
    }

    It "壊れた・別の版の zip は例外になり、ワークスペースが変わらず作業フォルダが残らない" -TestCases @(
        @{ label = "zip でないファイル" }
        @{ label = "目録が JSON でない" }
    ) {
        param ($label)
        $wsB = [Workspace]::new("$TestDrive\import_bad_$([Guid]::NewGuid().ToString('N'))")
        $settingsB = "$($wsB.Dir)\setting.config"
        $path = "$($wsB.Dir)\bad.zip"
        [System.IO.Directory]::CreateDirectory($wsB.Dir) | Out-Null
        switch ($label) {
            "zip でないファイル" { Set-Content -LiteralPath $path -Value "これは zip ではありません" -Encoding UTF8 }
            "目録が JSON でない" {
                $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
                try {
                    $entry = $archive.CreateEntry("tebunko-index.json")
                    $writer = New-Object System.IO.StreamWriter($entry.Open())
                    $writer.Write("これは JSON ではありません")
                    $writer.Dispose()
                } finally {
                    $archive.Dispose()
                }
            }
        }

        { importIndex $path ${importCollisionRename} "" "C:\新しい場所" $wsB $settingsB } | Should -Throw
        Test-Path -LiteralPath $wsB.IndexDir | Should -Be $false
        Test-Path -LiteralPath $wsB.StatusFile | Should -Be $false
        Test-Path -LiteralPath $settingsB | Should -Be $false
        Test-Path -LiteralPath $wsB.PublishDir | Should -Be $false
    }

    It "zip slip: <label> は例外になり、ワークスペースの外にも中にもファイルができない" -TestCases @(
        @{ label = "エントリーが .."; entry = "../x.tsv" }
        @{ label = "エントリーが content_index の外へ抜ける"; entry = "content_index/../../x.tsv" }
        @{ label = "エントリーが / 始まり"; entry = "/x.tsv" }
        @{ label = "エントリーがドライブ文字"; entry = "C:/x.tsv" }
        @{ label = "content_index の直下に \ が入った名前"; entry = "content_index\x.tsv" }
        @{ label = "代替データストリーム"; entry = "content_index/a:b.tsv" }
        @{ label = "予約名のフォルダ"; entry = "content_index/CON/x.tsv" }
    ) {
        param ($label, $entry)
        $wsB = [Workspace]::new("$TestDrive\slip_$([Guid]::NewGuid().ToString('N'))")
        $settingsB = "$($wsB.Dir)\setting.config"
        [System.IO.Directory]::CreateDirectory($wsB.Dir) | Out-Null
        $path = "$($wsB.Dir)\bad.zip"

        $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create, [System.Text.Encoding]::UTF8)
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes("x")
            $sha = getBytesSha256 $bytes
            writeArchiveBytes $archive $entry $bytes
            $manifest = [ordered]@{
                format = ${indexArchiveFormat}; formatVersion = ${indexArchiveFormatVersion}; appVersion = ""
                exportedAt = "2026-09-27T10:00:00+09:00"; indexName = "営業"; sourceFolder = "C:\共有\営業部"
                files = @([ordered]@{ path = $entry; size = $bytes.Length; sha256 = $sha })
            }
            writeArchiveBytes $archive "tebunko-index.json" ((New-Object System.Text.UTF8Encoding($false)).GetBytes((ConvertTo-Json $manifest -Depth 5)))
        } finally {
            $archive.Dispose()
        }

        { importIndex $path ${importCollisionRename} "" "C:\新しい場所" $wsB $settingsB } | Should -Throw
        Test-Path -LiteralPath $wsB.Dir | Should -Be $true
        Test-Path -LiteralPath $wsB.IndexDir | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\x.tsv" | Should -Be $false
    }

    It "大きさの細工: 目録より大きく展開されるエントリーで例外になり、目録の大きさを超えて書かない" {
        $fixtureA = newIndexFixture "$TestDrive\bigfile_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\bigfile.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        # 本文インデックスのファイルの目録の大きさを、実際より小さい値に書き換える
        $archive = [System.IO.Compression.ZipFile]::Open($dest, [System.IO.Compression.ZipArchiveMode]::Update, [System.Text.Encoding]::UTF8)
        try {
            $entry = $archive.GetEntry("tebunko-index.json")
            $reader = New-Object System.IO.StreamReader($entry.Open())
            $manifest = ConvertFrom-Json $reader.ReadToEnd()
            $reader.Dispose()
            foreach ($file in $manifest.files) {
                if ($file.path -like "content_index/*") {
                    $file.size = 1
                }
            }
            $entry.Delete()
            $newEntry = $archive.CreateEntry("tebunko-index.json")
            $writer = New-Object System.IO.StreamWriter($newEntry.Open())
            $writer.Write((ConvertTo-Json $manifest -Depth 5))
            $writer.Dispose()
        } finally {
            $archive.Dispose()
        }

        $wsB = [Workspace]::new("$TestDrive\bigfile_b")
        $settingsB = "$TestDrive\bigfile_b\setting.config"
        { importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB $settingsB } | Should -Throw "*目録の大きさより大きく展開*"
        Test-Path -LiteralPath $wsB.IndexDir | Should -Be $false
    }

    It "大きさの細工: 空き容量が「合計 + 1GB」に足りなければ例外になる" {
        $fixtureA = newIndexFixture "$TestDrive\space_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\space.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\space_b")
        $settingsB = "$TestDrive\space_b\setting.config"

        { importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB $settingsB { param ($root) [long]0 } } | Should -Throw "*空き容量*"
        Test-Path -LiteralPath $wsB.IndexDir | Should -Be $false
    }

    It "戻す: 取り込み一覧を書けなければ、上書きの前の状態に戻る（フォルダ・取り込み一覧・設定の並びとほかの項目・検索から外した記録）" {
        $fixtureA = newIndexFixture "$TestDrive\revert_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\revert.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $fixtureB = newIndexFixture "$TestDrive\revert_b" "営業" "C:\前の場所"
        $wsB = $fixtureB.Workspace
        $settingsB = $fixtureB.SettingsPath
        writeTargetFolders @(
            [pscustomobject]@{ Name = "経理"; Path = "C:\経理"; Enabled = $false }
            [pscustomobject]@{ Name = "営業"; Path = "C:\前の場所"; Enabled = $true }
            [pscustomobject]@{ Name = "総務"; Path = "C:\総務"; Enabled = $true }
        ) $settingsB
        writeSearchOption @{ UseRegex = $true } $settingsB
        writeSearchExcludes @([pscustomobject]@{ Path = (Join-Path $wsB.IndexDir "営業\見積"); Subfolders = $true }) $settingsB
        $packBefore = [System.IO.File]::ReadAllBytes($fixtureB.PackPath)
        $statusBefore = [System.IO.File]::ReadAllText($wsB.StatusFile)

        # 取り込み一覧の書き込みだけを失敗させる（設定の登録と content_index の入れ替えまでは進む）
        Mock writeTextLinesAtomic { throw "取り込み一覧を書けない" } -ParameterFilter { $path -eq $wsB.StatusFile }
        { importIndex $dest ${importCollisionOverwrite} "" "C:\新しい場所" $wsB $settingsB } | Should -Throw "*取り込み一覧を書けない*"

        # content_index\営業 は前の中身のまま
        [System.IO.File]::ReadAllBytes($fixtureB.PackPath) | Should -Be $packBefore
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業\source_folder.txt") | Should -Be $false
        Test-Path -LiteralPath (Join-Path $wsB.PublishDir "import") | Should -Be $false
        # 取り込み一覧は変わらない
        [System.IO.File]::ReadAllText($wsB.StatusFile) | Should -Be $statusBefore
        # 設定の並びと、ほかの項目・検索から外した記録が元に戻る
        $targets = @(getTargetFolders $settingsB)
        @($targets | ForEach-Object { "$($_.Name)|$($_.Path)|$($_.Enabled)" }) | Should -Be @("経理|C:\経理|False", "営業|C:\前の場所|True", "総務|C:\総務|True")
        (readSearchOption $settingsB).UseRegex | Should -Be $true
        @(readSearchExcludes $settingsB).Count | Should -Be 1
    }

    It "戻す: 設定に書けなければ、何も変えずに終わり、作業フォルダも残らない" {
        $fixtureA = newIndexFixture "$TestDrive\revert_set_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\revert_set.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $fixtureB = newIndexFixture "$TestDrive\revert_set_b" "営業" "C:\前の場所"
        $wsB = $fixtureB.Workspace
        $settingsB = $fixtureB.SettingsPath
        $packBefore = [System.IO.File]::ReadAllBytes($fixtureB.PackPath)
        $statusBefore = [System.IO.File]::ReadAllText($wsB.StatusFile)

        Mock writeTargetFolders { throw "設定を書けない" }
        { importIndex $dest ${importCollisionOverwrite} "" "C:\新しい場所" $wsB $settingsB } | Should -Throw "*設定を書けない*"

        [System.IO.File]::ReadAllBytes($fixtureB.PackPath) | Should -Be $packBefore
        [System.IO.File]::ReadAllText($wsB.StatusFile) | Should -Be $statusBefore
        Test-Path -LiteralPath (Join-Path $wsB.PublishDir "import") | Should -Be $false
        @(getTargetFolders $settingsB)[0].Path | Should -Be "C:\前の場所"
    }

    It "戻す: content_index を入れ替えられなければ、前のフォルダを戻し、設定も元に戻る" {
        $fixtureA = newIndexFixture "$TestDrive\revert_swap_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\revert_swap.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $fixtureB = newIndexFixture "$TestDrive\revert_swap_b" "営業" "C:\前の場所"
        $wsB = $fixtureB.Workspace
        $settingsB = $fixtureB.SettingsPath
        $packBefore = [System.IO.File]::ReadAllBytes($fixtureB.PackPath)

        # 設定の登録の後の、content_index の入れ替えを失敗させる
        Mock swapInImportedIndexDir { throw "入れ替えられない" }
        { importIndex $dest ${importCollisionOverwrite} "" "C:\新しい場所" $wsB $settingsB } | Should -Throw "*入れ替えられない*"

        [System.IO.File]::ReadAllBytes($fixtureB.PackPath) | Should -Be $packBefore
        @(getTargetFolders $settingsB)[0].Path | Should -Be "C:\前の場所"
        Test-Path -LiteralPath (Join-Path $wsB.PublishDir "import") | Should -Be $false
    }

    It "インポート後にインデックス作成の名前の割り当てを通しても、<label>: どのインデックスも消えず、設定・取り込み一覧・content_index が一致する" -TestCases @(
        @{ label = "上書き"; mode = "Overwrite"; expected = @("経理", "営業", "総務") }
        @{ label = "別名（別のフォルダ）"; mode = "Rename"; expected = @("経理", "営業", "総務", "営業(2)") }
    ) {
        param ($label, $mode, $expected)
        $fixtureA = newIndexFixture "$TestDrive\assign_a_$mode" "営業" "C:\共有\営業部" "見積\A社.xlsx" "品名`tX`n新規`t1`n"
        $dest = "$TestDrive\assign_$mode.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $fixtureB = newIndexFixture "$TestDrive\assign_b_$mode" "営業" "C:\前の場所" "前の資料\旧.xlsx"
        $wsB = $fixtureB.Workspace
        $settingsB = $fixtureB.SettingsPath
        foreach ($other in @("経理", "総務")) {
            [System.IO.Directory]::CreateDirectory((Join-Path $wsB.IndexDir $other)) | Out-Null
        }
        writeTargetFolders @(
            [pscustomobject]@{ Name = "経理"; Path = "C:\経理"; Enabled = $true }
            [pscustomobject]@{ Name = "営業"; Path = "C:\前の場所"; Enabled = $false }
            [pscustomobject]@{ Name = "総務"; Path = "C:\総務"; Enabled = $true }
        ) $settingsB
        updateSettings "caseSensitive" $true $settingsB
        updateSettings "fileFilter" "*.xlsx" $settingsB
        $status = readStatusFile $wsB.StatusFile
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\経理"; Name = "経理" }
            [pscustomobject]@{ Path = "C:\前の場所"; Name = "営業" }
            [pscustomobject]@{ Path = "C:\総務"; Name = "総務" }
        ) (getStatusRowArray $status) $wsB.StatusFile

        $newFolder = if ($mode -eq "Overwrite") { "D:\新しい場所" } else { "D:\別の場所" }
        $result = importIndex $dest $mode "" $newFolder $wsB $settingsB

        # 設定: 並び・path・enabled・ほかの項目
        $targets = @(getTargetFolders $settingsB)
        @($targets | ForEach-Object { $_.Name }) | Should -Be $expected
        $imported = @($targets | Where-Object { $_.Name -eq $result.Name })
        $imported.Count | Should -Be 1
        $imported[0].Path | Should -Be $newFolder
        $imported[0].Enabled | Should -Be $false   # D: の元のフォルダは無い
        @($targets | Where-Object { $_.Name -eq "経理" })[0].Path | Should -Be "C:\経理"
        $settings = readSettings $settingsB
        $settings.caseSensitive | Should -Be $true
        $settings.fileFilter | Should -Be "*.xlsx"

        # インデックス作成の始めと同じ手順: 名前の割り当て・消えたフォルダの整理。今の設定にあるインデックスは消えない
        $folders = @(assignIndexNames $targets (readStatusFile $wsB.StatusFile).Folders)
        function writeIndexerLog { param ($m, $c) }
        removeDroppedFolders $folders (readStatusFile $wsB.StatusFile).Folders $wsB
        foreach ($name in $expected) {
            $folders.Name | Should -Contain $name
            # 名前の割り当てだけでなく、removeDroppedFolders が実際に何も消していないこと
            Test-Path -LiteralPath (Join-Path $wsB.IndexDir $name) -PathType Container | Should -Be $true
        }
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir $result.Name) | Should -Be $true
        # 取り込み一覧・content_index が設定の名前と一致する
        $statusAfter = readStatusFile $wsB.StatusFile
        @($statusAfter.Folders | ForEach-Object { $_.Name }) | Should -Contain $result.Name
        @($statusAfter.Rows.Keys | Where-Object { $_ -like "$($result.Name)\*" }).Count | Should -Be 1
        (readSearchWord $wsB "新規").Count | Should -BeGreaterThan 0
        if ($mode -eq "Overwrite") {
            (readSearchWord $wsB "鉛筆").Count | Should -Be 0
        } else {
            # 別名では、元の 営業 は場所も中身もそのまま残る
            @(getTargetFolders $settingsB | Where-Object { $_.Name -eq "営業" })[0].Path | Should -Be "C:\前の場所"
            @($statusAfter.Folders | Where-Object { $_.Name -eq "営業" })[0].Path | Should -Be "C:\前の場所"
            (readSearchWord $wsB "鉛筆").Count | Should -BeGreaterThan 0
        }
    }

    It "止める: 同じ元のフォルダが別の名前のクロール対象フォルダにあれば、別名でも例外にし、設定・content_index を変えない（getTargetFolders は同じフォルダの 2 つ目以降を読まず、登録するとインデックス作成で消えるため）" {
        $fixtureA = newIndexFixture "$TestDrive\samefolder_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\samefolder.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null
        $fixtureB = newIndexFixture "$TestDrive\samefolder_b" "営業" "C:\共有\営業部"
        $wsB = $fixtureB.Workspace

        { importIndex $dest ${importCollisionRename} "" "C:\共有\営業部" $wsB $fixtureB.SettingsPath } | Should -Throw "*[営業]*が既にあります*"

        @(getTargetFolders $fixtureB.SettingsPath | ForEach-Object { "$($_.Name)|$($_.Path)" }) | Should -Be @("営業|C:\共有\営業部")
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業(2)") | Should -Be $false
        Test-Path -LiteralPath (Join-Path $wsB.PublishDir "import") | Should -Be $false
    }

    It "検索だけのインデックスと同じ名前: 上書きでインポートすると indexSources から外れ、targetFolders に入る" {
        $fixtureA = newIndexFixture "$TestDrive\src_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\src.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\src_b")
        $settingsB = "$($wsB.Dir)\setting.config"
        writeIndexSources @(
            [pscustomobject]@{ Name = "営業"; Path = "C:\別のPCの営業部" }
            [pscustomobject]@{ Name = "経理"; Path = "C:\別のPCの経理部" }
        ) $settingsB

        $result = importIndex $dest ${importCollisionOverwrite} "" "D:\別の場所\営業部" $wsB $settingsB

        $result.Name | Should -Be "営業"
        @(readIndexSources $settingsB | ForEach-Object { "$($_.Name)|$($_.Path)" }) | Should -Be @("経理|C:\別のPCの経理部")
        @(getTargetFolders $settingsB | ForEach-Object { "$($_.Name)|$($_.Path)" }) | Should -Be @("営業|D:\別の場所\営業部")
    }

    It "戻す: 検索だけのインデックスと同じ名前でインポートして失敗すると、indexSources も元に戻る" {
        $fixtureA = newIndexFixture "$TestDrive\src_rev_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\src_rev.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\src_rev_b")
        $settingsB = "$($wsB.Dir)\setting.config"
        writeIndexSources @(
            [pscustomobject]@{ Name = "経理"; Path = "C:\別のPCの経理部" }
            [pscustomobject]@{ Name = "営業"; Path = "C:\別のPCの営業部" }
        ) $settingsB

        Mock writeTextLinesAtomic { throw "取り込み一覧を書けない" } -ParameterFilter { $path -eq $wsB.StatusFile }
        { importIndex $dest ${importCollisionOverwrite} "" "D:\別の場所\営業部" $wsB $settingsB } | Should -Throw "*取り込み一覧を書けない*"

        @(readIndexSources $settingsB | ForEach-Object { "$($_.Name)|$($_.Path)" }) | Should -Be @("経理|C:\別のPCの経理部", "営業|C:\別のPCの営業部")
        @(getTargetFolders $settingsB).Count | Should -Be 0
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業") | Should -Be $false
    }

    It "パス: 260 文字を超えるパスの本文インデックスで往復できる" {
        $long = "とても長いフォルダ名" * 15
        $fixtureA = newIndexFixture "$TestDrive\longpath_a" "営業" "C:\共有\営業部" "$long\A社.xlsx"
        $dest = "$TestDrive\longpath.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\longpath_b")
        $settingsB = "$TestDrive\longpath_b\setting.config"
        try {
            $result = importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB $settingsB

            $result.Files | Should -Be 2
            Test-Path -LiteralPath (toLongPath (Join-Path $wsB.IndexDir "営業\$long\content_index.xlsx.001.tsv")) | Should -Be $true
        } finally {
            # TestDrive の後片付けは 260 文字を超えるパスを消せない（一時フォルダが残る）ため、ここで消す
            removeDirectoryRetry $fixtureA.Workspace.Dir
            removeDirectoryRetry $wsB.Dir
        }
    }

    It "パス: [ ] を含む名前で往復できる" {
        $fixtureA = newIndexFixture "$TestDrive\bracket_a" "[確定]見積" "C:\共有\営業部" "見積\A社.xlsx"
        $dest = "$TestDrive\bracket.zip"
        exportIndex "[確定]見積" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\bracket_b")
        $settingsB = "$TestDrive\bracket_b\setting.config"
        $result = importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB $settingsB

        $result.Name | Should -Be "[確定]見積"
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "[確定]見積\見積\content_index.xlsx.001.tsv") | Should -Be $true
    }

    It "壊れた行: <label> でも、インポートは止まらず、検索も例外にならない" -TestCases @(
        @{ label = "メタ情報の行に = が無い"; broken = { param ($t) $t.Replace("シート=Sheet1", "シート") } }
        @{ label = "中身の行が途中で切れている（末尾が欠けた）"; broken = { param ($t) $t.Substring(0, $t.IndexOf("鉛筆") + 1) } }
        @{ label = "メタ情報の行だけで中身が無い"; broken = { param ($t) $t.Substring(0, $t.IndexOf("品名")) } }
    ) {
        param ($label, $broken)
        # インポートは本文インデックスのファイルの中身の形を確かめない（大きさと SHA-256 で、作ったときのままであることだけを見る）。
        # 壊れた行があっても、検索が例外で止まらないことをここで確かめる
        $fixtureA = newIndexFixture "$TestDrive\brokenpack_a_$([Guid]::NewGuid().ToString('N'))" "営業" "C:\共有\営業部"
        $original = [System.IO.File]::ReadAllText($fixtureA.PackPath)
        [System.IO.File]::WriteAllText($fixtureA.PackPath, (& $broken $original), (New-Object System.Text.UTF8Encoding($true)))

        $dest = "$($fixtureA.Workspace.Dir)\brokenpack.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\brokenpack_b_$([Guid]::NewGuid().ToString('N'))")
        $result = importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB "$($wsB.Dir)\setting.config"

        $result.Files | Should -Be 2
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業\見積\content_index.xlsx.001.tsv") | Should -Be $true
        { readSearchWord $wsB "鉛筆" } | Should -Not -Throw
    }
}

Describe "getWorkspaceFreeSpace" -Tag Io {
    It "ドライブのルートなら空き容量（バイト）を返す" {
        $root = [System.IO.Path]::GetPathRoot($TestDrive)

        getWorkspaceFreeSpace $root | Should -BeGreaterThan 0
    }

    It "UNC（ドライブとして扱えないパス）は `$null を返す（例外にしない）" -TestCases @(
        @{ root = "\\server\share\" }
        @{ root = "\\server\share" }
    ) {
        param ($root)
        getWorkspaceFreeSpace $root | Should -Be $null
    }
}

Describe "testImportFreeSpace" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "足りていれば空文字列"; free = 5GB; total = 100MB; expectEmpty = $true }
        @{ label = "「合計 + 1GB」に足りなければ理由を返す"; free = 1GB; total = 100MB; expectEmpty = $false }
        @{ label = "調べられなければ（`$null）確かめない"; free = $null; total = 100MB; expectEmpty = $true }
    ) {
        param ($label, $free, $total, $expectEmpty)
        $block = { param ($root) $free }.GetNewClosure()

        $reason = testImportFreeSpace ([long]$total) "C:\" $block

        if ($expectEmpty) { $reason | Should -Be "" } else { $reason | Should -BeLike "*空き容量*" }
    }
}
