# インデックスのエクスポート・インポート（tebunko\index\index_archive.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"

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

        $result.Path | Should -Be $dest
        $result.Files | Should -Be 2
        Test-Path -LiteralPath $dest | Should -Be $true
        Test-Path -LiteralPath "${dest}.tmp" | Should -Be $false

        $archive = [System.IO.Compression.ZipFile]::Open($dest, [System.IO.Compression.ZipArchiveMode]::Read, [System.Text.Encoding]::UTF8)
        try {
            @($archive.Entries | ForEach-Object { $_.FullName }) | Sort-Object | Should -Be @(
                "content_index/見積/content_index.xlsx.001.tsv", "tebunko-index.json", "取り込み一覧.tsv"
            )
            $manifestEntry = $archive.GetEntry("tebunko-index.json")
            $reader = New-Object System.IO.StreamReader($manifestEntry.Open())
            $manifest = ConvertFrom-Json $reader.ReadToEnd()
            $reader.Dispose()
            $manifest.indexName | Should -Be "営業"
            $manifest.sourceFolder | Should -Be "C:\共有\営業部"
            $manifest.formatVersion | Should -Be 1

            $statusEntry = $archive.GetEntry("取り込み一覧.tsv")
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

    It "登録: targetFolders に名前・場所が入り、元のフォルダが無ければ Enabled が `$false。取り込み一覧・元のフォルダ.txt もできる" {
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
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業\元のフォルダ.txt") | Should -Be $true
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

    It "同じ元のフォルダが別の名前で既に登録されていれば Warnings に入る（止めない）" {
        $fixtureA = newIndexFixture "$TestDrive\warn_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\warn.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\warn_b")
        $settingsB = "$TestDrive\warn_b\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "既存"; Path = "C:\共有\営業部"; Enabled = $true }) $settingsB

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

    It "戻す: 取り込み一覧を書けなければ、上書きの前の状態に戻り、設定のほかの項目は消えない" {
        $fixtureA = newIndexFixture "$TestDrive\revert_a" "営業" "C:\共有\営業部"
        $dest = "$TestDrive\revert.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $fixtureB = newIndexFixture "$TestDrive\revert_b" "営業" "C:\前の場所"
        $wsB = $fixtureB.Workspace
        $settingsB = $fixtureB.SettingsPath
        writeSearchOption @{ UseRegex = $true } $settingsB

        # 取り込み一覧を開いたままにして writeTextLinesAtomic を失敗させる
        $stream = [System.IO.File]::Open($wsB.StatusFile, "Open", "Read", "None")
        try {
            { importIndex $dest ${importCollisionOverwrite} "" "C:\新しい場所" $wsB $settingsB } | Should -Throw
        } finally {
            $stream.Dispose()
        }

        $targets = @(getTargetFolders $settingsB)
        $targets.Count | Should -Be 1
        $targets[0].Path | Should -Be "C:\前の場所"
        (readSearchOption $settingsB).UseRegex | Should -Be $true
        Test-Path -LiteralPath (Join-Path $wsB.IndexDir "営業\元のフォルダ.txt") | Should -Be $false
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

    It "壊れた行: 本文インデックスの中身（版の行）が壊れていても、大きさと SHA-256 が合っていればインポート自体は止まらない" {
        # インポートは本文インデックスのファイルの中身の形（pack_format.ps1 の版など）を確かめない
        # （大きさと SHA-256 で、作ったときのままであることだけを見る）。検索側での扱いはこの PR の範囲外
        $fixtureA = newIndexFixture "$TestDrive\brokenpack_a" "営業" "C:\共有\営業部"
        $original = [System.IO.File]::ReadAllText($fixtureA.PackPath)
        $broken = $original -replace "版=1", "版=9"
        [System.IO.File]::WriteAllText($fixtureA.PackPath, $broken, (New-Object System.Text.UTF8Encoding($true)))

        $dest = "$TestDrive\brokenpack.zip"
        exportIndex "営業" $dest $fixtureA.Workspace $fixtureA.SettingsPath | Out-Null

        $wsB = [Workspace]::new("$TestDrive\brokenpack_b")
        $settingsB = "$TestDrive\brokenpack_b\setting.config"
        $result = importIndex $dest ${importCollisionRename} "" "C:\新しい場所" $wsB $settingsB

        $result.Files | Should -Be 2
        $importedPath = Join-Path $wsB.IndexDir "営業\見積\content_index.xlsx.001.tsv"
        Test-Path -LiteralPath $importedPath | Should -Be $true
        [System.IO.File]::ReadAllText($importedPath) | Should -Match "版=9"
    }
}
