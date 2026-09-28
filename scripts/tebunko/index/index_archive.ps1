# インデックスのエクスポート・インポート（状態層。zip の読み書き・ワークスペースの書き換え）。
# 画面にも利用者にも問い合わせない（画面なしで呼べる形にする）。安全性の確かめは index_archive_rules.ps1（判断層）で行う。

# System.IO.Compression.ZipFile・ZipArchive は既定では読み込まれていないアセンブリにあるため、ここで読み込む
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function getBytesSha256 {
    # バイト列の SHA-256（16進64文字）を返す
    param (
        [byte[]]$bytes
    )

    $sha = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    try {
        return [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace("-", "")
    } finally {
        $sha.Dispose()
    }
}

function writeArchiveBytes {
    # バイト列を zip の1エントリーとして書く
    param (
        $archive,
        [string]$entryPath,
        [byte[]]$bytes
    )

    $entry = $archive.CreateEntry($entryPath, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
    }
}

function writeArchiveFile {
    # ファイルを1つ、ストリームで zip の1エントリーに写す（メモリに全体を読まない）。
    # 写しながら SHA-256 を計算し、@{ Size; Sha256 } を返す
    param (
        $archive,
        [string]$entryPath,
        [string]$sourcePath
    )

    $entry = $archive.CreateEntry($entryPath, [System.IO.Compression.CompressionLevel]::Optimal)
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $source = New-Object System.IO.FileStream((toLongPath $sourcePath), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $sha = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    try {
        $size = $source.Length
        $entryStream = $entry.Open()
        try {
            $crypto = New-Object System.Security.Cryptography.CryptoStream($entryStream, $sha, [System.Security.Cryptography.CryptoStreamMode]::Write)
            try {
                $source.CopyTo($crypto)
                $crypto.FlushFinalBlock()
            } finally {
                $crypto.Dispose()
            }
        } finally {
            $entryStream.Dispose()
        }
        return @{ Size = $size; Sha256 = [BitConverter]::ToString($sha.Hash).Replace("-", "") }
    } finally {
        $source.Dispose()
        $sha.Dispose()
    }
}

function exportIndex {
    # 1 つのインデックスを 1 つの zip に書き出す: @{ Path; Files; Bytes }。画面にも利用者にも問い合わせない。
    # インデックス作成のロックを取れなければ例外（ほかで実行中）
    param (
        [string]$name,
        [string]$destPath,
        $ws = $workspace,
        [string]$settingsPath = ${settingsFile}
    )

    $lock = newAppMutex "indexer" $ws.Dir
    if (!$lock.Acquired) {
        $lock.Mutex.Dispose()
        throw "インデックス作成中はエクスポートできません。インデックス作成が終わってからやり直してください。"
    }
    try {
        return (exportIndexCore $name $destPath $ws $settingsPath)
    } finally {
        $lock.Mutex.ReleaseMutex()
        $lock.Mutex.Dispose()
    }
}

function exportIndexCore {
    param (
        [string]$name,
        [string]$destPath,
        $ws,
        [string]$settingsPath
    )

    $indexDir = Join-Path $ws.IndexDir $name
    if (!(Test-Path -LiteralPath (toLongPath $indexDir) -PathType Container)) {
        throw "インデックス [${name}] が見つかりません。"
    }
    $status = readStatusFile $ws.StatusFile
    if (@($status.Folders | Where-Object { [string]::Equals([string]$_.Name, $name, [System.StringComparison]::OrdinalIgnoreCase) }).Count -eq 0) {
        throw "取り込み一覧に、インデックス [${name}] の記録がありません。"
    }

    # 入れる前の TSV（本文インデックスのファイル以外の *.tsv）が残っていれば、途中で止まったインデックスとして止める
    $longIndexDir = toLongPath $indexDir
    $prefixLength = $indexDir.TrimEnd("\").Length + 1
    $packFiles = New-Object System.Collections.Generic.List[string]
    foreach ($file in [System.IO.Directory]::EnumerateFiles($longIndexDir, "*.tsv", [System.IO.SearchOption]::AllDirectories)) {
        $rel = (fromLongPath $file).Substring($prefixLength)
        $fileName = [System.IO.Path]::GetFileName($rel)
        if ($null -eq (readPackFileName $fileName)) {
            throw "インデックス作成を最後まで行ってからエクスポートしてください（取り込みの途中のファイルが残っています: ${rel}）。"
        }
        $packFiles.Add($rel)
    }

    # 取り込み一覧のうち、このインデックスの行だけを抜き出し、相対パスの先頭の "<名前>\" を外す
    $prefix = "$name\"
    $rows = New-Object System.Collections.Generic.List[string]
    $rows.Add((${statusColumns} -join "`t"))
    foreach ($entry in $status.Rows.GetEnumerator()) {
        if ($entry.Key.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $row = $entry.Value
            $stripped = [pscustomobject]@{
                相対パス = $entry.Key.Substring($prefix.Length); 更新日時 = $row.更新日時; サイズ = $row.サイズ; 状態 = $row.状態
                TSV数 = $row.TSV数; 取り込み日時 = $row.取り込み日時; エラー = $row.エラー; 抽出版 = $row.抽出版
            }
            $rows.Add((toStatusLine $stripped))
        }
    }
    $statusText = [string]::Join("`r`n", $rows) + "`r`n"

    $tmpPath = "${destPath}.tmp"
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($tmpPath)) | Out-Null
    if (Test-Path -LiteralPath $tmpPath) {
        Remove-Item -LiteralPath $tmpPath -Force
    }

    try {
        $manifestFiles = New-Object System.Collections.Generic.List[object]
        $totalBytes = [long]0
        $archive = [System.IO.Compression.ZipFile]::Open($tmpPath, [System.IO.Compression.ZipArchiveMode]::Create, [System.Text.Encoding]::UTF8)
        try {
            $statusBytes = ${utf8Bom}.GetBytes($statusText)
            writeArchiveBytes $archive ${indexArchiveStatusEntryName} $statusBytes
            $manifestFiles.Add([ordered]@{ path = ${indexArchiveStatusEntryName}; size = [long]$statusBytes.Length; sha256 = (getBytesSha256 $statusBytes) })
            $totalBytes += $statusBytes.Length

            foreach ($rel in $packFiles) {
                $entryPath = "content_index/" + $rel.Replace("\", "/")
                $written = writeArchiveFile $archive $entryPath (Join-Path $indexDir $rel)
                $manifestFiles.Add([ordered]@{ path = $entryPath; size = [long]$written.Size; sha256 = $written.Sha256 })
                $totalBytes += $written.Size
            }

            $sourceFolder = ""
            $ownFile = readSourceFolderFile $indexDir
            if ($ownFile.ContainsKey($name)) {
                $sourceFolder = $ownFile[$name]
            } else {
                $sourceMap = getSourceFolderMap $ws.IndexDir $ws.StatusFile $settingsPath
                if ($sourceMap.ContainsKey($name)) {
                    $sourceFolder = $sourceMap[$name]
                }
            }
            $version = readVersionFile (Join-Path ${rootDir} "VERSION.txt")
            $manifest = [ordered]@{
                format        = ${indexArchiveFormat}
                formatVersion = ${indexArchiveFormatVersion}
                appVersion    = $(if ($version) { $version.Tag } else { "" })
                exportedAt    = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
                indexName     = $name
                sourceFolder  = $sourceFolder
                files         = $manifestFiles.ToArray()
            }
            $manifestJson = ConvertTo-Json -InputObject $manifest -Depth 5
            writeArchiveBytes $archive ${indexArchiveManifestFileName} ((New-Object System.Text.UTF8Encoding($false)).GetBytes($manifestJson))
        } finally {
            $archive.Dispose()
        }
    } catch {
        if (Test-Path -LiteralPath $tmpPath) {
            Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
        }
        throw
    }
    try {
        if (Test-Path -LiteralPath $destPath) {
            Remove-Item -LiteralPath $destPath -Force
        }
        [System.IO.File]::Move((toLongPath $tmpPath), (toLongPath $destPath))
    } catch {
        if (Test-Path -LiteralPath $tmpPath) {
            Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
        }
        throw
    }
    return @{ Path = $destPath; Files = $manifestFiles.Count; Bytes = $totalBytes }
}

function openIndexArchive {
    # zip を開き、目録（testIndexArchiveManifest）を確かめる。確かめられなければ、開いたアーカイブを閉じて例外にする。
    # 返すもの: @{ Archive; Manifest }（呼んだ側が Archive.Dispose() を行う）
    param (
        [string]$zipPath
    )

    try {
        $archive = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Read, [System.Text.Encoding]::UTF8)
    } catch {
        throw "tebunko のインデックスの zip ではない、または壊れています（$($_.Exception.Message)）。"
    }
    try {
        $entryNames = @($archive.Entries | ForEach-Object { $_.FullName })
        $manifestEntry = $archive.Entries | Where-Object { $_.FullName -ceq ${indexArchiveManifestFileName} } | Select-Object -First 1
        $manifest = $null
        $manifestBytes = [long]0
        if ($null -ne $manifestEntry) {
            $manifestBytes = $manifestEntry.Length
            if ($manifestBytes -le ${indexArchiveMaxManifestBytes}) {
                $reader = New-Object System.IO.StreamReader($manifestEntry.Open(), [System.Text.Encoding]::UTF8)
                try {
                    $json = $reader.ReadToEnd()
                } finally {
                    $reader.Dispose()
                }
                try {
                    $manifest = ConvertFrom-Json $json
                } catch {
                    $manifest = $null
                }
            }
        }
        $reason = testIndexArchiveManifest $manifest $entryNames $manifestBytes
        if ($reason) {
            throw $reason
        }
        return @{ Archive = $archive; Manifest = $manifest }
    } catch {
        $archive.Dispose()
        throw
    }
}

function readIndexArchiveInfo {
    # 目録だけを読んで確かめる（インポートせずに、画面が既定の名前・元のフォルダ・合計の大きさを出すために使う）。
    # 返すもの: @{ IndexName; SourceFolder; Files; Bytes; FormatVersion; AppVersion }
    param (
        [string]$zipPath
    )

    $opened = openIndexArchive $zipPath
    try {
        $manifest = $opened.Manifest
        $bytes = [long]0
        foreach ($file in @($manifest.files)) {
            $bytes += [long]$file.size
        }
        return @{
            IndexName     = [string]$manifest.indexName
            SourceFolder  = (getManifestSourceFolder $manifest)
            Files         = @($manifest.files).Count
            Bytes         = $bytes
            FormatVersion = [int]$manifest.formatVersion
            AppVersion    = [string]$manifest.appVersion
        }
    } finally {
        $opened.Archive.Dispose()
    }
}

function getImportUsedIndexNames {
    # 今使われているインデックス名（設定の targetFolders・indexSources、content_index\ 直下のフォルダ、
    # 取り込み一覧のクロール対象フォルダの行）を集めて返す（大文字・小文字を区別しない集合）
    param (
        $ws,
        [string]$settingsPath
    )

    $names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($folder in @(getTargetFolders $settingsPath)) {
        if ($folder.Name) {
            [void]$names.Add($folder.Name)
        }
    }
    foreach ($source in @(readIndexSources $settingsPath)) {
        if ($source.Name) {
            [void]$names.Add($source.Name)
        }
    }
    if (Test-Path -LiteralPath (toLongPath $ws.IndexDir) -PathType Container) {
        foreach ($dir in [System.IO.Directory]::GetDirectories((toLongPath $ws.IndexDir))) {
            [void]$names.Add([System.IO.Path]::GetFileName($dir))
        }
    }
    foreach ($folder in (readStatusFile $ws.StatusFile).Folders.ToArray()) {
        if ($folder.Name) {
            [void]$names.Add($folder.Name)
        }
    }
    return , $names
}

function extractArchiveEntryToFile {
    # zip のエントリーを、目録の大きさ（expectedSize）を超えないようストリームで写す。
    # 書き終えたら大きさと SHA-256 を比べ、違えば例外にする（写したファイルは呼び出し側が消す）
    param (
        $entry,
        [long]$expectedSize,
        [string]$expectedSha256,
        [string]$destPath
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName((toLongPath $destPath))) | Out-Null
    $source = $entry.Open()
    $sha = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    $total = [long]0
    try {
        $dest = New-Object System.IO.FileStream((toLongPath $destPath), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $buffer = New-Object byte[] 1MB
            while ($true) {
                $read = $source.Read($buffer, 0, $buffer.Length)
                if ($read -le 0) {
                    break
                }
                $total += $read
                if ($total -gt $expectedSize) {
                    throw "エントリー「$($entry.FullName)」が目録の大きさより大きく展開されました。"
                }
                [void]$sha.TransformBlock($buffer, 0, $read, $buffer, 0)
                $dest.Write($buffer, 0, $read)
            }
        } finally {
            $dest.Dispose()
        }
    } finally {
        $source.Dispose()
    }
    [void]$sha.TransformFinalBlock([byte[]]@(), 0, 0)
    $actualHash = [BitConverter]::ToString($sha.Hash).Replace("-", "")
    $sha.Dispose()
    if ($total -ne $expectedSize -or $actualHash -ne $expectedSha256) {
        throw "エントリー「$($entry.FullName)」の大きさか内容が目録と違います。"
    }
}

function importIndex {
    # zip から 1 つのインデックスをインポートする: @{ Name; SourcePath; Enabled; Files; Bytes; Warnings }。
    # collisionMode（Rename・Overwrite・Cancel）が Cancel なら $null を返す（ワークスペースは変えない）。
    # name・sourceFolder を省くと目録の値を使う。インデックス作成のロックを取れなければ例外
    param (
        [string]$zipPath,
        [string]$collisionMode,
        [string]$name = "",
        [string]$sourceFolder = "",
        $ws = $workspace,
        [string]$settingsPath = ${settingsFile},
        [scriptblock]$getFreeSpace = { param ($root) ([System.IO.DriveInfo]$root).AvailableFreeSpace }
    )

    $lock = newAppMutex "indexer" $ws.Dir
    if (!$lock.Acquired) {
        $lock.Mutex.Dispose()
        throw "インデックス作成中はインポートできません。インデックス作成が終わってからやり直してください。"
    }
    try {
        return (importIndexCore $zipPath $collisionMode $name $sourceFolder $ws $settingsPath $getFreeSpace)
    } finally {
        $lock.Mutex.ReleaseMutex()
        $lock.Mutex.Dispose()
    }
}

function importIndexCore {
    param (
        [string]$zipPath,
        [string]$collisionMode,
        [string]$name,
        [string]$sourceFolder,
        $ws,
        [string]$settingsPath,
        [scriptblock]$getFreeSpace
    )

    # 前の版のワークスペース（index\ あり・content_index\ が空）を、取り込み直しを始めるときと同じ手順で片付ける。
    # content_index\ にこのインポートの分を足す前に行う（片付けに失敗したら、以降は行わない）
    $legacyState = getLegacyIndexState $ws.Dir
    if (testLegacyCleanupNeeded $legacyState) {
        $cleanup = clearLegacySystemIndex $ws.Dir
        if (!$cleanup.Ok) {
            throw "前の版のインデックス（高速検索用）を片付けられませんでした（$($cleanup.Reason)）。tebunko の画面やエクスプローラーで開いていれば閉じてから、もう一度インポートしてください。"
        }
    }

    $opened = openIndexArchive $zipPath
    $archive = $opened.Archive
    try {
        $manifest = $opened.Manifest
        $suggestedName = if ($name) { $name } else { [string]$manifest.indexName }
        $usedNames = @(getImportUsedIndexNames $ws $settingsPath)
        $finalName = getImportIndexName $suggestedName $usedNames $collisionMode
        if ($null -eq $finalName) {
            return $null
        }
        $nameReason = testIndexName $finalName @()
        if ($nameReason) {
            throw $nameReason
        }
        $overwrite = ($collisionMode -eq ${importCollisionOverwrite})

        $folder = if ($sourceFolder) { normalizeFolderPath $sourceFolder } else { getManifestSourceFolder $manifest }
        if ($folder -eq "") {
            throw "元のフォルダを指定してください。"
        }

        $totalBytes = [long]0
        $fileEntries = @($manifest.files)
        foreach ($file in $fileEntries) {
            $totalBytes += [long]$file.size
        }
        $freeSpace = & $getFreeSpace ([System.IO.Path]::GetPathRoot($ws.Dir))
        if ($freeSpace -lt ($totalBytes + ${indexArchiveFreeSpaceMargin})) {
            throw "保存先の空き容量が足りません（必要 約 $([Math]::Ceiling(($totalBytes + ${indexArchiveFreeSpaceMargin}) / 1MB)) MB）。"
        }

        # 同じ元のフォルダが、別の名前で既に登録されていないか（止めないが知らせる）
        $warnings = New-Object System.Collections.Generic.List[string]
        $otherFolders = @(@(getTargetFolders $settingsPath) + @(readIndexSources $settingsPath) | Where-Object { $_.Name -ine $finalName })
        foreach ($other in $otherFolders) {
            if (testSameFolder $other.Path $folder) {
                $warnings.Add("元のフォルダ「${folder}」は、インデックス [$($other.Name)] としても登録されています。")
                break
            }
        }

        # 7. 展開する（目録のパスからだけ作る。エントリーの名前からは作らない）
        $workDir = Join-Path $ws.PublishDir "import"
        $newDir = Join-Path $workDir "new"
        removeDirectoryRetry $workDir
        [System.IO.Directory]::CreateDirectory((toLongPath $newDir)) | Out-Null
        try {
            $statusTmpPath = Join-Path $workDir ${indexArchiveStatusEntryName}
            $statusLines = $null
            foreach ($file in $fileEntries) {
                $path = [string]$file.path
                $entry = $archive.GetEntry($path)
                if ($null -eq $entry) {
                    throw "目録にあるのに zip に無いエントリーがあります: ${path}"
                }
                $size = [long]$file.size
                $sha256 = [string]$file.sha256
                if ($path -ceq ${indexArchiveStatusEntryName}) {
                    extractArchiveEntryToFile $entry $size $sha256 $statusTmpPath
                    $bytes = [System.IO.File]::ReadAllBytes((toLongPath $statusTmpPath))
                    $text = ${utf8Bom}.GetString($bytes)
                    $statusLines = @($text -split "\r?\n" | Where-Object { $_ -ne "" })
                    $reason = testImportedStatusLines $statusLines
                    if ($reason) {
                        throw "取り込み一覧.tsv を読み込めません（${reason}）"
                    }
                } else {
                    $rel = $path.Substring("content_index/".Length).Replace("/", "\")
                    extractArchiveEntryToFile $entry $size $sha256 (Join-Path $newDir $rel)
                }
            }
            if ($null -eq $statusLines) {
                throw "目録に ${indexArchiveStatusEntryName} がありません。"
            }
            $importedLines = @($statusLines | Select-Object -Skip 1)

            $header = "# 検索結果から元のファイルを開くときに使う、インデックス名とクロール対象フォルダの対応です（インデックス作成のたびに作り直します）"
            writeListFile (Join-Path $newDir ${sourceFolderFileName}) @($header, "${finalName}`t${folder}")
        } catch {
            removeDirectoryRetry $workDir
            throw
        }

        # 8. 設定に登録する（逆の操作で戻せるよう、変える前の内容を覚えておく）
        $before = invokeSettingsLocked -path $settingsPath -action {
            $targets = @(getTargetFolders $settingsPath)
            $existingTarget = @($targets | Where-Object { $_.Name -ieq $finalName }) | Select-Object -First 1
            $sources = @(readIndexSources $settingsPath)
            $existingSource = @($sources | Where-Object { $_.Name -ieq $finalName }) | Select-Object -First 1
            $enabled = [System.IO.Directory]::Exists((toLongPath $folder))
            if ($existingTarget) {
                writeTargetFolders (@($targets | ForEach-Object {
                    if ($_.Name -ieq $finalName) { [pscustomobject]@{ Name = $finalName; Path = $folder; Enabled = $enabled } } else { $_ }
                })) $settingsPath
            } else {
                writeTargetFolders (@($targets) + @([pscustomobject]@{ Name = $finalName; Path = $folder; Enabled = $enabled })) $settingsPath
            }
            if ($existingSource) {
                writeIndexSources (@($sources | Where-Object { $_.Name -ine $finalName })) $settingsPath
            }
            return @{ HadTarget = ($null -ne $existingTarget); Target = $existingTarget; Source = $existingSource }
        }
        if ($overwrite) {
            [void](removeSearchExcludesUnder (Join-Path $ws.IndexDir $finalName) $settingsPath)
        }

        try {
            # 9. content_index\<名前> を入れ替える
            $targetIndexDir = Join-Path $ws.IndexDir $finalName
            $previousDir = Join-Path $workDir "previous"
            $hadPrevious = Test-Path -LiteralPath (toLongPath $targetIndexDir) -PathType Container
            if ($hadPrevious) {
                [System.IO.Directory]::Move((toLongPath $targetIndexDir), (toLongPath $previousDir))
            }
            try {
                [System.IO.Directory]::CreateDirectory((toLongPath $ws.IndexDir)) | Out-Null
                [System.IO.Directory]::Move((toLongPath $newDir), (toLongPath $targetIndexDir))
            } catch {
                if ($hadPrevious) {
                    [System.IO.Directory]::Move((toLongPath $previousDir), (toLongPath $targetIndexDir))
                }
                throw
            }

            try {
                # 10. 取り込み一覧を書き直す（このインデックスの前の行を消し、新しい行を足す）
                $existingLines = @(readStatusLines $ws.StatusFile)
                $kept = New-Object System.Collections.Generic.List[string]
                foreach ($line in $existingLines) {
                    $fields = $line.Split("`t")
                    if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3 -and [string]::Equals($fields[2], $finalName, [System.StringComparison]::OrdinalIgnoreCase)) {
                        continue
                    }
                    if ($fields.Count -eq ${statusColumns}.Count -and $fields[0] -ne "") {
                        $split = splitIndexRelPath $fields[0]
                        if ($split.Rest -ne "" -and [string]::Equals($split.Name, $finalName, [System.StringComparison]::OrdinalIgnoreCase)) {
                            continue
                        }
                    }
                    $kept.Add($line)
                }
                $kept.Add("${statusFolderKey}`t${folder}`t${finalName}")
                foreach ($line in $importedLines) {
                    $fields = $line.Split("`t")
                    $fields[0] = "${finalName}\$($fields[0])"
                    $kept.Add(($fields -join "`t"))
                }
                writeTextLinesAtomic $ws.StatusFile $kept
            } catch {
                # 9 を戻す
                removeDirectoryRetry $targetIndexDir
                if ($hadPrevious) {
                    [System.IO.Directory]::Move((toLongPath $previousDir), (toLongPath $targetIndexDir))
                }
                throw
            }
        } catch {
            # 8 を戻す
            invokeSettingsLocked -path $settingsPath -action {
                $targets = @(getTargetFolders $settingsPath)
                if ($before.HadTarget) {
                    writeTargetFolders (@($targets | ForEach-Object {
                        if ($_.Name -ieq $finalName) { $before.Target } else { $_ }
                    })) $settingsPath
                } else {
                    writeTargetFolders (@($targets | Where-Object { $_.Name -ine $finalName })) $settingsPath
                }
                if ($null -ne $before.Source) {
                    writeIndexSources ((@(readIndexSources $settingsPath)) + @($before.Source)) $settingsPath
                }
            } | Out-Null
            removeDirectoryRetry $workDir
            throw
        }

        # 11. 上書きなら、前のシステムインデックスを消す（失敗しても続ける。次のインデックス作成で整理される）
        if ($overwrite) {
            try {
                [void](removeSystemIndexOf $finalName $ws.SystemIndexDir $ws.SystemIndexStateFile)
            } catch {
            }
        }
        # 12. 作業フォルダを消す
        removeDirectoryRetry $workDir

        return @{
            Name = $finalName; SourcePath = $folder; Enabled = [System.IO.Directory]::Exists((toLongPath $folder))
            Files = $fileEntries.Count; Bytes = $totalBytes; Warnings = $warnings.ToArray()
        }
    } finally {
        $archive.Dispose()
    }
}
