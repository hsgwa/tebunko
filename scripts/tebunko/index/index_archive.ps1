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
    # 1 つのインデックスを 1 つの zip に書き出す: @{ Name; Path; Files; Bytes }。画面にも利用者にも問い合わせない。
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
        throw "更新中はエクスポートできません。更新が終わってからやり直してください。"
    }
    try {
        return (exportIndexCore $name $destPath $ws $settingsPath)
    } finally {
        $lock.Mutex.ReleaseMutex()
        $lock.Mutex.Dispose()
    }
}

function exportIndexToFolder {
    # 1 つのインデックスを、書き出し先のフォルダの中の zip に書き出す（ファイル名は getExportFileName で決め、
    # フォルダに同じ名前があれば番号を付ける）。書き出し先のフォルダが無ければ例外。戻り値は exportIndex と同じ。
    # 届かないネットワークのフォルダで止まりうるので、画面のスレッドでは呼ばず、別スレッドの仕事の中で呼ぶ
    param (
        [string]$name,
        [string]$folder,
        $ws = $workspace,
        [string]$settingsPath = ${settingsFile}
    )

    if (!(Test-Path -LiteralPath (toLongPath $folder) -PathType Container)) {
        throw "書き出し先のフォルダが見つかりません：$folder"
    }
    $used = @(Get-ChildItem -LiteralPath (toLongPath $folder) -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
    $destPath = Join-Path $folder (getExportFileName $name $used)
    return (exportIndex $name $destPath $ws $settingsPath)
}

function exportIndexes {
    # 選んだインデックスを、書き出し先のフォルダの中の zip にまとめて書き出す（インデックス 1 つにつき zip 1 つ）。
    # 1 つずつ exportIndexToFolder を呼び、途中で 1 つ失敗しても残りを続ける。
    # 戻り値は名前ごとの結果の配列（@{ Name; Ok; Reason; Path }）。Reason は失敗したときの理由（成功なら ""）、Path は書き出した zip（失敗なら ""）。
    # 画面のスレッドでは呼ばず、別スレッドの仕事の中で呼ぶ
    param (
        [string[]]$names,
        [string]$destination,
        $ws = $workspace,
        [string]$settingsPath = ${settingsFile}
    )

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($name in @($names)) {
        try {
            $exported = exportIndexToFolder $name $destination $ws $settingsPath
            $results.Add([pscustomobject]@{ Name = $name; Ok = $true; Reason = ""; Path = $exported.Path })
        } catch {
            $results.Add([pscustomobject]@{ Name = $name; Ok = $false; Reason = $_.Exception.Message; Path = "" })
        }
    }
    return , $results.ToArray()
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
    # 途中で throw して抜けても下のフォルダを掴んだまま残らないよう、列挙子（EnumerateFiles）でなく配列（GetFiles）で受ける
    foreach ($file in [System.IO.Directory]::GetFiles($longIndexDir, "*.tsv", [System.IO.SearchOption]::AllDirectories)) {
        $rel = (fromLongPath $file).Substring($prefixLength)
        $fileName = [System.IO.Path]::GetFileName($rel)
        if ($null -eq (readPackFileName $fileName)) {
            throw "更新を最後まで行ってからエクスポートしてください（更新の途中のファイルが残っています: ${rel}）。"
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
    return @{ Name = $name; Path = $destPath; Files = $manifestFiles.Count; Bytes = $totalBytes }
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
    # 返すもの: @{ IndexName; SourceFolder; Files; Bytes; FormatVersion; AppVersion; ExportedAt（目録の exportedAt のまま。無ければ空） }
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
            ExportedAt    = [string]$manifest.exportedAt
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
        [void]$sha.TransformFinalBlock([byte[]]@(), 0, 0)
        $actualHash = [BitConverter]::ToString($sha.Hash).Replace("-", "")
    } finally {
        $source.Dispose()
        $sha.Dispose()
    }
    if ($total -ne $expectedSize -or $actualHash -ne $expectedSha256) {
        throw "エントリー「$($entry.FullName)」の大きさか内容が目録と違います。"
    }
}

function getWorkspaceFreeSpace {
    # root（ドライブのルート）の空き容量（バイト）を返す。UNC（\\server\share）など、ドライブとして扱えず調べられないときは $null
    # （調べられないだけでインポートを止めない。呼び出し側が確かめを飛ばす）
    param (
        [string]$root
    )

    try {
        return ([System.IO.DriveInfo]$root).AvailableFreeSpace
    } catch {
        return $null
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
        [scriptblock]$getFreeSpace = { param ($root) getWorkspaceFreeSpace $root }
    )

    $lock = newAppMutex "indexer" $ws.Dir
    if (!$lock.Acquired) {
        $lock.Mutex.Dispose()
        throw "更新中はインポートできません。更新が終わってからやり直してください。"
    }
    try {
        return (importIndexCore $zipPath $collisionMode $name $sourceFolder $ws $settingsPath $getFreeSpace)
    } finally {
        $lock.Mutex.ReleaseMutex()
        $lock.Mutex.Dispose()
    }
}

function testImportFreeSpace {
    # 空き容量が「目録の合計 + 余裕」に足りなければ、理由を返す（足りていれば空文字列）。
    # 空き容量を調べられないとき（getFreeSpace が $null を返す・例外になる。UNC のワークスペースなど）は確かめない
    param (
        [long]$totalBytes,
        [string]$root,
        [scriptblock]$getFreeSpace
    )

    $freeSpace = $null
    try {
        $freeSpace = & $getFreeSpace $root
    } catch {
    }
    if ($null -eq $freeSpace) {
        return ""
    }
    $needed = $totalBytes + ${indexArchiveFreeSpaceMargin}
    if ($freeSpace -lt $needed) {
        return "保存先の空き容量が足りません（必要 約 $([Math]::Ceiling($needed / 1MB)) MB）。"
    }
    return ""
}

function expandImportArchive {
    # 7. 目録にあるエントリーを作業フォルダ（<workDir>\new）に展開する（目録のパスからだけ作る。エントリーの名前からは作らない）。
    # 書き込む前に、大きさと SHA-256 を目録と比べる。source_folder.txt もここで作る。
    # 返すもの: @{ NewDir; Rows }（Rows は取り込み一覧に足す行。相対パスの先頭にインデックス名を付けたもの）。
    # 途中で失敗したら、作業フォルダを消して例外にする
    param (
        $archive,
        $fileEntries,
        [string]$workDir,
        [string]$finalName,
        [string]$folder
    )

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
                    throw "ingest_status.tsv を読み込めません（${reason}）"
                }
            } else {
                $rel = $path.Substring("content_index/".Length).Replace("/", "\")
                extractArchiveEntryToFile $entry $size $sha256 (Join-Path $newDir $rel)
            }
        }
        if ($null -eq $statusLines) {
            throw "目録に ${indexArchiveStatusEntryName} がありません。"
        }

        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($line in @($statusLines | Select-Object -Skip 1)) {
            $fields = $line.Split("`t")
            $rows.Add((newStatusRow "${finalName}\$($fields[0])" $fields[1] $fields[2] $fields[3] $fields[4] $fields[5] $fields[6] $fields[7]))
        }

        $header = "# 検索結果から元のファイルを開くときに使う、インデックス名とクロール対象フォルダの対応です（インデックス作成のたびに作り直します）"
        writeListFile (Join-Path $newDir ${sourceFolderFileName}) @($header, "${finalName}`t${folder}")
        return @{ NewDir = $newDir; Rows = $rows.ToArray() }
    } catch {
        removeDirectoryRetry $workDir
        throw
    }
}

function restoreImportedSettings {
    # registerImportedIndexInSettings で変える前の内容（before）に、設定の 3 つの項目（クロール対象・検索だけのインデックス・検索から外したフォルダ）を戻す。
    # ほかの項目には触らない
    param (
        $before,
        [string]$settingsPath
    )

    invokeSettingsLocked -path $settingsPath -action {
        writeTargetFolders $before.Targets $settingsPath
        if ($before.SourcesChanged) {
            writeIndexSources $before.Sources $settingsPath
        }
        if ($before.ExcludesChanged) {
            writeSearchExcludes $before.Excludes $settingsPath
        }
    } | Out-Null
}

function registerImportedIndexInSettings {
    # 8. 設定にインデックスを登録する（クロール対象に名前と元のフォルダを入れる。同じ名前があれば置き換え、同じ名前の検索だけのインデックスは外す。
    # 上書きなら、前のインデックスの検索から外したフォルダの記録も消す）。
    # 途中で失敗したら、変える前に戻して例外にする。成功したら、戻すための @{ Targets; Sources; Excludes; SourcesChanged; ExcludesChanged } を返す
    param (
        [string]$finalName,
        [string]$folder,
        [bool]$overwrite,
        [string]$indexDir,
        [string]$settingsPath
    )

    $before = invokeSettingsLocked -path $settingsPath -action {
        return @{
            Targets = @(getTargetFolders $settingsPath); Sources = @(readIndexSources $settingsPath)
            Excludes = @(readSearchExcludes $settingsPath); SourcesChanged = $false; ExcludesChanged = $false
        }
    }
    try {
        invokeSettingsLocked -path $settingsPath -action {
            $enabled = [System.IO.Directory]::Exists((toLongPath $folder))
            $entry = [pscustomobject]@{ Name = $finalName; Path = $folder; Enabled = $enabled }
            $targets = New-Object System.Collections.Generic.List[object]
            $found = $false
            foreach ($target in $before.Targets) {
                if ($target.Name -ieq $finalName) {
                    $targets.Add($entry)
                    $found = $true
                } else {
                    $targets.Add($target)
                }
            }
            if (!$found) {
                $targets.Add($entry)
            }
            writeTargetFolders $targets.ToArray() $settingsPath
            if (@($before.Sources | Where-Object { $_.Name -ieq $finalName }).Count -gt 0) {
                $before.SourcesChanged = $true
                writeIndexSources @($before.Sources | Where-Object { $_.Name -ine $finalName }) $settingsPath
            }
            if ($overwrite) {
                $before.ExcludesChanged = $true
                [void](removeSearchExcludesUnder $indexDir $settingsPath)
            }
        } | Out-Null
    } catch {
        restoreImportedSettings $before $settingsPath
        throw
    }
    return $before
}

function swapInImportedIndexDir {
    # 9. content_index\<名前> を展開したフォルダ（newDir）に入れ替える。前のフォルダは previous に退避する。
    # 失敗したら、退避したフォルダを戻して例外にする。成功したら、戻すための @{ TargetDir; PreviousDir; HadPrevious } を返す
    param (
        [string]$newDir,
        [string]$targetDir,
        [string]$previousDir,
        [string]$indexRoot
    )

    # ウイルス対策ソフトが、作ったばかり・書いたばかりのフォルダ（$targetDir は前の取り込みで、$newDir はこの
    # インポートの展開で、それぞれ書いたばかり）を一時的に掴んでいることがあるため、moveDirectoryRetry で試し直す
    # （removeDirectoryRetry と同じ理由。docs/design/index-data/format.md「インポート」を参照）
    $hadPrevious = Test-Path -LiteralPath (toLongPath $targetDir) -PathType Container
    if ($hadPrevious) {
        moveDirectoryRetry $targetDir $previousDir
    }
    try {
        [System.IO.Directory]::CreateDirectory((toLongPath $indexRoot)) | Out-Null
        moveDirectoryRetry $newDir $targetDir
    } catch {
        if ($hadPrevious) {
            moveDirectoryRetry $previousDir $targetDir
        }
        throw
    }
    return @{ TargetDir = $targetDir; PreviousDir = $previousDir; HadPrevious = $hadPrevious }
}

function restoreSwappedIndexDir {
    # swapInImportedIndexDir で入れ替えたフォルダを、入れ替える前に戻す（入れたものを消し、退避した前のフォルダを戻す）
    param (
        $swap
    )

    removeDirectoryRetry $swap.TargetDir
    if ($swap.HadPrevious) {
        moveDirectoryRetry $swap.PreviousDir $swap.TargetDir
    }
}

function rewriteStatusForImport {
    # 10. 取り込み一覧を書き直す（このインデックスの前のクロール対象フォルダの行と各行を消し、新しい行を足す）。
    # ほかのインデックスの行は残す。見出しの行を含め、インデクサが書くのと同じ形（writeStatusFile）で書く
    param (
        [string]$statusPath,
        [string]$finalName,
        [string]$folder,
        $importedRows
    )

    $status = readStatusFile $statusPath
    $folders = New-Object System.Collections.Generic.List[object]
    foreach ($other in $status.Folders) {
        if (![string]::Equals($other.Name, $finalName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $folders.Add($other)
        }
    }
    $folders.Add([pscustomobject]@{ Path = $folder; Name = $finalName })

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $status.Rows.Values) {
        if (![string]::Equals((splitIndexRelPath $row.相対パス).Name, $finalName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $rows.Add($row)
        }
    }
    foreach ($row in $importedRows) {
        $rows.Add($row)
    }
    writeStatusFile $folders.ToArray() $rows.ToArray() $statusPath
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
        # 集合を返す関数なので、@() で包まない（包むと集合が 1 要素の配列になり、newIndexName が名前を突き合わせられず、
        # 「別名」でも同じ名前のインデックスを上書きしてしまう）
        $usedNames = getImportUsedIndexNames $ws $settingsPath
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
        $spaceReason = testImportFreeSpace $totalBytes ([System.IO.Path]::GetPathRoot($ws.Dir)) $getFreeSpace
        if ($spaceReason) {
            throw $spaceReason
        }

        # 同じ元のフォルダが、別の名前のクロール対象フォルダに既にある・入れ子になっていれば止める（getTargetFolders は同じフォルダの 2 つ目以降を読まないため、
        # 登録しても設定に残らず、次のインデックス作成で removeDroppedFolders がこのインデックスを消す。画面の追加・編集と同じ getIndexFolderConflict の決まり）。
        # 検索だけのインデックス（indexSources）の元のフォルダと同じなら、止めずに知らせる
        $conflict = getIndexFolderConflict $folder @(getTargetFolders $settingsPath | Where-Object { $_.Name -ine $finalName })
        if ($conflict -ne "") {
            throw $conflict
        }
        $warnings = New-Object System.Collections.Generic.List[string]
        foreach ($other in @(readIndexSources $settingsPath | Where-Object { $_.Name -ine $finalName })) {
            if (testSameFolder $other.Path $folder) {
                $warnings.Add("元のフォルダ「${folder}」は、インデックス [$($other.Name)] としても登録されています。")
                break
            }
        }

        $workDir = Join-Path $ws.PublishDir "import"
        $expanded = expandImportArchive $archive $fileEntries $workDir $finalName $folder
        $targetIndexDir = Join-Path $ws.IndexDir $finalName
        try {
            $before = registerImportedIndexInSettings $finalName $folder $overwrite $targetIndexDir $settingsPath
            try {
                $swap = swapInImportedIndexDir $expanded.NewDir $targetIndexDir (Join-Path $workDir "previous") $ws.IndexDir
                # Directory.Move で入れたフォルダは、work\content_index の NotContentIndexed を受け継がない。
                # 根に付いたあとの -Recurse は中へ降りないため（publish.md）、入れた直後に付ける（失敗しても止めない）
                setNotContentIndexed $targetIndexDir -Recurse | Out-Null
                try {
                    rewriteStatusForImport $ws.StatusFile $finalName $folder $expanded.Rows
                } catch {
                    restoreSwappedIndexDir $swap
                    throw
                }
            } catch {
                restoreImportedSettings $before $settingsPath
                throw
            }
        } catch {
            removeDirectoryRetry $workDir
            throw
        }

        # 11. 上書きなら、前のシステムインデックスを消す（消せなくても続ける。次のインデックス作成で整理される）
        if ($overwrite) {
            removeSystemIndexOfWorkspace $finalName $ws.IndexDir
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
