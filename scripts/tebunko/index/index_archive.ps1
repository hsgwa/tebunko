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
    # 1 つのインデックスを 1 つの zip に書き出す: @{ Name; Path; Files; Bytes; Omitted }。Omitted は大きすぎて入れなかった本文インデックスのファイル（writeContentIndexFilesToArchive）。画面にも利用者にも問い合わせない。
    # インデックス作成のロックを取れなければ例外（ほかで実行中）
    param (
        [string]$name,
        [string]$destPath,
        $ws = $workspace,
        [string]$settingsPath = ${settingsFile}
    )

    $lock = enterIndexerLock $ws.Dir
    if (!$lock.Acquired) {
        throw "更新中はエクスポートできません。更新が終わってからやり直してください。"
    }
    try {
        return (exportIndexCore $name $destPath $ws $settingsPath)
    } finally {
        exitIndexerLock $lock
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
    # 戻り値は名前ごとの結果の配列（@{ Name; Ok; Reason; Path; Omitted }）。Omitted は exportIndex と同じ（失敗なら空）。Reason は失敗したときの理由（成功なら ""）、Path は書き出した zip（失敗なら ""）。
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
            $results.Add([pscustomobject]@{ Name = $name; Ok = $true; Reason = ""; Path = $exported.Path; Omitted = @($exported.Omitted) })
        } catch {
            $results.Add([pscustomobject]@{ Name = $name; Ok = $false; Reason = $_.Exception.Message; Path = ""; Omitted = @() })
        }
    }
    return , $results.ToArray()
}

function getIndexTsvFilesNoFollow {
    # インデックスのフォルダの下の *.tsv（FileInfo）を返す。下のリンク（ジャンクション・シンボリックリンク）は辿らない
    # （先の、インデックスの外のファイルを zip に入れて持ち出さないため。親を指すリンクで輪になることも無い）。
    # 読めなかったフォルダがあれば、一部だけをエクスポートしないよう例外にする
    param (
        [string]$indexDir
    )

    $tree = getFilesNoFollow $indexDir -IncludeHidden
    if ($tree.HasError) {
        throw "インデックスのフォルダを最後まで読めませんでした: $($tree.Reason)"
    }
    return @($tree.Files | Where-Object { [string]::Equals($_.Extension, ".tsv", [System.StringComparison]::OrdinalIgnoreCase) })
}

function getIndexExportSize {
    # エクスポートする前に、選んだインデックスが zip に入れる本文インデックスのファイル（*.tsv）の大きさの合計を数える。
    # 戻り値: @{ Bytes（合計のバイト数）; Missing（インデックスのフォルダが無かった名前の配列）}。
    # 数えるのは zip に入る元のファイルの大きさで、zip にした後の大きさではない。
    # 画面のスレッドでは呼ばず、別スレッドの仕事の中で呼ぶ
    param (
        [string[]]$names,
        $ws = $workspace
    )

    $bytes = [long]0
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($name in @($names)) {
        $indexDir = Join-Path $ws.IndexDir $name
        if (!(Test-Path -LiteralPath (toLongPath $indexDir) -PathType Container)) {
            $missing.Add($name)
            continue
        }
        foreach ($file in @(getIndexTsvFilesNoFollow $indexDir)) {
            $bytes += $file.Length
        }
    }
    return @{ Bytes = $bytes; Missing = $missing.ToArray() }
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
    $prefixLength = $indexDir.TrimEnd("\").Length + 1
    $contentIndexFiles = New-Object System.Collections.Generic.List[string]
    # 途中で throw して抜けても下のフォルダを掴んだまま残らないよう、列挙子（EnumerateFiles）でなく配列で受ける
    foreach ($tsvFile in @(getIndexTsvFilesNoFollow $indexDir)) {
        $rel = (fromLongPath $tsvFile.FullName).Substring($prefixLength)
        $fileName = [System.IO.Path]::GetFileName($rel)
        if ($null -eq (readContentIndexFileName $fileName)) {
            throw "インデックスとして扱えないファイルが残っています（${rel}）。更新を最後まで行うか、そのファイルを取り除いてからエクスポートしてください。"
        }
        $contentIndexFiles.Add($rel)
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

            $contentResult = writeContentIndexFilesToArchive $archive $indexDir $contentIndexFiles $manifestFiles
            $totalBytes += $contentResult.Bytes

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
    return @{ Name = $name; Path = $destPath; Files = $manifestFiles.Count; Bytes = $totalBytes; Omitted = @($contentResult.Omitted) }
}

function writeContentIndexFilesToArchive {
    # 本文インデックスのファイルを zip に入れ、目録の files（manifestFiles）に足す。
    # インポートが断る大きさ（testImportedContentIndexFile と同じ判定）のファイルは、書き出せても読み込めないので、入れずに Omitted に控える。
    # 戻り値: @{ Bytes（入れたバイト数）; Omitted（入れなかったファイルの配列 @{ Path（インデックスの中の相対パス）; Books（中に入っていた元のファイルの相対パス）} ）}
    param (
        $archive,
        [string]$indexDir,
        [string[]]$contentIndexFiles,   # インデックスのフォルダからの相対パス
        $manifestFiles
    )

    $bytes = [long]0
    $omitted = New-Object System.Collections.Generic.List[object]
    foreach ($rel in $contentIndexFiles) {
        $source = Join-Path $indexDir $rel
        $size = ([System.IO.FileInfo]::new((toLongPath $source))).Length
        if (testImportedContentIndexFile $size $null) {
            $omitted.Add(@{ Path = $rel; Books = (readContentIndexBookNamesStreaming $source) })
            continue
        }
        $entryPath = "content_index/" + $rel.Replace("\", "/")
        $written = writeArchiveFile $archive $entryPath $source
        $manifestFiles.Add([ordered]@{ path = $entryPath; size = [long]$written.Size; sha256 = $written.Sha256 })
        $bytes += $written.Size
    }
    return @{ Bytes = $bytes; Omitted = $omitted.ToArray() }
}

function readArchiveEntryBytesLimited {
    # zip のエントリーを、実際に読んだバイト数を数えながらバイト列で読む。maxBytes を超えたら、それ以上読まずに $null を返す
    # （中央ディレクトリの展開後の大きさの申告は、送り手が偽れるため当てにしない）
    param (
        $entry,
        [long]$maxBytes
    )

    $source = $entry.Open()
    try {
        $memory = New-Object System.IO.MemoryStream
        try {
            $buffer = New-Object byte[] 1MB
            while ($true) {
                $read = $source.Read($buffer, 0, $buffer.Length)
                if ($read -le 0) {
                    break
                }
                if ($memory.Length + $read -gt $maxBytes) {
                    return $null
                }
                $memory.Write($buffer, 0, $read)
            }
            return , $memory.ToArray()
        } finally {
            $memory.Dispose()
        }
    } finally {
        $source.Dispose()
    }
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
                # 申告の大きさ（送り手が書ける）に頼らず、実際に読んだ大きさで打ち切る。超えたら大きすぎるものとして断る
                $manifestData = readArchiveEntryBytesLimited $manifestEntry ${indexArchiveMaxManifestBytes}
                if ($null -eq $manifestData) {
                    $manifestBytes = ${indexArchiveMaxManifestBytes} + 1
                } else {
                    try {
                        $manifest = ConvertFrom-Json ([System.Text.Encoding]::UTF8.GetString($manifestData).TrimStart([char]0xFEFF))
                    } catch {
                        $manifest = $null
                    }
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

function getImportArchiveInfo {
    # readIndexArchiveInfo の結果に、元のフォルダがこの PC にあるか（SourceExists）を足して返す（インポートの裏の仕事で使う）。
    # 目録の元のフォルダは zip の送り手が自由に書けるため、ネットワークの場所なら接続せず（Test-Path を呼ばず）$null（分からない）にする。
    # 利用者が元のフォルダを見て確かめる前に、サインイン情報が相手のサーバーへ送られないようにするため
    param (
        [string]$zipPath
    )

    $info = readIndexArchiveInfo $zipPath
    $exists = $false
    try {
        $folder = normalizeFolderPath ([string]$info.SourceFolder)
        if (testLocalDrivePath $folder) {
            $exists = [bool](Test-Path -LiteralPath $folder -PathType Container)
        } elseif ($folder) {
            # ローカルのドライブと分かるときだけ調べる。ネットワークの場所や、デバイス名の形（\\.\UNC\・\\.\GLOBALROOT\ など）は
            # 調べず「分からない」にする
            $exists = $null
        }
    } catch {
        # 使えない文字を含むなど。無いものとして扱う
    }
    $info.SourceExists = $exists
    return $info
}

function getImportUsedIndexNames {
    # 今使われているインデックス名（設定の targetFolders・indexSources、インデックスのフォルダ（store\content\・古い content_index\）直下のフォルダ、
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
    # 作り直す前の形式のインデックス（古い置き場）の名前も使われているものとして数える（同じ名前を取り込んで、新旧が混ざらないように）。
    # ワークスペースに残っているインデックスの集め方は getWorkspaceIndexFolders（画面の追加・名前の変更と同じ）
    foreach ($folder in @(getWorkspaceIndexFolders $ws)) {
        [void]$names.Add($folder.Name)
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
        [string]$destPath,
        [string]$root,
        $progress = $null
    )

    # 書き出すのは、root の下で途中にリンクが無いときだけ。同じ名前のファイルが既にあれば上書きせず、例外にする（CreateNew）
    if (!(newDirectoryUnder ([System.IO.Path]::GetDirectoryName($destPath)) $root)) {
        throw "置き場の外か、途中にリンクがあるため、展開できません: ${destPath}"
    }
    $source = $entry.Open()
    $sha = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    $total = [long]0
    try {
        $dest = New-Object System.IO.FileStream((toLongPath $destPath), [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
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
                if ($null -ne $progress) {
                    $progress.Done = [long]$progress.Done + $read
                }
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
    # progress は省ける: [hashtable]::Synchronized(@{ Done = 0L; Total = 0L })。展開の前に Total を目録の大きさの合計にし、展開するたびに Done を足す（画面が別スレッドから読む）。
    # collisionMode（Rename・Overwrite・Cancel）が Cancel なら $null を返す（ワークスペースは変えない）。
    # name・sourceFolder を省くと目録の値を使う。インデックス作成のロックを取れなければ例外
    param (
        [string]$zipPath,
        [string]$collisionMode,
        [string]$name = "",
        [string]$sourceFolder = "",
        $ws = $workspace,
        [string]$settingsPath = ${settingsFile},
        [scriptblock]$getFreeSpace = { param ($root) getWorkspaceFreeSpace $root },
        $progress = $null
    )

    $lock = enterIndexerLock $ws.Dir
    if (!$lock.Acquired) {
        throw "更新中はインポートできません。更新が終わってからやり直してください。"
    }
    try {
        return (importIndexCore $zipPath $collisionMode $name $sourceFolder $ws $settingsPath $getFreeSpace $progress)
    } finally {
        exitIndexerLock $lock
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
        [string]$folder,
        [string]$workRoot,
        $progress = $null
    )

    $newDir = Join-Path $workDir "new"
    # 消す・作るのは、workRoot の下で途中にリンクが無いときだけ（偽なら何も書かずに例外にする）
    if (!(removeDirectoryUnder $workDir $workRoot) -or !(newDirectoryUnder $newDir $workRoot)) {
        throw "置き場の外か、途中にリンクがあるため、展開できません: ${workDir}"
    }
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
                extractArchiveEntryToFile $entry $size $sha256 $statusTmpPath $workDir $progress
                $bytes = [System.IO.File]::ReadAllBytes((toLongPath $statusTmpPath))
                $text = ${utf8Bom}.GetString($bytes)
                $statusLines = @($text -split "\r?\n" | Where-Object { $_ -ne "" })
                $reason = testImportedStatusLines $statusLines
                if ($reason) {
                    throw "ingest_status.tsv を読み込めません（${reason}）"
                }
            } else {
                # 本文インデックスのファイル: 大きすぎれば展開しない。展開して SHA-256 を比べたあと、検索と更新が読めるかを確かめる
                $sizeReason = testImportedContentIndexFile $size $null
                if ($sizeReason) {
                    throw "インデックスのファイル「${path}」を読み込めません（${sizeReason}）。この版の tebunko でエクスポートし直すと、大きすぎるファイルを含めずに書き出せます。"
                }
                $rel = $path.Substring("content_index/".Length).Replace("/", "\")
                $target = Join-Path $newDir $rel
                extractArchiveEntryToFile $entry $size $sha256 $target $workDir $progress
                $readReason = testImportedContentIndexFile $size (readContentIndexText $target)
                if ($readReason) {
                    throw "インデックスのファイル「${path}」を読み込めません（${readReason}）。エクスポートした側で更新し直してから、もう一度エクスポートしてください。"
                }
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
        $original = $_
        try {
            [void](removeDirectoryUnder $workDir $workRoot)
        } catch {
            # 片付けに失敗しても、元の失敗の理由を隠さない
        }
        throw $original
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
        [string]$indexRoot,
        [string]$workRoot
    )

    # ウイルス対策ソフトが、作ったばかり・書いたばかりのフォルダ（$targetDir は前の取り込みで、$newDir はこの
    # インポートの展開で、それぞれ書いたばかり）を一時的に掴んでいることがあるため、moveDirectoryRetry で試し直す
    # （removeDirectoryRetry と同じ理由。docs/design/index-data/format.md「インポート」を参照）
    # 移すのは、targetDir が indexRoot の下、newDir と previousDir が workRoot の下で、どれも途中にリンクが無いときだけ
    # （moveDirectoryUnder。偽なら何も移さずに例外にする）
    $outside = "置き場の外か、途中にリンクがあるため、インデックスを入れ替えられません: ${targetDir}"
    $hadPrevious = Test-Path -LiteralPath (toLongPath $targetDir) -PathType Container
    if ($hadPrevious) {
        if (!(moveDirectoryUnder $targetDir $previousDir $workRoot $indexRoot)) {
            throw $outside
        }
    }
    try {
        [System.IO.Directory]::CreateDirectory((toLongPath $indexRoot)) | Out-Null
        if (!(moveDirectoryUnder $newDir $targetDir $indexRoot $workRoot)) {
            throw $outside
        }
    } catch {
        if ($hadPrevious) {
            [void](moveDirectoryUnder $previousDir $targetDir $indexRoot $workRoot)
        }
        throw
    }
    return @{ TargetDir = $targetDir; PreviousDir = $previousDir; HadPrevious = $hadPrevious; IndexRoot = $indexRoot; WorkRoot = $workRoot }
}

function restoreSwappedIndexDir {
    # swapInImportedIndexDir で入れ替えたフォルダを、入れ替える前に戻す（入れたものを消し、退避した前のフォルダを戻す）
    param (
        $swap
    )

    if (!(removeDirectoryUnder $swap.TargetDir $swap.IndexRoot)) {
        throw "置き場の外か、途中にリンクがあるため、入れ替えを戻せません: $($swap.TargetDir)"
    }
    if ($swap.HadPrevious) {
        if (!(moveDirectoryUnder $swap.PreviousDir $swap.TargetDir $swap.IndexRoot $swap.WorkRoot)) {
            throw "置き場の外か、途中にリンクがあるため、入れ替えを戻せません: $($swap.TargetDir)"
        }
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
        [scriptblock]$getFreeSpace,
        $progress = $null
    )

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

        $fileEntries = @($manifest.files)
        $totalBytes = [long](($fileEntries | ForEach-Object { [long]$_.size } | Measure-Object -Sum).Sum)
        if ($null -ne $progress) { $progress.Total = $totalBytes }
        $spaceReason = testImportFreeSpace $totalBytes ([System.IO.Path]::GetPathRoot($ws.Dir)) $getFreeSpace
        if ($spaceReason) {
            throw $spaceReason
        }

        # 同じ元のフォルダが、別の名前のクロール対象フォルダに既にある・入れ子になっていれば止める（getTargetFolders は同じフォルダの 2 つ目以降を読まないため、
        # 登録しても設定に残らず、次のインデックス作成の確認に、このインデックスが「削除予定」として出る。画面の追加・編集と同じ getIndexFolderConflict の決まり）。
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
        $expanded = expandImportArchive $archive $fileEntries $workDir $finalName $folder $ws.PublishDir $progress
        $targetIndexDir = Join-Path $ws.IndexDir $finalName
        try {
            # 前の版が作ったシステムインデックスは、この版では使わない。消せなくても取り込みは続け、警告に 1 行足す（次のインデックス作成でも消す）
            $oldSearchFiles = removeOldSystemIndex $ws.Dir
            if (!$oldSearchFiles.Ok) {
                $warnings.Add("前の版が作った検索用のファイルを消せませんでした（$($oldSearchFiles.Reason)）。次のインデックス作成でもう一度消します。")
            }
            $before = registerImportedIndexInSettings $finalName $folder $overwrite $targetIndexDir $settingsPath
            try {
                $swap = swapInImportedIndexDir $expanded.NewDir $targetIndexDir (Join-Path $workDir "previous") $ws.IndexDir $ws.PublishDir
                try {
                    # 取り込んだインデックスは、作り終えたものとして目録に載せる
                    if (!(setIndexFormatRow $finalName ${catalogRowFinished} $ws.FormatFile)) {
                        throw "インデックスの目録を書けません（新しい版の tebunko で作られたワークスペースの可能性があります）。"
                    }
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
            $original = $_
            try {
                [void](removeDirectoryUnder $workDir $ws.PublishDir)
            } catch {
                # 片付けに失敗しても、元の失敗の理由を隠さない
            }
            throw $original
        }

        # 11. 上書きなら、前の gram 索引を消す（消せなくても続ける。次のインデックス作成で整理される）
        if ($overwrite) {
            removeGramIndexOfWorkspace $finalName $ws.IndexDir
        }
        # 12. 作業フォルダを消す
        if (!(removeDirectoryUnder $workDir $ws.PublishDir)) {
            throw "置き場の外か、途中にリンクがあるため、作業フォルダを消せません: ${workDir}"
        }

        return @{
            Name = $finalName; SourcePath = $folder; Enabled = [System.IO.Directory]::Exists((toLongPath $folder))
            Files = $fileEntries.Count; Bytes = $totalBytes; Warnings = $warnings.ToArray()
        }
    } finally {
        $archive.Dispose()
    }
}
