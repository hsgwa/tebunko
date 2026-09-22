# 制限モードのインデックス作成（状態層）。制限言語モードで動く書き方だけで書く。
#
# いつものインデクサ（indexer.ps1）と同じ取り込み一覧・同じ形のインデックスを作る。取り込む対象の決め方（createTargetList）、
# 取り込み一覧の読み書き（indexer_state.ps1）、インデックスへの入れ替え（publishTsv）は、いつものインデクサと同じ関数を使う。
# 違うのは次の点:
#   ・読めるのは Word・PowerPoint の新形式（.docx .docm .pptx .pptm）の ZIP だけ（office_reader_clm.ps1。Office は使わない）。
#     Excel・旧形式（.xls .doc .ppt 等）・ZIP でないファイル（パスワード付きなど）は取り込まずに飛ばし、
#     取り込み一覧では「未取り込み」のまま残す（いつもの画面が使える PC で取り込めるようにする。「失敗」にすると、
#     いつもの画面で「失敗分も再取り込み」を選ぶまで取り込まれないため）
#   ・Office を使わないため、制限時間・Office の再起動は無い
#   ・二重起動は名前付きミューテックスの代わりに、PID を書いたロックファイル（work\インデックス作成中.lock）で防ぐ
#   ・進み具合は Write-Progress でコンソールに出す（画面の進捗ファイルは書かない）
#   ・以前の版の形式のインデックス（取り込み一覧にインデックス名の無いもの）は移せないため、いつもの画面で作ってもらう
# 取り込み済みの分は 1 件ごとに取り込み一覧へ追記するため、Ctrl+C で止めても次回は続きから取り込める。

# 制限モードで取り込める拡張子
${restrictedIngestExtensions} = @(".docx", ".docm", ".pptx", ".pptm")

# 二重起動を防ぐロックファイル（中身は "PID<TAB>コンピューター名<TAB>開始日時"）
${restrictedLockFile} = "${workDir}\インデックス作成中.lock"

# 取り込み中に続けて強制終了した回数がこれに達したファイルは失敗とする（indexer.ps1 の $interruptLimit と同じ）
${restrictedInterruptLimit} = 2

function readZipSignature {
    # ファイルの先頭が ZIP のシグネチャ（PK\x03\x04）か（office_reader.ps1 の isZipFile と同じ。FileStream の代わりに Get-Content で読む）
    param (
        [string]$path
    )

    $head = @(Get-Content -LiteralPath (toLongPath $path) -Encoding Byte -TotalCount 4 -ErrorAction Stop)
    return ($head.Count -eq 4 -and $head[0] -eq 0x50 -and $head[1] -eq 0x4B -and $head[2] -eq 0x03 -and $head[3] -eq 0x04)
}

function enterIndexingLock {
    # ロックファイルを作る。ほかのインデックス作成（同じ PC の制限モード・ほかの PC）が使っていれば、その理由を返す（取れたら空）。
    # 同じ PC で、書いた PID のプロセスが終わっていれば、強制終了で残ったものとみなして取り直す
    param (
        [string]$path = ${restrictedLockFile}
    )

    if (Test-Path -LiteralPath $path) {
        $fields = @((@(readListFile $path) + @(""))[0].Split("`t"))
        if ($fields.Count -ge 3) {
            $alive = $false
            if ($fields[1] -eq $env:COMPUTERNAME -and $fields[0] -match '^[0-9]{1,9}$') {
                $alive = [bool](Get-Process -Id ([int]$fields[0]) -ErrorAction SilentlyContinue)
            } elseif ($fields[1] -ne $env:COMPUTERNAME) {
                $alive = $true  # ほかの PC のプロセスは確かめられない
            }
            if ($alive) {
                return "ほかのインデックス作成が実行中です（$($fields[1]) で $($fields[2]) から）。終わってから実行してください。" +
                    "（実行中でないのにこの表示が続くときは、$path を削除してください）"
            }
        }
    }
    writeListFile $path @("$PID`t$($env:COMPUTERNAME)`t$(formatFileTime (Get-Date))")
    return ""
}

function exitIndexingLock {
    param (
        [string]$path = ${restrictedLockFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
    }
}

function getRestrictedTsvCounts {
    # getIndexTsvCounts（インデックスのフォルダごとの TSV の数）と同じものを、制限言語モードで使える書き方で返す。
    # 利用者が work\index のフォルダ・TSV を直接削除した場合に、取り込み一覧が「済」のまま検索できなくなるのを防ぐ
    # （createTargetList → testIndexComplete が使う）。
    # .NET の列挙ではなく Get-ChildItem で 1 回だけ列挙する。ハッシュテーブルのキーは大文字・小文字を区別しない
    param (
        [string]$dir = ${indexDir}
    )

    $counts = @{}
    $root = (toLongPath $dir).TrimEnd("\")
    if (!(Test-Path -LiteralPath $root -PathType Container)) {
        return $counts
    }
    $prefix = $root.Length + 1
    foreach ($entry in @(Get-ChildItem -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue)) {
        $relative = $entry.FullName.Substring($prefix)
        if ($entry.PSIsContainer) {
            # 内容が空のファイル（TSV 0 件）はフォルダだけが残るため、フォルダも 0 件として数える
            if (!$counts.ContainsKey($relative)) {
                $counts[$relative] = 0
            }
            continue
        }
        if ($entry.Extension -ne ".tsv") {
            continue
        }
        $key = getPathParent $relative
        if ($key -eq "") {
            continue  # インデックスのフォルダの直下の TSV（以前の形式）は、どのファイルのものか分からないため数えない
        }
        $count = 0
        if ($counts.ContainsKey($key)) {
            $count = $counts[$key]
        }
        # 0 バイトの TSV があるフォルダは壊れているものとして扱い、取り込み直す
        if ($entry.Length -eq 0 -or $count -eq ${indexBrokenCount}) {
            $counts[$key] = ${indexBrokenCount}
        } else {
            $counts[$key] = $count + 1
        }
    }
    return $counts
}

function testRestrictedIngestable {
    # 制限モードで取り込める拡張子か
    param (
        [string]$relPath
    )

    $extension = if ($relPath -match '(\.[^.\\]*)$') { $Matches[1].ToLowerInvariant() } else { "" }
    return (${restrictedIngestExtensions} -contains $extension)
}

function ingestRestrictedFile {
    # 1 ファイルを読んで作業フォルダ（tmpDir）に TSV を書き、書いた数を返す。ZIP でなければ $null（取り込まずに飛ばす）
    param (
        [string]$sourcePath
    )

    if (!(readZipSignature $sourcePath)) {
        return $null
    }
    if ($sourcePath -match '\.ppt[xm]$') {
        $units = readPptxUnitsClm $sourcePath ${tmpDir}
    } else {
        $units = readDocxUnitsClm $sourcePath ${tmpDir}
    }
    return (writeUnitsClm $units ${tmpDir})
}

function invokeRestrictedIndexing {
    # 設定のクロール対象フォルダ（チェックの付いたもの）からインデックスを作り、結果を @{ Success; Failed; Skipped; Dropped; Stopped; Messages } で返す。
    # 続けられないエラー（クロール対象フォルダが無い・ほかで実行中など）は例外にする
    param (
        [bool]$retryFailed = $false
    )

    $reason = enterIndexingLock
    if ($reason) {
        throw $reason
    }
    $result = @{ Success = 0; Failed = @(); Skipped = @(); Dropped = 0; Targets = 0 }
    $folders = $null
    $rows = @{}
    $dropped = @{}
    $started = $false
    try {
        $targetFolders = @(getTargetFolders)
        if ($targetFolders.Count -eq 0) {
            throw "クロール対象フォルダがありません。メニューの［インデックスを管理する］でフォルダを追加してください。"
        }
        if (@($targetFolders | Where-Object { $_.Enabled }).Count -eq 0) {
            throw "チェックの付いたクロール対象フォルダがありません。メニューの［インデックスを管理する］でチェックを付けてください。"
        }

        foreach ($dir in @(${workDir}, ${indexDir}, ${tmpDir}, ${publishDir})) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        removeStaleTmpDirs

        $status = readStatusFile
        if (@($status.Folders | Where-Object { -not $_.Name }).Count -gt 0 -or (Test-Path -LiteralPath "${indexDir}_移行中")) {
            throw "以前の版の形式のインデックスが残っています。制限モードでは形式を移せないため、いつもの画面が使える PC で一度インデックスを作成してください。"
        }
        $folders = @(assignIndexNames $targetFolders $status.Folders)
        removeDroppedFolders $folders $status.Folders
        if (@($targetFolders | Where-Object { -not $_.Name }).Count -gt 0) {
            writeTargetFolders $folders
        }

        # クロール（いつものインデクサと同じ。Excel・旧形式も数え、取り込み一覧の行を保つ）
        $previous = $status.Rows
        $counts = getRestrictedTsvCounts
        $targets = @{}
        $failedRows = @()
        foreach ($folder in $folders) {
            if (-not $folder.Enabled -or !(Test-Path -LiteralPath $folder.Path -PathType Container)) {
                if ($folder.Enabled) {
                    Write-Host "  [$($folder.Name)] $($folder.Path) … フォルダが見つからないため取り込みません" -ForegroundColor Yellow
                } else {
                    Write-Host "  [$($folder.Name)] $($folder.Path) … チェックなしのため取り込みません（インデックスはそのまま残します）"
                }
                $prefix = "$($folder.Name)\"
                foreach ($key in @($previous.Keys)) {
                    if ($key.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $rows[$rows.Count] = $previous[$key]
                    }
                }
                continue
            }
            Write-Host "  [$($folder.Name)] $($folder.Path) を調べています…"
            $list = createTargetList $folder $previous $counts
            foreach ($row in $list.Rows) { $rows[$rows.Count] = $row }
            foreach ($row in $list.Targets) { $targets[$targets.Count] = $row }
            $failedRows += $list.Failed
        }
        if ($retryFailed) {
            foreach ($row in $failedRows) { $targets[$targets.Count] = $row }
        }
        $targetList = @(for ($i = 0; $i -lt $targets.Count; $i++) { $targets[$i] })

        # 前回、取り込み中に強制終了したファイルは最後に回し、続けて止まったものは失敗にする（いつものインデクサと同じ）
        $interrupted = readIngestingFile
        if ($interrupted) {
            $row = @($targetList | Where-Object { $_.相対パス -eq $interrupted.RelPath }) | Select-Object -First 1
            if (!$row) {
                $interrupted = $null
                removeIngestingFile
            } elseif ($interrupted.Count -ge ${restrictedInterruptLimit}) {
                $targetList = @($targetList | Where-Object { $_.相対パス -ne $interrupted.RelPath })
                $row.状態 = ${stateFailed}
                $row.TSV数 = ""
                $row.取り込み日時 = formatFileTime (Get-Date)
                $row.エラー = "取り込み中に $($interrupted.Count) 回続けて強制終了されたため、取り込みを中止しました"
                $interrupted = $null
                removeIngestingFile
            } else {
                $targetList = @($targetList | Where-Object { $_.相対パス -ne $interrupted.RelPath }) + @($row)
            }
        }

        $rowList = @(for ($i = 0; $i -lt $rows.Count; $i++) { $rows[$i] })
        writeStatusFile $folders $rowList
        writeSourceFolderFile $folders

        $folderByName = @{}
        foreach ($folder in $folders) { $folderByName[$folder.Name] = $folder.Path.TrimEnd("\") }
        $result.Targets = $targetList.Count
        $started = $true
        for ($i = 0; $i -lt $targetList.Count; $i++) {
            $row = $targetList[$i]
            $relPath = $row.相対パス
            Write-Progress -Activity "インデックスを作成しています" -Status ("{0} / {1}  {2}" -f ($i + 1), $targetList.Count, $relPath) -PercentComplete ([int](100 * $i / $targetList.Count))
            if (!(testRestrictedIngestable $relPath)) {
                $result.Skipped += $relPath
                continue
            }
            $parts = splitIndexRelPath $relPath
            $sourceFolder = $folderByName[$parts.Name]
            $sourcePath = Join-Path $sourceFolder $parts.Rest
            if (!(Test-Path -LiteralPath (toLongPath $sourcePath) -PathType Leaf)) {
                if (!(Test-Path -LiteralPath $sourceFolder -PathType Container)) {
                    throw "クロール対象フォルダが見つからなくなったため、インデックス作成を中止しました: ${sourceFolder}（残りは未取り込みのまま残しました）"
                }
                removeBookDir (getBookDir $relPath)
                $dropped[$relPath] = $true
                $result.Dropped++
                continue
            }

            $startCount = 1
            if ($interrupted -and $interrupted.RelPath -eq $relPath) {
                $startCount = $interrupted.Count + 1
            }
            writeIngestingFile $relPath $startCount
            try {
                clearTmpDir
                $tsvCount = ingestRestrictedFile $sourcePath
                if ($null -eq $tsvCount) {
                    # ZIP でない（パスワード付き・中身が旧形式など）: 取り込まずに「未取り込み」のまま残す
                    $result.Skipped += $relPath
                    removeIngestingFile
                    continue
                }
                publishTsv (getBookDir $relPath)
                $row.状態 = ${stateDone}
                $row.TSV数 = [string]$tsvCount
                $row.エラー = ""
                $row.抽出版 = [string](getExtractVersion $relPath)
                $result.Success++
            } catch {
                $message = ($_.Exception.Message -replace "\s+", " ").Trim()
                Write-Host "    取り込みに失敗しました: $relPath（$message）" -ForegroundColor Red
                $row.状態 = ${stateFailed}
                $row.TSV数 = ""
                $row.エラー = $message
                $row.抽出版 = ""
                $result.Failed += @{ RelPath = $relPath; Message = $message }
            }
            $row.取り込み日時 = formatFileTime (Get-Date)
            addStatusRow $row
            removeIngestingFile
        }
    } finally {
        Write-Progress -Activity "インデックスを作成しています" -Completed
        if ($started) {
            removeIngestingFile
            writeStatusFile $folders @($rowList | Where-Object { $_ -and !$dropped.ContainsKey([string]$_.相対パス) })
            writeSourceFolderFile $folders
        }
        removeTmpDir
        exitIndexingLock
    }
    return $result
}
