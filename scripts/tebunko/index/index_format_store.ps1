# インデックスの格納形式の目録（store\format.tsv）の読み書きと、前の版の置き場の片付け（状態層）。
# 版・状態の判定そのものは content_index_format.ps1（判断層）の getIndexFormatState に任せ、ここはファイルを読み書きする。
#
# 目録の形（UTF-8 BOM 付き・タブ区切り）:
#   版=2
#   アプリの版=<書いた tebunko の版。調べるためで、判定には使わない>
#   <インデックス名><TAB>作成中|作り終えた
# 知らない行・知らない列は読み飛ばす（後の版が版を上げずに足せるように。足してよいものの規則は format.md「目録の版の規則」）

# 作り直しが最後まで終わったときに消す、前の版の置き場（Workspace のプロパティ名）。
# 後の版が古い置き場を足すときは、ここに名前を足す
${oldIndexEntries} = @("OldIndexDir", "OldStatusFile")

${indexFormatVersionKey} = "版"
${indexFormatAppVersionKey} = "アプリの版"


function readIndexFormat {
    # 目録を読む: @{ Status; Version; AppVersion; Indexes }
    #   Status : Ok（今の版）・Missing（ファイルが無い）・Broken（読めない・版の行が無い）・Newer（今より新しい版）・Older（今より古い版）
    #   Indexes: インデックス名 → 作成中・作り終えた（Ok のときだけ入る）
    param (
        [string]$path = $workspace.FormatFile
    )

    $result = @{ Status = "Missing"; Version = 0; AppVersion = ""; Indexes = [ordered]@{} }
    if (![System.IO.File]::Exists((toLongPath $path))) {
        return $result
    }
    $lines = $null
    for ($attempt = 1; $attempt -le 5 -and $null -eq $lines; $attempt++) {
        try {
            $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
            $stream = New-Object System.IO.FileStream((toLongPath $path), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
            try {
                $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
                $lines = @($reader.ReadToEnd() -split "\r?\n")
            } finally {
                $stream.Dispose()
            }
        } catch [System.IO.FileNotFoundException] {
            return $result
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds 100
        }
    }
    if ($null -eq $lines) {
        throw "インデックスの目録を読めません: $path"
    }

    $nonEmpty = @($lines | Where-Object { $_ -ne "" })
    $prefix = "${indexFormatVersionKey}="
    $version = 0
    if ($nonEmpty.Count -eq 0 -or !([string]$nonEmpty[0]).StartsWith($prefix, [System.StringComparison]::Ordinal) -or
        !([int]::TryParse(([string]$nonEmpty[0]).Substring($prefix.Length), [ref]$version))) {
        $result.Status = "Broken"
        return $result
    }
    $result.Version = $version
    if ($version -gt ${indexFormatVersion}) {
        $result.Status = "Newer"
        return $result
    }
    if ($version -lt ${indexFormatVersion}) {
        $result.Status = "Older"
        return $result
    }
    $result.Status = "Ok"
    $appKey = "${indexFormatAppVersionKey}="
    foreach ($line in @($nonEmpty | Select-Object -Skip 1)) {
        if ($line.StartsWith($appKey, [System.StringComparison]::Ordinal)) {
            $result.AppVersion = $line.Substring($appKey.Length)
            continue
        }
        $columns = $line -split "`t"
        if ($columns.Count -lt 2 -or $columns[0] -eq "") {
            continue
        }
        $result.Indexes[$columns[0]] = $columns[1]
    }
    return $result
}


function writeIndexFormat {
    # 目録を書く（一時ファイルに書いてから入れ替える）。indexes は 名前 → 作成中・作り終えた
    param (
        $indexes,
        [string]$path = $workspace.FormatFile,
        [string]$appVersion = ""
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("${indexFormatVersionKey}=${indexFormatVersion}")
    $lines.Add("${indexFormatAppVersionKey}=${appVersion}")
    foreach ($name in @($indexes.Keys | Sort-Object)) {
        $lines.Add("${name}`t$($indexes[$name])")
    }
    writeTextLinesAtomic $path $lines
}


function getIndexFormatAppVersion {
    # 目録に書く tebunko の版（調べるためで、判定には使わない）。分からなければ空
    try {
        $version = readVersionFile (Join-Path ${rootDir} "VERSION.txt")
        if ($version) { return [string]$version.Tag }
    } catch {
    }
    return ""
}

function updateIndexFormat {
    # 目録を排他で読み、change（{ param($indexes) }。名前 → 状態の表を書き換える）を当てて書く。
    # 今より新しい版の目録は書き換えない（$false を返す）。壊れた・無い目録は空の表から始める
    # （壊れたものは呼び出し側が先に repairIndexFormat で救う）。ロックを取れなければ $false
    param (
        [scriptblock]$change,
        [string]$path = $workspace.FormatFile,
        [string]$appVersion = ""
    )

    if ($appVersion -eq "") {
        $appVersion = getIndexFormatAppVersion
    }
    [System.IO.Directory]::CreateDirectory((toLongPath ([System.IO.Path]::GetDirectoryName($path)))) | Out-Null
    $lock = $null
    for ($attempt = 1; $attempt -le 20 -and $null -eq $lock; $attempt++) {
        try {
            $lock = New-Object System.IO.FileStream((toLongPath "${path}.lock"), [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None, 8, [System.IO.FileOptions]::DeleteOnClose)
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds 100
        }
    }
    if ($null -eq $lock) {
        return $false
    }
    try {
        $current = readIndexFormat $path
        if ($current.Status -eq "Newer") {
            return $false
        }
        $indexes = [ordered]@{}
        if ($current.Status -eq "Ok") {
            foreach ($key in $current.Indexes.Keys) { $indexes[$key] = $current.Indexes[$key] }
        }
        & $change $indexes
        writeIndexFormat $indexes $path $appVersion
    } finally {
        $lock.Dispose()
    }
    return $true
}


function setIndexFormatRow {
    # 目録のインデックス 1 つの行の状態を書く（作成中・作り終えた）。行が無ければ足す
    param (
        [string]$name,
        [string]$state,
        [string]$path = $workspace.FormatFile,
        [string]$appVersion = ""
    )

    return (updateIndexFormat { param ($indexes) $indexes[$name] = $state }.GetNewClosure() $path $appVersion)
}


function removeIndexFormatRow {
    # 目録のインデックス 1 つの行を消す（インデックスの削除）。目録が無ければ何もしない
    param (
        [string]$name,
        [string]$path = $workspace.FormatFile
    )

    if (![System.IO.File]::Exists((toLongPath $path))) {
        return $true
    }
    return (updateIndexFormat { param ($indexes) $indexes.Remove($name) }.GetNewClosure() $path)
}


function renameIndexFormatRow {
    # 目録のインデックスの行の名前を変える（インデックスの改名）。行が無ければ何もしない
    param (
        [string]$oldName,
        [string]$newName,
        [string]$path = $workspace.FormatFile
    )

    if (![System.IO.File]::Exists((toLongPath $path))) {
        return $true
    }
    return (updateIndexFormat {
        param ($indexes)
        if ($indexes.Contains($oldName)) {
            $state = $indexes[$oldName]
            $indexes.Remove($oldName)
            $indexes[$newName] = $state
        }
    }.GetNewClosure() $path)
}


function testIndexDirHasContent {
    # インデックスのフォルダに、source_folder.txt 以外のファイル（本文インデックスや作成途中の TSV）が 1 つでもあるか
    param (
        [string]$dir
    )

    $long = toLongPath $dir
    if (![System.IO.Directory]::Exists($long)) {
        return $false
    }
    # リンクは辿らない（getFilesNoFollow）。リンクの先を「中身がある」と数えない
    $found = getFilesNoFollow $dir
    return (@($found.Files | Where-Object { $_.Name -ne ${sourceFolderFileName} }).Count -gt 0)
}


function getIndexNamesWithContent {
    # 置き場（root）の直下のフォルダのうち、本文インデックスのファイルがあるもの（インデックス）の名前
    param (
        [string]$root
    )

    $names = New-Object System.Collections.Generic.List[string]
    $long = toLongPath $root
    if (![System.IO.Directory]::Exists($long)) {
        return , $names.ToArray()
    }
    foreach ($sub in [System.IO.Directory]::GetDirectories($long)) {
        if (testIndexDirHasContent $sub) {
            $names.Add([System.IO.Path]::GetFileName($sub))
        }
    }
    return , $names.ToArray()
}


function findOldIndexNames {
    # 古い置き場（content_index\）で本文インデックスを持つインデックスの名前
    param (
        $workspace = $script:workspace
    )

    return , (getIndexNamesWithContent $workspace.OldIndexDir)
}


function testIndexFileCurrentVersion {
    # 本文インデックスのファイルの 1 行目が今の版か（先頭だけを読む）
    param (
        [string]$path
    )

    try {
        $stream = New-Object System.IO.FileStream((toLongPath $path), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::Unicode, $true)
            $head = $reader.ReadLine()
            return (testContentIndexVersionLine $head)
        } finally {
            $stream.Dispose()
        }
    } catch {
        return $false
    }
}


function repairIndexFormat {
    # 目録が壊れたとき、本文インデックスの版の行から書き直す。今の版のファイルがあるインデックスの行を「作成中」にする
    # （作り終えたかは分からないため。取り込み一覧が残っていれば、差分の更新で済む）。書き直した目録を返す
    param (
        $workspace = $script:workspace
    )

    $indexes = [ordered]@{}
    $long = toLongPath $workspace.IndexDir
    if ([System.IO.Directory]::Exists($long)) {
        foreach ($sub in [System.IO.Directory]::GetDirectories($long)) {
            $found = $false
            $subFiles = getFilesNoFollow $sub
            foreach ($file in $subFiles.Files) {
                if ($file.Name.StartsWith(${contentIndexFileNamePrefix}, [System.StringComparison]::Ordinal) -and (testIndexFileCurrentVersion $file.FullName)) {
                    $found = $true
                    break
                }
            }
            if ($found) {
                $indexes[[System.IO.Path]::GetFileName($sub)] = ${catalogRowBuilding}
            }
        }
    }
    # 画面の裏とインデクサーが同時に呼びうるため、インデクサーの更新と同じロックの中で書く。
    # ロックの中で、ほかが先に直していた（読める目録になっていた）ときは書き直さない
    $computed = $indexes
    $written = updateIndexFormat ({
            param ($table)
            if ($table.Count -eq 0) {
                foreach ($key in $computed.Keys) { $table[$key] = $computed[$key] }
            }
        }.GetNewClosure()) $workspace.FormatFile
    $result = readIndexFormat $workspace.FormatFile
    # ロックの中で、ほかが今より新しい版の目録にしていたとき（書き換えない）は、その目録をそのまま返す
    if (($result.Status -eq "Ok" -and $written) -or $result.Status -eq "Newer") {
        return $result
    }
    # ロックを取れない・書けないときは、書かずに求めた表を返す（次の機会に直す）
    return @{ Status = "Ok"; Indexes = $computed }
}


function renameOldIndexFolder {
    # 古い置き場（content_index\）にある、同じ名前のインデックスのフォルダの名前も変える（インデックスの改名）。無ければ何もしない。
    # 新しい名前のフォルダが既にあるときは動かさない（古い置き場のものは作り直しで消えるため、改名を止めない）
    param (
        [string]$oldName,
        [string]$newName,
        $workspace = $script:workspace
    )

    $from = Join-Path $workspace.OldIndexDir $oldName
    $to = Join-Path $workspace.OldIndexDir $newName
    if (![System.IO.Directory]::Exists((toLongPath $from))) {
        return
    }
    if ([string]::Equals($oldName, $newName, [System.StringComparison]::OrdinalIgnoreCase)) {
        $tmp = Join-Path $workspace.OldIndexDir "${oldName}_rename_$([System.Diagnostics.Process]::GetCurrentProcess().Id)"
        [System.IO.Directory]::Move((toLongPath $from), (toLongPath $tmp))
        [System.IO.Directory]::Move((toLongPath $tmp), (toLongPath $to))
    } elseif (![System.IO.Directory]::Exists((toLongPath $to))) {
        [System.IO.Directory]::Move((toLongPath $from), (toLongPath $to))
    }
}


function getWorkspaceIndexFormat {
    # ワークスペースと各インデックスの状態（getIndexFormatState）を、今のファイルから求める。
    # 目録が壊れていれば、先に本文インデックスの版の行から書き直す
    param (
        $workspace = $script:workspace
    )

    $catalog = readIndexFormat $workspace.FormatFile
    if ($catalog.Status -eq "Broken") {
        # 書き直した後の状態は Ok だが、「目録だけ壊れていた」ことを今回の結果に残す（Indexes は書き直したもの）
        $repaired = repairIndexFormat $workspace
        $catalog = if ($repaired.Status -eq "Newer") { $repaired } else { @{ Status = "Broken"; Indexes = $repaired.Indexes } }
    }
    $legacy = getLegacyIndexState $workspace.Dir
    return (getIndexFormatState $catalog (findOldIndexNames $workspace) (getIndexNamesWithContent $workspace.IndexDir) $legacy.HasLegacyIndex)
}


function removeOldIndexOf {
    # 作り終えたインデックス name の、古い置き場の分（OldIndexDir\<name>）を消す。古い置き場に本文インデックスのあるフォルダが
    # 無くなれば、${oldIndexEntries} の古い置き場そのもの（content_index\・ingest_status.tsv）も消す。
    # 消せなければ残して結果で返し、例外にしない（古い版の PC が開いているときなど。次の作成の終わりにもう一度消す）。
    # 戻り値: @{ Removed（消したものの数）; Failed（消せなかったものの説明の配列） }
    param (
        [string]$name,
        $workspace = $script:workspace
    )

    $removed = 0
    $failed = New-Object System.Collections.Generic.List[string]
    # 名前は設定や取り込み一覧から来る。置き場の外を指す名前（.. や区切り文字を含むもの）では何も消さない
    if ($name -eq "" -or $name -eq "." -or $name -eq ".." -or $name.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        return @{ Removed = 0; Failed = @("インデックスの名前が正しくないため、古い形式の分を消していません: $name") }
    }
    $oldDir = Join-Path $workspace.OldIndexDir $name
    if ([System.IO.Directory]::Exists((toLongPath $oldDir))) {
        try {
            removeDirectoryRetry $oldDir
            $removed++
        } catch {
            $failed.Add("${oldDir}（$($_.Exception.Message)）")
        }
    }
    $remainingOld = getIndexNamesWithContent $workspace.OldIndexDir
    if ($failed.Count -eq 0 -and @($remainingOld).Count -eq 0) {
        foreach ($property in ${oldIndexEntries}) {
            $entry = [string]$workspace.$property
            try {
                if ([System.IO.Directory]::Exists((toLongPath $entry))) {
                    removeDirectoryRetry $entry
                    $removed++
                } elseif ([System.IO.File]::Exists((toLongPath $entry))) {
                    [System.IO.File]::Delete((toLongPath $entry))
                    $removed++
                }
            } catch {
                $failed.Add("${entry}（$($_.Exception.Message)）")
            }
        }
    }
    return @{ Removed = $removed; Failed = $failed.ToArray() }
}


function markIndexFormatBuilding {
    # この回で作るインデックスのうち、目録に行が無いものを「作成中」で足す（新しい場所の初回作成。行のあるものは変えない）。
    # 書き換えられなければ（今より新しい版の目録・ロックを取れない）$false
    param (
        [string[]]$names,
        $workspace = $script:workspace
    )

    $targetNames = @($names)
    $building = ${catalogRowBuilding}
    return (updateIndexFormat {
        param ($indexes)
        foreach ($name in $targetNames) {
            if (!$indexes.Contains($name)) { $indexes[$name] = $building }
        }
    }.GetNewClosure() $workspace.FormatFile)
}


function completeIndexFormat {
    # 作り終えたインデックス names の行を「作り終えた」にし、古い置き場の分を消す（removeOldIndexOf）。
    # 戻り値: @{ Updated（目録を書けたか）; Failed（消せなかったものの説明の配列） }
    param (
        [string[]]$names,
        $workspace = $script:workspace
    )

    $targetNames = @($names)
    $finished = ${catalogRowFinished}
    $updated = updateIndexFormat {
        param ($indexes)
        foreach ($name in $targetNames) { $indexes[$name] = $finished }
    }.GetNewClosure() $workspace.FormatFile
    $failed = New-Object System.Collections.Generic.List[string]
    if ($updated) {
        foreach ($name in $targetNames) {
            $result = removeOldIndexOf $name $workspace
            $failed.AddRange([string[]]@($result.Failed))
        }
    }
    return @{ Updated = $updated; Failed = $failed.ToArray() }
}
