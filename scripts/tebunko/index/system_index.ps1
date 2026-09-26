# システムインデックス（system_index の txt）と、その状態ファイル（システムインデックスの状態.tsv）の読み書き（状態層）。
# txt は index の中のフォルダ 1 つにつき 1 つ（分けたときは複数）で、index と同じ相対パスの system_index の中に置く。
# 中身は、そのフォルダ直下の集約ファイル（content.<拡張子>.tsv）と、インデックス作成の途中で残った、直下のブックのフォルダ（<ファイル名.xlsx>）の中の TSV から作る（search_gram.ps1）。
# 集約ファイルのメタ情報の行は除く（getPackContentText）。

function getSystemIndexFolderTsvPaths {
    # index の中のフォルダ 1 つの、システムインデックスの元になるファイル（直下の集約ファイル・TSV と、直下のブックのフォルダの中の TSV。\\?\ 付き）。
    # 検索対象のツリーで「フォルダ直下のファイル」を選んだときと同じ範囲（getIndexPackFiles の Recurse = $false）
    param (
        [string]$folder
    )

    $long = toLongPath $folder
    $paths = New-Object System.Collections.Generic.List[string]
    if (![System.IO.Directory]::Exists($long)) {
        return , $paths.ToArray()
    }
    $paths.AddRange([System.IO.Directory]::GetFiles($long, "*.tsv"))
    foreach ($sub in [System.IO.Directory]::GetDirectories($long)) {
        if (testIndexBookDir $sub) {
            $paths.AddRange([System.IO.Directory]::GetFiles($sub, "*.tsv"))
        }
    }
    return , $paths.ToArray()
}

function writeSystemIndexFolder {
    # index の中のフォルダ 1 つについて、system_index の txt を作り直す。
    # @{ Rel（index からの相対パス）; Files（@{ Rel（system_index からの txt の相対パス）; Ticks（更新日時。UTC の Ticks） } の配列）;
    #    Excluded（パスが長すぎて作らなかった） } を返す。集約ファイル・TSV が無くなったフォルダは txt を消して Files を空で返す
    #   texts: そのフォルダの集約ファイルの中身（インデックス作成で書いたばかりのもの）。渡せばファイルを読み直さない
    param (
        [string]$folder,
        [string]$indexRoot,
        [string]$systemRoot,
        [string[]]$texts = $null
    )

    $rel = $folder.Substring($indexRoot.TrimEnd("\").Length).Trim("\")
    $outDir = "$($systemRoot.TrimEnd('\'))\$rel"
    $longOut = toLongPath $outDir
    $result = @{ Rel = $rel; Files = @(); Excluded = $false }

    # 前の txt（分けた数が変わることもあるため、すべて）を消す
    if ([System.IO.Directory]::Exists($longOut)) {
        foreach ($old in [System.IO.Directory]::GetFiles($longOut, "$([System.IO.Path]::GetFileNameWithoutExtension(${systemIndexFileName}))*.txt")) {
            [System.IO.File]::Delete($old)
        }
    }
    $set = New-Object 'System.Collections.Generic.HashSet[uint32]'
    if ($null -ne $texts) {
        if ($texts.Count -eq 0) {
            return $result
        }
        foreach ($text in $texts) {
            addTextGrams $set (getPackContentText $text)
        }
    } else {
        $tsvPaths = getSystemIndexFolderTsvPaths $folder
        if ($tsvPaths.Count -eq 0) {
            return $result
        }
        foreach ($path in $tsvPaths) {
            $text = [System.IO.File]::ReadAllText($path)
            if ([System.IO.Path]::GetFileName($path) -like ${packFilePattern}) {
                $text = getPackContentText $text
            }
            addTextGrams $set $text
        }
    }
    $values = New-Object 'uint32[]' $set.Count
    $set.CopyTo($values)
    [Array]::Sort($values)
    $parts = getGramPartCount $values.Length
    if (!(testSystemIndexPath $outDir $parts)) {
        $result.Excluded = $true
        return $result
    }

    [System.IO.Directory]::CreateDirectory($longOut) | Out-Null
    $perPart = [int][Math]::Ceiling($values.Length / $parts)
    $names = getSystemIndexFileNames $parts
    $files = New-Object System.Collections.Generic.List[hashtable]
    for ($p = 0; $p -lt $parts; $p++) {
        $start = $p * $perPart
        $text = convertToGramText $values $start ([Math]::Min($perPart, $values.Length - $start))
        $path = "$longOut\$($names[$p])"
        [System.IO.File]::WriteAllText($path, $text, [System.Text.Encoding]::ASCII)
        $files.Add(@{ Rel = "$rel\$($names[$p])"; Ticks = [System.IO.File]::GetLastWriteTimeUtc($path).Ticks })
    }
    $result.Files = $files.ToArray()
    return $result
}

# txt をまとめて作るときのスレッドの数の上限（物理メモリ 8GB の PC で 4 スレッドのときにメモリが足りなくなったことがあるため）
${systemIndexWorkerMax} = 4

function writeSystemIndexFolders {
    # index の中のフォルダの txt を、まとめて作り直す（並列）。writeSystemIndexFolder の結果の配列を返す。
    #   shouldStop: $true を返すと、始めていない分を作らずに止める（作った分だけを返す）
    param (
        [string[]]$folders,
        [string]$indexRoot,
        [string]$systemRoot,
        [int]$workers = 0,
        [scriptblock]$shouldStop = $null
    )

    $results = New-Object System.Collections.Generic.List[hashtable]
    if ($folders.Count -eq 0) {
        return , $results.ToArray()
    }
    if ($workers -le 0) {
        $workers = [Math]::Min([Environment]::ProcessorCount, ${systemIndexWorkerMax})
    }
    $workers = [Math]::Min($workers, $folders.Count)
    if ($workers -le 1) {
        foreach ($folder in $folders) {
            if ($shouldStop -and (& $shouldStop)) {
                break
            }
            $results.Add((writeSystemIndexFolder $folder $indexRoot $systemRoot))
        }
        return , $results.ToArray()
    }

    # 各スレッドには必要な関数・値だけを読み込む（lib.ps1 全体を読み込むと、スレッドを用意するだけで時間がかかるため）。
    # インデックス作成の処理のため、スレッドの優先度を下げる（画面・検索を先に動かす。docs/00_共通_4_プロセスとスレッド.md 7.2）
    $state = newWorkerState @("writeSystemIndexFolder", "getSystemIndexFolderTsvPaths", "addTextGrams", "convertToGramText",
        "getGramPartCount", "getSystemIndexFileNames", "testSystemIndexPath", "toLongPath", "getPackContentText", "testIndexBookDir") `
        @("systemIndexFileName", "systemIndexPartBytes", "systemIndexPathMax", "indexBookDirPattern", "packFilePattern")
    $pool = [WorkerPool]::new($workers, $state, $Host, "BelowNormal")
    $pending = New-Object System.Collections.Generic.Queue[hashtable]
    $jobScript = {
        param ($folder, $indexRoot, $systemRoot)
        # 別スレッドは既定では .NET の例外で止まらず、書けなかった txt を作ったものとして返してしまう。例外で止めて呼び出し元に伝える
        $ErrorActionPreference = "Stop"
        @{ Result = writeSystemIndexFolder $folder $indexRoot $systemRoot }
    }.ToString()
    try {
        $next = 0
        while ($next -lt $folders.Count -or $pending.Count -gt 0) {
            # スレッド数の 2 倍まで先に始めておき、終わった順ではなく始めた順に受け取る
            while ($pending.Count -lt $workers * 2 -and $next -lt $folders.Count) {
                if ($shouldStop -and (& $shouldStop)) {
                    $next = $folders.Count
                    break
                }
                $pending.Enqueue($pool.Submit($jobScript, @($folders[$next], $indexRoot, $systemRoot)))
                $next++
            }
            if ($pending.Count -eq 0) {
                break
            }
            $output = $pool.Receive($pending.Dequeue())
            if ($output.Count -gt 0) {
                $results.Add($output[0].Result)
            }
        }
    } finally {
        foreach ($job in $pending) {
            $pool.Cancel($job)
        }
        $pool.Close()
    }
    return , $results.ToArray()
}

function readSystemIndexState {
    # 状態ファイルを読む（convertFromSystemIndexState の形）。無ければ空。
    # 書き込み中などで読めなければ $null（高速検索を使わず、すべてを照合する）
    param (
        [string]$path = ${systemIndexStateFile}
    )

    if (![System.IO.File]::Exists($path)) {
        return newSystemIndexState
    }
    for ($i = 1; $i -le 10; $i++) {
        try {
            $stream = [System.IO.FileStream]::new($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
            try {
                $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
                return convertFromSystemIndexState ($reader.ReadToEnd() -split "\r?\n")
            } finally {
                $stream.Dispose()
            }
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds 100
        }
    }
    return $null
}

function updateSystemIndexState {
    # 状態ファイルを排他で開き、change（{ param($state) }）で書き換えて保存する。
    # インデクサ（取り込み）と画面（検索のたびの整理）が同時に書かないよう、開いている間はほかから開けない。
    # 開けなければ少し待って数回試し、それでも開けなければ $false を返す（書き換えは次の機会に回る）
    param (
        [scriptblock]$change,
        [string]$path = ${systemIndexStateFile}
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    $stream = $null
    for ($i = 1; $i -le 20 -and $null -eq $stream; $i++) {
        try {
            $stream = [System.IO.FileStream]::new($path, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds 100
        }
    }
    if ($null -eq $stream) {
        return $false
    }
    try {
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true, 4096, $true)
        $state = convertFromSystemIndexState ($reader.ReadToEnd() -split "\r?\n")
        $reader.Dispose()
        & $change $state
        $bytes = ${utf8Bom}.GetPreamble() + ${utf8Bom}.GetBytes(((convertToSystemIndexState $state) -join "`r`n") + "`r`n")
        $stream.SetLength(0)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
    }
    return $true
}

function removeSystemIndexEntries {
    # 状態から、相対パス rel のフォルダ（とその中）の行を消す。indexName を渡したら、そのインデックスの「対応済み」も消す
    param (
        $state,
        [string]$rel,
        [string]$indexName = ""
    )

    foreach ($key in @($state.Pending.Keys)) {
        if ($key.StartsWith("$rel\", [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$state.Pending.Remove($key)
        }
    }
    foreach ($key in @($state.Excluded)) {
        if ($key.Equals($rel, [System.StringComparison]::OrdinalIgnoreCase) -or $key.StartsWith("$rel\", [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$state.Excluded.Remove($key)
        }
    }
    if ($indexName) {
        [void]$state.Covered.Remove($indexName)
    }
}

function setSystemIndexResults {
    # writeSystemIndexFolder の結果を状態に書く（フォルダの前の行は消して、作った txt を反映待ちにする）
    param (
        $state,
        [object[]]$results
    )

    foreach ($result in $results) {
        # フォルダ直下の txt の行だけを消す（サブフォルダの行は、サブフォルダの結果で書き換わる）
        foreach ($key in @($state.Pending.Keys)) {
            if ([System.IO.Path]::GetDirectoryName($key) -eq $result.Rel) {
                [void]$state.Pending.Remove($key)
            }
        }
        [void]$state.Excluded.Remove($result.Rel)
        foreach ($file in $result.Files) {
            $state.Pending[$file.Rel] = $file.Ticks
        }
        if ($result.Excluded) {
            [void]$state.Excluded.Add($result.Rel)
        }
    }
}

function markSystemIndexChanged {
    # TSV を入れ替えたフォルダを「反映待ち（日時は 0）」にする（txt を作り直すまで、そのフォルダは .NET で照合させる）。
    # 0 は txt の更新日時とも Windows Search の DateModified とも一致しないため、前の txt が索引されていても反映済みにならない
    # （今の日時にすると、txt を書いたのと同じ秒の中では一致して、反映済みと取り違える）
    param (
        [string[]]$rels,
        [string]$path = ${systemIndexStateFile}
    )

    return updateSystemIndexState {
        param ($state)
        foreach ($rel in $rels) {
            $state.Pending["$rel\${systemIndexFileName}"] = 0
        }
    } $path
}

function getSystemIndexStaleFolders {
    # txt の作り直しが要る index の中のフォルダ（フルパス）を返す。
    #   txt が無い・TSV より古い・状態が「反映待ち」なのに txt の更新日時と合わない（取り込みの途中で止まった）もの
    param (
        [string]$indexRoot,
        [string]$systemRoot,
        $state
    )

    $stale = New-Object System.Collections.Generic.List[string]
    $longRoot = toLongPath $indexRoot
    if (![System.IO.Directory]::Exists($longRoot)) {
        return , $stale.ToArray()
    }
    $rootLength = $indexRoot.TrimEnd("\").Length
    foreach ($longDir in [System.IO.Directory]::EnumerateDirectories($longRoot, "*", [System.IO.SearchOption]::AllDirectories)) {
        # 元のファイルごとのフォルダ（集約する前の TSV・中身が空のファイル）は、親のフォルダの txt に入る
        if (testIndexBookDir $longDir $false) {
            continue
        }
        $folder = fromLongPath $longDir
        $rel = $folder.Substring($rootLength).Trim("\")
        if ($state.Excluded.Contains($rel)) {
            continue
        }
        $tsvPaths = getSystemIndexFolderTsvPaths $folder
        $txtDir = toLongPath "$($systemRoot.TrimEnd('\'))\$rel"
        $txts = if ([System.IO.Directory]::Exists($txtDir)) { @([System.IO.Directory]::GetFiles($txtDir, "$([System.IO.Path]::GetFileNameWithoutExtension(${systemIndexFileName}))*.txt")) } else { @() }
        if ($tsvPaths.Count -eq 0) {
            if ($txts.Count -gt 0) {
                $stale.Add($folder)   # TSV が無くなった（txt を消す）
            }
            continue
        }
        if ($txts.Count -eq 0) {
            $stale.Add($folder)
            continue
        }
        $oldest = ($txts | ForEach-Object { [System.IO.File]::GetLastWriteTimeUtc($_) } | Measure-Object -Minimum).Minimum
        $newestTsv = ($tsvPaths | ForEach-Object { [System.IO.File]::GetLastWriteTimeUtc($_) } | Measure-Object -Maximum).Maximum
        $marked = $state.Pending.ContainsKey("$rel\${systemIndexFileName}") -and
            $state.Pending["$rel\${systemIndexFileName}"] -ne [System.IO.File]::GetLastWriteTimeUtc("$txtDir\${systemIndexFileName}").Ticks
        if ($newestTsv -gt $oldest -or $marked) {
            $stale.Add($folder)
        }
    }
    return , $stale.ToArray()
}

function updateSystemIndexes {
    # インデックス作成の終わりに、システムインデックスの作り直しが要るフォルダをまとめて作り直し、状態ファイルに書く。
    # すべてのフォルダの txt がそろったインデックスは「対応済み」にする（対応済みでないインデックスは、高速検索でもすべてを照合する）。
    # 利用者の作業の邪魔にならないよう、作るスレッドの優先度を下げる（writeSystemIndexFolders）。shouldStop が $true を返せば、始めていない分は作らない。
    # 作り直したフォルダの数と、作り終えていないインデックスの数を @{ Built; Unfinished } で返す
    param (
        [string]$indexRoot = ${indexDir},
        [string]$systemRoot = ${systemIndexDir},
        [string]$statePath = ${systemIndexStateFile},
        [scriptblock]$shouldStop = $null
    )

    $state = readSystemIndexState $statePath
    if ($null -eq $state) {
        writeIndexerLog "システムインデックスの状態を読めないため、作り直しは次のインデックス作成に回します。" "Yellow"
        return @{ Built = 0; Unfinished = -1 }
    }
    # 無くなったインデックス（設定から外した・画面で削除した）の txt と状態の行を消す
    $names = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    if ([System.IO.Directory]::Exists((toLongPath $indexRoot))) {
        foreach ($dir in [System.IO.Directory]::GetDirectories((toLongPath $indexRoot))) {
            [void]$names.Add([System.IO.Path]::GetFileName($dir))
        }
    }
    $gone = New-Object System.Collections.Generic.List[string]
    if ([System.IO.Directory]::Exists((toLongPath $systemRoot))) {
        foreach ($dir in [System.IO.Directory]::GetDirectories((toLongPath $systemRoot))) {
            $gone.Add([System.IO.Path]::GetFileName($dir))
        }
    }
    $gone.AddRange([string[]]@($state.Covered))
    foreach ($name in ($gone | Sort-Object -Unique)) {
        if (!$names.Contains($name)) {
            [void](removeSystemIndexOf $name $systemRoot $statePath)
            [void]$state.Covered.Remove($name)
        }
    }

    $stale = getSystemIndexStaleFolders $indexRoot $systemRoot $state
    $results = @()
    if ($stale.Count -gt 0) {
        writeIndexerLog "システムインデックス（高速検索用）を作っています…（$($stale.Count) フォルダ）"
        $results = writeSystemIndexFolders $stale $indexRoot $systemRoot 0 $shouldStop

    }
    # 作り終えていないフォルダがあるインデックスは、対応済みにしない
    $built = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($result in $results) {
        [void]$built.Add($result.Rel)
    }
    $unfinished = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $rootLength = $indexRoot.TrimEnd("\").Length
    foreach ($folder in $stale) {
        $rel = $folder.Substring($rootLength).Trim("\")
        if (!$built.Contains($rel)) {
            [void]$unfinished.Add((getIndexNameOfRelPath $rel))
        }
    }
    $saved = updateSystemIndexState {
        param ($current)
        setSystemIndexResults $current $results
        foreach ($name in $names) {
            if ($unfinished.Contains($name)) {
                [void]$current.Covered.Remove($name)
            } else {
                [void]$current.Covered.Add($name)
            }
        }
    } $statePath
    if (!$saved) {
        writeIndexerLog "システムインデックスの状態を書き込めなかったため、次のインデックス作成で作り直します。" "Yellow"
    } elseif ($unfinished.Count -gt 0) {
        writeIndexerLog "中止したため、システムインデックスの一部（$($stale.Count - $built.Count) フォルダ）は次のインデックス作成で作ります。" "Yellow"
    }
    return @{ Built = $built.Count; Unfinished = $unfinished.Count }
}

function removeSystemIndexOf {
    # インデックス（または index の中のフォルダ）を消したとき、system_index の同じフォルダと状態の行を消す
    param (
        [string]$rel,
        [string]$systemRoot = ${systemIndexDir},
        [string]$path = ${systemIndexStateFile}
    )

    $dir = toLongPath "$($systemRoot.TrimEnd('\'))\$rel"
    if ([System.IO.Directory]::Exists($dir)) {
        removeDirectoryRetry "$($systemRoot.TrimEnd('\'))\$rel"
    }
    $name = getIndexNameOfRelPath $rel
    $indexName = if ($name -eq $rel) { $name } else { "" }
    return updateSystemIndexState { param ($state) removeSystemIndexEntries $state $rel $indexName } $path
}
