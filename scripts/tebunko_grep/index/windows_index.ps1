# Windows インデックス（system_index の txt）と、その状態ファイル（Windowsインデックスの状態.tsv）の読み書き（状態層）。
# txt は index の中のフォルダ 1 つにつき 1 つ（分けたときは複数）で、index と同じ相対パスの system_index の中に置く。
# 中身は、そのフォルダ直下の TSV と、直下のブックのフォルダ（<ファイル名.xlsx>）の中の TSV から作る（search_gram.ps1）。

function getWindowsIndexFolderTsvPaths {
    # index の中のフォルダ 1 つの TSV（直下の TSV と、直下のブックのフォルダの中の TSV。\\?\ 付き）。
    # 検索対象のツリーで「フォルダ直下のファイル」を選んだときと同じ範囲（getIndexTsvFiles の Recurse = $false）
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
        if ([System.IO.Path]::GetFileName($sub) -match ${indexBookDirPattern}) {
            $paths.AddRange([System.IO.Directory]::GetFiles($sub, "*.tsv"))
        }
    }
    return , $paths.ToArray()
}

function writeWindowsIndexFolder {
    # index の中のフォルダ 1 つについて、system_index の txt を作り直す。
    # @{ Rel（index からの相対パス）; Files（@{ Rel（system_index からの txt の相対パス）; Ticks（更新日時。UTC の Ticks） } の配列）;
    #    Excluded（パスが長すぎて作らなかった） } を返す。TSV が無くなったフォルダは txt を消して Files を空で返す
    param (
        [string]$folder,
        [string]$indexRoot,
        [string]$systemRoot
    )

    $rel = $folder.Substring($indexRoot.TrimEnd("\").Length).Trim("\")
    $outDir = "$($systemRoot.TrimEnd('\'))\$rel"
    $longOut = toLongPath $outDir
    $result = @{ Rel = $rel; Files = @(); Excluded = $false }

    # 前の txt（分けた数が変わることもあるため、すべて）を消す
    if ([System.IO.Directory]::Exists($longOut)) {
        foreach ($old in [System.IO.Directory]::GetFiles($longOut, "$([System.IO.Path]::GetFileNameWithoutExtension(${windowsIndexFileName}))*.txt")) {
            [System.IO.File]::Delete($old)
        }
    }
    $tsvPaths = getWindowsIndexFolderTsvPaths $folder
    if ($tsvPaths.Count -eq 0) {
        return $result
    }

    $set = New-Object 'System.Collections.Generic.HashSet[uint32]'
    foreach ($path in $tsvPaths) {
        addTextGrams $set ([System.IO.File]::ReadAllText($path))
    }
    $values = New-Object 'uint32[]' $set.Count
    $set.CopyTo($values)
    [Array]::Sort($values)
    $parts = getGramPartCount $values.Length
    if (!(testWindowsIndexPath $outDir $parts)) {
        $result.Excluded = $true
        return $result
    }

    [System.IO.Directory]::CreateDirectory($longOut) | Out-Null
    $perPart = [int][Math]::Ceiling($values.Length / $parts)
    $names = getWindowsIndexFileNames $parts
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
${windowsIndexWorkerMax} = 4

function writeWindowsIndexFolders {
    # index の中のフォルダの txt を、まとめて作り直す（並列）。writeWindowsIndexFolder の結果の配列を返す。
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
        $workers = [Math]::Min([Environment]::ProcessorCount, ${windowsIndexWorkerMax})
    }
    $workers = [Math]::Min($workers, $folders.Count)
    if ($workers -le 1) {
        foreach ($folder in $folders) {
            if ($shouldStop -and (& $shouldStop)) {
                break
            }
            $results.Add((writeWindowsIndexFolder $folder $indexRoot $systemRoot))
        }
        return , $results.ToArray()
    }

    # 各スレッドには必要な関数・値だけを読み込む（lib.ps1 全体を読み込むと、スレッドを用意するだけで時間がかかるため）
    $state = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    foreach ($name in @("writeWindowsIndexFolder", "getWindowsIndexFolderTsvPaths", "addTextGrams", "convertToGramText",
            "getGramPartCount", "getWindowsIndexFileNames", "testWindowsIndexPath", "toLongPath")) {
        $state.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($name, (Get-Command $name -CommandType Function).Definition))
    }
    foreach ($name in @("windowsIndexFileName", "windowsIndexPartBytes", "windowsIndexPathMax", "indexBookDirPattern")) {
        $state.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new($name, (Get-Variable $name -ValueOnly), ""))
    }
    $pool = [runspacefactory]::CreateRunspacePool(1, $workers, $state, $Host)
    $pending = New-Object System.Collections.Generic.Queue[hashtable]
    try {
        $pool.Open()
        $next = 0
        while ($next -lt $folders.Count -or $pending.Count -gt 0) {
            # スレッド数の 2 倍まで先に始めておき、終わった順ではなく始めた順に受け取る
            while ($pending.Count -lt $workers * 2 -and $next -lt $folders.Count) {
                if ($shouldStop -and (& $shouldStop)) {
                    $next = $folders.Count
                    break
                }
                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                [void]$ps.AddScript({
                    param ($folder, $indexRoot, $systemRoot)
                    @{ Result = writeWindowsIndexFolder $folder $indexRoot $systemRoot }
                }).AddArgument($folders[$next]).AddArgument($indexRoot).AddArgument($systemRoot)
                $pending.Enqueue(@{ PowerShell = $ps; Handle = $ps.BeginInvoke() })
                $next++
            }
            if ($pending.Count -eq 0) {
                break
            }
            $job = $pending.Dequeue()
            try {
                $output = $job.PowerShell.EndInvoke($job.Handle)
                if ($output.Count -gt 0) {
                    $results.Add($output[0].Result)
                }
            } finally {
                $job.PowerShell.Dispose()
            }
        }
    } finally {
        foreach ($job in $pending) {
            try { $job.PowerShell.Stop() } catch {}
            $job.PowerShell.Dispose()
        }
        $pool.Dispose()
    }
    return , $results.ToArray()
}

function readWindowsIndexState {
    # 状態ファイルを読む（convertFromWindowsIndexState の形）。無ければ空。
    # 書き込み中などで読めなければ $null（高速検索を使わず、すべてを照合する）
    param (
        [string]$path = ${windowsIndexStateFile}
    )

    if (![System.IO.File]::Exists($path)) {
        return newWindowsIndexState
    }
    for ($i = 1; $i -le 10; $i++) {
        try {
            $stream = [System.IO.FileStream]::new($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
            try {
                $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
                return convertFromWindowsIndexState ($reader.ReadToEnd() -split "\r?\n")
            } finally {
                $stream.Dispose()
            }
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds 100
        }
    }
    return $null
}

function updateWindowsIndexState {
    # 状態ファイルを排他で開き、change（{ param($state) }）で書き換えて保存する。
    # インデクサ（取り込み）と画面（検索のたびの整理）が同時に書かないよう、開いている間はほかから開けない。
    # 開けなければ少し待って数回試し、それでも開けなければ $false を返す（書き換えは次の機会に回る）
    param (
        [scriptblock]$change,
        [string]$path = ${windowsIndexStateFile}
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
        $state = convertFromWindowsIndexState ($reader.ReadToEnd() -split "\r?\n")
        $reader.Dispose()
        & $change $state
        $bytes = ${utf8Bom}.GetPreamble() + ${utf8Bom}.GetBytes(((convertToWindowsIndexState $state) -join "`r`n") + "`r`n")
        $stream.SetLength(0)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
    }
    return $true
}

function removeWindowsIndexEntries {
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

function setWindowsIndexResults {
    # writeWindowsIndexFolder の結果を状態に書く（フォルダの前の行は消して、作った txt を反映待ちにする）
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

function markWindowsIndexChanged {
    # TSV を入れ替えたフォルダを「反映待ち（日時は 0）」にする（txt を作り直すまで、そのフォルダは .NET で照合させる）。
    # 0 は txt の更新日時とも Windows Search の DateModified とも一致しないため、前の txt が索引されていても反映済みにならない
    # （今の日時にすると、txt を書いたのと同じ秒の中では一致して、反映済みと取り違える）
    param (
        [string[]]$rels,
        [string]$path = ${windowsIndexStateFile}
    )

    return updateWindowsIndexState {
        param ($state)
        foreach ($rel in $rels) {
            $state.Pending["$rel\${windowsIndexFileName}"] = 0
        }
    } $path
}

function getWindowsIndexStaleFolders {
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
        if ([System.IO.Path]::GetFileName($longDir) -match ${indexBookDirPattern}) {
            continue
        }
        $folder = fromLongPath $longDir
        $rel = $folder.Substring($rootLength).Trim("\")
        if ($state.Excluded.Contains($rel)) {
            continue
        }
        $tsvPaths = getWindowsIndexFolderTsvPaths $folder
        $txtDir = toLongPath "$($systemRoot.TrimEnd('\'))\$rel"
        $txts = if ([System.IO.Directory]::Exists($txtDir)) { @([System.IO.Directory]::GetFiles($txtDir, "$([System.IO.Path]::GetFileNameWithoutExtension(${windowsIndexFileName}))*.txt")) } else { @() }
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
        $marked = $state.Pending.ContainsKey("$rel\${windowsIndexFileName}") -and
            $state.Pending["$rel\${windowsIndexFileName}"] -ne [System.IO.File]::GetLastWriteTimeUtc("$txtDir\${windowsIndexFileName}").Ticks
        if ($newestTsv -gt $oldest -or $marked) {
            $stale.Add($folder)
        }
    }
    return , $stale.ToArray()
}

function removeWindowsIndexOf {
    # インデックス（または index の中のフォルダ）を消したとき、system_index の同じフォルダと状態の行を消す
    param (
        [string]$rel,
        [string]$systemRoot = ${systemIndexDir},
        [string]$path = ${windowsIndexStateFile}
    )

    $dir = toLongPath "$($systemRoot.TrimEnd('\'))\$rel"
    if ([System.IO.Directory]::Exists($dir)) {
        removeDirectoryRetry "$($systemRoot.TrimEnd('\'))\$rel"
    }
    $name = getIndexNameOfRelPath $rel
    $indexName = if ($name -eq $rel) { $name } else { "" }
    return updateWindowsIndexState { param ($state) removeWindowsIndexEntries $state $rel $indexName } $path
}
