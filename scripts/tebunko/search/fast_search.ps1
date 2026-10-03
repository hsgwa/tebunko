# 高速検索: Windows Search で検索語を含みうるフォルダを先に絞り、照合する集約ファイルを集める（状態層）。
# 集めた集約ファイルは、今までどおり searchPackIndex で照合する。結果（行・行番号・順番）は、すべてを照合したときと同じになる:
#   反映済みの システムインデックスには、そのフォルダの集約ファイルのすべての 2-gram が入っていて、検索語の語はその一部のため、
#   当たる集約ファイルのフォルダは必ず候補に入る。反映済みでない・対象外・対応済みでないものは、候補に関係なく照合する。

function getReflectedSystemIndexEntries {
    # 状態ファイルの反映待ちの行を、Windows Search の索引と照らして、反映済みになったものとまだのものに分ける。
    # @{ Reflected（反映済みになった行の相対パス）; Unreflected（まだ反映されていない行があるフォルダの相対パス） } を返す。
    # 反映待ちが無ければ、Windows Search には問い合わせない。問い合わせに失敗したら例外のまま投げる
    #   ask: { param($sql) 行（object[]）の一覧 }（呼び出し元の $query・$connection を見る。& で呼ぶと呼び出し元の変数が見える）
    param (
        $state,
        [string]$systemRootPath,
        [scriptblock]$ask
    )

    $reflected = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $unreflected = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    if ($state.Pending.Count -gt 0) {
        foreach ($row in (& $ask (newSystemIndexStateQuery $systemRootPath))) {
            $rel = getRelativePath (convertItemUrl ([string]$row[0])) $systemRootPath
            if ($rel -and $state.Pending.ContainsKey($rel) -and (testSystemIndexReflected $row[1] $row[2] $state.Pending[$rel])) {
                [void]$reflected.Add($rel)
            }
        }
        foreach ($rel in $state.Pending.Keys) {
            if (!$reflected.Contains($rel)) {
                [void]$unreflected.Add([System.IO.Path]::GetDirectoryName($rel))
            }
        }
    }
    return @{ Reflected = $reflected; Unreflected = $unreflected }
}

function getFastSearchPackFiles {
    # 高速検索で照合する集約ファイルを、getIndexPackFiles と同じ形（@{ Folders; Packs }）に Fast（@{ Candidates; Unreflected }）を足して返す。
    # 高速検索を使えないとき（検索語から語を作れない・状態ファイルを読めない・Windows Search に問い合わせられない）は $null
    # （呼び出し側が getIndexPackFiles ですべての集約ファイルを集める）。
    #   folders: getIndexPackFiles と同じ（文字列、または @{ Root; RelPath; Recurse }）
    #   query  : { param($sql) 行（object[]）の一覧 }。$null なら Windows Search を開いて問い合わせる（テストで差し替える）
    param (
        [string]$word,
        [object[]]$folders,
        [scriptblock]$query = $null,
        [string]$indexRoot = $workspace.IndexDir,
        [string]$systemRoot = $workspace.SystemIndexDir,
        [string]$statePath = $workspace.SystemIndexStateFile,
        [scriptblock]$onProgress = $null
    )

    $grams = getSearchGrams $word
    if ($grams.Count -eq 0) {
        return $null
    }
    $state = readSystemIndexState $statePath
    if ($null -eq $state) {
        return $null
    }
    $connection = $null
    if ($null -eq $query) {
        $connection = openWindowsSearch
        if ($null -eq $connection) {
            return $null
        }
    }
    # 問い合わせ（差し替えが無ければ開いた接続で）。& で呼ぶと、この関数の $query・$connection が見える
    $ask = {
        param ($sql)
        if ($query) { & $query $sql } else { invokeWindowsSearch $connection $sql }
    }

    $indexRootPath = $indexRoot.TrimEnd("\")
    $systemRootPath = $systemRoot.TrimEnd("\")
    $targets = New-Object System.Collections.Generic.List[object]
    $candidates = 0
    try {
        # 反映の判定（状態ファイルに反映待ちがあるものだけ。反映済みになった行は、この後で状態から消す）
        $entries = getReflectedSystemIndexEntries $state $systemRootPath $ask
        $reflected = $entries.Reflected
        $unreflected = $entries.Unreflected

        foreach ($folder in $folders) {
            if ($folder -is [string]) {
                $folder = @{ Root = $folder; RelPath = ""; Recurse = $true }
            }
            $root = ([string]$folder.Root).TrimEnd("\")
            $relPath = ([string]$folder.RelPath).Trim("\")
            $targetDir = if ($relPath) { "$root\$relPath" } else { $root }
            # 別の場所のインデックス・インデックス全体・対応済みでないインデックス・無いフォルダは、今までどおりすべてを列挙する
            if (!$root.Equals($indexRootPath, [System.StringComparison]::OrdinalIgnoreCase) -or $relPath -eq "" -or
                !$state.Covered.Contains((getIndexNameOfRelPath $relPath)) -or ![System.IO.Directory]::Exists((toLongPath $targetDir))) {
                $targets.Add($folder)
                continue
            }
            $found = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
            $scopeDir = "$systemRootPath\$relPath"
            if ([System.IO.Directory]::Exists((toLongPath $scopeDir))) {
                foreach ($row in (& $ask (newSystemIndexQuery $scopeDir $grams))) {
                    $rel = getRelativePath ([System.IO.Path]::GetDirectoryName((convertItemUrl ([string]$row[0])))) $systemRootPath
                    if ($rel) { [void]$found.Add($rel) }
                }
                # 分けた txt は、語ごとに問い合わせ、すべての語がどれかの部分で見つかったフォルダを候補にする
                $split = $null
                foreach ($gram in $grams) {
                    $hit = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
                    foreach ($row in (& $ask (newSystemIndexSplitQuery $scopeDir $gram))) {
                        $rel = getRelativePath ([System.IO.Path]::GetDirectoryName((convertItemUrl ([string]$row[0])))) $systemRootPath
                        if ($rel) { [void]$hit.Add($rel) }
                    }
                    if ($null -eq $split) { $split = $hit } else { $split.IntersectWith($hit) }
                    if ($split.Count -eq 0) { break }
                }
                if ($split) { $found.UnionWith($split) }
            }
            $candidates += $found.Count
            $found.UnionWith($unreflected)
            $found.UnionWith($state.Excluded)
            foreach ($rel in $found) {
                $dir = "$indexRootPath\$rel"
                if ((testFolderInTarget $dir $targetDir ([bool]$folder.Recurse)) -and [System.IO.Directory]::Exists((toLongPath $dir))) {
                    $targets.Add(@{ Root = $indexRootPath; RelPath = $rel; Recurse = $false })
                }
            }
        }
    } catch {
        return $null
    } finally {
        if ($connection) {
            $connection.Dispose()
        }
    }

    $index = getIndexPackFiles $targets.ToArray() $onProgress
    if ($reflected.Count -gt 0) {
        # 反映済みになった行を消す（txt が書き直されて日時が変わった行は残す）。書けなくても検索は続ける。
        # 書き換えの中から見る値は、updateSystemIndexState の変数と名前が重ならないようにする（呼び出し先の $state が見えてしまう）
        $readTicks = $state.Pending
        [void](updateSystemIndexState {
            param ($current)
            foreach ($rel in $reflected) {
                if ($current.Pending.ContainsKey($rel) -and $current.Pending[$rel] -eq $readTicks[$rel]) {
                    [void]$current.Pending.Remove($rel)
                }
            }
        } $statePath)
    }
    $index.Fast = @{ Candidates = $candidates; Unreflected = $unreflected.Count }
    return $index
}

function getSystemIndexProgress {
    # システムインデックス（高速検索用の txt）が Windows Search にどこまで反映されたかを数える。画面が「確かめるとき」だけ呼ぶ（検索では呼ばない）。
    # @{ Folders; Waiting; ContentIndexed; ByIndex } を返す。状態ファイルを読めない・Windows Search に問い合わせられないときは $null。
    #   Folders       : txt があるフォルダと、状態ファイルの反映待ちのフォルダを合わせた数（分けた txt は 1 つのフォルダに数える）
    #   Waiting       : 反映待ちのうち、Windows Search でまだ反映済みでないフォルダの数（getFastSearchPackFiles と同じ判定）
    #   ContentIndexed: 本文インデックス（index の TSV）も Windows Search に索引されているか（反映が遅い理由の案内に使う）
    #   ByIndex       : インデックス名（system_index からの相対パスの先頭のフォルダ名） → @{ Folders; Waiting } の辞書（OrdinalIgnoreCase）
    #   query         : getFastSearchPackFiles と同じ。$null なら Windows Search を開いて問い合わせる（テストで差し替える）
    #   connection    : 開いた接続（渡したら呼び出し側が閉じる）。$null なら query が無いときだけ自分で開いて閉じる
    param (
        [scriptblock]$query = $null,
        [string]$indexRoot = $workspace.IndexDir,
        [string]$systemRoot = $workspace.SystemIndexDir,
        [string]$statePath = $workspace.SystemIndexStateFile,
        $connection = $null
    )

    $state = readSystemIndexState $statePath
    if ($null -eq $state) {
        return $null
    }
    $own = $false
    if ($null -eq $query -and $null -eq $connection) {
        $connection = openWindowsSearch
        if ($null -eq $connection) {
            return $null
        }
        $own = $true
    }
    # 問い合わせ（差し替えが無ければ開いた接続で）。& で呼ぶと、この関数の $query・$connection が見える
    $ask = {
        param ($sql)
        if ($query) { & $query $sql } else { invokeWindowsSearch $connection $sql }
    }

    $systemRootPath = $systemRoot.TrimEnd("\")
    try {
        $entries = getReflectedSystemIndexEntries $state $systemRootPath $ask
        $folders = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        $longRoot = toLongPath $systemRootPath
        if ([System.IO.Directory]::Exists($longRoot)) {
            $pattern = "$([System.IO.Path]::GetFileNameWithoutExtension(${systemIndexFileName}))*.txt"
            foreach ($file in [System.IO.Directory]::EnumerateFiles($longRoot, $pattern, [System.IO.SearchOption]::AllDirectories)) {
                $rel = getRelativePath (fromLongPath ([System.IO.Path]::GetDirectoryName($file))) $systemRootPath
                if ($rel) { [void]$folders.Add($rel) }
            }
        }
        foreach ($key in $state.Pending.Keys) {
            [void]$folders.Add([System.IO.Path]::GetDirectoryName($key))
        }
        $contentIndexed = testTsvIndexedByWindowsSearch $indexRoot $connection
    } catch {
        return $null
    } finally {
        if ($own) {
            $connection.Dispose()
        }
    }
    $byIndex = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($rel in $folders) {
        $name = $rel.Split([char[]]"\")[0]
        if (!$byIndex.ContainsKey($name)) {
            $byIndex[$name] = @{ Folders = 0; Waiting = 0 }
        }
        $byIndex[$name].Folders++
    }
    foreach ($rel in $entries.Unreflected) {
        $name = $rel.Split([char[]]"\")[0]
        if ($byIndex.ContainsKey($name)) {
            $byIndex[$name].Waiting++
        }
    }
    return @{ Folders = $folders.Count; Waiting = $entries.Unreflected.Count; ContentIndexed = $contentIndexed; ByIndex = $byIndex }
}
