# gram 索引（gram_index.ps1・gram_index_store.ps1）で絞った、照合する本文インデックスのファイルを集める（状態層）。
# 集めた本文インデックスのファイルは、今までどおり searchContentIndex で照合する。結果（行・行番号・順番）は、すべてを照合したときと同じになる:
#   検索語を含むファイルには、検索語のすべての 2-gram が必ず入っているため、含まないと分かったファイルだけを外す。
#   索引が無い・読めない・壊れている、ファイルが索引より新しい（更新日時か大きさが違う）・索引に載っていないものは、必ず照合する。

function getGramContentIndexFiles {
    # 照合する本文インデックスのファイルを、getContentIndexFiles と同じ形（@{ Folders; ContentIndexFiles }）に Gram（@{ Total; Candidates }）を足して返す。
    # 絞り込みを使えないとき（正規表現・検索語から 2-gram を作れない・gram 索引の置き場が無い）と、1 つも外せなかったときは $null
    # （呼び出し側が getContentIndexFiles ですべての本文インデックスのファイルを集める）。
    #   simpleMatch: 文字どおりに探すか（正規表現のときは絞れない）
    #   folders    : getContentIndexFiles と同じ（文字列、または @{ Root; RelPath; Recurse }）
    #   stopSignal : Stop を持つもの（検索の要求。$null は止めない）。Stop が立ったら、フォルダの列挙とインデックス名ごとの絞り込みの切れ目で止め、$null を返す（呼び出し側が Stop を見て終える）
    param (
        [string]$word,
        [bool]$simpleMatch,
        [object[]]$folders,
        [string]$indexRoot = $workspace.IndexDir,
        [string]$gramRoot = $workspace.GramDir,
        [scriptblock]$onProgress = $null,
        $stopSignal = $null
    )

    $grams = getSearchGramValues $word
    if (!(testGramIndexUsable $simpleMatch $grams) -or ![System.IO.Directory]::Exists((toLongPath $gramRoot))) {
        return $null
    }
    $index = getContentIndexFiles $folders $onProgress $stopSignal
    $all = $index.ContentIndexFiles
    $indexRootPath = $indexRoot.TrimEnd("\")

    # インデックス名ごとに分ける（別の場所のインデックスは絞らない）
    $groups = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $all) {
        if (!([string]$item.Root).Equals($indexRootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        $separator = ([string]$item.RelPath).IndexOf("\")
        if ($separator -lt 1) {
            continue
        }
        $name = ([string]$item.RelPath).Substring(0, $separator)
        if (!$groups.ContainsKey($name)) {
            $groups[$name] = New-Object 'System.Collections.Generic.List[object]'
        }
        $groups[$name].Add($item)
    }

    $dropped = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $groups.Keys) {
        if ($null -ne $stopSignal -and $stopSignal.Stop) { return $null }
        $segments = New-Object System.Collections.Generic.List[object]
        try {
            $dir = "$($gramRoot.TrimEnd('\'))\$name"
            $toc = readGramToc $dir
            if ($null -eq $toc -or $toc.Segments.Count -eq 0) {
                continue
            }
            $usable = $true
            for ($i = $toc.Segments.Count - 1; $i -ge 0; $i--) {
                $segment = openGramSegment "$dir\$($toc.Segments[$i].Name)" $toc.Segments[$i].Length
                if ($null -eq $segment) {
                    $usable = $false
                    break
                }
                $segments.Add($segment)
            }
            if (!$usable) {
                continue
            }
            foreach ($segment in $segments) {
                $segment.Candidates = getGramCandidates $segment $grams
            }
            $selected = selectGramFiles $groups[$name].ToArray() $segments.ToArray()
            $keptPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($item in $selected.Files) {
                [void]$keptPaths.Add([string]$item.Path)
            }
            foreach ($item in $groups[$name]) {
                if (!$keptPaths.Contains([string]$item.Path)) {
                    [void]$dropped.Add([string]$item.Path)
                }
            }
        } catch {
            # 読めなかったインデックスは絞らない（すべて照合する）
        } finally {
            closeGramSegments $segments.ToArray()
        }
    }

    if ($null -ne $stopSignal -and $stopSignal.Stop) { return $null }
    $files = New-Object System.Collections.Generic.List[hashtable]
    foreach ($item in $all) {
        if (!$dropped.Contains([string]$item.Path)) {
            $files.Add($item)
        }
    }
    if ($dropped.Count -eq 0) {
        # 1 つも外せなかった（索引が無い・古い・どのファイルも語を含みうる）。絞れていないので、$null を返し、呼び出し側が全部のファイルを照合する
        return $null
    }
    $index.ContentIndexFiles = $files.ToArray()
    $index.Gram = @{ Total = $all.Count; Candidates = $files.Count }
    return $index
}
