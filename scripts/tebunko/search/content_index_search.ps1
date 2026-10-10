# 検索用の本文インデックスのファイル（content_index_format.ps1）を検索する。
# 本文インデックスのファイルの全文に 1 回照合し、一致しないファイルは飛ばす。一致したら、位置から場所と行を求める。
# 照合のしかた（lines・filter・scan。search_query.ps1 の getRegexScanMode）で、全文への照合と 1 行ずつの照合の結果を同じにする。
# Excel の図形・コメントの行は「セル番地 + タブ + 文字」の形のため、区切りのタブより後から始まる一致だけをヒットにする（セル番地だけの一致を除く）。

function testCellPrefixedHit {
    # セル番地 + 区切りのタブ + 文字 の行で、区切りのタブ（最初のタブ）より後から始まる一致があるか。文字の中のタブは文字として扱う。
    # 開始位置つきの照合なので、後読みは前の文字（区切りのタブ）も見える。区切りのタブが無い行は一致なし
    param (
        [regex]$regex,
        [string]$line
    )

    $tab = $line.IndexOf([char]9)
    return ($tab -ge 0) -and $regex.IsMatch($line, $tab + 1)
}


function searchContentIndexFiles {
    # contentIndexFiles の start から count 件の本文インデックスのファイルを読み、regex に一致する行を PSCustomObject で返す（1 行に複数一致しても 1 件）。
    # max 以上（max+1 件目）が見つかった時点で打ち切る（負は上限なし）。読めないファイルは飛ばす。
    #   include / exclude: 元のファイル名の条件（$null は条件なし） / excludePlace: 除く場所（図形・コメント）
    #   cache: newTsvTextCache。更新日時・大きさが同じなら、前に読んだ内容と場所の一覧を使う
    # 照合が時間切れ（RegexMatchTimeoutException）なら、そのまま呼び出し元へ伝える
    param (
        $contentIndexFiles,
        [int]$start,
        [int]$count,
        [regex]$regex,
        [int]$max,
        [regex]$textRegex = $null,
        [string]$scanMode = "scan",
        $cache = $null,
        [regex]$include = $null,
        [regex]$exclude = $null,
        [regex]$excludePlace = $null
    )

    # 並列検索の別スレッドでも動くよう、コマンドレットは使わない
    $hits = [System.Collections.Generic.List[psobject]]::new()
    $end = [Math]::Min($contentIndexFiles.Count, $start + $count)
    $lf = [char]10
    $mark = [char]0x1E
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $mode = if ($null -eq $textRegex) { "scan" } else { $scanMode }
    for ($i = $start; $i -lt $end; $i++) {
        $contentIndexFile = $contentIndexFiles[$i]
        $text = $null
        $places = $null
        $cached = $false
        if ($null -ne $cache) {
            $entry = $null
            if ($cache.Texts.TryGetValue($contentIndexFile.Path, [ref]$entry) -and $entry[0] -eq $contentIndexFile.Ticks -and $entry[1] -eq $contentIndexFile.Size) {
                $text = $entry[2]
                $places = $entry[3]
                $cached = $true
                # この検索で使ったことを残す（trimTsvTextCache は、使われていない古いものから追い出す）
                $entry[4] = $cache.Generation[0]
            }
        }
        if ($null -eq $text) {
            try {
                $stream = [System.IO.FileStream]::new($contentIndexFile.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
                $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::Unicode, $true)
                try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
            } catch [System.IO.IOException] {
                continue
            } catch [System.UnauthorizedAccessException] {
                continue
            }
        }
        # 一致しなければ、場所の一覧も作らずに次へ（場所の一覧は、一致したときか、キャッシュに入れるときだけ作る）。
        # 全文への照合が時間切れなら、このファイルは 1 行ずつ照合する
        $contentIndexMode = $mode
        $matched = $true
        if ($contentIndexMode -ne "scan") {
            try {
                $matched = $textRegex.IsMatch($text)
            } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                $contentIndexMode = "scan"
            }
        }
        if ($null -eq $places -and ($matched -or $null -ne $cache)) {
            $places = readContentIndexPlaces $text
        }
        if ($null -ne $cache -and !$cached) {
            [System.Threading.Monitor]::Enter($cache)
            try {
                $old = $null
                if ($cache.Texts.TryRemove($contentIndexFile.Path, [ref]$old)) {
                    $cache.Chars[0] -= $old[2].Length
                }
                if ($cache.Chars[0] + $text.Length -le $cache.MaxChars) {
                    $cache.Texts[$contentIndexFile.Path] = [object[]]@($contentIndexFile.Ticks, $contentIndexFile.Size, $text, $places, $cache.Generation[0])
                    $cache.Chars[0] += $text.Length
                }
            } finally {
                [System.Threading.Monitor]::Exit($cache)
            }
        }
        if (!$matched) { continue }

        $before = $hits.Count
        $lineMode = $contentIndexMode -eq "lines"
        if ($lineMode) {
            try {
                $length = $text.Length
                $p = 0            # 今の場所（places の番号）
                $number = 0       # 今の場所で、pos までに数えた行の数
                $pos = -1         # 行を数え終えた位置（-1 は場所が変わったので数え直し）
                $lastP = -1       # 直前に判定した場所の番号（isTextBook を場所が変わったときだけ判定し直す）
                $isTextBook = $false
                $m = $textRegex.Match($text)
                while ($m.Success) {
                    $index = $m.Index
                    if ($index -ge $length) { break }
                    $lineStart = if ($index -eq 0) { 0 } else { $text.LastIndexOf($lf, $index - 1) + 1 }
                    $lineEnd = $text.IndexOf($lf, $index)
                    if ($lineEnd -lt 0) { $lineEnd = $length }
                    if ($lineStart -lt $length -and $text[$lineStart] -eq $mark) {
                        # メタ情報の行（ファイル名・場所の名前）の中の一致は捨てる
                        if ($lineEnd -ge $length) { break }
                        $m = $textRegex.Match($text, $lineEnd + 1)
                        continue
                    }
                    while ($p -lt $places.Count -and $places[$p].End -le $lineStart) { $p++; $pos = -1 }
                    if ($p -ge $places.Count) { break }
                    $place = $places[$p]
                    if ($p -ne $lastP) {
                        $isTextBook = ((getContentIndexFileKind $place.Book) -eq "テキスト")
                        $lastP = $p
                    }
                    $skip = ($include -and !$include.IsMatch($place.Book)) -or ($exclude -and $exclude.IsMatch($place.Book)) -or
                        ($excludePlace -and $excludePlace.IsMatch($place.Location))
                    if ($skip) {
                        if ($place.End -ge $length) { break }
                        $m = $textRegex.Match($text, $place.End)
                        continue
                    }
                    if ($pos -lt 0) { $pos = $place.Start; $number = 1 }
                    # 前に数えた位置から行の先頭までの改行を数える（Substring を作らない）
                    while ($pos -lt $lineStart) {
                        $n = $text.IndexOf($lf, $pos, $lineStart - $pos)
                        if ($n -lt 0) { break }
                        $number++
                        $pos = $n + 1
                    }
                    $pos = $lineStart
                    $line = $text.Substring($lineStart, $lineEnd - $lineStart)
                    if ($place.CellPrefixed -and !(testCellPrefixedHit $regex $line)) {
                        # セル番地だけに一致した行はヒットにしない（行番号の数え方は崩さない）
                        if ($lineEnd -ge $length) { break }
                        $m = $textRegex.Match($text, $lineEnd + 1)
                        continue
                    }
                    if ($isTextBook) {
                        $line = truncateHitLine $line ($index - $lineStart)
                    }
                    $hits.Add([pscustomobject]@{
                        Root = $contentIndexFile.Root; RelPath = $contentIndexFile.RelPath; RelDir = $contentIndexFile.RelDir; FileName = $place.Book
                        Book = $place.Book; Location = $place.Location; LineNumber = $number; Line = $line
                    })
                    if ($max -ge 0 -and $hits.Count -gt $max) { return , $hits }
                    if ($lineEnd -ge $length) { break }
                    $m = $textRegex.Match($text, $lineEnd + 1)
                }
                continue
            } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                # 全文への照合が時間切れなら、1 行ずつ照合し直す
                if ($hits.Count -gt $before) { $hits.RemoveRange($before, $hits.Count - $before) }
            }
        }
        # filter・scan（と lines の時間切れ）: 場所ごとに 1 行ずつ照合する
        foreach ($place in $places) {
            if ($include -and !$include.IsMatch($place.Book)) { continue }
            if ($exclude -and $exclude.IsMatch($place.Book)) { continue }
            if ($excludePlace -and $excludePlace.IsMatch($place.Location)) { continue }
            $isTextBook = ((getContentIndexFileKind $place.Book) -eq "テキスト")
            $number = 0
            $pos = $place.Start
            while ($pos -lt $place.End) {
                $n = $text.IndexOf($lf, $pos, $place.End - $pos)
                if ($n -lt 0) { $n = $place.End }
                $number++
                $line = $text.Substring($pos, $n - $pos)
                $pos = $n + 1
                $lineMatch = $regex.Match($line)
                if (!$lineMatch.Success) { continue }
                if ($place.CellPrefixed -and !(testCellPrefixedHit $regex $line)) { continue }
                if ($isTextBook) {
                    $line = truncateHitLine $line $lineMatch.Index
                }
                $hits.Add([pscustomobject]@{
                    Root = $contentIndexFile.Root; RelPath = $contentIndexFile.RelPath; RelDir = $contentIndexFile.RelDir; FileName = $place.Book
                    Book = $place.Book; Location = $place.Location; LineNumber = $number; Line = $line
                })
                if ($max -ge 0 -and $hits.Count -gt $max) { return , $hits }
            }
        }
    }
    return , $hits
}


# 本文インデックスのファイルの検索で、1 つのスレッドにまとめて渡す大きさの目安（バイト）
${contentIndexTaskBytes} = 16MB

# 本文インデックスのファイルの検索の各スレッドで動かすスクリプト。@{ Hits; Timeout } を返す
${contentIndexWorkerScript} = {
    param ($contentIndexFiles, $start, $count, $regex, $max, $textRegex, $scanMode, $cache, $include, $exclude, $excludePlace)
    try {
        $hits = searchContentIndexFiles $contentIndexFiles $start $count $regex $max $textRegex $scanMode $cache $include $exclude $excludePlace
        @{ Hits = $hits; Timeout = $false }
    } catch {
        if ($_.Exception -is [System.Text.RegularExpressions.RegexMatchTimeoutException] -or
            $_.Exception.InnerException -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) {
            @{ Hits = $null; Timeout = $true }
        } else {
            throw
        }
    }
}


function newContentIndexWorkerPool {
    # 本文インデックスのファイルの照合のプール（WorkerPool）を作る。各スレッドには照合に要る関数と値だけを読み込む。
    # 利用者が結果を待つ処理のため、優先度は下げない（docs/design/structure/threads.md「スレッドの一覧」）
    param (
        [int]$workers = (getWorkerCount)
    )

    $state = newWorkerState @("searchContentIndexFiles", "testCellPrefixedHit", "readContentIndexPlaces", "testContentIndexCellPrefixed", "convertContentIndexMetaToPlace", "decodeContentIndexValue", "getContentIndexFileKind", "testTextExtension", "truncateHitLine") `
        @("contentIndexMark", "contentIndexVersion", "contentIndexPlaceKeys", "placeKindShape", "placeKindComment", "placeKindHeaderFooter", "placeKindEmbed", "textExtensions", "hitLineMaxChars", "hitLineBeforeMatchChars")
    return [WorkerPool]::new($workers, $state, $Host, "Normal")
}


function splitContentIndexTasks {
    # 本文インデックスのファイルの並びを、1 つのスレッドに渡す単位（@{ Start; Count }）に分ける（合計がおよそ contentIndexTaskBytes になるまでまとめる）
    param (
        $contentIndexFiles,
        [long]$taskBytes = ${contentIndexTaskBytes}
    )

    $tasks = New-Object System.Collections.Generic.List[hashtable]
    $start = 0
    $bytes = 0L
    for ($i = 0; $i -lt $contentIndexFiles.Count; $i++) {
        $bytes += [long]$contentIndexFiles[$i].Size
        if ($bytes -ge $taskBytes -or $i -eq $contentIndexFiles.Count - 1) {
            $tasks.Add(@{ Start = $start; Count = $i - $start + 1 })
            $start = $i + 1
            $bytes = 0L
        }
    }
    return , $tasks
}


function searchContentIndex {
    # 本文インデックスのファイルをワードで検索し、ヒットした行を返す（画面の検索処理）。
    #   contentIndexFiles        : findContentIndexFiles・getContentIndexFiles の結果
    #   simpleMatch  : $true なら文字どおりに検索する。$false なら正規表現として検索し、正規表現として不正なら文字どおりに検索する
    #   limit        : 件数の上限（0 は上限なし）。超えたら打ち切る
    #   shouldStop   : $true を返すと中止する
    #   caseSensitive: 英字の大文字・小文字を区別する（newSearchRegex）
    #   fileFilter   : 対象ファイル（newFileFilter）。元のファイル名が一致しないものは検索しない
    #   workerCount  : 並列に検索するスレッドの数（0 は CPU のコア数から決める。最大 4）
    #   cache        : 読んだ本文インデックスのファイルの内容を次の検索で使い回す入れ物（newTsvTextCache。$null は使い回さない）
    #   includeShapes / includeComments: 図形・コメントの場所（"<シート名>[図形]" 等）も検索する（newPlaceExclude）
    #   taskBytes    : 1 つのスレッドにまとめて渡す大きさの目安（バイト）
    #   onProgress   : 1 つの作業を照合するたびに呼ぶ { param($done, $total, $newHits) }（done・total は本文インデックスのファイルの数）
    #   pool         : 照合に使うプール（newContentIndexWorkerPool。検索の司令が使い回す）。$null なら、並列にするときだけ作って最後に閉じる
    # @{ Hits; SimpleMatch（実際に文字どおり検索したか）; Total（本文インデックスのファイルの数）; Truncated; Cancelled } を返す。
    # Hits の各要素は PSCustomObject（Root; RelPath（本文インデックスのファイル）; RelDir; FileName; Book; Location; LineNumber; Line）
    param (
        [string]$word,
        $contentIndexFiles,
        [bool]$simpleMatch = $false,
        [int]$limit = 0,
        [scriptblock]$shouldStop = $null,
        [bool]$caseSensitive = $false,
        [string]$fileFilter = "",
        [int]$workerCount = 0,
        $cache = $null,
        [bool]$includeShapes = $true,
        [bool]$includeComments = $true,
        [long]$taskBytes = ${contentIndexTaskBytes},
        [scriptblock]$onProgress = $null,
        $pool = $null
    )

    $search = newSearchRegex $word $simpleMatch $caseSensitive
    $filter = newFileFilter $fileFilter
    $excludePlace = newPlaceExclude $includeShapes $includeComments
    $hits = New-Object System.Collections.Generic.List[psobject]
    $result = @{ Hits = $hits; SimpleMatch = $search.SimpleMatch; Total = $contentIndexFiles.Count; Truncated = $false; Cancelled = $false }
    $timeoutMessage = "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"
    $tasks = splitContentIndexTasks $contentIndexFiles $taskBytes
    $workers = if ($pool) { $pool.Size } elseif ($workerCount -gt 0) { $workerCount } else { getWorkerCount }
    if ($tasks.Count -lt 2) { $workers = 1 }

    $ownPool = $null
    $pending = New-Object System.Collections.Generic.Queue[hashtable]
    try {
        if ($workers -le 1) {
            $pool = $null
        } elseif ($null -eq $pool) {
            $ownPool = newContentIndexWorkerPool $workers
            $pool = $ownPool
        }
        $next = 0
        while ($next -lt $tasks.Count -or $pending.Count -gt 0) {
            if ($shouldStop -and (& $shouldStop)) {
                $result.Cancelled = $true
                break
            }
            if ($pool) {
                while ($pending.Count -lt $workers * 2 -and $next -lt $tasks.Count) {
                    $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
                    $task = $tasks[$next]
                    $job = $pool.Submit(${contentIndexWorkerScript}.ToString(), @($contentIndexFiles, $task.Start, $task.Count, $search.Regex, $max, $search.TextRegex,
                            $search.ScanMode, $cache, $filter.Include, $filter.Exclude, $excludePlace))
                    $job.Done = $task.Start + $task.Count
                    $pending.Enqueue($job)
                    $next++
                }
                $output = $pool.Receive($pending.Dequeue())
                if ($output[0].Timeout) { throw $timeoutMessage }
                $newHits = $output[0].Hits
                $done = $job.Done
            } else {
                $task = $tasks[$next]
                $next++
                $done = $task.Start + $task.Count
                $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
                try {
                    $newHits = searchContentIndexFiles $contentIndexFiles $task.Start $task.Count $search.Regex $max $search.TextRegex $search.ScanMode $cache $filter.Include $filter.Exclude $excludePlace
                } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                    throw $timeoutMessage
                }
            }
            $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
            if ($max -ge 0 -and $newHits.Count -gt $max) {
                $newHits.RemoveRange($max, $newHits.Count - $max)
                $result.Truncated = $true
            }
            $hits.AddRange($newHits)
            if ($onProgress) {
                & $onProgress $done $contentIndexFiles.Count $newHits
            }
            if ($result.Truncated) { break }
        }
    } finally {
        foreach ($job in $pending) {
            $pool.Cancel($job)
        }
        if ($ownPool) { $ownPool.Close() }
    }
    return $result
}


function getContentIndexFiles {
    # 検索対象の本文インデックスのファイルを集め、@{ Folders; ContentIndexFiles } を返す。
    #   folders: 検索対象インデックスのフォルダ（文字列。フォルダ以下すべて）、または
    #            @{ Root（インデックスのフォルダ）; RelPath（その中のフォルダ。空は Root 自身）; Recurse（$false は直下のファイルだけ） }
    #   Folders: フォルダごとの @{ Path; Root（フルパス）; Exists; Count }
    #   ContentIndexFiles  : findContentIndexFiles の結果をつないだもの（入れ子のフォルダを選んでも重複しない）
    #   onProgress: 数えた件数を知らせる { param($count) }
    param (
        [object[]]$folders = @($workspace.IndexDir),
        [scriptblock]$onProgress = $null
    )

    $folderInfo = New-Object System.Collections.Generic.List[object]
    $contentIndexFiles = New-Object System.Collections.Generic.List[hashtable]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($target in $folders) {
        if ($target -is [string]) {
            $target = @{ Root = $target; RelPath = ""; Recurse = $true }
        }
        $relPath = ([string]$target.RelPath).Trim("\")
        $dir = if ($relPath) { "$(([string]$target.Root).TrimEnd('\'))\${relPath}" } else { [string]$target.Root }
        if (!(Test-Path -LiteralPath $target.Root -PathType Container) -or ![System.IO.Directory]::Exists((toLongPath $dir))) {
            $folderInfo.Add(@{ Path = $dir; Root = ""; Exists = $false; Count = 0 })
            continue
        }
        $root = (Resolve-Path -LiteralPath $target.Root).ProviderPath.TrimEnd("\")
        $found = findContentIndexFiles $root $relPath ([bool]$target.Recurse)
        $folderInfo.Add(@{ Path = $dir; Root = $root; Exists = $true; Count = $found.Count })
        foreach ($contentIndexFile in $found) {
            if ($seen.Add($contentIndexFile.Path)) { $contentIndexFiles.Add($contentIndexFile) }
        }
        if ($onProgress) { & $onProgress $contentIndexFiles.Count }
    }
    # フォルダの順、フォルダの中は名前の順（findContentIndexFiles と同じ）。検索対象を複数選んだときも順が崩れないよう並べ直す
    $sorted = [hashtable[]]@($contentIndexFiles | Sort-Object @{ Expression = { $_.Root } }, @{ Expression = { $_.RelDir } }, @{ Expression = { [System.IO.Path]::GetFileName($_.RelPath) } })
    return @{ Folders = $folderInfo.ToArray(); ContentIndexFiles = $sorted }
}