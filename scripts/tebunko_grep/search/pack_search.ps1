# 検索用のまとめファイル（pack_format.ps1）を検索する。
# まとめファイルの全文に 1 回照合し、一致しないファイルは飛ばす。一致したら、位置から場所と行を求める。
# 照合のしかた（lines・filter・scan。search_query.ps1 の getRegexScanMode）で、全文への照合と 1 行ずつの照合の結果を同じにする。

function searchPackFiles {
    # packs の start から count 件のまとめファイルを読み、regex に一致する行を PSCustomObject で返す（1 行に複数一致しても 1 件）。
    # max 以上（max+1 件目）が見つかった時点で打ち切る（負は上限なし）。読めないファイルは飛ばす。
    #   include / exclude: 元のファイル名の条件（$null は条件なし） / excludePlace: 除く場所（図形・コメント）
    #   cache: newTsvTextCache。更新日時・大きさが同じなら、前に読んだ内容と場所の一覧を使う
    # 照合が時間切れ（RegexMatchTimeoutException）なら、そのまま呼び出し元へ伝える
    param (
        $packs,
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
    $end = [Math]::Min($packs.Count, $start + $count)
    $lf = [char]10
    $mark = [char]0x1E
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $mode = if ($null -eq $textRegex) { "scan" } else { $scanMode }
    for ($i = $start; $i -lt $end; $i++) {
        $pack = $packs[$i]
        $text = $null
        $places = $null
        $cached = $false
        if ($null -ne $cache) {
            $entry = $null
            if ($cache.Texts.TryGetValue($pack.Path, [ref]$entry) -and $entry[0] -eq $pack.Ticks -and $entry[1] -eq $pack.Size) {
                $text = $entry[2]
                $places = $entry[3]
                $cached = $true
            }
        }
        if ($null -eq $text) {
            try {
                $stream = [System.IO.FileStream]::new($pack.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
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
        $packMode = $mode
        $matched = $true
        if ($packMode -ne "scan") {
            try {
                $matched = $textRegex.IsMatch($text)
            } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                $packMode = "scan"
            }
        }
        if ($null -eq $places -and ($matched -or $null -ne $cache)) {
            $places = readPackPlaces $text
        }
        if ($null -ne $cache -and !$cached) {
            [System.Threading.Monitor]::Enter($cache)
            try {
                $old = $null
                if ($cache.Texts.TryRemove($pack.Path, [ref]$old)) {
                    $cache.Chars[0] -= $old[2].Length
                }
                if ($cache.Chars[0] + $text.Length -le $cache.MaxChars) {
                    $cache.Texts[$pack.Path] = [object[]]@($pack.Ticks, $pack.Size, $text, $places)
                    $cache.Chars[0] += $text.Length
                }
            } finally {
                [System.Threading.Monitor]::Exit($cache)
            }
        }
        if (!$matched) { continue }

        $before = $hits.Count
        $lineMode = $packMode -eq "lines"
        if ($lineMode) {
            try {
                $length = $text.Length
                $p = 0            # 今の場所（places の番号）
                $number = 0       # 今の場所で、pos までに数えた行の数
                $pos = -1         # 行を数え終えた位置（-1 は場所が変わったので数え直し）
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
                    $hits.Add([pscustomobject]@{
                        Root = $pack.Root; RelPath = $pack.RelPath; RelDir = $pack.RelDir; FileName = $place.Book
                        Book = $place.Book; Location = $place.Location; LineNumber = $number; Line = $text.Substring($lineStart, $lineEnd - $lineStart)
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
            $number = 0
            $pos = $place.Start
            while ($pos -lt $place.End) {
                $n = $text.IndexOf($lf, $pos, $place.End - $pos)
                if ($n -lt 0) { $n = $place.End }
                $number++
                $line = $text.Substring($pos, $n - $pos)
                $pos = $n + 1
                if (!$regex.IsMatch($line)) { continue }
                $hits.Add([pscustomobject]@{
                    Root = $pack.Root; RelPath = $pack.RelPath; RelDir = $pack.RelDir; FileName = $place.Book
                    Book = $place.Book; Location = $place.Location; LineNumber = $number; Line = $line
                })
                if ($max -ge 0 -and $hits.Count -gt $max) { return , $hits }
            }
        }
    }
    return , $hits
}


# まとめファイルの検索で、1 つのスレッドにまとめて渡す大きさの目安（バイト）
${packTaskBytes} = 16MB

# まとめファイルの検索の各スレッドで動かすスクリプト。@{ Hits; Timeout } を返す
${packWorkerScript} = {
    param ($packs, $start, $count, $regex, $max, $textRegex, $scanMode, $cache, $include, $exclude, $excludePlace)
    try {
        $hits = searchPackFiles $packs $start $count $regex $max $textRegex $scanMode $cache $include $exclude $excludePlace
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


function newPackWorkerPool {
    # まとめファイルの検索のスレッドを用意する。各スレッドには検索に要る関数と値だけを読み込む
    param (
        [int]$workers
    )

    $state = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    foreach ($name in "searchPackFiles", "readPackPlaces", "convertPackMetaToPlace", "decodePackValue") {
        $state.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($name, (Get-Item "function:$name").ScriptBlock.ToString()))
    }
    foreach ($name in "packMark", "packVersion", "packPlaceKeys", "placeKindShape", "placeKindComment") {
        $state.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new($name, (Get-Variable -Name $name -ValueOnly), ""))
    }
    $pool = [runspacefactory]::CreateRunspacePool(1, $workers, $state, $Host)
    $pool.Open()
    return $pool
}


function splitPackTasks {
    # まとめファイルの並びを、1 つのスレッドに渡す単位（@{ Start; Count }）に分ける（合計がおよそ packTaskBytes になるまでまとめる）
    param (
        $packs,
        [long]$taskBytes = ${packTaskBytes}
    )

    $tasks = New-Object System.Collections.Generic.List[hashtable]
    $start = 0
    $bytes = 0L
    for ($i = 0; $i -lt $packs.Count; $i++) {
        $bytes += [long]$packs[$i].Size
        if ($bytes -ge $taskBytes -or $i -eq $packs.Count - 1) {
            $tasks.Add(@{ Start = $start; Count = $i - $start + 1 })
            $start = $i + 1
            $bytes = 0L
        }
    }
    return , $tasks
}


function searchPackIndex {
    # まとめファイルをワードで検索し、ヒットした行を返す（画面の検索処理）。
    #   packs        : getPackFiles・getIndexPackFiles の結果
    #   simpleMatch  : $true なら文字どおりに検索する。$false なら正規表現として検索し、正規表現として不正なら文字どおりに検索する
    #   limit        : 件数の上限（0 は上限なし）。超えたら打ち切る
    #   shouldStop   : $true を返すと中止する
    #   caseSensitive: 英字の大文字・小文字を区別する（newSearchRegex）
    #   fileFilter   : 対象ファイル（newFileFilter）。元のファイル名が一致しないものは検索しない
    #   workerCount  : 並列に検索するスレッドの数（0 は CPU のコア数から決める。最大 4）
    #   cache        : 読んだまとめファイルの内容を次の検索で使い回す入れ物（newTsvTextCache。$null は使い回さない）
    #   includeShapes / includeComments: 図形・コメントの場所（"<シート名>[図形]" 等）も検索する（newPlaceExclude）
    #   taskBytes    : 1 つのスレッドにまとめて渡す大きさの目安（バイト）
    #   onProgress   : 1 つの作業を照合するたびに呼ぶ { param($done, $total, $newHits) }（done・total はまとめファイルの数）
    # @{ Hits; SimpleMatch（実際に文字どおり検索したか）; Total（まとめファイルの数）; Truncated; Cancelled } を返す。
    # Hits の各要素は PSCustomObject（Root; RelPath（まとめファイル）; RelDir; FileName; Book; Location; LineNumber; Line）
    param (
        [string]$word,
        $packs,
        [bool]$simpleMatch = $false,
        [int]$limit = 0,
        [scriptblock]$shouldStop = $null,
        [bool]$caseSensitive = $false,
        [string]$fileFilter = "",
        [int]$workerCount = 0,
        $cache = $null,
        [bool]$includeShapes = $true,
        [bool]$includeComments = $true,
        [long]$taskBytes = ${packTaskBytes},
        [scriptblock]$onProgress = $null
    )

    $search = newSearchRegex $word $simpleMatch $caseSensitive
    $filter = newFileFilter $fileFilter
    $excludePlace = newPlaceExclude $includeShapes $includeComments
    $hits = New-Object System.Collections.Generic.List[psobject]
    $result = @{ Hits = $hits; SimpleMatch = $search.SimpleMatch; Total = $packs.Count; Truncated = $false; Cancelled = $false }
    $timeoutMessage = "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"
    $tasks = splitPackTasks $packs $taskBytes
    $workers = if ($workerCount -gt 0) { $workerCount } else { [Math]::Min([Environment]::ProcessorCount, 4) }
    if ($tasks.Count -lt 2) { $workers = 1 }

    $pool = $null
    $pending = New-Object System.Collections.Generic.Queue[hashtable]
    try {
        if ($workers -gt 1) { $pool = newPackWorkerPool $workers }
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
                    $ps = [powershell]::Create()
                    $ps.RunspacePool = $pool
                    [void]$ps.AddScript(${packWorkerScript}).AddArgument($packs).AddArgument($task.Start).AddArgument($task.Count).AddArgument($search.Regex).AddArgument($max).AddArgument($search.TextRegex).AddArgument($search.ScanMode).AddArgument($cache).AddArgument($filter.Include).AddArgument($filter.Exclude).AddArgument($excludePlace)
                    $pending.Enqueue(@{ PowerShell = $ps; Handle = $ps.BeginInvoke(); Done = $task.Start + $task.Count })
                    $next++
                }
                $job = $pending.Dequeue()
                try { $output = $job.PowerShell.EndInvoke($job.Handle) } finally { $job.PowerShell.Dispose() }
                if ($output[0].Timeout) { throw $timeoutMessage }
                $newHits = $output[0].Hits
                $done = $job.Done
            } else {
                $task = $tasks[$next]
                $next++
                $done = $task.Start + $task.Count
                $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
                try {
                    $newHits = searchPackFiles $packs $task.Start $task.Count $search.Regex $max $search.TextRegex $search.ScanMode $cache $filter.Include $filter.Exclude $excludePlace
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
                & $onProgress $done $packs.Count $newHits
            }
            if ($result.Truncated) { break }
        }
    } finally {
        foreach ($job in $pending) {
            try { $job.PowerShell.Stop() } catch {}
            $job.PowerShell.Dispose()
        }
        if ($pool) { $pool.Dispose() }
    }
    return $result
}


function getIndexPackFiles {
    # 検索対象のまとめファイルを集め、@{ Folders; Packs } を返す。
    #   folders: 検索対象インデックスのフォルダ（文字列。フォルダ以下すべて）、または
    #            @{ Root（インデックスのフォルダ）; RelPath（その中のフォルダ。空は Root 自身）; Recurse（$false は直下のファイルだけ） }
    #   Folders: フォルダごとの @{ Path; Root（フルパス）; Exists; Count }
    #   Packs  : getPackFiles の結果をつないだもの（入れ子のフォルダを選んでも重複しない）
    #   onProgress: 数えた件数を知らせる { param($count) }
    param (
        [object[]]$folders = @(${indexDir}),
        [scriptblock]$onProgress = $null
    )

    $folderInfo = New-Object System.Collections.Generic.List[object]
    $packs = New-Object System.Collections.Generic.List[hashtable]
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
        $found = getPackFiles $root $relPath ([bool]$target.Recurse)
        $folderInfo.Add(@{ Path = $dir; Root = $root; Exists = $true; Count = $found.Count })
        foreach ($pack in $found) {
            if ($seen.Add($pack.Path)) { $packs.Add($pack) }
        }
        if ($onProgress) { & $onProgress $packs.Count }
    }
    # フォルダの順、フォルダの中は名前の順（getPackFiles と同じ）。検索対象を複数選んだときも順が崩れないよう並べ直す
    $sorted = [hashtable[]]@($packs | Sort-Object @{ Expression = { $_.Root } }, @{ Expression = { $_.RelDir } }, @{ Expression = { [System.IO.Path]::GetFileName($_.RelPath) } })
    return @{ Folders = $folderInfo.ToArray(); Packs = $sorted }
}