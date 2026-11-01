# 検索結果の組み立てと、インデックスの有無・件数（検索そのものは content_index_search.ps1）。

function toResultLine {
    # 検索結果1件を "ファイル名<TAB>場所<TAB>種別<TAB>行番号<TAB>該当行" に整形する。場所・種別は場所ごとの表記（describePlace。画面の見出しの要約と同じ）。
    # Excelに貼り付けたとき、該当行の各セルが元の列の順（5列目 = A列）に並ぶようにする
    param (
        [string]$book,
        [string]$location,
        [int]$lineNumber,
        [string]$line
    )

    if ((getContentIndexFileKind $book) -eq "Excel") {
        # セル内改行を戻す（改行を含むセルは " で囲まれているため、貼り付けても1セルのまま）
        $line = $line.Replace(${cellNewLine}, "`n")
    } else {
        # Word・PowerPointの行（段落、または表の1行をタブ区切りにしたもの）は " で囲まれていない。
        # Excelは " で始まるセルを囲みの " と解釈して後続がずれるため、そのセルだけ " で囲む
        $line = @($line.Split("`t") | ForEach-Object {
            if ($_.StartsWith('"')) { '"' + $_.Replace('"', '""') + '"' } else { $_ }
        }) -join "`t"
    }
    # 場所（シート名）にタブ・改行が入っていると、列・行が分かれてしまうためスペースにする
    # （Excelのシート名はタブ・改行を含められる）
    $described = describePlace $book $location
    $place = [regex]::Replace($described.Place, '[\x00-\x1F]', " ")
    return "${book}`t${place}`t$($described.Kind)`t${lineNumber}`t${line}"
}

function toResultHeader {
    # 検索結果の見出し行 "ファイル名<TAB>場所<TAB>種別<TAB>行<TAB>A<TAB>B…" を返す（列名は columnCount 列分）
    param (
        [int]$columnCount
    )

    $names = New-Object System.Collections.Generic.List[string]
    $names.AddRange([string[]]@("ファイル名", "場所", "種別", "行"))
    for ($i = 1; $i -le $columnCount; $i++) {
        $names.Add((toColumnName $i))
    }
    return ($names -join "`t")
}


# 検索で読んだ本文インデックスのファイルの内容を残しておく量の上限（文字数。1 文字 2 バイトのため約 128MB）
${searchCacheMaxChars} = 64000000

function newTsvTextCache {
    # 検索で読んだ本文インデックスのファイルの内容を、次の検索で使い回すための入れ物を作る（画面が 1 つ持ち、検索のたびに searchContentIndex に渡す）。
    # 2 回目以降の検索ではファイルを開かない。更新日時・サイズが列挙したときと違う本文インデックスのファイル（書き直した等）は読み直す。
    # 並列検索の各スレッドから使うため、中身は ConcurrentDictionary。
    #   Texts: 本文インデックスのファイルの \\?\ 付きのパス → @(更新日時（UTC の Ticks）, サイズ, 内容, 場所の一覧, 最後に使った世代) / Chars: 残している文字数 / MaxChars: 上限
    #   Generation: 今の世代（検索 1 回ごとに trimTsvTextCache が 1 つ進める）
    param (
        [long]$maxChars = ${searchCacheMaxChars}
    )

    return @{
        Texts      = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        Chars      = [long[]]::new(1)
        MaxChars   = $maxChars
        Generation = [long[]]::new(1)
    }
}

# 検索の後、キャッシュがこの割合を超えていたら、空きを作る（次の検索で新しく読んだものが入るように）
${searchCacheKeepRatio} = 0.9

function trimTsvTextCache {
    # 検索 1 回の後に呼ぶ（検索の司令のスレッド）。上限の keepRatio を超えていたら、今の世代で使わなかったものを、
    # 古い世代から追い出す（消した本文インデックスのファイル・検索しなくなったフォルダの分）。今の世代で使ったものは残す
    # （毎回同じ本文インデックスのファイルを順に読むため、使った順だけで追い出すと、上限より大きいインデックスでは何も残らない）。
    # 最後に世代を 1 つ進める。追い出した数を返す
    param (
        $cache,
        [double]$keepRatio = ${searchCacheKeepRatio}
    )

    $removed = 0
    [System.Threading.Monitor]::Enter($cache)
    try {
        $current = $cache.Generation[0]
        $limit = [long]($cache.MaxChars * $keepRatio)
        if ($cache.Chars[0] -gt $limit) {
            $old = @($cache.Texts.GetEnumerator() | Where-Object { $_.Value[4] -lt $current } | Sort-Object { $_.Value[4] })
            foreach ($item in $old) {
                if ($cache.Chars[0] -le $limit) {
                    break
                }
                $entry = $null
                if ($cache.Texts.TryRemove($item.Key, [ref]$entry)) {
                    $cache.Chars[0] -= $entry[2].Length
                    $removed++
                }
            }
        }
        $cache.Generation[0] = $current + 1
    } finally {
        [System.Threading.Monitor]::Exit($cache)
    }
    return $removed
}

# 検索の要求の番号（溢れの置き場所の名前に使う。画面を開いている間、要求ごとに増える）
$script:searchRequestNumber = 0

function newSearchRequest {
    # 検索の要求 1 つを作る（画面が作り、検索の司令のスレッド（SearchService）が invokeSearchRequest で実行する）。
    # 画面と司令のスレッドの両方から読み書きするため、Synchronized の hashtable にする。
    #   画面が書く: 検索条件・Stop（取り消し）
    #   司令が書く: Queue（ヒット）・Done / Total / IndexTotal / Scanned（進み具合）・Folders・GramUsed（gram 索引で絞ったか）・Truncated（上限を超えた。数えている途中でも立つ）/ HitTotal（上限を超えた分も含む件数。数えている途中は今までの件数。中止・総数の時間切れは -1）/ CountTimeout / OverflowDir（溢れの置き場所。書き溜めないときは空）/ OverflowFailed / Cancelled / Error・Skipped（飛ばした本文インデックスのファイルのパス）・Finished
    #   workDir: gram 索引を探すワークスペース（司令のスレッドは lib.ps1 を読み込んだときのワークスペースを覚えているため、画面から渡す）
    param (
        [string]$word,
        [bool]$simpleMatch,
        $folders,
        [int]$limit,
        [hashtable]$option = @{},
        [string]$workDir = $workspace.Dir
    )

    $script:searchRequestNumber++
    return [hashtable]::Synchronized(@{
        Number = $script:searchRequestNumber
        Word = $word; SimpleMatch = $simpleMatch; Folders = $folders; Limit = $limit; WorkDir = $workDir
        CaseSensitive = [bool]$option.CaseSensitive; FileFilter = (newFileKindFilter $option.FileKinds)
        IncludeShapes = ($option.IncludeShapes -ne $false); IncludeComments = ($option.IncludeComments -ne $false)
        GramUsed = $false
        Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        Stop = $false; Finished = $false; Done = 0; Total = -1; IndexTotal = -1; Scanned = 0
        Truncated = $false; HitTotal = 0; CountTimeout = $false; OverflowDir = ""; OverflowFailed = $false; Cancelled = $false; Error = $null; Skipped = @()
    })
}

function invokeSearchRequest {
    # 検索の要求（newSearchRequest）を実行し、ヒットと進み具合を要求に少しずつ入れる。最後に Finished を立てる。
    # 例外は投げずに Error に入れる。始める前に取り消されていたら（次の要求が来た等）、何もせずに Cancelled にする。
    #   pool : 照合のプール（newContentIndexWorkerPool）。$null なら searchContentIndex が必要なときだけ作る
    #   cache: 読んだ本文インデックスのファイルの内容の入れ物（newTsvTextCache）
    param (
        [hashtable]$request,
        $pool = $null,
        $cache = $null
    )

    try {
        if ($request.Stop) {
            $request.Cancelled = $true
            return
        }
        $word = $request.Word
        $folders = $request.Folders
        $own = [Workspace]::new($request.WorkDir)
        # 前の検索の溢れを消し、上限があるときだけ、この検索の溢れの置き場所を決める（作るのは、溢れが出たときに照合のスレッドが行う）
        [void](removeSearchOverflow (getSearchOverflowPidDir $own.TmpRoot) $own.TmpRoot)
        if ($request.Limit -gt 0) {
            if (testSearchOverflowRoot (getSearchOverflowRoot $own.TmpRoot) $own.TmpRoot) {
                $request.OverflowDir = getSearchOverflowDir $own.TmpRoot $request.Number
            } else {
                $request.OverflowFailed = $true
            }
        }
        # gram 索引で、検索語のすべての 2-gram を含むファイルだけに絞る（索引を使えなければ $null）
        $index = getGramContentIndexFiles $word $request.SimpleMatch $folders -indexRoot $own.IndexDir -gramRoot $own.GramDir `
            -onProgress { param ($count) $request.Scanned = $count } -stopSignal $request
        $request.GramUsed = ($null -ne $index)
        # gram 索引の絞り込みの途中で止められたとき、$null を受けて全部の列挙に進まない
        if ($request.Stop) {
            $request.Cancelled = $true
            return
        }
        if ($null -eq $index) {
            # 途中の件数を画面に伝える（止まって見えないように）
            $index = getContentIndexFiles $folders { param ($count) $request.Scanned = $count } -stopSignal $request
        }
        # 列挙の途中で止められたとき、照合に進まない
        if ($request.Stop) {
            $request.Cancelled = $true
            return
        }
        $request.Folders = $index.Folders
        $request.Total = $index.ContentIndexFiles.Count
        $request.IndexTotal = $index.ContentIndexFiles.Count
        $result = searchContentIndex $word $index.ContentIndexFiles $request.SimpleMatch $request.Limit -caseSensitive $request.CaseSensitive `
            -fileFilter $request.FileFilter -cache $cache -includeShapes $request.IncludeShapes -includeComments $request.IncludeComments -pool $pool `
            -overflowDir $request.OverflowDir `
            -onProgress {
            param ($done, $total, $newHits, $hitTotal)
            foreach ($hit in $newHits) {
                $request.Queue.Enqueue($hit)
            }
            $request.Total = $total
            $request.Done = $done
            # 数えている途中でも、上限を超えたことと今までの件数を画面に伝える
            $request.HitTotal = $hitTotal
            if ($request.Limit -gt 0 -and $hitTotal -gt $request.Limit) { $request.Truncated = $true }
        } -stopSignal $request
        $request.Total = $result.Total
        $request.Truncated = $result.Truncated
        $request.HitTotal = $result.HitTotal
        $request.CountTimeout = $result.CountTimeout
        if ($result.OverflowFailed) { $request.OverflowFailed = $true }
        $request.Cancelled = $result.Cancelled
        $request.Skipped = @($result.Skipped)
    } catch {
        $request.Error = $_.Exception.Message
    } finally {
        $request.Finished = $true
    }
}

function testIndexExists {
    # 検索対象インデックスに本文インデックスのファイルが1件でもあるか（最初の1件が見つかった時点で打ち切る）
    param (
        [string[]]$folders = @($workspace.IndexDir)
    )

    foreach ($dir in $folders) {
        if (!(Test-Path -LiteralPath $dir -PathType Container)) {
            continue
        }
        try {
            # 1 件見つけたらやめる。列挙子を閉じずに戻ると content_index の下のフォルダを掴んだまま残り、
            # 続くインポートの上書きで移動できなくなるため testAnyEntry を通す
            if (testAnyEntry ([System.IO.Directory]::EnumerateFiles((toLongPath (Resolve-Path -LiteralPath $dir).ProviderPath), ${contentIndexFilePattern}, [System.IO.SearchOption]::AllDirectories))) {
                return $true
            }
        } catch {
            # アクセスできないフォルダがある場合は、件数を数える方で判定する
            if (@(Get-ChildItem -LiteralPath $dir -Filter ${contentIndexFilePattern} -File -Recurse -ErrorAction SilentlyContinue).Count -gt 0) {
                return $true
            }
        }
    }
    return $false
}

function getIndexSummary {
    # 検索対象インデックスの本文インデックスのファイルの件数と最新の更新日時を返す: @{ Count; LastWrite（無ければ $null）; Missing（存在しないフォルダ） }
    param (
        [string[]]$folders = @($workspace.IndexDir)
    )

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $summary = @{ Count = 0; LastWrite = $null; Missing = @() }
    foreach ($dir in $folders) {
        if (!(Test-Path -LiteralPath $dir -PathType Container)) {
            $summary.Missing += $dir
            continue
        }
        $longDir = toLongPath (Resolve-Path -LiteralPath $dir).ProviderPath
        # アクセスできないフォルダがあると .NET の列挙は途中で止まるため、そのときは Get-ChildItem で数える（読めるものだけ）
        try {
            $found = [System.IO.DirectoryInfo]::new($longDir).GetFiles(${contentIndexFilePattern}, [System.IO.SearchOption]::AllDirectories)
        } catch {
            $found = @(Get-ChildItem -LiteralPath $longDir -Filter ${contentIndexFilePattern} -File -Recurse -ErrorAction SilentlyContinue)
        }
        foreach ($file in $found) {
            if (!$seen.Add($file.FullName)) {
                continue
            }
            $summary.Count++
            if ($null -eq $summary.LastWrite -or $file.LastWriteTime -gt $summary.LastWrite) {
                $summary.LastWrite = $file.LastWriteTime
            }
        }
    }
    return $summary
}

function getSearchResultLine {
    # 検索結果の 1 行（TSV）。ファイル名の前に、クロール対象フォルダからの相対フォルダを付ける（toSearchResultLines・画面のファイル出力で共通）
    param (
        [string]$relDir,
        [string]$book,
        [string]$location,
        [int]$lineNumber,
        [string]$line
    )

    $text = toResultLine $book $location $lineNumber $line
    if ($relDir) {
        return "${relDir}\${text}"
    }
    return $text
}

function getSearchResultColumnCount {
    # 結果の 1 行の、見出しに要る列の数（ファイル名・場所・種別・行番号の 4 列を除いたセルの数）
    param (
        [string]$line
    )

    return (countTsvFields $line) - 4
}

function toSearchResultLines {
    # 検索結果を検索結果ファイルの形式にし、@{ Header（見出し行）; Lines } を返す（画面のファイル出力・コピーで共通）。
    # Excelに貼り付けたときに元のセル位置が分かるよう、行番号（TSVの行番号 = Excelの行番号）を付け、
    # 見出しには該当行の最大セル数分の列名（A, B, C…）を付ける
    param (
        [object[]]$hits
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $columnCount = 0
    foreach ($hit in $hits) {
        $line = getSearchResultLine $hit.RelDir $hit.Book $hit.Location $hit.LineNumber $hit.Line
        $lines.Add($line)
        $columnCount = [math]::Max($columnCount, (getSearchResultColumnCount $line))
    }
    return @{ Header = (toResultHeader $columnCount); Lines = $lines }
}

function toResultCells {
    # 検索結果 1 件を、ファイルに出力するときの列（string[]）にする。純粋な関数（形式ごとの列の分け方はここで決める）。
    #   xlsx: ファイル名・場所・種別・行・該当行の 5 列（該当行のセルは画面の表と同じ " │ " でつないで 1 セル）
    #   csv・tsv: ファイル名・場所・種別・行・該当行のセル…（セルごとに 1 列。2 つは区切りだけが違い、列は同じ）
    # ファイル名（相対フォルダ付き）・場所の制御文字はスペースにし、セルの中の CR は LF にそろえる
    param (
        [string]$relDir,
        [string]$book,
        [string]$location,
        [int]$lineNumber,
        [string]$line,
        [string]$format
    )

    $controls = '[\x00-\x1F]'
    $book = [regex]::Replace($book, $controls, " ")
    $described = describePlace $book $location
    $name = $book
    if ($relDir) {
        $name = [regex]::Replace($relDir, $controls, " ") + "\" + $name
    }
    $cells = New-Object System.Collections.Generic.List[string]
    $cells.Add($name)
    $cells.Add([regex]::Replace($described.Place, $controls, " "))
    $cells.Add($described.Kind)
    $cells.Add([string]$lineNumber)

    if ((getContentIndexFileKind $book) -eq "Excel") {
        # 囲みの " を外し、セル内改行（U+2028）を LF に戻す
        $hitCells = @((splitTsvCells $line) | ForEach-Object { $_.Replace(${cellNewLine}, "`n") })
    } else {
        $hitCells = @($line.Split("`t"))
    }
    $hitCells = @($hitCells | ForEach-Object { $_.Replace("`r`n", "`n").Replace("`r", "`n") })

    if ($format -eq "xlsx") {
        $cells.Add(($hitCells -join " │ "))
    } else {
        $cells.AddRange([string[]]$hitCells)
    }
    return , $cells.ToArray()
}
