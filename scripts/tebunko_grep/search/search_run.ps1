# 検索結果の組み立て・書き出しと、インデックスの有無・件数（検索そのものは pack_search.ps1）。

function toResultLine {
    # 検索結果1件を "ファイル名<TAB>場所<TAB>種別<TAB>行番号<TAB>該当行" に整形する。場所・種別は画面と同じ表示（describePlace）。
    # Excelに貼り付けたとき、該当行の各セルが元の列の順（5列目 = A列）に並ぶようにする
    param (
        [string]$book,
        [string]$location,
        [int]$lineNumber,
        [string]$line
    )

    if ($book -match "\.xls[a-z]?$") {
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


# 検索で読んだ集約ファイルの内容を残しておく量の上限（文字数。1 文字 2 バイトのため約 128MB）
${searchCacheMaxChars} = 64000000

function newTsvTextCache {
    # 検索で読んだ集約ファイルの内容を、次の検索で使い回すための入れ物を作る（画面が 1 つ持ち、検索のたびに searchPackIndex に渡す）。
    # 2 回目以降の検索ではファイルを開かない。更新日時・サイズが列挙したときと違う集約ファイル（書き直した等）は読み直す。
    # 並列検索の各スレッドから使うため、中身は ConcurrentDictionary。
    #   Texts: 集約ファイルの \\?\ 付きのパス → @(更新日時（UTC の Ticks）, サイズ, 内容, 場所の一覧, 最後に使った世代) / Chars: 残している文字数 / MaxChars: 上限
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
    # 古い世代から追い出す（消した集約ファイル・検索しなくなったフォルダの分）。今の世代で使ったものは残す
    # （毎回同じ集約ファイルを順に読むため、使った順だけで追い出すと、上限より大きいインデックスでは何も残らない）。
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

function newSearchRequest {
    # 検索の要求 1 つを作る（画面が作り、検索の司令のスレッド（SearchService）が invokeSearchRequest で実行する）。
    # 画面と司令のスレッドの両方から読み書きするため、Synchronized の hashtable にする。
    #   画面が書く: 検索条件・Stop（取り消し）
    #   司令が書く: Queue（ヒット）・Done / Total / IndexTotal / Scanned（進み具合）・Folders・FastUsed / FastAvailable・Truncated / Cancelled / Error・Finished
    param (
        [string]$word,
        [bool]$simpleMatch,
        $folders,
        [int]$limit,
        [hashtable]$option = @{},
        [bool]$useFast = $false
    )

    return [hashtable]::Synchronized(@{
        Word = $word; SimpleMatch = $simpleMatch; Folders = $folders; Limit = $limit
        CaseSensitive = [bool]$option.CaseSensitive; FileFilter = [string]$option.FileFilter
        IncludeShapes = ($option.IncludeShapes -ne $false); IncludeComments = ($option.IncludeComments -ne $false)
        UseFast = $useFast; FastUsed = $false; FastAvailable = $null
        Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        Stop = $false; Finished = $false; Done = 0; Total = -1; IndexTotal = -1; Scanned = 0
        Truncated = $false; Cancelled = $false; Error = $null
    })
}

function invokeSearchRequest {
    # 検索の要求（newSearchRequest）を実行し、ヒットと進み具合を要求に少しずつ入れる。最後に Finished を立てる。
    # 例外は投げずに Error に入れる。始める前に取り消されていたら（次の要求が来た等）、何もせずに Cancelled にする。
    #   pool : 照合のプール（newPackWorkerPool）。$null なら searchPackIndex が必要なときだけ作る
    #   cache: 読んだ集約ファイルの内容の入れ物（newTsvTextCache）
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
        # 高速検索が使えるなら、Windows Search で検索語を含みうるフォルダを先に絞る（使えなければ $null で、すべてを集める）
        $index = $null
        if ($request.UseFast) {
            $request.FastAvailable = testWindowsSearch
            if ($request.FastAvailable) {
                $index = getFastSearchPackFiles $word $folders -onProgress { param ($count) $request.Scanned = $count }
            }
        }
        $request.FastUsed = ($null -ne $index)
        if ($null -eq $index) {
            # 途中の件数を画面に伝える（止まって見えないように）
            $index = getIndexPackFiles $folders { param ($count) $request.Scanned = $count }
        }
        $request.Folders = $index.Folders
        $request.Total = $index.Packs.Count
        $request.IndexTotal = $index.Packs.Count
        $result = searchPackIndex $word $index.Packs $request.SimpleMatch $request.Limit -caseSensitive $request.CaseSensitive `
            -fileFilter $request.FileFilter -cache $cache -includeShapes $request.IncludeShapes -includeComments $request.IncludeComments -pool $pool `
            -onProgress {
            param ($done, $total, $newHits)
            foreach ($hit in $newHits) {
                $request.Queue.Enqueue($hit)
            }
            $request.Total = $total
            $request.Done = $done
        } -shouldStop { $request.Stop }
        $request.Total = $result.Total
        $request.Truncated = $result.Truncated
        $request.Cancelled = $result.Cancelled
    } catch {
        $request.Error = $_.Exception.Message
    } finally {
        $request.Finished = $true
    }
}

function testIndexExists {
    # 検索対象インデックスに集約ファイルが1件でもあるか（最初の1件が見つかった時点で打ち切る）
    param (
        [string[]]$folders = @(${indexDir})
    )

    foreach ($dir in $folders) {
        if (!(Test-Path -LiteralPath $dir -PathType Container)) {
            continue
        }
        try {
            $enumerator = [System.IO.Directory]::EnumerateFiles((toLongPath (Resolve-Path -LiteralPath $dir).ProviderPath), ${packFilePattern}, [System.IO.SearchOption]::AllDirectories).GetEnumerator()
            if ($enumerator.MoveNext()) {
                return $true
            }
        } catch {
            # アクセスできないフォルダがある場合は、件数を数える方で判定する
            if (@(Get-ChildItem -LiteralPath $dir -Filter ${packFilePattern} -File -Recurse -ErrorAction SilentlyContinue).Count -gt 0) {
                return $true
            }
        }
    }
    return $false
}

function getIndexSummary {
    # 検索対象インデックスの集約ファイルの件数と最新の更新日時を返す: @{ Count; LastWrite（無ければ $null）; Missing（存在しないフォルダ） }
    param (
        [string[]]$folders = @(${indexDir})
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
            $found = [System.IO.DirectoryInfo]::new($longDir).GetFiles(${packFilePattern}, [System.IO.SearchOption]::AllDirectories)
        } catch {
            $found = @(Get-ChildItem -LiteralPath $longDir -Filter ${packFilePattern} -File -Recurse -ErrorAction SilentlyContinue)
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
        # ファイル名の前に、クロール対象フォルダからの相対フォルダを付ける
        $line = toResultLine $hit.Book $hit.Location $hit.LineNumber $hit.Line
        if ($hit.RelDir) {
            $line = "$($hit.RelDir)\${line}"
        }
        $lines.Add($line)
        $columnCount = [math]::Max($columnCount, (countTsvFields $line) - 4)  # ファイル名・場所・種別・行番号の4列を除く
    }
    return @{ Header = (toResultHeader $columnCount); Lines = $lines }
}

function writeSearchResult {
    # 1ワード分の検索結果を検索結果ファイルに書き出す
    param (
        [System.IO.TextWriter]$writer,
        [string]$word,
        [object[]]$hits
    )

    $result = toSearchResultLines $hits
    $writer.WriteLine("【検索文字列　${word}】 $($result.Lines.Count) 件")
    $writer.WriteLine($result.Header)
    foreach ($line in $result.Lines) {
        $writer.WriteLine($line)
    }
    $writer.WriteLine("")
}
