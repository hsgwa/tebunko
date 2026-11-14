# 前の検索の結果を使い回す（状態層）。検索語が前の語を含む形で長くなったとき（「見積」→「見積書」）、
# 前の検索のヒットの行（場所ごとのまとまり）だけを新しい語で照合し直し、本文インデックスのファイルを読まずに結果を出す。結果は全体を探し直したときと同じになる。
# 絞ってよいかの判定は getNarrowSearchReason（search_query.ps1）。ここでは、前の検索の記録を作る・本文インデックスのファイルの一覧が同じかを比べる・
# ヒットを選ぶ・絞った検索を実行する。

function testRecordableHitLines {
    # ヒットの行（まとまりの一覧 runs の Lines）が、元の行そのままか。テキストファイルの長い行は表示用に切って "…" を付ける（truncateHitLine）ため、
    # 切られた可能性のある行（テキストの行で、先頭か末尾が "…"）があれば $false（その行では新しい語の照合が全体の検索と変わりうる）
    param (
        $runs
    )

    foreach ($run in $runs) {
        if ((getContentIndexFileKind $run.Book) -ne "テキスト") {
            continue
        }
        foreach ($line in $run.Lines) {
            if ($line.StartsWith("…") -or $line.EndsWith("…")) {
                return $false
            }
        }
    }
    return $true
}

function newSearchRecord {
    # 最後まで検索し終えた要求の記録を作る（次の要求が絞れるかの判定と、絞った検索で使う）。記録できないとき（上限なし・打ち切り・飛ばしたファイルあり・切った行あり）は $null
    #   request: 終えた要求（Word などの条件）/ folders: 要求の検索対象（invokeSearchRequest が Folders を列挙の結果で上書きする前の値）
    #   files  : 列挙した本文インデックスのファイルのすべて（gram 索引で絞る前）/ hits: ヒットのまとまりの一覧（画面に渡した同じオブジェクト）
    param (
        [hashtable]$request,
        $folders,
        $files,
        $runs
    )

    if ($request.Limit -le 0 -or $request.Truncated -or $request.CountTimeout -or $request.Cancelled -or @($request.Skipped).Count -gt 0) {
        return $null
    }
    if (!(testRecordableHitLines $runs)) {
        return $null
    }
    $paths = [string[]]::new($files.Count)
    $ticks = [long[]]::new($files.Count)
    $sizes = [long[]]::new($files.Count)
    for ($i = 0; $i -lt $files.Count; $i++) {
        $paths[$i] = [string]$files[$i].Path
        $ticks[$i] = [long]$files[$i].Ticks
        $sizes[$i] = [long]$files[$i].Size
    }
    return @{
        Word = $request.Word; SimpleMatch = $request.SimpleMatch; CaseSensitive = $request.CaseSensitive; FileFilter = $request.FileFilter
        IncludeShapes = $request.IncludeShapes; IncludeComments = $request.IncludeComments; Folders = $folders; WorkDir = $request.WorkDir; Limit = $request.Limit
        Paths = $paths; Ticks = $ticks; Sizes = $sizes; Runs = $runs
    }
}

function testSearchRecordFiles {
    # 今の本文インデックスのファイルの一覧が、記録した一覧と同じか（件数・順番・パス・更新日時・大きさ）。1 つでも違えば $false（書き直された・足された・消えた）
    param (
        $record,
        $files
    )

    if ($files.Count -ne $record.Paths.Length) {
        return $false
    }
    for ($i = 0; $i -lt $files.Count; $i++) {
        $file = $files[$i]
        if (!([string]$file.Path).Equals($record.Paths[$i], [System.StringComparison]::OrdinalIgnoreCase) -or [long]$file.Ticks -ne $record.Ticks[$i] -or [long]$file.Size -ne $record.Sizes[$i]) {
            return $false
        }
    }
    return $true
}

function selectNarrowHits {
    # 記録のヒット（まとまりの一覧）のうち、新しい語の正規表現に一致する行だけを、順番を保ったまとまりの一覧にして返す（行が残らないまとまりは落とす）。判定は全体の検索の 1 行ずつの照合と同じ
    # （Excel の図形・コメントの行は、区切りのタブより後から始まる一致だけをヒットにする）
    param (
        $record,
        [string]$word,
        [bool]$caseSensitive
    )

    $regex = (newSearchRegex $word $true $caseSensitive).Regex
    $selected = New-Object System.Collections.Generic.List[hashtable]
    foreach ($run in $record.Runs) {
        $cellPrefixed = testCellPrefixedPlace $run.Book (splitObjectPlace $run.Location).Kind
        $kept = $null
        for ($i = 0; $i -lt $run.Lines.Count; $i++) {
            $line = [string]$run.Lines[$i]
            if (!$regex.IsMatch($line) -or ($cellPrefixed -and !(testCellPrefixedHit $regex $line))) {
                continue
            }
            if ($null -eq $kept) {
                $kept = @{} + $run
                $kept.LineNumbers = [System.Collections.Generic.List[int]]::new()
                $kept.Lines = [System.Collections.Generic.List[string]]::new()
                $selected.Add($kept)
            }
            $kept.LineNumbers.Add($run.LineNumbers[$i])
            $kept.Lines.Add($line)
        }
    }
    return , $selected
}

function invokeNarrowSearch {
    # 絞ってよいと判定された要求を、記録のヒットから実行する。今の本文インデックスのファイルの一覧が記録と違えば（IndexChanged）、
    # 何も変えずに $false を返し、呼び出し側が全体の検索に進む。絞れたら要求にヒットと進み具合を入れ、記録を置き換えて $true を返す。
    # 取り消されたら要求を Cancelled にして $true を返す（記録は変えない）
    #   state: @{ Record } 前の検索の記録の入れ物
    param (
        [hashtable]$request,
        [hashtable]$state
    )

    $record = $state.Record
    $listing = getContentIndexFiles $request.Folders { param ($count) $request.Scanned = $count } -stopSignal $request
    if ($request.Stop) {
        $request.Cancelled = $true
        return $true
    }
    if (!(testSearchRecordFiles $record $listing.ContentIndexFiles)) {
        $request.NarrowReason = "IndexChanged"
        return $false
    }
    $selected = selectNarrowHits $record $request.Word $request.CaseSensitive
    if ($request.Stop) {
        $request.Cancelled = $true
        return $true
    }
    foreach ($run in $selected) {
        $request.Queue.Enqueue($run)
    }
    $count = $listing.ContentIndexFiles.Count
    $request.Folders = $listing.Folders
    $request.Scanned = $count
    $request.Total = $count
    $request.IndexTotal = $count
    $request.Done = $count
    $request.HitTotal = getHitRunCount $selected
    $request.Narrowed = $true
    $request.NarrowReason = "Narrow"
    $state.Record = $record.Clone()
    $state.Record.Word = $request.Word
    $state.Record.SimpleMatch = $request.SimpleMatch
    $state.Record.Runs = $selected
    return $true
}

function tryNarrowSearch {
    # 前の検索の記録から絞ってよいか判定し（NarrowReason に符号を入れる）、絞ってよければ絞った検索を実行する。要求に答えた（取り消しを含む）なら $true、全体の検索に進むなら $false
    param (
        [hashtable]$request,
        [hashtable]$state
    )

    $request.NarrowReason = getNarrowSearchReason $state.Record $request
    return ($request.NarrowReason -eq "Narrow") -and (invokeNarrowSearch $request $state)
}

function updateSearchRecord {
    # 最後まで検索し終えた要求なら、記録を置き換える（記録できないときは前の記録を残す）
    param (
        [hashtable]$request,
        [hashtable]$state,
        $folders,
        $files,
        $runs
    )

    $record = newSearchRecord $request $folders $files $runs
    if ($null -ne $record) {
        $state.Record = $record
    }
}
