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
    # contentIndexFiles の start から count 件の本文インデックスのファイルを読み、regex に一致する行を @{ Hits; Skipped; Count; Timeout; OverflowFailed } で返す。
    # Hits は PSCustomObject の並び（1 行に複数一致しても 1 件）。max 以上（max+1 件目）が見つかった時点で、gate が無ければ打ち切る（負は上限なし）。
    # 読めない・形が崩れている・contentIndexSearchMaxBytes より大きいファイルは、そのファイルだけ飛ばし、パスを Skipped に入れる（キャッシュには入れない）。
    #   include / exclude: 元のファイル名の条件（$null は条件なし） / excludePlace: 除く場所（図形・コメント）
    #   stopSignal: Stop を持つもの（検索の要求。$null は止めない）。Stop が立ったら、次のファイルから読まずに、ここまでのヒットを返す
    #   cache: newTsvTextCache。更新日時・大きさが同じなら、前に読んだ内容と場所の一覧を使う
    #   gate: 照合の合図（searchContentIndex が作る [int[]]。$null は今までどおり）。型は付けない（写しにすると、ほかのスレッドの書き込みが見えない）。
    #     [0]=止める印・[1]=件数の上限・[2..]=作業ごとの件数の欄・その後ろ=作業ごとの時間切れの欄。欄は書く側が 1 つで、値は増えるだけ（小さい値は書かない）
    #     gate があるとき、max+1 件目を作った時点か、自分より前の作業の欄の和 + 自分の件数が上限を超えた時点で、行を作らずに件数だけを数える（戻らない）。
    #     照合の時間切れは、投げずに Timeout で返す（それまでに作った行を返す）
    #   slot: gate の自分の番号 / checkEvery: gate を見る間隔（ヒットの数）
    #   overflowDir: 数えるだけの間のヒットを <slot>.tsv に書き溜める置き場所（空は書かない。書けなければ OverflowFailed を立てて書くのをやめる）
    # 照合が時間切れ（RegexMatchTimeoutException）で、gate が無ければ、そのまま呼び出し元へ伝える
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
        [regex]$excludePlace = $null,
        $stopSignal = $null,
        $gate = $null,
        [int]$slot = 0,
        [int]$checkEvery = 64,
        [string]$overflowDir = ""
    )

    # 並列検索の別スレッドでも動くよう、コマンドレットは使わない
    $hits = [System.Collections.Generic.List[psobject]]::new()
    $skipped = [System.Collections.Generic.List[string]]::new()
    $end = [Math]::Min($contentIndexFiles.Count, $start + $count)
    $lf = [char]10
    $mark = [char]0x1E
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $mode = if ($null -eq $textRegex) { "scan" } else { $scanMode }
    $gated = $null -ne $gate
    $slotIndex = 2 + $slot
    $limit = if ($gated) { $gate[1] } else { 0 }
    $timeoutIndex = if ($gated) { 2 + ($gate.Length - 2) / 2 + $slot } else { 0 }
    $counting = $false        # 数えるだけの照合に切り替えたか（戻らない）
    $total = 0                # 作った行と数えた件数の合計
    $written = 0              # 自分の欄に書いた最大の値
    $sinceCheck = 0
    $lineTick = 0
    $halt = $false
    $timedOut = $false
    $writer = $null
    $overflowFailed = $false
    $wantOverflow = $gated -and ($overflowDir -ne "")
    try {
        for ($i = $start; $i -lt $end; $i++) {
            # 本文インデックスのファイル 1 件ごとに中止を見る（行ごとには見ない）。ここまでのヒットを返す
            if ($halt -or ($null -ne $stopSignal -and $stopSignal.Stop)) { break }
            if ($gated) {
                if ($gate[0] -ne 0) { break }
                if (!$counting) {
                    $prior = 0
                    for ($g = 2; $g -lt $slotIndex; $g++) { $prior += $gate[$g] }
                    if ($prior + $total -gt $limit) { $counting = $true }
                }
                if ($total -gt $written) { $written = $total; $gate[$slotIndex] = $total }
                $sinceCheck = 0
            }
            $contentIndexFile = $contentIndexFiles[$i]
            $hitsBefore = $hits.Count
            $fileTotal = $total
            $fileCounting = $counting
            $startPos = 0
            if ($null -ne $writer) {
                try { $writer.Flush(); $startPos = $writer.BaseStream.Position } catch { }
            }
            try {
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
                    if ($contentIndexFile.Size -gt ${contentIndexSearchMaxBytes}) {
                        # 全量を読んでメモリを使い切らないよう、大きすぎるファイルは読まずに飛ばす
                        throw "インデックスのファイルが大きすぎます"
                    }
                    $stream = [System.IO.FileStream]::new($contentIndexFile.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
                    $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::Unicode, $true)
                    try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
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
                            # 行番号と行の文字列は、行を作るとき・溢れに書くときだけ作る（数えるだけなら、セル番地の判定に要るときの行だけ）
                            $needLine = !$counting -or $wantOverflow
                            $line = $null
                            if ($needLine) {
                                if ($pos -lt 0) { $pos = $place.Start; $number = $place.FirstLine }
                                # 前に数えた位置から行の先頭までの改行を数える（Substring を作らない）
                                while ($pos -lt $lineStart) {
                                    $n = $text.IndexOf($lf, $pos, $lineStart - $pos)
                                    if ($n -lt 0) { break }
                                    $number++
                                    $pos = $n + 1
                                }
                                $pos = $lineStart
                            }
                            if ($needLine -or $place.CellPrefixed) {
                                $line = $text.Substring($lineStart, $lineEnd - $lineStart)
                            }
                            if ($place.CellPrefixed -and !(testCellPrefixedHit $regex $line)) {
                                # セル番地だけに一致した行はヒットにしない（行番号の数え方は崩さない）
                                if ($lineEnd -ge $length) { break }
                                $m = $textRegex.Match($text, $lineEnd + 1)
                                continue
                            }
                            if ($counting) {
                                $total++
                                if ($wantOverflow) {
                                    if ($isTextBook) {
                                        $line = truncateHitLine $line ($index - $lineStart)
                                    }
                                    if ($null -eq $writer) {
                                        $writer = newSearchOverflowWriter $overflowDir "$slot.tsv"
                                        if ($null -eq $writer) { $overflowFailed = $true; $wantOverflow = $false }
                                    }
                                    if ($null -ne $writer) {
                                        try {
                                            $writer.WriteLine("$($contentIndexFile.RelDir)`t$($place.Book)`t$($place.Location)`t$number`t$line")
                                        } catch {
                                            $overflowFailed = $true
                                            $wantOverflow = $false
                                            try { $writer.Dispose() } catch { }
                                            $writer = $null
                                        }
                                    }
                                }
                            } else {
                                if ($isTextBook) {
                                    $line = truncateHitLine $line ($index - $lineStart)
                                }
                                $hits.Add([pscustomobject]@{
                                    Root = $contentIndexFile.Root; RelPath = $contentIndexFile.RelPath; RelDir = $contentIndexFile.RelDir; FileName = $place.Book
                                    Book = $place.Book; Location = $place.Location; LineNumber = $number; Line = $line
                                })
                                $total++
                                if ($max -ge 0 -and $hits.Count -gt $max) {
                                    if (!$gated) { return @{ Hits = $hits; Skipped = $skipped; Count = $total; Timeout = $false; OverflowFailed = $false } }
                                    $counting = $true
                                }
                            }
                            if ($gated) {
                                $sinceCheck++
                                if ($sinceCheck -ge $checkEvery) {
                                    $sinceCheck = 0
                                    if ($gate[0] -ne 0 -or ($null -ne $stopSignal -and $stopSignal.Stop)) {
                                        $halt = $true
                                    } elseif (!$counting) {
                                        $prior = 0
                                        for ($g = 2; $g -lt $slotIndex; $g++) { $prior += $gate[$g] }
                                        if ($prior + $total -gt $limit) { $counting = $true }
                                    }
                                    if ($total -gt $written) { $written = $total; $gate[$slotIndex] = $total }
                                    if ($halt) { break }
                                }
                            }
                            if ($lineEnd -ge $length) { break }
                            $m = $textRegex.Match($text, $lineEnd + 1)
                        }
                        continue
                    } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                        # 全文への照合が時間切れなら、1 行ずつ照合し直す（このファイルで足した行・数えた件数・溢れの行は元に戻す）
                        if ($hits.Count -gt $before) { $hits.RemoveRange($before, $hits.Count - $before) }
                        $total = $fileTotal
                        $counting = $fileCounting
                        # 溢れのファイルを、ファイルの始めの位置まで戻す。Position も戻さないと次の書き込みが穴を作る
                        if ($null -ne $writer) {
                            try {
                                $writer.Flush()
                                $writer.BaseStream.SetLength($startPos)
                                $writer.BaseStream.Position = $startPos
                            } catch {
                                $overflowFailed = $true
                                $wantOverflow = $false
                                try { $writer.Dispose() } catch { }
                                $writer = $null
                            }
                        }
                    }
                }
                # filter・scan（と lines の時間切れ）: 場所ごとに 1 行ずつ照合する
                foreach ($place in $places) {
                    if ($include -and !$include.IsMatch($place.Book)) { continue }
                    if ($exclude -and $exclude.IsMatch($place.Book)) { continue }
                    if ($excludePlace -and $excludePlace.IsMatch($place.Location)) { continue }
                    $isTextBook = ((getContentIndexFileKind $place.Book) -eq "テキスト")
                    $number = $place.FirstLine - 1
                    $pos = $place.Start
                    while ($pos -lt $place.End) {
                        $n = $text.IndexOf($lf, $pos, $place.End - $pos)
                        if ($n -lt 0) { $n = $place.End }
                        $number++
                        $line = $text.Substring($pos, $n - $pos)
                        $pos = $n + 1
                        if ($counting) {
                            # 数えるだけの間は、1,024 行ごとにも止める印を見る
                            $lineTick++
                            if ($lineTick -ge 1024) {
                                $lineTick = 0
                                if (($null -ne $stopSignal -and $stopSignal.Stop) -or $gate[0] -ne 0) { $halt = $true; break }
                            }
                        }
                        $lineMatch = $regex.Match($line)
                        if (!$lineMatch.Success) { continue }
                        if ($place.CellPrefixed -and !(testCellPrefixedHit $regex $line)) { continue }
                        if ($counting) {
                            $total++
                            if ($wantOverflow) {
                                if ($isTextBook) {
                                    $line = truncateHitLine $line $lineMatch.Index
                                }
                                if ($null -eq $writer) {
                                    $writer = newSearchOverflowWriter $overflowDir "$slot.tsv"
                                    if ($null -eq $writer) { $overflowFailed = $true; $wantOverflow = $false }
                                }
                                if ($null -ne $writer) {
                                    try {
                                        $writer.WriteLine("$($contentIndexFile.RelDir)`t$($place.Book)`t$($place.Location)`t$number`t$line")
                                    } catch {
                                        $overflowFailed = $true
                                        $wantOverflow = $false
                                        try { $writer.Dispose() } catch { }
                                        $writer = $null
                                    }
                                }
                            }
                        } else {
                            if ($isTextBook) {
                                $line = truncateHitLine $line $lineMatch.Index
                            }
                            $hits.Add([pscustomobject]@{
                                Root = $contentIndexFile.Root; RelPath = $contentIndexFile.RelPath; RelDir = $contentIndexFile.RelDir; FileName = $place.Book
                                Book = $place.Book; Location = $place.Location; LineNumber = $number; Line = $line
                            })
                            $total++
                            if ($max -ge 0 -and $hits.Count -gt $max) {
                                if (!$gated) { return @{ Hits = $hits; Skipped = $skipped; Count = $total; Timeout = $false; OverflowFailed = $false } }
                                $counting = $true
                            }
                        }
                        if ($gated) {
                            $sinceCheck++
                            if ($sinceCheck -ge $checkEvery) {
                                $sinceCheck = 0
                                if ($gate[0] -ne 0 -or ($null -ne $stopSignal -and $stopSignal.Stop)) {
                                    $halt = $true
                                } elseif (!$counting) {
                                    $prior = 0
                                    for ($g = 2; $g -lt $slotIndex; $g++) { $prior += $gate[$g] }
                                    if ($prior + $total -gt $limit) { $counting = $true }
                                }
                                if ($total -gt $written) { $written = $total; $gate[$slotIndex] = $total }
                                if ($halt) { break }
                            }
                        }
                    }
                    if ($halt) { break }
                }
            } catch {
                # 照合の時間切れは検索語の問題。合図が無ければ全体を中止する（そのまま伝える）。合図があれば、それまでに作った行を返して時間切れの欄に 1 を書き、
                # 中止にするか総数だけ数えられないことにするかは、受け取る側（searchContentIndex）が決める。
                # それ以外（読めない・形が崩れている・大きすぎる）は、そのファイルだけ飛ばして知らせる
                if ($_.Exception -is [System.Text.RegularExpressions.RegexMatchTimeoutException] -or $_.Exception.InnerException -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) {
                    if (!$gated) { throw }
                    $timedOut = $true
                    $gate[$timeoutIndex] = 1
                    break
                }
                if ($hits.Count -gt $hitsBefore) { $hits.RemoveRange($hitsBefore, $hits.Count - $hitsBefore) }
                $total = $fileTotal
                $counting = $fileCounting
                # 溢れのファイルを、ファイルの始めの位置まで戻す。Position も戻さないと次の書き込みが穴を作る
                if ($null -ne $writer) {
                    try {
                        $writer.Flush()
                        $writer.BaseStream.SetLength($startPos)
                        $writer.BaseStream.Position = $startPos
                    } catch {
                        $overflowFailed = $true
                        $wantOverflow = $false
                        try { $writer.Dispose() } catch { }
                        $writer = $null
                    }
                }
                $skipped.Add($contentIndexFile.Path)
            }
        }
        if ($gated -and $total -gt $written) { $gate[$slotIndex] = $total }
    } finally {
        if ($null -ne $writer) {
            try { $writer.Dispose() } catch { $overflowFailed = $true }
        }
    }
    return @{ Hits = $hits; Skipped = $skipped; Count = $total; Timeout = $timedOut; OverflowFailed = $overflowFailed }
}


# 本文インデックスのファイルの検索で、1 つのスレッドにまとめて渡す大きさの目安（バイト）
${contentIndexTaskBytes} = 16MB

# 本文インデックスのファイルの検索の各スレッドで動かすスクリプト。@{ Hits; Timeout; Skipped; Count; OverflowFailed } を返す
# gate が無いときの照合の時間切れは、Hits = $null・Timeout = $true で返す（gate があるときは、それまでの Hits を付けて Timeout = $true で返る）
${contentIndexWorkerScript} = {
    param ($contentIndexFiles, $start, $count, $regex, $max, $textRegex, $scanMode, $cache, $include, $exclude, $excludePlace, $stopSignal, $gate = $null, $slot = 0, $checkEvery = 64, $overflowDir = "")
    try {
        $found = searchContentIndexFiles $contentIndexFiles $start $count $regex $max $textRegex $scanMode $cache $include $exclude $excludePlace $stopSignal $gate $slot $checkEvery $overflowDir
        @{ Hits = $found.Hits; Timeout = $found.Timeout; Skipped = $found.Skipped; Count = $found.Count; OverflowFailed = $found.OverflowFailed }
    } catch {
        if ($_.Exception -is [System.Text.RegularExpressions.RegexMatchTimeoutException] -or
            $_.Exception.InnerException -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) {
            @{ Hits = $null; Timeout = $true; Skipped = $null; Count = 0; OverflowFailed = $false }
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

    $state = newWorkerState @("searchContentIndexFiles", "testCellPrefixedHit", "readContentIndexPlaces", "addContentIndexPlace", "getIndexRebuildMessage", "testContentIndexCellPrefixed", "testCellPrefixedPlace", "convertContentIndexMetaToPlace", "decodeContentIndexValue", "getContentIndexFileKind", "getOfficeFileKind", "testTextExtension", "testPdfExtension", "truncateHitLine", "newSearchOverflowWriter", "testSourceBookName") `
        @("contentIndexMark", "contentIndexVersion", "contentIndexSplitVersion", "indexRebuildMessage", "officeFileKinds", "contentIndexPlaceKeys", "placeKindShape", "placeKindComment", "placeKindHeaderFooter", "placeKindEmbed", "textExtensions", "pdfExtensions", "officeExtensions", "displayControlChars", "hitLineMaxChars", "hitLineBeforeMatchChars", "contentIndexSearchMaxBytes")
    return [WorkerPool]::new($workers, $state, $Host, "Normal")
}


function splitContentIndexTasks {
    # 本文インデックスのファイルの並びを、1 つのスレッドに渡す単位（@{ Start; Count }）に分ける（合計がおよそ contentIndexTaskBytes になるまでまとめる）
    param (
        $contentIndexFiles,
        [long]$taskBytes = ${contentIndexTaskBytes},
        [long]$headBytes = 0
    )

    # headBytes（0 は使わない）: 先頭の作業の大きさ。作業ごとに 2 倍にして taskBytes まで大きくする。
    # 上限を超えて数えるとき、表に入る先頭の分を早く返すため（先頭の作業が大きいと、その全体を照合し終えるまで表が埋まらない）
    $tasks = New-Object System.Collections.Generic.List[hashtable]
    $start = 0
    $bytes = 0L
    $target = if ($headBytes -gt 0 -and $headBytes -lt $taskBytes) { $headBytes } else { $taskBytes }
    for ($i = 0; $i -lt $contentIndexFiles.Count; $i++) {
        $bytes += [long]$contentIndexFiles[$i].Size
        if ($bytes -ge $target -or $i -eq $contentIndexFiles.Count - 1) {
            $tasks.Add(@{ Start = $start; Count = $i - $start + 1 })
            $start = $i + 1
            $bytes = 0L
            $target = [Math]::Min($taskBytes, $target * 2)
        }
    }
    return , $tasks
}


function searchContentIndex {
    # 本文インデックスのファイルをワードで検索し、ヒットした行を返す（画面の検索処理）。
    #   contentIndexFiles        : findContentIndexFiles・getContentIndexFiles の結果
    #   simpleMatch  : $true なら文字どおりに検索する。$false なら正規表現として検索し、正規表現として不正なら文字どおりに検索する
    #   limit        : 件数の上限（0 は上限なし）。表（Hits）には上限までを入れ、超えた分は止めずに数え続けて HitTotal に入れる。overflowDir があれば超えた分の行を書き溜める
    #   stopSignal   : Stop を持つもの（検索の要求。$null は止めない）。Stop が立ったら、本文インデックスのファイル 1 件ごとに止める（作業スレッドにも同じものを渡す）
    #   caseSensitive: 英字の大文字・小文字を区別する（newSearchRegex）
    #   fileFilter   : 対象ファイル（newFileFilter）。元のファイル名が一致しないものは検索しない
    #   workerCount  : 並列に検索するスレッドの数（0 は CPU のコア数から決める。最大 4）
    #   cache        : 読んだ本文インデックスのファイルの内容を次の検索で使い回す入れ物（newTsvTextCache。$null は使い回さない）
    #   includeShapes / includeComments: 図形・コメントの場所（"<シート名>[図形]" 等）も検索する（newPlaceExclude）
    #   taskBytes    : 1 つのスレッドにまとめて渡す大きさの目安（バイト）
    #   onProgress   : 1 つの作業を照合するたびに呼ぶ { param($done, $total, $newHits, $hitTotal) }（done・total は本文インデックスのファイルの数。hitTotal は今までの件数）
    #   pool         : 照合に使うプール（newContentIndexWorkerPool。検索の司令が使い回す）。$null なら、並列にするときだけ作って最後に閉じる
    #   checkEvery   : 作業が照合の合図（件数・止める印）を見る間隔（ヒットの数）
    #   overflowDir  : 溢れの置き場所（search_overflow.ps1。空は書き溜めない）
    # @{ Hits; SimpleMatch（実際に文字どおり検索したか）; Total（本文インデックスのファイルの数）; Truncated（上限を超えた）; HitTotal（上限を超えた分も含む件数。中止・総数の時間切れは -1）;
    #    CountTimeout（上限を超えた後の照合が時間切れで、総数が数えられなかった）; OverflowFailed（溢れの書き溜めに失敗した）; Cancelled; Skipped（飛ばしたファイルのパスの一覧） } を返す。
    # 中止（Cancelled）のとき、途中で止めたジョブのヒット・Skipped は結果にも画面にも入れない（止める前に終えたジョブの分だけを返す）。
    # Hits の各要素は PSCustomObject（Root; RelPath（本文インデックスのファイル）; RelDir; FileName; Book; Location; LineNumber; Line）
    param (
        [string]$word,
        $contentIndexFiles,
        [bool]$simpleMatch = $false,
        [int]$limit = 0,
        $stopSignal = $null,
        [bool]$caseSensitive = $false,
        [string]$fileFilter = "",
        [int]$workerCount = 0,
        $cache = $null,
        [bool]$includeShapes = $true,
        [bool]$includeComments = $true,
        [long]$taskBytes = ${contentIndexTaskBytes},
        [scriptblock]$onProgress = $null,
        $pool = $null,
        [int]$checkEvery = 64,
        [string]$overflowDir = ""
    )

    $search = newSearchRegex $word $simpleMatch $caseSensitive
    $filter = newFileFilter $fileFilter
    $excludePlace = newPlaceExclude $includeShapes $includeComments
    $hits = New-Object System.Collections.Generic.List[psobject]
    $skipped = New-Object System.Collections.Generic.List[string]
    $result = @{ Hits = $hits; SimpleMatch = $search.SimpleMatch; Total = $contentIndexFiles.Count; Truncated = $false; HitTotal = 0; CountTimeout = $false; OverflowFailed = $false; Cancelled = $false; Skipped = $skipped }
    $timeoutMessage = "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"
    $tasks = splitContentIndexTasks $contentIndexFiles $taskBytes $(if ($limit -gt 0) { [long]($taskBytes / 64) } else { 0 })
    $workers = if ($pool) { $pool.Size } elseif ($workerCount -gt 0) { $workerCount } else { getWorkerCount }
    if ($tasks.Count -lt 2) { $workers = 1 }
    # 照合の合図（件数の上限を超えたら、表に入らない分を行を作らずに数えるため）。[0]=止める印・[1]=上限・[2..]=作業ごとの件数・その後ろ=作業ごとの時間切れ
    $gate = $null
    if ($limit -gt 0) {
        $gate = [int[]]::new(2 + 2 * $tasks.Count)
        $gate[1] = $limit
    }

    $ownPool = $null
    $pending = New-Object System.Collections.Generic.Queue[hashtable]
    $drain = $false   # 中止したとき、残った作業を受け取って捨てる（プールを使い回せる）。時間切れ・打ち切り・例外のときは Cancel で捨てる
    try {
        if ($workers -le 1) {
            $pool = $null
        } elseif ($null -eq $pool) {
            $ownPool = newContentIndexWorkerPool $workers
            $pool = $ownPool
        }
        $next = 0
        while ($next -lt $tasks.Count -or $pending.Count -gt 0) {
            if ($null -ne $stopSignal -and $stopSignal.Stop) {
                $result.Cancelled = $true
                $drain = $true
                break
            }
            if ($pool) {
                while ($pending.Count -lt $workers * 2 -and $next -lt $tasks.Count) {
                    $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
                    $task = $tasks[$next]
                    $job = $pool.Submit(${contentIndexWorkerScript}.ToString(), @($contentIndexFiles, $task.Start, $task.Count, $search.Regex, $max, $search.TextRegex,
                            $search.ScanMode, $cache, $filter.Include, $filter.Exclude, $excludePlace, $stopSignal, $gate, $next, $checkEvery, $overflowDir))
                    $job.Done = $task.Start + $task.Count
                    $job.Slot = $next
                    $pending.Enqueue($job)
                    $next++
                }
                $received = $pending.Dequeue()
                $output = $pool.Receive($received)
                if ($output[0].Timeout -and $null -eq $output[0].Hits) { $result.HitTotal = -1; throw $timeoutMessage }
                $found = $output[0]
                $slot = $received.Slot
                $done = $received.Done
            } else {
                $task = $tasks[$next]
                $slot = $next
                $next++
                $done = $task.Start + $task.Count
                $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
                try {
                    $found = searchContentIndexFiles $contentIndexFiles $task.Start $task.Count $search.Regex $max $search.TextRegex $search.ScanMode $cache $filter.Include $filter.Exclude $excludePlace $stopSignal $gate $slot $checkEvery $overflowDir
                } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                    $result.HitTotal = -1
                    throw $timeoutMessage
                }
            }
            $newHits = $found.Hits
            $newSkipped = $found.Skipped
            # 作業の途中で止められたときは、その作業のヒットを結果にも進み具合にも入れない（途中までの作業のため）
            if ($null -ne $stopSignal -and $stopSignal.Stop) {
                $result.Cancelled = $true
                $drain = $true
                break
            }
            $skipped.AddRange([string[]]@($newSkipped))
            if ($found.OverflowFailed) { $result.OverflowFailed = $true }
            $result.HitTotal += $found.Count
            # 表に入りきらない行は、止めずに溢れの head のファイルに書く（作業のファイルより前の順）
            $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
            if ($max -ge 0 -and $newHits.Count -gt $max) {
                $extra = $newHits.GetRange($max, $newHits.Count - $max)
                $newHits.RemoveRange($max, $newHits.Count - $max)
                if ($overflowDir -ne "" -and !$result.OverflowFailed) {
                    $headWriter = newSearchOverflowWriter $overflowDir "$slot.head.tsv"
                    $written = $false
                    if ($null -ne $headWriter) {
                        try {
                            $written = writeSearchOverflowHits $headWriter $extra
                        } finally {
                            try { $headWriter.Dispose() } catch { $written = $false }
                        }
                    }
                    if (!$written) { $result.OverflowFailed = $true }
                }
            }
            if ($limit -gt 0 -and $result.HitTotal -gt $limit) { $result.Truncated = $true }
            $hits.AddRange($newHits)
            if ($onProgress) {
                & $onProgress $done $contentIndexFiles.Count $newHits $result.HitTotal
            }
            if ($found.Timeout) {
                # 上限を超えた後の照合の時間切れ: 表は埋まっているので、総数だけが数えられない。表に入る分の途中で切れたなら、これまでどおり中止にする
                if (!$result.Truncated) { $result.HitTotal = -1; throw $timeoutMessage }
                $result.CountTimeout = $true
                break
            }
        }
    } finally {
        if ($null -ne $gate) { $gate[0] = 1 }
        foreach ($job in $pending) {
            if ($drain) {
                # 各作業はファイル 1 件ごとに Stop を見て早く終わる。受け取って捨てれば、プールを次の検索で使い回せる
                try { [void]$pool.Receive($job) } catch { }
            } else {
                $pool.Cancel($job)
            }
        }
        if ($ownPool) { $ownPool.Close() }
    }
    if ($result.Cancelled -or $result.CountTimeout) { $result.HitTotal = -1 }
    return $result
}


function getContentIndexFiles {
    # 検索対象の本文インデックスのファイルを集め、@{ Folders; ContentIndexFiles } を返す。
    #   folders: 検索対象インデックスのフォルダ（文字列。フォルダ以下すべて）、または
    #            @{ Root（インデックスのフォルダ）; RelPath（その中のフォルダ。空は Root 自身）; Recurse（$false は直下のファイルだけ） }
    #   Folders: フォルダごとの @{ Path; Root（フルパス）; Exists; Count }
    #   ContentIndexFiles  : findContentIndexFiles の結果をつないだもの（入れ子のフォルダを選んでも重複しない）
    #   onProgress: 数えた件数を知らせる { param($count) }
    #   stopSignal: Stop を持つもの（検索の要求。$null は止めない）。Stop が立ったら、検索対象のフォルダ 1 つごとに止め、ここまでの結果を返す
    param (
        [object[]]$folders = @($workspace.IndexDir),
        [scriptblock]$onProgress = $null,
        $stopSignal = $null
    )

    $folderInfo = New-Object System.Collections.Generic.List[object]
    $contentIndexFiles = New-Object System.Collections.Generic.List[hashtable]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($target in $folders) {
        if ($null -ne $stopSignal -and $stopSignal.Stop) { break }
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