# gram 索引（2-gram の索引。決まりは search/gram_index.ps1）の読み書き（状態層）。
# 索引はインデックス名ごとのフォルダ（$workspace.GramDir\<インデックス名>）に、区画のファイル（g<番号>.seg）と目録（toc.tsv）を置く。
# 区画は書き終えてから目録に足すため、途中で止まっても前の目録のままで使える。
# 読む側は共有モード（ReadWrite|Delete）で開くため、書く側が古い区画を消すのを妨げない。消せなくても（開かれている）エラーにしない。
# 消すのは次のインデックス作成の始め（目録に載っていない区画）で、まとめ直して外した区画は 1 世代だけ残す（目録の old）。

# 本文インデックスのファイルの 2-gram を並列に数えるときのスレッドの数の上限（system_index と同じ理由）と、1 つの作業にまとめる量
${gramWorkerMax} = 4
${gramTaskBytes} = 8MB
${gramTaskFiles} = 200

function openGramShared {
    # 区画・目録を共有モード（ReadWrite|Delete）で開く。bufferSize 1 は読み込みの先読みをしない（位置を指して少しずつ読むため）
    param (
        [string]$path,
        [int]$bufferSize = 1
    )

    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    return [System.IO.FileStream]::new((toLongPath $path), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share, $bufferSize)
}

function readGramBytes {
    # stream の position から count バイト読む（足りなければ例外）
    param (
        $stream,
        [long]$position,
        [int]$count
    )

    $buffer = New-Object 'byte[]' $count
    if ($count -eq 0) {
        return , $buffer
    }
    [void]$stream.Seek($position, [System.IO.SeekOrigin]::Begin)
    $done = 0
    while ($done -lt $count) {
        $read = $stream.Read($buffer, $done, $count - $done)
        if ($read -le 0) {
            throw "区画のファイルが途中で終わっています"
        }
        $done += $read
    }
    return , $buffer
}

function readGramArray {
    # stream の position から、要素 count 個（1 つ elementBytes バイト）の配列を読む。type は 'int' 'long' 'uint32'
    param (
        $stream,
        [long]$position,
        [int]$count,
        [string]$type,
        [int]$elementBytes
    )

    $values = switch ($type) {
        'int' { , (New-Object 'int[]' $count) }
        'long' { , (New-Object 'long[]' $count) }
        default { , (New-Object 'uint32[]' $count) }
    }
    if ($count -gt 0) {
        $bytes = readGramBytes $stream $position ($count * $elementBytes)
        [System.Buffer]::BlockCopy($bytes, 0, $values, 0, $count * $elementBytes)
    }
    return , $values
}

function writeGramArray {
    # 配列 array の startIndex から count 個を stream に書く（4MB ずつ）
    param (
        $stream,
        $array,
        [int]$elementBytes,
        [long]$startIndex,
        [long]$count
    )

    $chunkElements = [int](4MB / $elementBytes)
    $buffer = New-Object 'byte[]' ($chunkElements * $elementBytes)
    $position = $startIndex
    $end = $startIndex + $count
    while ($position -lt $end) {
        $take = [int][Math]::Min($chunkElements, $end - $position)
        [System.Buffer]::BlockCopy($array, [int]($position * $elementBytes), $buffer, 0, $take * $elementBytes)
        $stream.Write($buffer, 0, $take * $elementBytes)
        $position += $take
    }
}

function fillGramInts {
    # int 配列 array の start から count 個を value にする（最初の 1 つに入れて、Buffer.BlockCopy を倍々に使って広げる）
    param (
        [int[]]$array,
        [int]$start,
        [int]$count,
        [int]$value
    )

    if ($count -le 0) {
        return
    }
    $array[$start] = $value
    $filled = 1
    while ($filled -lt $count) {
        $copyCount = [Math]::Min($filled, $count - $filled)
        [System.Buffer]::BlockCopy($array, $start * 4, $array, ($start + $filled) * 4, $copyCount * 4)
        $filled += $copyCount
    }
}

function writeGramSegment {
    # 区画のファイルを書く。一時ファイルに書き終えてから名前を付けるため、途中で止まっても区画としては見えない。
    #   files     : 名前（RelPath。大文字と小文字を区別せず昇順）の順の @{ RelPath; Ticks; Size }
    #   gramArrays: files と同じ順の、ファイルごとの 2-gram の値（uint32[]。重複なし）
    # @{ Length; Postings } を返す
    param (
        [string]$path,
        [object[]]$files,
        [object[]]$gramArrays
    )

    $count = $files.Count
    $ticks = New-Object 'long[]' $count
    $sizes = New-Object 'long[]' $count
    $total = 0
    for ($i = 0; $i -lt $count; $i++) {
        $ticks[$i] = [long]$files[$i].Ticks
        $sizes[$i] = [long]$files[$i].Size
        $total += $gramArrays[$i].Length
    }
    $nameBytes = [System.Text.Encoding]::UTF8.GetBytes(([string[]]@($files | ForEach-Object { [string]$_.RelPath }) -join "`n"))

    # (2-gram, ファイルの番号) の組を、2-gram の順に並べる
    $grams = New-Object 'uint32[]' $total
    $owners = New-Object 'int[]' $total
    $position = 0
    $distinct = New-Object 'System.Collections.Generic.HashSet[uint32]'
    for ($i = 0; $i -lt $count; $i++) {
        $length = $gramArrays[$i].Length
        if ($length -gt 0) {
            [System.Array]::Copy($gramArrays[$i], 0, $grams, $position, $length)
            fillGramInts $owners $position $length $i
            $distinct.UnionWith([uint32[]]$gramArrays[$i])
            $position += $length
        }
    }
    [System.Array]::Sort($grams, $owners)
    $gramCount = $distinct.Count
    $distinct = $null

    $dictionary = New-Object 'uint32[]' $gramCount
    $offsets = New-Object 'int[]' ($gramCount + 1)
    $group = -1
    $last = [uint32]0
    for ($i = 0; $i -lt $total; $i++) {
        $value = $grams[$i]
        if ($group -lt 0 -or $value -ne $last) {
            $group++
            $dictionary[$group] = $value
            $offsets[$group] = $i
            $last = $value
        }
    }
    $offsets[$gramCount] = $total
    $grams = $null

    $layout = getGramSegmentLayout $count $nameBytes.Length $gramCount $total
    $longPath = toLongPath $path
    $tmp = "$longPath.tmp"
    $stream = [System.IO.FileStream]::new($tmp, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $header = convertToGramSegmentHeader $count $nameBytes.Length $gramCount $total
        $stream.Write($header, 0, $header.Length)
        writeGramArray $stream $ticks 8 0 $count
        writeGramArray $stream $sizes 8 0 $count
        $stream.Write($nameBytes, 0, $nameBytes.Length)
        writeGramArray $stream $dictionary 4 0 $gramCount
        writeGramArray $stream $offsets 4 0 ($gramCount + 1)
        writeGramArray $stream $owners 4 0 $total
        $footer = [System.BitConverter]::GetBytes([int]${gramSegmentFooter})
        $stream.Write($footer, 0, 4)
    } finally {
        $stream.Dispose()
    }
    [System.IO.File]::Move($tmp, $longPath)
    return @{ Length = $layout.Length; Postings = $total }
}

function openGramSegment {
    # 区画のファイルを開き、壊れていなければ @{ Stream; Header; Names; Ticks; Sizes; Path } を返す（Candidates は検索のとき足す）。
    # 無い・読めない・壊れている（長さが先頭と目録に合わない・末尾の印が違う）ときは $null
    param (
        [string]$path,
        [long]$tocLength
    )

    $stream = $null
    try {
        $stream = openGramShared $path
        $length = $stream.Length
        if ($length -lt (${gramSegmentHeaderBytes} + 4)) {
            $stream.Dispose()
            return $null
        }
        $header = convertFromGramSegmentHeader (readGramBytes $stream 0 ${gramSegmentHeaderBytes})
        $footer = [System.BitConverter]::ToInt32((readGramBytes $stream ($length - 4) 4), 0)
        if (!(testGramSegmentIntact $header $length $tocLength $footer)) {
            $stream.Dispose()
            return $null
        }
        $files = $header.FileCount
        $layout = $header.Layout
        $names = @()
        if ($files -gt 0) {
            $text = [System.Text.Encoding]::UTF8.GetString((readGramBytes $stream $layout.NamesPos $header.NameBytes))
            $names = [string[]]$text.Split([char]"`n")
            if ($names.Length -ne $files) {
                $stream.Dispose()
                return $null
            }
        } else {
            $names = [string[]]@()
        }
        return @{
            Stream = $stream; Header = $header; Path = $path; Names = $names
            Ticks = (readGramArray $stream $layout.TicksPos $files 'long' 8)
            Sizes = (readGramArray $stream $layout.SizesPos $files 'long' 8)
        }
    } catch {
        if ($stream) {
            $stream.Dispose()
        }
        return $null
    }
}

function closeGramSegments {
    param (
        [object[]]$segments
    )

    foreach ($segment in $segments) {
        if ($segment -and $segment.Stream) {
            $segment.Stream.Dispose()
        }
    }
}

function findGramInSegment {
    # 区画の 2-gram の表（昇順）から値 gram を二分探索し、位置を返す（無ければ -1）。表は読み込まず、位置を指して 4 バイトずつ読む
    param (
        $segment,
        [uint32]$gram
    )

    $low = 0
    $high = $segment.Header.GramCount - 1
    $base = $segment.Header.Layout.GramsPos
    while ($low -le $high) {
        $middle = $low + (($high - $low) -shr 1)
        $found = [System.BitConverter]::ToUInt32((readGramBytes $segment.Stream ($base + 4L * $middle) 4), 0)
        if ($found -eq $gram) {
            return $middle
        }
        if ($found -lt $gram) {
            $low = $middle + 1
        } else {
            $high = $middle - 1
        }
    }
    return -1
}

function getGramCandidates {
    # 区画の中で、検索語のすべての 2-gram（grams）を含むファイルの番号（Names の位置）の集まりを HashSet[int] で返す
    param (
        $segment,
        [uint32[]]$grams
    )

    $starts = New-Object 'int[]' $grams.Length
    $counts = New-Object 'int[]' $grams.Length
    for ($i = 0; $i -lt $grams.Length; $i++) {
        $k = findGramInSegment $segment $grams[$i]
        if ($k -lt 0) {
            return New-Object 'System.Collections.Generic.HashSet[int]'
        }
        $pair = readGramArray $segment.Stream ($segment.Header.Layout.OffsetsPos + 4L * $k) 2 'int' 4
        $starts[$i] = $pair[0]
        $counts[$i] = $pair[1] - $pair[0]
    }
    # 少ないものから読む（共通部分が空になればそれ以上読まない）
    [System.Array]::Sort($counts, $starts)
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $grams.Length; $i++) {
        if ($counts[$i] -le 0) {
            return New-Object 'System.Collections.Generic.HashSet[int]'
        }
        $rows.Add((readGramArray $segment.Stream ($segment.Header.Layout.PostingsPos + 4L * $starts[$i]) $counts[$i] 'int' 4))
    }
    return , (intersectGramRows $rows.ToArray())
}

function readGramToc {
    # 目録を読む（convertFromGramToc の形）。無い・読めない・形が違うときは $null
    param (
        [string]$dir
    )

    $path = "$($dir.TrimEnd('\'))\${gramTocFileName}"
    if (![System.IO.File]::Exists((toLongPath $path))) {
        return $null
    }
    try {
        $stream = openGramShared $path 4096
        try {
            $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
            return convertFromGramToc ($reader.ReadToEnd() -split "\r?\n")
        } finally {
            $stream.Dispose()
        }
    } catch {
        return $null
    }
}

function writeGramToc {
    # 目録を書く。一時ファイルに書き終えてから置き換える（途中で止まっても前の目録が残る）
    param (
        [string]$dir,
        $toc
    )

    $path = toLongPath "$($dir.TrimEnd('\'))\${gramTocFileName}"
    $tmp = "$path.tmp"
    [System.IO.File]::WriteAllText($tmp, ((convertToGramToc $toc) -join "`r`n") + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
    if ([System.IO.File]::Exists($path)) {
        [System.IO.File]::Replace($tmp, $path, [NullString]::Value)
    } else {
        [System.IO.File]::Move($tmp, $path)
    }
}

function removeGramFilesNotInToc {
    # 目録（seg・old）に載っていない区画と、書きかけの一時ファイルを消す。消せなくても（開かれている）エラーにしない。
    # toc が $null（目録が無い・読めない）なら、区画をすべて消す
    param (
        [string]$dir,
        $toc
    )

    $longDir = toLongPath $dir
    if (![System.IO.Directory]::Exists($longDir)) {
        return
    }
    $keep = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($toc) {
        foreach ($segment in $toc.Segments) { [void]$keep.Add($segment.Name) }
        foreach ($name in $toc.Old) { [void]$keep.Add($name) }
    }
    foreach ($pattern in @(${gramSegmentFilePattern}, "*.tmp")) {
        foreach ($file in [System.IO.Directory]::GetFiles($longDir, $pattern)) {
            if (!$keep.Contains([System.IO.Path]::GetFileName($file))) {
                try { [System.IO.File]::Delete($file) } catch { }
            }
        }
    }
}

function getFileGramValues {
    # 本文インデックスのファイル 1 つの 2-gram を、重複なし・昇順の uint32[] で返す。
    # 読む前と後でファイルの更新日時・大きさが、数えたとき（ticks・size）と同じときだけ返す（読む間に書き換わったら、
    # 古い日時に新しい中身の 2-gram を付けてしまうため）。読めない・大きすぎるときは $null（索引に載せず、必ず照合する）
    param (
        [string]$path,
        [long]$ticks,
        [long]$size
    )

    try {
        if ($size -gt ${contentIndexSearchMaxBytes}) {
            return $null
        }
        $info = [System.IO.FileInfo]::new((toLongPath $path))
        if ($info.Length -ne $size -or $info.LastWriteTimeUtc.Ticks -ne $ticks) {
            return $null
        }
        $set = New-Object 'System.Collections.Generic.HashSet[uint32]'
        addTextGrams $set (getContentIndexBodyText (readContentIndexText $path))
        $after = [System.IO.FileInfo]::new((toLongPath $path))
        if ($after.Length -ne $size -or $after.LastWriteTimeUtc.Ticks -ne $ticks) {
            return $null
        }
        $values = New-Object 'uint32[]' $set.Count
        $set.CopyTo($values)
        [System.Array]::Sort($values)
        return , $values
    } catch {
        return $null
    }
}

function getFilesGramValues {
    # 本文インデックスのファイル（items。@{ Path; Ticks; Size }）の 2-gram を数える（並列）。items と同じ順の配列を返す
    # （数えられなかったものは $null）。shouldStop が $true を返したら、始めていない分を数えずに $null を返す
    param (
        [object[]]$items,
        [int]$workers = 0,
        [scriptblock]$shouldStop = $null
    )

    $results = New-Object 'object[]' $items.Count
    if ($items.Count -eq 0) {
        return , $results
    }
    # 1 つの作業にまとめるファイルの範囲
    $tasks = New-Object System.Collections.Generic.List[hashtable]
    $start = 0
    $bytes = 0L
    for ($i = 0; $i -lt $items.Count; $i++) {
        $bytes += [long]$items[$i].Size
        if ($bytes -ge ${gramTaskBytes} -or ($i - $start + 1) -ge ${gramTaskFiles} -or $i -eq $items.Count - 1) {
            $tasks.Add(@{ Start = $start; Count = $i - $start + 1 })
            $start = $i + 1
            $bytes = 0L
        }
    }
    if ($workers -le 0) {
        $workers = [Math]::Min([Environment]::ProcessorCount, ${gramWorkerMax})
    }
    $workers = [Math]::Min($workers, $tasks.Count)
    if ($workers -le 1) {
        foreach ($task in $tasks) {
            if ($shouldStop -and (& $shouldStop)) {
                break
            }
            for ($i = $task.Start; $i -lt $task.Start + $task.Count; $i++) {
                $results[$i] = getFileGramValues $items[$i].Path $items[$i].Ticks $items[$i].Size
            }
        }
        return , $results
    }

    # 各スレッドには必要な関数・値だけを読み込む。インデックス作成の処理のため、スレッドの優先度を下げる
    $state = newWorkerState @("getFileGramValues", "addTextGrams", "getContentIndexBodyText", "readContentIndexText", "toLongPath") @("contentIndexSearchMaxBytes")
    $pool = [WorkerPool]::new($workers, $state, $Host, "BelowNormal")
    $pending = New-Object System.Collections.Generic.Queue[hashtable]
    $jobScript = {
        param ($paths, $ticks, $sizes)
        $ErrorActionPreference = "Stop"
        $found = New-Object 'object[]' $paths.Count
        for ($i = 0; $i -lt $paths.Count; $i++) {
            $found[$i] = getFileGramValues $paths[$i] $ticks[$i] $sizes[$i]
        }
        @{ Grams = $found }
    }.ToString()
    try {
        $next = 0
        while ($next -lt $tasks.Count -or $pending.Count -gt 0) {
            while ($pending.Count -lt $workers * 2 -and $next -lt $tasks.Count) {
                if ($shouldStop -and (& $shouldStop)) {
                    $next = $tasks.Count
                    break
                }
                $task = $tasks[$next]
                $range = $items[$task.Start..($task.Start + $task.Count - 1)]
                $paths = [string[]]@($range | ForEach-Object { $_.Path })
                $ticks = [long[]]@($range | ForEach-Object { $_.Ticks })
                $sizes = [long[]]@($range | ForEach-Object { $_.Size })
                $job = $pool.Submit($jobScript, @($paths, $ticks, $sizes))
                $pending.Enqueue(@{ Job = $job; Start = $task.Start })
                $next++
            }
            if ($pending.Count -eq 0) {
                break
            }
            $entry = $pending.Dequeue()
            $output = $pool.Receive($entry.Job)
            if ($output.Count -gt 0) {
                $found = $output[0].Grams
                for ($i = 0; $i -lt $found.Count; $i++) {
                    $results[$entry.Start + $i] = $found[$i]
                }
            }
        }
    } finally {
        foreach ($entry in $pending) {
            $pool.Cancel($entry.Job)
        }
        $pool.Close()
    }
    return , $results
}

function updateGramIndexOf {
    # インデックス 1 つ（name）の gram 索引を、本文インデックスのファイルの今の状態に合わせる。
    # 区画に載っていないファイル（新しい・更新日時か大きさが変わった）の 2-gram を数え、区画として足す。
    # 区画が多くなったら（testGramMergeNeeded）、すべてのファイルを数え直して作り直す。
    # 目録に載っていない区画は始めに消す。shouldStop が $true を返したら、目録を書かずに止める（書いた区画は次の回に消える）。
    # @{ Added（足した区画の数）; Files（数えたファイルの数）; Merged（作り直したか）; Stopped } を返す
    param (
        [string]$indexRoot,
        [string]$gramRoot,
        [string]$name,
        [scriptblock]$shouldStop = $null,
        [int]$workers = 0
    )

    $result = @{ Added = 0; Files = 0; Merged = $false; Stopped = $false }
    $dir = "$($gramRoot.TrimEnd('\'))\$name"
    $items = [object[]](findContentIndexFiles $indexRoot $name $true)
    if ($items.Count -eq 0) {
        [void](removeDirectoryUnder $dir $gramRoot)
        return $result
    }
    $toc = readGramToc $dir
    removeGramFilesNotInToc $dir $toc

    # 今の区画（新しい順）。ひとつでも開けない・壊れているときは、目録ごと捨てて作り直す（開けない区画の行を、他の区画の行が隠さないため）
    $segments = New-Object System.Collections.Generic.List[object]
    try {
        if ($toc) {
            for ($i = $toc.Segments.Count - 1; $i -ge 0; $i--) {
                $segment = openGramSegment "$dir\$($toc.Segments[$i].Name)" $toc.Segments[$i].Length
                if ($null -eq $segment) {
                    closeGramSegments $segments.ToArray()
                    $segments.Clear()
                    $discardedNext = $toc.Next
                    $toc = $null
                    removeGramFilesNotInToc $dir $null
                    break
                }
                $segment.Name = $toc.Segments[$i].Name
                $segments.Add($segment)
            }
        }
        # まとめた後の数の基準は、生きている組の数。古い区画に残った重複（変わる前の版）と消えたファイルの分を、
        # 「今のファイルの数 / 全区画のファイルの数」の割合で除く
        $postings = 0L
        $listedFiles = 0L
        foreach ($segment in $segments) {
            $postings += $segment.Header.PostingCount
            $listedFiles += $segment.Header.FileCount
        }
        if ($listedFiles -gt $items.Count) {
            $postings = [long][Math]::Floor([double]$postings * $items.Count / $listedFiles)
        }
        $existing = if ($toc) { @($toc.Segments | ForEach-Object { $_.Name }) } else { @() }
        $merge = ($toc -and (testGramMergeNeeded ($segments.Count + 1) $postings))
        $todo = if ($merge) { $items } else { getGramFilesToAdd $items $segments.ToArray() }
        if ($todo.Count -eq 0) {
            return $result
        }
        # 作り直すときは、まとめた区画が新しい版になるよう、今の区画は使わない
        $comparer = [System.StringComparer]::OrdinalIgnoreCase
        $todo = [object[]]@($todo)
        $keys = [string[]]@($todo | ForEach-Object { [string]$_.RelPath })
        [System.Array]::Sort($keys, $todo, $comparer)

        if (!(newDirectoryUnder $dir $gramRoot)) {
            throw "gram 索引の置き場の外か、途中にリンクがあるため、作れません: $dir"
        }
        $values = getFilesGramValues $todo $workers $shouldStop
        if ($shouldStop -and (& $shouldStop)) {
            $result.Stopped = $true
            return $result
        }
        # 数えられたファイルだけを区画に載せる（載せなかったものは必ず照合される）
        $listed = New-Object System.Collections.Generic.List[object]
        $arrays = New-Object System.Collections.Generic.List[object]
        for ($i = 0; $i -lt $todo.Count; $i++) {
            if ($null -ne $values[$i]) {
                $listed.Add($todo[$i])
                $arrays.Add($values[$i])
            }
        }
        $values = $null
        $next = if ($toc) { $toc.Next } elseif ($discardedNext) { $discardedNext } else { 1 }
        $newSegments = New-Object System.Collections.Generic.List[hashtable]
        if ($listed.Count -gt 0) {
            $counts = New-Object 'int[]' $arrays.Count
            for ($i = 0; $i -lt $arrays.Count; $i++) { $counts[$i] = $arrays[$i].Length }
            $batches = splitGramBatches $counts
            foreach ($batch in $batches) {
                $segmentName = getGramSegmentName $next
                $next++
                $written = writeGramSegment "$dir\$segmentName" $listed.GetRange($batch.Start, $batch.Count).ToArray() $arrays.GetRange($batch.Start, $batch.Count).ToArray()
                $newSegments.Add(@{ Name = $segmentName; Length = $written.Length })
                $result.Added++
            }
        }
        $newToc = @{ Segments = New-Object System.Collections.Generic.List[hashtable]; Old = New-Object System.Collections.Generic.List[string]; Next = $next }
        if ($toc -and !$merge) {
            foreach ($segment in $toc.Segments) { $newToc.Segments.Add($segment) }
        } elseif ($toc) {
            foreach ($segmentName in $existing) { $newToc.Old.Add($segmentName) }
        }
        foreach ($segment in $newSegments) { $newToc.Segments.Add($segment) }
        writeGramToc $dir $newToc
        $result.Files = $listed.Count
        $result.Merged = [bool]$merge
    } finally {
        closeGramSegments $segments.ToArray()
    }
    return $result
}

function updateGramIndexes {
    # インデックスのフォルダ（indexRoot）の直下にあるインデックスすべての gram 索引を合わせ、無くなったインデックスの索引を消す。
    # インデックス 1 つの失敗は、そのインデックスを飛ばして続ける（索引が古いだけで検索は安全側に倒れる）。
    # @{ Added; Files; Failed（落ちたインデックスの @{ Name; Reason }）} を返す。shouldStop が $true を返したら止める
    param (
        [string]$indexRoot = $workspace.IndexDir,
        [string]$gramRoot = $workspace.GramDir,
        [scriptblock]$shouldStop = $null,
        [int]$workers = 0
    )

    $summary = @{ Added = 0; Files = 0; Failed = (New-Object System.Collections.Generic.List[hashtable]) }
    $longIndex = toLongPath $indexRoot.TrimEnd("\")
    $names = @()
    if ([System.IO.Directory]::Exists($longIndex)) {
        $names = getGramIndexNames ([string[]]@([System.IO.Directory]::GetDirectories($longIndex) | ForEach-Object { [System.IO.Path]::GetFileName($_) }))
    }
    $known = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $names) {
        [void]$known.Add($name)
        if ($shouldStop -and (& $shouldStop)) {
            return $summary
        }
        try {
            $one = updateGramIndexOf $indexRoot $gramRoot $name $shouldStop $workers
            $summary.Added += $one.Added
            $summary.Files += $one.Files
        } catch {
            $summary.Failed.Add(@{ Name = $name; Reason = $_.Exception.Message })
        }
    }
    # 無くなったインデックスの索引
    $longGram = toLongPath $gramRoot.TrimEnd("\")
    if ([System.IO.Directory]::Exists($longGram)) {
        foreach ($dir in [System.IO.Directory]::GetDirectories($longGram)) {
            $name = [System.IO.Path]::GetFileName($dir)
            if (!$known.Contains($name)) {
                try {
                    [void](removeDirectoryUnder (fromLongPath $dir) $gramRoot)
                } catch {
                    $summary.Failed.Add(@{ Name = $name; Reason = $_.Exception.Message })
                }
            }
        }
    }
    return $summary
}

function removeGramIndexOf {
    # インデックス name を消したとき、gram 索引の同じフォルダを消す。置き場の下で途中にリンクが無いときだけ消す。
    # 消せなくても（開かれている）例外のまま返す。呼び出し側が続けるかを決める
    param (
        [string]$name,
        [string]$gramRoot = $workspace.GramDir
    )

    return removeDirectoryUnder "$($gramRoot.TrimEnd('\'))\$name" $gramRoot
}
