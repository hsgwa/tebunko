# 検索用の本文インデックスのファイル（content_index_format.ps1）の読み書き（状態層）。
# 本文インデックスのファイルは work\content_index の中のフォルダごと・元のファイルの拡張子ごとに、大きさで分けて置く（content_index.xlsx.001.tsv など）。UTF-16LE（BOM 付き）で書く
# （UTF-8 より文字列への変換が速い。日本語が多いと大きさはほとんど変わらない）。

function writeContentIndexFile {
    # 本文インデックスのファイルを書く。一時ファイル（<名前>.tmp）に書き終えてから置き換えるため、途中で止まっても前のファイルが残る
    param (
        [string]$path,
        [string]$text
    )

    $longPath = toLongPath $path
    $tmp = "$longPath.tmp"
    [System.IO.File]::WriteAllText($tmp, $text, [System.Text.UnicodeEncoding]::new($false, $true))
    if ([System.IO.File]::Exists($longPath)) {
        # 3 つ目の引数（控え）に $null を渡すと PowerShell が空の文字列にして例外になるため、NullString を渡す
        [System.IO.File]::Replace($tmp, $longPath, [NullString]::Value)
    } else {
        [System.IO.File]::Move($tmp, $longPath)
    }
}


function readContentIndexText {
    # 本文インデックスのファイルを文字列で読む（インデックス作成中の置き換えと同時に読めるよう、共有モードは ReadWrite|Delete）
    param (
        [string]$path
    )

    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = [System.IO.FileStream]::new((toLongPath $path), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::Unicode, $true)
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
}


function readContentIndexRange {
    # 本文インデックスのファイル（UTF-16LE・BOM 付き）の、文字の位置 start（含む）から end（含まない）までを start から数えて、
    # firstLine 行目の先頭（firstLine が 1 以下なら start）から lines 個目の改行までを文字列で返す（それより前の行は持たない）。
    # BOM が無い（UTF-16LE でない）ときは $null（呼び出し側が全体を読む）。
    # checkpoints（ハッシュテーブル）を渡すと、通った 200 行ごとの行の先頭（行の番号 → 文字の位置）を足す
    # （start の行の番号が baseLine。次に読むとき、目当ての行の手前の位置から読めるようにする）
    param (
        [string]$path,
        [int]$start,
        [int]$end,
        [int]$firstLine,
        [int]$lines,
        [int]$baseLine = 1,
        $checkpoints = $null
    )

    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = [System.IO.FileStream]::new((toLongPath $path), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        $bom = New-Object byte[] 2
        if ($stream.Read($bom, 0, 2) -ne 2 -or $bom[0] -ne 0xFF -or $bom[1] -ne 0xFE) { return $null }
        [void]$stream.Seek(2L + 2L * $start, [System.IO.SeekOrigin]::Begin)
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.UnicodeEncoding]::new($false, $false), $false)
        $builder = New-Object System.Text.StringBuilder
        $buffer = New-Object char[] 65536
        $remaining = $end - $start
        $newlines = 0
        $skip = [Math]::Max(0, $firstLine - 1)
        $started = ($skip -eq 0)
        while ($remaining -gt 0 -and $newlines -lt $lines) {
            $consumed = ($end - $start) - $remaining
            $got = $reader.Read($buffer, 0, [Math]::Min($buffer.Length, $remaining))
            if ($got -le 0) { break }
            $keepFrom = if ($started) { 0 } else { -1 }
            $from = 0
            while ($from -lt $got) {
                $found = [Array]::IndexOf($buffer, [char]10, $from, $got - $from)
                if ($found -lt 0) { break }
                $newlines++
                $from = $found + 1
                if ($null -ne $checkpoints) {
                    $nextLine = $baseLine + $newlines
                    if (($nextLine - 1) % 200 -eq 0) { $checkpoints[$nextLine] = $start + $consumed + $from }
                }
                if (!$started -and $newlines -eq $skip) { $started = $true; $keepFrom = $from }
            }
            # 読み飛ばす行のチャンクは捨てる（行数だけ数える）
            if ($keepFrom -ge 0 -and $keepFrom -lt $got) { [void]$builder.Append($buffer, $keepFrom, $got - $keepFrom) }
            $remaining -= $got
        }
        return $builder.ToString()
    } finally {
        $stream.Dispose()
    }
}


function testIndexBookDir {
    # インデックスの中のフォルダが、元のファイルごとのフォルダ（<ファイル名.xlsx>\<場所>.tsv。本文インデックスのファイルに入れる前の TSV の置き場所）か。
    # 名前だけでは、名前が .xlsx などで終わる本物のフォルダ（元のフォルダの名前をそのまま使う）と区別できないため、中身も見る:
    #   ・名前が Office・テキストの拡張子で終わる（indexBookDirPattern）
    #   ・サブフォルダも本文インデックスのファイル（content_index.<拡張子>.tsv）も無い
    #   ・withTsv なら、TSV が 1 つ以上ある（取り込んだが中身が空のファイルのフォルダは、本文インデックスのファイルに入れるものが無い）
    # 読めないフォルダは $false（本文インデックスのファイルに入れる・消す対象にしない）
    param (
        [string]$dir,
        [bool]$withTsv = $true
    )

    if ([System.IO.Path]::GetFileName($dir.TrimEnd("\")) -notmatch ${indexBookDirPattern}) {
        return $false
    }
    $long = toLongPath $dir
    try {
        # 途中で戻っても、調べたフォルダを掴んだまま残さない（testAnyEntry・配列で受ける GetFiles）
        if (testAnyEntry ([System.IO.Directory]::EnumerateDirectories($long))) {
            return $false
        }
        $hasTsv = $false
        foreach ($file in [System.IO.Directory]::GetFiles($long, "*.tsv")) {
            if ([System.IO.Path]::GetFileName($file) -like ${contentIndexFilePattern}) {
                return $false
            }
            $hasTsv = $true
        }
        return ($hasTsv -or !$withTsv)
    } catch {
        return $false
    }
}

function getIndexFolderBooks {
    # 今の形式のインデックスのフォルダ 1 つ（直下の <ファイル名.xlsx>\<場所>.tsv）から、本文インデックスのファイルに入れる元のファイルの並びを作る。
    # 並びは今の検索結果と同じ順（前の名前の付け方（フォルダのパス + encodeIndexPlace 場所 + ".tsv"）で並べたときと同じ順に
    # なるよう、実際のファイル名（英語の固定名・符号化したシート名）から場所を戻し、その前の名前の付け方の文字列を並べ替えの鍵にする）。
    # @{ Name; Places（@{ Place; Path } の並び） } の並びを返す
    param (
        [string]$folder
    )

    $longDir = toLongPath $folder
    $entries = New-Object System.Collections.Generic.List[hashtable]
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($sub in [System.IO.Directory]::EnumerateDirectories($longDir)) {
        if (!(testIndexBookDir $sub)) {
            continue
        }
        $name = [System.IO.Path]::GetFileName($sub)
        foreach ($path in [System.IO.Directory]::GetFiles($sub, "*.tsv", [System.IO.SearchOption]::TopDirectoryOnly)) {
            # 長いパス（\\?\ 付き）は [System.IO.Path]::GetDirectoryName で崩れることがあるため、
            # ここで分かっている $sub・$name をそのまま使い、$path から作り直さない。
            $place = convertIndexFileNameToPlace ([System.IO.Path]::GetFileNameWithoutExtension($path))
            $legacyName = "{0}.tsv" -f (encodeIndexPlace $place)
            # Join-Path は \\?\ 付きの長いパスで「drive」が null というエラーになるため、
            # プロバイダーを介さない [System.IO.Path]::Combine で組み立てる。
            $key = [System.IO.Path]::Combine($sub, $legacyName)
            $entries.Add(@{ Name = $name; Place = $place; Path = $path })
            $keys.Add($key)
        }
    }
    $keyArray = $keys.ToArray()
    $entryArray = $entries.ToArray()
    [System.Array]::Sort($keyArray, $entryArray, [System.StringComparer]::CurrentCultureIgnoreCase)

    $books = New-Object System.Collections.Generic.List[hashtable]
    $current = $null
    foreach ($entry in $entryArray) {
        if (!$current -or $current.Name -ne $entry.Name) {
            $current = @{ Name = $entry.Name; Places = New-Object System.Collections.Generic.List[hashtable] }
            $books.Add($current)
        }
        $current.Places.Add(@{ Place = $entry.Place; Path = $entry.Path })
    }
    return , $books
}


function convertFolderToContentIndex {
    # 今の形式のインデックスのフォルダ 1 つ（直下の <ファイル名.xlsx>\<場所>.tsv）から、拡張子ごと・番号ごとの本文インデックスのファイル
    # （destFolder\content_index.xlsx.001.tsv など）を書く。destFolder に前の本文インデックスのファイルがあれば、それとまぜる（planContentIndexParts）:
    #   ・TSV のある元のファイルは、TSV の中身で入れ替える（前の本文インデックスのファイルに無ければ、最後の番号の本文インデックスのファイルに足す。
    #     contentIndexFileMaxBytes 以上なら次の番号の本文インデックスのファイルを作る）
    #   ・removeBooks に挙げた元のファイルは外す（元のファイルが無くなった）
    #   ・それ以外の元のファイルは、前の本文インデックスのファイルのまま。変わらない本文インデックスのファイルは書き直さない
    # 元のファイルが無くなった本文インデックスのファイルは消す。removeTsv なら、本文インデックスのファイルを書き終えた後に、
    # 読み込んだ元のファイルのフォルダ（TSV）を消す（TSV は一時的な置き場で、残すとインデックスの容量が倍になるため）。
    # 書き終える前に止まっても、TSV か前の本文インデックスのファイルのどちらかに中身が残る。
    # @{ Books; Tsv; Chars; Files（本文インデックスのファイルの数）; Written（書き直した数）;
    #    Texts（そのフォルダのすべての本文インデックスのファイルの中身。システムインデックスを読み直さずに作るため） } を返す
    param (
        [string]$folder,
        [string]$destFolder,
        [string[]]$removeBooks = @(),
        [bool]$removeTsv = $false,
        [long]$maxBytes = ${contentIndexFileMaxBytes}
    )

    $books = getIndexFolderBooks $folder
    $tsvCount = 0
    $newBooks = New-Object System.Collections.Generic.List[hashtable]
    foreach ($book in $books) {
        foreach ($place in $book.Places) {
            $place.Text = [System.IO.File]::ReadAllText($place.Path)
            $tsvCount++
        }
        $newBooks.Add(@{ Name = $book.Name; Block = (convertBookToContentIndexBlock $book) })
    }
    $longDest = toLongPath $destFolder
    # 前の本文インデックスのファイル（番号の付いた名前のもの）を読む
    $parts = New-Object System.Collections.Generic.List[hashtable]
    $oldTexts = @{}
    if ([System.IO.Directory]::Exists($longDest)) {
        foreach ($old in [System.IO.Directory]::GetFiles($longDest, ${contentIndexFilePattern})) {
            $info = readContentIndexFileName ([System.IO.Path]::GetFileName($old))
            if ($null -eq $info) { continue }
            $text = readContentIndexText $old
            $oldTexts[(getContentIndexFileName $info.Extension $info.Part)] = $text
            $parts.Add(@{ Extension = $info.Extension; Part = $info.Part; Books = (splitContentIndexTextByBook $text) })
        }
    }

    $plan = planContentIndexParts $parts $newBooks $removeBooks $maxBytes
    $texts = New-Object System.Collections.Generic.List[string]
    $bookTotal = 0
    $written = 0
    $chars = 0L
    foreach ($part in $plan) {
        $name = getContentIndexFileName $part.Extension $part.Part
        $path = Join-Path $destFolder $name
        if ($part.Books.Count -eq 0) {
            if ([System.IO.File]::Exists((toLongPath $path))) { [System.IO.File]::Delete((toLongPath $path)) }
            continue
        }
        if ($part.Changed) {
            [void][System.IO.Directory]::CreateDirectory($longDest)
            $text = convertToContentIndexText $part.Books
            writeContentIndexFile $path $text
            $written++
        } else {
            $text = $oldTexts[$name]
        }
        $texts.Add($text)
        $bookTotal += $part.Books.Count
        $chars += $text.Length
    }
    if ($removeTsv) {
        foreach ($book in $books) {
            $bookDir = [System.IO.Path]::GetDirectoryName($book.Places[0].Path)
            [System.IO.Directory]::Delete($bookDir, $true)
        }
    }
    return @{ Books = $bookTotal; Tsv = $tsvCount; Chars = $chars; Files = $texts.Count; Written = $written; Texts = [string[]]$texts.ToArray() }
}

function updateFolderContentIndex {
    # インデックスのフォルダ 1 つで、置かれた TSV（追加・更新した元のファイル）を本文インデックスのファイルに入れ、TSV を消す。
    # removeBooks に挙げた元のファイル（無くなったもの）は本文インデックスのファイルから外す
    param (
        [string]$folder,
        [string[]]$removeBooks = @()
    )

    return convertFolderToContentIndex $folder $folder $removeBooks $true
}

function findContentIndexFiles {
    # インデックスのフォルダ以下の本文インデックスのファイルを列挙し、フォルダの順・フォルダの中は名前の順に並べて返す。
    #   root   : インデックスのフォルダ（相対パスの基準）
    #   relPath: その中のフォルダ（空は root 自身）
    #   recurse: $false なら、そのフォルダの本文インデックスのファイルだけ
    # 各要素は @{ Path（\\?\ 付き）; Root; RelDir（root からのフォルダ）; RelPath; Ticks; Size }
    param (
        [string]$root,
        [string]$relPath = "",
        [bool]$recurse = $true
    )

    $root = $root.TrimEnd("\")
    $relPath = $relPath.Trim("\")
    $dir = if ($relPath) { "${root}\${relPath}" } else { $root }
    $longDir = toLongPath $dir
    if (![System.IO.Directory]::Exists($longDir)) {
        return , @()
    }
    $option = if ($recurse) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }
    $found = [System.IO.DirectoryInfo]::new($longDir).GetFiles(${contentIndexFilePattern}, $option)
    $plainDir = fromLongPath $longDir
    $list = New-Object System.Collections.Generic.List[hashtable]
    foreach ($file in $found) {
        $full = $plainDir + $file.FullName.Substring($longDir.Length)
        $rel = $full.Substring($root.Length + 1)
        $list.Add(@{
            Path = $file.FullName; Root = $root; RelPath = $rel
            RelDir = [string][System.IO.Path]::GetDirectoryName($rel)
            Ticks = $file.LastWriteTimeUtc.Ticks; Size = $file.Length
        })
    }
    # フォルダの順、フォルダの中は本文インデックスのファイルの名前の順（現在のカルチャ・大文字と小文字を区別しない）
    $items = [hashtable[]]@($list | Sort-Object @{ Expression = { $_.RelDir } }, @{ Expression = { [System.IO.Path]::GetFileName($_.RelPath) } })
    return , $items
}


function findIndexFoldersWithBooks {
    # インデックスのフォルダ以下で、元のファイルごとのフォルダ（<ファイル名.xlsx>。本文インデックスのファイルに入れる前の TSV）が
    # 直下にあるフォルダを返す（インデックス作成が途中で止まったとき）。root 自身も含む
    param (
        [string]$root
    )

    $found = New-Object System.Collections.Generic.List[string]
    $longRoot = toLongPath $root.TrimEnd("\")
    if (![System.IO.Directory]::Exists($longRoot)) {
        return , $found.ToArray()
    }
    $dirs = @($longRoot) + @([System.IO.Directory]::GetDirectories($longRoot, "*", [System.IO.SearchOption]::AllDirectories))
    foreach ($dir in $dirs) {
        if ($dir -ne $longRoot -and (testIndexBookDir $dir $false)) { continue }
        if (testAnyEntry ([System.IO.Directory]::EnumerateDirectories($dir)) { param ($sub) testIndexBookDir $sub }) {
            $found.Add((fromLongPath $dir))
        }
    }
    return , $found.ToArray()
}


function publishIndexFolders {
    # インデックス作成で TSV を置いた・元のファイルが無くなったフォルダを、まとめて書き出す。フォルダごとに次を続けて行う:
    #   1. 本文インデックスのファイルを書く（前の本文インデックスのファイルとまぜ、無くなった元のファイルは外す）
    #   2. 元のファイルごとのフォルダの TSV を消す
    #   3. 書いた本文インデックスのファイルの中身から、そのフォルダのシステムインデックスの txt を作る（読み直さない）
    # txt の状態（反映待ち）はまとめて状態ファイルに書く。書き出したフォルダの数を返す。
    #   pending: フォルダ（フルパス）→ 無くなった元のファイル名の集まり
    param (
        $pending,
        [string]$indexRoot = $workspace.IndexDir,
        [string]$systemRoot = $workspace.SystemIndexDir,
        [string]$statePath = $workspace.SystemIndexStateFile
    )

    $results = New-Object System.Collections.Generic.List[hashtable]
    foreach ($folder in @($pending.Keys)) {
        $contentIndexFile = updateFolderContentIndex $folder ([string[]]@($pending[$folder]))
        # 書いた本文インデックスのファイルは、フォルダごとに書いた直後にここで NotContentIndexed を付ける（件数が多いため失敗してもログに出さない）
        [void](setNotContentIndexed $folder)
        $results.Add((writeSystemIndexFolder $folder $indexRoot $systemRoot $contentIndexFile.Texts))
    }
    if ($results.Count -gt 0) {
        $saved = updateSystemIndexState { param ($state) setSystemIndexResults $state $results.ToArray() } $statePath
        if (!$saved) {
            writeIndexerLog "システムインデックスの状態を書き込めませんでした（インデックス作成の終わりに作り直します）。" "Yellow"
        }
    }
    return $results.Count
}


function readContentIndexLines {
    # 本文インデックスのファイルの中の、元のファイル book・場所 location の firstLine 行目から count 行を返す（画面の選択行のプレビュー。
    # 行の数え方は検索と同じ）。戻り値は @{ Rows = @{ LineNumber; Line } の配列; End = 頼んだ数より少なく返ったか（場所の終わりまで読んだ）; Marks = 更新した控え }。
    # cache（検索のキャッシュ）に同じ本文インデックスのファイルの内容があれば、ファイルを読み直さない。読めない・見つからなければ Rows は空・End は $true。
    #
    # marks は、繰り返し読むときに先頭から読み飛ばさないための控え（@{ Path; Ticks; Size; Places; Checkpoints }。
    # ファイルのパス・更新日時・大きさが同じ間だけ使う）。呼び出し側が引数で渡し、戻り値の Marks で受け取る
    # （読み足しは別の runspace で動くため、$script: には置かない。渡された控えは書き換えず、新しいものを返す）。
    # Places は場所の一覧、Checkpoints は場所ごと（Book と Location をタブでつないだ名前）の、200 行ごとの行の先頭の位置（行の番号 → 文字の位置）
    param (
        [string]$path,
        [string]$book,
        [string]$location,
        [int]$firstLine,
        [int]$count,
        $cache = $null,
        $marks = $null
    )

    $first = [Math]::Max(1, $firstLine)
    $last = $first + $count - 1
    if ($count -le 0) { return @{ Rows = @(); End = $false; Marks = $marks } }
    $nothing = @{ Rows = @(); End = $true; Marks = $marks }
    $text = $null
    $places = $null
    $longPath = toLongPath $path
    $entry = $null
    if ($null -ne $cache -and $cache.Texts.TryGetValue($longPath, [ref]$entry) -and $entry.Count -ge 4) {
        $info = [System.IO.FileInfo]::new($longPath)
        if ($info.Exists -and $entry[0] -eq $info.LastWriteTimeUtc.Ticks -and $entry[1] -eq $info.Length) {
            $text = $entry[2]
            $places = $entry[3]
        }
    }
    $isRange = $false
    $newMarks = $marks
    $checkpoints = $null
    $allCheckpoints = $null
    $placeKey = "$book`t$location"
    try {
        if ($null -eq $text) {
            # キャッシュに無いときは、ファイルの全体を文字列にしない。控えがあれば、その場所の目当ての行の手前の位置から必要な行までだけを読む
            # （続けて読むたびに、全体を読み直さず、場所の先頭からも読み飛ばさない）
            $info = [System.IO.FileInfo]::new($longPath)
            if (!$info.Exists) { return $nothing }
            if ($null -ne $marks -and $marks.Path -eq $longPath -and $marks.Ticks -eq $info.LastWriteTimeUtc.Ticks -and $marks.Size -eq $info.Length) {
                $places = $marks.Places
                $allCheckpoints = $(if ($null -ne $marks.Checkpoints) { $marks.Checkpoints.Clone() } else { @{} })
                $found = $false
                foreach ($place in $places) {
                    if ($place.Book -ne $book -or $place.Location -ne $location) { continue }
                    $found = $true
                    # 位置は場所ごとに控える（行の番号が同じでも、別の場所の位置は使わない）
                    $checkpoints = $(if ($allCheckpoints.ContainsKey($placeKey)) { $allCheckpoints[$placeKey].Clone() } else { @{} })
                    $allCheckpoints[$placeKey] = $checkpoints
                    $base = 1
                    $startPosition = $place.Start
                    foreach ($key in $checkpoints.Keys) {
                        if ($key -le $first -and $key -gt $base) { $base = $key; $startPosition = $checkpoints[$key] }
                    }
                    $text = readContentIndexRange $path $startPosition $place.End ($first - $base + 1) ($last - $base + 1) $base $checkpoints
                    break
                }
                # 控えが今のファイルのものなら、無い場所はファイルを読まずに空
                if (!$found) { return $nothing }
                if ($null -ne $text) {
                    # 読んでいる間にファイルが置き換わっていたら、古い位置で読んだ行は使わず、全体を読み直す
                    $again = [System.IO.FileInfo]::new($longPath)
                    if ($again.Exists -and $again.LastWriteTimeUtc.Ticks -eq $info.LastWriteTimeUtc.Ticks -and $again.Length -eq $info.Length) { $isRange = $true } else { $text = $null }
                }
            }
            if ($null -eq $text) {
                $text = readContentIndexText $path
                $places = readContentIndexPlaces $text
                $checkpoints = @{}
                $allCheckpoints = @{ $placeKey = $checkpoints }
            }
            $newMarks = @{ Path = $longPath; Ticks = $info.LastWriteTimeUtc.Ticks; Size = $info.Length; Places = $places; Checkpoints = $allCheckpoints }
        }
    } catch [System.IO.IOException] {
        return $nothing
    } catch [System.UnauthorizedAccessException] {
        return $nothing
    }
    $rows = New-Object System.Collections.Generic.List[psobject]
    $isTextBook = (getContentIndexFileKind $book) -eq "テキスト"
    foreach ($place in $places) {
        if ($place.Book -ne $book -or $place.Location -ne $location) { continue }
        # 範囲だけ読んだときは、text の先頭が $first 行目の先頭
        $number = if ($isRange) { $first - 1 } else { 0 }
        $placeEnd = if ($isRange) { $text.Length } else { [Math]::Min($place.End, $text.Length) }
        $pos = if ($isRange) { 0 } else { $place.Start }
        # ファイルの全体を読んだときは、通った 200 行ごとの行の先頭を控える（範囲を読んだときは readContentIndexRange が控える）
        $record = !$isRange -and $null -ne $checkpoints
        while ($pos -lt $placeEnd) {
            $n = $text.IndexOf([char]10, $pos, $placeEnd - $pos)
            if ($n -lt 0) { $n = $placeEnd }
            $number++
            if ($record -and $number -gt 1 -and ($number - 1) % 200 -eq 0) { $checkpoints[$number] = $pos }
            if ($number -gt $last) { break }
            if ($number -ge $first) {
                $line = $text.Substring($pos, $n - $pos)
                if ($isTextBook) {
                    # 前後の行はどこが一致したかが分からないため、先頭から切る
                    $line = truncateHitLine $line -1
                }
                $rows.Add([pscustomobject]@{ LineNumber = $number; Line = $line })
            }
            $pos = $n + 1
        }
        break
    }
    return @{ Rows = $rows.ToArray(); End = ($rows.Count -lt $count); Marks = $newMarks }
}

