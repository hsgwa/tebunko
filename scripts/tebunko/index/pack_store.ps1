# 検索用の集約ファイル（pack_format.ps1）の読み書き（状態層）。
# 集約ファイルは work\index の中のフォルダごと・元のファイルの拡張子ごとに、大きさで分けて置く（content.xlsx.001.tsv など）。UTF-16LE（BOM 付き）で書く
# （UTF-8 より文字列への変換が速い。日本語が多いと大きさはほとんど変わらない）。

function writePackFile {
    # 集約ファイルを書く。一時ファイル（<名前>.tmp）に書き終えてから置き換えるため、途中で止まっても前のファイルが残る
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


function readPackText {
    # 集約ファイルを文字列で読む（インデックス作成中の置き換えと同時に読めるよう、共有モードは ReadWrite|Delete）
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


function testIndexBookDir {
    # インデックスの中のフォルダが、元のファイルごとのフォルダ（<ファイル名.xlsx>\<場所>.tsv。集約ファイルに入れる前の TSV の置き場所）か。
    # 名前だけでは、名前が .xlsx などで終わる本物のフォルダ（元のフォルダの名前をそのまま使う）と区別できないため、中身も見る:
    #   ・名前が Office の拡張子で終わる（indexBookDirPattern）
    #   ・サブフォルダも集約ファイル（content.<拡張子>.tsv）も無い
    #   ・withTsv なら、TSV が 1 つ以上ある（取り込んだが中身が空のファイルのフォルダは、集約ファイルに入れるものが無い）
    # 読めないフォルダは $false（集約ファイルに入れる・消す対象にしない）
    param (
        [string]$dir,
        [bool]$withTsv = $true
    )

    if ([System.IO.Path]::GetFileName($dir.TrimEnd("\")) -notmatch ${indexBookDirPattern}) {
        return $false
    }
    $long = toLongPath $dir
    try {
        foreach ($sub in [System.IO.Directory]::EnumerateDirectories($long)) {
            return $false
        }
        $hasTsv = $false
        foreach ($file in [System.IO.Directory]::EnumerateFiles($long, "*.tsv")) {
            if ([System.IO.Path]::GetFileName($file) -like ${packFilePattern}) {
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
    # 今の形式のインデックスのフォルダ 1 つ（直下の <ファイル名.xlsx>\<場所>.tsv）から、集約ファイルに入れる元のファイルの並びを作る。
    # 並びは今の検索結果と同じ順（TSV のパスを現在のカルチャ・大文字と小文字を区別しない順に並べたもの）。
    # @{ Name; Places（@{ Place; Path } の並び） } の並びを返す
    param (
        [string]$folder
    )

    $longDir = toLongPath $folder
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($sub in [System.IO.Directory]::EnumerateDirectories($longDir)) {
        if (!(testIndexBookDir $sub)) {
            continue
        }
        $paths.AddRange([System.IO.Directory]::GetFiles($sub, "*.tsv", [System.IO.SearchOption]::TopDirectoryOnly))
    }
    $sorted = $paths.ToArray()
    [System.Array]::Sort($sorted, [System.StringComparer]::CurrentCultureIgnoreCase)
    $books = New-Object System.Collections.Generic.List[hashtable]
    $current = $null
    foreach ($path in $sorted) {
        $name = [System.IO.Path]::GetFileName([System.IO.Path]::GetDirectoryName($path))
        if (!$current -or $current.Name -ne $name) {
            $current = @{ Name = $name; Places = New-Object System.Collections.Generic.List[hashtable] }
            $books.Add($current)
        }
        $current.Places.Add(@{ Place = (decodeIndexPlace ([System.IO.Path]::GetFileNameWithoutExtension($path))); Path = $path })
    }
    return , $books
}


function convertIndexFolderToPack {
    # 今の形式のインデックスのフォルダ 1 つ（直下の <ファイル名.xlsx>\<場所>.tsv）から、拡張子ごと・番号ごとの集約ファイル
    # （destFolder\content.xlsx.001.tsv など）を書く。destFolder に前の集約ファイルがあれば、それとまぜる（planPackParts）:
    #   ・TSV のある元のファイルは、TSV の中身で入れ替える（前の集約ファイルに無ければ、最後の番号の集約ファイルに足す。
    #     packFileMaxBytes 以上なら次の番号の集約ファイルを作る）
    #   ・removeBooks に挙げた元のファイルは外す（元のファイルが無くなった）
    #   ・それ以外の元のファイルは、前の集約ファイルのまま。変わらない集約ファイルは書き直さない
    # 元のファイルが無くなった集約ファイルは消す。removeTsv なら、集約ファイルを書き終えた後に、
    # 読み込んだ元のファイルのフォルダ（TSV）を消す（TSV は一時的な置き場で、残すとインデックスの容量が倍になるため）。
    # 書き終える前に止まっても、TSV か前の集約ファイルのどちらかに中身が残る。
    # @{ Books; Tsv; Chars; Files（集約ファイルの数）; Written（書き直した数）;
    #    Texts（そのフォルダのすべての集約ファイルの中身。システムインデックスを読み直さずに作るため） } を返す
    param (
        [string]$folder,
        [string]$destFolder,
        [string[]]$removeBooks = @(),
        [bool]$removeTsv = $false,
        [long]$maxBytes = ${packFileMaxBytes}
    )

    $books = getIndexFolderBooks $folder
    $tsvCount = 0
    $newBooks = New-Object System.Collections.Generic.List[hashtable]
    foreach ($book in $books) {
        foreach ($place in $book.Places) {
            $place.Text = [System.IO.File]::ReadAllText($place.Path)
            $tsvCount++
        }
        $newBooks.Add(@{ Name = $book.Name; Block = (convertBookToPackBlock $book) })
    }
    $longDest = toLongPath $destFolder
    # 前の集約ファイル（番号の付いた名前のもの）を読む
    $parts = New-Object System.Collections.Generic.List[hashtable]
    $oldTexts = @{}
    if ([System.IO.Directory]::Exists($longDest)) {
        foreach ($old in [System.IO.Directory]::GetFiles($longDest, ${packFilePattern})) {
            $info = readPackFileName ([System.IO.Path]::GetFileName($old))
            if ($null -eq $info) { continue }
            $text = readPackText $old
            $oldTexts[(getPackFileName $info.Extension $info.Part)] = $text
            $parts.Add(@{ Extension = $info.Extension; Part = $info.Part; Books = (splitPackTextByBook $text) })
        }
    }

    $plan = planPackParts $parts $newBooks $removeBooks $maxBytes
    $texts = New-Object System.Collections.Generic.List[string]
    $bookTotal = 0
    $written = 0
    $chars = 0L
    foreach ($part in $plan) {
        $name = getPackFileName $part.Extension $part.Part
        $path = Join-Path $destFolder $name
        if ($part.Books.Count -eq 0) {
            if ([System.IO.File]::Exists((toLongPath $path))) { [System.IO.File]::Delete((toLongPath $path)) }
            continue
        }
        if ($part.Changed) {
            [void][System.IO.Directory]::CreateDirectory($longDest)
            $text = convertToPackText $part.Books
            writePackFile $path $text
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

function updateIndexFolderPack {
    # インデックスのフォルダ 1 つで、置かれた TSV（追加・更新した元のファイル）を集約ファイルに入れ、TSV を消す。
    # removeBooks に挙げた元のファイル（無くなったもの）は集約ファイルから外す
    param (
        [string]$folder,
        [string[]]$removeBooks = @()
    )

    return convertIndexFolderToPack $folder $folder $removeBooks $true
}

function getPackFiles {
    # インデックスのフォルダ以下の集約ファイルを列挙し、フォルダの順・フォルダの中は名前の順に並べて返す。
    #   root   : インデックスのフォルダ（相対パスの基準）
    #   relPath: その中のフォルダ（空は root 自身）
    #   recurse: $false なら、そのフォルダの集約ファイルだけ
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
    $found = [System.IO.DirectoryInfo]::new($longDir).GetFiles(${packFilePattern}, $option)
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
    # フォルダの順、フォルダの中は集約ファイルの名前の順（現在のカルチャ・大文字と小文字を区別しない）
    $items = [hashtable[]]@($list | Sort-Object @{ Expression = { $_.RelDir } }, @{ Expression = { [System.IO.Path]::GetFileName($_.RelPath) } })
    return , $items
}


function findIndexFoldersWithBooks {
    # インデックスのフォルダ以下で、元のファイルごとのフォルダ（<ファイル名.xlsx>。集約ファイルに入れる前の TSV）が
    # 直下にあるフォルダを返す（インデックス作成が途中で止まった・前の形式のインデックス）。root 自身も含む
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
        foreach ($sub in [System.IO.Directory]::EnumerateDirectories($dir)) {
            if (testIndexBookDir $sub) {
                $found.Add((fromLongPath $dir))
                break
            }
        }
    }
    return , $found.ToArray()
}


function publishIndexFolders {
    # インデックス作成で TSV を置いた・元のファイルが無くなったフォルダを、まとめて書き出す。フォルダごとに次を続けて行う:
    #   1. 集約ファイルを書く（前の集約ファイルとまぜ、無くなった元のファイルは外す）
    #   2. 元のファイルごとのフォルダの TSV を消す
    #   3. 書いた集約ファイルの中身から、そのフォルダのシステムインデックスの txt を作る（読み直さない）
    # txt の状態（反映待ち）はまとめて状態ファイルに書く。書き出したフォルダの数を返す。
    #   pending: フォルダ（フルパス）→ 無くなった元のファイル名の集まり
    param (
        $pending,
        [string]$indexRoot = ${indexDir},
        [string]$systemRoot = ${systemIndexDir},
        [string]$statePath = ${systemIndexStateFile}
    )

    $results = New-Object System.Collections.Generic.List[hashtable]
    foreach ($folder in @($pending.Keys)) {
        $pack = updateIndexFolderPack $folder ([string[]]@($pending[$folder]))
        $results.Add((writeSystemIndexFolder $folder $indexRoot $systemRoot $pack.Texts))
    }
    if ($results.Count -gt 0) {
        $saved = updateSystemIndexState { param ($state) setSystemIndexResults $state $results.ToArray() } $statePath
        if (!$saved) {
            writeIndexerLog "システムインデックスの状態を書き込めませんでした（インデックス作成の終わりに作り直します）。" "Yellow"
        }
    }
    return $results.Count
}


function readPackContext {
    # 集約ファイルの中の、元のファイル book・場所 location の lineNumber 行目と、その前後 before 行・after 行を
    # @{ LineNumber; Line } の配列で返す（画面の選択行のプレビュー。行の数え方は検索と同じ）。
    # cache（検索のキャッシュ）に同じ集約ファイルの内容があれば、ファイルを読み直さない。読めない・見つからなければ空
    param (
        [string]$path,
        [string]$book,
        [string]$location,
        [int]$lineNumber,
        [int]$before = 3,
        [int]$after = 3,
        $cache = $null
    )

    $rows = New-Object System.Collections.Generic.List[psobject]
    $first = [Math]::Max(1, $lineNumber - $before)
    $last = $lineNumber + $after
    if ($last -lt $first) { return @() }
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
    try {
        if ($null -eq $text) {
            $text = readPackText $path
            $places = readPackPlaces $text
        }
    } catch [System.IO.IOException] {
        return @()
    } catch [System.UnauthorizedAccessException] {
        return @()
    }
    foreach ($place in $places) {
        if ($place.Book -ne $book -or $place.Location -ne $location) { continue }
        $number = 0
        $pos = $place.Start
        while ($pos -lt $place.End) {
            $n = $text.IndexOf([char]10, $pos, $place.End - $pos)
            if ($n -lt 0) { $n = $place.End }
            $number++
            if ($number -gt $last) { break }
            if ($number -ge $first) {
                $rows.Add([pscustomobject]@{ LineNumber = $number; Line = $text.Substring($pos, $n - $pos) })
            }
            $pos = $n + 1
        }
        break
    }
    # 呼び出し側で @() にして使う
    return $rows.ToArray()
}