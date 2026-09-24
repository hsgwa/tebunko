# 検索用のまとめファイル（pack_format.ps1）の読み書き（状態層）。
# まとめファイルは work\index の中のフォルダごと・元のファイルの拡張子ごとに 1 つ（本文.xlsx.tsv など）。UTF-16LE（BOM 付き）で書く
# （UTF-8 より文字列への変換が速い。日本語が多いと大きさはほとんど変わらない）。

function writePackFile {
    # まとめファイルを書く。一時ファイル（<名前>.tmp）に書き終えてから置き換えるため、途中で止まっても前のファイルが残る
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
    # まとめファイルを文字列で読む（インデックス作成中の置き換えと同時に読めるよう、共有モードは ReadWrite|Delete）
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


function getIndexFolderBooks {
    # 今の形式のインデックスのフォルダ 1 つ（直下の <ファイル名.xlsx>\<場所>.tsv）から、まとめファイルに入れる元のファイルの並びを作る。
    # 並びは今の検索結果と同じ順（TSV のパスを現在のカルチャ・大文字と小文字を区別しない順に並べたもの）。
    # @{ Name; Places（@{ Place; Path } の並び） } の並びを返す
    param (
        [string]$folder
    )

    $longDir = toLongPath $folder
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($sub in [System.IO.Directory]::EnumerateDirectories($longDir)) {
        if ([System.IO.Path]::GetFileName($sub) -notmatch ${indexBookDirPattern}) {
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
    # 今の形式のインデックスのフォルダ 1 つ（直下の <ファイル名.xlsx>\<場所>.tsv）から、拡張子ごとのまとめファイル
    # （destFolder\本文.xlsx.tsv など）を書く。destFolder に前のまとめファイルがあれば、それとまぜる:
    #   ・TSV のある元のファイルは、TSV の中身で入れ替える（追加・更新）
    #   ・removeBooks に挙げた元のファイルは外す（元のファイルが無くなった）
    #   ・それ以外の元のファイルは、前のまとめファイルからそのまま写す
    # 元のファイルが無くなった拡張子のまとめファイルは消す。removeTsv なら、まとめファイルを書き終えた後に、
    # 読み込んだ元のファイルのフォルダ（TSV）を消す（TSV は一時的な置き場で、残すとインデックスの容量が倍になるため）。
    # 書き終える前に止まっても、TSV か前のまとめファイルのどちらかに中身が残る。@{ Books; Tsv; Chars; Files } を返す
    param (
        [string]$folder,
        [string]$destFolder,
        [string[]]$removeBooks = @(),
        [bool]$removeTsv = $false
    )

    $books = getIndexFolderBooks $folder
    $tsvCount = 0
    foreach ($book in $books) {
        foreach ($place in $book.Places) {
            $place.Text = [System.IO.File]::ReadAllText($place.Path)
            $tsvCount++
        }
    }
    $longDest = toLongPath $destFolder
    # 前のまとめファイルから、入れ替えない元のファイルを取り出す
    $replaced = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($book in $books) { [void]$replaced.Add($book.Name) }
    foreach ($name in $removeBooks) { [void]$replaced.Add($name) }
    $merged = New-Object System.Collections.Generic.List[object]
    $merged.AddRange([object[]]@($books))
    if ([System.IO.Directory]::Exists($longDest)) {
        foreach ($old in [System.IO.Directory]::GetFiles($longDest, ${packFilePattern})) {
            foreach ($kept in (splitPackTextByBook (readPackText $old))) {
                if (!$replaced.Contains($kept.Name)) { $merged.Add($kept) }
            }
        }
    }
    # 元のファイル名の順（現在のカルチャ・大文字と小文字を区別しない）
    $sorted = @($merged | Sort-Object { [string]$_.Name })

    $written = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $chars = 0L
    if ($sorted.Count -gt 0) {
        [void][System.IO.Directory]::CreateDirectory($longDest)
        $groups = splitPackBooksByExtension $sorted
        foreach ($extension in $groups.Keys) {
            $name = getPackFileName $extension
            $text = convertToPackText $groups[$extension]
            writePackFile (Join-Path $destFolder $name) $text
            [void]$written.Add($name)
            $chars += $text.Length
        }
    }
    if ([System.IO.Directory]::Exists($longDest)) {
        foreach ($old in [System.IO.Directory]::GetFiles($longDest, ${packFilePattern})) {
            if (!$written.Contains([System.IO.Path]::GetFileName($old))) {
                [System.IO.File]::Delete($old)
            }
        }
    }
    if ($removeTsv) {
        foreach ($book in $books) {
            $bookDir = [System.IO.Path]::GetDirectoryName($book.Places[0].Path)
            [System.IO.Directory]::Delete($bookDir, $true)
        }
    }
    return @{ Books = $sorted.Count; Tsv = $tsvCount; Chars = $chars; Files = $written.Count }
}


function updateIndexFolderPack {
    # インデックスのフォルダ 1 つで、置かれた TSV（追加・更新した元のファイル）をまとめファイルに入れ、TSV を消す。
    # removeBooks に挙げた元のファイル（無くなったもの）はまとめファイルから外す
    param (
        [string]$folder,
        [string[]]$removeBooks = @()
    )

    return convertIndexFolderToPack $folder $folder $removeBooks $true
}

function getPackFiles {
    # インデックスのフォルダ以下のまとめファイルを列挙し、フォルダの順・フォルダの中は名前の順に並べて返す。
    #   root   : インデックスのフォルダ（相対パスの基準）
    #   relPath: その中のフォルダ（空は root 自身）
    #   recurse: $false なら、そのフォルダのまとめファイルだけ
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
    # フォルダの順、フォルダの中はまとめファイルの名前の順（現在のカルチャ・大文字と小文字を区別しない）
    $items = [hashtable[]]@($list | Sort-Object @{ Expression = { $_.RelDir } }, @{ Expression = { [System.IO.Path]::GetFileName($_.RelPath) } })
    return , $items
}
