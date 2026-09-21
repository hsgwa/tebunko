# インデックスの TSV を検索し、結果を組み立てる。

function toResultLine {
    # 検索結果1件を "ファイル名<TAB>場所<TAB>行番号<TAB>該当行" に整形する。
    # Excelに貼り付けたとき、該当行の各セルが元の列の順（4列目 = A列）に並ぶようにする
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
    $place = [regex]::Replace($location, '[\x00-\x1F]', " ")
    return "${book}`t${place}`t${lineNumber}`t${line}"
}

function toResultHeader {
    # 検索結果の見出し行 "ファイル名<TAB>場所<TAB>行<TAB>A<TAB>B…" を返す（列名は columnCount 列分）
    param (
        [int]$columnCount
    )

    $names = New-Object System.Collections.Generic.List[string]
    $names.AddRange([string[]]@("ファイル名", "場所", "行"))
    for ($i = 1; $i -le $columnCount; $i++) {
        $names.Add((toColumnName $i))
    }
    return ($names -join "`t")
}


# 検索処理の本体（TSVを読んで照合し、結果を作る）。Select-String と PowerShell で1件ずつ結果を作ると、
# 検索処理。以前は C# にして Add-Type でコンパイルしていたが、実行時コンパイル（csc.exe）を無くすため
# .NET を直接呼ぶ PowerShell 関数にした。数万件でも [StreamReader]＋[regex] のタイトループで実用的な速度
# （実測: 50 万行で約 0.6 秒）。ヒットは PSCustomObject（Root; RelPath; RelDir; FileName; Book; Location; LineNumber; Line）で返す。


function newTsvFiles {
    # 検索対象のTSVを、元のファイル名（Book）・場所（Location）付きに整える。
    # include に一致しない・exclude に一致する Book は除く（$null は条件なし）。パスの分解は splitIndexTsvPath に合わせる。
    # excludePlace に一致する場所（図形・コメント。newPlaceExclude）も除く。
    #   tsvFiles: getIndexTsvFiles の Files（TSVのフルパス → @{ Root; RelPath; LongPath; Ticks; Size }。LongPath 以降は無くてもよい）
    param (
        $tsvFiles,
        [regex]$include = $null,
        [regex]$exclude = $null,
        [regex]$excludePlace = $null
    )

    $files = New-Object System.Collections.Generic.List[hashtable]
    # -match と同じく大文字と小文字を区別しない
    $bookDirPattern = [regex]::new(${indexBookDirPattern}, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($path in $tsvFiles.Keys) {
        $info = $tsvFiles[$path]
        $relPath = [string]$info.RelPath
        $fileName = [System.IO.Path]::GetFileName($relPath)
        $dir = [System.IO.Path]::GetDirectoryName($relPath)
        $bookDir = if ($dir) { [System.IO.Path]::GetFileName($dir) } else { "" }
        if ($fileName.IndexOfAny([char[]]"_%") -lt 0 -and $bookDirPattern.IsMatch($bookDir)) {
            # 今の形式で、場所に符号化した文字が無いもの（ほとんどのTSV）。splitIndexTsvPath と同じ結果を、関数を呼ばずに作る
            # （TSV が数万件あると、関数呼び出しだけで検索を始めるまでに数秒かかるため）
            $book = $bookDir
            $place = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $relDir = [string][System.IO.Path]::GetDirectoryName($dir)
        } else {
            $parts = splitIndexTsvPath $relPath
            $book = [string]$parts.Book
            $place = [string]$parts.Place
            $relDir = [string]$parts.RelDir
        }
        if ($include -and !$include.IsMatch($book)) { continue }
        if ($exclude -and $exclude.IsMatch($book)) { continue }
        if ($excludePlace -and $excludePlace.IsMatch($place)) { continue }
        # PSCustomObject より作るのが速いハッシュテーブルにする（TSV の数だけ作るため）
        $files.Add(@{
            # 260文字を超えるパスのTSVも読めるよう \\?\ 付きで読む（getIndexTsvFiles が列挙したパスがあればそれを使う）
            Path     = if ($info.LongPath) { $info.LongPath } else { toLongPath $path }
            Root     = $info.Root
            RelPath  = $relPath
            RelDir   = $relDir
            FileName = $fileName
            Book     = $book
            Location = $place
            Ticks    = $info.Ticks
            Size     = $info.Size
        })
    }
    return , $files
}

# 全文を読んで照合する TSV の大きさの上限（バイト）。これより大きい TSV は、メモリを使いすぎないよう 1 行ずつ読む
${searchWholeFileMax} = 64MB

# 検索で読んだ TSV の内容を残しておく量の上限（文字数。1 文字 2 バイトのため約 128MB）
${searchCacheMaxChars} = 64000000

function newTsvTextCache {
    # 検索で読んだ TSV の内容を、次の検索で使い回すための入れ物を作る（画面が 1 つ持ち、検索のたびに searchIndex に渡す）。
    # TSV を 1 つずつ開いて読む時間が検索時間の大半のため（TSV 1 万件で約 1 秒）、2 回目以降の検索ではファイルを開かない。
    # 更新日時・サイズが列挙したときと違う TSV（変換し直した等）は読み直す。並列検索の各スレッドから使うため、中身は ConcurrentDictionary。
    #   Texts: TSV の \\?\ 付きのパス → @(更新日時（UTC の Ticks）, サイズ, 内容) / Chars: 残している文字数 / MaxChars: 上限
    param (
        [long]$maxChars = ${searchCacheMaxChars}
    )

    return @{
        Texts    = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        Chars    = [long[]]::new(1)
        MaxChars = $maxChars
    }
}

function searchTsvFiles {
    # files の start から count 件を読み、regex に一致する行を PSCustomObject で返す（1行に複数一致しても1件）。
    # max 以上（max+1 件目）が見つかった時点で打ち切る（負は上限なし）。読めないTSV（変換中に削除された等）は飛ばす。
    # 変換中のTSVも読めるよう共有モードは ReadWrite|Delete にする。正規表現の照合が時間切れ（RegexMatchTimeoutException）なら
    # 例外はそのまま呼び出し元（searchIndex）へ伝わる。
    #
    # 1 行ずつ PowerShell で照合すると、行数に比例して遅い（1 行あたり十数マイクロ秒）。textRegex・scanMode（newSearchRegex）を
    # 渡すと、TSV を丸ごと読んで全文に 1 回照合し（.NET の中で走る）、1 行ずつの照合を減らす。
    #   lines : 全文での一致の位置から行を切り出す（1 行ずつの照合はしない）
    #   filter: 全文で一致しない TSV は飛ばし、一致した TSV だけ 1 行ずつ照合する
    #   scan  : 1 行ずつ照合する
    # 全文の改行は LF にそろえる（StreamReader.ReadLine と同じく CRLF・LF・CR を行の区切りとし、行の中身・行番号は変わらない）。
    # 全文への照合が時間切れになった TSV は、1 行ずつ照合し直す（行ごとの時間切れの扱いは従来どおり）。
    # cache（newTsvTextCache）を渡すと、読んだ内容を残し、次の検索では更新日時・サイズが同じ TSV をファイルから読まない
    param (
        $files,
        [int]$start,
        [int]$count,
        [regex]$regex,
        [int]$max,
        [regex]$textRegex = $null,
        [string]$scanMode = "scan",
        $cache = $null
    )

    # 並列検索の別スレッド（searchIndex）でも動くよう、コマンドレット（New-Object 等）は使わない
    $hits = [System.Collections.Generic.List[psobject]]::new()
    $end = [Math]::Min($files.Count, $start + $count)
    $lf = [char]10
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $rawCheck = ($null -ne $textRegex -and $textRegex.ToString().IndexOfAny([char[]]"^$") -lt 0)
    for ($i = $start; $i -lt $end; $i++) {
        $f = $files[$i]
        $reader = $null
        try {
            $mode = if ($null -eq $textRegex) { "scan" } else { $scanMode }
            # 前の検索で読んだ内容があり、更新日時・サイズが列挙したときと同じなら、ファイルを開かずに使う
            $text = $null
            $cacheable = ($null -ne $cache -and $null -ne $f.Ticks -and $f.Size -le ${searchWholeFileMax})
            if ($cacheable) {
                $entry = $null
                if ($cache.Texts.TryGetValue($f.Path, [ref]$entry) -and $entry[0] -eq $f.Ticks -and $entry[1] -eq $f.Size) {
                    $text = $entry[2]
                }
            }
            if ($null -eq $text) {
                $stream = [System.IO.FileStream]::new($f.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
                $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
                if ($stream.Length -gt ${searchWholeFileMax}) {
                    $mode = "scan"
                } elseif ($mode -ne "scan" -or $cacheable) {
                    $text = $reader.ReadToEnd()
                    $reader.Dispose()
                    $reader = $null
                    if ($cacheable) {
                        # 上限を超える分は残さない（前の内容は、新しい内容に置き換えられなければ消す）
                        [System.Threading.Monitor]::Enter($cache)
                        try {
                            $old = $null
                            if ($cache.Texts.TryRemove($f.Path, [ref]$old)) {
                                $cache.Chars[0] -= $old[2].Length
                            }
                            if ($cache.Chars[0] + $text.Length -le $cache.MaxChars) {
                                $cache.Texts[$f.Path] = [object[]]@($f.Ticks, $f.Size, $text)
                                $cache.Chars[0] += $text.Length
                            }
                        } finally {
                            [System.Threading.Monitor]::Exit($cache)
                        }
                    }
                }
            }
            if ($null -ne $text -and $mode -ne "scan") {
                $before = $hits.Count
                try {
                    # ^ $ を含まなければ、改行をそろえる前の全文で一致しない TSV は、そろえても一致しない
                    # （ヒットしない TSV の文字列のコピーを省く。^ $ は CR だけの改行・CRLF で位置が変わるため、そろえてから照合する）
                    $rawChecked = $false
                    if ($rawCheck) {
                        if (!$textRegex.IsMatch($text)) { continue }
                        $rawChecked = $true
                    }
                    $text = $text.Replace("`r`n", "`n").Replace("`r", "`n")
                    if ($mode -eq "filter") {
                        if (!$rawChecked -and !$textRegex.IsMatch($text)) { continue }
                    } else {
                        $length = $text.Length
                        # 末尾の改行の後ろは行ではない（ReadLine は空の行を返さない）
                        $tailIsLine = $length -gt 0 -and $text[$length - 1] -ne $lf
                        $number = 1
                        $pos = 0  # 行 number の先頭
                        $m = $textRegex.Match($text)
                        while ($m.Success) {
                            $index = $m.Index
                            if ($index -ge $length -and !$tailIsLine) { break }
                            $lineStart = if ($index -eq 0) { 0 } else { $text.LastIndexOf($lf, $index - 1) + 1 }
                            $lineEnd = $text.IndexOf($lf, $index)
                            if ($lineEnd -lt 0) { $lineEnd = $length }
                            if ($lineStart -gt $pos) {
                                $skipped = $text.Substring($pos, $lineStart - $pos)
                                $number += $skipped.Length - $skipped.Replace("`n", "").Length
                                $pos = $lineStart
                            }
                            $hits.Add([pscustomobject]@{
                                Root = $f.Root; RelPath = $f.RelPath; RelDir = $f.RelDir; FileName = $f.FileName
                                Book = $f.Book; Location = $f.Location; LineNumber = $number; Line = $text.Substring($lineStart, $lineEnd - $lineStart)
                            })
                            if ($max -ge 0 -and $hits.Count -gt $max) { return , $hits }
                            if ($lineEnd -ge $length) { break }
                            # 1 行に複数一致しても 1 件にするため、次の行の先頭から探す
                            $m = $textRegex.Match($text, $lineEnd + 1)
                        }
                        continue
                    }
                } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                    if ($hits.Count -gt $before) { $hits.RemoveRange($before, $hits.Count - $before) }
                }
            }
            if ($null -ne $text) {
                $reader = [System.IO.StringReader]::new($text)
            }
            $number = 0
            while ($null -ne ($line = $reader.ReadLine())) {
                $number++
                if (!$regex.IsMatch($line)) { continue }
                $hits.Add([pscustomobject]@{
                    Root = $f.Root; RelPath = $f.RelPath; RelDir = $f.RelDir; FileName = $f.FileName
                    Book = $f.Book; Location = $f.Location; LineNumber = $number; Line = $line
                })
                if ($max -ge 0 -and $hits.Count -gt $max) { return , $hits }
            }
        } catch [System.IO.IOException] {
        } catch [System.UnauthorizedAccessException] {
        } finally {
            if ($reader) { $reader.Dispose() }
        }
    }
    return , $hits
}

function findTsvFilesParallel {
    # フォルダ以下（再帰）の *.tsv を FileInfo で返す（順不同。呼び出し元で並べる）。
    # FileInfo の更新日時・サイズは走査のときに得られる（1 件ずつ問い合わせない）。検索の読み込み結果の再利用に使う。
    # フォルダの走査は、ファイルの数が多いとそれだけで検索前の待ち時間になるため（TSV 1 万件で約 0.6 秒）、
    # 直下のフォルダごとに別スレッドで並行して走査する。直下のフォルダが少なければ 1 スレッドで走査する
    param (
        [string]$longDir
    )

    $found = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    $top = [System.IO.DirectoryInfo]::new($longDir)
    $subDirs = [System.IO.Directory]::GetDirectories($longDir)
    $workers = [Math]::Min([Math]::Min([Environment]::ProcessorCount, 4), $subDirs.Length)
    if ($workers -lt 2) {
        $found.AddRange($top.GetFiles("*.tsv", [System.IO.SearchOption]::AllDirectories))
        return , $found
    }

    $found.AddRange($top.GetFiles("*.tsv", [System.IO.SearchOption]::TopDirectoryOnly))
    $pool = [runspacefactory]::CreateRunspacePool(1, $workers, [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2(), $Host)
    $jobs = New-Object System.Collections.Generic.List[hashtable]
    try {
        $pool.Open()
        foreach ($sub in $subDirs) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool
            # 出力に配列をそのまま出すと 1 件ずつに分かれるため、ハッシュテーブルに入れて返す
            [void]$ps.AddScript({
                param ($dir)
                @{ Files = [System.IO.DirectoryInfo]::new($dir).GetFiles("*.tsv", [System.IO.SearchOption]::AllDirectories) }
            }).AddArgument($sub)
            $jobs.Add(@{ PowerShell = $ps; Handle = $ps.BeginInvoke() })
        }
        foreach ($job in $jobs) {
            $output = $job.PowerShell.EndInvoke($job.Handle)
            $found.AddRange([System.IO.FileInfo[]]$output[0].Files)
        }
    } finally {
        foreach ($job in $jobs) {
            $job.PowerShell.Dispose()
        }
        $pool.Dispose()
    }
    return , $found
}

function getIndexTsvFiles {
    # 検索対象インデックスのTSVを列挙し、@{ Folders; Files } を返す。
    #   folders: 検索対象インデックスのフォルダ（文字列。フォルダ以下すべて）、または
    #            @{ Root（インデックスのフォルダ）; RelPath（その中のフォルダ。空は Root 自身）; Recurse（$false は直下のファイルだけ） }
    #            （画面の検索対象ツリーで一部のフォルダだけを選んだとき。結果の相対パスは Root から求める）
    #   Folders: フォルダごとの @{ Path（指定どおり。RelPath があれば Root\RelPath）; Root（フルパス）; Exists; Count }
    #   Files  : TSVのフルパス → @{ Root; RelPath（インデックスフォルダからの相対パス）; LongPath（\\?\ 付きのパス）; Ticks; Size（更新日時・サイズ） }。パス順。入れ子のフォルダでも重複しない
    #   onProgress: 数えた件数を知らせる { param($count) }（TSVが多いと数秒かかるため、画面が「確認中… N 件」を出せるようにする）
    param (
        [object[]]$folders = @(${indexDir}),
        [scriptblock]$onProgress = $null
    )

    $folderInfo = New-Object System.Collections.Generic.List[object]
    $files = @{}
    $scanned = 0
    $notifyEvery = 2000
    foreach ($target in $folders) {
        if ($target -is [string]) {
            $target = @{ Root = $target; RelPath = ""; Recurse = $true }
        }
        $relPath = ([string]$target.RelPath).Trim("\")
        $dir = if ($relPath) { "$(([string]$target.Root).TrimEnd('\'))\${relPath}" } else { [string]$target.Root }
        $exists = Test-Path -LiteralPath $target.Root -PathType Container
        if ($exists) {
            $root = (Resolve-Path -LiteralPath $target.Root).ProviderPath.TrimEnd("\")
            $fullDir = if ($relPath) { "${root}\${relPath}" } else { $root }
            $exists = [System.IO.Directory]::Exists((toLongPath $fullDir))
        }
        if (!$exists) {
            $folderInfo.Add(@{ Path = $dir; Root = ""; Exists = $false; Count = 0 })
            continue
        }

        # 長いパス（フォルダが約248文字超）の中も列挙できるよう \\?\ 付きで列挙し、キーは \\?\ の無い通常のパスにする。
        # 件数が多いと検索を始めるまでの待ち時間になるため、1件ずつオブジェクトを作る Get-ChildItem ではなく
        # .NET の列挙（文字列）を使い、相対パスも関数呼び出し無しで切り出す（TSV 2 万件で約 6 秒 → 約 2 秒）
        $longDir = toLongPath $fullDir
        if ([bool]$target.Recurse) {
            $found = findTsvFilesParallel $longDir
        } else {
            $found = New-Object System.Collections.Generic.List[System.IO.FileInfo]
            $found.AddRange([System.IO.DirectoryInfo]::new($longDir).GetFiles("*.tsv", [System.IO.SearchOption]::TopDirectoryOnly))
            # 「フォルダ直下のファイル」には、元のファイル名のフォルダ（<ファイル名.xlsx>\<場所>.tsv）の中のTSVも含める
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories($longDir)) {
                if ([System.IO.Path]::GetFileName($sub) -notmatch ${indexBookDirPattern}) {
                    continue
                }
                $found.AddRange([System.IO.DirectoryInfo]::new($sub).GetFiles("*.tsv", [System.IO.SearchOption]::TopDirectoryOnly))
            }
        }

        $folderInfo.Add(@{ Path = $dir; Root = $root; Exists = $true; Count = $found.Count })
        $rootLength = $root.Length + 1
        # 列挙したパスはどれも longDir で始まるため、\\?\ を外した形は先頭を差し替えて作る（1件ずつ fromLongPath を呼ばない）
        $plainDir = fromLongPath $longDir
        $longDirLength = $longDir.Length
        foreach ($file in $found) {
            $path = $file.FullName
            $fullName = $plainDir + $path.Substring($longDirLength)
            $relative = if ($fullName.Length -gt $rootLength) { $fullName.Substring($rootLength) } else { [System.IO.Path]::GetFileName($fullName) }
            # LongPath: 検索で開くときのパス（searchIndex が1件ずつ toLongPath しなくて済むよう、列挙したパスを持たせる）
            # Ticks・Size: 更新日時（UTC）とサイズ。検索で読み込んだ内容を再利用してよいかの判定に使う（searchTsvFiles）
            $files[$fullName] = @{ Root = $root; RelPath = $relative; LongPath = $path; Ticks = $file.LastWriteTimeUtc.Ticks; Size = $file.Length }
            $scanned++
            if ($onProgress -and ($scanned % $notifyEvery) -eq 0) {
                & $onProgress $scanned
            }
        }
    }

    # Sort-Object と同じ順（現在のカルチャ・大文字と小文字を区別しない）。Sort-Object は件数が多いと遅いため .NET で並べる
    $keys = [string[]]@($files.Keys)
    [System.Array]::Sort($keys, [System.StringComparer]::CurrentCultureIgnoreCase)
    $sorted = [ordered]@{}
    foreach ($path in $keys) {
        $sorted[$path] = $files[$path]
    }
    return @{ Folders = $folderInfo.ToArray(); Files = $sorted }
}

function testIndexExists {
    # 検索対象インデックスにTSVが1件でもあるか（最初の1件が見つかった時点で打ち切る）
    param (
        [string[]]$folders = @(${indexDir})
    )

    foreach ($dir in $folders) {
        if (!(Test-Path -LiteralPath $dir -PathType Container)) {
            continue
        }
        try {
            $enumerator = [System.IO.Directory]::EnumerateFiles((toLongPath (Resolve-Path -LiteralPath $dir).ProviderPath), "*.tsv", [System.IO.SearchOption]::AllDirectories).GetEnumerator()
            if ($enumerator.MoveNext()) {
                return $true
            }
        } catch {
            # アクセスできないフォルダがある場合は、件数を数える方で判定する
            if (@(Get-ChildItem -LiteralPath $dir -Filter "*.tsv" -File -Recurse -ErrorAction SilentlyContinue).Count -gt 0) {
                return $true
            }
        }
    }
    return $false
}

function getIndexSummary {
    # 検索対象インデックスのTSVの件数と最新の更新日時を返す: @{ Count; LastWrite（無ければ $null）; Missing（存在しないフォルダ） }
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
        # TSV が多いと Get-ChildItem では数秒かかるため、検索と同じ .NET の列挙で数える。
        # アクセスできないフォルダがあると .NET の列挙は途中で止まるため、そのときは Get-ChildItem で数える（読めるものだけ）
        try {
            $found = findTsvFilesParallel $longDir
        } catch {
            $found = @(Get-ChildItem -LiteralPath $longDir -Filter "*.tsv" -File -Recurse -ErrorAction SilentlyContinue)
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

function searchIndex {
    # インデックスのTSVをワードで検索し、ヒットした行を返す（画面の検索処理）。
    #   tsvFiles     : getIndexTsvFiles の Files
    #   simpleMatch  : $true なら文字どおりに検索する。$false なら正規表現として検索し、正規表現として不正なら文字どおりに検索する
    #   limit        : 件数の上限（0 は上限なし）。超えたら打ち切る
    #   onProgress   : chunkSize 件のTSVを検索するたびに呼ぶ { param($done, $total, $newHits) }
    #   shouldStop   : $true を返すと中止する
    #   caseSensitive: 英字の大文字・小文字を区別する（newSearchRegex）
    #   fileFilter   : 対象ファイル（newFileFilter）。元のファイル名が一致しないTSVは検索しない
    #   workerCount  : 並列に検索するスレッドの数（0 は TSV の数と CPU のコア数から決める）
    #   cache        : 読んだ TSV の内容を次の検索で使い回す入れ物（newTsvTextCache。$null は使い回さない）
    #   includeShapes / includeComments: 図形・コメントの場所（"<シート名>[図形]" 等）の TSV も検索する（newPlaceExclude）
    # @{ Hits; SimpleMatch（実際に文字どおり検索したか）; Total（対象ファイルで絞った後のTSVの数）; Truncated; Cancelled } を返す。
    # Hits の各要素は PSCustomObject（Root; RelPath; RelDir; FileName; Book; Location; LineNumber; Line）
    param (
        [string]$word,
        $tsvFiles,
        [bool]$simpleMatch = $false,
        [int]$limit = 0,
        [int]$chunkSize = 100,
        [scriptblock]$onProgress = $null,
        [scriptblock]$shouldStop = $null,
        [bool]$caseSensitive = $false,
        [string]$fileFilter = "",
        [int]$workerCount = 0,
        $cache = $null,
        [bool]$includeShapes = $true,
        [bool]$includeComments = $true
    )

    $search = newSearchRegex $word $simpleMatch $caseSensitive
    $filter = newFileFilter $fileFilter

    $files = newTsvFiles $tsvFiles $filter.Include $filter.Exclude (newPlaceExclude $includeShapes $includeComments)

    $hits = New-Object System.Collections.Generic.List[psobject]
    $result = @{ Hits = $hits; SimpleMatch = $search.SimpleMatch; Total = $files.Count; Truncated = $false; Cancelled = $false }
    $timeoutMessage = "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"

    # TSV が多ければ、chunkSize 件ずつを別スレッドで並行して検索する（TSV の読み込み・照合は CPU を使うため、コア数に応じて速くなる）。
    # 結果は TSV の順に取り込み、上限・中止・進み具合の扱いは 1 スレッドのときと同じにする。
    # 少ないときは、スレッドを用意する時間の方が長いため 1 スレッドで検索する
    $workers = if ($workerCount -gt 0) { $workerCount } else { [Math]::Min([Environment]::ProcessorCount, ${searchWorkerMax}) }
    if ($workerCount -le 0 -and $files.Count -lt [Math]::Max($chunkSize * 2, ${searchParallelMin})) {
        $workers = 1
    }
    $pool = $null
    $pending = New-Object System.Collections.Generic.Queue[hashtable]
    try {
        if ($workers -gt 1) {
            $pool = newSearchWorkerPool $workers
        }
        $next = 0
        while ($next -lt $files.Count -or $pending.Count -gt 0) {
            if ($shouldStop -and (& $shouldStop)) {
                $result.Cancelled = $true
                break
            }

            # 上限があれば、残りの件数を超えた時点で止める（残りちょうどで終われば打ち切りにしない）
            if ($pool) {
                # 先の TSV の分から順に、スレッド数の 2 倍まで検索を始めておく（先に始めた分ほど上限が緩いが、取り込むときに切る）
                while ($pending.Count -lt $workers * 2 -and $next -lt $files.Count) {
                    $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
                    $ps = [powershell]::Create()
                    $ps.RunspacePool = $pool
                    [void]$ps.AddScript(${searchWorkerScript}).AddArgument($files).AddArgument($next).AddArgument($chunkSize).AddArgument($search.Regex).AddArgument($max).AddArgument($search.TextRegex).AddArgument($search.ScanMode).AddArgument($cache)
                    $pending.Enqueue(@{ PowerShell = $ps; Handle = $ps.BeginInvoke(); Start = $next })
                    $next += $chunkSize
                }
                $job = $pending.Dequeue()
                try {
                    $output = $job.PowerShell.EndInvoke($job.Handle)
                } finally {
                    $job.PowerShell.Dispose()
                }
                $answer = $output[0]
                if ($answer.Timeout) {
                    throw $timeoutMessage
                }
                $newHits = $answer.Hits
                $start = $job.Start
            } else {
                $start = $next
                $next += $chunkSize
                $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
                try {
                    $newHits = searchTsvFiles $files $start $chunkSize $search.Regex $max $search.TextRegex $search.ScanMode $cache
                } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                    throw $timeoutMessage
                } catch {
                    if ($_.Exception.InnerException -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) {
                        throw $timeoutMessage
                    }
                    throw
                }
            }

            $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
            if ($max -ge 0 -and $newHits.Count -gt $max) {
                $newHits.RemoveRange($max, $newHits.Count - $max)
                $result.Truncated = $true
            }
            $hits.AddRange($newHits)

            if ($onProgress) {
                & $onProgress ([math]::Min($start + $chunkSize, $files.Count)) $files.Count $newHits
            }
            if ($result.Truncated) {
                break
            }
        }
    } finally {
        # 打ち切り・中止・例外で残った検索は止める
        foreach ($job in $pending) {
            try { $job.PowerShell.Stop() } catch {}
            $job.PowerShell.Dispose()
        }
        if ($pool) {
            $pool.Dispose()
        }
    }
    return $result
}

# 並列検索のスレッド数の上限と、並列にする TSV の数の下限
${searchWorkerMax} = 8
${searchParallelMin} = 400

# 並列検索の各スレッドで動かすスクリプト（searchTsvFiles を呼ぶ）。@{ Hits; Timeout（照合が時間切れ） } を返す
${searchWorkerScript} = {
    param ($files, $start, $count, $regex, $max, $textRegex, $scanMode, $cache)
    # 出力に List をそのまま出すと 1 件ずつに分かれるため、ハッシュテーブルに入れて返す
    try {
        $hits = searchTsvFiles $files $start $count $regex $max $textRegex $scanMode $cache
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

function newSearchWorkerPool {
    # 並列検索のスレッド（RunspacePool）を用意する。各スレッドには searchTsvFiles だけを読み込む
    # （lib.ps1 全体を読み込むと、スレッドを用意するだけで時間がかかるため）
    param (
        [int]$workers
    )

    $state = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $state.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new("searchTsvFiles", ${function:searchTsvFiles}.ToString()))
    $state.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new("searchWholeFileMax", ${searchWholeFileMax}, ""))
    $pool = [runspacefactory]::CreateRunspacePool(1, $workers, $state, $Host)
    $pool.Open()
    return $pool
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
        # ファイル名の前に、変換対象フォルダからの相対フォルダを付ける
        $line = toResultLine $hit.Book $hit.Location $hit.LineNumber $hit.Line
        if ($hit.RelDir) {
            $line = "$($hit.RelDir)\${line}"
        }
        $lines.Add($line)
        $columnCount = [math]::Max($columnCount, (countTsvFields $line) - 3)  # ファイル名・場所・行番号の3列を除く
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
