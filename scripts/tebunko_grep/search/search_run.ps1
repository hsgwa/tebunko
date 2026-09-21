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
    param (
        [string[]]$paths,
        [string[]]$roots,
        [string[]]$relPaths,
        [regex]$include = $null,
        [regex]$exclude = $null
    )

    $files = New-Object System.Collections.Generic.List[psobject]
    for ($i = 0; $i -lt $paths.Length; $i++) {
        $relPath = $relPaths[$i]
        $parts = splitIndexTsvPath $relPath
        $book = [string]$parts.Book
        if ($include -and !$include.IsMatch($book)) { continue }
        if ($exclude -and $exclude.IsMatch($book)) { continue }
        $files.Add([pscustomobject]@{
            Path     = $paths[$i]
            Root     = $roots[$i]
            RelPath  = $relPath
            RelDir   = [string]$parts.RelDir
            FileName = [System.IO.Path]::GetFileName($relPath)
            Book     = $book
            Location = [string]$parts.Place
        })
    }
    return , $files
}

function searchTsvFiles {
    # files の start から count 件を読み、regex に一致する行を PSCustomObject で返す（1行に複数一致しても1件）。
    # max 以上（max+1 件目）が見つかった時点で打ち切る（負は上限なし）。読めないTSV（変換中に削除された等）は飛ばす。
    # 変換中のTSVも読めるよう共有モードは ReadWrite|Delete にする。正規表現の照合が時間切れ（RegexMatchTimeoutException）なら
    # 例外はそのまま呼び出し元（searchIndex）へ伝わる。
    param (
        $files,
        [int]$start,
        [int]$count,
        [regex]$regex,
        [int]$max
    )

    $hits = New-Object System.Collections.Generic.List[psobject]
    $end = [Math]::Min($files.Count, $start + $count)
    for ($i = $start; $i -lt $end; $i++) {
        $f = $files[$i]
        $reader = $null
        try {
            $stream = New-Object System.IO.FileStream($f.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
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

function getIndexTsvFiles {
    # 検索対象インデックスのTSVを列挙し、@{ Folders; Files } を返す。
    #   folders: 検索対象インデックスのフォルダ（文字列。フォルダ以下すべて）、または
    #            @{ Root（インデックスのフォルダ）; RelPath（その中のフォルダ。空は Root 自身）; Recurse（$false は直下のファイルだけ） }
    #            （画面の検索対象ツリーで一部のフォルダだけを選んだとき。結果の相対パスは Root から求める）
    #   Folders: フォルダごとの @{ Path（指定どおり。RelPath があれば Root\RelPath）; Root（フルパス）; Exists; Count }
    #   Files  : TSVのフルパス → @{ Root; RelPath（インデックスフォルダからの相対パス） }。パス順。入れ子のフォルダでも重複しない
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
        $found = New-Object System.Collections.Generic.List[string]
        $option = if ([bool]$target.Recurse) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }
        foreach ($path in [System.IO.Directory]::EnumerateFiles($longDir, "*.tsv", $option)) {
            $found.Add($path)
        }
        if (!$target.Recurse) {
            # 「フォルダ直下のファイル」には、元のファイル名のフォルダ（<ファイル名.xlsx>\<場所>.tsv）の中のTSVも含める
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories($longDir)) {
                if ([System.IO.Path]::GetFileName($sub) -notmatch ${indexBookDirPattern}) {
                    continue
                }
                foreach ($path in [System.IO.Directory]::EnumerateFiles($sub, "*.tsv", [System.IO.SearchOption]::TopDirectoryOnly)) {
                    $found.Add($path)
                }
            }
        }

        $folderInfo.Add(@{ Path = $dir; Root = $root; Exists = $true; Count = $found.Count })
        $rootLength = $root.Length + 1
        foreach ($path in $found) {
            $fullName = fromLongPath $path
            $relative = if ($fullName.Length -gt $rootLength) { $fullName.Substring($rootLength) } else { [System.IO.Path]::GetFileName($fullName) }
            $files[$fullName] = @{ Root = $root; RelPath = $relative }
            $scanned++
            if ($onProgress -and ($scanned % $notifyEvery) -eq 0) {
                & $onProgress $scanned
            }
        }
    }

    $sorted = [ordered]@{}
    foreach ($path in @($files.Keys | Sort-Object)) {
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
        foreach ($file in @(Get-ChildItem -LiteralPath (toLongPath (Resolve-Path -LiteralPath $dir).ProviderPath) -Filter "*.tsv" -File -Recurse -ErrorAction SilentlyContinue)) {
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
        [string]$fileFilter = ""
    )

    $search = newSearchRegex $word $simpleMatch $caseSensitive
    $filter = newFileFilter $fileFilter

    $count = $tsvFiles.Count
    $paths = New-Object string[] $count
    $roots = New-Object string[] $count
    $relPaths = New-Object string[] $count
    $i = 0
    foreach ($path in $tsvFiles.Keys) {
        $info = $tsvFiles[$path]
        # 260文字を超えるパスのTSVも読めるよう \\?\ 付きで読む
        $paths[$i] = toLongPath $path
        $roots[$i] = $info.Root
        $relPaths[$i] = $info.RelPath
        $i++
    }
    $files = newTsvFiles $paths $roots $relPaths $filter.Include $filter.Exclude

    $hits = New-Object System.Collections.Generic.List[psobject]
    $result = @{ Hits = $hits; SimpleMatch = $search.SimpleMatch; Total = $files.Count; Truncated = $false; Cancelled = $false }

    for ($i = 0; $i -lt $files.Count; $i += $chunkSize) {
        if ($shouldStop -and (& $shouldStop)) {
            $result.Cancelled = $true
            break
        }

        # 上限があれば、残りの件数を超えた時点で止める（残りちょうどで終われば打ち切りにしない）
        $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
        try {
            $newHits = searchTsvFiles $files $i $chunkSize $search.Regex $max
        } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
            throw "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"
        } catch {
            if ($_.Exception.InnerException -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) {
                throw "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"
            }
            throw
        }
        if ($max -ge 0 -and $newHits.Count -gt $max) {
            $newHits.RemoveRange($max, $newHits.Count - $max)
            $result.Truncated = $true
        }
        $hits.AddRange($newHits)

        if ($onProgress) {
            & $onProgress ([math]::Min($i + $chunkSize, $files.Count)) $files.Count $newHits
        }
        if ($result.Truncated) {
            break
        }
    }
    return $result
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
