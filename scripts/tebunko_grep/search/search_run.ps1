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


# 検索で読んだまとめファイルの内容を残しておく量の上限（文字数。1 文字 2 バイトのため約 128MB）
${searchCacheMaxChars} = 64000000

function newTsvTextCache {
    # 検索で読んだまとめファイルの内容を、次の検索で使い回すための入れ物を作る（画面が 1 つ持ち、検索のたびに searchPackIndex に渡す）。
    # 2 回目以降の検索ではファイルを開かない。更新日時・サイズが列挙したときと違うまとめファイル（書き直した等）は読み直す。
    # 並列検索の各スレッドから使うため、中身は ConcurrentDictionary。
    #   Texts: まとめファイルの \\?\ 付きのパス → @(更新日時（UTC の Ticks）, サイズ, 内容, 場所の一覧) / Chars: 残している文字数 / MaxChars: 上限
    param (
        [long]$maxChars = ${searchCacheMaxChars}
    )

    return @{
        Texts    = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        Chars    = [long[]]::new(1)
        MaxChars = $maxChars
    }
}

function testIndexExists {
    # 検索対象インデックスにまとめファイルが1件でもあるか（最初の1件が見つかった時点で打ち切る）
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
    # 検索対象インデックスのまとめファイルの件数と最新の更新日時を返す: @{ Count; LastWrite（無ければ $null）; Missing（存在しないフォルダ） }
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
