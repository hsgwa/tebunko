# 検索用のまとめファイル（フォルダ 1 つ・元のファイルの拡張子 1 つにつき 1 つ。本文.xlsx.tsv など）の形式（判断層）。
# ファイル（元のファイル）ごと・場所（シート・ページ・スライドなど）ごとに、メタ情報の行と今の TSV の中身を並べる。
#
#   ␞ 版=1
#   ␞ ファイル名=(株)山田商事_見積書.xlsx
#   ␞ 種類=Excel
#   ␞ シート=見積
#   ␞ 対象=本文
#   品名→数量→単価        ← 今の TSV の中身そのまま（→ はタブ）。この行が 1 行目
#
# ・メタ情報の行は「RS（U+001E）、半角スペース、キー=値」。1 行に 1 項目。
#   ファイル名= で新しいファイル、シート= ページ= スライド= 部分= で新しい場所が始まる。知らないキーは読み飛ばす
# ・中身には U+001C〜U+001F を入れない（書くときに取り除く）ため、行の先頭が RS ならメタ情報の行
# ・改行は LF にそろえる。行の分け方は StreamReader.ReadLine と同じ（CRLF・LF・CR）にし、行番号を変えない
# ・文字コードは UTF-16LE（BOM 付き。pack_store.ps1 が読み書きする）

# まとめファイルの名前は「本文.<元のファイルの拡張子（小文字）>.tsv」。_ は使わない（以前の形式 <ブック>_<場所>.tsv と区別するため）
${packFilePattern} = "本文.*.tsv"
${packVersion} = 1
${packMark} = [char]0x1E

# 場所のまとまりを始めるキー
${packPlaceKeys} = @("シート", "ページ", "スライド", "部分")


function getPackFileKind {
    # 元のファイル名から種類（Excel・Word・PowerPoint）を返す。分からなければ空
    param (
        [string]$book
    )

    if ($book -match '\.xls[a-z]?$') { return "Excel" }
    if ($book -match '\.doc[a-z]?$') { return "Word" }
    if ($book -match '\.ppt[a-z]?$') { return "PowerPoint" }
    return ""
}

function getPackExtension {
    # 元のファイル名から、まとめファイルを分ける拡張子（小文字・先頭の . なし）を返す
    param (
        [string]$book
    )

    return [System.IO.Path]::GetExtension($book).TrimStart(".").ToLowerInvariant()
}


function getPackFileName {
    # 拡張子のまとめファイルの名前（本文.xlsx.tsv など）を返す
    param (
        [string]$extension
    )

    return "本文.{0}.tsv" -f $extension
}


function splitPackBooksByExtension {
    # 元のファイルの並びを、拡張子ごとの並び（[ordered] 拡張子 → 並び。拡張子は現れた順）に分ける。各並びの中の順は変えない
    param (
        $books
    )

    $groups = [ordered]@{}
    foreach ($book in $books) {
        $extension = getPackExtension ([string]$book.Name)
        if (!$groups.Contains($extension)) {
            $groups[$extension] = New-Object System.Collections.Generic.List[object]
        }
        $groups[$extension].Add($book)
    }
    return $groups
}


function encodePackValue {
    # メタ情報の値を 1 行に収まる形にする。制御文字（タブを除く）と % を "%XX" にする（decodePackValue で戻す）
    param (
        [string]$value
    )

    if ($value -notmatch '[\x00-\x08\x0A-\x1F%]') {
        return $value
    }
    return [regex]::Replace($value, '[\x00-\x08\x0A-\x1F%]', { param($m) '%{0:X2}' -f [int][char]$m.Value })
}


function decodePackValue {
    # encodePackValue で符号化した値を戻す
    param (
        [string]$value
    )

    if ($value.IndexOf("%") -lt 0) {
        return $value
    }
    return [regex]::Replace($value, '%(?:[01][0-9A-F]|25)', { param($m) [string][char][Convert]::ToInt32($m.Value.Substring(1), 16) })
}


function convertPlaceToPackMeta {
    # 今の場所の名前（TSV のファイル名。"見積[図形]"・"ページ001"・"スライド002（非表示）"・"スライド001_ノート" など）を、
    # 場所のメタ情報（[ordered] キー → 値）にする。convertPackMetaToPlace で必ず元の名前に戻る形にし、
    # 分けられない名前は「部分=名前そのまま・対象=本文」にする
    param (
        [string]$book,
        [string]$place
    )

    $split = splitObjectPlace $place
    $target = if ($split.Kind) { $split.Kind } else { "本文" }
    $meta = [ordered]@{}
    switch (getPackFileKind $book) {
        "Excel" {
            $meta["シート"] = $split.Base
            $meta["対象"] = $target
        }
        "Word" {
            if ($split.Base -match '^ページ(\d+)$') {
                $meta["ページ"] = [string][int]$Matches[1]
            } else {
                $meta["部分"] = $split.Base
            }
            $meta["対象"] = $target
        }
        "PowerPoint" {
            if (!$split.Kind -and $split.Base -match '^スライド(\d+)_ノート$') {
                $meta["スライド"] = [string][int]$Matches[1]
                $meta["対象"] = "ノート"
            } elseif ($split.Base -match '^スライド(\d+)(（非表示）)?$') {
                $meta["スライド"] = [string][int]$Matches[1]
                $meta["対象"] = $target
                if ($Matches[2]) { $meta["非表示"] = "はい" }
            } else {
                $meta["部分"] = $split.Base
                $meta["対象"] = $target
            }
        }
        default {
            $meta["部分"] = $split.Base
            $meta["対象"] = $target
        }
    }
    if ((convertPackMetaToPlace $meta) -cne $place) {
        # 番号の桁が違う（"ページ1"）など、組み立て直すと同じ名前にならないものは、名前をそのまま持つ
        $meta = [ordered]@{ "部分" = $place; "対象" = "本文" }
    }
    return $meta
}


function convertPackMetaToPlace {
    # 場所のメタ情報から、今の場所の名前（画面の表示・図形とコメントの除外・元のファイルを開く処理が使う形）を組み立てる
    param (
        $meta
    )

    $target = [string]$meta["対象"]
    $suffix = if ($target -eq ${placeKindShape} -or $target -eq ${placeKindComment}) { "[$target]" } else { "" }
    if ($meta.Contains("シート")) {
        return [string]$meta["シート"] + $suffix
    }
    if ($meta.Contains("ページ")) {
        return ("ページ{0:D3}" -f [int]$meta["ページ"]) + $suffix
    }
    if ($meta.Contains("スライド")) {
        $base = "スライド{0:D3}" -f [int]$meta["スライド"]
        if ($target -eq "ノート") {
            return "${base}_ノート"
        }
        if ($meta["非表示"] -eq "はい") {
            $base += "（非表示）"
        }
        return $base + $suffix
    }
    if ($target -eq "本文") {
        return [string]$meta["部分"]
    }
    return [string]$meta["部分"] + $suffix
}


function convertToPackBody {
    # TSV の中身を、まとめファイルに入れる形（LF 区切り・末尾に LF・U+001C〜U+001F を除く）にする。
    # 行の分け方は StreamReader.ReadLine と同じ（CRLF・LF・CR で分け、末尾の改行の後ろは行にしない）。中身が空なら空を返す
    param (
        [string]$text
    )

    if ($text.Length -eq 0) {
        return ""
    }
    $body = $text.Replace("`r`n", "`n").Replace("`r", "`n")
    if ($body.IndexOfAny([char[]]@([char]0x1C, [char]0x1D, [char]0x1E, [char]0x1F)) -ge 0) {
        $body = [regex]::Replace($body, '[\x1C-\x1F]', "")
    }
    if ($body.EndsWith("`n")) {
        return $body
    }
    return $body + "`n"
}


function convertToPackText {
    # フォルダ直下の元のファイルの中身から、まとめファイルの文字列を作る。
    #   books: 次のどちらかの並び（この順に書く）
    #     @{ Name（元のファイル名）; Places（@{ Place（今の場所の名前）; Text（TSV の中身） } の並び） } … TSV から新しく作る
    #     @{ Name; Block（splitPackTextByBook で取り出した、そのファイルのまとまり） }       … 前のまとめファイルからそのまま写す
    param (
        $books
    )

    $mark = [string]${packMark}
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("$mark 版=${packVersion}`n")
    foreach ($book in $books) {
        if ($null -ne $book.Block) {
            [void]$sb.Append([string]$book.Block)
            continue
        }
        $name = [string]$book.Name
        [void]$sb.Append("$mark ファイル名=").Append((encodePackValue $name)).Append("`n")
        $kind = getPackFileKind $name
        if ($kind) {
            [void]$sb.Append("$mark 種類=$kind`n")
        }
        foreach ($place in $book.Places) {
            $meta = convertPlaceToPackMeta $name ([string]$place.Place)
            foreach ($key in $meta.Keys) {
                [void]$sb.Append("$mark ").Append($key).Append("=").Append((encodePackValue ([string]$meta[$key]))).Append("`n")
            }
            [void]$sb.Append((convertToPackBody ([string]$place.Text)))
        }
    }
    return $sb.ToString()
}


function readPackPlaces {
    # まとめファイルの文字列を読み、場所ごとの @{ Book; Location（今の場所の名前）; Start（中身の先頭の位置）; End（中身の終わりの次の位置） }
    # を先頭から順に返す。版が分からないときは例外にする
    param (
        [string]$text
    )

    $places = New-Object System.Collections.Generic.List[hashtable]
    $mark = ${packMark}
    $placeKeys = ${packPlaceKeys}
    $book = ""
    $current = $null
    $meta = $null
    $versionSeen = $false
    $pos = $text.IndexOf($mark)
    while ($pos -ge 0) {
        $lineEnd = $text.IndexOf([char]10, $pos)
        if ($lineEnd -lt 0) { $lineEnd = $text.Length }
        $line = $text.Substring($pos + 2, [Math]::Max(0, $lineEnd - $pos - 2))
        $eq = $line.IndexOf("=")
        if ($eq -gt 0) {
            $key = $line.Substring(0, $eq)
            $value = decodePackValue $line.Substring($eq + 1)
            if ($key -eq "版") {
                if ($value -ne [string]${packVersion}) {
                    throw "まとめファイルの版が違います（$value）"
                }
                $versionSeen = $true
            } elseif ($key -eq "ファイル名" -or $placeKeys -contains $key) {
                if ($current) {
                    $current.End = $pos
                    $current.Location = convertPackMetaToPlace $meta
                    $places.Add($current)
                    $current = $null
                }
                if ($key -eq "ファイル名") {
                    $book = $value
                } else {
                    $meta = [ordered]@{ $key = $value }
                    $current = @{ Book = $book; Location = ""; Start = $lineEnd + 1; End = $lineEnd + 1 }
                }
            } elseif ($current) {
                $meta[$key] = $value
                $current.Start = $lineEnd + 1
            }
        }
        $pos = if ($lineEnd -lt $text.Length) { $text.IndexOf($mark, $lineEnd + 1) } else { -1 }
    }
    if ($current) {
        $current.End = $text.Length
        $current.Location = convertPackMetaToPlace $meta
        $places.Add($current)
    }
    if (!$versionSeen -and $text.Length -gt 0) {
        throw "まとめファイルの版がありません"
    }
    return , $places
}


function splitPackTextByBook {
    # まとめファイルの文字列を、元のファイルごとのまとまり（@{ Name; Block }。Block は「ファイル名=」の行から次の「ファイル名=」の行の前まで）
    # に分けて先頭から順に返す（元のファイルを入れ替えるとき、変わらないファイルをそのまま写すため）。版が違えば例外にする
    param (
        [string]$text
    )

    $books = New-Object System.Collections.Generic.List[hashtable]
    if ($text.Length -eq 0) {
        return , $books
    }
    $head = "$([string]${packMark}) ファイル名="
    $versionLine = "$([string]${packMark}) 版="
    $firstEnd = $text.IndexOf([char]10)
    if (!$text.StartsWith($versionLine, [System.StringComparison]::Ordinal) -or $firstEnd -lt 0 -or $text.Substring($versionLine.Length, $firstEnd - $versionLine.Length) -ne [string]${packVersion}) {
        throw "まとめファイルの版が違います"
    }
    $pos = $text.IndexOf("`n$head", [System.StringComparison]::Ordinal)
    while ($pos -ge 0) {
        $start = $pos + 1
        $lineEnd = $text.IndexOf([char]10, $start)
        $name = decodePackValue $text.Substring($start + $head.Length, $lineEnd - $start - $head.Length)
        $next = $text.IndexOf("`n$head", $lineEnd, [System.StringComparison]::Ordinal)
        $end = if ($next -ge 0) { $next + 1 } else { $text.Length }
        $books.Add(@{ Name = $name; Block = $text.Substring($start, $end - $start) })
        $pos = $next
    }
    return , $books
}