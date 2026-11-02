# 2-gram の索引（gram 索引）の決まり（判断層。ファイルにも画面にも触らない）。
# 検索のたびに本文インデックスのファイルを全部読まないよう、本文インデックスのファイルごとの 2-gram を自前の索引に持ち、
# 検索語のすべての 2-gram を含むファイルだけを照合する（search/search_gram.ps1 の getSearchGramValues）。
# 検索語を含むファイルには、検索語のすべての 2-gram が必ず入っているため、絞っても結果は変わらない。
# 索引が使えないとき・ファイルが変わったとき・索引に無いファイルは、必ず照合する側に倒す（漏れさせない）。
# ファイルの読み書きは gram_index_store.ps1（状態層）、検索への組み込みは gram_search.ps1 が行う。
#
# 索引はインデックス名ごとのフォルダ（store\gram\<インデックス名>\）に、区画（segment）のファイルと目録（toc.tsv）を置く。
# 区画は 1 回のインデックス作成で増えた分を 1 つずつ足したもので、区画が多くなったら作り直して 1 つにまとめる。
# 同じファイルが新しい区画にも載っていれば、新しい区画の方が正しい（古い区画の行は使わない）。

# 区画のファイルの先頭と末尾の印（リトルエンディアンの 4 バイト。"TBGR"・"TBGE"）と、形式の版
${gramSegmentMagic}  = 0x52474254
${gramSegmentFooter} = 0x45474254
${gramSegmentVersion} = 1
${gramSegmentHeaderBytes} = 24

# 目録の名前と、1 行目の印
${gramTocFileName} = "toc.tsv"
${gramTocMark} = "tebunko-gram"
${gramSegmentFilePattern} = "g*.seg"

# 区画の数の上限（超えたらまとめ直す）と、1 つの区画に入れる 2-gram の組の数の上限（作るときのメモリを抑える）
${gramSegmentLimit} = 8
${gramBatchPostings} = 12000000

function testGramIndexUsable {
    # gram 索引で絞れるか。検索語（正規表現なら必ず含まれる文字）から 2-gram を作れるとき（grams は getSearchGramValues の結果）だけ $true
    param (
        [uint32[]]$grams
    )

    return ($grams.Count -gt 0)
}

function getGramSegmentLayout {
    # 区画のファイルの中身の置き場所（先頭からのバイト数）と全体の長さを返す。
    # 並び: 先頭（24 バイト）・更新日時（int64 × ファイル数）・大きさ（int64 × ファイル数）・ファイルの相対パス（UTF-8。改行区切り）・
    #       2-gram の値（uint32 × 2-gram の数。昇順）・組の始まり（int32 × (2-gram の数 + 1)）・ファイルの番号（int32 × 組の数）・末尾の印（4 バイト）
    param (
        [long]$fileCount,
        [long]$nameBytes,
        [long]$gramCount,
        [long]$postingCount
    )

    $ticksPos = ${gramSegmentHeaderBytes}
    $sizesPos = $ticksPos + 8 * $fileCount
    $namesPos = $sizesPos + 8 * $fileCount
    $gramsPos = $namesPos + $nameBytes
    $offsetsPos = $gramsPos + 4 * $gramCount
    $postingsPos = $offsetsPos + 4 * ($gramCount + 1)
    $footerPos = $postingsPos + 4 * $postingCount
    return @{
        TicksPos = $ticksPos; SizesPos = $sizesPos; NamesPos = $namesPos; GramsPos = $gramsPos
        OffsetsPos = $offsetsPos; PostingsPos = $postingsPos; FooterPos = $footerPos; Length = $footerPos + 4
    }
}

function convertToGramSegmentHeader {
    # 区画のファイルの先頭の 24 バイトを作る
    param (
        [int]$fileCount,
        [int]$nameBytes,
        [int]$gramCount,
        [int]$postingCount
    )

    $bytes = New-Object 'byte[]' ${gramSegmentHeaderBytes}
    $values = [int[]]@([int]${gramSegmentMagic}, ${gramSegmentVersion}, $fileCount, $nameBytes, $gramCount, $postingCount)
    [System.Buffer]::BlockCopy($values, 0, $bytes, 0, ${gramSegmentHeaderBytes})
    return , $bytes
}

function convertFromGramSegmentHeader {
    # 区画のファイルの先頭の 24 バイトを読む。@{ FileCount; NameBytes; GramCount; PostingCount; Layout } を返す。
    # 印・版が違う・数が負・短いときは $null（読めない区画。呼び出し側は安全側に倒す）
    param (
        [byte[]]$bytes
    )

    if ($null -eq $bytes -or $bytes.Length -lt ${gramSegmentHeaderBytes}) {
        return $null
    }
    $values = New-Object 'int[]' 6
    [System.Buffer]::BlockCopy($bytes, 0, $values, 0, ${gramSegmentHeaderBytes})
    if ($values[0] -ne [int]${gramSegmentMagic} -or $values[1] -ne ${gramSegmentVersion}) {
        return $null
    }
    for ($i = 2; $i -lt 6; $i++) {
        if ($values[$i] -lt 0) {
            return $null
        }
    }
    return @{
        FileCount = $values[2]; NameBytes = $values[3]; GramCount = $values[4]; PostingCount = $values[5]
        Layout = (getGramSegmentLayout $values[2] $values[3] $values[4] $values[5])
    }
}

function testGramSegmentIntact {
    # 区画のファイルが壊れていないか: 先頭が読めて、ファイルの長さが先頭から求めた長さと目録に書いた長さに合い、末尾の印が正しい
    #   header      : convertFromGramSegmentHeader の結果
    #   actualLength: ファイルの長さ
    #   tocLength   : 目録に書いてある長さ
    #   footer      : 末尾の 4 バイトの値
    param (
        $header,
        [long]$actualLength,
        [long]$tocLength,
        [int]$footer
    )

    if ($null -eq $header) {
        return $false
    }
    return ($actualLength -eq $header.Layout.Length -and $actualLength -eq $tocLength -and $footer -eq [int]${gramSegmentFooter})
}

function convertFromGramToc {
    # 目録の行を読む。@{ Segments（@{ Name; Length } の配列。古い順）; Old（前の世代の区画の名前）; Next（次の区画の番号） } を返す。
    # 1 行目の印が違う・形の違う行がある・区画の名前が使えないときは $null（読めない目録。呼び出し側は安全側に倒す）。
    # 行: "seg<TAB>名前<TAB>長さ" / "old<TAB>名前" / "next<TAB>番号"
    param (
        [string[]]$lines
    )

    $rows = @($lines | Where-Object { $_ -ne "" -and $null -ne $_ })
    if ($rows.Count -eq 0 -or $rows[0] -ne "${gramTocMark}`t${gramSegmentVersion}") {
        return $null
    }
    $toc = @{ Segments = New-Object System.Collections.Generic.List[hashtable]; Old = New-Object System.Collections.Generic.List[string]; Next = 1 }
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $cells = $rows[$i].Split("`t")
        switch ($cells[0]) {
            "seg" {
                $length = 0L
                if ($cells.Count -ne 3 -or !(testGramSegmentName $cells[1]) -or ![long]::TryParse($cells[2], [ref]$length) -or $length -lt 0) {
                    return $null
                }
                $toc.Segments.Add(@{ Name = $cells[1]; Length = $length })
            }
            "old" {
                if ($cells.Count -ne 2 -or !(testGramSegmentName $cells[1])) {
                    return $null
                }
                $toc.Old.Add($cells[1])
            }
            "next" {
                $next = 0
                if ($cells.Count -ne 2 -or ![int]::TryParse($cells[1], [ref]$next) -or $next -lt 1) {
                    return $null
                }
                $toc.Next = $next
            }
            default {
                return $null
            }
        }
    }
    return $toc
}

function convertToGramToc {
    # 目録を行にする（convertFromGramToc の逆）
    param (
        $toc
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("${gramTocMark}`t${gramSegmentVersion}")
    $lines.Add("next`t$($toc.Next)")
    foreach ($segment in $toc.Segments) {
        $lines.Add("seg`t$($segment.Name)`t$($segment.Length)")
    }
    foreach ($name in $toc.Old) {
        $lines.Add("old`t$name")
    }
    return , $lines.ToArray()
}

function getGramSegmentName {
    # 番号から区画のファイルの名前を作る（g0000000001.seg）
    param (
        [int]$number
    )

    return "g{0:D10}.seg" -f $number
}

function testGramSegmentName {
    # 目録から読んだ区画の名前として使えるか（目録は共有の場所で書き換えられうるため、パスにつなぐ前に形を確かめる）
    param (
        [string]$name
    )

    return ($name -cmatch '^g[0-9]{10}\.seg$')
}

function getGramIndexNames {
    # 索引を作る対象のインデックス名（インデックスのフォルダ直下のフォルダの名前）のうち、パスにつなげる名前だけを返す
    param (
        [string[]]$names
    )

    return , @($names | Where-Object { $_ -and !(testStoredIndexName $_) })
}

function intersectGramRows {
    # 検索語の 2-gram ごとの「それを含むファイルの番号」の並び（rows。載っていない 2-gram は $null か空）の共通部分を HashSet で返す
    # （検索語のすべての 2-gram を含むファイル）。1 つでも載っていない 2-gram があれば空
    param (
        [object[]]$rows
    )

    $set = $null
    foreach ($row in $rows) {
        if ($null -eq $row -or $row.Length -eq 0) {
            return New-Object 'System.Collections.Generic.HashSet[int]'
        }
        if ($null -eq $set) {
            $set = [System.Collections.Generic.HashSet[int]]::new([int[]]$row)
        } else {
            $set.IntersectWith([int[]]$row)
        }
        if ($set.Count -eq 0) {
            break
        }
    }
    if ($null -eq $set) {
        $set = New-Object 'System.Collections.Generic.HashSet[int]'
    }
    return , $set
}

function selectGramFiles {
    # 本文インデックスのファイル（items。@{ RelPath; Ticks; Size } を含む）から、照合するものを選んで返す（順番は変えない）。
    # 外してよいのは、新しい区画から順に探して最初に載っていた区画で、更新日時と大きさが同じで、検索語の 2-gram を含まないファイルだけ。
    # どの区画にも載っていない・更新日時か大きさが違うファイルは、必ず照合する。
    #   segments: 新しい順の区画（@{ Names（相対パス。大文字と小文字を区別せず昇順）; Ticks; Sizes; Candidates（HashSet[int]。Names の番号） }）
    # @{ Files; Excluded（外した数） } を返す
    param (
        [object[]]$items,
        [object[]]$segments
    )

    $comparer = [System.StringComparer]::OrdinalIgnoreCase
    $kept = New-Object System.Collections.Generic.List[object]
    $excluded = 0
    foreach ($item in $items) {
        $keep = $true
        foreach ($segment in $segments) {
            $position = [Array]::BinarySearch($segment.Names, [string]$item.RelPath, $comparer)
            if ($position -lt 0) {
                continue
            }
            if ($segment.Ticks[$position] -eq [long]$item.Ticks -and $segment.Sizes[$position] -eq [long]$item.Size) {
                $keep = $segment.Candidates.Contains($position)
            }
            break
        }
        if ($keep) {
            $kept.Add($item)
        } else {
            $excluded++
        }
    }
    return @{ Files = $kept.ToArray(); Excluded = $excluded }
}

function getGramFilesToAdd {
    # 本文インデックスのファイル（items）のうち、区画にまだ載っていない（どの区画にも無い、または新しい区画の更新日時か大きさが違う）ものを返す。
    # segments は selectGramFiles と同じ形（Candidates は要らない）
    param (
        [object[]]$items,
        [object[]]$segments
    )

    $comparer = [System.StringComparer]::OrdinalIgnoreCase
    $todo = New-Object System.Collections.Generic.List[object]
    foreach ($item in $items) {
        $listed = $false
        foreach ($segment in $segments) {
            $position = [Array]::BinarySearch($segment.Names, [string]$item.RelPath, $comparer)
            if ($position -lt 0) {
                continue
            }
            $listed = ($segment.Ticks[$position] -eq [long]$item.Ticks -and $segment.Sizes[$position] -eq [long]$item.Size)
            break
        }
        if (!$listed) {
            $todo.Add($item)
        }
    }
    return , $todo.ToArray()
}

function splitGramBatches {
    # ファイルごとの 2-gram の数（counts）を、組の数の合計が limit を超えないように前から区切る（1 つのファイルは分けない）。
    # @{ Start; Count } の配列を返す
    param (
        [int[]]$counts,
        [long]$limit = ${gramBatchPostings}
    )

    $batches = New-Object System.Collections.Generic.List[hashtable]
    $start = 0
    $sum = 0L
    for ($i = 0; $i -lt $counts.Length; $i++) {
        if ($i -gt $start -and $sum + $counts[$i] -gt $limit) {
            $batches.Add(@{ Start = $start; Count = $i - $start })
            $start = $i
            $sum = 0L
        }
        $sum += $counts[$i]
    }
    if ($counts.Length -gt $start) {
        $batches.Add(@{ Start = $start; Count = $counts.Length - $start })
    }
    return , $batches.ToArray()
}

function testGramMergeNeeded {
    # 区画をまとめ直すか: まとめたあとの区画の数（組の数 / limit の切り上げ）より、区画が上限の分以上に多いとき。
    # 組が多くて、まとめたあとの数が上限以上になる場合も、まとめた直後に毎回まとめ直さないよう、まとめたあとの数を基準にする
    param (
        [int]$segmentCount,
        [long]$totalPostings,
        [long]$limit = ${gramBatchPostings}
    )

    if ($segmentCount -le ${gramSegmentLimit}) {
        return $false
    }
    $after = [Math]::Max(1, [int][Math]::Ceiling($totalPostings / $limit))
    return ($segmentCount -ge $after + ${gramSegmentLimit})
}
