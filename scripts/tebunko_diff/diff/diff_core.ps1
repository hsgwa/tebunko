# 差分の中心（判断層。入力も出力も素の値で、画面・ファイルに触らない）。
#   ・行の並び 2 つの対応（どの行とどの行が同じか）を求める（getLineMatches）
#   ・対応から、左右に並べる行（同じ・変更・追加・削除）を組み立てる（buildAlignedRows）
#   ・変更の行の、違う文字の区切りを求める（getCharSegments）
#   ・同じ行が続くところをたたむ（foldSameRows）
#
# 行の対応の求め方（PowerShell は 1 回ごとの処理が遅いため、比べる量を先に減らす）:
#   1. 前後の一致する行を取り除く
#   2. 両方に 1 回ずつしか出てこない行を目印にし、目印の最長の並び（最長増加部分列）で区間に分ける（patience diff）
#   3. 目印が無い区間だけを Myers の O(ND) 法で比べる。違いが ${diffMaxMyersD} を超える区間は、まとめて置き換えにする

${diffMaxLines}          = 100000  # 1 つの場所で比べる行数の上限（どちらか一方）。超えたら行ごとには比べない
${diffMaxMyersD}         = 300     # Myers で比べる区間の違いの数の上限。超えた区間はまとめて置き換え（削除と追加）にする
${diffMaxCharLength}     = 2000    # 文字の単位の差分を取る 1 行の長さの上限。超えた行は行全体を変更として示す
${diffMaxCharRows}       = 3000    # 文字の単位の差分を取る変更の行の数の上限（1 つの場所で）
${diffPairMinSimilarity} = 0.3     # 削除と追加を「変更」として組む、似ている度合いの下限
${diffFoldContext}       = 2       # たたむときに、変更の前後に残す同じ行の数

# 左右に並べる 1 行。画面（WPF）はこのプロパティをそのまま表示する（メソッドは持たない）
class DiffRow {
    [string]$Type = "Line"      # Line（行）/ Header（PowerPoint のスライドの見出し）/ Section（ノート・図形などの小見出し）/ Fold（たたんだ同じ行）
    [string]$Kind = "same"      # same / change / insert（比較先だけ）/ delete（比較元だけ）
    [bool]$LeftEmpty            # 比較元の側が空き（insert の行）
    [bool]$RightEmpty           # 比較先の側が空き（delete の行）
    [string]$LeftNo = ""        # 行番号の欄（Excel は行番号、Word は "p.3 ¶2"、PowerPoint は段落番号）
    [string]$RightNo = ""
    [string]$LeftText = ""      # 全文（Excel はセルをタブでつないだもの）
    [string]$RightText = ""
    [object[]]$LeftSegs         # 文字の区切り（DiffSeg）。無ければ $null（全文をそのまま出す）
    [object[]]$RightSegs
    [object[]]$LeftCells        # Excel のセル（DiffCell）。Excel 以外は $null
    [object[]]$RightCells
    [int]$LeftLine = -1         # その場所の TSV での行の位置（0 から）。空きなら -1
    [int]$RightLine = -1
    [string]$LeftPlace = ""     # 開くときの場所（Excel のシート名）
    [string]$RightPlace = ""
    [string]$LeftCell = ""      # 開くときに選ぶセル（Excel。違うセルがあれば最初のもの）
    [string]$RightCell = ""
    [string]$Detail = ""        # 変更の中身の説明（Excel は "B8：10 → 12" など）
    [int]$FoldCount             # Fold の行: たたんだ行の数
    [object[]]$FoldRows         # Fold の行: たたんだ行（開くと、この行の代わりに並べる）
}

# 文字の区切り（同じ部分・違う部分）
class DiffSeg {
    [string]$Text
    [bool]$Changed
}

# Excel のセル 1 つ
class DiffCell {
    [string]$Text
    [bool]$Changed
    [bool]$Blank                # 相手側に行が無い（insert・delete の行のセル）
    [double]$Width
}

# ----------------------------------------------------------------------------
# 行の対応
# ----------------------------------------------------------------------------

function normalizeDiffLine {
    # 比べるための行の形にする（空白の違いを無視・大文字と小文字を区別しない）。表示する文字は変えない
    param (
        [string]$line,
        [bool]$ignoreWhitespace,
        [bool]$caseSensitive
    )

    if ($ignoreWhitespace) {
        # 連続する空白（半角・全角）を 1 つにし、セルの区切り（タブ）の前後と行頭・行末の空白を取る
        $line = [regex]::Replace($line, '[ 　]+', ' ')
        $line = [regex]::Replace($line, ' ?\t ?', "`t").Trim(' ')
    }
    if (!$caseSensitive) {
        $line = $line.ToLowerInvariant()
    }
    return $line
}

function getLineKeys {
    # 行を整数の番号に置き換える（同じ行は同じ番号）。左右で同じ辞書（dict）を使う
    param (
        [string[]]$lines,
        [System.Collections.Generic.Dictionary[string, int]]$dict,
        [bool]$ignoreWhitespace = $false,
        [bool]$caseSensitive = $true
    )

    $keys = New-Object int[] $lines.Count
    $normalize = $ignoreWhitespace -or !$caseSensitive
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $text = $lines[$i]
        if ($normalize) {
            $text = normalizeDiffLine $text $ignoreWhitespace $caseSensitive
        }
        $id = 0
        if (!$dict.TryGetValue($text, [ref]$id)) {
            $id = $dict.Count
            $dict.Add($text, $id)
        }
        $keys[$i] = $id
    }
    return , $keys
}

function getLineMatches {
    # 番号の並び a・b の対応を求め、a の各位置に対応する b の位置（無ければ -1）を返す（int[]）
    param (
        [int[]]$a,
        [int[]]$b,
        [int]$maxD = ${diffMaxMyersD}
    )

    $match = New-Object int[] $a.Count
    for ($i = 0; $i -lt $a.Count; $i++) {
        $match[$i] = -1
    }
    matchRange $a $b 0 $a.Count 0 $b.Count $match $maxD
    return , $match
}

function matchRange {
    # a[aLo..aHi) と b[bLo..bHi) の対応を match に書く（aHi・bHi は含まない）
    param (
        [int[]]$a,
        [int[]]$b,
        [int]$aLo,
        [int]$aHi,
        [int]$bLo,
        [int]$bHi,
        [int[]]$match,
        [int]$maxD
    )

    # 1. 前後の一致
    while ($aLo -lt $aHi -and $bLo -lt $bHi -and $a[$aLo] -eq $b[$bLo]) {
        $match[$aLo] = $bLo
        $aLo++
        $bLo++
    }
    while ($aLo -lt $aHi -and $bLo -lt $bHi -and $a[$aHi - 1] -eq $b[$bHi - 1]) {
        $match[$aHi - 1] = $bHi - 1
        $aHi--
        $bHi--
    }
    if ($aLo -ge $aHi -or $bLo -ge $bHi) {
        return
    }

    # 2. 両方に 1 回ずつしか出てこない行を目印にする
    $anchors = getUniqueAnchors $a $b $aLo $aHi $bLo $bHi
    if ($anchors.Count -gt 0) {
        $prevA = $aLo
        $prevB = $bLo
        foreach ($pair in $anchors) {
            # 片方が空の区間には対応が無い（関数の呼び出しは遅いため、呼ばずに飛ばす）
            if ($pair[0] -gt $prevA -and $pair[1] -gt $prevB) {
                matchRange $a $b $prevA $pair[0] $prevB $pair[1] $match $maxD
            }
            $match[$pair[0]] = $pair[1]
            $prevA = $pair[0] + 1
            $prevB = $pair[1] + 1
        }
        matchRange $a $b $prevA $aHi $prevB $bHi $match $maxD
        return
    }

    # 3. 目印が無い区間は Myers で比べる（違いが多すぎれば、対応なし = まとめて置き換え）
    matchMyers $a $b $aLo $aHi $bLo $bHi $match $maxD
}

function getUniqueAnchors {
    # 区間の中で、a・b のどちらにも 1 回ずつしか出てこない番号の組（a の位置, b の位置）を、
    # 両方の順番がそろう最長の並び（b の位置の最長増加部分列）にして返す
    param (
        [int[]]$a,
        [int[]]$b,
        [int]$aLo,
        [int]$aHi,
        [int]$bLo,
        [int]$bHi
    )

    # 番号ごとに、a での出現数（2 以上は -1）と位置
    $posA = New-Object 'System.Collections.Generic.Dictionary[int,int]'
    for ($i = $aLo; $i -lt $aHi; $i++) {
        $key = $a[$i]
        if ($posA.ContainsKey($key)) {
            $posA[$key] = -1
        } else {
            $posA.Add($key, $i)
        }
    }
    $posB = New-Object 'System.Collections.Generic.Dictionary[int,int]'
    for ($j = $bLo; $j -lt $bHi; $j++) {
        $key = $b[$j]
        if (!$posA.ContainsKey($key) -or $posA[$key] -lt 0) {
            continue
        }
        if ($posB.ContainsKey($key)) {
            $posB[$key] = -1
        } else {
            $posB.Add($key, $j)
        }
    }

    # a の順に並べた候補（b の位置）
    $candA = New-Object System.Collections.Generic.List[int]
    $candB = New-Object System.Collections.Generic.List[int]
    for ($i = $aLo; $i -lt $aHi; $i++) {
        $key = $a[$i]
        $j = 0
        if ($posA[$key] -eq $i -and $posB.TryGetValue($key, [ref]$j) -and $j -ge 0) {
            $candA.Add($i)
            $candB.Add($j)
        }
    }
    $result = New-Object System.Collections.Generic.List[object]
    if ($candA.Count -eq 0) {
        return , $result
    }

    # 最長増加部分列（patience sorting。tails[k] = 長さ k+1 の並びの最後の候補の添字）
    $count = $candA.Count
    $tails = New-Object int[] $count
    $prev = New-Object int[] $count
    $length = 0
    for ($c = 0; $c -lt $count; $c++) {
        $value = $candB[$c]
        $lo = 0
        $hi = $length
        while ($lo -lt $hi) {
            $mid = [int](($lo + $hi) -shr 1)
            if ($candB[$tails[$mid]] -lt $value) { $lo = $mid + 1 } else { $hi = $mid }
        }
        $prev[$c] = if ($lo -gt 0) { $tails[$lo - 1] } else { -1 }
        $tails[$lo] = $c
        if ($lo -eq $length) {
            $length++
        }
    }
    $picked = New-Object int[] $length
    $c = $tails[$length - 1]
    for ($k = $length - 1; $k -ge 0; $k--) {
        $picked[$k] = $c
        $c = $prev[$c]
    }
    foreach ($c in $picked) {
        $result.Add([int[]]@($candA[$c], $candB[$c]))
    }
    return , $result
}

function matchMyers {
    # Myers の O(ND) 法（貪欲法。途中の V を残して、後ろからたどる）。違いの数が maxD を超えたら何もしない
    param (
        [int[]]$a,
        [int[]]$b,
        [int]$aLo,
        [int]$aHi,
        [int]$bLo,
        [int]$bHi,
        [int[]]$match,
        [int]$maxD
    )

    $n = $aHi - $aLo
    $m = $bHi - $bLo
    $max = [Math]::Min($n + $m, $maxD)
    $offset = $max + 1
    $v = New-Object int[] (2 * $max + 3)
    $trace = New-Object System.Collections.Generic.List[int[]]
    $found = $false
    for ($d = 0; $d -le $max -and !$found; $d++) {
        $trace.Add([int[]]$v.Clone())
        for ($k = -$d; $k -le $d; $k += 2) {
            if ($k -eq -$d -or ($k -ne $d -and $v[$offset + $k - 1] -lt $v[$offset + $k + 1])) {
                $x = $v[$offset + $k + 1]
            } else {
                $x = $v[$offset + $k - 1] + 1
            }
            $y = $x - $k
            while ($x -lt $n -and $y -lt $m -and $a[$aLo + $x] -eq $b[$bLo + $y]) {
                $x++
                $y++
            }
            $v[$offset + $k] = $x
            if ($x -ge $n -and $y -ge $m) {
                $found = $true
                break
            }
        }
    }
    if (!$found) {
        return
    }

    # 後ろからたどって、斜めに進んだところ（一致）を書く
    $x = $n
    $y = $m
    for ($d = $trace.Count - 1; $d -ge 0; $d--) {
        $vd = $trace[$d]
        $k = $x - $y
        if ($k -eq -$d -or ($k -ne $d -and $vd[$offset + $k - 1] -lt $vd[$offset + $k + 1])) {
            $prevK = $k + 1
        } else {
            $prevK = $k - 1
        }
        $prevX = $vd[$offset + $prevK]
        $prevY = $prevX - $prevK
        while ($x -gt $prevX -and $y -gt $prevY) {
            $match[$aLo + $x - 1] = $bLo + $y - 1
            $x--
            $y--
        }
        if ($d -gt 0) {
            $x = $prevX
            $y = $prevY
        }
    }
}

# ----------------------------------------------------------------------------
# 似ている度合い
# ----------------------------------------------------------------------------

function getTextSimilarity {
    # 2 つの文字列の似ている度合い（0〜1）。2 文字ずつの組の重なり（Dice 係数）で測る。1 文字以下の文字列は一致かどうかだけ
    param (
        [string]$x,
        [string]$y
    )

    if ($x -eq $y) {
        return 1.0
    }
    if ($x.Length -lt 2 -or $y.Length -lt 2) {
        return 0.0
    }
    $set = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    for ($i = 0; $i -lt $x.Length - 1; $i++) {
        $pair = $x.Substring($i, 2)
        $count = 0
        [void]$set.TryGetValue($pair, [ref]$count)
        $set[$pair] = $count + 1
    }
    $common = 0
    for ($i = 0; $i -lt $y.Length - 1; $i++) {
        $pair = $y.Substring($i, 2)
        $count = 0
        if ($set.TryGetValue($pair, [ref]$count) -and $count -gt 0) {
            $common++
            $set[$pair] = $count - 1
        }
    }
    return (2.0 * $common) / (($x.Length - 1) + ($y.Length - 1))
}

function getCellSimilarity {
    # Excel の 2 行の似ている度合い（0〜1）。値のあるセルのうち、同じ列で同じ値のセルの割合
    param (
        [string]$x,
        [string]$y
    )

    if ($x -eq $y) {
        return 1.0
    }
    $left = $x.Split("`t")
    $right = $y.Split("`t")
    $width = [Math]::Max($left.Count, $right.Count)
    $filled = 0
    $same = 0
    for ($i = 0; $i -lt $width; $i++) {
        $l = if ($i -lt $left.Count) { $left[$i] } else { "" }
        $r = if ($i -lt $right.Count) { $right[$i] } else { "" }
        if ($l -eq "" -and $r -eq "") {
            continue
        }
        $filled++
        if ($l -eq $r) {
            $same++
        }
    }
    if ($filled -eq 0) {
        return 1.0
    }
    return $same / $filled
}

# ----------------------------------------------------------------------------
# 左右に並べる行
# ----------------------------------------------------------------------------

function getAlignedPairs {
    # 対応（match）から、左右に並べる順の組を返す。組は int[]（左の位置, 右の位置, 種類）。
    # 種類: 0 = 同じ、1 = 削除（左だけ）、2 = 追加（右だけ）、3 = 変更（削除と追加を組んだもの）。
    # 一致しなかった区間の削除と追加は、前から順に似ている度合いで組む（similarity: 2 つの文字列を受け取る scriptblock）
    param (
        [int[]]$match,
        [int]$rightCount,
        [string[]]$leftLines,
        [string[]]$rightLines,
        [scriptblock]$similarity,
        [double]$minSimilarity = ${diffPairMinSimilarity}
    )

    $pairs = New-Object System.Collections.Generic.List[int[]]
    $n = $match.Count
    $i = 0
    $j = 0
    while ($i -lt $n -or $j -lt $rightCount) {
        if ($i -lt $n -and $match[$i] -ge 0 -and $match[$i] -eq $j) {
            $same = [int[]]::new(3)
            $same[0] = $i
            $same[1] = $j
            $pairs.Add($same)
            $i++
            $j++
            continue
        }
        # 一致しない区間: 左は次に一致する行まで、右はその相手の位置まで
        $dels = [System.Collections.Generic.List[int]]::new()
        while ($i -lt $n -and $match[$i] -lt 0) {
            $dels.Add($i)
            $i++
        }
        $nextJ = if ($i -lt $n) { $match[$i] } else { $rightCount }
        $inss = [System.Collections.Generic.List[int]]::new()
        while ($j -lt $nextJ) {
            $inss.Add($j)
            $j++
        }
        addGapPairs $pairs $dels $inss $leftLines $rightLines $similarity $minSimilarity
    }
    return , $pairs
}

function addGapPairs {
    # 一致しなかった区間の削除（dels）と追加（inss）を、似ている度合いで組んで pairs に足す
    param (
        $pairs,
        $dels,
        $inss,
        [string[]]$leftLines,
        [string[]]$rightLines,
        [scriptblock]$similarity,
        [double]$minSimilarity
    )

    $p = 0
    $q = 0
    while ($p -lt $dels.Count -and $q -lt $inss.Count) {
        $left = $leftLines[$dels[$p]]
        $right = $rightLines[$inss[$q]]
        if ((& $similarity $left $right) -ge $minSimilarity) {
            $pairs.Add([int[]]@($dels[$p], $inss[$q], 3))
            $p++
            $q++
        } elseif ($q + 1 -lt $inss.Count -and (& $similarity $left $rightLines[$inss[$q + 1]]) -ge $minSimilarity) {
            # 右に 1 行多く挟まっている
            $pairs.Add([int[]]@(-1, $inss[$q], 2))
            $q++
        } elseif ($p + 1 -lt $dels.Count -and (& $similarity $leftLines[$dels[$p + 1]] $right) -ge $minSimilarity) {
            # 左に 1 行多く挟まっている
            $pairs.Add([int[]]@($dels[$p], -1, 1))
            $p++
        } else {
            $pairs.Add([int[]]@($dels[$p], -1, 1))
            $pairs.Add([int[]]@(-1, $inss[$q], 2))
            $p++
            $q++
        }
    }
    for (; $p -lt $dels.Count; $p++) {
        $pairs.Add([int[]]@($dels[$p], -1, 1))
    }
    for (; $q -lt $inss.Count; $q++) {
        $pairs.Add([int[]]@(-1, $inss[$q], 2))
    }
}

function getPairKind {
    # getAlignedPairs の種類の番号を、DiffRow.Kind の文字にする
    param (
        [int]$code
    )

    switch ($code) {
        0 { return "same" }
        1 { return "delete" }
        2 { return "insert" }
        default { return "change" }
    }
}

function getCharSegments {
    # 変更の 2 行（x・y）の、同じ部分と違う部分の区切りを返す: @{ Left = DiffSeg[]; Right = DiffSeg[] }。
    # 長すぎる行は、行全体を違う部分とする
    param (
        [string]$x,
        [string]$y
    )

    if ($x.Length -gt ${diffMaxCharLength} -or $y.Length -gt ${diffMaxCharLength}) {
        return @{ Left = @(newDiffSeg $x $true); Right = @(newDiffSeg $y $true) }
    }
    $a = New-Object int[] $x.Length
    for ($i = 0; $i -lt $x.Length; $i++) { $a[$i] = [int]$x[$i] }
    $b = New-Object int[] $y.Length
    for ($i = 0; $i -lt $y.Length; $i++) { $b[$i] = [int]$y[$i] }
    $match = getLineMatches $a $b

    # 左: 一致した文字は同じ部分。右: 左から対応された位置が同じ部分
    $rightSame = New-Object bool[] $y.Length
    foreach ($j in $match) {
        if ($j -ge 0) { $rightSame[$j] = $true }
    }
    $leftFlags = New-Object bool[] $x.Length
    for ($i = 0; $i -lt $x.Length; $i++) { $leftFlags[$i] = ($match[$i] -lt 0) }
    $rightFlags = New-Object bool[] $y.Length
    for ($j = 0; $j -lt $y.Length; $j++) { $rightFlags[$j] = !$rightSame[$j] }
    return @{ Left = (toDiffSegs $x $leftFlags); Right = (toDiffSegs $y $rightFlags) }
}

function toDiffSegs {
    # 文字ごとの「違う」の印（flags）から、続く同じ印をまとめた区切りを作る
    param (
        [string]$text,
        [bool[]]$flags
    )

    $segs = New-Object System.Collections.Generic.List[object]
    $start = 0
    for ($i = 1; $i -le $text.Length; $i++) {
        if ($i -eq $text.Length -or $flags[$i] -ne $flags[$start]) {
            $segs.Add((newDiffSeg $text.Substring($start, $i - $start) $flags[$start]))
            $start = $i
        }
    }
    return , $segs.ToArray()
}

function newDiffSeg {
    param (
        [string]$text,
        [bool]$changed
    )

    $seg = [DiffSeg]::new()
    $seg.Text = $text
    $seg.Changed = $changed
    return $seg
}

function foldSameRows {
    # 同じ行（Type Line・Kind same）が続くところを、変更の前後 context 行を残して Fold の行にまとめる。
    # label は、たたんだ行の説明の書き方（{0} = 行数、{1} = 最初の行番号、{2} = 最後の行番号）
    param (
        [object[]]$rows,
        [int]$context = ${diffFoldContext},
        [string]$label = "同じ行 {0} 行（行 {1}〜{2}）"
    )

    # 行ごとに関数を呼ぶと遅いため、たためる行かどうかを先に配列にしておく
    $foldable = New-Object bool[] $rows.Count
    for ($k = 0; $k -lt $rows.Count; $k++) {
        $foldable[$k] = ($rows[$k].Type -eq "Line" -and $rows[$k].Kind -eq "same")
    }
    $result = New-Object System.Collections.Generic.List[object]
    $i = 0
    while ($i -lt $rows.Count) {
        if (!$foldable[$i]) {
            $result.Add($rows[$i])
            $i++
            continue
        }
        $start = $i
        while ($i -lt $rows.Count -and $foldable[$i]) {
            $i++
        }
        $end = $i  # 含まない
        $keepHead = if ($start -eq 0) { 0 } else { $context }
        $keepTail = if ($end -eq $rows.Count) { 0 } else { $context }
        $hidden = ($end - $start) - $keepHead - $keepTail
        if ($hidden -lt 3) {
            # たたむほど長くない
            for ($k = $start; $k -lt $end; $k++) { $result.Add($rows[$k]) }
            continue
        }
        for ($k = $start; $k -lt $start + $keepHead; $k++) { $result.Add($rows[$k]) }
        $foldStart = $start + $keepHead
        $foldEnd = $end - $keepTail
        $fold = [DiffRow]::new()
        $fold.Type = "Fold"
        $fold.FoldCount = $foldEnd - $foldStart
        $fold.FoldRows = $rows[$foldStart..($foldEnd - 1)]
        $fold.LeftText = $label -f $fold.FoldCount, $rows[$foldStart].LeftNo, $rows[$foldEnd - 1].LeftNo
        $fold.RightText = $label -f $fold.FoldCount, $rows[$foldStart].RightNo, $rows[$foldEnd - 1].RightNo
        $result.Add($fold)
        for ($k = $foldEnd; $k -lt $end; $k++) { $result.Add($rows[$k]) }
    }
    return , $result.ToArray()
}

function getRowCounts {
    # 行の種類ごとの数: @{ Insert; Delete; Change }（Line の行だけを数える。Fold の中は同じ行なので数えない）
    param (
        [object[]]$rows
    )

    $counts = @{ Insert = 0; Delete = 0; Change = 0 }
    foreach ($row in $rows) {
        if ($row.Type -ne "Line") {
            continue
        }
        switch ($row.Kind) {
            "insert" { $counts.Insert++ }
            "delete" { $counts.Delete++ }
            "change" { $counts.Change++ }
        }
    }
    return $counts
}
