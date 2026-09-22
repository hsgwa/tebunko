# 種類ごとの比べ方（判断層）。場所（シート・ページ・スライド）ごとの行（抽出した TSV の中身）を受け取り、
# 画面に出す場所（PlaceDiff）と左右に並べる行（DiffRow）を組み立てる。
#   Excel      : シート名で場所を対応づけ（名前を変えたシートは中身で組む）、列を対応づけてから（列の挿入・削除）、
#                共通の列の値で行の挿入・削除を見つける。対応した行はセルごとに比べる
#   Word       : ページをつないだ段落の並びで比べる（ページの区切りは目安で、動きやすいため対応づけに使わない）
#   PowerPoint : スライドの本文でスライドを対応づけ、その中の段落・ノート・図形・コメントを比べる
#
# 入力の units は、場所の名前 → 行（string[]）の順序付きの辞書（抽出した順）。options は getDiffOptions の形。

${diffMaxGridColumns} = 100  # Excel の表として出す列の上限（これより右の列は表に出さない。比べるときは全部の列を使う）
${diffObjectKinds}    = @{ Shape = "[図形]"; Comment = "[コメント]" }
${diffPairMinSlideSimilarity} = 0.4  # PowerPoint の本文が一致しないスライドを対応づける、似ている度合いの下限
${diffPairMinCellSimilarity} = 0.4  # Excel の削除と追加の行を「変更」として組む下限（同じセルの割合と、行の文字の似ている度合いの大きい方）
${diffPairMinColumnSimilarity} = 0.5  # Excel の列を対応づける、値の重なりの割合の下限（下回れば列の追加・削除）
${diffColumnSampleRows} = 2000     # Excel の列の対応づけに使う、値のある行の数（先頭から）
${diffPairMinSheetSimilarity} = 0.5  # 名前の違うシートを同じシート（名前を変えた）とする、同じ中身の行の割合の下限
# 見えない文字（引用符・ゼロ幅スペース・BOM・セル内の改行）。これと空白だけのセル・段落は空として扱う（isBlankCell）
${diffInvisibleChars} = [char[]]@('"', [char]0x200B, [char]0xFEFF, [char]0x2028)

# 画面に出す場所 1 つ（Excel のシート、Word の本文・脚注、PowerPoint のスライドとノート など）
class PlaceDiff {
    [string]$Name               # 見出しに出す名前
    [string]$LeftName = ""      # 比較元での場所の名前（無ければ ""）
    [string]$RightName = ""
    [string]$Status = "same"    # same / change / insert（比較先だけ）/ delete（比較元だけ）
    [object[]]$Rows             # 左右に並べる行（DiffRow）
    [object[]]$FoldedRows       # 同じ行をたたんだもの
    [int]$Inserts
    [int]$Deletes
    [int]$Changes
    [int]$Moves                 # 動かした行の数
    [bool]$IsGrid               # Excel の表として出す
    [double[]]$ColumnWidths     # 表の列の幅（IsGrid のとき。左右そろえた列ごと）
    [string[]]$LeftColumnNames  # 表の列の見出し（A, B, …。相手側にだけある列は ""）
    [string[]]$RightColumnNames
    [string[]]$ColumnKinds      # 列の種類（"" / insert（比較先だけ。動かした先も）/ delete（比較元だけ。動かす前も））
    [int]$ColumnInserts         # 追加・削除した列の数（空の列は数えない）
    [int]$ColumnDeletes
    [int]$ColumnMoves           # 動かした列の数
    [string]$ColumnNote = ""    # 追加・削除した列の説明（"列の追加 C（役職）"）
    [string]$Note = ""          # 注意（大きすぎて行ごとに比べなかった等）
}

# 1 ファイルの比較の結果
class FileDiff {
    [string]$Kind               # Excel / Word / PowerPoint
    [object[]]$Places           # PlaceDiff
    [int]$Inserts
    [int]$Deletes
    [int]$Changes
    [int]$Moves                 # 動かした行・段落の数
    [int]$ColumnInserts         # Excel: 追加・削除した列の数
    [int]$ColumnDeletes
    [int]$ColumnMoves
    [int]$LeftSlides            # PowerPoint: スライドの枚数
    [int]$RightSlides
    [int]$SlideInserts          # PowerPoint: 追加・削除したスライドの枚数
    [int]$SlideDeletes
    [int]$SlideMoves            # PowerPoint: 並べ替えで動かしたスライドの数
}

function getDiffOptions {
    # 比べ方の設定（settings_diff.ps1 の diffOptions と同じキー）を、既定値で補って返す
    param (
        $options = $null
    )

    $result = @{ IncludeShapes = $true; IncludeComments = $true; IgnoreWhitespace = $false; CaseSensitive = $true }
    if ($null -ne $options) {
        foreach ($key in @("IncludeShapes", "IncludeComments", "IgnoreWhitespace", "CaseSensitive")) {
            $name = $key.Substring(0, 1).ToLowerInvariant() + $key.Substring(1)
            if ($options -is [System.Collections.IDictionary]) {
                if ($options.Contains($key)) { $result[$key] = [bool]$options[$key] }
                elseif ($options.Contains($name)) { $result[$key] = [bool]$options[$name] }
            }
        }
    }
    return $result
}

function compareOfficeUnits {
    # 1 ファイルの比較。kind は Excel / Word / PowerPoint
    param (
        [string]$kind,
        $left,
        $right,
        $options = $null
    )

    $options = getDiffOptions $options
    $left = filterObjectPlaces $left $options
    $right = filterObjectPlaces $right $options
    $result = [FileDiff]::new()
    $result.Kind = $kind
    switch ($kind) {
        "Excel"      { $result.Places = compareExcelUnits $left $right $options }
        "Word"       { $result.Places = compareWordUnits $left $right $options }
        "PowerPoint" { compareSlideUnits $left $right $options $result }
        default      { throw "比べられない種類です: $kind" }
    }
    foreach ($place in $result.Places) {
        $result.Inserts += $place.Inserts
        $result.Deletes += $place.Deletes
        $result.Changes += $place.Changes
        $result.Moves += $place.Moves
        $result.ColumnInserts += $place.ColumnInserts
        $result.ColumnDeletes += $place.ColumnDeletes
        $result.ColumnMoves += $place.ColumnMoves
    }
    return $result
}

function filterObjectPlaces {
    # 図形・コメントの場所を、設定に合わせて除く
    param (
        $units,
        $options
    )

    $result = [ordered]@{}
    if ($null -eq $units) {
        return $result
    }
    foreach ($name in @($units.Keys)) {
        if (!$options.IncludeShapes -and $name.EndsWith(${diffObjectKinds}.Shape)) { continue }
        if (!$options.IncludeComments -and $name.EndsWith(${diffObjectKinds}.Comment)) { continue }
        $result[$name] = $units[$name]
    }
    return $result
}

function mergePlaceOrder {
    # 左右の場所の名前を、大文字・小文字を区別せずに対応づけて並べる。
    # 並びは比較先の順を基にし、比較元にしか無いものは、比較元でその前にあった場所の後ろに入れる。
    # 返すのは @{ Left; Right } の配列（無い側は ""）
    param (
        [string[]]$leftNames,
        [string[]]$rightNames
    )

    $comparer = [System.StringComparer]::OrdinalIgnoreCase
    $rightIndex = New-Object 'System.Collections.Generic.Dictionary[string,string]' $comparer
    foreach ($name in $rightNames) { if (!$rightIndex.ContainsKey($name)) { $rightIndex.Add($name, $name) } }
    $leftIndex = New-Object 'System.Collections.Generic.Dictionary[string,string]' $comparer
    foreach ($name in $leftNames) { if (!$leftIndex.ContainsKey($name)) { $leftIndex.Add($name, $name) } }

    # 比較元にしか無い場所を、直前にある（両方にある）場所ごとにまとめる
    $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]' $comparer
    $head = New-Object System.Collections.Generic.List[string]
    $anchor = $null
    foreach ($name in $leftNames) {
        if ($rightIndex.ContainsKey($name)) {
            $anchor = $name
            continue
        }
        if ($null -eq $anchor) {
            $head.Add($name)
        } else {
            if (!$pending.ContainsKey($anchor)) { $pending.Add($anchor, (New-Object System.Collections.Generic.List[string])) }
            $pending[$anchor].Add($name)
        }
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($name in $head) { $result.Add(@{ Left = $name; Right = "" }) }
    foreach ($name in $rightNames) {
        $leftName = ""
        [void]$leftIndex.TryGetValue($name, [ref]$leftName)
        $result.Add(@{ Left = [string]$leftName; Right = $name })
        if ($leftName -and $pending.ContainsKey($leftName)) {
            foreach ($only in $pending[$leftName]) { $result.Add(@{ Left = $only; Right = "" }) }
        }
    }
    return , $result.ToArray()
}

# ----------------------------------------------------------------------------
# 行の比較（種類に共通）
# ----------------------------------------------------------------------------

function getUnitLines {
    # 場所の行を string[] で返す（場所が無い・名前が空なら空の配列）。
    # `$x = if (...) { $array }` と書くと、空の配列が $null に、1 行の配列が文字列になるため、この関数を使う
    param (
        $units,
        [string]$name
    )

    if ($null -eq $units -or !$name) {
        return , ([string[]]@())
    }
    $value = $units[$name]
    if ($null -eq $value) {
        return , ([string[]]@())
    }
    return , ([string[]]@($value))
}

function compareLines {
    # 文字の行の並び 2 つを比べ、getAlignedPairs の組（位置は元の行の位置）を返す（compareKeyLines）
    param (
        [string[]]$left,
        [string[]]$right,
        $options
    )

    return , (compareKeyLines (getCompareKeys $left $options) (getCompareKeys $right $options))
}

function compareKeyLines {
    # 比べる形にした行（getCompareKeys・getRowKeys）の並び 2 つを比べ、getAlignedPairs の組（位置は元の行の位置）を返す。
    # cells は Excel のセルの似ている度合いで組むとき。
    # 空の行（空白だけの行・値の無い Excel の行）は比べず、組にも入れない（空の行を足した・消しただけでは差分にしない）
    param (
        [string[]]$leftKeys,
        [string[]]$rightKeys,
        [bool]$cells = $false
    )

    $leftIndex = [System.Collections.Generic.List[int]]::new()
    $leftFilled = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $leftKeys.Count; $i++) {
        if (![string]::IsNullOrWhiteSpace($leftKeys[$i])) { $leftIndex.Add($i); $leftFilled.Add($leftKeys[$i]) }
    }
    $rightIndex = [System.Collections.Generic.List[int]]::new()
    $rightFilled = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $rightKeys.Count; $i++) {
        if (![string]::IsNullOrWhiteSpace($rightKeys[$i])) { $rightIndex.Add($i); $rightFilled.Add($rightKeys[$i]) }
    }
    $a2 = $leftFilled.ToArray()
    $b2 = $rightFilled.ToArray()

    # 比べる形はすでに整えてあるため、getLineKeys では整えない
    $dict = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $a = getLineKeys $a2 $dict
    $b = getLineKeys $b2 $dict
    $match = getLineMatches $a $b
    if ($cells) {
        $pairs = getAlignedPairs $match $b2.Count $a2 $b2 { param($x, $y) [Math]::Max([double](getCellSimilarity $x $y), [double](getTextSimilarity $x $y)) } ${diffPairMinCellSimilarity}
    } else {
        $pairs = getAlignedPairs $match $b2.Count $a2 $b2 { param($x, $y) getTextSimilarity $x $y }
    }
    # 空の行を除いた位置を、元の行の位置に戻す
    foreach ($pair in $pairs) {
        if ($pair[0] -ge 0) { $pair[0] = $leftIndex[$pair[0]] }
        if ($pair[1] -ge 0) { $pair[1] = $rightIndex[$pair[1]] }
    }
    return , $pairs
}

function getCompareKeys {
    # 比べるための文字の行の形（string[]）。空白の違い・大文字と小文字の設定に合わせて整え、見えない文字だけの段落は空にする
    param (
        [string[]]$lines,
        $options
    )

    # 行ごとに関数を呼ぶと遅いため、整える設定のあるときだけ normalizeDiffLine を呼ぶ
    $normalize = $options.IgnoreWhitespace -or !$options.CaseSensitive
    $keys = [string[]]::new($lines.Count)
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $text = $lines[$i]
        if ($text.IndexOfAny(${diffInvisibleChars}) -ge 0 -and (isBlankCell $text)) {
            $text = ""
        }
        if ($normalize) {
            $text = normalizeDiffLine $text $options.IgnoreWhitespace $options.CaseSensitive
        }
        $keys[$i] = $text
    }
    return , $keys
}

function isBlankCell {
    # 見た目が空のセル・段落か（引用符を外した中身が、空白と見えない文字だけ。セル内の改行 " ↵ " も見えない文字とする）
    param (
        [string]$cell
    )

    if ($cell.Length -ge 2 -and $cell[0] -eq '"' -and $cell[$cell.Length - 1] -eq '"') {
        $cell = $cell.Substring(1, $cell.Length - 2).Replace('""', '"')
    }
    return [string]::IsNullOrWhiteSpace($cell.Replace([string][char]0x200B, "").Replace([string][char]0xFEFF, "").Replace([string][char]0x2028, "").Replace("↵", ""))
}

function markMovedRows {
    # 削除と追加の行で中身（keys。行の位置 → 比べる形）が同じものを組み、説明（Detail）に移動の元と先を書く。
    # 並べ替え・切り取りと貼り付けで動いた行を、別々の削除と追加ではなく移動と分かるようにする
    param (
        [object[]]$rows,
        [string[]]$leftKeys,
        [string[]]$rightKeys,
        [string]$unit = "行 "
    )

    $deleted = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach ($row in $rows) {
        if ($row.Type -ne "Line" -or $row.Kind -ne "delete" -or $row.LeftLine -lt 0) { continue }
        $key = $leftKeys[$row.LeftLine]
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if (!$deleted.ContainsKey($key)) { $deleted.Add($key, (New-Object System.Collections.Generic.Queue[object])) }
        $deleted[$key].Enqueue($row)
    }
    if ($deleted.Count -eq 0) {
        return
    }
    foreach ($row in $rows) {
        if ($row.Type -ne "Line" -or $row.Kind -ne "insert" -or $row.RightLine -lt 0) { continue }
        $key = $rightKeys[$row.RightLine]
        if (!$deleted.ContainsKey($key) -or $deleted[$key].Count -eq 0) { continue }
        $from = $deleted[$key].Dequeue()
        $from.Detail = "$unit$($row.RightNo) へ移動"
        $row.Detail = "$unit$($from.LeftNo) から移動"
        $from.Moved = $true
        $row.Moved = $true
    }
}

function newTextRows {
    # 文字の行（Word の段落・PowerPoint の段落・Excel の図形とコメント）の組から、左右に並べる行を作る。
    # leftNos・rightNos は、行ごとの行番号の欄の文字（行の数と同じ長さ）。
    # 行ごとに関数を呼ぶと遅いため、ループの中では関数を呼ばない
    param (
        $pairs,
        [string[]]$left,
        [string[]]$right,
        [string[]]$leftNos,
        [string[]]$rightNos,
        [string]$leftPlace = "",
        [string]$rightPlace = ""
    )

    $kindNames = @("same", "delete", "insert", "change")
    $cellNewLine = [string][char]0x2028
    $rows = New-Object System.Collections.Generic.List[object]
    $charRows = 0
    foreach ($pair in $pairs) {
        $row = [DiffRow]::new()
        $row.Kind = $kindNames[$pair[2]]
        $row.LeftLine = $pair[0]
        $row.RightLine = $pair[1]
        $row.LeftPlace = $leftPlace
        $row.RightPlace = $rightPlace
        if ($pair[0] -ge 0) {
            $row.LeftText = $left[$pair[0]].Replace($cellNewLine, " ↵ ")
            $row.LeftNo = $leftNos[$pair[0]]
        } else {
            $row.LeftEmpty = $true
        }
        if ($pair[1] -ge 0) {
            $row.RightText = $right[$pair[1]].Replace($cellNewLine, " ↵ ")
            $row.RightNo = $rightNos[$pair[1]]
        } else {
            $row.RightEmpty = $true
        }
        if ($row.Kind -eq "change" -and $charRows -lt ${diffMaxCharRows}) {
            $segs = getCharSegments $row.LeftText $row.RightText
            $row.LeftSegs = $segs.Left
            $row.RightSegs = $segs.Right
            $charRows++
        }
        $rows.Add($row)
    }
    return , $rows.ToArray()
}

function getParagraphNos {
    # 段落番号の欄の文字（"¶1", "¶2", …）を count 個
    param (
        [int]$count
    )

    $nos = New-Object string[] $count
    for ($i = 0; $i -lt $count; $i++) {
        $nos[$i] = "¶$($i + 1)"
    }
    return , $nos
}

function newPlaceDiff {
    # 場所の結果を作る（行の数を数え、たたんだ行も作る）
    param (
        [string]$name,
        [string]$leftName,
        [string]$rightName,
        [object[]]$rows,
        [string]$foldLabel = "同じ行 {0} 行（行 {1}〜{2}）"
    )

    $place = [PlaceDiff]::new()
    $place.Name = $name
    $place.LeftName = $leftName
    $place.RightName = $rightName
    $place.Rows = $rows
    $counts = getRowCounts $rows
    $place.Inserts = $counts.Insert
    $place.Deletes = $counts.Delete
    $place.Changes = $counts.Change
    $place.Moves = $counts.Move
    $place.Status = if (!$leftName) { "insert" } elseif (!$rightName) { "delete" } elseif (($counts.Insert + $counts.Delete + $counts.Change + $counts.Move) -gt 0) { "change" } else { "same" }
    $place.FoldedRows = foldSameRows $rows ${diffFoldContext} $foldLabel
    return $place
}

function newTooLargePlace {
    # 行が多すぎて行ごとに比べない場所。同じかどうかだけを出す
    param (
        [string]$name,
        [string]$leftName,
        [string]$rightName,
        [string[]]$left,
        [string[]]$right
    )

    $place = [PlaceDiff]::new()
    $place.Name = $name
    $place.LeftName = $leftName
    $place.RightName = $rightName
    $same = ($left.Count -eq $right.Count) -and (($left -join "`n") -ceq ($right -join "`n"))
    $place.Status = if ($same) { "same" } else { "change" }
    $place.Changes = if ($same) { 0 } else { 1 }
    $place.Rows = @()
    $place.FoldedRows = @()
    $place.Note = "行が多すぎるため（上限 ${diffMaxLines} 行）、行ごとには比べていません。" + $(if ($same) { "中身は同じです。" } else { "中身が違います。" })
    return $place
}

# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------

function compareExcelUnits {
    param (
        $left,
        $right,
        $options
    )

    $order = pairRenamedSheets (mergePlaceOrder @($left.Keys) @($right.Keys)) $left $right
    $places = New-Object System.Collections.Generic.List[object]
    foreach ($pair in $order) {
        $leftLines = getUnitLines $left $pair.Left
        $rightLines = getUnitLines $right $pair.Right
        $name = if ($pair.Right) { $pair.Right } else { $pair.Left }
        if ($leftLines.Count -gt ${diffMaxLines} -or $rightLines.Count -gt ${diffMaxLines}) {
            $place = newTooLargePlace $name $pair.Left $pair.Right $leftLines $rightLines
        } elseif ($name.EndsWith(${diffObjectKinds}.Shape) -or $name.EndsWith(${diffObjectKinds}.Comment)) {
            $place = compareExcelObjects $name $pair.Left $pair.Right $leftLines $rightLines $options
        } else {
            $place = compareExcelSheet $name $pair.Left $pair.Right $leftLines $rightLines $options
        }
        # 名前を変えたシート（大文字・小文字だけの違いは同じ名前とする）
        if ($pair.Left -and $pair.Right -and $pair.Left -ne $pair.Right) {
            $place.Note = ("シート名を変えました（$($pair.Left) → $($pair.Right)）。" + $place.Note)
            if ($place.Status -eq "same") { $place.Status = "change" }
        }
        $places.Add($place)
    }
    return , $places.ToArray()
}

function pairRenamedSheets {
    # 名前で対応しなかったシート（比較元だけ・比較先だけ）を、中身が似ていれば同じシートとして組む（シート名を変えたとき）。
    # そのシートの図形・コメントの場所も同じように組む。order は mergePlaceOrder の形で、同じ形で返す
    param (
        [object[]]$order,
        $left,
        $right
    )

    $isObject = { param($n) $n.EndsWith(${diffObjectKinds}.Shape) -or $n.EndsWith(${diffObjectKinds}.Comment) }
    $leftOnly = @($order | Where-Object { $_.Left -and !$_.Right -and !(& $isObject $_.Left) } | ForEach-Object { $_.Left })
    $rightOnly = @($order | Where-Object { !$_.Left -and $_.Right -and !(& $isObject $_.Right) } | ForEach-Object { $_.Right })
    if ($leftOnly.Count -eq 0 -or $rightOnly.Count -eq 0) {
        return , $order
    }

    # 比較先のシートごとに、いちばん似ている比較元のシートを選ぶ（組んだものは使わない）
    $comparer = [System.StringComparer]::OrdinalIgnoreCase
    $renamed = New-Object 'System.Collections.Generic.Dictionary[string,string]' $comparer  # 比較先の名前 → 比較元の名前
    $used = New-Object 'System.Collections.Generic.HashSet[string]' $comparer
    foreach ($rightName in $rightOnly) {
        $best = $null
        $bestScore = 0.0
        foreach ($leftName in $leftOnly) {
            if ($used.Contains($leftName)) { continue }
            $score = getSheetSimilarity (getUnitLines $left $leftName) (getUnitLines $right $rightName)
            if ($score -gt $bestScore) { $best = $leftName; $bestScore = $score }
        }
        if ($null -ne $best -and $bestScore -ge ${diffPairMinSheetSimilarity}) {
            $renamed.Add($rightName, $best)
            [void]$used.Add($best)
        }
    }
    if ($renamed.Count -eq 0) {
        return , $order
    }
    # 図形・コメントの場所（"<シート名>[図形]"）も、シートと同じ組にする
    $leftNames = New-Object 'System.Collections.Generic.HashSet[string]' $comparer
    foreach ($entry in $order) { if ($entry.Left) { [void]$leftNames.Add($entry.Left) } }
    foreach ($entry in @($order | Where-Object { !$_.Left -and $_.Right -and (& $isObject $_.Right) })) {
        $at = $entry.Right.LastIndexOf("[")
        $base = $entry.Right.Substring(0, $at)
        if ($renamed.ContainsKey($base)) {
            $leftName = $renamed[$base] + $entry.Right.Substring($at)
            if ($leftNames.Contains($leftName) -and !$used.Contains($leftName)) {
                $renamed.Add($entry.Right, $leftName)
                [void]$used.Add($leftName)
            }
        }
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $order) {
        if ($entry.Left -and !$entry.Right -and $used.Contains($entry.Left)) {
            continue
        }
        if (!$entry.Left -and $renamed.ContainsKey($entry.Right)) {
            $result.Add(@{ Left = $renamed[$entry.Right]; Right = $entry.Right })
            continue
        }
        $result.Add($entry)
    }
    return , $result.ToArray()
}

function getSheetSimilarity {
    # 2 枚のシートの似ている度合い（0〜1）。値のある行のうち、同じ中身の行の割合（多い方の行数で割る。行の順番は見ない）
    param (
        [string[]]$left,
        [string[]]$right
    )

    $counts = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $leftCount = 0
    foreach ($line in $left) {
        $key = $line.TrimEnd("`t")
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $leftCount++
        $n = 0
        [void]$counts.TryGetValue($key, [ref]$n)
        $counts[$key] = $n + 1
    }
    $rightCount = 0
    $common = 0
    foreach ($line in $right) {
        $key = $line.TrimEnd("`t")
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $rightCount++
        $n = 0
        if ($counts.TryGetValue($key, [ref]$n) -and $n -gt 0) {
            $common++
            $counts[$key] = $n - 1
        }
    }
    $max = [Math]::Max($leftCount, $rightCount)
    if ($max -eq 0) {
        return 0.0
    }
    return [double]$common / $max
}

function compareExcelSheet {
    # シート 1 枚。行番号は TSV の行の位置 + 1（TSV の N 行目 = シートの N 行目）。
    # 先に列を対応づけ（alignExcelColumns）、左右の表の列をそろえる。追加・削除した列と空の列は行の比べ方に入れず、
    # 共通の列だけで行を対応づける（列を挿入しても、ほかの行を変更にしない）
    param (
        [string]$name,
        [string]$leftName,
        [string]$rightName,
        [string[]]$left,
        [string[]]$right,
        $options
    )

    $leftCells = splitExcelLines $left
    $rightCells = splitExcelLines $right
    $leftValues = getCellValues $leftCells $options
    $rightValues = getCellValues $rightCells $options
    $slots = alignExcelColumns $leftValues $rightValues
    if ($slots.Count -eq 0) {
        # 両方とも空のシート
        $slots = @(, [int[]]@(0, 0, 0, -1))
    }
    $leftPaired = [System.Collections.Generic.List[int]]::new()
    $rightPaired = [System.Collections.Generic.List[int]]::new()
    foreach ($slot in $slots) {
        if ($slot[0] -ge 0 -and $slot[1] -ge 0) { $leftPaired.Add($slot[0]); $rightPaired.Add($slot[1]) }
    }
    if ($leftPaired.Count -eq 0) {
        # 片側にしか無い（または片側が空の）シートは、行ごと追加・削除になるため、列の追加・削除とはしない
        foreach ($slot in $slots) { $slot[2] = 0 }
    }
    $leftKeys = getRowKeys $leftValues $leftPaired.ToArray()
    $rightKeys = getRowKeys $rightValues $rightPaired.ToArray()
    $pairs = compareKeyLines $leftKeys $rightKeys $true

    # 表の列（左右そろえた列）。表に出すのは ${diffMaxGridColumns} 列まで。列の幅は左右の長い方の値に合わせる
    $columns = [Math]::Min($slots.Count, ${diffMaxGridColumns})
    $leftWidth = 0
    foreach ($cells in $leftCells) { if ($cells.Count -gt $leftWidth) { $leftWidth = $cells.Count } }
    $rightWidth = 0
    foreach ($cells in $rightCells) { if ($cells.Count -gt $rightWidth) { $rightWidth = $cells.Count } }
    $leftColumnWidths = getGridColumnWidths $leftCells @() ([Math]::Max(1, $leftWidth))
    $rightColumnWidths = getGridColumnWidths @() $rightCells ([Math]::Max(1, $rightWidth))
    $leftSlotOf = [int[]]::new([Math]::Max(1, $leftWidth))
    $rightSlotOf = [int[]]::new([Math]::Max(1, $rightWidth))
    for ($c = 0; $c -lt $leftSlotOf.Count; $c++) { $leftSlotOf[$c] = -1 }
    for ($c = 0; $c -lt $rightSlotOf.Count; $c++) { $rightSlotOf[$c] = -1 }
    $widths = [double[]]::new($columns)
    $leftNames = [string[]]::new($columns)
    $rightNames = [string[]]::new($columns)
    $kinds = [string[]]::new($columns)
    $lastLeftGap = -1
    $lastRightGap = -1
    for ($s = 0; $s -lt $columns; $s++) {
        $lc = $slots[$s][0]
        $rc = $slots[$s][1]
        $width = 44.0
        if ($lc -ge 0) {
            if ($lc -lt $leftSlotOf.Count) { $leftSlotOf[$lc] = $s }
            if ($lc -lt $leftColumnWidths.Count) { $width = [Math]::Max($width, $leftColumnWidths[$lc]) }
            $leftNames[$s] = getColumnName ($lc + 1)
        } else {
            $leftNames[$s] = ""
            $lastLeftGap = $s
        }
        if ($rc -ge 0) {
            if ($rc -lt $rightSlotOf.Count) { $rightSlotOf[$rc] = $s }
            if ($rc -lt $rightColumnWidths.Count) { $width = [Math]::Max($width, $rightColumnWidths[$rc]) }
            $rightNames[$s] = getColumnName ($rc + 1)
        } else {
            $rightNames[$s] = ""
            $lastRightGap = $s
        }
        $widths[$s] = $width
        $kinds[$s] = switch ($slots[$s][2]) { 1 { "delete" } 4 { "delete" } 2 { "insert" } 5 { "insert" } default { "" } }
    }

    # 表の列ごとの、比較元・比較先の列の番号（1 行ごとに回るところで組を引かないよう、先に配列にしておく）
    $leftColumnOf = [int[]]::new($columns)
    $rightColumnOf = [int[]]::new($columns)
    $allPaired = $true
    for ($s = 0; $s -lt $columns; $s++) {
        $leftColumnOf[$s] = $slots[$s][0]
        $rightColumnOf[$s] = $slots[$s][1]
        if ($slots[$s][2] -ne 0) { $allPaired = $false }
    }

    $kindNames = @("same", "delete", "insert", "change")
    $noChange = [bool[]]::new($columns)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($pair in $pairs) {
        $row = [DiffRow]::new()
        $row.Kind = $kindNames[$pair[2]]
        $row.LeftLine = $pair[0]
        $row.RightLine = $pair[1]
        $row.LeftPlace = $leftName
        $row.RightPlace = $rightName
        $row.HasCells = $true
        # `$x = if (...) { $array }` と書くと配列がばらされる（1 セルの行が文字列になる）ため、分けて代入する
        $lc = $null
        $lv = $null
        if ($pair[0] -ge 0) { $lc = $leftCells[$pair[0]]; $lv = $leftValues[$pair[0]] }
        $rc = $null
        $rv = $null
        if ($pair[1] -ge 0) { $rc = $rightCells[$pair[1]]; $rv = $rightValues[$pair[1]] }
        # 値の違うセル（共通の列だけを比べる。比べる値は整えたもの）
        # New-Object はコマンドレットの呼び出しで遅いため、1 行ごとに回るところでは ::new を使う
        $changedSlots = [System.Collections.Generic.List[int]]::new()
        if ($row.Kind -eq "change") {
            for ($s = 0; $s -lt $slots.Count; $s++) {
                $x = $slots[$s][0]
                $y = $slots[$s][1]
                if ($x -lt 0 -or $y -lt 0) { continue }
                $l = if ($x -lt $lv.Count) { $lv[$x] } else { "" }
                $r = if ($y -lt $rv.Count) { $rv[$y] } else { "" }
                if ($l -cne $r) { $changedSlots.Add($s) }
            }
            # 値の違うセルが無ければ（空白の違いを無視したときなど）同じ行にする
            if ($changedSlots.Count -eq 0) { $row.Kind = "same" }
        }
        $changed = $noChange
        $lastChanged = -1
        if ($changedSlots.Count -gt 0) {
            $changed = [bool[]]::new($columns)
            foreach ($s in $changedSlots) { if ($s -lt $columns) { $changed[$s] = $true; $lastChanged = $s } }
        }
        $whole = ($row.Kind -eq "delete" -or $row.Kind -eq "insert")
        # 1 行ごとに回るところのため、関数や名前を組み立てたプロパティの読み書きを使わず、左右を別々に書く。
        # セルの部品は、値のある列・違うセルの列・相手側にだけある列（空き）の分だけ作る
        if ($null -ne $lc) {
            $row.LeftNo = [string]($pair[0] + 1)
            $row.LeftText = [string]::Join("`t", $lc)
            $last = [Math]::Max($lastChanged, $lastLeftGap)
            for ($c = 0; $c -lt $lc.Count -and $c -lt $leftSlotOf.Count; $c++) {
                if ($lc[$c].Length -gt 0 -and $leftSlotOf[$c] -gt $last -and $leftSlotOf[$c] -lt $columns) { $last = $leftSlotOf[$c] }
            }
            $items = [object[]]::new($last + 1)
            for ($s = 0; $s -le $last; $s++) {
                $cell = [DiffCell]::new()
                $cell.Width = $widths[$s]
                $x = $leftColumnOf[$s]
                if ($x -lt 0) {
                    $cell.Text = ""
                    $cell.Kind = "empty"
                } else {
                    if ($x -lt $lc.Count) { $cell.Text = $lc[$x] } else { $cell.Text = "" }
                    if ($whole -or $changed[$s]) {
                        $cell.Changed = $true
                        $cell.Kind = $row.Kind
                    } elseif (($slots[$s][2] -eq 1 -or $slots[$s][2] -eq 4) -and $x -lt $lv.Count -and $lv[$x].Length -gt 0) {
                        # 削除した（動かす前の）列の値
                        $cell.Changed = $true
                        $cell.Kind = "delete"
                        $row.HasColumnChange = $true
                    }
                }
                $items[$s] = $cell
            }
            $row.LeftCells = $items
        } else {
            $row.LeftEmpty = $true
        }
        if ($null -ne $rc) {
            $row.RightNo = [string]($pair[1] + 1)
            $row.RightText = [string]::Join("`t", $rc)
            $last = [Math]::Max($lastChanged, $lastRightGap)
            for ($c = 0; $c -lt $rc.Count -and $c -lt $rightSlotOf.Count; $c++) {
                if ($rc[$c].Length -gt 0 -and $rightSlotOf[$c] -gt $last -and $rightSlotOf[$c] -lt $columns) { $last = $rightSlotOf[$c] }
            }
            # 同じ行で、表の列が左右で 1 対 1 に対応し、値もすべて同じなら、比較元のセルの部品をそのまま使う
            # （部品を作るのは遅いため。大きいシートの同じ行の分がほぼ半分になる）
            $shared = $false
            if ($row.Kind -eq "same" -and $allPaired -and $null -ne $row.LeftCells -and $row.LeftCells.Count -eq $last + 1) {
                $shared = $true
                for ($s = 0; $s -le $last; $s++) {
                    $y = $rightColumnOf[$s]
                    $text = if ($y -lt $rc.Count) { $rc[$y] } else { "" }
                    if ($text -cne $row.LeftCells[$s].Text) { $shared = $false; break }
                }
            }
            if ($shared) {
                $row.RightCells = $row.LeftCells
            }
            $items = [object[]]::new($last + 1)
            for ($s = 0; $s -le $last -and !$shared; $s++) {
                $cell = [DiffCell]::new()
                $cell.Width = $widths[$s]
                $y = $rightColumnOf[$s]
                if ($y -lt 0) {
                    $cell.Text = ""
                    $cell.Kind = "empty"
                } else {
                    if ($y -lt $rc.Count) { $cell.Text = $rc[$y] } else { $cell.Text = "" }
                    if ($whole -or $changed[$s]) {
                        $cell.Changed = $true
                        $cell.Kind = $row.Kind
                    } elseif (($slots[$s][2] -eq 2 -or $slots[$s][2] -eq 5) -and $y -lt $rv.Count -and $rv[$y].Length -gt 0) {
                        # 追加した（動かした先の）列の値
                        $cell.Changed = $true
                        $cell.Kind = "insert"
                        $row.HasColumnChange = $true
                    }
                }
                $items[$s] = $cell
            }
            if (!$shared) { $row.RightCells = $items }
        } else {
            $row.RightEmpty = $true
        }
        if ($changedSlots.Count -gt 0) {
            $first = $slots[$changedSlots[0]]
            $row.LeftCell = "$(getColumnName ($first[0] + 1))$($pair[0] + 1)"
            $row.RightCell = "$(getColumnName ($first[1] + 1))$($pair[1] + 1)"
            $row.Detail = getCellChangeDetail $lc $rc $slots $changedSlots ($pair[0] + 1) ($pair[1] + 1)
        } elseif ($null -ne $lc -or $null -ne $rc) {
            $row.LeftCell = if ($null -ne $lc) { "A$($pair[0] + 1)" } else { "" }
            $row.RightCell = if ($null -ne $rc) { "A$($pair[1] + 1)" } else { "" }
        }
        $rows.Add($row)
    }
    $rowArray = $rows.ToArray()
    markMovedRows $rowArray $leftKeys $rightKeys
    $place = newPlaceDiff $name $leftName $rightName $rowArray
    $place.IsGrid = $true
    $place.ColumnWidths = $widths
    $place.LeftColumnNames = $leftNames
    $place.RightColumnNames = $rightNames
    $place.ColumnKinds = $kinds

    # 追加・削除・移動した列（空の列は数えない）。見出しには、その列の最初の値を添える
    $notes = New-Object System.Collections.Generic.List[string]
    foreach ($slot in $slots) {
        if ($slot[2] -eq 4) {
            $place.ColumnMoves++
            $to = $slots[$slot[3]][1]
            $notes.Add("列の移動 $(getColumnName ($slot[0] + 1)) → $(getColumnName ($to + 1))$(getColumnLabel $leftCells $slot[0])")
        } elseif ($slot[2] -eq 1) {
            $place.ColumnDeletes++
            $notes.Add("列の削除 $(getColumnName ($slot[0] + 1))$(getColumnLabel $leftCells $slot[0])")
        } elseif ($slot[2] -eq 2) {
            $place.ColumnInserts++
            $notes.Add("列の追加 $(getColumnName ($slot[1] + 1))$(getColumnLabel $rightCells $slot[1])")
        }
    }
    $place.ColumnNote = $notes -join "・"
    if ($place.Status -eq "same" -and $notes.Count -gt 0) {
        $place.Status = "change"
    }
    return $place
}

function getCellValues {
    # 比べるためのセルの値（行ごとの string[]）。見た目が空のセル（空白・見えない文字だけ）は ""、
    # 空白の違い・大文字と小文字の設定に合わせて整える。cells は splitExcelLines の結果
    param (
        [object[]]$cells,
        $options
    )

    # セルごとに関数を呼ぶと遅いため、見えない文字を含むセル・整える設定のあるときだけ呼ぶ
    $normalize = $options.IgnoreWhitespace -or !$options.CaseSensitive
    $invisible = [char[]]@([char]0x200B, [char]0xFEFF, [char]0x21B5)
    $result = [object[]]::new($cells.Count)
    for ($i = 0; $i -lt $cells.Count; $i++) {
        $source = $cells[$i]
        $values = [string[]]::new($source.Count)
        for ($c = 0; $c -lt $source.Count; $c++) {
            $value = $source[$c]
            if ($value.Length -gt 0) {
                if ([string]::IsNullOrWhiteSpace($value)) {
                    $value = ""
                } elseif ($value.IndexOfAny($invisible) -ge 0 -and (isBlankCell $value)) {
                    $value = ""
                } elseif ($normalize) {
                    $value = normalizeDiffLine $value $options.IgnoreWhitespace $options.CaseSensitive
                }
            }
            $values[$c] = $value
        }
        $result[$i] = $values
    }
    return , $result
}

function getRowKeys {
    # 行を比べるための形（string[]）。共通の列（columns。列の番号の並び）の値をタブでつなぎ、右端の空の値を落とす。
    # 共通の列に値が無く、追加・削除した列にだけ値のある行は、その行の値すべてで比べる（ほかの行と組まない）
    param (
        [object[]]$values,
        [int[]]$columns
    )

    $keys = [string[]]::new($values.Count)
    $parts = [string[]]::new($columns.Count)
    $mark = [string][char]1
    for ($i = 0; $i -lt $values.Count; $i++) {
        $row = $values[$i]
        for ($k = 0; $k -lt $columns.Count; $k++) {
            $c = $columns[$k]
            if ($c -lt $row.Count) { $parts[$k] = $row[$c] } else { $parts[$k] = "" }
        }
        $key = [string]::Join("`t", $parts).TrimEnd("`t")
        if ($key.Length -eq 0) {
            $all = [string]::Join("`t", $row).Trim("`t")
            if ($all.Length -gt 0) { $key = $mark + $all }
        }
        $keys[$i] = $key
    }
    return , $keys
}

function alignExcelColumns {
    # 左右の列を対応づけ、表に並べる順の組を返す。
    # 組は int[]（比較元の列, 比較先の列, 種類, 相手の位置）。種類は 0 = 対応・1 = 削除・2 = 追加・4 = 移動の元・5 = 移動の先。
    # 列は 0 から数え、無い側は -1。空の列（値のあるセルが 1 つも無い列）は組に入れない（空の列を足した・消しただけでは差分にしない）。
    # 列の値の並び（値のある行の先頭 ${diffColumnSampleRows} 行）が同じ列を目印にし、残りは値の重なりの似ている度合いで組む
    param (
        [object[]]$leftValues,
        [object[]]$rightValues
    )

    $leftColumns = getColumnKeys $leftValues
    $rightColumns = getColumnKeys $rightValues
    $slots = New-Object System.Collections.Generic.List[int[]]
    $a2 = $leftColumns.Keys
    $b2 = $rightColumns.Keys
    $dict = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $a = getLineKeys $a2 $dict
    $b = getLineKeys $b2 $dict
    $match = getLineMatches $a $b
    # 値の並びが同じ列を目印にし、その間の列は、左右の数が同じなら位置どうしで組む（行を足した・値を直しただけなら、
    # どの列も並びが変わるため）。数が違う区間だけ、値の重なりで追加・削除の列を決める
    $pairs = New-Object System.Collections.Generic.List[int[]]
    $i = 0
    $j = 0
    while ($i -lt $a2.Count -or $j -lt $b2.Count) {
        if ($i -lt $a2.Count -and $match[$i] -ge 0 -and $match[$i] -eq $j) {
            $pairs.Add([int[]]@($i, $j, 0))
            $i++
            $j++
            continue
        }
        $dels = [System.Collections.Generic.List[int]]::new()
        while ($i -lt $a2.Count -and $match[$i] -lt 0) {
            $dels.Add($i)
            $i++
        }
        $nextJ = if ($i -lt $a2.Count) { $match[$i] } else { $b2.Count }
        $inss = [System.Collections.Generic.List[int]]::new()
        while ($j -lt $nextJ) {
            $inss.Add($j)
            $j++
        }
        if ($dels.Count -eq $inss.Count) {
            for ($k = 0; $k -lt $dels.Count; $k++) { $pairs.Add([int[]]@($dels[$k], $inss[$k], 3)) }
        } else {
            addGapPairs $pairs $dels $inss $a2 $b2 { param($x, $y) getColumnSimilarity $x $y } ${diffPairMinColumnSimilarity}
        }
    }
    $keys = [System.Collections.Generic.List[string]]::new()
    foreach ($pair in $pairs) {
        $x = if ($pair[0] -ge 0) { $leftColumns.Index[$pair[0]] } else { -1 }
        $y = if ($pair[1] -ge 0) { $rightColumns.Index[$pair[1]] } else { -1 }
        $kind = switch ($pair[2]) { 1 { 1 } 2 { 2 } default { 0 } }
        $slots.Add([int[]]@($x, $y, $kind, -1))
        $keys.Add($(if ($pair[2] -eq 1) { $a2[$pair[0]] } elseif ($pair[2] -eq 2) { $b2[$pair[1]] } else { "" }))
    }
    # 動かした列（削除と追加の列で、値の並びが同じもの）は、種類を 4（元の位置）・5（動かした先）にし、4 つ目に相手の位置を入れる
    for ($s = 0; $s -lt $slots.Count; $s++) {
        if ($slots[$s][2] -ne 1) { continue }
        for ($t = 0; $t -lt $slots.Count; $t++) {
            if ($slots[$t][2] -eq 2 -and $keys[$t] -ceq $keys[$s]) {
                $slots[$s][2] = 4
                $slots[$s][3] = $t
                $slots[$t][2] = 5
                $slots[$t][3] = $s
                break
            }
        }
    }
    return , $slots.ToArray()
}

function getColumnKeys {
    # 値のある列ごとの、値の並び（値のある行の先頭 ${diffColumnSampleRows} 行の値を改行でつないだもの）。
    # 返すのは @{ Index = int[]（列の番号）; Keys = string[] }
    param (
        [object[]]$values
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $width = 0
    foreach ($row in $values) {
        if ($rows.Count -ge ${diffColumnSampleRows}) { break }
        $filled = $false
        foreach ($value in $row) { if ($value.Length -gt 0) { $filled = $true; break } }
        if (!$filled) { continue }
        $rows.Add($row)
        if ($row.Count -gt $width) { $width = $row.Count }
    }
    $index = [System.Collections.Generic.List[int]]::new()
    $keys = [System.Collections.Generic.List[string]]::new()
    $parts = [string[]]::new($rows.Count)
    for ($c = 0; $c -lt $width; $c++) {
        $filled = $false
        for ($k = 0; $k -lt $rows.Count; $k++) {
            $row = $rows[$k]
            if ($c -lt $row.Count) { $parts[$k] = $row[$c] } else { $parts[$k] = "" }
            if ($parts[$k].Length -gt 0) { $filled = $true }
        }
        if ($filled) {
            $index.Add($c)
            $keys.Add([string]::Join("`n", $parts))
        }
    }
    return @{ Index = $index.ToArray(); Keys = $keys.ToArray() }
}

function getColumnSimilarity {
    # 2 つの列（getColumnKeys の値の並び）の似ている度合い（0〜1）。値のあるセルのうち、同じ値のセルの割合
    # （多い方の数で割る。行の位置は見ない。行を足したり消したりしても、ほかの値は重なるため）
    param (
        [string]$x,
        [string]$y
    )

    if ($x -ceq $y) {
        return 1.0
    }
    $counts = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $leftCount = 0
    foreach ($value in $x.Split("`n")) {
        if ($value.Length -eq 0) { continue }
        $leftCount++
        $n = 0
        [void]$counts.TryGetValue($value, [ref]$n)
        $counts[$value] = $n + 1
    }
    $rightCount = 0
    $common = 0
    foreach ($value in $y.Split("`n")) {
        if ($value.Length -eq 0) { continue }
        $rightCount++
        $n = 0
        if ($counts.TryGetValue($value, [ref]$n) -and $n -gt 0) {
            $common++
            $counts[$value] = $n - 1
        }
    }
    $max = [Math]::Max($leftCount, $rightCount)
    if ($max -eq 0) {
        return 1.0
    }
    return [double]$common / $max
}

function getColumnLabel {
    # 追加・削除した列の説明に添える、その列の最初の値（"（役職）"。長ければ切る。値が無ければ ""）
    param (
        [object[]]$cells,
        [int]$column
    )

    foreach ($row in $cells) {
        if ($column -lt $row.Count -and ![string]::IsNullOrWhiteSpace($row[$column])) {
            $text = $row[$column]
            if ($text.Length -gt 20) { $text = $text.Substring(0, 20) + "…" }
            return "（$text）"
        }
    }
    return ""
}

function compareExcelObjects {
    # 図形・コメント（1 行が "<セル番地><TAB><文字>"）。文字の行として比べ、行番号の欄にセル番地を出す
    param (
        [string]$name,
        [string]$leftName,
        [string]$rightName,
        [string[]]$left,
        [string[]]$right,
        $options
    )

    $pairs = compareLines $left $right $options
    $leftAddress = @($left | ForEach-Object { ($_ -split "`t", 2)[0] })
    $rightAddress = @($right | ForEach-Object { ($_ -split "`t", 2)[0] })
    $leftText = [string[]]@($left | ForEach-Object { $parts = $_ -split "`t", 2; if ($parts.Count -gt 1) { unquoteCell $parts[1] } else { "" } })
    $rightText = [string[]]@($right | ForEach-Object { $parts = $_ -split "`t", 2; if ($parts.Count -gt 1) { unquoteCell $parts[1] } else { "" } })
    $rows = newTextRows $pairs $leftText $rightText ([string[]]$leftAddress) ([string[]]$rightAddress)
    $base = $name.Substring(0, $name.LastIndexOf("["))
    foreach ($row in $rows) {
        # 開くときは元のシートのそのセル
        $row.LeftPlace = $base
        $row.RightPlace = $base
        $row.LeftCell = $row.LeftNo
        $row.RightCell = $row.RightNo
    }
    return (newPlaceDiff $name $leftName $rightName $rows "同じ {0} 件（{1}〜{2}）")
}

function splitExcelLines {
    # TSV の行をセルに分ける（引用符で囲まれたセルは外す）
    param (
        [string[]]$lines
    )

    # 行ごと・セルごとに関数を呼ぶと遅いため、引用符を外すのは " で始まるセルだけ、その場で行う
    $cellNewLine = [string][char]0x2028
    $result = New-Object 'object[]' $lines.Count
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $cells = $lines[$i].Replace($cellNewLine, " ↵ ").Split("`t")
        for ($c = 0; $c -lt $cells.Count; $c++) {
            $cell = $cells[$c]
            if ($cell.Length -ge 2 -and $cell[0] -eq '"' -and $cell[$cell.Length - 1] -eq '"') {
                $cells[$c] = $cell.Substring(1, $cell.Length - 2).Replace('""', '"')
            }
        }
        $result[$i] = $cells
    }
    return , $result
}

function unquoteCell {
    # Excel のテキスト保存で " に囲まれたセル（改行・" を含む）を元の値に戻す。セル内の改行 U+2028 は ↵ にする
    param (
        [string]$cell
    )

    if ($cell.Length -ge 2 -and $cell.StartsWith('"') -and $cell.EndsWith('"')) {
        $cell = $cell.Substring(1, $cell.Length - 2).Replace('""', '"')
    }
    return $cell.Replace([string][char]0x2028, " ↵ ")
}

function getColumnName {
    # 列の番号（1 から）を Excel の列名（A, B, …, Z, AA, …）にする
    param (
        [int]$number
    )

    $name = ""
    while ($number -gt 0) {
        $number--
        $name = [string][char]([int][char]'A' + ($number % 26)) + $name
        $number = [int][Math]::Floor($number / 26)
    }
    return $name
}

function getGridColumnWidths {
    # 表の列の幅（左右で同じ）。両方の値の長い方に合わせる（全角は 2 文字分。先頭の 3000 行で測る）
    param (
        [object[]]$leftCells,
        [object[]]$rightCells,
        [int]$columns
    )

    # 表示幅は Shift_JIS のバイト数（半角 1・全角 2）で測る。セルごとに関数を呼ぶと遅いため、.NET の関数を直接呼ぶ
    $encoding = [System.Text.Encoding]::GetEncoding(932)
    $chars = New-Object int[] $columns
    foreach ($side in @($leftCells, $rightCells)) {
        $limit = [Math]::Min($side.Count, 3000)
        for ($i = 0; $i -lt $limit; $i++) {
            $cells = $side[$i]
            $count = [Math]::Min($cells.Count, $columns)
            for ($c = 0; $c -lt $count; $c++) {
                $cell = $cells[$c]
                if ($cell.Length -eq 0) { continue }
                $length = if ($cell.Length -gt 40) { 40 } else { $encoding.GetByteCount($cell) }
                if ($length -gt $chars[$c]) { $chars[$c] = [Math]::Min($length, 40) }
            }
        }
    }
    $widths = New-Object double[] $columns
    for ($c = 0; $c -lt $columns; $c++) {
        $widths[$c] = [Math]::Min(240, [Math]::Max(44, $chars[$c] * 7 + 14))
    }
    return , $widths
}

function getCellChangeDetail {
    # 変更の行の、違うセルの説明（"B8：10 → 12　D8：120,000 → 144,000"。番地が左右で違えば "B8→C9：…"）。5 つまで。
    # slots は alignExcelColumns の組、changed は違うセルの組の位置
    param (
        [string[]]$left,
        [string[]]$right,
        $slots,
        $changed,
        [int]$leftRow,
        [int]$rightRow
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($s in $changed) {
        if ($parts.Count -ge 5) {
            $parts.Add("ほか $($changed.Count - 5) セル")
            break
        }
        $x = $slots[$s][0]
        $y = $slots[$s][1]
        $leftAddress = "$(getColumnName ($x + 1))$leftRow"
        $rightAddress = "$(getColumnName ($y + 1))$rightRow"
        $address = if ($leftAddress -eq $rightAddress) { $leftAddress } else { "$leftAddress→$rightAddress" }
        $l = if ($x -lt $left.Count -and $left[$x] -ne "") { $left[$x] } else { "（空）" }
        $r = if ($y -lt $right.Count -and $right[$y] -ne "") { $right[$y] } else { "（空）" }
        $parts.Add("${address}：$l → $r")
    }
    return ($parts -join "　")
}

# ----------------------------------------------------------------------------
# Word
# ----------------------------------------------------------------------------

function getWordGroups {
    # Word の場所を、比べる単位（本文・脚注・コメント・図形・ヘッダー・フッター・その他）にまとめる。
    # 返すのは 単位の名前 → @{ Lines = string[]; Nos = string[]（行番号の欄）; Places = string[]（元の場所） } の順序付きの辞書
    param (
        $units
    )

    $pages = New-Object System.Collections.Generic.List[object]
    $shapes = New-Object System.Collections.Generic.List[object]
    $comments = New-Object System.Collections.Generic.List[object]
    $others = [ordered]@{}
    foreach ($name in @($units.Keys)) {
        $m = [regex]::Match($name, '^ページ(\d+)(\[図形\]|\[コメント\])?$')
        if ($m.Success) {
            $entry = @{ Page = [int]$m.Groups[1].Value; Name = $name }
            switch ($m.Groups[2].Value) {
                ""          { $pages.Add($entry) }
                "[図形]"    { $shapes.Add($entry) }
                "[コメント]" { $comments.Add($entry) }
            }
        } elseif ($name.EndsWith(${diffObjectKinds}.Comment)) {
            $comments.Add(@{ Page = [int]::MaxValue; Name = $name })
        } else {
            $others[$name] = $name
        }
    }

    $groups = [ordered]@{}
    if ($pages.Count -gt 0) { $groups["本文"] = joinWordPages $units $pages $true }
    if ($others.Contains("脚注")) { $groups["脚注"] = joinWordPages $units @(@{ Page = 0; Name = "脚注" }) $false }
    if ($comments.Count -gt 0) { $groups["コメント"] = joinWordPages $units $comments $true }
    if ($shapes.Count -gt 0) { $groups["図形"] = joinWordPages $units $shapes $true }
    foreach ($name in @($others.Keys)) {
        if ($name -eq "脚注") { continue }
        $groups[$name] = joinWordPages $units @(@{ Page = 0; Name = $name }) $false
    }
    return $groups
}

function joinWordPages {
    # ページの順につないだ行と、行番号の欄（"p.3 ¶2"。ページの無い場所は段落番号だけ）
    param (
        $units,
        [object[]]$entries,
        [bool]$withPage
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $nos = New-Object System.Collections.Generic.List[string]
    $places = New-Object System.Collections.Generic.List[string]
    foreach ($entry in @($entries | Sort-Object { $_.Page }, { $_.Name })) {
        $number = 0
        foreach ($line in @($units[$entry.Name])) {
            $number++
            $lines.Add([string]$line)
            $nos.Add($(if ($withPage -and $entry.Page -lt [int]::MaxValue) { "p.$($entry.Page) ¶$number" } else { "¶$number" }))
            $places.Add($entry.Name)
        }
    }
    return @{ Lines = $lines.ToArray(); Nos = $nos.ToArray(); Places = $places.ToArray() }
}

function compareWordUnits {
    param (
        $left,
        $right,
        $options
    )

    $leftGroups = getWordGroups $left
    $rightGroups = getWordGroups $right
    $places = New-Object System.Collections.Generic.List[object]
    foreach ($pair in (mergePlaceOrder @($leftGroups.Keys) @($rightGroups.Keys))) {
        $name = if ($pair.Right) { $pair.Right } else { $pair.Left }
        $empty = @{ Lines = [string[]]@(); Nos = [string[]]@(); Places = [string[]]@() }
        $l = $empty
        if ($pair.Left) { $l = $leftGroups[$pair.Left] }
        $r = $empty
        if ($pair.Right) { $r = $rightGroups[$pair.Right] }
        if ($l.Lines.Count -gt ${diffMaxLines} -or $r.Lines.Count -gt ${diffMaxLines}) {
            $places.Add((newTooLargePlace $name $pair.Left $pair.Right $l.Lines $r.Lines))
            continue
        }
        $leftKeys = getCompareKeys $l.Lines $options
        $rightKeys = getCompareKeys $r.Lines $options
        $pairs = compareKeyLines $leftKeys $rightKeys
        $rows = newTextRows $pairs $l.Lines $r.Lines $l.Nos $r.Nos $pair.Left $pair.Right
        markMovedRows $rows $leftKeys $rightKeys ""
        $places.Add((newPlaceDiff $name $pair.Left $pair.Right $rows "同じ段落 {0} 個（{1}〜{2}）"))
    }
    return , $places.ToArray()
}

# ----------------------------------------------------------------------------
# PowerPoint
# ----------------------------------------------------------------------------

function getSlides {
    # PowerPoint の場所を、スライドごと（本文・ノート・図形・コメント・非表示）とそれ以外に分ける。
    # 返すのは @{ Slides = スライドの配列（番号の順）; Others = 場所の名前 → 行 }
    param (
        $units
    )

    $slides = @{}
    $others = [ordered]@{}
    foreach ($name in @($units.Keys)) {
        $m = [regex]::Match($name, '^スライド(\d+)(（非表示）)?(_ノート|\[図形\]|\[コメント\])?$')
        if (!$m.Success) {
            $others[$name] = [string[]]@($units[$name])
            continue
        }
        $number = [int]$m.Groups[1].Value
        if (!$slides.ContainsKey($number)) {
            $slides[$number] = @{ Number = $number; Hidden = $false; Body = [string[]]@(); Notes = [string[]]@(); Shapes = [string[]]@(); Comments = [string[]]@(); Name = "スライド$($number.ToString('000'))" }
        }
        $slide = $slides[$number]
        if ($m.Groups[2].Success) { $slide.Hidden = $true }
        $lines = [string[]]@($units[$name])
        switch ($m.Groups[3].Value) {
            ""          { $slide.Body = $lines; $slide.Name = $name }
            "_ノート"    { $slide.Notes = $lines }
            "[図形]"    { $slide.Shapes = $lines }
            "[コメント]" { $slide.Comments = $lines }
        }
    }
    $ordered = @($slides.Keys | Sort-Object | ForEach-Object { $slides[$_] })
    return @{ Slides = $ordered; Others = $others }
}

function getSlideKey {
    # スライドの対応づけに使う文字（本文の空でない段落を改行でつないだもの）
    param (
        $slide,
        $options
    )

    $body = @($slide.Body | Where-Object { ![string]::IsNullOrWhiteSpace($_) })
    return (normalizeDiffLine ($body -join "`n") $options.IgnoreWhitespace $options.CaseSensitive)
}

function getSlideSimilarity {
    # 2 枚のスライドの似ている度合い（本文とノートをつないだ文字の似ている度合い。getTextSimilarity）
    param (
        $x,
        $y
    )

    $a = (@($x.Body) + @($x.Notes)) -join "`n"
    $b = (@($y.Body) + @($y.Notes)) -join "`n"
    return (getTextSimilarity $a $b)
}

function compareSlideUnits {
    # スライドを対応づけ、スライドごとに見出しの行と、本文・ノート・図形・コメントの行を並べる（1 つの場所「スライドとノート」）。
    # ヘッダー・フッターなど、スライドに属さない場所はそれぞれ 1 つの場所にする
    param (
        $left,
        $right,
        $options,
        $result
    )

    $l = getSlides $left
    $r = getSlides $right
    $result.LeftSlides = $l.Slides.Count
    $result.RightSlides = $r.Slides.Count

    # スライドの本文で対応づける（本文が同じスライドは同じ番号にし、残りは似ている度合いで組む）
    $dict = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $leftKeys = [string[]]@($l.Slides | ForEach-Object { getSlideKey $_ $options })
    $rightKeys = [string[]]@($r.Slides | ForEach-Object { getSlideKey $_ $options })
    $a = getLineKeys $leftKeys $dict
    $b = getLineKeys $rightKeys $dict
    $match = getLineMatches $a $b
    $leftSlides = $l.Slides
    $rightSlides = $r.Slides
    $similarity = { param($x, $y) 0.0 }  # 組むかどうかは下で判断する（本文の文字ではなくスライドの似ている度合いで）
    $pairs = getAlignedPairs $match $rightSlides.Count $leftKeys $rightKeys $similarity 2.0
    $pairs = pairSimilarSlides $pairs $leftSlides $rightSlides
    $moves = getMovedSlides $pairs $leftKeys $rightKeys $leftSlides $rightSlides

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($pair in $pairs) {
        $ls = if ($pair[0] -ge 0) { $leftSlides[$pair[0]] } else { $null }
        $rs = if ($pair[1] -ge 0) { $rightSlides[$pair[1]] } else { $null }
        $slideRows = newSlideRows $ls $rs $options
        # 並べ替えで動いたスライドは、見出しの説明に移動の元と先を書き、追加・削除ではなく移動として数える
        # （中の段落も、動かした行として件数に入れない）
        $move = $null
        if ($null -eq $rs -and $moves.Left.ContainsKey($pair[0])) { $move = $moves.Left[$pair[0]] }
        if ($null -eq $ls -and $moves.Right.ContainsKey($pair[1])) { $move = $moves.Right[$pair[1]]; $result.SlideMoves++ }
        if ($move) {
            $slideRows[0].Detail = (@($move, $slideRows[0].Detail) | Where-Object { $_ }) -join "・"
            foreach ($row in $slideRows) { $row.Moved = $true }
        } else {
            if ($null -eq $ls) { $result.SlideInserts++ }
            if ($null -eq $rs) { $result.SlideDeletes++ }
        }
        foreach ($row in $slideRows) { $rows.Add($row) }
    }
    $place = newPlaceDiff "スライドとノート" "スライド" "スライド" $rows.ToArray()
    # 動かしたスライドは SlideMoves で数えるため、中の段落を移動として数えない
    $place.Moves = 0
    $place.FoldedRows = foldSameSlides $rows.ToArray()
    # スライドの追加・削除・非表示の切り替えは、行の数に出ないことがあるため見出しの行でも判断する
    if ($place.Status -eq "same" -and @($rows | Where-Object { $_.Type -eq "Header" -and $_.Kind -ne "same" }).Count -gt 0) {
        $place.Status = "change"
    }
    $places = New-Object System.Collections.Generic.List[object]
    if ($pairs.Count -gt 0) {
        $places.Add($place)
    }

    foreach ($pair in (mergePlaceOrder @($l.Others.Keys) @($r.Others.Keys))) {
        $name = if ($pair.Right) { $pair.Right } else { $pair.Left }
        $lines = getUnitLines $l.Others $pair.Left
        $rlines = getUnitLines $r.Others $pair.Right
        $linePairs = compareLines $lines $rlines $options
        $textRows = newTextRows $linePairs $lines $rlines (getParagraphNos $lines.Count) (getParagraphNos $rlines.Count) $pair.Left $pair.Right
        $places.Add((newPlaceDiff $name $pair.Left $pair.Right $textRows "同じ段落 {0} 個（{1}〜{2}）"))
    }
    $result.Places = $places.ToArray()
}

function getMovedSlides {
    # 削除と追加のスライドで本文が同じもの（並べ替えで動いたスライド）を組み、説明の文字を返す。
    # 返すのは @{ Left = 比較元の位置 → "スライド 4 へ移動"; Right = 比較先の位置 → "スライド 1 から移動" }
    param (
        $pairs,
        [string[]]$leftKeys,
        [string[]]$rightKeys,
        [object[]]$leftSlides,
        [object[]]$rightSlides
    )

    $result = @{ Left = @{}; Right = @{} }
    $deleted = @{}
    foreach ($pair in $pairs) {
        if ($pair[1] -lt 0 -and $leftKeys[$pair[0]]) {
            if (!$deleted.ContainsKey($leftKeys[$pair[0]])) { $deleted[$leftKeys[$pair[0]]] = New-Object System.Collections.Generic.Queue[int] }
            $deleted[$leftKeys[$pair[0]]].Enqueue($pair[0])
        }
    }
    foreach ($pair in $pairs) {
        if ($pair[0] -ge 0) { continue }
        $key = $rightKeys[$pair[1]]
        if (!$key -or !$deleted.ContainsKey($key) -or $deleted[$key].Count -eq 0) { continue }
        $from = $deleted[$key].Dequeue()
        $result.Left[$from] = "スライド $($rightSlides[$pair[1]].Number) へ移動"
        $result.Right[$pair[1]] = "スライド $($leftSlides[$from].Number) から移動"
    }
    return $result
}

function pairSimilarSlides {
    # 本文が一致しなかったスライド（削除・追加の続き）を、似ている度合いが ${diffPairMinSlideSimilarity} 以上の組を順に対応づける
    param (
        $pairs,
        [object[]]$leftSlides,
        [object[]]$rightSlides
    )

    $result = New-Object System.Collections.Generic.List[int[]]
    $i = 0
    while ($i -lt $pairs.Count) {
        if ($pairs[$i][2] -eq 0) {
            $result.Add($pairs[$i])
            $i++
            continue
        }
        $dels = New-Object System.Collections.Generic.List[int]
        $inss = New-Object System.Collections.Generic.List[int]
        while ($i -lt $pairs.Count -and $pairs[$i][2] -ne 0) {
            if ($pairs[$i][0] -ge 0) { $dels.Add($pairs[$i][0]) }
            if ($pairs[$i][1] -ge 0) { $inss.Add($pairs[$i][1]) }
            $i++
        }
        $leftBodies = [string[]]@(0..([Math]::Max($leftSlides.Count, 1) - 1) | ForEach-Object { "$_" })
        $rightBodies = [string[]]@(0..([Math]::Max($rightSlides.Count, 1) - 1) | ForEach-Object { "$_" })
        # GetNewClosure は使わない（閉じた scriptblock からは、読み込んだ関数が見えなくなる）。
        # 呼ばれる addGapPairs の呼び出し元（この関数）の $leftSlides・$rightSlides を使う
        $similarity = { param($x, $y) getSlideSimilarity $leftSlides[[int]$x] $rightSlides[[int]$y] }
        addGapPairs $result $dels $inss $leftBodies $rightBodies $similarity ${diffPairMinSlideSimilarity}
    }
    return , $result
}

function newSlideRows {
    # スライド 1 組の行: 見出し、本文、ノート・図形・コメント（それぞれ小見出しの後に）
    param (
        $left,
        $right,
        $options
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $parts = @(
        @{ Key = "Body"; Title = "" }
        @{ Key = "Notes"; Title = "ノート" }
        @{ Key = "Shapes"; Title = "図形" }
        @{ Key = "Comments"; Title = "コメント" }
    )
    $changed = $false
    $partRows = New-Object System.Collections.Generic.List[object]
    foreach ($part in $parts) {
        $l = getUnitLines $left $part.Key
        $r = getUnitLines $right $part.Key
        if ($l.Count -eq 0 -and $r.Count -eq 0) {
            continue
        }
        $pairs = compareLines $l $r $options
        $lineRows = newTextRows $pairs $l $r (getParagraphNos $l.Count) (getParagraphNos $r.Count) $(if ($left) { $left.Name } else { "" }) $(if ($right) { $right.Name } else { "" })
        if ($part.Title) {
            $section = [DiffRow]::new()
            $section.Type = "Section"
            $section.LeftText = $part.Title
            $section.RightText = $part.Title
            $section.LeftEmpty = ($l.Count -eq 0)
            $section.RightEmpty = ($r.Count -eq 0)
            $partRows.Add($section)
        }
        foreach ($row in $lineRows) {
            if ($row.Kind -ne "same") { $changed = $true }
            $partRows.Add($row)
        }
    }

    $header = [DiffRow]::new()
    $header.Type = "Header"
    if ($null -eq $left) {
        $header.Kind = "insert"
        $header.LeftEmpty = $true
    } elseif ($null -eq $right) {
        $header.Kind = "delete"
        $header.RightEmpty = $true
    } elseif ($changed -or $left.Hidden -ne $right.Hidden) {
        $header.Kind = "change"
    }
    if ($left) {
        $header.LeftNo = "スライド $($left.Number)"
        $header.LeftText = getSlideTitle $left
        $header.LeftPlace = $left.Name
    }
    if ($right) {
        $header.RightNo = "スライド $($right.Number)"
        $header.RightText = getSlideTitle $right
        $header.RightPlace = $right.Name
    }
    $notes = New-Object System.Collections.Generic.List[string]
    if ($left -and $right -and $left.Number -ne $right.Number) {
        $notes.Add("比較元はスライド $($left.Number)")
    }
    if ($left -and $right -and $left.Hidden -ne $right.Hidden) {
        $notes.Add($(if ($right.Hidden) { "非表示にした" } else { "表示にした" }))
    } elseif ($right -and $right.Hidden) {
        $notes.Add("非表示")
    } elseif ($left -and !$right -and $left.Hidden) {
        $notes.Add("非表示")
    }
    $header.Detail = $notes -join "・"
    $rows.Add($header)
    foreach ($row in $partRows) { $rows.Add($row) }
    return , $rows.ToArray()
}

function getSlideTitle {
    # 見出しに出すスライドの名前（本文の最初の段落。長ければ切る）
    param (
        $slide
    )

    $first = @($slide.Body | Where-Object { $_.Trim() -ne "" } | Select-Object -First 1)
    if ($first.Count -eq 0) {
        return ""
    }
    $text = [string]$first[0]
    if ($text.Length -gt 40) {
        $text = $text.Substring(0, 40) + "…"
    }
    return $text
}

function foldSameSlides {
    # 変わらないスライド（見出しが same で、中の行もすべて same）が続くところを、1 行にまとめる
    param (
        [object[]]$rows
    )

    # スライドごとに分ける
    $slides = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        if ($row.Type -eq "Header") {
            $slides.Add((New-Object System.Collections.Generic.List[object]))
        }
        if ($slides.Count -gt 0) { $slides[$slides.Count - 1].Add($row) }
    }
    $result = New-Object System.Collections.Generic.List[object]
    $run = New-Object System.Collections.Generic.List[object]
    $flush = {
        if ($run.Count -eq 0) { return }
        $hidden = New-Object System.Collections.Generic.List[object]
        foreach ($slide in $run) { foreach ($row in $slide) { $hidden.Add($row) } }
        $fold = [DiffRow]::new()
        $fold.Type = "Fold"
        $fold.FoldCount = $run.Count
        $fold.FoldRows = $hidden.ToArray()
        $first = $run[0][0]
        $last = $run[$run.Count - 1][0]
        $fold.LeftText = if ($run.Count -eq 1) { "変わらないスライド 1 枚（$($first.LeftNo)）" } else { "変わらないスライド $($run.Count) 枚（$($first.LeftNo)〜$($last.LeftNo)）" }
        $fold.RightText = if ($run.Count -eq 1) { "変わらないスライド 1 枚（$($first.RightNo)）" } else { "変わらないスライド $($run.Count) 枚（$($first.RightNo)〜$($last.RightNo)）" }
        $result.Add($fold)
        $run.Clear()
    }
    foreach ($slide in $slides) {
        $same = $true
        foreach ($row in $slide) {
            if ($row.Kind -ne "same") { $same = $false; break }
        }
        if ($same) {
            $run.Add($slide)
        } else {
            & $flush
            foreach ($row in $slide) { $result.Add($row) }
        }
    }
    & $flush
    return , $result.ToArray()
}
