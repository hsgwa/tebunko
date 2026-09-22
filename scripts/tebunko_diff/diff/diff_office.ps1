# 種類ごとの比べ方（判断層）。場所（シート・ページ・スライド）ごとの行（抽出した TSV の中身）を受け取り、
# 画面に出す場所（PlaceDiff）と左右に並べる行（DiffRow）を組み立てる。
#   Excel      : シート名で場所を対応づけ、行の LCS で行の挿入・削除を見つける。対応した行はセルごとに比べる
#   Word       : ページをつないだ段落の並びで比べる（ページの区切りは目安で、動きやすいため対応づけに使わない）
#   PowerPoint : スライドの本文でスライドを対応づけ、その中の段落・ノート・図形・コメントを比べる
#
# 入力の units は、場所の名前 → 行（string[]）の順序付きの辞書（抽出した順）。options は getDiffOptions の形。

${diffMaxGridColumns} = 100  # Excel の表として出す列の上限（これより右の列は表に出さない。比べるときは全部の列を使う）
${diffObjectKinds}    = @{ Shape = "[図形]"; Comment = "[コメント]" }
${diffPairMinSlideSimilarity} = 0.4  # PowerPoint の本文が一致しないスライドを対応づける、似ている度合いの下限
${diffPairMinCellSimilarity} = 0.4  # Excel の削除と追加の行を「変更」として組む下限（同じセルの割合と、行の文字の似ている度合いの大きい方）

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
    [bool]$IsGrid               # Excel の表として出す
    [double[]]$ColumnWidths     # 表の列の幅（IsGrid のとき）
    [string[]]$ColumnNames      # 表の列の見出し（A, B, …）
    [string]$Note = ""          # 注意（大きすぎて行ごとに比べなかった等）
}

# 1 ファイルの比較の結果
class FileDiff {
    [string]$Kind               # Excel / Word / PowerPoint
    [object[]]$Places           # PlaceDiff
    [int]$Inserts
    [int]$Deletes
    [int]$Changes
    [int]$LeftSlides            # PowerPoint: スライドの枚数
    [int]$RightSlides
    [int]$SlideInserts          # PowerPoint: 追加・削除したスライドの枚数
    [int]$SlideDeletes
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
    # 行の並び 2 つを比べ、getAlignedPairs の組を返す。cells は Excel のセルの似ている度合いで組むとき
    param (
        [string[]]$left,
        [string[]]$right,
        $options,
        [bool]$cells = $false
    )

    $dict = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $a = getLineKeys $left $dict $options.IgnoreWhitespace $options.CaseSensitive
    $b = getLineKeys $right $dict $options.IgnoreWhitespace $options.CaseSensitive
    $match = getLineMatches $a $b
    if ($cells) {
        return , (getAlignedPairs $match $right.Count $left $right { param($x, $y) [Math]::Max((getCellSimilarity $x $y), (getTextSimilarity $x $y)) } ${diffPairMinCellSimilarity})
    }
    return , (getAlignedPairs $match $right.Count $left $right { param($x, $y) getTextSimilarity $x $y })
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

function toDisplayText {
    # TSV の 1 行を画面に出す文字にする（セル内の改行 U+2028 は ↵ にする）
    param (
        [string]$line
    )

    return $line.Replace([string][char]0x2028, " ↵ ")
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
    $place.Status = if (!$leftName) { "insert" } elseif (!$rightName) { "delete" } elseif (($counts.Insert + $counts.Delete + $counts.Change) -gt 0) { "change" } else { "same" }
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

    $places = New-Object System.Collections.Generic.List[object]
    foreach ($pair in (mergePlaceOrder @($left.Keys) @($right.Keys))) {
        $leftLines = getUnitLines $left $pair.Left
        $rightLines = getUnitLines $right $pair.Right
        $name = if ($pair.Right) { $pair.Right } else { $pair.Left }
        if ($leftLines.Count -gt ${diffMaxLines} -or $rightLines.Count -gt ${diffMaxLines}) {
            $places.Add((newTooLargePlace $name $pair.Left $pair.Right $leftLines $rightLines))
            continue
        }
        if ($name.EndsWith(${diffObjectKinds}.Shape) -or $name.EndsWith(${diffObjectKinds}.Comment)) {
            $places.Add((compareExcelObjects $name $pair.Left $pair.Right $leftLines $rightLines $options))
        } else {
            $places.Add((compareExcelSheet $name $pair.Left $pair.Right $leftLines $rightLines $options))
        }
    }
    return , $places.ToArray()
}

function compareExcelSheet {
    # シート 1 枚。行番号は TSV の行の位置 + 1（TSV の N 行目 = シートの N 行目）
    param (
        [string]$name,
        [string]$leftName,
        [string]$rightName,
        [string[]]$left,
        [string[]]$right,
        $options
    )

    $pairs = compareLines $left $right $options $true
    $leftCells = splitExcelLines $left
    $rightCells = splitExcelLines $right
    $columns = 0
    foreach ($cells in @($leftCells + $rightCells)) {
        if ($cells.Count -gt $columns) { $columns = $cells.Count }
    }
    $columns = [Math]::Max(1, [Math]::Min($columns, ${diffMaxGridColumns}))
    $widths = getGridColumnWidths $leftCells $rightCells $columns

    $kindNames = @("same", "delete", "insert", "change")
    $noChange = New-Object bool[] $columns
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($pair in $pairs) {
        $row = [DiffRow]::new()
        $row.Kind = $kindNames[$pair[2]]
        $row.LeftLine = $pair[0]
        $row.RightLine = $pair[1]
        $row.LeftPlace = $leftName
        $row.RightPlace = $rightName
        # `$x = if (...) { $array }` と書くと配列がばらされる（1 セルの行が文字列になる）ため、分けて代入する
        $lc = $null
        if ($pair[0] -ge 0) { $lc = $leftCells[$pair[0]] }
        $rc = $null
        if ($pair[1] -ge 0) { $rc = $rightCells[$pair[1]] }
        # New-Object はコマンドレットの呼び出しで遅いため、1 行ごとに回るところでは ::new を使う
        $changedColumns = [System.Collections.Generic.List[int]]::new()
        if ($row.Kind -eq "change") {
            $width = [Math]::Max($lc.Count, $rc.Count)
            for ($c = 0; $c -lt $width; $c++) {
                $l = if ($c -lt $lc.Count) { $lc[$c] } else { "" }
                $r = if ($c -lt $rc.Count) { $rc[$c] } else { "" }
                if ((normalizeDiffLine $l $options.IgnoreWhitespace $options.CaseSensitive) -ne (normalizeDiffLine $r $options.IgnoreWhitespace $options.CaseSensitive)) {
                    $changedColumns.Add($c)
                }
            }
        }
        # 違うセルの印（列ごと）。行ごと追加・削除なら全部の列。
        # 1 行ごとに回るところのため、関数や名前を組み立てたプロパティの読み書きを使わず、左右を別々に書く。
        # セルの部品は、値のある列（右端の空のセルは作らない）と違うセルの列の分だけ作る
        $changed = $noChange
        if ($changedColumns.Count -gt 0) {
            $changed = [bool[]]::new($columns)
            foreach ($c in $changedColumns) { if ($c -lt $columns) { $changed[$c] = $true } }
        }
        $whole = ($row.Kind -eq "delete" -or $row.Kind -eq "insert")
        $last = -1
        if ($changedColumns.Count -gt 0) { $last = [Math]::Min($changedColumns[$changedColumns.Count - 1], $columns - 1) }
        if ($null -ne $lc) {
            $row.LeftNo = [string]($pair[0] + 1)
            $row.LeftText = [string]::Join("`t", $lc)
            $count = [Math]::Max([Math]::Min($lc.Count, $columns), $last + 1)
            $items = [object[]]::new($count)
            for ($c = 0; $c -lt $count; $c++) {
                $cell = [DiffCell]::new()
                if ($c -lt $lc.Count) { $cell.Text = $lc[$c] } else { $cell.Text = "" }
                $cell.Width = $widths[$c]
                $cell.Changed = $whole -or $changed[$c]
                $items[$c] = $cell
            }
            $row.LeftCells = $items
        } else {
            $row.LeftEmpty = $true
        }
        if ($null -ne $rc) {
            $row.RightNo = [string]($pair[1] + 1)
            $row.RightText = [string]::Join("`t", $rc)
            $count = [Math]::Max([Math]::Min($rc.Count, $columns), $last + 1)
            $items = [object[]]::new($count)
            for ($c = 0; $c -lt $count; $c++) {
                $cell = [DiffCell]::new()
                if ($c -lt $rc.Count) { $cell.Text = $rc[$c] } else { $cell.Text = "" }
                $cell.Width = $widths[$c]
                $cell.Changed = $whole -or $changed[$c]
                $items[$c] = $cell
            }
            $row.RightCells = $items
        } else {
            $row.RightEmpty = $true
        }
        if ($changedColumns.Count -gt 0) {
            $first = getColumnName ($changedColumns[0] + 1)
            $row.LeftCell = "$first$($pair[0] + 1)"
            $row.RightCell = "$first$($pair[1] + 1)"
            $row.Detail = getCellChangeDetail $lc $rc $changedColumns ($pair[0] + 1) ($pair[1] + 1)
        } elseif ($null -ne $lc -or $null -ne $rc) {
            $row.LeftCell = if ($null -ne $lc) { "A$($pair[0] + 1)" } else { "" }
            $row.RightCell = if ($null -ne $rc) { "A$($pair[1] + 1)" } else { "" }
        }
        $rows.Add($row)
    }
    $place = newPlaceDiff $name $leftName $rightName $rows.ToArray()
    $place.IsGrid = $true
    $place.ColumnWidths = $widths
    $place.ColumnNames = @(1..$columns | ForEach-Object { getColumnName $_ })
    return $place
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
    # 変更の行の、違うセルの説明（"B8：10 → 12　D8：120,000 → 144,000"。行番号が違えば "B8→B9：…"）。5 つまで
    param (
        [string[]]$left,
        [string[]]$right,
        $columns,
        [int]$leftRow,
        [int]$rightRow
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($c in $columns) {
        if ($parts.Count -ge 5) {
            $parts.Add("ほか $($columns.Count - 5) セル")
            break
        }
        $column = getColumnName ($c + 1)
        $address = if ($leftRow -eq $rightRow) { "$column$leftRow" } else { "$column$leftRow→$column$rightRow" }
        $l = if ($c -lt $left.Count -and $left[$c] -ne "") { $left[$c] } else { "（空）" }
        $r = if ($c -lt $right.Count -and $right[$c] -ne "") { $right[$c] } else { "（空）" }
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
        $pairs = compareLines $l.Lines $r.Lines $options
        $rows = newTextRows $pairs $l.Lines $r.Lines $l.Nos $r.Nos $pair.Left $pair.Right
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
    # スライドの対応づけに使う文字（本文の段落を改行でつないだもの）
    param (
        $slide,
        $options
    )

    return (normalizeDiffLine ($slide.Body -join "`n") $options.IgnoreWhitespace $options.CaseSensitive)
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

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($pair in $pairs) {
        $ls = if ($pair[0] -ge 0) { $leftSlides[$pair[0]] } else { $null }
        $rs = if ($pair[1] -ge 0) { $rightSlides[$pair[1]] } else { $null }
        if ($null -eq $ls) { $result.SlideInserts++ }
        if ($null -eq $rs) { $result.SlideDeletes++ }
        foreach ($row in (newSlideRows $ls $rs $options)) { $rows.Add($row) }
    }
    $place = newPlaceDiff "スライドとノート" "スライド" "スライド" $rows.ToArray()
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
