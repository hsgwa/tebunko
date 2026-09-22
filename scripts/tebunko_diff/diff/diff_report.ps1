# 比較結果ファイル（［結果をファイルに出力］。work\tebunko_diff\比較結果.txt）の文字を作る（判断層）。

function getOptionsText {
    # 比べ方の設定を 1 行にする（"図形も比較・コメントも比較"）
    param (
        $options
    )

    $options = getDiffOptions $options
    $parts = @()
    if ($options.IncludeShapes) { $parts += "図形も比較" }
    if ($options.IncludeComments) { $parts += "コメントも比較" }
    if ($options.IgnoreWhitespace) { $parts += "空白の違いを無視" }
    if (!$options.CaseSensitive) { $parts += "大文字と小文字を区別しない" }
    if ($parts.Count -eq 0) { return "（なし）" }
    return ($parts -join "・")
}

function getFileReportLines {
    # 1 ファイルの比較の結果を、場所ごとに変更の行だけ並べた文字にする（見出しの行は含めない）
    param (
        $fileDiff
    )

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($place in $fileDiff.Places) {
        if ($place.Status -eq "same") {
            continue
        }
        $mark = switch ($place.Status) { "insert" { "（追加）" } "delete" { "（削除）" } default { "" } }
        $lines.Add("■ $($place.Name)$mark")
        if ($place.Note) {
            $lines.Add($place.Note)
        }
        if ($place.ColumnNote) {
            $lines.Add($place.ColumnNote)
        }
        foreach ($row in $place.Rows) {
            if ($row.Type -eq "Header" -and $row.Kind -ne "same") {
                $title = if ($row.RightNo) { $row.RightNo } else { $row.LeftNo }
                $lines.Add("□ $title $(getKindLabel $row.Kind)$(if ($row.Detail) { "（$($row.Detail)）" })")
                continue
            }
            if ($row.Type -ne "Line" -or $row.Kind -eq "same") {
                continue
            }
            $position = "$(if ($row.LeftEmpty) { "($(getNearNo $place.Rows $row 'Left'))" } else { $row.LeftNo }) → $(if ($row.RightEmpty) { "($(getNearNo $place.Rows $row 'Right'))" } else { $row.RightNo })"
            $content = switch ($row.Kind) {
                "change" { if ($row.Detail) { $row.Detail } else { "$($row.LeftText) → $($row.RightText)" } }
                "insert" { $row.RightText + $(if ($row.Detail) { "（$($row.Detail)）" }) }
                "delete" { $row.LeftText + $(if ($row.Detail) { "（$($row.Detail)）" }) }
            }
            $label = if ($row.Moved) { "移動" } else { getKindLabel $row.Kind }
            $lines.Add("$label`t$position`t$content")
        }
        $lines.Add("")
    }
    return , $lines.ToArray()
}

function getKindLabel {
    param (
        [string]$kind
    )

    switch ($kind) {
        "insert" { return "追加" }
        "delete" { return "削除" }
        "change" { return "変更" }
    }
    return "同じ"
}

function getNearNo {
    # 空きの側の位置の目安（直前にある、その側の行番号）
    param (
        [object[]]$rows,
        $row,
        [string]$side
    )

    $index = [Array]::IndexOf($rows, $row)
    for ($i = $index - 1; $i -ge 0; $i--) {
        $no = $rows[$i]."${side}No"
        if ($no -and !$rows[$i]."${side}Empty" -and $rows[$i].Type -eq "Line") {
            return $no
        }
    }
    return "先頭"
}

function getFileReport {
    # ファイル同士の比較結果ファイルの中身（行の配列）
    param (
        [string]$leftPath,
        [string]$rightPath,
        $fileDiff,
        $options,
        [datetime]$time
    )

    $summary = getFileSummaryParts $fileDiff
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("比較元: $leftPath")
    $lines.Add("比較先: $rightPath")
    $lines.Add("日時:   $($time.ToString('yyyy/MM/dd HH:mm:ss'))")
    $lines.Add("設定:   $(getOptionsText $options)")
    $lines.Add("要約:   $summary")
    $lines.Add("")
    foreach ($line in (getFileReportLines $fileDiff)) { $lines.Add($line) }
    return , $lines.ToArray()
}

function getFileSummaryParts {
    # 要約の 1 行（getFileSummaryText の見出しと件数をつないだもの）
    param (
        $fileDiff
    )

    $text = getFileSummaryText $fileDiff
    if ($text.Counts) {
        return "$($text.Title)（$($text.Counts)）"
    }
    return $text.Title
}

function getFolderReport {
    # フォルダ同士の比較結果ファイルの中身。先頭にファイルの一覧、その後に変更のあるファイルごとの中身
    #   fileDiffs: 相対パス → FileDiff（中身を比べたファイル）
    param (
        [string]$leftPath,
        [string]$rightPath,
        [object[]]$entries,
        $fileDiffs,
        $options,
        [bool]$subfolders,
        [datetime]$time
    )

    $counts = getFolderCounts $entries
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("比較元: $leftPath")
    $lines.Add("比較先: $rightPath")
    $lines.Add("日時:   $($time.ToString('yyyy/MM/dd HH:mm:ss'))")
    $lines.Add("設定:   $(getOptionsText $options)$(if ($subfolders) { '・サブフォルダも比較' })")
    $lines.Add("要約:   $(getFolderSummaryText $counts)")
    $lines.Add("")
    foreach ($entry in (getEntriesInTreeOrder $entries)) {
        $note = switch ($entry.Status) {
            "change" { getChangeCountText $entry.Inserts $entry.Deletes $entry.Changes }
            "failed" { $entry.Error }
            default  { "" }
        }
        $lines.Add("$(getStatusLabel $entry.Status)`t$($entry.RelPath)`t$note")
    }
    foreach ($entry in (getEntriesInTreeOrder $entries)) {
        if ($entry.Status -ne "change" -or !$fileDiffs.Contains($entry.RelPath)) {
            continue
        }
        $lines.Add("")
        $lines.Add("==== $($entry.RelPath)（$(getFileSummaryParts $fileDiffs[$entry.RelPath])）")
        foreach ($line in (getFileReportLines $fileDiffs[$entry.RelPath])) { $lines.Add($line) }
    }
    return , $lines.ToArray()
}

function getStatusLabel {
    # ファイルの状態の名前
    param (
        [string]$status
    )

    switch ($status) {
        "same"    { return "同じ" }
        "similar" { return "中身は同じ" }
        "change"  { return "変更" }
        "insert"  { return "追加" }
        "delete"  { return "削除" }
        "failed"  { return "比較できない" }
        "running" { return "比較中" }
    }
    return "比較待ち"
}

function getFileSummaryText {
    # ファイル同士の比較の要約（"シート 4 のうち 3 に変更"）と、件数の文字
    param (
        $fileDiff
    )

    if ($fileDiff.Kind -eq "PowerPoint") {
        $title = "スライド $($fileDiff.LeftSlides) → $($fileDiff.RightSlides)"
    } else {
        $unit = if ($fileDiff.Kind -eq "Excel") { "シート" } else { "場所" }
        $total = @($fileDiff.Places).Count
        $changed = @($fileDiff.Places | Where-Object { $_.Status -ne "same" }).Count
        $title = if ($changed -eq 0) { "違いはありません。（$unit $total）" } else { "$unit $total のうち $changed に変更" }
    }
    $parts = @()
    if ($fileDiff.SlideInserts -gt 0) { $parts += "スライド追加 $($fileDiff.SlideInserts)" }
    if ($fileDiff.SlideDeletes -gt 0) { $parts += "スライド削除 $($fileDiff.SlideDeletes)" }
    if ($fileDiff.SlideMoves -gt 0) { $parts += "スライド移動 $($fileDiff.SlideMoves)" }
    if ($fileDiff.ColumnInserts -gt 0) { $parts += "列追加 $($fileDiff.ColumnInserts)" }
    if ($fileDiff.ColumnDeletes -gt 0) { $parts += "列削除 $($fileDiff.ColumnDeletes)" }
    if ($fileDiff.ColumnMoves -gt 0) { $parts += "列移動 $($fileDiff.ColumnMoves)" }
    if ($fileDiff.Inserts -gt 0) { $parts += "追加 $($fileDiff.Inserts)" }
    if ($fileDiff.Deletes -gt 0) { $parts += "削除 $($fileDiff.Deletes)" }
    if ($fileDiff.Changes -gt 0) { $parts += "変更 $($fileDiff.Changes)" }
    if ($fileDiff.Moves -gt 0) { $parts += "移動 $($fileDiff.Moves)" }
    if ($fileDiff.Kind -eq "PowerPoint" -and $parts.Count -eq 0) {
        $title += "（違いはありません）"
    }
    return @{ Title = $title; Counts = ($parts -join "・") }
}

function getFolderSummaryText {
    # フォルダ同士の比較の要約（"ファイル 120（変更 3・追加 2・…）"）
    param (
        $counts
    )

    $parts = @()
    foreach ($item in @(
            @{ Key = "change"; Label = "変更" }, @{ Key = "insert"; Label = "追加" }, @{ Key = "delete"; Label = "削除" },
            @{ Key = "failed"; Label = "比較できない" }, @{ Key = "similar"; Label = "中身は同じ" }, @{ Key = "same"; Label = "同じ" })) {
        if ($counts[$item.Key] -gt 0) { $parts += "$($item.Label) $($counts[$item.Key])" }
    }
    $text = "ファイル $($counts.Total)"
    if ($parts.Count -gt 0) {
        $text += "（$($parts -join '・')）"
    }
    return $text
}
