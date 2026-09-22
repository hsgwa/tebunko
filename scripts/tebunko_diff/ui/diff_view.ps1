# ［2 比較］タブに出す文言・件数と、操作の可否（判断層。$ui・WPF に触らない。入力も出力も素の値）。

${diffKinds} = @{
    ".xlsx" = "Excel"; ".xlsm" = "Excel"; ".xls" = "Excel"; ".xlsb" = "Excel"
    ".docx" = "Word"; ".docm" = "Word"; ".doc" = "Word"
    ".pptx" = "PowerPoint"; ".pptm" = "PowerPoint"; ".ppt" = "PowerPoint"
}

function getDiffKind {
    # ファイル名の拡張子から、Excel / Word / PowerPoint を返す（Office でなければ ""）
    param (
        [string]$path
    )

    $extension = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
    if (${diffKinds}.ContainsKey($extension)) {
        return ${diffKinds}[$extension]
    }
    return ""
}

function getModeView {
    # トグル（file / folder）で変わる、入力欄の文言と、使えるオプション
    param (
        [string]$mode
    )

    if ($mode -eq "folder") {
        return @{
            Placeholder       = "フォルダを選ぶか、ここへドロップ"
            PickLabel         = "フォルダ…"
            FolderOptions     = $true
            Note              = "フォルダの中の Office ファイルを、相対パスで対応づけて比べます。"
            FolderOptionsTip  = ""
        }
    }
    return @{
        Placeholder       = "ファイルを選ぶか、ここへドロップ"
        PickLabel         = "ファイル…"
        FolderOptions     = $false
        Note              = "同じ種類のファイル（Excel 同士・Word 同士・PowerPoint 同士）を比べます。"
        FolderOptionsTip  = "フォルダを比べるときだけ使えます"
    }
}

function getCompareBlockReason {
    # ［比較］を押せない理由（押せるなら ""）。left・right は @{ Path; Exists; IsFolder }（Exists・IsFolder は呼び出し側が調べる）
    param (
        [string]$mode,
        $left,
        $right,
        [bool]$subfolders = $true
    )

    $what = if ($mode -eq "folder") { "フォルダ" } else { "ファイル" }
    if ([string]::IsNullOrWhiteSpace($left.Path) -and [string]::IsNullOrWhiteSpace($right.Path)) {
        return "比較元と比較先の${what}を選んでください。"
    }
    if ([string]::IsNullOrWhiteSpace($left.Path)) {
        return "比較元の${what}を選んでください。"
    }
    if ([string]::IsNullOrWhiteSpace($right.Path)) {
        return "比較先の${what}を選んでください。"
    }
    foreach ($side in @(@{ Name = "比較元"; Item = $left }, @{ Name = "比較先"; Item = $right })) {
        $item = $side.Item
        if (!$item.Exists) {
            return "$($side.Name)の${what}が見つかりません。"
        }
        if ($mode -eq "folder" -and !$item.IsFolder) {
            return "$($side.Name)はフォルダではありません。［ファイル］に切り替えるか、フォルダを選んでください。"
        }
        if ($mode -ne "folder" -and $item.IsFolder) {
            return "$($side.Name)はフォルダです。［フォルダ］に切り替えるか、ファイルを選んでください。"
        }
    }
    $leftFull = $left.Path.TrimEnd("\")
    $rightFull = $right.Path.TrimEnd("\")
    if ([string]::Equals($leftFull, $rightFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return "同じ${what}が選ばれています。"
    }
    if ($mode -eq "folder") {
        if ($subfolders -and ($rightFull.StartsWith("$leftFull\", [System.StringComparison]::OrdinalIgnoreCase))) {
            return "比較先のフォルダが比較元のフォルダの中にあります。"
        }
        if ($subfolders -and ($leftFull.StartsWith("$rightFull\", [System.StringComparison]::OrdinalIgnoreCase))) {
            return "比較元のフォルダが比較先のフォルダの中にあります。"
        }
        return ""
    }
    $leftKind = getDiffKind $left.Path
    $rightKind = getDiffKind $right.Path
    if ($leftKind -eq "" -or $rightKind -eq "") {
        return "Excel・Word・PowerPoint のファイルを選んでください。"
    }
    if ($leftKind -ne $rightKind) {
        return "種類の違うファイル（$leftKind と $rightKind）は比べられません。"
    }
    return ""
}

function getDropAction {
    # ドロップされたもの（items: @{ Path; IsFolder } の配列）と、ドロップ先（target: left / right / ""）から、
    # 入力欄に入れる内容を返す: @{ Mode; Left; Right }（変えない側は $null）。入れられなければ $null
    param (
        [object[]]$items,
        [string]$target,
        [string]$mode
    )

    $items = @($items | Where-Object { $_ -and $_.Path })
    if ($items.Count -eq 0) {
        return $null
    }
    if ($items.Count -ge 2) {
        # 2 つまとめて: 名前の順に比較元・比較先へ。ファイルかフォルダかは 1 つ目に合わせる（そろっていなければ入れない）
        $pair = @($items | Select-Object -First 2 | Sort-Object { $_.Path })
        if ($pair[0].IsFolder -ne $pair[1].IsFolder) {
            return $null
        }
        $newMode = if ($pair[0].IsFolder) { "folder" } else { "file" }
        return @{ Mode = $newMode; Left = $pair[0].Path; Right = $pair[1].Path }
    }
    $item = $items[0]
    $newMode = if ($item.IsFolder) { "folder" } else { "file" }
    $side = if ($target -eq "right") { "right" } else { "left" }
    if ($side -eq "left") {
        return @{ Mode = $newMode; Left = $item.Path; Right = $null }
    }
    return @{ Mode = $newMode; Left = $null; Right = $item.Path }
}

function getFolderProgressText {
    # フォルダ同士で中身を比べている間の進み具合（"中身を比べています 5 / 9"）。終わっていれば ""
    param (
        $counts,
        [int]$compareTotal
    )

    $waiting = $counts.pending + $counts.running
    if ($waiting -eq 0 -or $compareTotal -le 0) {
        return ""
    }
    return "中身を比べています $($compareTotal - $waiting) / $compareTotal"
}

function getPlaceTabText {
    # 場所の見出しの文字（"明細 4"・"条件 追加"・"表紙"）
    param (
        $place
    )

    switch ($place.Status) {
        "insert" { return "$($place.Name)（追加）" }
        "delete" { return "$($place.Name)（削除）" }
        "change" {
            $count = $place.Inserts + $place.Deletes + $place.Changes
            if ($count -gt 0) { return "$($place.Name)  $count" }
            return $place.Name
        }
    }
    return $place.Name
}

function getDefaultPlaceIndex {
    # 最初に開く場所（違いのある最初の場所。無ければ先頭）
    param (
        [object[]]$places
    )

    for ($i = 0; $i -lt $places.Count; $i++) {
        if ($places[$i].Status -ne "same") {
            return $i
        }
    }
    return 0
}

function findNextChangeRow {
    # 行（rows）で、from の次（direction = 1）・前（-1）の変更の塊の先頭の位置。無ければ -1。
    # 変更の行（same 以外の Line・Header）が続くところは 1 つの塊とし、その先頭で止まる
    param (
        [object[]]$rows,
        [int]$from,
        [int]$direction
    )

    $i = $from + $direction
    while ($i -ge 0 -and $i -lt $rows.Count) {
        if ((isChangedRow $rows[$i]) -and ($i -eq 0 -or !(isChangedRow $rows[$i - 1]))) {
            return $i
        }
        $i += $direction
    }
    return -1
}

function isChangedRow {
    param ($row)

    return (($row.Type -eq "Line" -or $row.Type -eq "Header") -and $row.Kind -ne "same")
}

function getExtractFailureText {
    # 抽出に失敗したときに、差分の代わりに出す文
    param (
        [string]$side,
        [string]$message
    )

    $name = if ($side -eq "right") { "比較先" } else { "比較元" }
    return "${name}を読み取れませんでした。（$message）"
}
