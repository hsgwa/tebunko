# フォルダ同士の比較（判断層）。
#   ・左右のファイルを相対パスで対応づける（getFolderEntries）
#   ・ファイルの状態（同じ・中身は同じ・変更・追加・削除・比較できない・比較待ち・比較中）を決める
#   ・左右に並べるツリーの行を組み立てる（buildTreeRows。WPF の TreeView は使わず、開いている行だけの平らな一覧にする）
#
# ファイルは @{ RelPath; Path; Size; Time } の形で受け取る（Time は更新日時の DateTime）。

# 状態と、ツリーの中央に出す印
${diffStatusMarks} = @{
    same    = "="   # 同じ（バイトが同じ）
    similar = "≈"   # 中身は同じ（バイトは違うが、抽出した文字が同じ）
    change  = "≠"   # 変更
    insert  = "+"   # 追加（比較先だけ）
    delete  = "−"   # 削除（比較元だけ）
    failed  = "!"   # 比較できない
    pending = "…"   # 比較待ち
    running = "…"   # 比較中
}

# 左右に並べるツリーの 1 行（画面はこのプロパティをそのまま表示する）
class TreeRow {
    [int]$Depth
    [double]$Indent             # 字下げ（px）
    [bool]$IsFolder
    [string]$Name
    [string]$RelPath
    [string]$Status             # ${diffStatusMarks} のキー
    [string]$Mark
    [bool]$LeftEmpty            # 比較元に無い（斜線の空き）
    [bool]$RightEmpty
    [string]$LeftMeta = ""      # 名前の右に小さく出す文字（サイズ・更新日時など）
    [string]$RightMeta = ""
    [bool]$Expanded             # フォルダを開いている
    [string]$Glyph = ""         # フォルダの開閉の印（▾ / ▸）
}


# 名前の並べ順。Windows の表示言語に左右されず、どの PC でも日本語の並びにする
${diffSortCulture} = "ja-JP"
function getFolderEntries {
    # 左右のファイルを相対パス（大文字・小文字を区別しない）で対応づけ、相対パスの順に返す。
    # 1 件は @{ RelPath; Left; Right; Status; Inserts; Deletes; Changes; Error }。Status は、
    #   片方だけ → insert / delete、両方にあってサイズが違う → pending（中身を比べる）、サイズが同じ → pending（ハッシュを比べる。HashNeeded）
    param (
        [object[]]$leftFiles,
        [object[]]$rightFiles
    )

    $map = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $leftFiles) {
        $map[$file.RelPath] = @{ RelPath = $file.RelPath; Left = $file; Right = $null }
    }
    foreach ($file in $rightFiles) {
        if ($map.ContainsKey($file.RelPath)) {
            $map[$file.RelPath].Right = $file
        } else {
            $map[$file.RelPath] = @{ RelPath = $file.RelPath; Left = $null; Right = $file }
        }
    }
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($key in @($map.Keys | Sort-Object -Culture ${diffSortCulture})) {
        $entry = $map[$key]
        $entry.Inserts = 0
        $entry.Deletes = 0
        $entry.Changes = 0
        $entry.Error = ""
        $entry.HashNeeded = $false
        if ($null -eq $entry.Left) {
            $entry.Status = "insert"
        } elseif ($null -eq $entry.Right) {
            $entry.Status = "delete"
        } else {
            $entry.Status = "pending"
            $entry.HashNeeded = ([long]$entry.Left.Size -eq [long]$entry.Right.Size)
        }
        $result.Add($entry)
    }
    return , $result.ToArray()
}

function setEntryResult {
    # 中身を比べた結果（FileDiff）を、ファイルの状態に入れる。差が無ければ「中身は同じ」
    param (
        $entry,
        $fileDiff
    )

    $entry.Inserts = $fileDiff.Inserts + $fileDiff.SlideInserts
    $entry.Deletes = $fileDiff.Deletes + $fileDiff.SlideDeletes
    # 移動（行・スライド・列）と列の追加・削除も、ツリーの件数では「変更」に数える
    $entry.Changes = $fileDiff.Changes + $fileDiff.Moves + $fileDiff.SlideMoves + $fileDiff.ColumnInserts + $fileDiff.ColumnDeletes + $fileDiff.ColumnMoves
    $changed = @($fileDiff.Places | Where-Object { $_.Status -ne "same" }).Count -gt 0
    if ($changed -and ($entry.Inserts + $entry.Deletes + $entry.Changes) -eq 0) {
        # シート名を変えただけ等（行の数に出ない変更）
        $entry.Changes = 1
    }
    $entry.Status = if ($changed -or ($entry.Inserts + $entry.Deletes + $entry.Changes) -gt 0) { "change" } else { "similar" }
}

function getFolderCounts {
    # 状態ごとのファイルの数
    param (
        [object[]]$entries
    )

    $counts = [ordered]@{ Total = 0; change = 0; insert = 0; delete = 0; failed = 0; similar = 0; same = 0; pending = 0; running = 0 }
    foreach ($entry in $entries) {
        $counts.Total++
        $counts[$entry.Status]++
    }
    return $counts
}

function newFolderNode {
    param (
        [string]$name,
        [string]$relPath
    )

    return @{ Name = $name; RelPath = $relPath; Folders = [ordered]@{}; Files = New-Object System.Collections.Generic.List[object] }
}

function buildFolderTree {
    # 相対パスからフォルダの木を作る（名前の順。フォルダの名前は大文字・小文字を区別せずにまとめる）
    param (
        [object[]]$entries
    )

    $root = newFolderNode "" ""
    foreach ($entry in $entries) {
        $parts = $entry.RelPath.Split("\")
        $node = $root
        for ($i = 0; $i -lt $parts.Count - 1; $i++) {
            $key = $parts[$i].ToLowerInvariant()
            if (!$node.Folders.Contains($key)) {
                $rel = ($parts[0..$i] -join "\")
                $node.Folders[$key] = newFolderNode $parts[$i] $rel
            }
            $node = $node.Folders[$key]
        }
        $node.Files.Add($entry)
    }
    return $root
}

function getFolderStats {
    # フォルダの中（下のフォルダも含む）のファイルの状態をまとめる
    param (
        $node
    )

    $stats = @{ Total = 0; Left = 0; Right = 0; Diff = 0; Waiting = 0; Failed = 0 }
    foreach ($entry in $node.Files) {
        $stats.Total++
        if ($null -ne $entry.Left) { $stats.Left++ }
        if ($null -ne $entry.Right) { $stats.Right++ }
        switch ($entry.Status) {
            { $_ -in @("change", "insert", "delete") } { $stats.Diff++ }
            { $_ -in @("pending", "running") } { $stats.Waiting++ }
            "failed" { $stats.Failed++ }
        }
    }
    foreach ($child in $node.Folders.Values) {
        $sub = getFolderStats $child
        foreach ($key in @($stats.Keys)) { $stats[$key] += $sub[$key] }
    }
    return $stats
}

function getFolderStatus {
    # フォルダの状態（中のファイルから決める）
    param (
        $stats
    )

    if ($stats.Left -eq 0) { return "insert" }
    if ($stats.Right -eq 0) { return "delete" }
    if ($stats.Waiting -gt 0) { return "running" }
    if ($stats.Failed -gt 0 -and $stats.Diff -eq 0) { return "failed" }
    if ($stats.Diff -gt 0 -or $stats.Failed -gt 0) { return "change" }
    return "same"
}

function buildTreeRows {
    # 左右に並べるツリーの行を、開いているフォルダの中だけ平らに並べて返す（TreeRow の配列）。
    #   expanded / collapsed : 利用者が開いた・閉じたフォルダの相対パス（HashSet。大文字・小文字を区別しない）。
    #                          どちらにも無いフォルダは、中に違い（変更・追加・削除・比較できない・比較中）があれば開く
    #   hideSame             : 同じ（=）ファイルと、中がすべて同じフォルダを出さない
    param (
        [object[]]$entries,
        $expanded,
        $collapsed,
        [bool]$hideSame
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $root = buildFolderTree $entries
    addTreeRows $rows $root 0 $expanded $collapsed $hideSame
    return , $rows.ToArray()
}

function addTreeRows {
    param (
        $rows,
        $node,
        [int]$depth,
        $expanded,
        $collapsed,
        [bool]$hideSame
    )

    foreach ($key in @($node.Folders.Keys | Sort-Object -Culture ${diffSortCulture})) {
        $child = $node.Folders[$key]
        $stats = getFolderStats $child
        $status = getFolderStatus $stats
        if ($hideSame -and $status -eq "same") {
            continue
        }
        $isOpen = if ($expanded -and $expanded.Contains($child.RelPath)) { $true }
                  elseif ($collapsed -and $collapsed.Contains($child.RelPath)) { $false }
                  else { $status -ne "same" }
        $row = [TreeRow]::new()
        $row.Depth = $depth
        $row.Indent = $depth * 18
        $row.IsFolder = $true
        $row.Name = $child.Name
        $row.RelPath = $child.RelPath
        $row.Status = $status
        $row.Mark = ${diffStatusMarks}[$status]
        $row.Expanded = $isOpen
        $row.Glyph = if ($isOpen) { "▾" } else { "▸" }
        $row.LeftEmpty = ($stats.Left -eq 0)
        $row.RightEmpty = ($stats.Right -eq 0)
        $row.RightMeta = getFolderMeta $status $stats
        $rows.Add($row)
        if ($isOpen) {
            addTreeRows $rows $child ($depth + 1) $expanded $collapsed $hideSame
        }
    }
    foreach ($entry in @($node.Files | Sort-Object { [System.IO.Path]::GetFileName($_.RelPath) } -Culture ${diffSortCulture})) {
        if ($hideSame -and $entry.Status -eq "same") {
            continue
        }
        $row = [TreeRow]::new()
        $row.Depth = $depth
        $row.Indent = $depth * 18 + 14
        $row.Name = [System.IO.Path]::GetFileName($entry.RelPath)
        $row.RelPath = $entry.RelPath
        $row.Status = $entry.Status
        $row.Mark = ${diffStatusMarks}[$entry.Status]
        $row.LeftEmpty = ($null -eq $entry.Left)
        $row.RightEmpty = ($null -eq $entry.Right)
        if ($entry.Left) { $row.LeftMeta = getFileMeta $entry.Left }
        if ($entry.Right) { $row.RightMeta = (getEntryNote $entry) + (getFileMeta $entry.Right) }
        $rows.Add($row)
    }
}

function getFolderMeta {
    # フォルダの行の右側に出す文字
    param (
        [string]$status,
        $stats
    )

    switch ($status) {
        "insert"  { return "フォルダ · $($stats.Total) ファイル" }
        "running" { return "比較中 $($stats.Total - $stats.Waiting) / $($stats.Total)" }
        "failed"  { return "比較できない $($stats.Failed)" }
        "change"  { return "違い $($stats.Diff + $stats.Failed)" }
    }
    return ""
}

function getEntryNote {
    # ファイルの行の右側の先頭に出す、結果の説明（"変更 6 · " など）
    param (
        $entry
    )

    switch ($entry.Status) {
        "change"  { return "$(getChangeCountText $entry.Inserts $entry.Deletes $entry.Changes) · " }
        "similar" { return "中身は同じ · " }
        "failed"  { return "比較できない · " }
        "running" { return "比較中… · " }
        "pending" { return "比較待ち · " }
    }
    return ""
}

function getChangeCountText {
    # "追加 2・削除 1・変更 3"（0 の種類は出さない。すべて 0 なら "違いなし"）
    param (
        [int]$inserts,
        [int]$deletes,
        [int]$changes
    )

    $parts = @()
    if ($inserts -gt 0) { $parts += "追加 $inserts" }
    if ($deletes -gt 0) { $parts += "削除 $deletes" }
    if ($changes -gt 0) { $parts += "変更 $changes" }
    if ($parts.Count -eq 0) { return "違いなし" }
    return ($parts -join "・")
}

function getFileMeta {
    # ファイルのサイズと更新日時（"48 KB · 2024/10/15"）
    param (
        $file
    )

    return "$(formatFileSize ([long]$file.Size)) · $(([datetime]$file.Time).ToString('yyyy/MM/dd'))"
}

function formatFileSize {
    # バイト数を読みやすい大きさにする（1 KB 未満は B、1 MB 未満は KB、それ以上は MB）
    param (
        [long]$bytes
    )

    if ($bytes -lt 1024) { return "$bytes B" }
    if ($bytes -lt 1048576) { return "$([Math]::Ceiling($bytes / 1024)) KB" }
    return "$([Math]::Round($bytes / 1048576, 1).ToString('0.0')) MB"
}

function findNextDiffRow {
    # ツリーの行（rows）で、from の次（direction = 1）・前（-1）の違うファイル（変更・追加・削除・比較できない）の位置。無ければ -1
    param (
        [object[]]$rows,
        [int]$from,
        [int]$direction
    )

    $i = $from + $direction
    while ($i -ge 0 -and $i -lt $rows.Count) {
        $row = $rows[$i]
        if (!$row.IsFolder -and $row.Status -in @("change", "insert", "delete", "failed")) {
            return $i
        }
        $i += $direction
    }
    return -1
}

function getFolderPathsToOpen {
    # 相対パス（ファイル）までのフォルダの相対パスの一覧（Ctrl+↓ で閉じたフォルダの中へ移るとき、開くフォルダ）
    param (
        [string]$relPath
    )

    $parts = $relPath.Split("\")
    $result = @()
    for ($i = 0; $i -lt $parts.Count - 1; $i++) {
        $result += ($parts[0..$i] -join "\")
    }
    return , $result
}

function getEntriesInTreeOrder {
    # ファイルを、ツリーに並べる順（フォルダが先、同じ階層は名前の順）で返す。閉じたフォルダの中も含む
    param (
        [object[]]$entries
    )

    $result = New-Object System.Collections.Generic.List[object]
    addEntriesInTreeOrder $result (buildFolderTree $entries)
    return , $result.ToArray()
}

function addEntriesInTreeOrder {
    param (
        $result,
        $node
    )

    foreach ($key in @($node.Folders.Keys | Sort-Object -Culture ${diffSortCulture})) {
        addEntriesInTreeOrder $result $node.Folders[$key]
    }
    foreach ($entry in @($node.Files | Sort-Object { [System.IO.Path]::GetFileName($_.RelPath) } -Culture ${diffSortCulture})) {
        $result.Add($entry)
    }
}

function findNextDiffEntry {
    # ツリーの順で、relPath の次（direction = 1）・前（-1）の違うファイルの相対パス。無ければ ""（relPath が "" なら先頭・末尾から探す）
    param (
        [object[]]$entries,
        [string]$relPath,
        [int]$direction
    )

    $ordered = getEntriesInTreeOrder $entries
    $index = -1
    if ($relPath) {
        for ($i = 0; $i -lt $ordered.Count; $i++) {
            if ([string]::Equals($ordered[$i].RelPath, $relPath, [System.StringComparison]::OrdinalIgnoreCase)) { $index = $i; break }
        }
    }
    if ($index -lt 0 -and $direction -lt 0) { $index = $ordered.Count }
    $i = $index + $direction
    while ($i -ge 0 -and $i -lt $ordered.Count) {
        if ($ordered[$i].Status -in @("change", "insert", "delete", "failed")) {
            return $ordered[$i].RelPath
        }
        $i += $direction
    }
    return ""
}
