# ［2 検索］タブの判断（検索できるか・注意書き・検索条件の説明）。
# 画面に触らないため、そのままテストできる（tests\tebunko_grep\ui\search_view.Tests.ps1）。

function describeSearchOption {
    # 既定から変えた検索条件を「大文字と小文字を区別・対象ファイル：*.xlsx」のように返す（無ければ空）
    param (
        [hashtable]$option
    )

    $items = @()
    if ($option.CaseSensitive) {
        $items += "大文字と小文字を区別"
    }
    if ($option.FileFilter) {
        $items += "対象ファイル：$($option.FileFilter)"
    }
    # 既定はどちらも検索する（項目が無い古い形の条件も、検索するものとみなす）
    if ($option.ContainsKey("IncludeShapes") -and -not $option.IncludeShapes) {
        $items += "図形を除く"
    }
    if ($option.ContainsKey("IncludeComments") -and -not $option.IncludeComments) {
        $items += "コメントを除く"
    }
    return ($items -join "・")
}

function getWordNotice {
    # 検索ワードの下に出す注意書き（出さないときは空文字列）
    param (
        [string]$word,
        [bool]$useRegex
    )

    if ($useRegex -and $word -ne "" -and !(isValidRegex $word)) {
        return "正規表現として不正なため、文字どおり検索します。"
    }
    return ""
}

function getFastSearchView {
    # 検索ワードの下に出す、高速検索（Windows Search で先に絞る）の使用可否。
    #   available: Windows Search が使えるか（testWindowsSearch）。$null はまだ確かめていない（使えるものとして扱う）
    param (
        $available,
        [bool]$useRegex,
        [string]$word
    )

    $usable = testFastSearchUsable ($available -ne $false) $useRegex $word
    return @{ Usable = $usable; Text = if ($usable) { "高速検索：使用可" } else { "高速検索：使用不可" } }
}

function getSearchProgressText {
    # 検索中の要約欄
    param (
        [int]$hits
    )

    return "検索中…　該当 $($hits.ToString('N0')) 件"
}

function getSearchSummaryText {
    # 検索が終わったときの要約欄（ヒットがあるとき）
    param (
        [int]$hits,
        [int]$files,
        [double]$seconds
    )

    return "該当 $($hits.ToString('N0')) 件（$($files.ToString('N0')) ファイル） ・ $($seconds.ToString('0.0')) 秒"
}

function newSearchButtonState {
    # ［検索］ボタンの文言と、押せるかどうか
    param (
        [bool]$searching,    # 検索中か
        [bool]$stopping,     # 中止を頼んだ後か
        [string]$word,       # 検索ワード
        [bool]$hasIndex,     # 検索できるインデックスがあるか
        [int]$targetCount    # 検索対象に選ばれている数
    )

    if ($searching) {
        return @{ Content = "中止"; Enabled = !$stopping }
    }
    return @{ Content = "検索"; Enabled = ($word -ne "" -and $hasIndex -and $targetCount -gt 0) }
}

# ---- ファイルごとにまとめた表示 ----

function getAppKind {
    # 元のファイル名の拡張子から、アプリの種類（Excel / Word / PowerPoint。どれでもなければ空）を返す
    param (
        [string]$book
    )

    $extension = [System.IO.Path]::GetExtension($book).ToLowerInvariant()
    if ($extension -match '^\.xls') {
        return "Excel"
    }
    if ($extension -match '^\.doc') {
        return "Word"
    }
    if ($extension -match '^\.ppt') {
        return "PowerPoint"
    }
    return ""
}

function describeFileLocations {
    # ファイルの中でヒットした場所（見つかった順・重複なし）を、見出しの右端に出す文字列にする。
    # 1 か所ならその場所、2 か所以上なら「[シート] 4月 ほか 2 か所」（場所の表記は describePlace）
    param (
        [string[]]$labels
    )

    # 検索中にファイル・場所が増えるたびに呼ぶため、パイプライン（Where-Object）を使わない（1 回 1 ms を超えて検索が遅くなる）
    $first = ""
    $count = 0
    foreach ($label in $labels) {
        if ($label) {
            if ($count -eq 0) {
                $first = $label
            }
            $count++
        }
    }
    if ($count -le 1) {
        return $first
    }
    return "$first ほか $($count - 1) か所"
}

# ---- 結果の表に並べる項目（見出しと行） ----
# group は FileGroup（types_grep.ps1）と同じ項目（Rows・ShownRows・ShownCount・IsExpanded・Order）を持つもの、
# row は HitRow と同じ項目（Order・Contains(文字列)）を持つもの。

function selectShownRows {
    # 行のうち、絞り込み（空ならすべて）に合うものを、元の順のまま返す
    param (
        $rows,
        [string]$filterText
    )

    $shown = New-Object 'System.Collections.Generic.List[object]'
    foreach ($row in $rows) {
        if ($filterText -eq "" -or $row.Contains($filterText)) {
            $shown.Add($row)
        }
    }
    return , $shown
}

function getResultItems {
    # 結果の表に並べる項目。ファイルごとに見出しを 1 つ置き、開いているファイルだけ、その下に行を並べる。
    # 絞り込みで行が 1 つも残らないファイル（ShownCount が 0）は、見出しも出さない。開いているファイルの行は作ってあること
    param (
        $groups
    )

    $items = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $groups) {
        if ($group.ShownCount -eq 0) {
            continue
        }
        $items.Add($group)
        if ($group.IsExpanded) {
            $items.AddRange($group.ShownRows)
        }
    }
    return , $items
}

function getShownHitRows {
    # 絞り込みに合う行を、表の順（閉じているファイルの行も含む）に並べて返す（結果の出力に使う。行はすべて作ってあること）
    param (
        $groups
    )

    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $groups) {
        $rows.AddRange($group.ShownRows)
    }
    return , $rows
}

function sortFileGroups {
    # 列見出しのクリックでの並べ替え。各ファイルの中の行を property の順に並べ替え、
    # ファイルの順は、並べ替えた後の先頭の行の順にする。同じ値のときは見つかった順（Order）
    param (
        $groups,
        [string]$property,
        [bool]$descending
    )

    $byValue = @{ Expression = { $_.$property }; Descending = $descending }
    $byOrder = @{ Expression = { $_.Order }; Descending = $false }
    foreach ($group in $groups) {
        $sorted = @($group.Rows | Sort-Object $byValue, $byOrder)
        $group.Rows.Clear()
        $group.Rows.AddRange([object[]]$sorted)
    }
    $byFirst = @{ Expression = { if ($_.Rows.Count -gt 0) { $_.Rows[0].$property } }; Descending = $descending }
    $sortedGroups = New-Object 'System.Collections.Generic.List[object]'
    $sortedGroups.AddRange([object[]]@($groups | Sort-Object $byFirst, $byOrder))
    return , $sortedGroups
}