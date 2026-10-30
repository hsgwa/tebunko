# インデックスごとの「取り込み対象の拡張子」（インデックスを取り込む前の確認ダイアログから開くダイアログ）の、文言と判断（判断層）。
# 画面に触らないため、そのままテストできる（tests/tebunko/ui/index/text_ext_view.Tests.ps1）。
# 選べるのは shared\core\text_file.ps1 の ${textExtensions} の拡張子だけ（足せない。Excel・Word・PowerPoint・PDF は選べない）。
# 設定は core\settings.ps1 の indexTextExcludes（インデックス名 → 取り込まない拡張子）。件数は、インデックス作成の数える段階（findTargetFiles の TextCounts）が
# 数えた結果（拡張子 → 件数）。開き方（既定のアプリ・エディタ）は ui\settings\text_open_view.ps1。

function getIndexTextExtensionRows {
    # ダイアログの行。@{ Extension; Count（件数。数え中・数えていないときは $null）; Checked（取り込む）; Changed（既定から変えた＝外した） } の並び。
    # 件数の多い順（同じ件数は拡張子の名前順）。件数が 0 の拡張子は、末尾にまとめる（件数が分かっていないときは、固定の一覧の順）。
    # counts は拡張子 → 件数のハッシュテーブル（$null なら数え中）、excluded は取り込まない拡張子（そのインデックスの設定）
    param (
        $counts = $null,
        [object[]]$excluded = @()
    )

    $excludedSet = @(@($excluded) | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($extension in ${textExtensions}) {
        $count = $null
        if ($null -ne $counts) {
            $count = [int]$counts[$extension]
        }
        $checked = ($excludedSet -notcontains $extension)
        $rows.Add(@{ Extension = $extension; Count = $count; Checked = $checked; Changed = (-not $checked) })
    }
    if ($null -eq $counts) {
        # 件数が無いときは名前順
        return @($rows | Sort-Object -Property @{ Expression = { $_.Extension }; Descending = $false })
    }
    return @($rows | Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Extension }; Descending = $false })
}

function convertIndexTextExtensionRows {
    # ダイアログで選んだ行（@{ Extension; Checked } の並び）から、設定に書く、取り込まない拡張子の配列を作る（${textExtensions} の順）。
    # ${textExtensions} に無い拡張子は捨てる
    param (
        [object[]]$rows
    )

    $unchecked = @(@($rows) | Where-Object { -not $_.Checked } | ForEach-Object { ([string]$_.Extension).ToLowerInvariant() })
    return @(${textExtensions} | Where-Object { $unchecked -contains $_ })
}

function formatTextExtensionCount {
    # 件数を 3 桁区切りの文字にする。数え中（$null）は「…」
    param (
        $count
    )

    if ($null -eq $count) {
        return "…"
    }
    return ([int]$count).ToString("N0", [System.Globalization.CultureInfo]::InvariantCulture)
}

function getIndexTextExtensionCountSummary {
    # 取り込み対象のファイル数の要約。何も外していなければ「すべて取り込み（10,350 件）」、外していれば「10,350 件（対象外 2,800 件）」。
    # excluded は対象外にした拡張子、counts は拡張子 → 件数（${textExtensions} の中のものだけ数える）。件数が無い（$null）ときは空
    param (
        [object[]]$excluded = @(),
        $counts = $null
    )

    if ($null -eq $counts) {
        return ""
    }
    $skip = @(@($excluded) | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $kept = 0
    $removed = 0
    foreach ($extension in @($counts.Keys)) {
        if (${textExtensions} -notcontains ([string]$extension).ToLowerInvariant()) { continue }
        if ($skip -contains ([string]$extension).ToLowerInvariant()) { $removed += [int]$counts[$extension] } else { $kept += [int]$counts[$extension] }
    }
    if ($removed -eq 0) {
        return "すべて取り込み（$(formatTextExtensionCount $kept) 件）"
    }
    return "$(formatTextExtensionCount $kept) 件（対象外 $(formatTextExtensionCount $removed) 件）"
}

function getIndexTextExtensionBriefSummary {
    # 詳細の行・確認の表の要約。件数があれば getIndexTextExtensionCountSummary。件数が無い（$null）ときは、「更新すると、件数が出ます」（拡張子の数は言わない）
    param (
        [object[]]$excluded = @(),
        $counts = $null
    )

    if ($null -ne $counts) {
        return getIndexTextExtensionCountSummary $excluded $counts
    }
    return "更新すると、件数が出ます"
}

function testIndexTextExcludesChanged {
    # 今の設定（current）が、前回取り込んだときの設定（applied）と違うか。違えば［要更新］。
    # 順・大文字小文字・重なりは無視して比べる。$null や空は「何も外していない」
    param (
        $current,
        $applied
    )

    $normalize = { param ($list) (@(@($list) | Where-Object { $null -ne $_ -and ([string]$_) -ne "" } | ForEach-Object { ([string]$_).ToLowerInvariant() } | Sort-Object -Unique) -join "`n") }
    return ((& $normalize $current) -ne (& $normalize $applied))
}

function getIndexTextExtensionCountNote {
    # 詳細の「取り込む拡張子」の行の、件数の見出し。source は件数の出どころ:
    # previous（前回の控え。time は控えの日時）→「前回の更新（日時）の件数です」、none（控え無し）→「更新すると、件数が出ます」、now（確認ダイアログ）→ 空
    param (
        [string]$source = "none",
        [string]$time = ""
    )

    if ($source -eq "now") {
        return ""
    }
    if ($source -eq "previous" -and $time -ne "") {
        return "前回の更新（${time}）の件数です"
    }
    return "更新すると、件数が出ます"
}

function getIndexTextExtensionDetailRow {
    # 詳細の「取り込む拡張子」の行。excluded はそのインデックスの取り込まない拡張子、entry は件数の控え（readTextExtCounts のその名前の値。無ければ $null）。
    # 戻り値: @{ Summary（要約）; Caption（件数の見出し）; Source（previous か none）; Time（控えの日時）; Counts（控えの件数。無ければ $null） }
    param (
        [object[]]$excluded = @(),
        $entry = $null
    )

    $has = ($null -ne $entry -and $null -ne $entry.Counts)
    $time = $(if ($has) { [string]$entry.Time } else { "" })
    $source = $(if ($has) { "previous" } else { "none" })
    return @{
        Summary = getIndexTextExtensionBriefSummary $excluded $(if ($has) { $entry.Counts } else { $null })
        Caption = $(if ($has) { getIndexTextExtensionCountNote $source $time } else { "" })
        Source  = $source
        Time    = $time
        Counts  = $(if ($has) { $entry.Counts } else { $null })
    }
}

function getIndexTextExtensionDetailView {
    # 詳細の「取り込む拡張子」の行の見せ方。item は詳細の行（TextExcludes・TextExcludesIngested・TextCountsEntry・TextExtEditable。省いた項目は、全部取り込む・控え無し・押せる）。
    # $null（何も選んでいない）なら空で、［変更…］は押せない。
    # 戻り値: @{ Summary（要約）; Caption（件数の見出し）; Stale（［要更新］の札を出すか。今の設定と前回取り込んだときの差）; ChangeEnabled（［変更…］を押せるか）;
    #            CountsSource・CountsTime・Counts（ダイアログに渡す件数の出どころ・日時・件数）}
    param (
        $item = $null
    )

    if ($null -eq $item) {
        return @{ Summary = ""; Caption = ""; Stale = $false; ChangeEnabled = $false; CountsSource = "none"; CountsTime = ""; Counts = $null }
    }
    $row = getIndexTextExtensionDetailRow @($item.TextExcludes) $item.TextCountsEntry
    return @{
        Summary = $row.Summary; Caption = $row.Caption
        Stale = (testIndexTextExcludesChanged $item.TextExcludes $item.TextExcludesIngested)
        ChangeEnabled = ($item.TextExtEditable -ne $false)
        CountsSource = $row.Source; CountsTime = $row.Time; Counts = $row.Counts
    }
}

function getIndexTextExtensionSummary {
    # ダイアログの上の要約。rows は @{ Extension; Count; Checked } の並び。件数があれば「取り込み対象 10,350 件（対象外 2,800 件）」、
    # 何も外していなければ「取り込み対象 13,150 件（対象外 0 件）」。件数が無い（$null）ときは getIndexTextExtensionBriefSummary
    param (
        [object[]]$rows
    )

    $all = @($rows)
    $removed = @($all | Where-Object { -not $_.Checked })
    if (@($all | Where-Object { $null -eq $_.Count }).Count -gt 0) {
        return getIndexTextExtensionBriefSummary @($removed | ForEach-Object { $_.Extension })
    }
    $total = [int](@($all | ForEach-Object { [int]$_.Count } | Measure-Object -Sum).Sum)
    $removedCount = [int](@($removed | ForEach-Object { [int]$_.Count } | Measure-Object -Sum).Sum)
    return "取り込み対象 $(formatTextExtensionCount ($total - $removedCount)) 件（対象外 $(formatTextExtensionCount $removedCount) 件）"
}

function getIndexTextExtensionListView {
    # ダイアログの一覧の見せ方。rows は @{ Extension; Count; Checked } の並び、filter は絞り込み欄の文字、keptOnly は［取り込むものだけ］。
    # 返す値: @{ Visible（見せる拡張子の並び）; Summary（要約）; Empty（一致が無いか）; ZeroCount（件数 0 の拡張子の数）}
    # 絞り込みは、拡張子に含まれるか（大文字・小文字を区別せず、先頭の . は有っても無くてもよい）。要約は
    # 何も絞らないとき取り込む件数の要約、文字で絞るとき「75 個中 2 個を表示」、［取り込むものだけ］のとき「取り込むもの N 個を表示」
    param (
        [object[]]$rows,
        [string]$filter = "",
        [bool]$keptOnly = $false
    )

    $needle = ([string]$filter).Trim().ToLowerInvariant().TrimStart(".")
    $visible = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($rows)) {
        if ($keptOnly -and -not $row.Checked) {
            continue
        }
        if ($needle -ne "" -and -not $row.Extension.TrimStart(".").Contains($needle)) {
            continue
        }
        $visible.Add($row.Extension)
    }
    if ($needle -ne "") {
        $summary = "$(@($rows).Count) 個中 $($visible.Count) 個を表示"
    } else {
        $summary = getIndexTextExtensionSummary $rows
    }
    return @{
        Visible   = $visible.ToArray()
        Summary   = $summary
        Empty     = ($visible.Count -eq 0)
        ZeroCount = @(@($rows) | Where-Object { $null -ne $_.Count -and $_.Count -eq 0 }).Count
    }
}

function getIndexTextExtensionZeroHeader {
    # 一覧の末尾の、件数 0 の拡張子をまとめた行の文。expanded は開いているか
    param (
        [int]$count,
        [bool]$expanded = $false
    )

    if ($expanded) {
        return "$([char]0x25BE) 件数 0 の拡張子 ${count} 個（開いて選ぶ。見つかったときに効く設定）"
    }
    return "$([char]0x25B8) 件数 0 の拡張子 ${count} 個（折りたたみ中。開くと選べます）"
}

function getIndexTextExtensionDialogTexts {
    # ダイアログの文言。indexName はインデックスの名前。
    # countsSource は件数の出どころ（previous = 詳細から・前回の件数、now = 確認ダイアログから・今数えた件数、none = 件数なし）、countsTime は前回の控えの日時。
    # 注の文は、件数の出どころで 3 通りに切り替わる
    param (
        [string]$indexName,
        [string]$countsSource = "none",
        [string]$countsTime = ""
    )

    $countSentence = switch ($countsSource) {
        "now" { "件数は、いま数えた結果です。" }
        "previous" { $(if ($countsTime -ne "") { "件数は、前回の更新（${countsTime}）で数えた結果です。" } else { "まだ件数がありません。更新すると、件数が出ます。" }) }
        default { "まだ件数がありません。更新すると、件数が出ます。" }
    }
    return @{
        Title             = "取り込み対象の拡張子 - ${indexName}"
        Note              = "チェックを外した拡張子は対象外になり、そのファイルは、次のインデックス作成で、このインデックス（${indexName}）から外します。${countSentence}Excel・Word・PowerPoint は、ここでは選べません。"
        Extension         = "拡張子"
        Count             = "件数 $([char]0x25BC)"
        Reset             = "既定に戻す"
        UncheckAll        = "すべて外す"
        Ok                = "OK"
        Cancel            = "キャンセル"
        KeptOnly          = "取り込み対象のみ表示"
        FilterPlaceholder = "拡張子で絞り込み"
        Empty             = "一致する拡張子がありません"
    }
}

function getTextExtCountsFromPlan {
    # 取り込みを終えた回の計画の行（kind の行だけ）から、インデックス名 → @{ Time; Counts } の表を作る（件数の控えと同じ形）。
    # ネットワークの場所のワークスペースでは、画面が控えのファイルを読まないので、終わり方 0 の回の件数をこの形でメモリに持つ
    param (
        $plan,
        [string]$kind,
        [string]$time
    )

    $result = @{}
    foreach ($row in @($plan | Where-Object { $null -ne $_ -and $_.区分 -eq $kind })) {
        $counts = @{}
        foreach ($extension in @($row.TextExtensionCounts.Keys)) { $counts[[string]$extension] = [int]$row.TextExtensionCounts[$extension] }
        $result[[string]$row.インデックス名] = @{ Time = $time; Counts = $counts }
    }
    return $result
}
