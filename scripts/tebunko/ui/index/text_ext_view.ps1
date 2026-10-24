# インデックスごとの「取り込む拡張子」（インデックスを取り込む前の確認ダイアログから開くダイアログ）の、文言と判断（判断層）。
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
        return $rows.ToArray()
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

function getIndexTextExtensionSummary {
    # 確認ダイアログの行と、ダイアログの要約。rows は @{ Extension; Count; Checked } の並び。
    # 数え中（件数が $null）は「対象のフォルダを数えています…」、何も外していなければ「テキストファイル 13,150 件 ・ すべて取り込む」、
    # 外していれば「取り込む 10,350 件（外した 2,800 件）」。件数が分からないとき（数え中）に外していれば「外した N 個」
    param (
        [object[]]$rows
    )

    $all = @($rows)
    $removed = @($all | Where-Object { -not $_.Checked })
    $counting = (@($all | Where-Object { $null -eq $_.Count }).Count -gt 0)
    if ($counting) {
        if ($removed.Count -eq 0) {
            return "対象のフォルダを数えています…（件数は数え終わると出ます）"
        }
        return "外した拡張子 $($removed.Count) 個（対象のフォルダを数えています…）"
    }
    $total = [int](@($all | ForEach-Object { [int]$_.Count } | Measure-Object -Sum).Sum)
    if ($removed.Count -eq 0) {
        return "テキストファイル $(formatTextExtensionCount $total) 件 ・ すべて取り込む"
    }
    $removedCount = [int](@($removed | ForEach-Object { [int]$_.Count } | Measure-Object -Sum).Sum)
    return "取り込む $(formatTextExtensionCount ($total - $removedCount)) 件（外した $(formatTextExtensionCount $removedCount) 件）"
}

function getIndexTextExtensionListView {
    # ダイアログの一覧の見せ方。rows は @{ Extension; Count; Checked } の並び、filter は絞り込み欄の文字、removedOnly は［外したものだけ］。
    # 返す値: @{ Visible（見せる拡張子の並び）; Summary（要約）; Empty（一致が無いか）; ZeroCount（件数 0 の拡張子の数）}
    # 絞り込みは、拡張子に含まれるか（大文字・小文字を区別せず、先頭の . は有っても無くてもよい）。要約は
    # 何も絞らないとき取り込む件数の要約、文字で絞るとき「75 個中 2 個を表示」、［外したものだけ］のとき「外したもの N 個を表示」
    param (
        [object[]]$rows,
        [string]$filter = "",
        [bool]$removedOnly = $false
    )

    $needle = ([string]$filter).Trim().ToLowerInvariant().TrimStart(".")
    $visible = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($rows)) {
        if ($removedOnly -and $row.Checked) {
            continue
        }
        if ($needle -ne "" -and -not $row.Extension.TrimStart(".").Contains($needle)) {
            continue
        }
        $visible.Add($row.Extension)
    }
    if ($needle -ne "") {
        $summary = "$(@($rows).Count) 個中 $($visible.Count) 個を表示"
    } elseif ($removedOnly) {
        $summary = "外したもの $($visible.Count) 個を表示"
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

function getIndexTextExtensionDialogTexts {
    # ダイアログの文言。indexName はインデックスの名前
    param (
        [string]$indexName
    )

    return @{
        Title             = "取り込む拡張子 - ${indexName}"
        Note              = "チェックを外した拡張子のファイルは、次のインデックス作成で、このインデックス（${indexName}）から外します。件数は、対象のフォルダにあるファイルを数えた結果です（数えた時点の件数）。Excel・Word・PowerPoint は、ここでは選べません。"
        Extension         = "拡張子"
        Count             = "件数"
        Reset             = "既定に戻す"
        UncheckAll        = "すべて外す"
        Ok                = "OK"
        Cancel            = "キャンセル"
        RemovedOnly       = "外したものだけ"
        FilterPlaceholder = "拡張子で絞り込み"
        Empty             = "一致する拡張子がありません"
    }
}
