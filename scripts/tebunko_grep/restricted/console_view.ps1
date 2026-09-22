# 制限モードのコンソールに出す文言と、入力の読み方（判断層。ファイルに触らない）。制限言語モードで動く書き方だけで書く。

# 検索の条件（readSearchOption の項目）と、メニューに出す名前
${restrictedOptionItems} = @(
    @{ Key = "UseRegex"; Label = "正規表現を使う" },
    @{ Key = "CaseSensitive"; Label = "大文字と小文字を区別する" },
    @{ Key = "FileFilter"; Label = "対象ファイル" },
    @{ Key = "IncludeShapes"; Label = "図形も検索する" },
    @{ Key = "IncludeComments"; Label = "コメントも検索する" }
)

# コンソールに一覧で出すヒットの数の上限（Excel が無い PC。多いと読めないため）
${restrictedConsoleListMax} = 200

function describeRestrictedOption {
    # 検索の条件を 1 行で表す（例: 文字どおり・大文字と小文字を区別しない・対象ファイル すべて・図形とコメントも検索）
    param (
        [hashtable]$option
    )

    $parts = @()
    $parts += $(if ($option.UseRegex) { "正規表現" } else { "文字どおり" })
    $parts += $(if ($option.CaseSensitive) { "大文字と小文字を区別する" } else { "大文字と小文字を区別しない" })
    $parts += $(if ([string]$option.FileFilter) { "対象ファイル $($option.FileFilter)" } else { "対象ファイル すべて" })
    if ($option.IncludeShapes -and $option.IncludeComments) {
        $parts += "図形とコメントも検索"
    } elseif ($option.IncludeShapes) {
        $parts += "図形も検索（コメントは除く）"
    } elseif ($option.IncludeComments) {
        $parts += "コメントも検索（図形は除く）"
    } else {
        $parts += "図形とコメントは除く"
    }
    return ($parts -join "・")
}

function getOptionMenuLines {
    # 条件を変えるメニュー（番号で切り替える）
    param (
        [hashtable]$option
    )

    $lines = @("検索の条件（番号で切り替えます。空で Enter: 戻る）")
    for ($i = 0; $i -lt ${restrictedOptionItems}.Count; $i++) {
        $item = ${restrictedOptionItems}[$i]
        $value = $option[$item.Key]
        $shown = if ($item.Key -eq "FileFilter") {
            if ([string]$value) { [string]$value } else { "すべて" }
        } elseif ($value) { "オン" } else { "オフ" }
        $lines += "  $($i + 1). $($item.Label): $shown"
    }
    return , $lines
}

function testIndexChecked {
    # インデックスを検索するか（インデックスのフォルダごと検索から外していなければ検索する）
    param (
        $index,
        [object[]]$excludes
    )

    foreach ($exclude in $excludes) {
        if ($exclude.Subfolders -and $null -ne (getPathUnderFolder $index.Path $exclude.Path)) {
            return $false
        }
    }
    return $true
}

function describeRestrictedIndexes {
    # 検索するインデックスを 1 行で表す（先頭の 3 件まで）
    param (
        [object[]]$indexes,
        [object[]]$excludes
    )

    if ($indexes.Count -eq 0) {
        return "（インデックスがありません。いつもの画面が使える PC で作り、work\index に置いてください）"
    }
    $names = @($indexes | Where-Object { testIndexChecked $_ $excludes } | ForEach-Object { $_.Name })
    if ($names.Count -eq 0) {
        return "（どのインデックスも選んでいません。メニューの［検索するインデックスを選ぶ］で選んでください）"
    }
    $shown = if ($names.Count -gt 3) { ($names[0..2] -join "、") + " ほか $($names.Count - 3) 件" } else { $names -join "、" }
    return "$shown（$($names.Count) / $($indexes.Count) 件）"
}

function getIndexMenuLines {
    # 検索するインデックスを選ぶメニュー
    param (
        [object[]]$indexes,
        [object[]]$excludes
    )

    $lines = @("検索するインデックス（番号で切り替えます。a: すべて選ぶ、空で Enter: 戻る）")
    for ($i = 0; $i -lt $indexes.Count; $i++) {
        $mark = if (testIndexChecked $indexes[$i] $excludes) { "[x]" } else { "[ ]" }
        $source = if ($indexes[$i].SourcePath) { "（元のフォルダ: $($indexes[$i].SourcePath)）" } else { "" }
        $lines += "  $($i + 1). $mark $($indexes[$i].Name)$source"
    }
    return , $lines
}

function switchIndexExclude {
    # インデックスを検索する・しないを切り替えた後の、検索から外すフォルダ（writeSearchExcludes に渡す）を返す。
    # 外すときはインデックスのフォルダ以下すべてを外す。戻すときは、そのインデックスの中の記録をすべて消す
    # （いつもの画面の検索対象ツリーで、インデックスのチェックを付け直したときと同じ）
    param (
        $index,
        [object[]]$excludes
    )

    $inside = { param($exclude) $null -ne (getPathUnderFolder $exclude.Path $index.Path) }
    $others = @($excludes | Where-Object { !(& $inside $_) })
    if (testIndexChecked $index $excludes) {
        return , @($others + @(@{ Path = $index.Path; Subfolders = $true }))
    }
    return , $others
}

function parseMenuNumber {
    # メニューの入力を番号（1 から）にする。番号でなければ 0、範囲の外なら -1
    param (
        [string]$text,
        [int]$count
    )

    $text = $text.Trim()
    if ($text -notmatch '^[0-9０-９]+$') {
        return 0
    }
    # 全角の数字も読む
    $digits = @($text.ToCharArray() | ForEach-Object { if ([int]$_ -ge 0xFF10) { [string][char]([int]$_ - 0xFF10 + 48) } else { [string]$_ } }) -join ""
    $number = [int]$digits
    if ($number -lt 1 -or $number -gt $count) {
        return -1
    }
    return $number
}

function getSearchSummaryLines {
    # 検索の後に出す行（件数・打ち切り・文字どおりに切り替えたこと）
    param (
        [string]$word,
        $result,
        [bool]$useRegex,
        [int]$limit
    )

    $lines = @()
    if ($useRegex -and $result.SimpleMatch) {
        $lines += "「$word」は正規表現として正しくないため、文字どおりに検索しました。"
    }
    if ($result.Hits.Count -eq 0) {
        $lines += "「$word」は見つかりませんでした（TSV $($result.Total) 件を検索）。"
    } elseif ($result.Truncated) {
        $lines += "「$word」は $limit 件を超えたため、$limit 件で打ち切りました（条件を絞ってください）。"
    } else {
        $lines += "「$word」は $($result.Hits.Count) 件見つかりました（TSV $($result.Total) 件を検索）。"
    }
    return , $lines
}

function formatHitListLine {
    # Excel が無い PC で、ヒットを 1 行で出す（番号・ファイル・場所・行・該当行）。width 文字に切り詰める
    param (
        [int]$number,
        $hit,
        [int]$width = 120
    )

    $described = describePlace $hit.Book $hit.Location
    $file = if ($hit.RelDir) { "$($hit.RelDir)\$($hit.Book)" } else { $hit.Book }
    $line = ($hit.Line.Replace("`t", " ").Replace(${cellNewLine}, " ") -replace '[\x00-\x1F]', " ").Trim()
    $text = "{0,4}. {1}  {2}  {3} 行目: {4}" -f $number, $file, ($described.Place -replace '[\x00-\x1F]', " "), $hit.LineNumber, $line
    if ($width -gt 1 -and $text.Length -gt $width) {
        $text = $text.Substring(0, $width - 1) + "…"
    }
    return $text
}

function parseHitChoice {
    # ヒットの一覧の入力を読む: 番号 → @{ Kind = "file"; Number }、f番号 → @{ Kind = "folder"; Number }、読めなければ $null
    param (
        [string]$text,
        [int]$count
    )

    $text = $text.Trim()
    $kind = "file"
    if ($text -match '^[fｆFＦ]\s*(.+)$') {
        $kind = "folder"
        $text = $Matches[1]
    }
    $number = parseMenuNumber $text $count
    if ($number -le 0) {
        return $null
    }
    return @{ Kind = $kind; Number = $number }
}

function getResultInfoItems {
    # 検索結果のブックの［条件］シートに出す項目
    param (
        [string]$word,
        [hashtable]$option,
        [string]$indexes,
        $result,
        [int]$limit,
        [datetime]$now
    )

    return @(
        @("検索ワード", $word),
        @("条件", (describeRestrictedOption $option)),
        @("検索したインデックス", $indexes),
        @("件数", $(if ($result.Truncated) { "$limit 件で打ち切り" } else { "$($result.Hits.Count) 件" })),
        @("検索した TSV", "$($result.Total) 件"),
        @("検索した日時", $now.ToString("yyyy/MM/dd HH:mm:ss")),
        @("見方", "ファイルの行の + で、ファイルごとのヒットを開閉できます。ヒットの行の + で前後の行が見えます。ファイル名のリンクで元のファイル（Excel は該当のシートとセル）を開きます。")
    )
}
