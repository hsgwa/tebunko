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