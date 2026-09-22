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

function formatLocationLabel {
    # 結果の「場所」（TSV の名前）を、まとめ表示の見出しに出す表記にする。
    # Excel は「シート 4月」、Word は「3 ページ」、PowerPoint は「スライド 7」。それ以外（ヘッダー・フッターなど）はそのまま
    param (
        [string]$book,
        [string]$location
    )

    if ((getAppKind $book) -eq "Excel") {
        return "シート $location"
    }
    if ($location -match '^ページ0*(\d+)(.*)$') {
        return "$($Matches[1]) ページ$($Matches[2])"
    }
    if ($location -match '^スライド0*(\d+)(.*)$') {
        # 発表者ノート（スライド003_ノート）は「スライド 3 ノート」とする
        return "スライド $($Matches[1])$($Matches[2] -replace '^_', ' ')"
    }
    return $location
}

function describeFileLocations {
    # ファイルの中でヒットした場所（見つかった順・重複なし）を、見出しの右端に出す文字列にする。
    # 1 か所ならその場所、2 か所以上なら「シート 4月 ほか 2 か所」
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