# 検索対象のツリーの判断（名前の絞り込み）。画面に触らないため、そのままテストできる（tests\tebunko\ui\search\target_tree_view.Tests.ps1）。

function matchesTreeFilter {
    # ツリーの一番上の項目（インデックス）の名前が、絞り込みの文字に合うか。
    # 空（空白だけ）なら、すべて合う。大文字・小文字は区別せず、空白で区切った語は、すべてを含むものを合うとする
    param (
        [string]$name,
        [string]$filter
    )

    foreach ($word in @($filter -split "[\s　]+" | Where-Object { $_ -ne "" })) {
        if ($name.IndexOf($word, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
            return $false
        }
    }
    return $true
}
