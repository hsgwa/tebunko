# 起動と終了の速さの比べ方（tools\perf\measure_startup.ps1 から読む。プロセスを起こさない判定の部分）。
# zip 版（scripts\ をそのまま使う形）と単一 .ps1 版を同じ手順で測った値から、捨てる回・中央値・倍率・合否を出す。
# 設計は docs/design/structure/single-script.md「速さの測り方」。

. "$PSScriptRoot\perf_common.ps1"

# 単一 .ps1 版の中央値が zip 版の中央値の何倍以内なら合格とするか
$script:startupRatioLimit = 1.2

# 1 つの値の並び（測った順。ミリ秒）から、先頭の warmup 回を捨てた中央値を返す。捨てたあとに値が無ければ $null
function getStartupMedian {
    param (
        [double[]]$values,
        [int]$warmup = 1
    )

    $kept = @($values | Select-Object -Skip $warmup)
    $stats = getStats ([double[]]$kept)
    if ($null -eq $stats) { return $null }
    return $stats.Median
}

# zip 版と単一 .ps1 版の値（測った順。ミリ秒）を比べる。倍率は 単一 / zip の中央値。
# 中央値が出せない（捨てたあとに値が無い）・zip 版の中央値が 0 のときは、倍率も合否も $null（合格とは言わない）
function compareStartup {
    param (
        [double[]]$zipValues,
        [double[]]$singleValues,
        [int]$warmup = 1,
        [double]$limit = $script:startupRatioLimit
    )

    $zip = getStartupMedian $zipValues $warmup
    $single = getStartupMedian $singleValues $warmup
    $ratio = $null
    $pass = $null
    if ($null -ne $zip -and $null -ne $single -and $zip -gt 0) {
        $ratio = $single / $zip
        $pass = ($ratio -le $limit)
    }
    return [ordered]@{ ZipMedian = $zip; SingleMedian = $single; Ratio = $ratio; Limit = $limit; Pass = $pass }
}

# 終了コードが 0 でない回数（閉じたときの終了コードの揺れを、捨てずに数える）。$null（終わらなかった）も数える
function countNonZeroExit {
    param (
        $exitCodes
    )

    return @($exitCodes | Where-Object { $null -eq $_ -or $_ -ne 0 }).Count
}
