# PSScriptAnalyzer の指摘を、ベースライン（tests\meta\PSScriptAnalyzer.baseline.tsv）と比べる関数。
# tools\check_analyzer.ps1 が読み込む。PSScriptAnalyzer を使わずに単体テストできるよう、比べる処理だけをここに分ける。
#
# ベースラインは「ファイル<TAB>規則<TAB>件数」の行を、ファイル・規則の順（序数比較）に並べたもの。
# ファイルはリポジトリの直下からの相対パスで、区切りは /。行番号は持たない（ほかの行を直すたびにずれるため）。

# 指摘の一覧から、数えない指摘を除く。除くのは次の 2 つ。
#   - TypeNotFound: 規則でなく構文解析の指摘なので ExcludeRules では除けない。継承元の型が別ファイルにあり 1 ファイルだけでは解決できないもの（読み込む順で解決することは tests\meta\structure.Tests.ps1 の「型の読み込み」が確かめる）
#   - tests\testdata\text\ の見本: .ps1・.psm1・.psd1 の中身は文章で、チェックアウトの改行などで指摘が変わるため
function selectAnalyzerFindings($findings) {
    $kept = New-Object System.Collections.Generic.List[object]
    foreach ($f in $findings) {
        if ($f.RuleName -eq "TypeNotFound") { continue }
        if ($f.ScriptPath.Replace("/", "\").ToLowerInvariant().Contains("\tests\testdata\text\")) { continue }
        $kept.Add($f)
    }
    return $kept.ToArray()
}

# 指摘の一覧を「ファイル<TAB>規則」ごとの件数と行の一覧にまとめる。
# 指摘は ScriptPath・RuleName・Line・Message を持つ。戻り値のキーは「ファイル<TAB>規則」
function getAnalyzerCounts($findings, [string]$rootDir) {
    $prefix = [System.IO.Path]::GetFullPath($rootDir).TrimEnd("\", "/") + "\"
    $counts = @{}
    # List[object] を @() に渡さない（5.1 で型が合わずに落ちる）。配列か null を受け取る
    foreach ($f in $findings) {
        $full = [System.IO.Path]::GetFullPath($f.ScriptPath)
        $name = if ($full.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { $full.Substring($prefix.Length) } else { $full }
        $name = $name.Replace("\", "/")
        $key = "$name`t$($f.RuleName)"
        if (!$counts.ContainsKey($key)) { $counts[$key] = New-Object System.Collections.Generic.List[string] }
        $counts[$key].Add("$($f.RuleName) ${name}:$($f.Line) $($f.Message)")
    }
    return $counts
}

# ベースラインの本文を読む。戻り値のキーは「ファイル<TAB>規則」、値は件数
function readAnalyzerBaseline([string]$text) {
    $baseline = @{}
    foreach ($line in ($text -split "\r?\n")) {
        if (!$line) { continue }
        $cols = $line.Split("`t")
        if ($cols.Count -ne 3) { throw "ベースラインの行が 3 列ではありません: $line" }
        $baseline["$($cols[0])`t$($cols[1])"] = [int]$cols[2]
    }
    return $baseline
}

# 件数（getAnalyzerCounts の戻り値）からベースラインの本文を作る（並びは決まっている）
function formatAnalyzerBaseline($counts) {
    $keys = @($counts.Keys)
    [array]::Sort($keys, [System.StringComparer]::Ordinal)
    $lines = foreach ($key in $keys) { "$key`t$($counts[$key].Count)" }
    if (!$lines) { return "" }
    return (@($lines) -join "`r`n") + "`r`n"
}

# ベースラインと今の件数を比べる。増えた（ベースラインに無いものを含む）指摘は 1 行ずつ、減ったものは書き直す行を返す。
# 戻り値は @{ Increased = 行の配列; Decreased = 行の配列 }
function compareAnalyzerBaseline($baseline, $counts) {
    $increased = New-Object System.Collections.Generic.List[string]
    $decreased = New-Object System.Collections.Generic.List[string]
    $keys = @($counts.Keys) + @($baseline.Keys | Where-Object { !$counts.ContainsKey($_) })
    # 大文字小文字を区別して重複を除く（formatAnalyzerBaseline の並びと同じ序数比較）
    $set = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::Ordinal)
    foreach ($k in $keys) { [void]$set.Add($k) }
    $keys = @($set)
    foreach ($key in $keys) {
        $was = if ($baseline.ContainsKey($key)) { [int]$baseline[$key] } else { 0 }
        $now = if ($counts.ContainsKey($key)) { $counts[$key].Count } else { 0 }
        if ($now -gt $was) {
            # どの指摘が新しいかは分からないので、そのファイル・規則の指摘を全部出す
            foreach ($line in $counts[$key]) { $increased.Add("$line （ベースライン $was 件 → $now 件）") }
        } elseif ($now -lt $was) {
            $cols = $key.Split("`t")
            $decreased.Add("$($cols[1]) $($cols[0]) : ベースライン $was 件 → $now 件。-UpdateBaseline で書き直す")
        }
    }
    return @{ Increased = @($increased); Decreased = @($decreased) }
}
