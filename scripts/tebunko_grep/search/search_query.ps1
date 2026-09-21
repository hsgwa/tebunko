# 検索の条件（正規表現・対象ファイル）の組み立て（判断層。ファイルに触らない）。

function isValidRegex {
    # 正規表現として正しいか
    param (
        [string]$pattern
    )

    try {
        [void][regex]::new($pattern)
        return $true
    } catch {
        return $false
    }
}


# 1行の照合にかけてよい時間。正規表現によっては終わらなくなるため、超えたら検索を止める
${regexTimeout} = [timespan]::FromSeconds(5)

function getRegexScanMode {
    # 正規表現を TSV の全文（改行は LF にそろえ、^ $ は行の先頭・末尾に一致させる）にかけてよいかを返す。
    # 1 行ずつ PowerShell で照合すると遅いため、全文への照合（.NET の中で走る）で済ませられる場合を見分ける。
    #   "lines" : 一致は必ず 1 行の中に収まり、行の外の文字を見ない。全文での一致の位置から、ヒットした行が分かる
    #   "filter": 行をまたいで一致することはあるが、1 行で一致するなら全文でも必ず一致する。
    #             全文で一致しない TSV は読み飛ばせる（一致した TSV は 1 行ずつ照合し直す）
    #   "scan"  : 全文では 1 行と結果が変わりうる（\A \z、否定の先読み・後読み、(?i) 等）。1 行ずつ照合する
    # 判定は安全側に倒す（分からない書き方は "filter" か "scan"）。検索結果は照合のしかたによらず同じになる
    param (
        [string]$pattern
    )

    $lines = $true
    $inClass = $false
    $classStart = 0
    $length = $pattern.Length
    for ($i = 0; $i -lt $length; $i++) {
        $c = $pattern[$i]
        if ($c -eq '\') {
            $i++
            if ($i -ge $length) {
                return "scan"
            }
            $e = $pattern[$i]
            if ("AzZG".IndexOf($e) -ge 0) {
                return "scan"   # 文字列の先頭・末尾は、1 行と全文とで位置が違う
            }
            if ("sWDpPnrxuc0123456789".IndexOf($e) -ge 0) {
                $lines = $false  # 改行にも一致しうる（\s \W \D \p{..} \n \x0A \u000A \cJ 8 進数・後方参照）
            }
            if ($inClass -and $i + 1 -lt $length -and $pattern[$i + 1] -eq '-') {
                $lines = $false  # エスケープから始まる範囲（[\t-z] 等）は改行を含みうる
            }
            continue
        }
        if ([int]$c -lt 0x20) {
            $lines = $false      # 制御文字をそのまま含む（改行・タブから始まる範囲など）
        }
        if ($inClass) {
            if ($c -eq ']' -and $i -gt $classStart) {
                $inClass = $false
            }
            continue
        }
        if ($c -eq '[') {
            $inClass = $true
            if ($i + 1 -lt $length -and $pattern[$i + 1] -eq '^') {
                $lines = $false  # 否定の文字クラスは改行に一致する
                $i++
            }
            $classStart = $i + 1  # 先頭の ] は文字として扱われる
            continue
        }
        if ($c -eq '(' -and $i + 1 -lt $length -and $pattern[$i + 1] -eq '?') {
            $rest = $pattern.Substring($i + 2)
            if ($rest.StartsWith(":") -or $rest.StartsWith(">")) {
                continue
            }
            if ($rest.StartsWith("=") -or $rest.StartsWith("<=")) {
                $lines = $false  # 肯定の先読み・後読みは行の外（改行）を見る。全文で一致が増えるだけなので "filter" にはできる
                continue
            }
            if ($rest -match "^(<[A-Za-z_]|'[A-Za-z_])") {
                continue         # 名前付きグループ
            }
            return "scan"        # 否定の先読み・後読み、(?i) 等のオプション、条件分岐、コメント
        }
    }
    if ($lines) {
        return "lines"
    }
    return "filter"
}


function newSearchRegex {
    # 検索条件から、検索と一致箇所の強調で使う正規表現を作る（画面の検索と結果の表示で共通）。
    #   simpleMatch  : $true ならワードを文字どおりに探す。$false なら正規表現として探し、正規表現として不正なら文字どおりに探す
    #   caseSensitive: 英字の大文字・小文字を区別する（既定は区別しない）
    # @{ Regex; SimpleMatch（実際に文字どおり探すか）; TextRegex; ScanMode } を返す。
    #   TextRegex: TSV の全文にかける正規表現（Regex に Multiline を足したもの。^ $ が行の先頭・末尾に一致する）
    #   ScanMode : TextRegex を全文にかけてよいか（getRegexScanMode）
    param (
        [string]$word,
        [bool]$simpleMatch = $true,
        [bool]$caseSensitive = $false
    )

    if (!$simpleMatch -and !(isValidRegex $word)) {
        $simpleMatch = $true
    }
    $pattern = if ($simpleMatch) { [regex]::Escape($word) } else { $word }
    # CultureInvariant: 大文字・小文字を区別しないときの照合が、区別する場合と同程度に速くなる（日本語の照合結果は変わらない）
    $options = [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    if (!$caseSensitive) {
        $options = $options -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    }
    $textOptions = $options -bor [System.Text.RegularExpressions.RegexOptions]::Multiline
    return @{
        Regex       = New-Object System.Text.RegularExpressions.Regex($pattern, $options, ${regexTimeout})
        SimpleMatch = $simpleMatch
        TextRegex   = New-Object System.Text.RegularExpressions.Regex($pattern, $textOptions, ${regexTimeout})
        ScanMode    = getRegexScanMode $pattern
    }
}

function newPlaceExclude {
    # 検索から外す場所（図形・コメント）の正規表現を返す。どちらも検索するなら $null。
    # 場所の名前は "<シート名>[図形]" "<シート名>[コメント]"（index_name.ps1 の objectPlacePattern）
    param (
        [bool]$includeShapes = $true,
        [bool]$includeComments = $true
    )

    $kinds = @()
    if (!$includeShapes) { $kinds += "図形" }
    if (!$includeComments) { $kinds += "コメント" }
    if ($kinds.Count -eq 0) {
        return $null
    }
    return [regex]::new("\[(?:$($kinds -join '|'))\]$")
}

function newFileFilter {
    # 対象ファイルの指定（例: "*.xlsx;見積*;!*old*"）を、元のファイル名に対する正規表現 @{ Include; Exclude } にする（無い側は $null）。
    #   ; で区切る（全角の ； も可）。! で始まるものは除外。* は任意の文字列、? は任意の1文字。大文字・小文字を区別しない
    #   * も ? も無いものは部分一致（「見積」は「*見積*」）
    param (
        [string]$filter
    )

    $include = New-Object System.Collections.Generic.List[string]
    $exclude = New-Object System.Collections.Generic.List[string]
    foreach ($item in ([string]$filter).Split([char[]]";；")) {
        $item = $item.Trim()
        $list = $include
        if ($item.StartsWith("!") -or $item.StartsWith("！")) {
            $list = $exclude
            $item = $item.Substring(1).Trim()
        }
        if ($item -eq "") {
            continue
        }
        if ($item.IndexOfAny([char[]]"*?") -lt 0) {
            $item = "*${item}*"
        }
        $list.Add("^" + [regex]::Escape($item).Replace("\*", ".*").Replace("\?", ".") + "$")
    }

    $options = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    $result = @{ Include = $null; Exclude = $null }
    if ($include.Count -gt 0) {
        $result.Include = New-Object System.Text.RegularExpressions.Regex(($include -join "|"), $options)
    }
    if ($exclude.Count -gt 0) {
        $result.Exclude = New-Object System.Text.RegularExpressions.Regex(($exclude -join "|"), $options)
    }
    return $result
}
