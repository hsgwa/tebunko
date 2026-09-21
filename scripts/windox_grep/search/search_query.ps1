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


function newSearchRegex {
    # 検索条件から、検索と一致箇所の強調で使う正規表現を作る（画面の検索と結果の表示で共通）。
    #   simpleMatch  : $true ならワードを文字どおりに探す。$false なら正規表現として探し、正規表現として不正なら文字どおりに探す
    #   caseSensitive: 英字の大文字・小文字を区別する（既定は区別しない）
    # @{ Regex; SimpleMatch（実際に文字どおり探すか） } を返す
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
    return @{
        Regex       = New-Object System.Text.RegularExpressions.Regex($pattern, $options, ${regexTimeout})
        SimpleMatch = $simpleMatch
    }
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
