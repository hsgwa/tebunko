# 検索用の索引（gram 索引）に使う 2-gram の決まり（判断層。ファイルにも画面にも触らない）。
# 本文の隣り合う 2 文字を値（uint32）にして索引に入れる。ワードを含む本文には、ワードのすべての 2-gram が必ず入っているので、
# 検索語の 2-gram を含まないファイルは、照合しなくてよい。

# 1 回の検索に使う 2-gram の数の上限（多いときは均等に間引く。間引いても候補が増えるだけで漏れない）
${searchGramMax} = 16

function getSearchGramValues {
    # 検索語から、候補の絞り込みに使う 2-gram の値（uint32。addTextGrams と同じ値）を返す（重複なし・最大 searchGramMax 個）。
    # 空白で区切り、2 文字以上の部分の隣り合う 2 文字を値にする。値が無ければ空（絞り込みは使えない）
    param (
        [string]$word
    )

    $grams = New-Object System.Collections.Generic.List[uint32]
    $seen = New-Object System.Collections.Generic.HashSet[uint32]
    foreach ($part in ([string]$word).ToLowerInvariant() -split "[\s\u2028]+") {
        for ($i = 0; $i -lt $part.Length - 1; $i++) {
            $value = [uint32]([long][int]$part[$i] + 65536L * [long][int]$part[$i + 1])
            if ($seen.Add($value)) {
                $grams.Add($value)
            }
        }
    }
    if ($grams.Count -le ${searchGramMax}) {
        return , $grams.ToArray()
    }
    # 均等に間引く（先頭と末尾は残す）
    $picked = New-Object System.Collections.Generic.List[uint32]
    for ($k = 0; $k -lt ${searchGramMax}; $k++) {
        $picked.Add($grams[[int][Math]::Round($k * ($grams.Count - 1) / (${searchGramMax} - 1))])
    }
    return , $picked.ToArray()
}

function addTextGrams {
    # 本文の 2-gram を、語の値（uint32。UTF-16LE の 2 文字をそのまま 4 バイトの整数にしたもの）として set に足す。
    # 1 文字ずつ PowerShell で回すと遅いため、char[] を uint32[] に写して .NET の HashSet で重複を除く。
    # 空白を含む 2-gram も入るが、検索には使わないため害は無い（除く方が遅い）
    param (
        [System.Collections.Generic.HashSet[uint32]]$set,
        [string]$text
    )

    $chars = $text.ToLowerInvariant().ToCharArray()
    $even = $chars.Length -shr 1
    $odd = ($chars.Length - 1) -shr 1
    if ($even -gt 0) {
        $u0 = New-Object 'uint32[]' $even
        [System.Buffer]::BlockCopy($chars, 0, $u0, 0, $even * 4)
        $set.UnionWith($u0)
    }
    if ($odd -gt 0) {
        # 1 文字ずらした並びは、2 バイトずらして写せば作れる
        $u1 = New-Object 'uint32[]' $odd
        [System.Buffer]::BlockCopy($chars, 2, $u1, 0, $odd * 4)
        $set.UnionWith($u1)
    }
}

function skipRegexClass {
    # 正規表現の文字クラス（[ の位置 start から）を読み飛ばし、閉じ ] の次の位置を返す。閉じていなければ -1。
    # 文字クラスの引き算（[a-z-[aeiou]]）の入れ子は数で数える。\ の次の 1 文字は読み飛ばす
    param (
        [string]$pattern,
        [int]$start
    )

    $i = $start + 1
    $length = $pattern.Length
    if ($i -lt $length -and $pattern[$i] -eq '^') { $i++ }
    if ($i -lt $length -and $pattern[$i] -eq ']') { $i++ }
    $depth = 1
    while ($i -lt $length) {
        $c = $pattern[$i]
        if ($c -eq '\') {
            $i += 2
            continue
        }
        if ($c -eq '[') {
            # 入れ子になるのは引き算（-[）のときだけ。それ以外の [ は文字
            if ($pattern[$i - 1] -eq '-') { $depth++ }
        } elseif ($c -eq ']') {
            $depth--
            if ($depth -eq 0) { return $i + 1 }
        }
        $i++
    }
    return -1
}

function skipRegexGroup {
    # 正規表現の ( の位置 start から、対応する ) の次の位置を返す。閉じていない・( ? の後ろが : 以外（先読み・名前付き・インラインの指定など）が
    # 中にある・読めないときは -1（呼び出し側はあきらめる）。中は見ない。入れ子は数で数える（再帰しない）
    param (
        [string]$pattern,
        [int]$start
    )

    $length = $pattern.Length
    $i = $start
    $depth = 0
    while ($i -lt $length) {
        $c = $pattern[$i]
        if ($c -eq '\') {
            $i += 2
            continue
        }
        if ($c -eq '[') {
            $i = skipRegexClass $pattern $i
            if ($i -lt 0) { return -1 }
            continue
        }
        if ($c -eq '(') {
            if ($i + 1 -lt $length -and $pattern[$i + 1] -eq '?' -and !($i + 2 -lt $length -and $pattern[$i + 2] -eq ':')) { return -1 }
            $depth++
        } elseif ($c -eq ')') {
            $depth--
            if ($depth -eq 0) { return $i + 1 }
        }
        $i++
    }
    return -1
}

function skipRegexQuantifier {
    # 正規表現の { の位置 start が、量指定（{n}・{n,}・{n,m}）なら } の次の位置を返す。量指定の形でなければ -1
    # （.NET では量指定でない { はただの文字で、中に | や ( が来うる。呼び出し側は読むのをあきらめる）
    param (
        [string]$text,
        [int]$start
    )

    $length = $text.Length
    $i = $start + 1
    $digits = $i
    while ($i -lt $length -and $text[$i] -ge '0' -and $text[$i] -le '9') { $i++ }
    if ($i -eq $digits) { return -1 }
    if ($i -lt $length -and $text[$i] -eq ',') {
        $i++
        while ($i -lt $length -and $text[$i] -ge '0' -and $text[$i] -le '9') { $i++ }
    }
    if ($i -lt $length -and $text[$i] -eq '}') { return $i + 1 }
    return -1
}

function readRegexEscape {
    # 正規表現の \ の位置 start の逃がしを読み、@{ Next; Literal } を返す。Next は次の位置（読めなければ -1）。
    # Literal は英数字でない文字の逃がし（\. など）の文字で、英数字の逃がし（\d \x41 \k<a> など）は $null（区切る。デコードしない）。
    # 引数を持つ逃がしは、引数まで読み飛ばす（引数を文字として拾わない）
    param (
        [string]$text,
        [int]$start
    )

    $length = $text.Length
    if ($start + 1 -ge $length) {
        return @{ Next = -1; Literal = $null }
    }
    $e = $text[$start + 1]
    $i = $start + 2
    if (!([char]::IsLetterOrDigit($e) -or $e -ceq '_')) {
        return @{ Next = $i; Literal = $e }
    }
    if ($e -ceq 'x') {
        $i += 2
    } elseif ($e -ceq 'u') {
        $i += 4
    } elseif ($e -ceq 'c') {
        $i += 1
    } elseif ($e -ceq 'k') {
        if ($i -ge $length -or ($text[$i] -ne '<' -and $text[$i] -ne "'")) {
            return @{ Next = -1; Literal = $null }
        }
        $close = $text.IndexOf($(if ($text[$i] -eq '<') { '>' } else { "'" }), $i + 1)
        if ($close -lt 0) {
            return @{ Next = -1; Literal = $null }
        }
        $i = $close + 1
    } elseif ($e -ceq 'p' -or $e -ceq 'P') {
        # \p{L} の {...} まで読み飛ばす
        $close = if ($i -lt $length -and $text[$i] -eq '{') { $text.IndexOf('}', $i + 1) } else { -1 }
        if ($close -lt 0) {
            return @{ Next = -1; Literal = $null }
        }
        $i = $close + 1
    } elseif ([char]::IsDigit($e)) {
        while ($i -lt $length -and [char]::IsDigit($text[$i])) { $i++ }
    }
    return @{ Next = $i; Literal = $null }
}

function getRegexRequiredText {
    # 正規表現から、一致するなら必ず含まれる文字の並びを、空白区切りの文字列で返す（getSearchGramValues にそのまま渡せる）。取り出せなければ空文字。
    # 式を前から 1 回だけ読む（正規表現を使わない・再帰しない）。いちばん外側の連接だけを見て、迷う書き方は全部あきらめる（空文字）。
    #   取り出す: ふつうの文字・英数字でない文字の逃がし（\. など）
    #   区切る: . ^ $ [...] (...) (?:...) 英数字の逃がし（\d \x41 \k<a> など。デコードしない）・空白
    #   直前の 1 文字を捨てて区切る: ? * {n} {n,} {n,m}（量指定の形でない { はあきらめる）。+ は直前の 1 文字を残して区切る。量指定の後ろの ? + は読み飛ばす
    #   あきらめる: いちばん外側の |・量指定の形でない {・(? の後ろが : 以外・閉じていない・閉じすぎ・末尾の \
    param (
        [string]$pattern
    )

    $text = [string]$pattern
    $length = $text.Length
    $parts = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Text.StringBuilder
    $special = [char[]]@('(', ')', '[', '\')
    $i = 0
    while ($i -lt $length) {
        $c = $text[$i]
        $flush = $false
        $quantifier = $false
        if ($c -ceq '\') {
            $escape = readRegexEscape $text $i
            if ($escape.Next -lt 0) { return "" }
            $i = $escape.Next
            if ($escape.Literal) { [void]$current.Append($escape.Literal) } else { $flush = $true }
        } elseif ($c -ceq '(') {
            $next = if ($i + 1 -lt $length -and $text[$i + 1] -ceq '?') { -1 } else { $text.IndexOfAny($special, $i + 1) }
            if ($next -ge 0 -and $text[$next] -ceq ')') {
                # 中に ( [ \ が無い単純な組は、そのまま閉じまで進める
                $i = $next + 1
            } else {
                $i = skipRegexGroup $text $i
                if ($i -lt 0) { return "" }
            }
            $flush = $true
        } elseif ($c -ceq '[') {
            $i = skipRegexClass $text $i
            if ($i -lt 0) { return "" }
            $flush = $true
        } elseif ($c -ceq '|' -or $c -ceq ')') {
            return ""
        } elseif ($c -ceq '.' -or $c -ceq '^' -or $c -ceq '$') {
            $i++
            $flush = $true
        } elseif ($c -ceq '?' -or $c -ceq '*') {
            $i++
            if ($current.Length -gt 0) { $current.Length-- }
            $flush = $true
            $quantifier = $true
        } elseif ($c -ceq '+') {
            $i++
            $flush = $true
            $quantifier = $true
        } elseif ($c -ceq '{') {
            $i = skipRegexQuantifier $text $i
            if ($i -lt 0) { return "" }
            if ($current.Length -gt 0) { $current.Length-- }
            $flush = $true
            $quantifier = $true
        } else {
            $i++
            if ([char]::IsWhiteSpace($c) -or [int]$c -eq 0x2028) { $flush = $true } else { [void]$current.Append($c) }
        }
        if ($quantifier -and $i -lt $length -and ($text[$i] -ceq '?' -or $text[$i] -ceq '+')) { $i++ }
        if ($flush -and $current.Length -gt 0) {
            $parts.Add($current.ToString())
            [void]$current.Clear()
        }
    }
    if ($current.Length -gt 0) { $parts.Add($current.ToString()) }
    return ($parts -join " ")
}

function getGramSearchText {
    # gram 索引で絞るときに、2-gram を作る元の文字列を返す。照合（newSearchRegex）が文字どおりに探すとき
    # （simpleMatch、または正規表現として不正な式）は検索語そのまま、正規表現のときは必ず含まれる文字（getRegexRequiredText）
    param (
        [string]$word,
        [bool]$simpleMatch
    )

    if ($simpleMatch -or !(isValidRegex $word)) {
        return $word
    }
    return (getRegexRequiredText $word)
}
