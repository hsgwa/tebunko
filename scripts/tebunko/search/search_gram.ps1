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

