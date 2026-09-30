# テキストファイル（.txt など）の対象の拡張子・文字コードの判定・読み込み。
# 文字コードの判定（バイト列 → 文字コードまたはバイナリ）は、画面にもファイルにも触らない判断層の関数にする。
# ファイルを開いて読む関数（readTextFile）だけが状態層（ファイルに触る）。

# 取り込み対象にするテキストファイルの拡張子（小文字）。足す・減らすと tests\meta\safety.Tests.ps1 が固定の一覧で確かめる
${textExtensions} = @(".txt", ".csv", ".tsv", ".md", ".log", ".json", ".xml")

# 元のファイルがこの大きさ（バイト）を超えたら、読まずに失敗にする
${textFileMaxBytes} = 10MB

# BOM の無い UTF-16 の判定に使う閾値（docs/design/indexer/text.md）。
#   先頭 ${textUtf16SampleBytes} バイト（偶数バイトに切り詰める）を 2 バイトずつの組として数え、
#   偶数の位置・奇数の位置のどちらかの NUL の割合が ${textUtf16NulRatioThreshold} 以上で、
#   もう片方が「多い側 ÷ ${textUtf16NulSkewDivisor}」以下なら UTF-16 と判定する
${textUtf16SampleBytes}      = 64KB
${textUtf16NulRatioThreshold} = 0.2
${textUtf16NulSkewDivisor}    = 10

# NUL の無い BOM 無し UTF-16（日本語だけの文章）の判定に使う閾値（docs/design/indexing/text.md）。
#   先頭 ${textUtf16SampleBytes} バイトを UTF-16 として読み、ひらがな・カタカナが ${textJpUtf16KanaRatio} 以上あり、
#   日本語の文章に出る文字（ASCII・句読点・かな・漢字・全角）以外が ${textJpUtf16OtherRatio} 以下なら UTF-16 と判定する
${textJpUtf16MinBytes}   = 8
${textJpUtf16KanaRatio}  = 0.2
${textJpUtf16OtherRatio} = 0.05

# 7 ビットだけ（ASCII の範囲）のバイト列は、ひらがなの UTF-16 と ASCII の文字列の見分けが付きにくいため、
# UTF-16 とみなすかなの割合をもっと高く求める
${textJpUtf16AsciiKanaRatio} = 0.9

# Shift_JIS として読めたとき、半角カナが「ASCII 以外の文字」のこの割合を超えると、半角カナの並びとして自然かを確かめる
# （濁点・半濁点が付けられるカナの後にしか来ないか。GBK・Big5 などを Shift_JIS として読むと崩れる）。
# EUC-JP としても読めるときは、この割合を超える側を EUC-JP とする
${textLegacyHalfKanaRatio} = 0.3

# 半角カナの数えは、読んだ結果の先頭からこの文字数だけ見る（大きなファイルで遅くならないため）
${textLegacySampleChars} = 32768

function testTextExtension {
    # ファイル名（またはパス）の拡張子が、取り込み対象のテキストの拡張子か（大文字・小文字を区別しない）
    param (
        [string]$path
    )

    $extension = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
    return (${textExtensions} -contains $extension)
}

function getStrictUtf8Encoding {
    # 不正なバイト列（UTF-8として読めない並び）で例外にする UTF-8 のデコーダー（BOM は付けない）
    return [System.Text.Encoding]::GetEncoding("utf-8",
        [System.Text.EncoderExceptionFallback]::new(), [System.Text.DecoderExceptionFallback]::new())
}

function testTextPlausible {
    # 読んだ結果の文字列が、テキストとして自然か。制御文字（TAB・LF・FF・CR 以外）・私用領域・置き換え文字（U+FFFD）を含めば偽。
    # バイナリを Shift_JIS・EUC-JP などとして読んでしまったものを見分ける（Shift_JIS の外字（F040〜F9FC）は私用領域に読まれるため偽になる）
    param (
        [string]$text
    )

    return (-not [regex]::IsMatch($text, '[\x00-\x08\x0B\x0E-\x1F\x7F-\x9F\uE000-\uF8FF\uFFFD]'))
}

function tryDecodeStrict {
    # コードページで読めないバイト列があれば $null、読めれば文字列
    param (
        [byte[]]$bytes,
        [int]$codePage
    )

    try {
        $encoding = [System.Text.Encoding]::GetEncoding($codePage,
            [System.Text.EncoderExceptionFallback]::new(), [System.Text.DecoderExceptionFallback]::new())
        return $encoding.GetString($bytes)
    } catch [System.Text.DecoderFallbackException] {
        return $null
    } catch [System.ArgumentException] {
        return $null
    }
}

function detectJapaneseUtf16WithoutNul {
    # NUL の無い（または NUL が改行の分しか無い）BOM 無し UTF-16（日本語だけの文章）の判定。
    # ひらがな・カタカナ・漢字は NUL を含まない。LE・BE の片方だけが日本語の文章として自然なときにその文字コードを返す。
    # 両方・どちらも当たらなければ $null。7 ビットだけのバイト列は、かなの割合をもっと高く求める（${textJpUtf16AsciiKanaRatio}）
    param (
        [byte[]]$bytes
    )

    $sampleLength = [Math]::Min($bytes.Length, [int]${textUtf16SampleBytes})
    $sampleLength -= ($sampleLength % 2)
    if ($sampleLength -lt ${textJpUtf16MinBytes}) { return $null }

    $ascii = $true
    for ($i = 0; $i -lt $sampleLength; $i++) {
        if ($bytes[$i] -ge 0x80) { $ascii = $false; break }
    }
    $kanaRatio = if ($ascii) { ${textJpUtf16AsciiKanaRatio} } else { ${textJpUtf16KanaRatio} }

    $found = @()
    foreach ($candidate in @(@("UTF16LE", [System.Text.Encoding]::Unicode), @("UTF16BE", [System.Text.Encoding]::BigEndianUnicode))) {
        $text = $candidate[1].GetString($bytes, 0, $sampleLength)
        $kana = [regex]::Matches($text, '[\u3040-\u30FF]').Count
        $other = [regex]::Matches($text, '[^\u0009\u000A\u000D -~\u3000-\u30FF\u4E00-\u9FFF\uFF00-\uFFEF]').Count
        if ($kana -ge ($text.Length * $kanaRatio) -and $other -le ($text.Length * ${textJpUtf16OtherRatio})) {
            $found += $candidate[0]
        }
    }
    if ($found.Count -eq 1) { return $found[0] }
    return $null
}

function getIso2022JpVerdict {
    # ISO-2022-JP（JIS）かの判定。戻り値: "yes" / "broken" / "none"
    #   yes: 7 ビットだけで、漢字への切り替え（ESC $ B・ESC $ @）を含み、ESC の並びがすべて規格のもの（ESC ( B・ESC ( J・ESC ( I を含む）
    #   broken: 漢字への切り替えを含むが、8 ビットのバイトや規格外の ESC の並びがある（JIS のつもりの壊れたもの。UTF-8 として読むと化けるので取り込まない）
    #   none: ESC が無い、または漢字への切り替えが無い（ESC ( B だけ・ESC [ などの端末の制御の並び）
    param (
        [byte[]]$bytes
    )

    if ([Array]::IndexOf($bytes, [byte]0x1B) -lt 0) { return "none" }
    $hasKanjiShift = $false
    $broken = $false
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -ge 0x80) { $broken = $true; continue }
        if ($bytes[$i] -ne 0x1B) { continue }
        if ($i + 2 -ge $bytes.Length) { $broken = $true; continue }
        $seq = [string][char]$bytes[$i + 1] + [string][char]$bytes[$i + 2]
        if ($seq -eq '$B' -or $seq -eq '$@') {
            $hasKanjiShift = $true
        } elseif ($seq -ne '(B' -and $seq -ne '(J' -and $seq -ne '(I') {
            $broken = $true
        }
    }
    if (-not $hasKanjiShift) { return "none" }
    if ($broken) { return "broken" }
    return "yes"
}

function testHalfWidthKanaNatural {
    # 半角カナを多く含む文字列が、半角カナの並びとして自然か。濁点（ﾞ）・半濁点（ﾟ）は、付けられるカナ
    # （ｦ ｳ ｶ〜ﾄ ﾊ〜ﾎ ﾜ。半濁点は ﾊ〜ﾎ）の直後にしか来ない。GBK・Big5 などを Shift_JIS として読むと崩れる
    param (
        [string]$text
    )

    if ([regex]::IsMatch($text, '(?<![\uFF66\uFF73\uFF76-\uFF84\uFF8A-\uFF8E\uFF9C])\uFF9E')) { return $false }
    if ([regex]::IsMatch($text, '(?<![\uFF8A-\uFF8E])\uFF9F')) { return $false }
    return $true
}

function detectLegacyJapaneseEncoding {
    # UTF-8 として読めなかった NUL の無いバイト列が、Shift_JIS か EUC-JP か。決められなければ $null（取り込まない）。
    #   Shift_JIS: コードページ 932 として不正なバイト列が無く、制御文字などが出ない。半角カナが多い（${textLegacyHalfKanaRatio} 超）ときは、
    #              半角カナの並びとして自然（testHalfWidthKanaNatural）なものだけ
    #   EUC-JP: コードページ 51932 として不正なバイト列が無く、制御文字などが出ず、ひらがな・カタカナを含む
    #           （かなの無い漢字だけの文章は、GBK・EUC-KR と構造が同じで見分けが付かない）
    #   両方に当たるときは、Shift_JIS として読むと半角カナばかりになる（EUC-JP のかなの見え方）なら EUC-JP、そうでなければあいまいなので $null
    param (
        [byte[]]$bytes
    )

    $sjis = tryDecodeStrict $bytes 932
    $sjisOk = $false
    $sjisHalfKanaHeavy = $false
    if ($null -ne $sjis -and (testTextPlausible $sjis)) {
        $sample = if ($sjis.Length -gt ${textLegacySampleChars}) { $sjis.Substring(0, [int]${textLegacySampleChars}) } else { $sjis }
        $nonAscii = [regex]::Matches($sample, '[^\u0000-\u007F]').Count
        $halfKana = [regex]::Matches($sample, '[\uFF61-\uFF9F]').Count
        $sjisHalfKanaHeavy = ($halfKana -gt ($nonAscii * ${textLegacyHalfKanaRatio}))
        $sjisOk = (-not $sjisHalfKanaHeavy) -or (testHalfWidthKanaNatural $sample)
    }

    $euc = tryDecodeStrict $bytes 51932
    $eucOk = ($null -ne $euc -and (testTextPlausible $euc) -and [regex]::IsMatch($euc, '[\u3040-\u30FF]'))

    if ($sjisOk -and -not $eucOk) { return "ShiftJIS" }
    if ($eucOk -and -not $sjisOk) { return "EUCJP" }
    if ($eucOk -and $sjisOk -and $sjisHalfKanaHeavy) { return "EUCJP" }
    return $null
}

function detectTextEncoding {
    # バイト列だけから文字コードを判定する（ファイルに触らない）。
    # 順に確かめる: BOM → NUL の偏り（BOM の無い UTF-16）→ NUL の無い日本語の UTF-16 → ISO-2022-JP →
    #   UTF-8 として正しく読めるか → Shift_JIS・EUC-JP。判定できないもの（バイナリ・UTF-32・GBK など）は $null を返す
    #   戻り値: "UTF8" / "UTF16LE" / "UTF16BE" / "ShiftJIS" / "EUCJP" / "ISO2022JP" / $null（取り込まない）
    param (
        [byte[]]$bytes
    )

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return "UTF8"
    }
    # UTF-32 の BOM は UTF-16LE の BOM（FF FE）と同じ並びで始まるため、先に確かめて取り込まない
    if ($bytes.Length -ge 4 -and (
            ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE -and $bytes[2] -eq 0 -and $bytes[3] -eq 0) -or
            ($bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 0xFE -and $bytes[3] -eq 0xFF))) {
        return $null
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return "UTF16LE"
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return "UTF16BE"
    }

    $sampleLength = [Math]::Min($bytes.Length, [int]${textUtf16SampleBytes})
    $sampleLength -= ($sampleLength % 2)
    if ($sampleLength -gt 0) {
        $even = 0  # 偶数の位置（0, 2, …）の NUL の数
        $odd = 0   # 奇数の位置（1, 3, …）の NUL の数
        for ($i = 0; $i -lt $sampleLength; $i++) {
            if ($bytes[$i] -eq 0) {
                if (($i % 2) -eq 0) { $even++ } else { $odd++ }
            }
        }
        $n = $sampleLength / 2
        if ($odd -ge ($n * ${textUtf16NulRatioThreshold}) -and $even -le ($odd / ${textUtf16NulSkewDivisor})) {
            return "UTF16LE"
        }
        if ($even -ge ($n * ${textUtf16NulRatioThreshold}) -and $odd -le ($even / ${textUtf16NulSkewDivisor})) {
            return "UTF16BE"
        }
        if ($even -gt 0 -or $odd -gt 0) {
            # NUL が 20% に届かない。片側にしか無ければ、日本語の UTF-16 の改行（CRLF）だけが NUL のものかを確かめる。
            # それ以外（両側にある・かなが足りない）はバイナリ（UTF-32 もここ）
            if ($even -eq 0 -or $odd -eq 0) {
                return (detectJapaneseUtf16WithoutNul $bytes)
            }
            return $null
        }
    }

    $utf16 = detectJapaneseUtf16WithoutNul $bytes
    if ($null -ne $utf16) { return $utf16 }

    $iso = getIso2022JpVerdict $bytes
    if ($iso -eq "yes") { return "ISO2022JP" }
    if ($iso -eq "broken") { return $null }

    if ($null -ne (tryDecodeStrict $bytes 65001)) {
        return "UTF8"
    }
    return (detectLegacyJapaneseEncoding $bytes)
}

function decodeTextBytes {
    # detectTextEncoding が返した文字コードで、バイト列を文字列にする（BOM は取り除く）
    param (
        [byte[]]$bytes,
        [string]$encodingName
    )

    switch ($encodingName) {
        "UTF8" {
            $offset = if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { 3 } else { 0 }
            return (getStrictUtf8Encoding).GetString($bytes, $offset, $bytes.Length - $offset)
        }
        "UTF16LE" {
            $offset = if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { 2 } else { 0 }
            return [System.Text.Encoding]::Unicode.GetString($bytes, $offset, $bytes.Length - $offset)
        }
        "UTF16BE" {
            $offset = if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { 2 } else { 0 }
            return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, $offset, $bytes.Length - $offset)
        }
        "ShiftJIS" {
            return [System.Text.Encoding]::GetEncoding(932).GetString($bytes)
        }
        "EUCJP" {
            return [System.Text.Encoding]::GetEncoding(51932).GetString($bytes)
        }
        "ISO2022JP" {
            return [System.Text.Encoding]::GetEncoding(50221).GetString($bytes)
        }
    }
    throw "不明な文字コードです（${encodingName}）。"
}

function splitTextLines {
    # 文字列を行に分ける（CRLF・LF・CR のどれでも1行。StreamReader.ReadLine と同じ分け方）。
    # 途中の空の行は残し、行末の空白は取り除き（TrimEnd）、末尾の空の行は捨てる
    param (
        [string]$text
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $reader = New-Object System.IO.StringReader($text)
    try {
        $line = $reader.ReadLine()
        while ($null -ne $line) {
            $lines.Add($line.TrimEnd())
            $line = $reader.ReadLine()
        }
    } finally {
        $reader.Dispose()
    }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq "") {
        $lines.RemoveAt($lines.Count - 1)
    }
    return , $lines.ToArray()
}

function readTextFile {
    # テキストファイルを読み、行の並び（splitTextLines と同じ形）にして返す。
    # 元のファイルは読み取りだけで開く（共有は copyFileShared と同じ ReadWrite|Delete。コピーは作らない）。
    # 大きさの上限を超える・バイナリと判定したときは、取り込み一覧・ログにそのまま出す文言で例外にする。
    # maxBytes はテストで上限を小さく差し替えるための引数（既定は textFileMaxBytes）
    param (
        [string]$path,
        [long]$maxBytes = ${textFileMaxBytes}
    )

    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = [System.IO.FileStream]::new((toLongPath $path), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        if ($stream.Length -gt $maxBytes) {
            throw "ファイルサイズが大きすぎるため取り込めません。"
        }
        $bytes = New-Object byte[] ([int]$stream.Length)
        $read = 0
        while ($read -lt $bytes.Length) {
            $n = $stream.Read($bytes, $read, $bytes.Length - $read)
            if ($n -eq 0) { break }
            $read += $n
        }
    } finally {
        $stream.Dispose()
    }

    $encodingName = detectTextEncoding $bytes
    if ($null -eq $encodingName) {
        throw "テキストファイルではないため取り込めません。"
    }
    return (splitTextLines (decodeTextBytes $bytes $encodingName))
}
