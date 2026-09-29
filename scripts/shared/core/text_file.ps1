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

function detectTextEncoding {
    # バイト列だけから文字コードを判定する（ファイルに触らない）。
    # 順に確かめる: BOM → BOM の無い UTF-16（NUL の偏り）→ UTF-8 として正しく読めるか → Shift_JIS。
    # 偏りの無い NUL があるとき（判定できないとき）はバイナリとみなし、$null を返す
    #   戻り値: "UTF8" / "UTF16LE" / "UTF16BE" / "ShiftJIS" / $null（バイナリ）
    param (
        [byte[]]$bytes
    )

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return "UTF8"
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
            return $null  # NUL はあるが、どちらの位置にも偏っていない → バイナリ
        }
    }

    try {
        [void](getStrictUtf8Encoding).GetString($bytes)
        return "UTF8"
    } catch [System.Text.DecoderFallbackException] {
        return "ShiftJIS"
    }
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
