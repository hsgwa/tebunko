# テキストファイルの対象の拡張子・文字コードの判定・読み込み（shared\core\text_file.ps1）のテスト
BeforeDiscovery {
    # -TestCases の表がバイト列を組み立てるのに使う（Discovery 段階で評価されるため、BeforeAll ではなくここに置く）
    function newNulPairBytes {
        # 2 バイトずつの組を pairs 個作り、先頭から evenNul 個の組は偶数の位置（0番目）を、
        # 先頭から oddNul 個の組は奇数の位置（1番目）を NUL にする（それ以外は非 0 のダミーのバイト）
        param (
            [int]$pairs,
            [int]$evenNul,
            [int]$oddNul
        )

        $bytes = New-Object byte[] ($pairs * 2)
        for ($i = 0; $i -lt $pairs; $i++) {
            $bytes[$i * 2] = if ($i -lt $evenNul) { 0 } else { 0x41 }
            $bytes[$i * 2 + 1] = if ($i -lt $oddNul) { 0 } else { 0x42 }
        }
        return , $bytes
    }

    function toShiftJisBytes([string]$text) {
        return , ([System.Text.Encoding]::GetEncoding(932).GetBytes($text))
    }

    function toUtf16LeBytes([string]$text, [bool]$bom = $false) {
        $bytes = [System.Text.Encoding]::Unicode.GetBytes($text)
        if ($bom) { return , (@(0xFF, 0xFE) + $bytes) }
        return , $bytes
    }

    function toUtf16BeBytes([string]$text, [bool]$bom = $false) {
        $bytes = [System.Text.Encoding]::BigEndianUnicode.GetBytes($text)
        if ($bom) { return , (@(0xFE, 0xFF) + $bytes) }
        return , $bytes
    }

    function toUtf8Bytes([string]$text, [bool]$bom = $false) {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
        if ($bom) { return , (@(0xEF, 0xBB, 0xBF) + $bytes) }
        return , $bytes
    }
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    # Run 段階の It（-TestCases の表を使わないもの）でも使うため、BeforeDiscovery と同じものをここにも置く
    function toShiftJisBytes([string]$text) {
        return , ([System.Text.Encoding]::GetEncoding(932).GetBytes($text))
    }

    function newNulPairBytes {
        param (
            [int]$pairs,
            [int]$evenNul,
            [int]$oddNul
        )

        $bytes = New-Object byte[] ($pairs * 2)
        for ($i = 0; $i -lt $pairs; $i++) {
            $bytes[$i * 2] = if ($i -lt $evenNul) { 0 } else { 0x41 }
            $bytes[$i * 2 + 1] = if ($i -lt $oddNul) { 0 } else { 0x42 }
        }
        return , $bytes
    }

    function toUtf8Bytes([string]$text, [bool]$bom = $false) {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
        if ($bom) { return , (@(0xEF, 0xBB, 0xBF) + $bytes) }
        return , $bytes
    }
}

Describe "testTextExtension" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "対象の拡張子（7つ）はどれも対象"; path = "a.txt"; expected = $true }
        @{ name = "大文字・小文字を区別しない"; path = "a.TXT"; expected = $true }
        @{ name = "csv も対象"; path = "a.csv"; expected = $true }
        @{ name = "対象外の拡張子"; path = "a.pdf"; expected = $false }
        @{ name = "Office の拡張子は対象外（testTextExtension としては）"; path = "a.xlsx"; expected = $false }
    ) {
        param ($name, $path, $expected)
        testTextExtension $path | Should -Be $expected
    }
}

Describe "detectTextEncoding" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "UTF-8 の BOM"; bytes = (toUtf8Bytes "あ" $true); expected = "UTF8" }
        @{ name = "UTF-16LE の BOM"; bytes = (toUtf16LeBytes "あ" $true); expected = "UTF16LE" }
        @{ name = "UTF-16BE の BOM"; bytes = (toUtf16BeBytes "あ" $true); expected = "UTF16BE" }
        @{ name = "ASCII だけ（BOM 無し）→ UTF-8"; bytes = ([System.Text.Encoding]::ASCII.GetBytes("abc")); expected = "UTF8" }
        @{ name = "空のファイル → UTF-8"; bytes = ([byte[]]@()); expected = "UTF8" }
        @{ name = "BOM の無い UTF-8（日本語。NUL が無いので UTF-8 として読める）"; bytes = (toUtf8Bytes "日本語"); expected = "UTF8" }
        @{ name = "ASCII を含む BOM 無し UTF-16LE"; bytes = (toUtf16LeBytes "hello"); expected = "UTF16LE" }
        @{ name = "ASCII を含む BOM 無し UTF-16BE"; bytes = (toUtf16BeBytes "hello"); expected = "UTF16BE" }
        @{ name = "NUL が偏らないバイナリ → バイナリ"; bytes = (newNulPairBytes 10 5 5); expected = $null }
        @{ name = "閾値ちょうど20%（奇数側）→ UTF-16LE"; bytes = (newNulPairBytes 10 0 2); expected = "UTF16LE" }
        @{ name = "20%を1組下回る（奇数側）→ バイナリ"; bytes = (newNulPairBytes 10 0 1); expected = $null }
        @{ name = "少ない側がちょうど1/10（奇数側が多い）→ UTF-16LE"; bytes = (newNulPairBytes 10 1 10); expected = "UTF16LE" }
        @{ name = "少ない側が1/10を1つ上回る → バイナリ"; bytes = (newNulPairBytes 10 2 10); expected = $null }
        @{ name = "閾値ちょうど20%（偶数側）→ UTF-16BE"; bytes = (newNulPairBytes 10 2 0); expected = "UTF16BE" }
        @{ name = "ASCII と日本語が混ざり NUL が20%に届かない BOM 無し UTF-16LE → バイナリ（限界）"; bytes = (toUtf16LeBytes (("日" * 20) + "abc")); expected = $null }
    ) {
        param ($name, $bytes, $expected)
        detectTextEncoding $bytes | Should -Be $expected
    }

    It "短い Shift_JIS の半角カナ（例: ﾃｽ = C3 BD）は UTF-8 としても正しく読めるため UTF-8 と判定される（限界。判定の順が変わったら気付けるようにする）" {
        $bytes = toShiftJisBytes "ﾃｽ"
        ([BitConverter]::ToString($bytes)) | Should -Be "C3-BD"
        detectTextEncoding $bytes | Should -Be "UTF8"
    }

    It "壊れた UTF-8（不正なバイト列。NUL は無い）→ Shift_JIS" {
        $bytes = toShiftJisBytes "日本語のテスト"
        detectTextEncoding $bytes | Should -Be "ShiftJIS"
    }
}

Describe "decodeTextBytes" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "UTF-8（BOM 無し）"; bytes = (toUtf8Bytes "あ"); encodingName = "UTF8"; expected = "あ" }
        @{ name = "UTF-8（BOM 付き。BOM を取り除く）"; bytes = (toUtf8Bytes "あ" $true); encodingName = "UTF8"; expected = "あ" }
        @{ name = "UTF-16LE（BOM 付き。BOM を取り除く）"; bytes = (toUtf16LeBytes "あ" $true); encodingName = "UTF16LE"; expected = "あ" }
        @{ name = "UTF-16LE（BOM 無し）"; bytes = (toUtf16LeBytes "あ"); encodingName = "UTF16LE"; expected = "あ" }
        @{ name = "UTF-16BE（BOM 付き。BOM を取り除く）"; bytes = (toUtf16BeBytes "あ" $true); encodingName = "UTF16BE"; expected = "あ" }
        @{ name = "UTF-16BE（BOM 無し）"; bytes = (toUtf16BeBytes "あ"); encodingName = "UTF16BE"; expected = "あ" }
    ) {
        param ($name, $bytes, $encodingName, $expected)
        $text = decodeTextBytes $bytes $encodingName
        $text | Should -Be $expected
        ([int][char]$text[0]) | Should -Not -Be 0xFEFF
    }

    It "Shift_JIS" {
        (decodeTextBytes (toShiftJisBytes "あ") "ShiftJIS") | Should -Be "あ"
    }
}

Describe "readTextFile" -Tag Io {
    BeforeAll {
        function writeBytesFile([byte[]]$bytes) {
            $path = Join-Path $TestDrive ("f_" + [guid]::NewGuid().ToString("N") + ".txt")
            [System.IO.File]::WriteAllBytes($path, $bytes)
            return $path
        }
    }

    It "<name>" -TestCases @(
        @{ name = "UTF-8（BOM 無し）"; bytes = (toUtf8Bytes "1行目`r`n2行目`r`n") }
        @{ name = "UTF-8（BOM 付き。BOM が先頭行に残らない）"; bytes = (toUtf8Bytes "1行目`r`n2行目`r`n" $true) }
        @{ name = "UTF-16LE（BOM 付き。BOM が先頭行に残らない）"; bytes = (toUtf16LeBytes "1行目`r`n2行目`r`n" $true) }
        @{ name = "UTF-16LE（BOM 無し。ASCII を含む）"; bytes = (toUtf16LeBytes "1line`r`n2line`r`n") }
        @{ name = "UTF-16BE（BOM 付き。BOM が先頭行に残らない）"; bytes = (toUtf16BeBytes "1行目`r`n2行目`r`n" $true) }
        @{ name = "UTF-16BE（BOM 無し。ASCII を含む）"; bytes = (toUtf16BeBytes "1line`r`n2line`r`n") }
    ) {
        param ($name, $bytes)
        $lines = readTextFile (writeBytesFile $bytes)
        $lines.Count | Should -Be 2
        $lines[0] | Should -BeIn @("1行目", "1line")
        $lines[1] | Should -BeIn @("2行目", "2line")
        ([int][char]$lines[0][0]) | Should -Not -Be 0xFEFF
    }

    It "Shift_JIS のファイルを行の並びとして読める" {
        $lines = readTextFile (writeBytesFile (toShiftJisBytes "1行目`r`n2行目`r`n"))
        $lines | Should -Be @("1行目", "2行目")
    }

    It "CRLF・LF・CR が混ざったファイルの N 行目が、元の N 行目のまま読める（途中の空行も残る）" {
        $lines = readTextFile (writeBytesFile (toUtf8Bytes "1行目`r`n2行目`n3行目`r4行目`n`n6行目`n"))
        $lines | Should -Be @("1行目", "2行目", "3行目", "4行目", "", "6行目")
    }

    It "大きさの上限（差し替えた小さい値）を超えるファイルは、決めた文言で失敗にする" {
        $path = writeBytesFile (toUtf8Bytes "12345678")
        { readTextFile $path 4 } | Should -Throw "ファイルサイズが大きすぎるため取り込めません。"
    }

    It "バイナリと判定したファイルは、決めた文言で失敗にする" {
        $path = writeBytesFile (newNulPairBytes 10 5 5)
        { readTextFile $path } | Should -Throw "テキストファイルではないため取り込めません。"
    }
}

Describe "splitTextLines" -Tag Unit {
    It "CRLF・LF・CR のどれでも1行にし、途中の空行を残し、行末の空白を取り除き、末尾の空行は捨てる" {
        # "a " CRLF "b" LF "c" CR "d" LF LF "e " CRLF CRLF
        # → 行: a / b / c / d / （空） / e / （末尾の空行。捨てる）
        $lines = splitTextLines "a `r`nb`nc`rd`n`ne `r`n`r`n"
        ($lines -join "|") | Should -Be "a|b|c|d||e"
    }

    It "空文字列は空の配列を返す" {
        (splitTextLines "").Count | Should -Be 0
    }
}
