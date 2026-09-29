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
