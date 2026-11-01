# 2-gram の決まり（tebunko\search\search_gram.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "addTextGrams" -Tag Unit {
    It "隣り合う 2 文字を uint32 にして集合に足す（英字の大小は小文字にそろえる）" {
        $set = New-Object 'System.Collections.Generic.HashSet[uint32]'
        addTextGrams $set "abCd"
        $set.Count | Should -Be 3
        $set.Contains([uint32]([int][char]"a" + 65536 * [int][char]"b")) | Should -BeTrue
        $set.Contains([uint32]([int][char]"c" + 65536 * [int][char]"d")) | Should -BeTrue
        $set.Contains([uint32]([int][char]"a" + 65536 * [int][char]"c")) | Should -BeFalse
    }

    It "1 文字・空の本文" {
        $set = New-Object 'System.Collections.Generic.HashSet[uint32]'
        addTextGrams $set "a"
        addTextGrams $set ""
        $set.Count | Should -Be 0
    }
}

Describe "getSearchGramValues" -Tag Unit {
    It "隣り合う 2 文字を uint32 にし（先の文字 + 65536 * 後の文字）、小文字にそろえる" {
        $values = getSearchGramValues "AB"
        $values.Count | Should -Be 1
        $values[0] | Should -Be ([uint32]([int][char]"a" + 65536 * [int][char]"b"))
    }

    It "本文の 2-gram（addTextGrams）に、本文に含まれる語の 2-gram がすべて入る" -TestCases @(
        @{ text = "保守サービスの見積"; word = "サービス" }
        @{ text = "Hello World"; word = "LO WO" }
        @{ text = "x ab y"; word = "ab" }
    ) {
        $set = New-Object "System.Collections.Generic.HashSet[uint32]"
        addTextGrams $set $text
        foreach ($gram in (getSearchGramValues $word)) {
            $set.Contains($gram) | Should -BeTrue
        }
    }

    It "空・1 文字は空" {
        (getSearchGramValues "").Count | Should -Be 0
        (getSearchGramValues "a").Count | Should -Be 0
    }

    It "多いときは searchGramMax 個に間引く" {
        $word = -join (0..40 | ForEach-Object { [char](0x4E00 + $_) })
        (getSearchGramValues $word).Count | Should -Be ${searchGramMax}
    }
}

Describe "skipRegexClass・skipRegexGroup" -Tag Unit {
    It "skipRegexClass <pattern> -> <expected>" -TestCases @(
        @{ pattern = '[ab]c'; expected = 4 }
        @{ pattern = '[^]a]c'; expected = 5 }
        @{ pattern = '[a\]b]c'; expected = 6 }
        @{ pattern = '[a-z-[aeiou]]c'; expected = 13 }
        @{ pattern = '[[]a]c'; expected = 3 }
        @{ pattern = '[ab'; expected = -1 }
    ) {
        skipRegexClass $pattern 0 | Should -Be $expected
    }

    It "skipRegexGroup <pattern> -> <expected>" -TestCases @(
        @{ pattern = '(ab)c'; expected = 4 }
        @{ pattern = '(a(b)c)d'; expected = 7 }
        @{ pattern = '(a[)]b)c'; expected = 7 }
        @{ pattern = '(a\)b)c'; expected = 6 }
        @{ pattern = '(a(?:b))c'; expected = 8 }
        @{ pattern = '(a(?=b))c'; expected = -1 }
        @{ pattern = '(ab'; expected = -1 }
    ) {
        skipRegexGroup $pattern 0 | Should -Be $expected
    }
}

Describe "readRegexEscape" -Tag Unit {
    It "<pattern> -> 次 <next>・文字 <literal>" -TestCases @(
        @{ pattern = '\.a'; next = 2; literal = '.' }
        @{ pattern = '\da'; next = 2; literal = $null }
        @{ pattern = '\x41a'; next = 4; literal = $null }
        @{ pattern = '\u0041a'; next = 6; literal = $null }
        @{ pattern = '\k<n>a'; next = 5; literal = $null }
        @{ pattern = '\k'; next = -1; literal = $null }
        @{ pattern = '\12a'; next = 3; literal = $null }
        @{ pattern = '\'; next = -1; literal = $null }
    ) {
        $got = readRegexEscape $pattern 0
        $got.Next | Should -Be $next
        $got.Literal | Should -Be $literal
    }
}

Describe "getRegexRequiredText" -Tag Unit {
    It "<pattern> -> <expected>" -TestCases @(
        @{ pattern = '見積.*書'; expected = '見積 書' }
        @{ pattern = '請求書\d+'; expected = '請求書' }
        @{ pattern = 'ab?c'; expected = 'a c' }
        @{ pattern = 'ab+c'; expected = 'ab c' }
        @{ pattern = 'a{2}bc'; expected = 'bc' }
        @{ pattern = '\(株\)山田'; expected = '(株)山田' }
        @{ pattern = '[見積]書類'; expected = '書類' }
        @{ pattern = '(見積)書類'; expected = '書類' }
        @{ pattern = '(?:見積)書類'; expected = '書類' }
        @{ pattern = '第\d+期 決算'; expected = '第 期 決算' }
        @{ pattern = '見積|請求'; expected = '' }
        @{ pattern = '(?i)abc'; expected = '' }
        @{ pattern = '(?=ab)cd'; expected = '' }
        @{ pattern = '(?x)a b'; expected = '' }
        @{ pattern = '(a(?=b))cd'; expected = '' }
        @{ pattern = '^\d{5}\t'; expected = '' }
        @{ pattern = 'a'; expected = 'a' }
        @{ pattern = '[ab'; expected = '' }
        @{ pattern = 'a\.?b'; expected = 'a b' }
        @{ pattern = '\(株\)+山田'; expected = '(株) 山田' }
        @{ pattern = '(ab)?cd'; expected = 'cd' }
        @{ pattern = '[ab]+cd'; expected = 'cd' }
        @{ pattern = '\d*cd'; expected = 'cd' }
        @{ pattern = 'ab{1,3}?c'; expected = 'a c' }
        @{ pattern = 'ab+?c'; expected = 'ab c' }
        @{ pattern = 'a{x}bc'; expected = '' }
        @{ pattern = 'a{x'; expected = '' }
        @{ pattern = 'ab{c|d}ef'; expected = '' }
        @{ pattern = '{x|y}zw'; expected = '' }
        @{ pattern = 'a{1|(b)}cd'; expected = '' }
        @{ pattern = 'a{,5}bc'; expected = '' }
        @{ pattern = 'a{2,}bc'; expected = 'bc' }
        @{ pattern = 'a{2,3}bc'; expected = 'bc' }
        @{ pattern = '[[]a|b]cd'; expected = '' }
        @{ pattern = '[[]ab]cd'; expected = 'ab]cd' }
        @{ pattern = '\p{L'; expected = '' }
        @{ pattern = 'ab\'; expected = '' }
        @{ pattern = 'ab)c'; expected = '' }
        @{ pattern = 'ＡＢ.*ｃ'; expected = 'ＡＢ ｃ' }
        @{ pattern = 'a ?bc'; expected = 'a bc' }
        @{ pattern = '\x41bc'; expected = 'bc' }
        @{ pattern = '\u0041bc'; expected = 'bc' }
        @{ pattern = '\012bc'; expected = 'bc' }
        @{ pattern = '\k<n>bc'; expected = 'bc' }
        @{ pattern = '\p{L}bc'; expected = 'bc' }
        @{ pattern = '[a-z-[aeiou]]bc'; expected = 'bc' }
        @{ pattern = '[]a]bc'; expected = 'bc' }
        @{ pattern = '(a[)]b)cd'; expected = 'cd' }
        @{ pattern = '(a(b)c)de'; expected = 'de' }
        @{ pattern = '(ab'; expected = '' }
    ) {
        getRegexRequiredText $pattern | Should -Be $expected
    }

    It "長い式（(a+)+ を 1 万回つないだもの）が 1 秒以内に返る" {
        $long = "(a+)+" * 10000
        $result = $null
        $elapsed = (Measure-Command { $result = getRegexRequiredText ($long + "bc") }).TotalSeconds
        $result | Should -Be "bc"
        $elapsed | Should -BeLessThan 1
    }
}

Describe "getRegexRequiredText と実際の一致" -Tag Unit {
    It "取り出した文字は、正しい式のどの一致にも含まれる（記号多めの乱数の式）" {
        $random = New-Object System.Random 20261101
        $alphabet = [char[]]'ab{}|()[]?*+.^$\-,12:'
        $samples = @('ab', 'abc', 'a{b', 'b}a', '[a]', 'a|b', 'aab', 'abab', 'a1b2', '(a)', 'ba', 'a-b', 'a,b', 'a.b', 'a{1}', 'a{1,2}b', 'b{', '{a}', 'ab)', '[a', 'a]', 'a?b')
        $bad = New-Object System.Collections.Generic.List[string]
        for ($n = 0; $n -lt 3000; $n++) {
            $length = $random.Next(2, 10)
            $chars = for ($k = 0; $k -lt $length; $k++) { $alphabet[$random.Next($alphabet.Length)] }
            $pattern = -join $chars
            $regex = $null
            try { $regex = New-Object System.Text.RegularExpressions.Regex($pattern, [System.Text.RegularExpressions.RegexOptions]::CultureInvariant, [TimeSpan]::FromSeconds(1)) } catch { continue }
            $required = getRegexRequiredText $pattern
            if ($required -eq '') { continue }
            $parts = $required -split ' '
            foreach ($sample in $samples) {
                foreach ($m in $regex.Matches($sample)) {
                    foreach ($part in $parts) {
                        if (!$m.Value.Contains($part)) { $bad.Add("$pattern -> $required (一致 $($m.Value))") }
                    }
                }
            }
        }
        ($bad | Select-Object -First 5) -join "`n" | Should -BeNullOrEmpty
    }
}

Describe "getGramSearchText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "文字どおりは語そのまま"; word = 'a.b'; simple = $true; expected = 'a.b' }
        @{ name = "正規表現は必ず含まれる文字"; word = 'a.b'; simple = $false; expected = 'a b' }
        @{ name = "不正な式（入れ子の量指定）は語そのまま"; word = 'a\.bc**'; simple = $false; expected = 'a\.bc**' }
        @{ name = "不正な式（範囲が逆）は語そのまま"; word = 'a\.b{2,1}'; simple = $false; expected = 'a\.b{2,1}' }
        @{ name = "不正な式（無い後方参照）は語そのまま"; word = '見積\.書\1'; simple = $false; expected = '見積\.書\1' }
        @{ name = "不正な式（閉じすぎ）は語そのまま"; word = 'a\.b)'; simple = $false; expected = 'a\.b)' }
    ) {
        getGramSearchText $word $simple | Should -Be $expected
    }
}
