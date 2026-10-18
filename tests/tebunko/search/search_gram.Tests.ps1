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