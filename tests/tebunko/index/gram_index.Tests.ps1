# gram 索引の決まり（tebunko\index\gram_index.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:newSegment {
        # 区画（新しい順に並べて渡す）。names は大文字と小文字を区別せず昇順
        param ([string[]]$names, [long]$ticks, [long]$size, [int[]]$candidates)
        $count = $names.Count
        return @{
            Names = $names
            Ticks = [long[]]@(1..$count | ForEach-Object { $ticks })
            Sizes = [long[]]@(1..$count | ForEach-Object { $size })
            Candidates = [System.Collections.Generic.HashSet[int]]::new([int[]]$candidates)
        }
    }

    function script:newItem {
        param ([string]$rel, [long]$ticks = 5, [long]$size = 7)
        return @{ RelPath = $rel; Ticks = $ticks; Size = $size }
    }
}

Describe "testGramIndexUsable" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "文字どおりで 2-gram があれば使える"; simple = $true; word = "ab"; expected = $true }
        @{ name = "正規表現は使えない"; simple = $false; word = "ab"; expected = $false }
        @{ name = "2-gram が作れない語は使えない"; simple = $true; word = "a"; expected = $false }
    ) {
        testGramIndexUsable $simple (getSearchGramValues $word) | Should -Be $expected
    }
}

Describe "区画の先頭と置き場所" -Tag Unit {
    It "先頭を作って読み戻せる" {
        $bytes = convertToGramSegmentHeader 3 20 5 9
        $bytes.Length | Should -Be ${gramSegmentHeaderBytes}
        $header = convertFromGramSegmentHeader $bytes
        $header.FileCount | Should -Be 3
        $header.NameBytes | Should -Be 20
        $header.GramCount | Should -Be 5
        $header.PostingCount | Should -Be 9
        $header.Layout.Length | Should -Be (24 + 16 * 3 + 20 + 4 * 5 + 4 * 6 + 4 * 9 + 4)
    }

    It "印・版が違う、短い、数が負の先頭は読めない" {
        convertFromGramSegmentHeader $null | Should -BeNullOrEmpty
        convertFromGramSegmentHeader ([byte[]]@(1, 2, 3)) | Should -BeNullOrEmpty
        $badMagic = convertToGramSegmentHeader 1 1 1 1
        $badMagic[0] = 0
        convertFromGramSegmentHeader $badMagic | Should -BeNullOrEmpty
        $badVersion = convertToGramSegmentHeader 1 1 1 1
        $badVersion[4] = 9
        convertFromGramSegmentHeader $badVersion | Should -BeNullOrEmpty
        $negative = convertToGramSegmentHeader -1 1 1 1
        convertFromGramSegmentHeader $negative | Should -BeNullOrEmpty
    }

    It "壊れていないかを長さ・目録の長さ・末尾の印で確かめる" {
        $header = convertFromGramSegmentHeader (convertToGramSegmentHeader 1 4 1 1)
        $length = $header.Layout.Length
        $footer = [int]${gramSegmentFooter}
        testGramSegmentIntact $header $length $length $footer | Should -BeTrue
        testGramSegmentIntact $header ($length - 1) ($length - 1) $footer | Should -BeFalse
        testGramSegmentIntact $header $length ($length + 1) $footer | Should -BeFalse
        testGramSegmentIntact $header $length $length 0 | Should -BeFalse
        testGramSegmentIntact $null $length $length $footer | Should -BeFalse
    }
}

Describe "目録" -Tag Unit {
    It "書いて読み戻せる" {
        $toc = @{
            Next = 4
            Segments = @(@{ Name = (getGramSegmentName 1); Length = 100 }, @{ Name = (getGramSegmentName 3); Length = 200 })
            Old = @((getGramSegmentName 2))
        }
        $back = convertFromGramToc (convertToGramToc $toc)
        $back.Next | Should -Be 4
        $back.Segments.Count | Should -Be 2
        $back.Segments[1].Name | Should -Be "g0000000003.seg"
        $back.Segments[1].Length | Should -Be 200
        $back.Old[0] | Should -Be "g0000000002.seg"
    }

    It "<name>は読めない" -TestCases @(
        @{ name = "空"; lines = @() }
        @{ name = "印が違う"; lines = @("other`t1") }
        @{ name = "版が違う"; lines = @("tebunko-gram`t2") }
        @{ name = "区画の名前が使えない"; lines = @("tebunko-gram`t1", "seg`t..\x.seg`t1") }
        @{ name = "長さが数でない"; lines = @("tebunko-gram`t1", "seg`tg0000000001.seg`tx") }
        @{ name = "長さが負"; lines = @("tebunko-gram`t1", "seg`tg0000000001.seg`t-1") }
        @{ name = "知らない行"; lines = @("tebunko-gram`t1", "foo`t1") }
        @{ name = "番号が 0"; lines = @("tebunko-gram`t1", "next`t0") }
        @{ name = "列が足りない"; lines = @("tebunko-gram`t1", "seg`tg0000000001.seg") }
    ) {
        convertFromGramToc $lines | Should -BeNullOrEmpty
    }

    It "区画の名前は g と 10 桁の数字と .seg だけ" -TestCases @(
        @{ name = "g0000000001.seg"; expected = $true }
        @{ name = "g000000001.seg"; expected = $false }
        @{ name = "G0000000001.seg"; expected = $false }
        @{ name = "g0000000001.seg.tmp"; expected = $false }
        @{ name = "..\g0000000001.seg"; expected = $false }
        @{ name = ""; expected = $false }
    ) {
        testGramSegmentName $name | Should -Be $expected
    }
}

Describe "共通部分と選び方" -Tag Unit {
    It "intersectGramRows は共通のファイルの番号を返す" {
        $set = intersectGramRows @(, @(1, 2, 3))
        @($set | Sort-Object) -join "," | Should -Be "1,2,3"
        $set = intersectGramRows @([int[]]@(1, 2, 3), [int[]]@(2, 3, 4), [int[]]@(3, 9))
        @($set) -join "," | Should -Be "3"
    }

    It "intersectGramRows は載っていない 2-gram があれば空" {
        (intersectGramRows @([int[]]@(1, 2), $null)).Count | Should -Be 0
        (intersectGramRows @([int[]]@(1, 2), [int[]]@())).Count | Should -Be 0
        (intersectGramRows @([int[]]@(1), [int[]]@(2))).Count | Should -Be 0
        (intersectGramRows @()).Count | Should -Be 0
    }

    It "selectGramFiles は候補に無い同じ版のファイルだけを外す" {
        $segment = newSegment @("a.tsv", "b.tsv", "c.tsv") 5 7 @(1)
        $items = @((newItem "a.tsv"), (newItem "b.tsv"), (newItem "c.tsv"))
        $selected = selectGramFiles $items @($segment)
        @($selected.Files | ForEach-Object { $_.RelPath }) -join "," | Should -Be "b.tsv"
        $selected.Excluded | Should -Be 2
    }

    It "selectGramFiles は更新日時か大きさが違う・載っていないファイルを必ず残す" {
        $segment = newSegment @("a.tsv") 5 7 @()
        $items = @((newItem "a.tsv" 6 7), (newItem "a.tsv" 5 8), (newItem "z.tsv"))
        $selected = selectGramFiles $items @($segment)
        $selected.Files.Count | Should -Be 3
        $selected.Excluded | Should -Be 0
    }

    It "selectGramFiles は区画が無ければすべて残す" {
        (selectGramFiles @((newItem "a.tsv")) @()).Files.Count | Should -Be 1
    }

    It "selectGramFiles は新しい区画の記録を優先する（古い区画の候補は使わない）" {
        $newer = newSegment @("a.tsv") 5 7 @()
        $older = newSegment @("a.tsv") 5 7 @(0)
        (selectGramFiles @((newItem "a.tsv")) @($newer, $older)).Files.Count | Should -Be 0
        (selectGramFiles @((newItem "a.tsv")) @($older, $newer)).Files.Count | Should -Be 1
    }

    It "selectGramFiles は新しい区画の版が違えば、古い区画で外さない" {
        $newer = newSegment @("a.tsv") 9 7 @(0)
        $older = newSegment @("a.tsv") 5 7 @()
        (selectGramFiles @((newItem "a.tsv" 5 7)) @($newer, $older)).Files.Count | Should -Be 1
    }

    It "selectGramFiles は名前の大文字と小文字を区別しない" {
        $segment = newSegment @("A.tsv") 5 7 @()
        (selectGramFiles @((newItem "a.tsv")) @($segment)).Files.Count | Should -Be 0
    }

    It "getGramFilesToAdd は載っていない・版の違うファイルだけを返す" {
        $segment = newSegment @("a.tsv", "b.tsv") 5 7 @()
        $items = @((newItem "a.tsv"), (newItem "b.tsv" 6 7), (newItem "c.tsv"))
        @(getGramFilesToAdd $items @($segment) | ForEach-Object { $_.RelPath }) -join "," | Should -Be "b.tsv,c.tsv"
        (getGramFilesToAdd $items @()).Count | Should -Be 3
    }
}

Describe "区切りとまとめ直し" -Tag Unit {
    It "splitGramBatches は組の数の合計が上限を超えないように前から区切る" {
        $batches = splitGramBatches ([int[]]@(5, 5, 5, 5)) 10
        $batches.Count | Should -Be 2
        $batches[0].Start | Should -Be 0
        $batches[0].Count | Should -Be 2
        $batches[1].Start | Should -Be 2
        $batches[1].Count | Should -Be 2
    }

    It "splitGramBatches は 1 つで上限を超えるファイルを分けない" {
        $batches = splitGramBatches ([int[]]@(1, 50, 1)) 10
        @($batches | ForEach-Object { $_.Count }) -join "," | Should -Be "1,1,1"
        (splitGramBatches ([int[]]@()) 10).Count | Should -Be 0
    }

    It "testGramMergeNeeded は<name>" -TestCases @(
        @{ name = "区画が上限以下ならまとめない"; count = 8; postings = 10; limit = 100; expected = $false }
        @{ name = "上限を超え、まとめると減るならまとめる"; count = 9; postings = 10; limit = 100; expected = $true }
        @{ name = "まとめても同じ数に分かれるならまとめない"; count = 9; postings = 900; limit = 100; expected = $false }
    ) {
        testGramMergeNeeded $count $postings $limit | Should -Be $expected
    }

    It "getGramIndexNames はパスにつなげない名前を外す" {
        (getGramIndexNames @("a", "..\b", "")) -join "," | Should -Be "a"
    }
}
