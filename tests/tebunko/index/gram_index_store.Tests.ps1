# gram 索引の読み書き（tebunko\index\gram_index_store.ps1）と、検索への組み込み（tebunko\search\gram_search.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:putContentIndex {
        # インデックス index の中に、本文インデックスのファイルを 1 つ置く（folder はインデックス内のフォルダ）
        param ([string]$root, [string]$index, [string]$folder, [string]$bookName, [string]$text)
        $path = "$root\$index\$folder\content_index.xlsx.001.tsv"
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
        writeContentIndexFile $path (convertToContentIndexText @(@{ Name = $bookName; Places = @(@{ Place = "S"; Text = $text }) }))
        return $path
    }

    function script:newGramTree {
        # 3 つのファイルを持つインデックス「営業」を作る。{ Index; Gram } を返す
        param ([string]$base)
        $index = "$base\index"
        putContentIndex $index "営業" "a" "A.xlsx" "モニター 27インチ 保守サービス" | Out-Null
        putContentIndex $index "営業" "b" "B.xlsx" "千代田区 見積書" | Out-Null
        putContentIndex $index "営業" "c" "C.xlsx" "Hello World" | Out-Null
        return @{ Index = $index; Gram = "$base\gram" }
    }

    function script:candidateRels {
        # gram 索引で絞った結果の、本文インデックスのファイルの相対パス（; 区切り）。$null は絞れなかった
        param ([string]$word, $tree, [bool]$simple = $true)
        $got = getGramContentIndexFiles $word $simple @($tree.Index) $tree.Index $tree.Gram
        if ($null -eq $got) { return $null }
        return (@($got.ContentIndexFiles | ForEach-Object { $_.RelPath }) -join ";")
    }
}

Describe "gram 索引を作って絞る" -Tag Io {
    It "作ったあと、検索語のすべての 2-gram を含むファイルだけに絞る" {
        $tree = newGramTree "$TestDrive\s1"
        $result = updateGramIndexOf $tree.Index $tree.Gram "営業"
        $result.Files | Should -Be 3
        $result.Added | Should -Be 1
        candidateRels "モニター" $tree | Should -Be "営業\a\content_index.xlsx.001.tsv"
        candidateRels "千代田区" $tree | Should -Be "営業\b\content_index.xlsx.001.tsv"
        candidateRels "hello" $tree | Should -Be "営業\c\content_index.xlsx.001.tsv"
        candidateRels "存在しない語" $tree | Should -Be ""
    }

    It "検索語を含むファイルは必ず残る（全部照合した結果に漏れが無い）" -TestCases @(
        @{ word = "27インチ" }
        @{ word = "WORLD" }
        @{ word = "hello world" }
        @{ word = "見積" }
        @{ word = "サー" }
    ) {
        $tree = newGramTree "$TestDrive\eq$([guid]::NewGuid().ToString('N'))"
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        $kept = (candidateRels $word $tree) -split ";"
        foreach ($item in (findContentIndexFiles $tree.Index "営業" $true)) {
            $body = (getContentIndexBodyText (readContentIndexText $item.Path)).ToLowerInvariant()
            if ($body.Contains($word.ToLowerInvariant())) {
                $kept | Should -Contain $item.RelPath
            }
        }
    }

    It "正規表現・2-gram を作れない語・置き場が無いときは絞らない（\$null）" {
        $tree = newGramTree "$TestDrive\s2"
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        candidateRels "モニター" $tree $false | Should -BeNullOrEmpty
        candidateRels "a" $tree | Should -BeNullOrEmpty
        $none = getGramContentIndexFiles "モニター" $true @($tree.Index) $tree.Index "$TestDrive\s2\nothing"
        $none | Should -BeNullOrEmpty
    }

    It "変わっていなければ 2 回目は何も足さない" {
        $tree = newGramTree "$TestDrive\s3"
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        $again = updateGramIndexOf $tree.Index $tree.Gram "営業"
        $again.Added | Should -Be 0
        $again.Files | Should -Be 0
    }

    It "新しいファイルは区画として足し、更新されたファイルは新しい区画の内容を使う" {
        $tree = newGramTree "$TestDrive\s4"
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        putContentIndex $tree.Index "営業" "d" "D.xlsx" "追加した新しい帳票" | Out-Null
        $path = putContentIndex $tree.Index "営業" "b" "B.xlsx" "全く違う本文になった"
        $result = updateGramIndexOf $tree.Index $tree.Gram "営業"
        $result.Added | Should -Be 1
        $result.Files | Should -Be 2
        candidateRels "新しい帳票" $tree | Should -Be "営業\d\content_index.xlsx.001.tsv"
        candidateRels "千代田区" $tree | Should -Be ""
        candidateRels "違う本文" $tree | Should -Be "営業\b\content_index.xlsx.001.tsv"
    }

    It "索引に載っていないファイル・更新日時か大きさが索引と違うファイルは、必ず照合する" {
        $tree = newGramTree "$TestDrive\s5"
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        putContentIndex $tree.Index "営業" "e" "E.xlsx" "索引にまだ無い" | Out-Null
        (candidateRels "モニター" $tree) -split ";" | Should -Contain "営業\e\content_index.xlsx.001.tsv"
        # 索引を作ったあとで書き換えた（モニターという語が入った）
        putContentIndex $tree.Index "営業" "b" "B.xlsx" "書き換えた後にモニターが入った" | Out-Null
        (candidateRels "モニター" $tree) -split ";" | Should -Contain "営業\b\content_index.xlsx.001.tsv"
    }

    It "区画が上限を超えたら 1 つにまとめ直し、外した区画は 1 世代だけ残し、その次に消す" {
        $tree = newGramTree "$TestDrive\s6"
        $dir = "$($tree.Gram)\営業"
        for ($i = 0; $i -lt 9; $i++) {
            putContentIndex $tree.Index "営業" "n$i" "N$i.xlsx" "追加の本文 $i" | Out-Null
            $r = updateGramIndexOf $tree.Index $tree.Gram "営業"
        }
        $r.Merged | Should -BeTrue
        $toc = readGramToc $dir
        $toc.Segments.Count | Should -Be 1
        $toc.Old.Count | Should -BeGreaterThan 0
        foreach ($old in $toc.Old) { [System.IO.File]::Exists("$dir\$old") | Should -BeTrue }
        putContentIndex $tree.Index "営業" "z" "Z.xlsx" "さらに追加" | Out-Null
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        foreach ($old in $toc.Old) { [System.IO.File]::Exists("$dir\$old") | Should -BeTrue }
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        foreach ($old in $toc.Old) { [System.IO.File]::Exists("$dir\$old") | Should -BeFalse }
        candidateRels "追加の本文" $tree | Should -Not -BeNullOrEmpty
    }

    It "<name>ときは絞らず、次の回に作り直す" -TestCases @(
        @{ name = "区画を壊した"; action = { param($dir) [System.IO.File]::WriteAllBytes((Get-ChildItem $dir -Filter "g*.seg" | Select-Object -First 1).FullName, [byte[]](1, 2, 3)) } }
        @{ name = "区画を消した"; action = { param($dir) Remove-Item (Get-ChildItem $dir -Filter "g*.seg" | Select-Object -First 1).FullName } }
        @{ name = "目録を壊した"; action = { param($dir) [System.IO.File]::WriteAllText("$dir\toc.tsv", "壊れた") } }
        @{ name = "目録を消した"; action = { param($dir) Remove-Item "$dir\toc.tsv" } }
    ) {
        $tree = newGramTree "$TestDrive\d$([guid]::NewGuid().ToString('N'))"
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        $dir = "$($tree.Gram)\営業"
        & $action $dir
        # 絞らない = 3 つとも残る
        (candidateRels "モニター" $tree) -split ";" | Should -HaveCount 3
        $again = updateGramIndexOf $tree.Index $tree.Gram "営業"
        $again.Files | Should -Be 3
        candidateRels "モニター" $tree | Should -Be "営業\a\content_index.xlsx.001.tsv"
    }

    It "目録に載っていない区画と書きかけの一時ファイルは、次の回の始めに消す" {
        $tree = newGramTree "$TestDrive\s7"
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        $dir = "$($tree.Gram)\営業"
        [System.IO.File]::WriteAllText("$dir\g0000000099.seg", "x")
        [System.IO.File]::WriteAllText("$dir\g0000000098.seg.tmp", "x")
        [void](updateGramIndexOf $tree.Index $tree.Gram "営業")
        [System.IO.File]::Exists("$dir\g0000000099.seg") | Should -BeFalse
        [System.IO.File]::Exists("$dir\g0000000098.seg.tmp") | Should -BeFalse
        [System.IO.File]::Exists("$dir\toc.tsv") | Should -BeTrue
    }

    It "止めるよう言われたら目録を書かずに止める" {
        $tree = newGramTree "$TestDrive\s8"
        $result = updateGramIndexOf $tree.Index $tree.Gram "営業" { $true }
        $result.Stopped | Should -BeTrue
        readGramToc "$($tree.Gram)\営業" | Should -BeNullOrEmpty
        (candidateRels "モニター" $tree) -split ";" | Should -HaveCount 3
    }
}

Describe "updateGramIndexes と removeGramIndexOf" -Tag Io {
    It "インデックスごとに作り、無くなったインデックスの索引を消す" {
        $tree = newGramTree "$TestDrive\u1"
        putContentIndex $tree.Index "人事" "a" "H.xlsx" "給与明細" | Out-Null
        $summary = updateGramIndexes $tree.Index $tree.Gram
        $summary.Failed.Count | Should -Be 0
        $summary.Files | Should -Be 4
        [System.IO.Directory]::Exists("$($tree.Gram)\人事") | Should -BeTrue
        Remove-Item "$($tree.Index)\人事" -Recurse
        [void](updateGramIndexes $tree.Index $tree.Gram)
        [System.IO.Directory]::Exists("$($tree.Gram)\人事") | Should -BeFalse
        [System.IO.Directory]::Exists("$($tree.Gram)\営業") | Should -BeTrue
    }

    It "removeGramIndexOf は索引のフォルダを消す" {
        $tree = newGramTree "$TestDrive\u2"
        [void](updateGramIndexes $tree.Index $tree.Gram)
        [void](removeGramIndexOf "営業" $tree.Gram)
        [System.IO.Directory]::Exists("$($tree.Gram)\営業") | Should -BeFalse
    }
}

Describe "getFileGramValues" -Tag Io {
    It "読む間に変わりうる・大きすぎる・読めないファイルは \$null" {
        $tree = newGramTree "$TestDrive\f1"
        $item = (findContentIndexFiles $tree.Index "営業" $true)[0]
        (getFileGramValues $item.Path $item.Ticks $item.Size).Count | Should -BeGreaterThan 0
        getFileGramValues $item.Path ($item.Ticks + 1) $item.Size | Should -BeNullOrEmpty
        getFileGramValues $item.Path $item.Ticks ($item.Size + 1) | Should -BeNullOrEmpty
        getFileGramValues "$TestDrive\f1\none.tsv" 1 1 | Should -BeNullOrEmpty
    }

    It "検索語の 2-gram は、その語を含む本文の 2-gram に必ず入る（addTextGrams と同じ値）" {
        $tree = newGramTree "$TestDrive\f2"
        $item = (findContentIndexFiles $tree.Index "営業" $true) | Where-Object { $_.RelPath -like "*\c\*" }
        $values = getFileGramValues $item.Path $item.Ticks $item.Size
        foreach ($gram in (getSearchGramValues "hello world")) {
            $values | Should -Contain $gram
        }
    }
}

Describe "特殊な文字でも、全部を照合した結果が絞り込みで漏れない" -Tag Io {
    It "<word>（大文字小文字の区別: <case>）" -TestCases @(
        @{ word = "kb"; case = $false }
        @{ word = "KB"; case = $false }
        @{ word = "k"; case = $false }
        @{ word = "kel"; case = $false }
        @{ word = "istanbul"; case = $false }
        @{ word = "ISTANBUL"; case = $false }
        @{ word = "i̇stanbul"; case = $false }
        @{ word = "straße"; case = $false }
        @{ word = "STRASSE"; case = $false }
        @{ word = "strasse"; case = $false }
        @{ word = "ßa"; case = $false }
        @{ word = "ＡＢ"; case = $false }
        @{ word = "ab"; case = $false }
        @{ word = "Kelvin"; case = $true }
        @{ word = "İstanbul"; case = $true }
    ) {
        $base = "$TestDrive\sp$([guid]::NewGuid().ToString('N'))"
        $index = "$base\index"
        $texts = @{
            k = "Kelvin $([char]0x212A)B 測定"
            i = "$([char]0x0130)stanbul i$([char]0x0307)stanbul Istanbul"
            s = "Straße STRASSE strasse"
            f = "ＡＢ ab AB"
        }
        foreach ($key in $texts.Keys) {
            putContentIndex $index "営業" $key "$key.xlsx" $texts[$key] | Out-Null
        }
        [void](updateGramIndexOf $index "$base\gram" "営業")
        $files = getContentIndexFiles @($index)
        $full = searchContentIndex $word $files.ContentIndexFiles $true 0 $null $case
        $gram = getGramContentIndexFiles $word $true @($index) $index "$base\gram"
        if ($null -eq $gram) {
            return
        }
        $kept = @($gram.ContentIndexFiles | ForEach-Object { $_.RelPath })
        foreach ($rel in @($full.Hits | ForEach-Object { $_.RelPath } | Select-Object -Unique)) {
            $kept | Should -Contain $rel
        }
        $narrow = searchContentIndex $word $gram.ContentIndexFiles $true 0 $null $case
        (@($narrow.Hits | ForEach-Object { "$($_.RelPath):$($_.LineNumber)" }) -join "|") | Should -Be (@($full.Hits | ForEach-Object { "$($_.RelPath):$($_.LineNumber)" }) -join "|")
    }
}

Describe "前の版の gram 索引（tests\testdata\compat\gram\v1）が今のコードで読める" -Tag Io {
    It "目録と区画を読み、検索語の候補を求められる（この見本は変えない・消さない。形を変えるなら新しい見本を足す）" -TestCases @(
        @{ word = "モニター"; expected = "営業\a\content_index.xlsx.001.tsv" }
        @{ word = "千代田区"; expected = "営業\b\content_index.xlsx.001.tsv" }
        @{ word = "world"; expected = "営業\c\content_index.xlsx.001.tsv" }
        @{ word = "存在しない"; expected = "" }
    ) {
        $dir = "${testDataDir}\compat\gram\v1"
        $toc = readGramToc $dir
        $toc.Segments.Count | Should -Be 1
        $segment = openGramSegment "$dir\$($toc.Segments[0].Name)" $toc.Segments[0].Length
        $segment | Should -Not -BeNullOrEmpty
        try {
            $candidates = getGramCandidates $segment (getSearchGramValues $word)
            (@($candidates | Sort-Object | ForEach-Object { $segment.Names[$_] }) -join ";") | Should -Be $expected
            $segment.Ticks[0] | Should -Be 638000000000000001
            $segment.Sizes[2] | Should -Be 300
        } finally {
            closeGramSegments @($segment)
        }
    }
}