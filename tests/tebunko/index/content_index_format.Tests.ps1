# 検索用の本文インデックスのファイルの形式（tebunko\index\content_index_format.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "convertPlaceToContentIndexMeta / convertContentIndexMetaToPlace" -Tag Unit {
    It "今の場所の名前をメタ情報に分け、元の名前に戻せる" {
        $cases = @(
            @("見積.xlsx", "見積", "シート=見積|対象=本文"),
            @("見積.xlsx", "見積[図形]", "シート=見積|対象=図形"),
            @("見積.xlsx", "見積[コメント]", "シート=見積|対象=コメント"),
            @("見積.xlsx", "見積[ヘッダー・フッター]", "シート=見積|対象=ヘッダー・フッター"),
            @("見積.xlsx", "ページ001", "シート=ページ001|対象=本文"),
            @("議事録.docx", "ページ001", "ページ=1|対象=本文"),
            @("議事録.docx", "ページ012[図形]", "ページ=12|対象=図形"),
            @("議事録.docx", "ヘッダー・フッター", "部分=ヘッダー・フッター|対象=本文"),
            @("議事録.docx", "文書[コメント]", "部分=文書|対象=コメント"),
            @("議事録.doc", "脚注", "部分=脚注|対象=本文"),
            @("提案.pptx", "スライド001", "スライド=1|対象=本文"),
            @("提案.pptx", "スライド001_ノート", "スライド=1|対象=ノート"),
            @("提案.pptx", "スライド002（非表示）", "スライド=2|対象=本文|非表示=はい"),
            @("提案.pptx", "スライド002（非表示）[図形]", "スライド=2|対象=図形|非表示=はい"),
            @("提案.pptx", "ヘッダー・フッター", "部分=ヘッダー・フッター|対象=本文")
        )
        foreach ($c in $cases) {
            $meta = convertPlaceToContentIndexMeta $c[0] $c[1]
            (($meta.Keys | ForEach-Object { "$_=$($meta[$_])" }) -join "|") | Should -Be $c[2]
            convertContentIndexMetaToPlace $meta | Should -BeExactly $c[1]
        }
    }

    It "埋め込みの場所は、対象=埋め込みと番号をメタ情報に持ち、元の名前に戻せる" {
        foreach ($c in @(@("議事録.docx", "ページ003[埋め込み2]"), @("提案.pptx", "スライド001[埋め込み12]"), @("議事録.docx", "ヘッダー・フッター[埋め込み1]"))) {
            $meta = convertPlaceToContentIndexMeta $c[0] $c[1]
            $meta["対象"] | Should -Be "埋め込み"
            $meta["埋め込み"] | Should -Be ($c[1] -replace '^.*埋め込み(\d+)\]$', '$1')
            convertContentIndexMetaToPlace $meta | Should -BeExactly $c[1]
        }
    }

    It "種類の分からないファイルは、部分に場所の名前を持つ" {
        getContentIndexFileKind "メモ.pdf" | Should -Be ""
        $meta = convertPlaceToContentIndexMeta "メモ.pdf" "本文[図形]"
        (($meta.Keys | ForEach-Object { "$_=$($meta[$_])" }) -join "|") | Should -Be "部分=本文|対象=図形"
        convertContentIndexMetaToPlace $meta | Should -BeExactly "本文[図形]"
    }

    It "組み立て直すと同じにならない名前は、部分にそのまま持つ" {
        $meta = convertPlaceToContentIndexMeta "議事録.docx" "ページ1"
        $meta["部分"] | Should -Be "ページ1"
        convertContentIndexMetaToPlace $meta | Should -BeExactly "ページ1"
    }

    It "テキストの拡張子は「テキスト」、場所は「本文」だけ（部分=本文・対象=本文）で、そのまま元の名前に戻る" {
        getContentIndexFileKind "議事メモ.txt" | Should -Be "テキスト"
        $meta = convertPlaceToContentIndexMeta "議事メモ.txt" "本文"
        (($meta.Keys | ForEach-Object { "$_=$($meta[$_])" }) -join "|") | Should -Be "部分=本文|対象=本文"
        convertContentIndexMetaToPlace $meta | Should -BeExactly "本文"
    }
}

Describe "getContentIndexFileName / splitContentIndexBooksByExtension" -Tag Unit {
    It "本文インデックスのファイルの名前に元のファイルの拡張子（小文字）を入れる" {
        getContentIndexExtension "見積.XLSX" | Should -Be "xlsx"
    }

    It "元のファイルを拡張子ごとに分け、それぞれの中の順は変えない" {
        $books = @(@{ Name = "a.xlsx" }, @{ Name = "b.docx" }, @{ Name = "c.XLSX" }, @{ Name = "d.xlsm" })
        $groups = splitContentIndexBooksByExtension $books
        @($groups.Keys) -join "," | Should -Be "xlsx,docx,xlsm"
        @($groups["xlsx"] | ForEach-Object { $_.Name }) -join "," | Should -Be "a.xlsx,c.XLSX"
    }
}

Describe "getContentIndexFileName / readContentIndexFileName" -Tag Unit {
    It "名前に拡張子と 3 桁の番号を入れ、名前から取り出せる" {
        getContentIndexFileName "xlsx" | Should -Be "content_index.xlsx.001.tsv"
        getContentIndexFileName "docx" 12 | Should -Be "content_index.docx.012.tsv"
        $info = readContentIndexFileName "content_index.XLSX.002.tsv"
        $info.Extension | Should -Be "xlsx"
        $info.Part | Should -Be 2
        readContentIndexFileName "content_index.xlsx.tsv" | Should -Be $null
        readContentIndexFileName "見積.xlsx_S.tsv" | Should -Be $null
        # 前の版の名前（content_index に名前をそろえる前）は読まない
        readContentIndexFileName "content.xlsx.001.tsv" | Should -Be $null
    }
}

Describe "planContentIndexParts" -Tag Unit {
    BeforeAll {
        function newBook([string]$name, [int]$chars) {
            return @{ Name = $name; Block = ("x" * $chars) }
        }
        function describePlan($plan) {
            return (@($plan | ForEach-Object { "{0}.{1}:{2}:{3}" -f $_.Extension, $_.Part, (@($_.Books | ForEach-Object { $_.Name }) -join "+"), $(if ($_.Changed) { "書く" } else { "そのまま" }) }) -join " | ")
        }
    }

    It "新しい元のファイルは、上限に達するまで同じ本文インデックスのファイルに足し、達したら次の番号に足す" {
        # 1 冊 1,000 文字（2,000 バイト）。上限 4,000 バイトなら 2 冊で上限に達する
        $plan = planContentIndexParts @() @((newBook "c.xlsx" 1000), (newBook "a.xlsx" 1000), (newBook "b.xlsx" 1000), (newBook "d.docx" 10)) @() 4000
        describePlan $plan | Should -Be "docx.1:d.docx:書く | xlsx.1:a.xlsx+b.xlsx:書く | xlsx.2:c.xlsx:書く"
    }

    It "入れ替えは同じ位置で、外したものは除き、変わらない本文インデックスのファイルは書き直さない。新しいものは最後の番号に足す" {
        $parts = @(
            @{ Extension = "xlsx"; Part = 1; Books = @((newBook "a.xlsx" 10), (newBook "b.xlsx" 10)) },
            @{ Extension = "xlsx"; Part = 2; Books = @((newBook "c.xlsx" 10)) },
            @{ Extension = "docx"; Part = 1; Books = @((newBook "d.docx" 10)) }
        )
        $plan = planContentIndexParts $parts @((newBook "B.XLSX" 20), (newBook "e.xlsx" 10)) @("d.docx") 4000
        describePlan $plan | Should -Be "docx.1::書く | xlsx.1:a.xlsx+B.XLSX:書く | xlsx.2:c.xlsx+e.xlsx:書く"
        ($plan | Where-Object { $_.Extension -eq "xlsx" -and $_.Part -eq 1 }).Books[1].Block.Length | Should -Be 20
        describePlan (planContentIndexParts $parts @() @() 4000) | Should -Be "docx.1:d.docx:そのまま | xlsx.1:a.xlsx+b.xlsx:そのまま | xlsx.2:c.xlsx:そのまま"
    }

    It "最後の番号の本文インデックスのファイルが上限以上なら、次の番号の本文インデックスのファイルを作る（1 冊が上限を超えても、その 1 冊で 1 つ）" {
        $parts = @(@{ Extension = "xlsx"; Part = 3; Books = @((newBook "a.xlsx" 5000)) })
        describePlan (planContentIndexParts $parts @((newBook "b.xlsx" 5000)) @() 4000) | Should -Be "xlsx.3:a.xlsx:そのまま | xlsx.4:b.xlsx:書く"
    }
}

Describe "encodeContentIndexValue / decodeContentIndexValue" -Tag Unit {
    It "改行・制御文字・% を符号化し、タブはそのまま残して戻せる" {
        $value = "50%引き`n2行目`t" + [char]0x1E
        $encoded = encodeContentIndexValue $value
        $encoded | Should -Be "50%25引き%0A2行目`t%1E"
        decodeContentIndexValue $encoded | Should -BeExactly $value
    }
}

Describe "convertToContentIndexBody" -Tag Unit {
    It "改行を LF にそろえ、ReadLine と同じ行に分かれる形にする" {
        convertToContentIndexBody "a`r`n`r`nb`r`n" | Should -BeExactly "a`n`nb`n"
        convertToContentIndexBody "a`rb" | Should -BeExactly "a`nb`n"
        convertToContentIndexBody "" | Should -BeExactly ""
        convertToContentIndexBody "`r`n" | Should -BeExactly "`n"
    }

    It "区切りに使う U+001C〜U+001F を取り除く" {
        convertToContentIndexBody ("a" + [char]0x1E + "b") | Should -BeExactly "ab`n"
    }
}

Describe "convertToContentIndexText / readContentIndexPlaces" -Tag Unit {
    BeforeAll {
        $books = @(
            @{ Name = "見積.xlsx"; Places = @(@{ Place = "見積"; Text = "品名`t単価`r`n`r`nりんご`t100`r`n" }, @{ Place = "見積[図形]"; Text = "E2`t承認済み`r`n" }) },
            @{ Name = "議事録.docx"; Places = @(@{ Place = "ページ001"; Text = "方針`r`n" }, @{ Place = "空"; Text = "" }) }
        )
        $text = convertToContentIndexText $books
    }

    It "メタ情報の行と中身を並べる" {
        $mark = [string][char]0x1E
        $lines = $text.Split("`n")
        $lines[0] | Should -Be "$mark 版=1"
        $lines[1] | Should -Be "$mark ファイル名=見積.xlsx"
        $lines[2] | Should -Be "$mark 種類=Excel"
        $lines[3] | Should -Be "$mark シート=見積"
        $lines[4] | Should -Be "$mark 対象=本文"
        $lines[5] | Should -Be "品名`t単価"
        $lines[6] | Should -Be ""
    }

    It "場所ごとに、元のファイル名・場所の名前・中身の範囲を読み取る" {
        $places = readContentIndexPlaces $text
        $places.Count | Should -Be 4
        $places[0].Book | Should -Be "見積.xlsx"
        $places[0].Location | Should -Be "見積"
        $text.Substring($places[0].Start, $places[0].End - $places[0].Start) | Should -BeExactly "品名`t単価`n`nりんご`t100`n"
        $places[1].Location | Should -Be "見積[図形]"
        $places[2].Book | Should -Be "議事録.docx"
        $places[2].Location | Should -Be "ページ001"
        $places[3].Location | Should -Be "空"
        $places[3].End - $places[3].Start | Should -Be 0
    }

    It "元のファイルごとのまとまりに分け、そのまま並べ直すと元に戻る" {
        $blocks = splitContentIndexTextByBook $text
        @($blocks | ForEach-Object { $_.Name }) -join "," | Should -Be "見積.xlsx,議事録.docx"
        $blocks[0].Block.StartsWith([string][char]0x1E + " ファイル名=見積.xlsx`n") | Should -Be $true
        convertToContentIndexText $blocks | Should -BeExactly $text
        (splitContentIndexTextByBook "").Count | Should -Be 0
        { splitContentIndexTextByBook ([string][char]0x1E + " 版=2`n") } | Should -Throw
    }

    It "版が違う・無いときは例外にする" {
        { readContentIndexPlaces ([string][char]0x1E + " 版=2`n") } | Should -Throw
        { readContentIndexPlaces ([string][char]0x1E + " ファイル名=a.xlsx`n") } | Should -Throw
    }

    It "末尾に改行が無くても、最後のメタ情報の行まで読む" {
        $mark = [string][char]0x1E
        $places = readContentIndexPlaces "$mark 版=1`n$mark ファイル名=a.xlsx`n$mark シート=S`n$mark 対象=図形"
        $places.Count | Should -Be 1
        $places[0].Location | Should -Be "S[図形]"
    }

    It "<Book> の <Target> の場所は、行の先頭がセル番地か（CellPrefixed）= <Expected>" -ForEach @(
        @{ Book = "a.xlsx"; Target = "図形"; Expected = $true }
        @{ Book = "a.xlsm"; Target = "図形"; Expected = $true }
        @{ Book = "A.XLSX"; Target = "コメント"; Expected = $true }
        @{ Book = "a.xlsx"; Target = "コメント"; Expected = $true }
        @{ Book = "a.xlsx"; Target = "本文"; Expected = $false }
        @{ Book = "a.xlsx"; Target = "ヘッダー・フッター"; Expected = $false }
        @{ Book = "a.docx"; Target = "図形"; Expected = $false }
        @{ Book = "a.docx"; Target = "コメント"; Expected = $false }
        @{ Book = "a.pptx"; Target = "図形"; Expected = $false }
        @{ Book = "a.pptx"; Target = "コメント"; Expected = $false }
        @{ Book = "a.pptx"; Target = "ノート"; Expected = $false }
    ) {
        $mark = [string][char]0x1E
        $unit = if ($Book -match '\.xls') { "シート" } elseif ($Book -match '\.doc') { "ページ" } else { "スライド" }
        $value = if ($unit -eq "シート") { "S" } else { "1" }
        $places = readContentIndexPlaces "$mark 版=1`n$mark ファイル名=$Book`n$mark $unit=$value`n$mark 対象=$Target`nC2`t文字`n"
        $places.Count | Should -Be 1
        $places[0].CellPrefixed | Should -Be $Expected
    }

    It "Excel のヘッダー・フッターの場所（シート名 + 対象）を、場所の名前に戻して読む" {
        $mark = [string][char]0x1E
        $places = readContentIndexPlaces "$mark 版=1`n$mark ファイル名=a.xlsx`n$mark シート=S`n$mark 対象=ヘッダー・フッター`n中央`n"
        $places.Count | Should -Be 1
        $places[0].Location | Should -Be "S[ヘッダー・フッター]"
    }
}
