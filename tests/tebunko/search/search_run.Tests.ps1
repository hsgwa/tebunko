# インデックスの検索（tebunko\search\search_run.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "toResultLine" -Tag Io {
    It "ブック名・場所・種別・行番号・該当行をタブ区切りにする（場所・種別は画面と同じ表示）" {
        toResultLine "book.xlsx" "Sheet1" 12 "時刻`t12:34:56" | Should -Be "book.xlsx`t[シート]Sheet1`tセル`t12`t時刻`t12:34:56"
    }

    It "Excelのセル内改行を改行に戻す" {
        toResultLine "book.xlsx" "Sheet1" 3 "a`t`"1行目${cellNewLine}2行目`"`tb" | Should -Be "book.xlsx`t[シート]Sheet1`tセル`t3`ta`t`"1行目`n2行目`"`tb"
    }

    It "シート名のタブ・改行はスペースにする（列・行が分かれないようにする）" {
        toResultLine "book.xlsx" "タブ`tあり" 1 "x" | Should -Be "book.xlsx`t[シート]タブ あり`tセル`t1`tx"
        toResultLine "book.xlsx" "改行`nあり" 2 "y" | Should -Be "book.xlsx`t[シート]改行 あり`tセル`t2`ty"
    }

    It 'Word・PowerPointの行は、" で始まるセルだけを " で囲む' {
        toResultLine "doc.docx" "ページ001" 1 "`"引用`"と言った" | Should -Be "doc.docx`t1 ページ（目安）`t本文`t1`t`"`"`"引用`"`"と言った`""
        toResultLine "doc.docx" "ページ001" 2 "彼は`"引用`"と言った" | Should -Be "doc.docx`t1 ページ（目安）`t本文`t2`t彼は`"引用`"と言った"
        toResultLine "doc.docx" "ページ001" 3 "表`t`"見出し`"`tx" | Should -Be "doc.docx`t1 ページ（目安）`t本文`t3`t表`t`"`"`"見出し`"`"`"`tx"
    }
}

Describe "testIndexExists / getIndexSummary" -Tag Io {
    BeforeAll {
        $index = Join-Path $TestDrive "index[1]"
        $other = Join-Path $TestDrive "other"
        foreach ($item in @(@{ Folder = "$index"; Name = "b.xlsx" }, @{ Folder = "$index\sub"; Name = "a.xlsx" }, @{ Folder = "$other"; Name = "c.docx" })) {
            [void][System.IO.Directory]::CreateDirectory($item.Folder)
            writeContentIndexFile "$($item.Folder)\$(getContentIndexFileName (getContentIndexExtension $item.Name))" (convertToContentIndexText @(@{ Name = $item.Name; Places = @(@{ Place = "S"; Text = "x" }) }))
        }
        (Get-Item -LiteralPath "$other\content_index.docx.001.tsv").LastWriteTime = [datetime]"2030-01-02 03:04:05"
        # 本文インデックスにする前の TSV（インデックス作成の途中）は数えない
        newTsv "$other\d.xlsx\S.tsv" @("d")
        $missing = Join-Path $TestDrive "missing"
    }

    It "本文インデックスのファイルの有無を判定する" {
        testIndexExists @($missing, $other) | Should -Be $true
        testIndexExists @($missing) | Should -Be $false
    }

    It "有無を調べた直後に、下のフォルダを移動できる（起動直後の上書きのインポート）" {
        # 1 件見つけたところでやめても、下のフォルダを掴んだまま残さないこと（残るとインポートの上書きで移動できない）
        $root = Join-Path $TestDrive "exists_handle"
        foreach ($name in @("a.xlsx", "b.xlsx")) {
            [void][System.IO.Directory]::CreateDirectory("$root\営業\見積")
            writeContentIndexFile "$root\営業\見積\$(getContentIndexFileName (getContentIndexExtension $name))" (convertToContentIndexText @(@{ Name = $name; Places = @(@{ Place = "S"; Text = "x" }) }))
        }
        newTsv "$root\営業\見積\c.xlsx\S.tsv" @("c")

        testIndexExists @($root) | Should -Be $true

        { [System.IO.Directory]::Move("$root\営業", "$TestDrive\exists_handle_moved") } | Should -Not -Throw
    }

    It "本文インデックスのファイルの件数・最新の更新日時・存在しないフォルダを返す" {
        $summary = getIndexSummary @($index, $other, $missing, $index)
        $summary.Count | Should -Be 3
        $summary.LastWrite | Should -Be ([datetime]"2030-01-02 03:04:05")
        $summary.Missing.Count | Should -Be 1
    }

    It "本文インデックスのファイルの無いフォルダ・存在しないフォルダだけなら、なし・件数 0・更新日時なし" {
        $empty = Join-Path $TestDrive "empty_index"
        New-Item -ItemType Directory -Path "$empty\sub" -Force | Out-Null
        newTsv "$empty\sub\e.xlsx\S.tsv" @("e")
        testIndexExists @($empty, $missing) | Should -Be $false
        testIndexExists @() | Should -Be $false
        $summary = getIndexSummary @($empty)
        $summary.Count | Should -Be 0
        $summary.LastWrite | Should -Be $null
        @($summary.Missing).Count | Should -Be 0
    }
}

Describe "検索結果の行（本文インデックスのファイルのヒット）" -Tag Io {
    It "ファイル名に .xlsx_ を含んでも、元のファイル名と場所で組み立てる" {
        $index = Join-Path $TestDrive "search_dir"
        newTsv "$index\A社.xlsx_old.xlsx\$(toIndexFileName "Sheet1")" @("りんご`t200")
        foreach ($folder in (findIndexFoldersWithBooks $index)) { [void](updateFolderContentIndex $folder) }
        $hit = (expandHitRuns (searchContentIndex "りんご" (getContentIndexFiles @($index)).ContentIndexFiles $true).Runs)[0]
        $hit.Book | Should -Be "A社.xlsx_old.xlsx"
        (toSearchResultLines @($hit)).Lines[0] | Should -Be "A社.xlsx_old.xlsx`t[シート]Sheet1`tセル`t1`tりんご`t200"
    }

    It "Excel の図形の場所（名前に [ ] を含む）も検索でき、セル内改行を戻して出力する" {
        $objectIndex = Join-Path $TestDrive "search_object"
        newTsv "$objectIndex\[確定]見積.xlsx\$(toIndexFileName "見積[図形]")" @("F2`t`"納期は$([char]0x2028)別途`"")
        foreach ($folder in (findIndexFoldersWithBooks $objectIndex)) { [void](updateFolderContentIndex $folder) }
        $hit = (expandHitRuns (searchContentIndex "納期" (getContentIndexFiles @($objectIndex)).ContentIndexFiles $true).Runs)[0]
        $hit.Book | Should -Be "[確定]見積.xlsx"
        $hit.Location | Should -Be "見積[図形]"
        (toSearchResultLines @($hit)).Lines[0] | Should -Be "[確定]見積.xlsx`t[シート]見積`t図形`t1`tF2`t`"納期は`n別途`""
    }
}
Describe "getSearchResultLine / getSearchResultColumnCount" -Tag Unit {
    It "相対フォルダがあればファイル名の前に付け、無ければ付けない" -ForEach @(
        @{ RelDir = ""; Expected = "A社.xlsx`t[シート]Sheet1`tセル`t3`tりんご" }
        @{ RelDir = "営業部\2024"; Expected = "営業部\2024\A社.xlsx`t[シート]Sheet1`tセル`t3`tりんご" }
    ) {
        getSearchResultLine $RelDir "A社.xlsx" "Sheet1" 3 "りんご" | Should -Be $Expected
    }

    It "列の数は、ファイル名・場所・種別・行番号の 4 列を除いたセルの数" -ForEach @(
        @{ Line = "a`tb`tセル`t1`tx"; Expected = 1 }
        @{ Line = "a`tb`tセル`t1`tx`ty`tz"; Expected = 3 }
    ) {
        getSearchResultColumnCount $Line | Should -Be $Expected
    }
}

Describe "toSearchResultLines" -Tag Io {
    BeforeAll {
        $index = Join-Path $TestDrive "result"
        newTsv "$index\x\A社.xlsx\Sheet1.tsv" @("a`t`"りんご${cellNewLine}みかん`"`tc")
        newTsv "$index\文書.docx\ページ001.tsv" @("`"引用`"で始まる りんご")
        foreach ($folder in (findIndexFoldersWithBooks $index)) { [void](updateFolderContentIndex $folder) }
        $hits = @((expandHitRuns (searchContentIndex "りんご" (getContentIndexFiles @($index)).ContentIndexFiles $true).Runs) | Sort-Object RelDir -Descending)
    }

    It "相対フォルダ付きのファイル名・場所・行番号・該当行にし、見出しに最大セル数分の列名を付ける" {
        $result = toSearchResultLines $hits
        $result.Header | Should -Be "ファイル名`t場所`t種別`t行`tA`tB`tC"
        $result.Lines[0] | Should -Be "x\A社.xlsx`t[シート]Sheet1`tセル`t1`ta`t`"りんご`nみかん`"`tc"
        $result.Lines[1] | Should -Be "文書.docx`t1 ページ（目安）`t本文`t1`t`"`"`"引用`"`"で始まる りんご`""
    }
}

Describe "toResultCells" -Tag Unit {
    BeforeDiscovery {
        $nl = [string][char]0x2028
        $xl = "a`t`"1行目${nl}2行目`"`tc"
        $wd = "表`t`"見出し`"`tx"
        $cases = @(
            @{ name = "xlsx: Excel の行は セルを ` │ ` でつなぎ 1 セルにする"; Format = "xlsx"; Book = "A.xlsx"; Line = $xl; Rest = @("a │ 1行目`n2行目 │ c") }
            @{ name = "csv: Excel の行は 囲みを外し セルに分ける"; Format = "csv"; Book = "A.xlsx"; Line = $xl; Rest = @("a", "1行目`n2行目", "c") }
            @{ name = "tsv: Excel の行は csv と同じ列になる"; Format = "tsv"; Book = "A.xlsx"; Line = $xl; Rest = @("a", "1行目`n2行目", "c") }
            @{ name = "xlsx: Word の行は タブを ` │ ` にする"; Format = "xlsx"; Book = "d.docx"; Line = $wd; Rest = @("表 │ `"見出し`" │ x") }
            @{ name = "csv: Word の行は タブで分けるだけ（`" を加工しない）"; Format = "csv"; Book = "d.docx"; Line = $wd; Rest = @("表", "`"見出し`"", "x") }
            @{ name = "tsv: Word の行は csv と同じ"; Format = "tsv"; Book = "d.docx"; Line = $wd; Rest = @("表", "`"見出し`"", "x") }
            @{ name = "csv: セルの中の CR は LF にそろえる"; Format = "csv"; Book = "d.docx"; Line = "a`r`nb`rc"; Rest = @("a`nb`nc") }
            @{ name = "xlsx: セルの中の CR は LF にそろえる"; Format = "xlsx"; Book = "d.docx"; Line = "a`r`nb`rc"; Rest = @("a`nb`nc") }
        )
    }

    It "<name>" -ForEach $cases {
        $described = describePlace $Book "S"
        $cells = toResultCells "" $Book "S" 3 $Line $Format
        $expected = @($Book, $described.Place, $described.Kind, "3") + $Rest
        $cells.Count | Should -Be $expected.Count
        for ($i = 0; $i -lt $expected.Count; $i++) { $cells[$i] | Should -Be $expected[$i] }
    }

    It "相対フォルダを付け、ファイル名・相対フォルダ・場所の制御文字はスペースにする（<_>）" -ForEach @("xlsx", "csv", "tsv") {
        $cells = toResultCells "営`t業\x`r`n" "A`tB.xlsx" "シ`tート`n名" 5 "x" $_
        $cells[0] | Should -Be "営 業\x  \A B.xlsx"
        $cells[1] | Should -Not -Match '[\x00-\x1F]'
        $cells[3] | Should -Be "5"
    }
}
