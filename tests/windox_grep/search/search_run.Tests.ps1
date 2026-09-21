# インデックスの検索（windox_grep\search\search_run.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "toResultLine" -Tag Io {
    It "ブック名・シート名・行番号・該当行をタブ区切りにする" {
        toResultLine "book.xlsx" "Sheet1" 12 "時刻`t12:34:56" | Should Be "book.xlsx`tSheet1`t12`t時刻`t12:34:56"
    }

    It "Excelのセル内改行を改行に戻す" {
        toResultLine "book.xlsx" "Sheet1" 3 "a`t`"1行目${cellNewLine}2行目`"`tb" | Should Be "book.xlsx`tSheet1`t3`ta`t`"1行目`n2行目`"`tb"
    }

    It "シート名のタブ・改行はスペースにする（列・行が分かれないようにする）" {
        toResultLine "book.xlsx" "タブ`tあり" 1 "x" | Should Be "book.xlsx`tタブ あり`t1`tx"
        toResultLine "book.xlsx" "改行`nあり" 2 "y" | Should Be "book.xlsx`t改行 あり`t2`ty"
    }

    It 'Word・PowerPointの行は、" で始まるセルだけを " で囲む' {
        toResultLine "doc.docx" "ページ001" 1 "`"引用`"と言った" | Should Be "doc.docx`tページ001`t1`t`"`"`"引用`"`"と言った`""
        toResultLine "doc.docx" "ページ001" 2 "彼は`"引用`"と言った" | Should Be "doc.docx`tページ001`t2`t彼は`"引用`"と言った"
        toResultLine "doc.docx" "ページ001" 3 "表`t`"見出し`"`tx" | Should Be "doc.docx`tページ001`t3`t表`t`"`"`"見出し`"`"`"`tx"
    }
}

Describe "toResultHeader" -Tag Io {
    It "ファイル名・場所・行と、列名を並べる" {
        toResultHeader 3 | Should Be "ファイル名`t場所`t行`tA`tB`tC"
        toResultHeader 0 | Should Be "ファイル名`t場所`t行"
    }
}

Describe "getIndexTsvFiles / testIndexExists / getIndexSummary" -Tag Io {
    $index = Join-Path $TestDrive "index[1]"
    $other = Join-Path $TestDrive "other"
    newTsv "$index\b.xlsx_S.tsv" @("b")
    newTsv "$index\sub\a.xlsx_S.tsv" @("a")
    newTsv "$other\c.docx_ページ001.tsv" @("c")
    (Get-Item -LiteralPath "$other\c.docx_ページ001.tsv").LastWriteTime = [datetime]"2030-01-02 03:04:05"
    $missing = Join-Path $TestDrive "missing"

    It "数え上げの途中の件数を知らせる（画面が「確認中… N 件」を出すため）" {
        $many = Join-Path $TestDrive "many"
        for ($i = 0; $i -lt 5; $i++) {
            newTsv "$many\book$i.xlsx\S.tsv" @("x")
        }
        $counts = New-Object System.Collections.Generic.List[int]
        # 既定は 2000 件ごとのため、ここでは呼ばれないことを確かめる（件数の通知が検索を遅くしない）
        $result = getIndexTsvFiles @($many) { param ($count) $counts.Add($count) }
        $result.Files.Count | Should Be 5
        $counts.Count | Should Be 0
    }

    It "フォルダごとの件数と、インデックスフォルダからの相対パスを返す" {
        $result = getIndexTsvFiles @($index, $missing, $other)
        $result.Folders.Count | Should Be 3
        $result.Folders[0].Exists | Should Be $true
        $result.Folders[0].Count | Should Be 2
        $result.Folders[1].Exists | Should Be $false
        $result.Files.Count | Should Be 3
        $result.Files["$index\sub\a.xlsx_S.tsv"].RelPath | Should Be "sub\a.xlsx_S.tsv"
    }

    It "入れ子のフォルダを指定しても同じTSVを重複させない" {
        $result = getIndexTsvFiles @($index, "$index\sub")
        $result.Files.Count | Should Be 2
    }

    It "インデックスの中のフォルダだけを指定でき、相対パスはインデックスのフォルダから求める" {
        $result = getIndexTsvFiles @(@{ Root = $index; RelPath = "sub"; Recurse = $true })
        $result.Files.Count | Should Be 1
        $result.Files["$index\sub\a.xlsx_S.tsv"].Root | Should Be $index
        $result.Files["$index\sub\a.xlsx_S.tsv"].RelPath | Should Be "sub\a.xlsx_S.tsv"
        $result.Folders[0].Path | Should Be "$index\sub"
    }

    It "Recurse が false ならフォルダ直下のファイルだけ" {
        $result = getIndexTsvFiles @(@{ Root = $index; RelPath = ""; Recurse = $false })
        $result.Files.Count | Should Be 1
        $result.Files.Contains("$index\b.xlsx_S.tsv") | Should Be $true
    }

    It "Recurse が false でも、元のファイル名のフォルダの中のTSVは直下のファイルとして数える" {
        $dir = Join-Path $TestDrive "index_direct"
        newTsv "$dir\A社.xlsx\$(toIndexFileName "Sheet1")" @("a")
        newTsv "$dir\sub\B社.xlsx\$(toIndexFileName "Sheet1")" @("b")

        $result = getIndexTsvFiles @(@{ Root = $dir; RelPath = ""; Recurse = $false })
        $result.Files.Count | Should Be 1
        $result.Files.Contains("$dir\A社.xlsx\Sheet1.tsv") | Should Be $true
    }

    It "存在しない中のフォルダは Exists が false" {
        $result = getIndexTsvFiles @(@{ Root = $index; RelPath = "なし"; Recurse = $true }, @{ Root = $missing; RelPath = "sub"; Recurse = $true })
        $result.Folders[0].Exists | Should Be $false
        $result.Folders[1].Exists | Should Be $false
        $result.Files.Count | Should Be 0
    }

    It "TSVの有無を判定する" {
        testIndexExists @($missing, $other) | Should Be $true
        testIndexExists @($missing) | Should Be $false
    }

    It "TSVの件数・最新の更新日時・存在しないフォルダを返す" {
        $summary = getIndexSummary @($index, $other, $missing, $index)
        $summary.Count | Should Be 3
        $summary.LastWrite | Should Be ([datetime]"2030-01-02 03:04:05")
        $summary.Missing.Count | Should Be 1
    }
}

Describe "searchIndex（今の形式: <ファイル名>\<場所>.tsv）" -Tag Io {
    $index = Join-Path $TestDrive "search_dir"
    # ファイル名のフォルダの中に、場所を名前にしたTSVを置く
    newTsv "$index\営業\A社.xlsx\$(toIndexFileName "見積_2024")" @("りんご`t100")
    newTsv "$index\A社.xlsx_old.xlsx\$(toIndexFileName "Sheet1")" @("りんご`t200")
    newTsv "$index\営業\文書.docx\$(toIndexFileName "ページ001")" @("りんご (株)")
    $files = (getIndexTsvFiles @($index)).Files

    It "フォルダ名をファイル名、TSVの名前を場所として返す（相対フォルダにファイル名のフォルダは入れない）" {
        $hits = @((searchIndex "りんご" $files $true).Hits | Sort-Object Book)
        $hits.Count | Should Be 3
        $hits[0].Book | Should Be "A社.xlsx"
        $hits[0].Location | Should Be "見積_2024"
        $hits[0].RelDir | Should Be "営業"
        $hits[0].LineNumber | Should Be 1
        # ファイル名に .xlsx_ を含んでも、フォルダ名のまま返る
        $hits[1].Book | Should Be "A社.xlsx_old.xlsx"
        $hits[1].Location | Should Be "Sheet1"
        $hits[1].RelDir | Should Be ""
        $hits[2].Book | Should Be "文書.docx"
        $hits[2].Location | Should Be "ページ001"
    }

    It "対象ファイルの絞り込みは、フォルダ名（元のファイル名）で判定する" {
        $result = searchIndex "りんご" $files $true 0 100 $null $null $false "*.docx"
        $result.Hits.Count | Should Be 1
        $result.Hits[0].Book | Should Be "文書.docx"
    }

    It "検索結果の行は、フォルダ名のファイル名と場所で組み立てる" {
        $hit = @((searchIndex "りんご" $files $true 0 100 $null $null $false "A社.xlsx_old.xlsx").Hits)[0]
        (toSearchResultLines @($hit)).Lines[0] | Should Be "A社.xlsx_old.xlsx`tSheet1`t1`tりんご`t200"
    }
}

Describe "searchIndex" -Tag Io {
    $index = Join-Path $TestDrive "search"
    newTsv "$index\A社.xlsx_Sheet1.tsv" @("見積先：`t(株)山田商事", "", "株式会社`t1.5", "ABC`t105")
    newTsv "$index\sub\文書.docx_ページ001.tsv" @("りんご (株) abc")
    $files = (getIndexTsvFiles @($index)).Files

    It "文字どおりに検索すると (株) は記号のまま探す" {
        $result = searchIndex "(株)" $files $true
        $result.Hits.Count | Should Be 2
        $result.SimpleMatch | Should Be $true
    }

    It "正規表現として検索すると (株) は「株」に一致する" {
        $result = searchIndex "(株)" $files $false
        $result.Hits.Count | Should Be 3
        $result.SimpleMatch | Should Be $false
    }

    It "正規表現として不正なワードは文字どおりに検索する" {
        $result = searchIndex "(" $files $false
        $result.SimpleMatch | Should Be $true
        $result.Hits.Count | Should Be 2
    }

    It "大文字・小文字を区別しない" {
        (searchIndex "abc" $files $true).Hits.Count | Should Be 2
    }

    It "ファイル名・場所・行番号・相対フォルダを返す（空行も行番号に数える）" {
        $hit = @((searchIndex "1.5" $files $true).Hits)[0]
        $hit.Book | Should Be "A社.xlsx"
        $hit.Location | Should Be "Sheet1"
        $hit.LineNumber | Should Be 3
        $hit.RelDir | Should Be ""
        $hit.Root | Should Be $index
        @((searchIndex "りんご" $files $true).Hits)[0].RelDir | Should Be "sub"
    }

    It "上限を超えたら打ち切る" {
        $result = searchIndex "株" $files $true 2
        $result.Hits.Count | Should Be 2
        $result.Truncated | Should Be $true
    }

    It "上限ちょうどなら打ち切りにしない" {
        $result = searchIndex "株" $files $true 3
        $result.Hits.Count | Should Be 3
        $result.Truncated | Should Be $false
    }

    It "進み具合を知らせ、中止できる" {
        $script:progress = @()
        $result = searchIndex "株" $files $true 0 1 { param($done, $total, $newHits) $script:progress += "$done/$total" } { $script:progress.Count -ge 1 }
        $result.Cancelled | Should Be $true
        $script:progress.Count | Should Be 1
        $script:progress[0] | Should Be "1/2"
    }

    It "結果を @() で配列にできる" {
        @((searchIndex "株" $files $true).Hits).Count | Should Be 3
    }

    It "大文字と小文字を区別できる" {
        @((searchIndex "abc" $files $true -caseSensitive $true).Hits).Count | Should Be 1
        @((searchIndex "ABC" $files $true -caseSensitive $true).Hits)[0].LineNumber | Should Be 4
    }

    It "対象ファイルで、元のファイル名が一致するTSVだけを検索する" {
        $result = searchIndex "株" $files $true -fileFilter "*.docx"
        $result.Total | Should Be 1
        @($result.Hits).Count | Should Be 1
        @((searchIndex "株" $files $true -fileFilter "!*.docx").Hits).Count | Should Be 2
    }

    It "検索の途中で読めなくなったTSV（変換中に削除された等）は飛ばす" {
        $dir = Join-Path $TestDrive "deleted"
        newTsv "$dir\a.xlsx_S.tsv" @("株")
        newTsv "$dir\b.xlsx_S.tsv" @("株")
        $targets = (getIndexTsvFiles @($dir)).Files
        Remove-Item -LiteralPath "$dir\a.xlsx_S.tsv"
        $hits = @((searchIndex "株" $targets $true).Hits)
        $hits.Count | Should Be 1
        $hits[0].Book | Should Be "b.xlsx"
    }
}

Describe "toSearchResultLines / writeSearchResult" -Tag Io {
    $index = Join-Path $TestDrive "result"
    newTsv "$index\x\A社.xlsx_Sheet1.tsv" @("a`t`"りんご${cellNewLine}みかん`"`tc")
    newTsv "$index\文書.docx_ページ001.tsv" @("`"引用`"で始まる りんご")
    $files = (getIndexTsvFiles @($index)).Files
    $hits = @((searchIndex "りんご" $files $true).Hits)

    It "相対フォルダ付きのファイル名・場所・行番号・該当行にし、見出しに最大セル数分の列名を付ける" {
        $result = toSearchResultLines $hits
        $result.Header | Should Be "ファイル名`t場所`t行`tA`tB`tC"
        $result.Lines[0] | Should Be "x\A社.xlsx`tSheet1`t1`ta`t`"りんご`nみかん`"`tc"
        $result.Lines[1] | Should Be "文書.docx`tページ001`t1`t`"`"`"引用`"`"で始まる りんご`""
    }

    It "検索結果ファイルの形式（02_検索.md）で書き出す" {
        $writer = New-Object System.IO.StringWriter
        writeSearchResult $writer "りんご" $hits
        $lines = $writer.ToString() -split "`r`n"
        $lines[0] | Should Be "【検索文字列　りんご】 2 件"
        $lines[1] | Should Be "ファイル名`t場所`t行`tA`tB`tC"
        $lines[4] | Should Be ""
    }

    It "0 件なら見出しの2行と空行だけ書き出す" {
        $writer = New-Object System.IO.StringWriter
        writeSearchResult $writer "無い" @()
        $writer.ToString() | Should Be "【検索文字列　無い】 0 件`r`nファイル名`t場所`t行`r`n`r`n"
    }
}
