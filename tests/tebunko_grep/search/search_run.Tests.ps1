# インデックスの検索（tebunko_grep\search\search_run.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "toResultLine" -Tag Io {
    It "ブック名・場所・種別・行番号・該当行をタブ区切りにする（場所・種別は画面と同じ表示）" {
        toResultLine "book.xlsx" "Sheet1" 12 "時刻`t12:34:56" | Should Be "book.xlsx`t[シート] Sheet1`tセル`t12`t時刻`t12:34:56"
    }

    It "Excelのセル内改行を改行に戻す" {
        toResultLine "book.xlsx" "Sheet1" 3 "a`t`"1行目${cellNewLine}2行目`"`tb" | Should Be "book.xlsx`t[シート] Sheet1`tセル`t3`ta`t`"1行目`n2行目`"`tb"
    }

    It "シート名のタブ・改行はスペースにする（列・行が分かれないようにする）" {
        toResultLine "book.xlsx" "タブ`tあり" 1 "x" | Should Be "book.xlsx`t[シート] タブ あり`tセル`t1`tx"
        toResultLine "book.xlsx" "改行`nあり" 2 "y" | Should Be "book.xlsx`t[シート] 改行 あり`tセル`t2`ty"
    }

    It 'Word・PowerPointの行は、" で始まるセルだけを " で囲む' {
        toResultLine "doc.docx" "ページ001" 1 "`"引用`"と言った" | Should Be "doc.docx`t[ページ] 1（目安）`t本文`t1`t`"`"`"引用`"`"と言った`""
        toResultLine "doc.docx" "ページ001" 2 "彼は`"引用`"と言った" | Should Be "doc.docx`t[ページ] 1（目安）`t本文`t2`t彼は`"引用`"と言った"
        toResultLine "doc.docx" "ページ001" 3 "表`t`"見出し`"`tx" | Should Be "doc.docx`t[ページ] 1（目安）`t本文`t3`t表`t`"`"`"見出し`"`"`"`tx"
    }
}

Describe "toResultHeader" -Tag Io {
    It "ファイル名・場所・種別・行と、列名を並べる" {
        toResultHeader 3 | Should Be "ファイル名`t場所`t種別`t行`tA`tB`tC"
        toResultHeader 0 | Should Be "ファイル名`t場所`t種別`t行"
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

    It "TSV の無いフォルダ・存在しないフォルダだけなら、TSV なし・件数 0・更新日時なし" {
        $empty = Join-Path $TestDrive "empty_index"
        New-Item -ItemType Directory -Path "$empty\sub" -Force | Out-Null
        testIndexExists @($empty, $missing) | Should Be $false
        testIndexExists @() | Should Be $false
        $summary = getIndexSummary @($empty)
        $summary.Count | Should Be 0
        $summary.LastWrite | Should Be $null
        @($summary.Missing).Count | Should Be 0
        $result = getIndexTsvFiles @($empty, $missing)
        $result.Files.Count | Should Be 0
        ($result.Folders | ForEach-Object { "$($_.Exists):$($_.Count)" }) -join "," | Should Be "True:0,False:0"
    }

    It "2000 件ごとに、数えた件数を知らせる" {
        $many = Join-Path $TestDrive "many2000\book.xlsx"
        [void][System.IO.Directory]::CreateDirectory($many)
        for ($i = 0; $i -lt 2001; $i++) {
            [System.IO.File]::WriteAllText("$many\S$i.tsv", "x")
        }
        $counts = New-Object System.Collections.Generic.List[int]
        $result = getIndexTsvFiles @((Split-Path $many -Parent)) { param ($count) $counts.Add($count) }
        $result.Files.Count | Should Be 2001
        ($counts -join ",") | Should Be "2000"
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
        (toSearchResultLines @($hit)).Lines[0] | Should Be "A社.xlsx_old.xlsx`t[シート] Sheet1`tセル`t1`tりんご`t200"
    }

    It "Excel の図形・コメントの場所（名前に [ ] を含む）も検索でき、セル内改行を戻して出力する" {
        $objectIndex = Join-Path $TestDrive "search_object"
        newTsv "$objectIndex\[確定]見積.xlsx\$(toIndexFileName "見積[図形]")" @("F2`t`"納期は$([char]0x2028)別途`"")
        $hit = @((searchIndex "納期" (getIndexTsvFiles @($objectIndex)).Files $true).Hits)[0]
        $hit.Book | Should Be "[確定]見積.xlsx"
        $hit.Location | Should Be "見積[図形]"
        (toSearchResultLines @($hit)).Lines[0] | Should Be "[確定]見積.xlsx`t[シート] 見積`t図形`t1`tF2`t`"納期は`n別途`""
    }

    It "図形・コメントを検索しない指定では、その場所の TSV を検索しない（件数にも入れない）" {
        $mixedIndex = Join-Path $TestDrive "search_object_option"
        newTsv "$mixedIndex\見積.xlsx\$(toIndexFileName "見積")" @("納期`t5日")
        newTsv "$mixedIndex\見積.xlsx\$(toIndexFileName "見積[図形]")" @("F2`t納期は別途")
        newTsv "$mixedIndex\見積.xlsx\$(toIndexFileName "見積[コメント]")" @("C2`t納期を確認")
        $mixedFiles = (getIndexTsvFiles @($mixedIndex)).Files

        $all = searchIndex "納期" $mixedFiles $true
        $all.Hits.Count | Should Be 3
        $noShapes = searchIndex "納期" $mixedFiles $true -includeShapes $false
        @($noShapes.Hits | ForEach-Object { $_.Location } | Sort-Object) -join "|" | Should Be "見積|見積[コメント]"
        $noShapes.Total | Should Be 2
        $cellsOnly = searchIndex "納期" $mixedFiles $true -includeShapes $false -includeComments $false
        @($cellsOnly.Hits | ForEach-Object { $_.Location }) -join "|" | Should Be "見積"
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

    It "検索の途中で読めなくなったTSV（インデックス作成中に削除された等）は飛ばす" {
        $dir = Join-Path $TestDrive "deleted"
        newTsv "$dir\a.xlsx_S.tsv" @("株")
        newTsv "$dir\b.xlsx_S.tsv" @("株")
        $targets = (getIndexTsvFiles @($dir)).Files
        Remove-Item -LiteralPath "$dir\a.xlsx_S.tsv"
        $hits = @((searchIndex "株" $targets $true).Hits)
        $hits.Count | Should Be 1
        $hits[0].Book | Should Be "b.xlsx"
    }

    It "上限 0 は上限なし。対象の TSV が無ければ 0 件（打ち切り・中止にしない）" {
        $all = searchIndex "株" $files $true 0
        $all.Hits.Count | Should Be 3
        $all.Truncated | Should Be $false
        $none = searchIndex "株" @{} $true
        $none.Total | Should Be 0
        $none.Hits.Count | Should Be 0
        $none.Truncated | Should Be $false
        $none.Cancelled | Should Be $false
    }

    It "対象ファイルに当たる TSV が無ければ、Total は 0" {
        $result = searchIndex "株" $files $true -fileFilter "*.pptx"
        $result.Total | Should Be 0
        $result.Hits.Count | Should Be 0
    }
}

Describe "searchIndex（正規表現の照合の時間切れ）" -Tag Io {
    # 入れ子の繰り返し（(a+)+$）は、一致しない文字列で照合の時間が指数的に増える
    $dir = Join-Path $TestDrive "timeout"
    for ($i = 0; $i -lt 3; $i++) {
        newTsv "$dir\book$i.xlsx\S.tsv" @(("a" * 40) + "!")
    }
    $files = (getIndexTsvFiles @($dir)).Files
    $message = "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"

    It "時間切れなら、分かるメッセージで検索を中止する" {
        $regexTimeout = [timespan]::FromMilliseconds(1)
        { searchIndex "(a+)+$" $files $false -workerCount 1 } | Should Throw $message
    }

    It "並列に検索していても、時間切れなら同じメッセージで中止する" {
        $regexTimeout = [timespan]::FromMilliseconds(1)
        { searchIndex "(a+)+$" $files $false 0 1 -workerCount 2 } | Should Throw $message
    }
}

Describe "searchTsvFiles（全文を読まない大きさの TSV）" -Tag Io {
    $dir = Join-Path $TestDrive "large"
    newTsv "$dir\a.xlsx\S.tsv" @("見積 1", "x", "見積 2")
    newTsv "$dir\b.docx\ページ001.tsv" @("見積 3")
    $targets = newTsvFiles (getIndexTsvFiles @($dir)).Files

    # searchTsvFiles は List を 1 つのまま返すため、1 件ずつ "ファイル名:行番号" にする
    function formatBookLines {
        param ($hits)
        return (@(foreach ($hit in $hits) { "$($hit.Book):$($hit.LineNumber)" }) -join ",")
    }

    It "上限より大きい TSV は 1 行ずつ照合し、結果・行番号は同じ" {
        $search = newSearchRegex "見積" $true
        $expected = formatBookLines (searchTsvFiles $targets 0 $targets.Count $search.Regex -1 $search.TextRegex $search.ScanMode)
        $expected | Should Be "a.xlsx:1,a.xlsx:3,b.docx:1"

        $searchWholeFileMax = 4
        $cache = newTsvTextCache
        formatBookLines (searchTsvFiles $targets 0 $targets.Count $search.Regex -1 $search.TextRegex $search.ScanMode $cache) |
            Should Be $expected
        # 大きい TSV は内容を残さない
        $cache.Texts.Count | Should Be 0
    }
}

Describe "searchTsvFiles（全文への照合と 1 行ずつの照合の結果が同じ）" -Tag Io {
    # 改行の種類（CRLF・LF・CR）、末尾の改行の有無、空行、空のファイル、BOM の無いファイルを混ぜる
    $dir = Join-Path $TestDrive "modes"
    [void][System.IO.Directory]::CreateDirectory("$dir\a.xlsx")
    [void][System.IO.Directory]::CreateDirectory("$dir\b.docx")
    [void][System.IO.Directory]::CreateDirectory("$dir\c.pptx")
    [void][System.IO.Directory]::CreateDirectory("$dir\d.xlsx")
    [System.IO.File]::WriteAllText("$dir\a.xlsx\S.tsv", "見積`t(株)山田商事`r`n`r`nabc 見積`r`n確定`tABC`r`n", ${utf8Bom})
    [System.IO.File]::WriteAllText("$dir\b.docx\ページ001.tsv", "りんご`nみかん abc`n`n見積 確定`nabc", ${utf8Bom})
    [System.IO.File]::WriteAllText("$dir\c.pptx\スライド001.tsv", "abc`rりんご`r`r見積`r", (New-Object System.Text.UTF8Encoding $false))
    [System.IO.File]::WriteAllText("$dir\d.xlsx\空.tsv", "", ${utf8Bom})
    $files = (getIndexTsvFiles @($dir)).Files
    $targets = newTsvFiles $files

    # searchTsvFiles は List を 1 つのまま返すため、1 件ずつ "相対パス:行番号:行" にする
    function formatHits {
        param ($hits)
        $list = New-Object System.Collections.Generic.List[string]
        foreach ($hit in $hits) {
            $list.Add("$($hit.RelPath):$($hit.LineNumber):$($hit.Line)")
        }
        return ($list -join "|")
    }

    $cases = @(
        @("abc", $true), @("見積", $true), @("(株)", $true), @("存在しない", $true), @("`t", $true),
        @("^abc", $false), @("abc$", $false), @("^$", $false), @("x*", $false), @("見積.*確定", $false),
        @("見積\s確定", $false), @("\s", $false), @("[^a]", $false), @("abc(?=\s)", $false),
        @("abc(?!\s)", $false), @("\Aabc", $false), @("abc\z", $false), @("(?i)ABC", $false)
    )
    foreach ($case in $cases) {
        It "「$($case[0])」（文字どおり=$($case[1])）" {
            $search = newSearchRegex $case[0] $case[1]
            $expected = formatHits (searchTsvFiles $targets 0 $targets.Count $search.Regex -1)
            formatHits (searchTsvFiles $targets 0 $targets.Count $search.Regex -1 $search.TextRegex $search.ScanMode) | Should Be $expected
            if ($search.ScanMode -ne "scan") {
                # 全文だけで行を切り出す場合と、全文で絞ってから 1 行ずつ照合する場合も同じ
                formatHits (searchTsvFiles $targets 0 $targets.Count $search.Regex -1 $search.TextRegex "filter") | Should Be $expected
            }
        }
    }

    It "行番号は 1 行ずつ読んだときと同じ（CR だけの改行・空行も数える）" {
        $search = newSearchRegex "見積" $true
        formatHits (searchTsvFiles $targets 0 $targets.Count $search.Regex -1 $search.TextRegex $search.ScanMode) |
            Should Be "a.xlsx\S.tsv:1:見積`t(株)山田商事|a.xlsx\S.tsv:3:abc 見積|b.docx\ページ001.tsv:4:見積 確定|c.pptx\スライド001.tsv:4:見積"
    }

    It "上限を超えたら打ち切る（全文から行を切り出す場合も）" {
        $search = newSearchRegex "abc" $true
        (searchTsvFiles $targets 0 $targets.Count $search.Regex 1 $search.TextRegex $search.ScanMode).Count | Should Be 2
    }

    It "読んだ内容を使い回しても結果は同じ（1 行ずつ照合する正規表現も）" {
        $cache = newTsvTextCache
        foreach ($case in @(@("見積", $true), @("abc(?!\s)", $false))) {
            $search = newSearchRegex $case[0] $case[1]
            $expected = formatHits (searchTsvFiles $targets 0 $targets.Count $search.Regex -1 $search.TextRegex $search.ScanMode)
            formatHits (searchTsvFiles $targets 0 $targets.Count $search.Regex -1 $search.TextRegex $search.ScanMode $cache) | Should Be $expected
            formatHits (searchTsvFiles $targets 0 $targets.Count $search.Regex -1 $search.TextRegex $search.ScanMode $cache) | Should Be $expected
        }
        $cache.Texts.Count | Should Be 4
    }
}

Describe "searchIndex（並列検索・読んだ内容の使い回し）" -Tag Io {
    $dir = Join-Path $TestDrive "parallel"
    for ($i = 0; $i -lt 30; $i++) {
        newTsv ("$dir\sub{0}\book{1:D2}.xlsx\S.tsv" -f ($i % 3), $i) @("見積 $i", "x", "見積 確定 $i")
    }
    $files = (getIndexTsvFiles @($dir)).Files

    # 1 件ずつ "ファイル名:行番号" にする
    function formatBookLines {
        param ($hits)
        return (@($hits | ForEach-Object { "$($_.Book):$($_.LineNumber)" }) -join ",")
    }

    It "並列に検索しても、TSV の順に同じ結果を返す" {
        $expected = formatBookLines (searchIndex "見積" $files $true 0 1 -workerCount 1).Hits
        $result = searchIndex "見積" $files $true 0 1 -workerCount 3
        formatBookLines $result.Hits | Should Be $expected
        $result.Hits.Count | Should Be 60
    }

    It "並列でも上限で打ち切り、進み具合を TSV の順に知らせる" {
        $script:done = New-Object System.Collections.Generic.List[int]
        $result = searchIndex "見積" $files $true 5 2 { param($done, $total, $newHits) $script:done.Add($done) } -workerCount 3
        $result.Truncated | Should Be $true
        formatBookLines $result.Hits | Should Be "book00.xlsx:1,book00.xlsx:3,book03.xlsx:1,book03.xlsx:3,book06.xlsx:1"
        ($script:done -join ",") | Should Be ((@($script:done | Sort-Object)) -join ",")
    }

    It "並列でも中止できる" {
        $script:calls = 0
        $result = searchIndex "見積" $files $true 0 1 { param($done, $total, $newHits) $script:calls++ } { $script:calls -ge 2 } -workerCount 3
        $result.Cancelled | Should Be $true
        $script:calls | Should Be 2
    }

    It "内容を使い回し、更新された TSV は読み直す" {
        $cache = newTsvTextCache
        (searchIndex "更新後" $files $true -cache $cache).Hits.Count | Should Be 0
        $path = @($files.Keys)[0]
        [System.IO.File]::WriteAllText($path, "更新後の内容`r`n", ${utf8Bom})
        (Get-Item -LiteralPath $path).LastWriteTime = (Get-Date).AddMinutes(1)
        $updated = (getIndexTsvFiles @($dir)).Files
        (searchIndex "更新後" $updated $true -cache $cache).Hits.Count | Should Be 1
    }

    It "上限を超える内容は残さない" {
        $cache = newTsvTextCache 20
        [void](searchIndex "見積" $files $true -cache $cache)
        $cache.Chars[0] | Should BeLessThan 21
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
        $result.Header | Should Be "ファイル名`t場所`t種別`t行`tA`tB`tC"
        $result.Lines[0] | Should Be "x\A社.xlsx`t[シート] Sheet1`tセル`t1`ta`t`"りんご`nみかん`"`tc"
        $result.Lines[1] | Should Be "文書.docx`t[ページ] 1（目安）`t本文`t1`t`"`"`"引用`"`"で始まる りんご`""
    }

    It "検索結果ファイルの形式（02_検索.md）で書き出す" {
        $writer = New-Object System.IO.StringWriter
        writeSearchResult $writer "りんご" $hits
        $lines = $writer.ToString() -split "`r`n"
        $lines[0] | Should Be "【検索文字列　りんご】 2 件"
        $lines[1] | Should Be "ファイル名`t場所`t種別`t行`tA`tB`tC"
        $lines[4] | Should Be ""
    }

    It "0 件なら見出しの2行と空行だけ書き出す" {
        $writer = New-Object System.IO.StringWriter
        writeSearchResult $writer "無い" @()
        $writer.ToString() | Should Be "【検索文字列　無い】 0 件`r`nファイル名`t場所`t種別`t行`r`n`r`n"
    }
}
