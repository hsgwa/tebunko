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

Describe "testIndexExists / getIndexSummary" -Tag Io {
    $index = Join-Path $TestDrive "index[1]"
    $other = Join-Path $TestDrive "other"
    foreach ($item in @(@{ Folder = "$index"; Name = "b.xlsx" }, @{ Folder = "$index\sub"; Name = "a.xlsx" }, @{ Folder = "$other"; Name = "c.docx" })) {
        [void][System.IO.Directory]::CreateDirectory($item.Folder)
        writePackFile "$($item.Folder)\$(getPackFileName (getPackExtension $item.Name))" (convertToPackText @(@{ Name = $item.Name; Places = @(@{ Place = "S"; Text = "x" }) }))
    }
    (Get-Item -LiteralPath "$other\content.docx.001.tsv").LastWriteTime = [datetime]"2030-01-02 03:04:05"
    # 集約する前の TSV（インデックス作成の途中）は数えない
    newTsv "$other\d.xlsx\S.tsv" @("d")
    $missing = Join-Path $TestDrive "missing"

    It "集約ファイルの有無を判定する" {
        testIndexExists @($missing, $other) | Should Be $true
        testIndexExists @($missing) | Should Be $false
    }

    It "集約ファイルの件数・最新の更新日時・存在しないフォルダを返す" {
        $summary = getIndexSummary @($index, $other, $missing, $index)
        $summary.Count | Should Be 3
        $summary.LastWrite | Should Be ([datetime]"2030-01-02 03:04:05")
        $summary.Missing.Count | Should Be 1
    }

    It "集約ファイルの無いフォルダ・存在しないフォルダだけなら、なし・件数 0・更新日時なし" {
        $empty = Join-Path $TestDrive "empty_index"
        New-Item -ItemType Directory -Path "$empty\sub" -Force | Out-Null
        newTsv "$empty\sub\e.xlsx\S.tsv" @("e")
        testIndexExists @($empty, $missing) | Should Be $false
        testIndexExists @() | Should Be $false
        $summary = getIndexSummary @($empty)
        $summary.Count | Should Be 0
        $summary.LastWrite | Should Be $null
        @($summary.Missing).Count | Should Be 0
    }
}

Describe "検索結果の行（集約ファイルのヒット）" -Tag Io {
    It "ファイル名に .xlsx_ を含んでも、元のファイル名と場所で組み立てる" {
        $index = Join-Path $TestDrive "search_dir"
        newTsv "$index\A社.xlsx_old.xlsx\$(toIndexFileName "Sheet1")" @("りんご`t200")
        foreach ($folder in (findIndexFoldersWithBooks $index)) { [void](updateIndexFolderPack $folder) }
        $hit = @((searchPackIndex "りんご" (getIndexPackFiles @($index)).Packs $true).Hits)[0]
        $hit.Book | Should Be "A社.xlsx_old.xlsx"
        (toSearchResultLines @($hit)).Lines[0] | Should Be "A社.xlsx_old.xlsx`t[シート] Sheet1`tセル`t1`tりんご`t200"
    }

    It "Excel の図形の場所（名前に [ ] を含む）も検索でき、セル内改行を戻して出力する" {
        $objectIndex = Join-Path $TestDrive "search_object"
        newTsv "$objectIndex\[確定]見積.xlsx\$(toIndexFileName "見積[図形]")" @("F2`t`"納期は$([char]0x2028)別途`"")
        foreach ($folder in (findIndexFoldersWithBooks $objectIndex)) { [void](updateIndexFolderPack $folder) }
        $hit = @((searchPackIndex "納期" (getIndexPackFiles @($objectIndex)).Packs $true).Hits)[0]
        $hit.Book | Should Be "[確定]見積.xlsx"
        $hit.Location | Should Be "見積[図形]"
        (toSearchResultLines @($hit)).Lines[0] | Should Be "[確定]見積.xlsx`t[シート] 見積`t図形`t1`tF2`t`"納期は`n別途`""
    }
}
Describe "toSearchResultLines / writeSearchResult" -Tag Io {
    $index = Join-Path $TestDrive "result"
    newTsv "$index\x\A社.xlsx\Sheet1.tsv" @("a`t`"りんご${cellNewLine}みかん`"`tc")
    newTsv "$index\文書.docx\ページ001.tsv" @("`"引用`"で始まる りんご")
    foreach ($folder in (findIndexFoldersWithBooks $index)) { [void](updateIndexFolderPack $folder) }
    $hits = @((searchPackIndex "りんご" (getIndexPackFiles @($index)).Packs $true).Hits | Sort-Object RelDir -Descending)

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
