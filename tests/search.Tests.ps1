# 検索・Officeプロセスなど、画面（config_gui.ps1）が使う common.ps1 の関数のテスト
# Pester 3.4 以降で実行: Invoke-Pester .\tests
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here\..\scripts\common.ps1"

function newTsv {
    param (
        [string]$path,
        [string[]]$lines
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), ${utf8Bom})
}

Describe "isValidRegex" {
    It "正しい正規表現は true" {
        isValidRegex "見積.*確定" | Should Be $true
    }

    It "不正な正規表現は false" {
        isValidRegex "(" | Should Be $false
    }
}

Describe "getIndexTsvFiles / testIndexExists / getIndexSummary" {
    $index = Join-Path $TestDrive "index[1]"
    $other = Join-Path $TestDrive "other"
    newTsv "$index\b.xlsx_S.tsv" @("b")
    newTsv "$index\sub\a.xlsx_S.tsv" @("a")
    newTsv "$other\c.docx_ページ001.tsv" @("c")
    (Get-Item -LiteralPath "$other\c.docx_ページ001.tsv").LastWriteTime = [datetime]"2030-01-02 03:04:05"
    $missing = Join-Path $TestDrive "missing"

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

Describe "searchIndex（今の形式: <ファイル名>\<場所>.tsv）" {
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

Describe "searchIndex" {
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

Describe "newSearchRegex" {
    It "文字どおりなら記号をそのまま探し、既定は大文字と小文字を区別しない" {
        $regex = (newSearchRegex "C++ (株)").Regex
        $regex.IsMatch("c++ (株)") | Should Be $true
        $regex.IsMatch("C (株)") | Should Be $false
    }

    It "正規表現として不正なワードは文字どおりにする" {
        $result = newSearchRegex "(" $false
        $result.SimpleMatch | Should Be $true
        $result.Regex.IsMatch("a(b") | Should Be $true
    }

    It "大文字と小文字を区別できる" {
        (newSearchRegex "ID" $true $true).Regex.IsMatch("社員id") | Should Be $false
        (newSearchRegex "ID" $true $true).Regex.IsMatch("社員ID") | Should Be $true
    }
}

Describe "newFileFilter" {
    It "; で区切ったワイルドカードで含め、! で始まるもので除く" {
        $filter = newFileFilter "*.xlsx；見積 ; !*old*"
        $filter.Include.IsMatch("A社.XLSX") | Should Be $true
        $filter.Include.IsMatch("2024見積書.docx") | Should Be $true
        $filter.Include.IsMatch("報告書.docx") | Should Be $false
        $filter.Exclude.IsMatch("A社_old.xlsx") | Should Be $true
    }

    It "空なら条件なし" {
        $filter = newFileFilter "  "
        $filter.Include | Should Be $null
        $filter.Exclude | Should Be $null
    }

    It "? は任意の1文字、ほかの記号は文字どおり" {
        $filter = newFileFilter "v?.[確定].xlsx"
        $filter.Include.IsMatch("v1.[確定].xlsx") | Should Be $true
        $filter.Include.IsMatch("v1x[確定].xlsx") | Should Be $false
    }
}

Describe "toSearchResultLines / writeSearchResult" {
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

Describe "splitTsvCells" {
    It 'ダブルクォートで囲まれたセルはタブを含んでも1セルとし、囲みを外して "" を " に戻す' {
        $cells = splitTsvCells "a`t`"b`tc`"`t`"d`"`"e`"`t"
        $cells.Count | Should Be 4
        $cells[0] | Should Be "a"
        $cells[1] | Should Be "b`tc"
        $cells[2] | Should Be "d`"e"
        $cells[3] | Should Be ""
    }

    It "1セルでも配列で返す" {
        (splitTsvCells "a").Count | Should Be 1
    }

    It "先頭のセルが空でも次のセルを落とさない" {
        $cells = splitTsvCells "`tりんご`t`"x`ty`""
        $cells.Count | Should Be 3
        $cells[0] | Should Be ""
        $cells[1] | Should Be "りんご"
        $cells[2] | Should Be "x`ty"
    }
}

Describe "readTsvContext" {
    $sep = [string][char]0x2028
    $path = Join-Path $TestDrive "context[1]\book.xlsx_Sheet1.tsv"
    newTsv $path @("r1", "r2", "r3${sep}改行", "r4", "`"r5`ta`"`tb", "r6", "r7", "r8", "r9", "r10")

    It "指定した行と前後の行を、行番号付きで返す（セル内改行の文字では行を分けない）" {
        $rows = @(readTsvContext $path 5 3 3)
        $rows.Count | Should Be 7
        $rows[0].LineNumber | Should Be 2
        $rows[1].Line | Should Be "r3${sep}改行"
        $rows[3].LineNumber | Should Be 5
        $rows[3].Line | Should Be "`"r5`ta`"`tb"
        $rows[6].LineNumber | Should Be 8
    }

    It "先頭の行では前の行を 1 行目までにする" {
        $rows = @(readTsvContext $path 2 3 1)
        ($rows | ForEach-Object { $_.LineNumber }) -join "," | Should Be "1,2,3"
    }

    It "末尾の行では後の行をファイルの最後までにする" {
        $rows = @(readTsvContext $path 9 1 3)
        ($rows | ForEach-Object { $_.LineNumber }) -join "," | Should Be "8,9,10"
    }

    It "1行だけでも配列で返す" {
        $rows = @(readTsvContext $path 1 0 0)
        $rows.Count | Should Be 1
        $rows[0].Line | Should Be "r1"
    }

    It "ファイルが無ければ空" {
        @(readTsvContext (Join-Path $TestDrive "missing.tsv") 1).Count | Should Be 0
    }
}

Describe "readSearchHistory / addSearchHistory" {
    $path = Join-Path $TestDrive "履歴.txt"

    It "新しい順に保存し、同じワードは先頭に移す" {
        [void](addSearchHistory "a" $path)
        [void](addSearchHistory "b" $path)
        $items = @(addSearchHistory " a " $path)
        $items -join "," | Should Be "a,b"
        @(readSearchHistory $path) -join "," | Should Be "a,b"
    }

    It "最大件数を超えた古いワードは消す" {
        [void](addSearchHistory "c" $path 2)
        @(readSearchHistory $path) -join "," | Should Be "c,a"
    }

    It "空のワードは保存しない" {
        @(addSearchHistory "  " $path) -join "," | Should Be "c,a"
    }
}

Describe "readSearchOption / writeSearchOption" {
    $path = Join-Path $TestDrive "setting.config"

    It "ファイルが無ければ、文字どおり・大文字と小文字を区別しない・対象ファイルはすべて" {
        $option = readSearchOption $path
        $option.UseRegex | Should Be $false
        $option.CaseSensitive | Should Be $false
        $option.FileFilter | Should Be ""
    }

    It "保存した値を読み込む" {
        writeSearchOption @{ UseRegex = $true } $path
        (readSearchOption $path).UseRegex | Should Be $true
        writeSearchOption @{ UseRegex = $false } $path
        (readSearchOption $path).UseRegex | Should Be $false
    }

    It "指定した項目だけを変え、ほかの項目は保つ" {
        writeSearchOption @{ UseRegex = $true; CaseSensitive = $true; FileFilter = "*.xlsx;!*old*" } $path
        writeSearchOption @{ CaseSensitive = $false } $path
        $option = readSearchOption $path
        $option.UseRegex | Should Be $true
        $option.CaseSensitive | Should Be $false
        $option.FileFilter | Should Be "*.xlsx;!*old*"
    }
}

Describe "getOfficeProcesses / stopOfficeProcesses" {
    It "ウィンドウを持たないプロセスをバックグラウンドとし、表示名を付ける" {
        Mock Get-Process {
            @(
                [pscustomobject]@{ Id = 1; ProcessName = "EXCEL"; MainWindowHandle = [IntPtr]::Zero; StartTime = [datetime]"2030-01-01"; WorkingSet64 = 10MB; MainWindowTitle = "" }
                [pscustomobject]@{ Id = 2; ProcessName = "WINWORD"; MainWindowHandle = [IntPtr]100; StartTime = [datetime]"2030-01-01"; WorkingSet64 = 20MB; MainWindowTitle = "文書 - Word" }
            )
        }
        $processes = @(getOfficeProcesses)
        $processes.Count | Should Be 2
        $processes[0].AppName | Should Be "Excel"
        $processes[0].Background | Should Be $true
        $processes[1].AppName | Should Be "Word"
        $processes[1].Background | Should Be $false
        $processes[1].MemoryMB | Should Be 20
    }

    It "プロセスが無ければ空配列" {
        Mock Get-Process { }
        @(getOfficeProcesses).Count | Should Be 0
    }

    It "終了できなかったプロセスは理由を返す" {
        Mock Stop-Process { if ($Id -eq 2) { throw "アクセスが拒否されました" } }
        $results = @(stopOfficeProcesses @(1, 2))
        $results.Count | Should Be 2
        $results[0].Stopped | Should Be $true
        $results[1].Stopped | Should Be $false
        $results[1].Message | Should Be "アクセスが拒否されました"
    }
}
