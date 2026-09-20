# Pester 3.4 以降で実行: Invoke-Pester .\tests
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here\..\scripts\win_grep\lib.ps1"

Describe "toSafeFileName" {
    It "ファイル名に使えない文字を全角に変換する" {
        toSafeFileName 'a<b>c\d*e:f?g|h' | Should Be 'a＜b＞c￥d＊e：f？g｜h'
    }

    It "スラッシュを全角に変換する" {
        toSafeFileName 'a/b' | Should Be 'a／b'
    }

    It "ダブルクォートを全角に変換する" {
        toSafeFileName 'a"b' | Should Be 'a”b'
    }

    It "使える文字はそのまま返す" {
        toSafeFileName 'シート1 (2)' | Should Be 'シート1 (2)'
    }
}

Describe "encodeIndexPlace / decodeIndexPlace" {
    It "ファイル名に使えない文字・_・% を %XX にし、元に戻せる" {
        $place = 'a<b>c\d*e:f?g|h/i"j_k%l'
        $encoded = encodeIndexPlace $place
        $encoded | Should Be 'a%3Cb%3Ec%5Cd%2Ae%3Af%3Fg%7Ch%2Fi%22j%5Fk%25l'
        decodeIndexPlace $encoded | Should BeExactly $place
    }

    It "制御文字も %XX にする" {
        encodeIndexPlace "a`tb" | Should Be 'a%09b'
        decodeIndexPlace 'a%09b' | Should Be "a`tb"
    }

    It "半角の記号と全角の記号は別の名前になる（衝突しない）" {
        encodeIndexPlace '衝突"' | Should Not Be (encodeIndexPlace '衝突”')
    }

    It "使える文字（全角記号・空白・&'#() 等）はそのまま返す" {
        # PowerShell は ” を " と同じに扱うため、' で囲む
        encodeIndexPlace 'シート1 (2)&''#＜” 全角　' | Should Be 'シート1 (2)&''#＜” 全角　'
    }

    It "符号化で作らない %XX はそのまま返す（以前の版のシート名 100% 等）" {
        decodeIndexPlace '100%' | Should Be '100%'
        decodeIndexPlace '%41%2a' | Should Be '%41%2a'
    }
}

Describe "toIndexFileName" {
    It "<場所>.tsv にし、場所は符号化する（元のファイル名はフォルダ名にするため入れない）" {
        toIndexFileName "記号<>_1" | Should Be "記号%3C%3E%5F1.tsv"
    }

    It "場所が長くても255文字を超えなければ返す" {
        (toIndexFileName ("あ" * 251)).Length | Should Be 255
    }

    It "255文字を超えると、文字数の分かるメッセージで例外にする" {
        { toIndexFileName ("あ" * 252) } | Should Throw "256 文字。上限 255 文字"
    }
}

Describe "splitIndexTsvPath" {
    It "今の形式（<ファイル名.xlsx>\<場所>.tsv）をファイル名・場所・フォルダに分ける" {
        foreach ($case in @(
            @("A.xlsx", "old.xlsx_1"),
            @("A.xlsx_old.xlsx", "1"),
            @("コピー.xls_old.xlsx", "Sheet1"),
            @("終わりが_.xlsx", "_"),
            @("100%.xlsx", "100%"),
            @("ア_イ_ウ.xlsx", "シ_ー_ト"),
            @("[確定]報告書.xlsx", "衝突`""),
            @("資料.pptx", "スライド003_ノート"),
            @("報告書.docx", "ヘッダー・フッター")
        )) {
            $parts = splitIndexTsvPath "営業\2024\$($case[0])\$(toIndexFileName $case[1])"
            $parts.Book | Should BeExactly $case[0]
            $parts.Place | Should BeExactly $case[1]
            $parts.RelDir | Should Be "営業\2024"
        }
    }

    It "インデックスフォルダの直下のファイルは、フォルダが空になる" {
        $parts = splitIndexTsvPath "A.xlsx\Sheet1.tsv"
        $parts.Book | Should Be "A.xlsx"
        $parts.Place | Should Be "Sheet1"
        $parts.RelDir | Should Be ""
    }

    It "以前の形式（<ファイル名.xlsx>_<場所>.tsv）も分けられる" {
        $parts = splitIndexTsvPath "営業\コピー.xls_old.xlsx_Sheet1.tsv"
        $parts.Book | Should Be "コピー.xls_old.xlsx"
        $parts.Place | Should Be "Sheet1"
        $parts.RelDir | Should Be "営業"

        # さらに以前の形式（場所に _ をそのまま入れていた版）
        $parts = splitIndexTsvPath "売上.xls_2024_上期.tsv"
        $parts.Book | Should Be "売上.xls"
        $parts.Place | Should Be "2024_上期"
        $parts.RelDir | Should Be ""
    }

    It "Officeファイル以外の名前のフォルダにあるTSVは、ファイル名をそのまま返す" {
        $parts = splitIndexTsvPath "メモ\other.tsv"
        $parts.Book | Should Be "other.tsv"
        $parts.Place | Should Be ""
        $parts.RelDir | Should Be "メモ"
    }
}

Describe "toLongPath / fromLongPath" {
    It "ドライブのパスに \\?\ を付け、外すと元に戻る" {
        toLongPath "C:\data\a.xlsx" | Should Be "\\?\C:\data\a.xlsx"
        fromLongPath "\\?\C:\data\a.xlsx" | Should Be "C:\data\a.xlsx"
    }

    It "ネットワークのパスは \\?\UNC\ にし、外すと元に戻る" {
        toLongPath "\\server\share\a.xlsx" | Should Be "\\?\UNC\server\share\a.xlsx"
        fromLongPath "\\?\UNC\server\share\a.xlsx" | Should Be "\\server\share\a.xlsx"
    }

    It "付いていればそのまま、付いていなければ外してもそのまま" {
        toLongPath "\\?\C:\a" | Should Be "\\?\C:\a"
        fromLongPath "C:\a" | Should Be "C:\a"
    }

    It "/ は \ にする" {
        toLongPath "C:/data/a.xlsx" | Should Be "\\?\C:\data\a.xlsx"
    }
}

Describe "copyFileShared" {
    It "ほかのアプリが書き込み用に開いているファイルもコピーでき、コピー中もほかのアプリの書き込みを妨げない" {
        $source = "$TestDrive\共有 [1]\元.xlsx"
        [void][System.IO.Directory]::CreateDirectory((Split-Path $source -Parent))
        [System.IO.File]::WriteAllText($source, "abc", [System.Text.Encoding]::ASCII)

        # 利用者が編集中（書き込みあり・ほかの読み取りだけ許可）の状態
        $editing = New-Object System.IO.FileStream($source, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        try {
            copyFileShared $source "$TestDrive\コピー.xlsx"
        } finally {
            $editing.Dispose()
        }
        [System.IO.File]::ReadAllText("$TestDrive\コピー.xlsx") | Should Be "abc"
    }

    It "読み取り専用のファイルも、通常の属性のコピーを作り、既にあれば上書きする" {
        $source = "$TestDrive\読み取り専用.docx"
        $dest = "$TestDrive\上書き.docx"
        [System.IO.File]::WriteAllText($source, "new", [System.Text.Encoding]::ASCII)
        [System.IO.File]::SetAttributes($source, [System.IO.FileAttributes]::ReadOnly)
        [System.IO.File]::WriteAllText($dest, "old-content", [System.Text.Encoding]::ASCII)
        try {
            copyFileShared $source $dest
            [System.IO.File]::ReadAllText($dest) | Should Be "new"
            ([System.IO.File]::GetAttributes($dest) -band [System.IO.FileAttributes]::ReadOnly) | Should Be 0
        } finally {
            [System.IO.File]::SetAttributes($source, [System.IO.FileAttributes]::Normal)
        }
    }

    It "長いパス（260文字超）のファイルもコピーできる" {
        $dir = "$TestDrive\copyLong\" + ("a" * 100) + "\" + ("b" * 100)
        [void][System.IO.Directory]::CreateDirectory((toLongPath $dir))
        $source = "$dir\book.xlsx"
        [System.IO.File]::WriteAllText((toLongPath $source), "long", [System.Text.Encoding]::ASCII)
        try {
            copyFileShared $source "$TestDrive\long_copy.xlsx"
            [System.IO.File]::ReadAllText("$TestDrive\long_copy.xlsx") | Should Be "long"
        } finally {
            Remove-Item -LiteralPath (toLongPath "$TestDrive\copyLong") -Recurse -Force
        }
    }
}

Describe "長いパス（260文字超）" {
    # フォルダのパスが約248文字、ファイルのパスが260文字を超えると、\\?\ を付けないと扱えない
    $deepRel = ("a" * 100) + "\" + ("b" * 100)

    It "prettyTsv で長いパスに保存できる" {
        $dir = "$TestDrive\prettyLong\$deepRel"
        [void][System.IO.Directory]::CreateDirectory((toLongPath $dir))
        $in = "$TestDrive\long_sheet.tmp"
        [System.IO.File]::WriteAllText($in, "x`r`n", [System.Text.Encoding]::Unicode)
        try {
            $out = "$dir\book.xlsx_Sheet1.tsv"
            $out.Length | Should BeGreaterThan 260
            prettyTsv $in $out | Should Be $true
            [System.IO.File]::ReadAllText((toLongPath $out)) | Should Be "x`r`n"
        } finally {
            Remove-Item -LiteralPath (toLongPath "$TestDrive\prettyLong") -Recurse -Force
        }
    }

    It "getIndexTsvFiles・searchIndex で長いパスのTSVも列挙・検索でき、相対パスは \\?\ の無い形になる" {
        $root = "$TestDrive\indexLong"
        $dir = "$root\$deepRel"
        [void][System.IO.Directory]::CreateDirectory((toLongPath $dir))
        [System.IO.File]::WriteAllText((toLongPath "$dir\book.xlsx_Sheet1.tsv"), "hello`r`n", $utf8Bom)
        [System.IO.File]::WriteAllText("$root\short.xlsx_Sheet1.tsv", "hello`r`n", $utf8Bom)
        try {
            $index = getIndexTsvFiles @($root)
            $index.Folders[0].Count | Should Be 2
            $rel = @($index.Files.Values | ForEach-Object { $_.RelPath } | Sort-Object)
            $rel | Should Be @("$deepRel\book.xlsx_Sheet1.tsv", "short.xlsx_Sheet1.tsv")

            $hits = @((searchIndex "hello" $index.Files).Hits)
            $hits.Count | Should Be 2
            @($hits | Where-Object { $_.RelDir -eq $deepRel }).Count | Should Be 1

            (getIndexSummary @($root)).Count | Should Be 2
            testIndexExists @($root) | Should Be $true
        } finally {
            Remove-Item -LiteralPath (toLongPath $root) -Recurse -Force
        }
    }
}

Describe "replaceCellNewLine" {
    It "ダブルクォート内の改行（LF・CR・CRLF）をセル内改行の文字に置き換える" {
        replaceCellNewLine "`"a`nb`rc`r`nd`"`te`r`n" | Should Be "`"a${cellNewLine}b${cellNewLine}c${cellNewLine}d`"`te`r`n"
    }

    It "ダブルクォート外の改行は残す" {
        replaceCellNewLine "a`r`nb" | Should Be "a`r`nb"
    }
}

Describe "formatTsv" {
    It "行末の空セルと末尾の空行を取り除き、途中の空行は残す" {
        $lines = (formatTsv "a`tb`t`t`r`n`t`r`nc  `r`n`t`r`n") -split "`r`n"
        $lines.Count | Should Be 3
        $lines[0] | Should Be "a`tb"
        $lines[1] | Should Be ""
        $lines[2] | Should Be "c  "
    }

    It "セル内改行を含む行を1行にまとめる" {
        $lines = (formatTsv "`"x`r`ny`"`tz`r`nw`r`n") -split "`r`n"
        $lines.Count | Should Be 2
        $lines[0] | Should Be "`"x${cellNewLine}y`"`tz"
        $lines[1] | Should Be "w"
    }

    It "出力範囲の左上の位置に合わせて、先頭に空行・空セルを補う（D5 → 5行目の4列目）" {
        $lines = (formatTsv "a`t`tb`r`n`r`nc`r`n" 5 4) -split "`r`n"
        $lines.Count | Should Be 7
        $lines[0..3] -join "|" | Should Be "|||"
        $lines[4] | Should Be "`t`t`ta`t`tb"
        $lines[5] | Should Be ""
        $lines[6] | Should Be "`t`t`tc"
    }

    It "空白だけなら空文字を返す" {
        formatTsv "`t `t`r`n`r`n" 3 2 | Should Be ""
    }
}

Describe "prettyTsv" {
    It "UTF-16のTSVを整形してBOM付きUTF-8で保存する（[ ] を含むパスも可）" {
        $in = "$TestDrive\sheet1.tmp"
        $out = "$TestDrive\見積[確定].xlsx_一覧.tsv"
        [System.IO.File]::WriteAllText($in, "`"a`nb`"`tc`t`r`n`t`r`n", [System.Text.Encoding]::Unicode)

        prettyTsv $in $out | Should Be $true
        $bytes = [System.IO.File]::ReadAllBytes($out)
        $bytes[0..2] -join "," | Should Be "239,187,191"
        [System.IO.File]::ReadAllText($out) | Should Be "`"a${cellNewLine}b`"`tc`r`n"
    }

    It "出力範囲の左上の位置を指定できる" {
        $in = "$TestDrive\sheet2.tmp"
        $out = "$TestDrive\book.xlsx_B2.tsv"
        [System.IO.File]::WriteAllText($in, "x`r`n", [System.Text.Encoding]::Unicode)

        prettyTsv $in $out 2 2 | Should Be $true
        [System.IO.File]::ReadAllText($out) | Should Be "`r`n`tx`r`n"
    }

    It "内容が空なら保存せず `$false を返す" {
        $in = "$TestDrive\empty.tmp"
        $out = "$TestDrive\empty.tsv"
        [System.IO.File]::WriteAllText($in, "`t`t`r`n`r`n", [System.Text.Encoding]::Unicode)

        prettyTsv $in $out | Should Be $false
        Test-Path -LiteralPath $out | Should Be $false
    }
}

Describe "readListFile / writeListFile" {
    It "書き込んだ行をそのまま読み込める（[ ] や先頭の空白を含むパス）" {
        $path = "$TestDrive\list[1].txt"
        writeListFile $path @("a\[確定]見積.xlsx", " b.xls")
        $lines = readListFile $path
        $lines.Count | Should Be 2
        $lines[0] | Should Be "a\[確定]見積.xlsx"
        $lines[1] | Should Be " b.xls"
    }

    It "ファイルが無ければ空配列を返す" {
        @(readListFile "$TestDrive\none_list.txt").Count | Should Be 0
    }
}

Describe "readStatusFile / writeStatusFile / addStatusRow" {
    It "まとめて書き出した行と、1件ずつ追記した行の整形が同じ（速さのために別々に書いているため）" {
        # writeStatusFile は数万行を速く書くため、toStatusLine と同じ整形をその場に展開している
        $row = newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "10420" $stateFailed "" "2026/09/20 10:00:00" "エラー`tの`r`n説明"
        $path = "$TestDrive\status_same.tsv"
        writeStatusFile @() @($row) $path
        $written = @(readStatusLines $path)[-1]
        $written | Should Be (toStatusLine $row)
        $written | Should Be "営業\見積.xlsx`t2025/01/10 12:34:56`t10420`t${stateFailed}`t`t2026/09/20 10:00:00`tエラー の 説明"
    }

    It "書き込んだ変換対象フォルダと行をそのまま読み込める（[ ] や先頭の空白を含むパス）" {
        $path = "$TestDrive\status[1].tsv"
        $rows = @(
            (newStatusRow "a\[確定]見積.xlsx" "2025/01/10 12:34:56" "10420" $stateDone "3" "2026/09/19 10:00:00"),
            (newStatusRow " b.xls" "2025/02/01 08:00:00" "0" $stateNew)
        )
        $folders = @(
            [pscustomobject]@{ Path = "C:\data [1]"; Name = "data [1]" },
            [pscustomobject]@{ Path = "D:\"; Name = "D" }
        )
        writeStatusFile $folders $rows $path

        $status = readStatusFile $path
        $status.Folders.Count | Should Be 2
        $status.Folders[0].Path | Should Be "C:\data [1]"
        $status.Folders[0].Name | Should Be "data [1]"
        $status.Folders[1].Path | Should Be "D:\"
        $status.Folders[1].Name | Should Be "D"
        $status.Rows.Count | Should Be 2
        $row = $status.Rows["a\[確定]見積.xlsx"]
        $row.更新日時 | Should Be "2025/01/10 12:34:56"
        $row.サイズ | Should Be "10420"
        $row.状態 | Should Be $stateDone
        $row.TSV数 | Should Be "3"
        $status.Rows[" b.xls"].状態 | Should Be $stateNew
    }

    It "先頭に変換対象フォルダ（パス・インデックス名）、次に見出しのTSVになる" {
        $path = "$TestDrive\status_lines.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        $lines = [System.IO.File]::ReadAllLines($path)
        $lines[0] | Should Be "変換対象フォルダ`tC:\data`tdata"
        $lines[1] | Should Be "相対パス`t更新日時`tサイズ`t状態`tTSV数`t変換日時`tエラー"
        $lines[2] | Should Be "a.xlsx`t2025/01/10 12:34:56`t1`t未変換`t`t`t"
    }

    It "追記した行が前の行より優先される。相対パスの大文字・小文字は区別しない" {
        $path = "$TestDrive\status_append.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "Dir\A.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        addStatusRow (newStatusRow "dir\a.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 10:00:00" "パスワードが違います") $path

        $status = readStatusFile $path
        $status.Rows.Count | Should Be 1
        $status.Rows["DIR\A.XLSX"].状態 | Should Be $stateFailed
        $status.Rows["DIR\A.XLSX"].エラー | Should Be "パスワードが違います"
    }

    It "エラーメッセージのタブ・改行はスペースにして1行に収める" {
        $path = "$TestDrive\status_error.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "" "" $stateFailed "" "" "1行目`r`n2行目`tタブ") $path
        (readStatusFile $path).Rows["a.xlsx"].エラー | Should Be "1行目 2行目 タブ"
    }

    It "列数の合わない行（書き込み途中で中断した行）は無視する" {
        $path = "$TestDrive\status_broken.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1") $path
        [System.IO.File]::AppendAllText($path, "a.xlsx`t2025/01/10", $utf8Bom)

        $status = readStatusFile $path
        $status.Rows["a.xlsx"].状態 | Should Be $stateDone
    }

    It "書き直すと既存のファイルを置き換え、一時ファイルは残らない" {
        $path = "$TestDrive\status_replace.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\old"; Name = "old" }) @(newStatusRow "old\a.xlsx") $path
        writeStatusFile @([pscustomobject]@{ Path = "C:\new"; Name = "new" }) @(newStatusRow "new\b.xlsx") $path

        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Path }) | Should Be @("C:\new")
        @($status.Rows.Keys) | Should Be @("new\b.xlsx")
        Test-Path -LiteralPath "${path}.tmp" | Should Be $false
    }

    It "以前の形式（変換対象フォルダが1つでインデックス名なし）も読める" {
        $path = "$TestDrive\status_legacy.tsv"
        [System.IO.File]::WriteAllLines($path, [string[]]@(
            "変換対象フォルダ`tC:\old",
            "相対パス`t更新日時`tサイズ`t状態`tTSV数`t変換日時`tエラー",
            "a.xlsx`t2025/01/10 12:34:56`t1`t済`t1`t`t"
        ), $utf8Bom)

        $status = readStatusFile $path
        $status.Folders[0].Path | Should Be "C:\old"
        $status.Folders[0].Name | Should Be ""
        $status.Rows["a.xlsx"].状態 | Should Be $stateDone
    }

    It "ファイルが無ければ空の一覧を返す" {
        $status = readStatusFile "$TestDrive\none_status.tsv"
        $status.Folders.Count | Should Be 0
        $status.Rows.Count | Should Be 0
    }
}

Describe "readConvertingFile / writeConvertingFile / removeConvertingFile" {
    It "書き込んだ相対パスと回数をそのまま読み込める（[ ] や空白を含むパス）" {
        $path = "$TestDrive\converting[1].txt"
        writeConvertingFile "フォルダ1\a [確定]\見積.xlsx" 2 $path

        $converting = readConvertingFile $path
        $converting.RelPath | Should Be "フォルダ1\a [確定]\見積.xlsx"
        $converting.Count | Should Be 2
    }

    It "削除すると記録なし（`$null）になる" {
        $path = "$TestDrive\converting_remove.txt"
        writeConvertingFile "a.xlsx" 1 $path
        removeConvertingFile $path

        Test-Path -LiteralPath $path | Should Be $false
        readConvertingFile $path | Should Be $null
    }

    It "ファイルが無くても削除でエラーにならない" {
        { removeConvertingFile "$TestDrive\none_converting.txt" } | Should Not Throw
    }

    It "壊れた記録（回数が数値でない・相対パスが無い・空）は `$null を返す" {
        $path = "$TestDrive\converting_broken.txt"
        foreach ($content in @("x`ta.xlsx", "0`ta.xlsx", "1`t", "a.xlsx", "")) {
            [System.IO.File]::WriteAllText($path, $content, $utf8Bom)
            readConvertingFile $path | Should Be $null
        }
    }
}

Describe "formatFileTime" {
    It "秒までの日時にする" {
        formatFileTime (New-Object DateTime 2025, 1, 2, 3, 4, 5, 678) | Should Be "2025/01/02 03:04:05"
    }
}

Describe "splitIndexFileName" {
    It "ブック名とシート名に分解する" {
        $name = splitIndexFileName "book.xlsx_Sheet1.tsv"
        $name.book | Should Be "book.xlsx"
        $name.sheet | Should Be "Sheet1"
    }

    It "場所の %XX を元に戻す" {
        $name = splitIndexFileName "売上.xls_2024%5F上期%22.tsv"
        $name.book | Should Be "売上.xls"
        $name.sheet | Should Be '2024_上期"'
    }

    It "ファイル名に .xls_ 等を含んでも、最後の _ で分解する" {
        $name = splitIndexFileName "コピー.xls_old.xlsx_Sheet1.tsv"
        $name.book | Should Be "コピー.xls_old.xlsx"
        $name.sheet | Should Be "Sheet1"
    }

    It "以前の版のTSV（シート名の _ を符号化していない）も分解できる" {
        $name = splitIndexFileName "売上.xls_2024_上期.tsv"
        $name.book | Should Be "売上.xls"
        $name.sheet | Should Be "2024_上期"

        $name = splitIndexFileName "ア_イ_ウ.xlsx_シ_ト＜＞.tsv"
        $name.book | Should Be "ア_イ_ウ.xlsx"
        $name.sheet | Should Be "シ_ト＜＞"
    }

    It "拡張子が大文字でも分解できる" {
        $name = splitIndexFileName "大文字.XLSX_Sheet1.tsv"
        $name.book | Should Be "大文字.XLSX"
        $name.sheet | Should Be "Sheet1"
    }

    It "xlsm も扱える" {
        (splitIndexFileName "macro.xlsm_A.tsv").book | Should Be "macro.xlsm"
    }

    It "Word・PowerPointのファイル名と場所に分解する" {
        $name = splitIndexFileName "報告書.docx_ページ001.tsv"
        $name.book | Should Be "報告書.docx"
        $name.sheet | Should Be "ページ001"

        $name = splitIndexFileName "旧.doc_ヘッダー・フッター.tsv"
        $name.book | Should Be "旧.doc"
        $name.sheet | Should Be "ヘッダー・フッター"

        $name = splitIndexFileName "提案.pptx_スライド003%5Fノート.tsv"
        $name.book | Should Be "提案.pptx"
        $name.sheet | Should Be "スライド003_ノート"
    }

    It "形式外のファイル名はそのままブック名として返す" {
        $name = splitIndexFileName "other.tsv"
        $name.book | Should Be "other.tsv"
        $name.sheet | Should Be ""
    }
}

Describe "toResultLine" {
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

Describe "countTsvFields" {
    It "タブ区切りのセル数を返す" {
        countTsvFields "a`t`tb" | Should Be 3
        countTsvFields "" | Should Be 1
    }

    It "先頭のセルが空でも数え落とさない" {
        countTsvFields "`tb`tc" | Should Be 3
        countTsvFields "`t`"b`tc`"" | Should Be 2
        countTsvFields "`t" | Should Be 2
    }

    It '" で始まるセル内のタブ・改行は区切りとしない' {
        countTsvFields "a`t`"左`t右`"`t`"`"`"x`"`"`"" | Should Be 3
        countTsvFields "`"1行目`n2行目`"`tb" | Should Be 2
    }

    It '" で始まらないセルの " は囲みとしない' {
        countTsvFields "a`"b`tc`"d" | Should Be 2
    }
}

Describe "toColumnName" {
    It "列番号を列名に変換する" {
        toColumnName 1 | Should Be "A"
        toColumnName 26 | Should Be "Z"
        toColumnName 27 | Should Be "AA"
        toColumnName 702 | Should Be "ZZ"
        toColumnName 703 | Should Be "AAA"
        toColumnName 16384 | Should Be "XFD"
    }
}

Describe "toResultHeader" {
    It "ファイル名・場所・行と、列名を並べる" {
        toResultHeader 3 | Should Be "ファイル名`t場所`t行`tA`tB`tC"
        toResultHeader 0 | Should Be "ファイル名`t場所`t行"
    }
}

Describe "readSettings / writeSettings" {
    It "ファイルが無ければ既定値を返し、ファイルを作らない" {
        $path = "$TestDrive\空\setting.config"
        $settings = readSettings $path
        @($settings.targetFolders).Count | Should Be 0
        @($settings.indexSources).Count | Should Be 0
        $settings.useRegex | Should Be $false
        $settings.openMode | Should Be "normal"
        Test-Path -LiteralPath $path | Should Be $false
    }

    It "開き方（openMode）を保存・読み込みできる" {
        $path = "$TestDrive\開き方\setting.config"
        foreach ($mode in ${openModes}) {
            writeOpenMode $mode $path
            readOpenMode $path | Should Be $mode
        }
    }

    It "開き方が無い・知らない値なら「通常」とする" {
        $path = "$TestDrive\開き方2\setting.config"
        readOpenMode $path | Should Be ${openModeNormal}        # ファイルが無い
        updateSettings "openMode" "知らない値" $path
        readOpenMode $path | Should Be ${openModeNormal}
    }

    It "保存した設定をそのまま読み込める（1件だけの一覧も配列のまま）" {
        $path = "$TestDrive\往復 [1]\setting.config"
        $settings = newSettings
        $settings.targetFolders = @([pscustomobject]@{ path = "C:\a [1]"; enabled = $true })
        $settings.indexSources = @([pscustomobject]@{ name = "index<1>"; path = "D:\index<1>" })
        $settings.useRegex = $true
        writeSettings $settings $path
        $read = readSettings $path
        @($read.targetFolders).Count | Should Be 1
        $read.targetFolders[0].path | Should Be "C:\a [1]"
        $read.indexSources[0].path | Should Be "D:\index<1>"
        $read.useRegex | Should Be $true
        # 1つの項目だけを変えると、ほかの項目は保つ
        updateSettings "useRegex" $false $path
        $read = readSettings $path
        $read.useRegex | Should Be $false
        $read.indexSources[0].path | Should Be "D:\index<1>"
    }

    It "記載の無い項目は既定値、空のファイルは既定値" {
        $path = "$TestDrive\一部.json"
        [System.IO.File]::WriteAllText($path, '{ "useRegex": true }', ${utf8Bom})
        $settings = readSettings $path
        $settings.useRegex | Should Be $true
        @($settings.targetFolders).Count | Should Be 0
        [System.IO.File]::WriteAllText($path, "", ${utf8Bom})
        (readSettings $path).useRegex | Should Be $false
    }

    It "JSON として読めなければ例外を投げる" {
        $path = "$TestDrive\壊れ.json"
        [System.IO.File]::WriteAllText($path, "{ targetFolders: ", ${utf8Bom})
        { readSettings $path } | Should Throw "読み込めません"
    }

    It "設定ファイルが無く config フォルダに以前の設定ファイル（*.txt）があれば、移して保存する" {
        $dir = "$TestDrive\以前"
        writeListFile "$dir\config\変換対象フォルダパス.txt" @("C:\データ\Excel\", "", "# D:\old\報告書")
        writeListFile "$dir\config\検索オプション.txt" @("正規表現=オン")
        $settings = readSettings "$dir\setting.config"
        Test-Path -LiteralPath "$dir\setting.config" | Should Be $true
        $folders = @(getTargetFolders "$dir\setting.config")
        $folders.Count | Should Be 2
        $folders[0].Path | Should Be "C:\データ\Excel"
        $folders[0].Enabled | Should Be $true
        $folders[1].Path | Should Be "D:\old\報告書"
        $folders[1].Enabled | Should Be $false
        (readSearchOption "$dir\setting.config").UseRegex | Should Be $true
        # 移した後は以前の設定ファイルを使わない
        writeListFile "$dir\config\検索オプション.txt" @("正規表現=オフ")
        (readSearchOption "$dir\setting.config").UseRegex | Should Be $true
    }
}

Describe "getTargetFolders / writeTargetFolders" {
    It "記載順に返し、enabled が false はチェックなし、記載が無ければチェックあり" {
        $path = "$TestDrive\targets.json"
        [System.IO.File]::WriteAllText($path, @'
{ "targetFolders": [
    { "path": "C:\\データ\\Excel", "enabled": true },
    { "path": "D:\\old\\報告書", "enabled": false },
    { "path": "\"F:\\引用符付き\\\"" },
    { "path": "  " }
] }
'@, ${utf8Bom})
        $folders = @(getTargetFolders $path)
        $folders.Count | Should Be 3
        $folders[0].Path | Should Be "C:\データ\Excel"
        $folders[0].Enabled | Should Be $true
        $folders[1].Path | Should Be "D:\old\報告書"
        $folders[1].Enabled | Should Be $false
        $folders[2].Path | Should Be "F:\引用符付き"
        $folders[2].Enabled | Should Be $true
    }

    It "同じフォルダ（大文字・小文字、末尾の \ の違い）は最初のものだけ使う" {
        $path = "$TestDrive\targets_dup.json"
        writeTargetFolders @(
            [pscustomobject]@{ Path = "C:\Data"; Enabled = $true },
            [pscustomobject]@{ Path = "c:\data\"; Enabled = $false }
        ) $path
        $folders = @(getTargetFolders $path)
        $folders.Count | Should Be 1
        $folders[0].Enabled | Should Be $true
    }

    It "設定が無ければ空の配列を返す" {
        @(getTargetFolders "$TestDrive\none_targets.json").Count | Should Be 0
    }

    It "保存した一覧をそのまま読み込め、ほかの設定は保つ" {
        $path = "$TestDrive\targets_write.json"
        writeSearchOption @{ UseRegex = $true } $path
        writeTargetFolders @(
            [pscustomobject]@{ Path = "C:\a [1]"; Enabled = $true },
            [pscustomobject]@{ Path = "D:\b"; Enabled = $false }
        ) $path
        $folders = @(getTargetFolders $path)
        $folders.Count | Should Be 2
        $folders[0].Path | Should Be "C:\a [1]"
        $folders[1].Path | Should Be "D:\b"
        $folders[1].Enabled | Should Be $false
        (readSearchOption $path).UseRegex | Should Be $true
        writeTargetFolders @() $path
        @(getTargetFolders $path).Count | Should Be 0
    }
}

Describe "normalizeFolderPath" {
    It "前後の空白・引用符と末尾の \ を取り除き、ドライブ直下は \ を残す" {
        normalizeFolderPath '  "C:\data\"  ' | Should Be "C:\data"
        normalizeFolderPath "D:\" | Should Be "D:\"
        normalizeFolderPath "D:" | Should Be "D:\"
        normalizeFolderPath "\\server\share\" | Should Be "\\server\share"
    }

    It "/ を \ にそろえ、長いパス用の \\?\ ・ \\?\UNC\ を外す" {
        normalizeFolderPath "C:/data/見積" | Should Be "C:\data\見積"
        normalizeFolderPath "//server/share/見積" | Should Be "\\server\share\見積"
        normalizeFolderPath "\\?\C:\data\見積" | Should Be "C:\data\見積"
        normalizeFolderPath "\\?\UNC\server\share\見積" | Should Be "\\server\share\見積"
    }

    It "重なった \ ・ . ・ .. を解決する" {
        normalizeFolderPath "C:\data\\見積" | Should Be "C:\data\見積"
        normalizeFolderPath "C:\data\.\見積" | Should Be "C:\data\見積"
        normalizeFolderPath "C:\data\売上\..\見積" | Should Be "C:\data\見積"
        normalizeFolderPath "\\server\share\売上\..\見積" | Should Be "\\server\share\見積"
    }

    It "環境変数を展開する" {
        $env:WIN_GREP_TEST_FOLDER = "C:\data\見積"
        try {
            normalizeFolderPath "%WIN_GREP_TEST_FOLDER%" | Should Be "C:\data\見積"
            normalizeFolderPath "%WIN_GREP_TEST_FOLDER%\2024" | Should Be "C:\data\見積\2024"
        } finally {
            Remove-Item Env:\WIN_GREP_TEST_FOLDER
        }
    }

    It "相対パスは win_grep のフォルダからとみなす" {
        normalizeFolderPath "work\index" | Should Be "${rootDir}\work\index"
        normalizeFolderPath ".\work\index" | Should Be "${rootDir}\work\index"
    }

    It "パスとして解釈できない場合は書かれたとおりに扱う" {
        normalizeFolderPath "C:\data*" | Should Be "C:\data*"
        normalizeFolderPath "\\server" | Should Be "\\server"
        # \ ひとつで始まるパスは、書き間違えた UNC パスのことが多いため、今のドライブのパスに直さない
        normalizeFolderPath "\server\share" | Should Be "\server\share"
    }
}

Describe "getPathUnderFolder" {
    It "フォルダからの相対パスを返す（フォルダ自身は空。末尾の \ ・大文字と小文字は区別しない）" {
        getPathUnderFolder "C:\data\見積\2024\a.xlsx" "C:\data\見積" | Should Be "2024\a.xlsx"
        getPathUnderFolder "c:\DATA\見積\" "C:\data\見積" | Should Be ""
        getPathUnderFolder "D:\a.xlsx" "D:\" | Should Be "a.xlsx"
        getPathUnderFolder "\\server\share\a.xlsx" "\\server\share" | Should Be "a.xlsx"
    }

    It "フォルダの下でなければ `$null（フォルダ名の途中では一致しない）" {
        ($null -eq (getPathUnderFolder "C:\data\見積2\a.xlsx" "C:\data\見積")) | Should Be $true
        ($null -eq (getPathUnderFolder "D:\a.xlsx" "C:\data")) | Should Be $true
        ($null -eq (getPathUnderFolder "C:\data" "")) | Should Be $true
    }
}

Describe "getFolderPathAliases / testSameFolder" {
    # ドライブの割り当ては PC によって違うため、テストでは割り当てを差し替える
    $drives = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $drives["Z:"] = "\\server\share"
    $drives["X:"] = "C:\data"

    It "ネットワークドライブと UNC パスを行き来する" {
        @(getFolderPathAliases "Z:\見積" $drives) | Should Be @("Z:\見積", "\\server\share\見積")
        @(getFolderPathAliases "\\server\share\見積" $drives) | Should Be @("\\server\share\見積", "Z:\見積")
    }

    It "subst のドライブと割り当て元のフォルダを行き来する" {
        @(getFolderPathAliases "X:\見積" $drives) | Should Be @("X:\見積", "C:\data\見積")
        @(getFolderPathAliases "C:\data\見積" $drives) | Should Be @("C:\data\見積", "X:\見積")
    }

    It "ドライブ直下・共有直下" {
        @(getFolderPathAliases "Z:\" $drives) | Should Be @("Z:\", "\\server\share")
        @(getFolderPathAliases "\\server\share" $drives) | Should Be @("\\server\share", "Z:\")
    }

    It "割り当ての無いパスは自身だけ" {
        @(getFolderPathAliases "D:\見積" $drives) | Should Be @("D:\見積")
    }

    It "書き方が違っても同じフォルダと分かる" {
        testSameFolder "Z:\見積" "\\server\share\見積" $drives | Should Be $true
        testSameFolder "\\server\share\見積\" "Z:\見積" $drives | Should Be $true
        testSameFolder "X:\見積" "C:\data\見積" $drives | Should Be $true
        testSameFolder "C:\data\見積" "X:\見積" $drives | Should Be $true
        testSameFolder "Z:\見積" "Z:\売上" $drives | Should Be $false
        testSameFolder "C:\見積" "\\server\share\見積" $drives | Should Be $false
    }

    It "この PC のドライブの割り当てを使っても例外にならない" {
        @(getFolderPathAliases "${rootDir}\work")[0] | Should Be "${rootDir}\work"
        testSameFolder "${rootDir}\work" "${rootDir}\work\" | Should Be $true
    }

    It "入れ子のフォルダが分かる（testFolderUnder）" {
        testFolderUnder "C:\data\見積\2024" "C:\data\見積" $drives | Should Be $true
        testFolderUnder "C:\data\見積" "C:\data\見積" $drives | Should Be $true   # 同じフォルダも「中」とする
        testFolderUnder "C:\data\見積" "C:\data\見積\2024" $drives | Should Be $false
        testFolderUnder "C:\data\見積2" "C:\data\見積" $drives | Should Be $false  # 名前の先頭が同じだけ
        testFolderUnder "Z:\見積\2024" "\\server\share\見積" $drives | Should Be $true
        testFolderUnder "C:\data\見積\2024" "X:\見積" $drives | Should Be $true
        testFolderUnder "" "C:\data" $drives | Should Be $false
        testFolderUnder "C:\data" "" $drives | Should Be $false
    }
}

Describe "testOfficeFile" {
    It "Excel・Word・PowerPoint のファイルを見分ける（大文字・小文字は区別しない）" {
        testOfficeFile "見積.xlsx" | Should Be $true
        testOfficeFile "報告.DOCM" | Should Be $true
        testOfficeFile "資料.ppt" | Should Be $true
    }

    It "Office 以外のファイル・拡張子の無い名前は false" {
        testOfficeFile "メモ.txt" | Should Be $false
        testOfficeFile "データ" | Should Be $false
        testOfficeFile "" | Should Be $false
    }
}

Describe "getParentFolderPath" {
    It "1つ上のフォルダを返す" {
        getParentFolderPath "C:\data\見積\2025" | Should Be "C:\data\見積"
        getParentFolderPath "C:\data" | Should Be "C:\"
        getParentFolderPath "\\server\share\見積\2025" | Should Be "\\server\share\見積"
        getParentFolderPath "\\server\share\見積" | Should Be "\\server\share"
    }

    It "これ以上たどれないフォルダ（ドライブ直下・共有フォルダ直下）は空文字列" {
        getParentFolderPath "C:\" | Should Be ""
        getParentFolderPath "C:" | Should Be ""
        getParentFolderPath "\\server\share" | Should Be ""
        getParentFolderPath "\\server" | Should Be ""
        getParentFolderPath "" | Should Be ""
    }
}

Describe "joinFolderPath" {
    It "ドライブ直下・共有フォルダ直下でも \ が重ならない" {
        joinFolderPath "C:\" "data" | Should Be "C:\data"
        joinFolderPath "C:\data" "見積" | Should Be "C:\data\見積"
        joinFolderPath "\\server\share" "見積" | Should Be "\\server\share\見積"
    }
}

Describe "getFolderEntries / testHasSubFolders" {
    $root = "$TestDrive\folderEntries"
    New-Item -ItemType Directory -Path "$root\B社" -Force | Out-Null
    New-Item -ItemType Directory -Path "$root\A社" -Force | Out-Null
    New-Item -ItemType Directory -Path "$root\A社\2025" -Force | Out-Null
    New-Item -ItemType Directory -Path "$root\_隠しフォルダ" -Force | Out-Null
    (Get-Item "$root\_隠しフォルダ").Attributes = "Directory, Hidden"
    Set-Content "$root\見積.xlsx" "x" -Encoding UTF8
    Set-Content "$root\メモ.txt" "x" -Encoding UTF8
    Set-Content "$root\隠し.txt" "x" -Encoding UTF8
    (Get-Item "$root\隠し.txt").Attributes = "Hidden"

    It "フォルダを先に、それぞれ名前順で返す" {
        $result = getFolderEntries $root
        # 名前順は Windows の並び（カタカナが漢字より先）
        @($result.Entries | ForEach-Object { $_.Name }) -join "," | Should Be "A社,B社,メモ.txt,見積.xlsx"
    }

    It "フォルダ数と Office ファイル数を返す" {
        $result = getFolderEntries $root
        $result.FolderCount | Should Be 2
        $result.OfficeCount | Should Be 1
        $result.Truncated | Should Be $false
        $result.Error | Should Be ""
    }

    It "隠し・システムのフォルダとファイルは返さない（エクスプローラーの既定と同じ）" {
        $names = @((getFolderEntries $root).Entries | ForEach-Object { $_.Name })
        $names -contains "_隠しフォルダ" | Should Be $false
        $names -contains "隠し.txt" | Should Be $false
    }

    It "パスと種別を返す" {
        $entry = @((getFolderEntries $root).Entries | Where-Object { $_.Name -eq "見積.xlsx" })[0]
        $entry.Path | Should Be "$root\見積.xlsx"
        $entry.IsFolder | Should Be $false
        $entry.IsOffice | Should Be $true
        $entry.Updated | Should Not BeNullOrEmpty
    }

    It "foldersOnly はフォルダだけを返す（ツリーの読み込み用）" {
        $result = getFolderEntries $root -foldersOnly
        @($result.Entries | ForEach-Object { $_.Name }) -join "," | Should Be "A社,B社"
        $result.OfficeCount | Should Be 0
    }

    It "件数が多いときは打ち切り、Truncated を立てる" {
        $result = getFolderEntries $root 2
        @($result.Entries).Count | Should Be 2
        $result.Truncated | Should Be $true
    }

    It "開けないフォルダは Error に理由を入れる（一覧は空）" {
        $result = getFolderEntries "$root\ありません"
        @($result.Entries).Count | Should Be 0
        $result.Error | Should Match "見つかりません"
    }

    It "フォルダを指定しなければ Error を返す" {
        (getFolderEntries "").Error | Should Match "指定してください"
    }

    It "サブフォルダがあるかを返す（ツリーの ▷ の判定）" {
        testHasSubFolders $root | Should Be $true
        testHasSubFolders "$root\B社" | Should Be $false
        testHasSubFolders "$root\ありません" | Should Be $false
    }
}

Describe "newIndexName / assignIndexNames" {
    It "フォルダ名（ドライブ直下はドライブ名、UNC は共有名）を使い、重複すれば (2) を付ける" {
        $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        newIndexName "C:\data\見積" $used | Should Be "見積"
        newIndexName "D:\" $used | Should Be "D"
        newIndexName "\\server\share" $used | Should Be "share"
        [void]$used.Add("見積")
        [void]$used.Add("見積(2)")
        newIndexName "E:\見積" $used | Should Be "見積(3)"
    }

    It '使用中の名前が $null・配列・空でも落ちずに名前を作る' {
        # インデックスが 1 つも無いとき、呼び出し側から $null が渡ることがある（画面の新規作成ダイアログ）
        newIndexName "C:\data\sample" $null | Should Be "sample"
        newIndexName "C:\data\見積" | Should Be "見積"
        newIndexName "C:\data\見積" @() | Should Be "見積"
        newIndexName "C:\data\見積" @("見積") | Should Be "見積(2)"
        # インデックスが 1 件のときは、集合が展開されて文字列 1 個で渡ることがある
        newIndexName "C:\data\見積" "見積" | Should Be "見積(2)"
        newIndexName "C:\data\見" "見積" | Should Be "見"
        # 前方一致・大文字小文字違いで取り違えない
        newIndexName "C:\data\見積" @("見積書") | Should Be "見積"
        newIndexName "C:\data\sample" @("SAMPLE") | Should Be "sample(2)"
    }

    It "設定に名前があればそれを使い、フォルダの場所が変わっても同じ名前のままにする" {
        $targets = @(
            [pscustomobject]@{ Name = "見積"; Path = "\server\新しい場所\見積書"; Enabled = $true },
            [pscustomobject]@{ Name = ""; Path = "F:\売上"; Enabled = $true }
        )
        # 前回は別の場所だったが、名前が同じなので同じインデックスとして扱う
        $previous = @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" })
        $folders = @(assignIndexNames $targets $previous)
        $folders.Count | Should Be 2
        $folders[0].Name | Should Be "見積"
        $folders[0].Path | Should Be "\server\新しい場所\見積書"
        $folders[1].Name | Should Be "売上"
    }

    It "名前が無ければ、前回の変換一覧の同じフォルダの名前を使い、無ければフォルダ名から作る" {
        $targets = @(
            [pscustomobject]@{ Name = ""; Path = "C:\data\見積"; Enabled = $false },
            [pscustomobject]@{ Name = ""; Path = "E:\new\見積"; Enabled = $true },
            [pscustomobject]@{ Name = ""; Path = "F:\売上"; Enabled = $true }
        )
        $previous = @(
            [pscustomobject]@{ Path = "c:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "G:\削除した\報告"; Name = "報告" }
        )
        $folders = @(assignIndexNames $targets $previous)
        $folders[0].Name | Should Be "見積"
        $folders[0].Enabled | Should Be $false
        # 前回の名前（削除したフォルダの名前を含む）と重複しない名前を付ける
        $folders[1].Name | Should Be "見積(2)"
        $folders[2].Name | Should Be "売上"
    }

    It "設定にある名前は、ほかのフォルダの名前には使わない" {
        $targets = @(
            [pscustomobject]@{ Name = ""; Path = "E:\新\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true }
        )
        $folders = @(assignIndexNames $targets @())
        $folders[0].Name | Should Be "見積(2)"
        $folders[1].Name | Should Be "見積"
    }
}

Describe "splitIndexRelPath" {
    It "先頭のインデックス名と残りに分ける（残りの \ や 2 はそのまま）" {
        $parts = splitIndexRelPath "excel\2024\見積\A社2.xlsx"
        $parts.Name | Should Be "excel"
        $parts.Rest | Should Be "2024\見積\A社2.xlsx"
        $parts = splitIndexRelPath "excel"
        $parts.Name | Should Be "excel"
        $parts.Rest | Should Be ""
    }
}

Describe "describeConvertError" {
    function newComError([string]$message, [string]$code) {
        return New-Object System.Runtime.InteropServices.COMException($message, [Convert]::ToInt32($code, 16))
    }

    It "パスワード付きのファイルは、Officeアプリの分かりにくいメッセージを付けずに原因だけを返す" {
        $expected = "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは変換できません）"
        describeConvertError (newComError "入力したパスワードが間違っています。CapsLock キーの状態に注意して…" "800A03EC") | Should Be $expected
        describeConvertError (newComError "パスワードが正しくありません。文書を開けません。 (C:\Users\a\AppData\...\source.doc)" "800A1520") | Should Be $expected
        describeConvertError (newComError "Presentations.Open : 読み取りパスワードをもう一度入力してください(&P):" "80004005") | Should Be $expected
    }

    It "メソッド呼び出しの例外は中の例外のメッセージを使う" {
        $inner = newComError "Excel でファイル 'a.xlsx' を開くことができません。ファイル形式またはファイル拡張子が正しくありません。" "800A03EC"
        $outer = New-Object System.Management.Automation.MethodInvocationException('"7" 個の引数を指定して "Open" を呼び出し中に例外が発生しました', $inner)
        describeConvertError $outer | Should Be "ファイルが壊れているか、拡張子と中身の形式が一致していません（詳細: Excel でファイル 'a.xlsx' を開くことができません。ファイル形式またはファイル拡張子が正しくありません。）"
    }

    It "スクリプト自身が throw したメッセージはそのまま返す" {
        $exception = $null
        try { throw "ファイルが壊れているか、PowerPointのファイルではありません。" } catch { $exception = $_.Exception }
        describeConvertError $exception | Should Be "ファイルが壊れているか、PowerPointのファイルではありません。"
    }

    It "使用中・アクセス権なし・ファイルなしは原因を付けて元のメッセージを詳細にする" {
        $locked = New-Object System.IO.IOException("別のプロセスで使用されているため、アクセスできません。", [Convert]::ToInt32("80070020", 16))
        describeConvertError $locked | Should Be "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再変換してください）（詳細: 別のプロセスで使用されているため、アクセスできません。）"
        describeConvertError (New-Object System.UnauthorizedAccessException("アクセスが拒否されました。")) | Should Match "^ファイルを読むアクセス権がありません（詳細: アクセスが拒否されました。）$"
        describeConvertError (New-Object System.IO.FileNotFoundException("見つかりません。")) | Should Match "^ファイルが見つかりません（"
    }

    It "Officeアプリの異常終了・応答なし・起動失敗は HRESULT で判断する" {
        describeConvertError (newComError "RPC サーバーを利用できません。" "800706BA") | Should Match "^Officeアプリが異常終了したか、内部でエラーが発生しました（.*（詳細: RPC サーバーを利用できません。）$"
        describeConvertError (newComError "呼び出し先が呼び出しを拒否しました。" "80010001") | Should Match "^Officeアプリが応答しませんでした"
        describeConvertError (newComError "クラスが登録されていません" "80040154") | Should Match "^Officeアプリ（Excel・Word・PowerPoint）を起動できませんでした"
    }

    It "メモリ不足（巨大なシート）は原因を付けて元のメッセージを詳細にする" {
        $inner = New-Object System.OutOfMemoryException("Exception of type 'System.OutOfMemoryException' was thrown.")
        $outer = New-Object System.Management.Automation.MethodInvocationException('"1" 個の引数を指定して "ReadAllText" を呼び出し中に例外が発生しました', $inner)
        describeConvertError $outer | Should Match "^シート・文書が大きすぎて変換できません（メモリが不足しました）（詳細: "
    }

    It "原因が分からないものは元のメッセージ（改行は詰める）、メッセージが無ければエラーコードを返す" {
        describeConvertError (newComError "予期しない`r`nエラーです。" "800A03EC") | Should Be "予期しない エラーです。"
        describeConvertError (New-Object System.Exception(" ")) | Should Match "^エラーコード 0x[0-9A-F]{8}$"
    }
}

Describe "getConversionState" {
    It "失敗したファイルの行を、変換日時の新しい順で FailedRows に返す" {
        $path = "$TestDrive\status_state.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(
            (newStatusRow "data\済.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/19 10:00:00"),
            (newStatusRow "data\古い失敗.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/18 09:00:00" "原因A"),
            (newStatusRow "data\新しい失敗.docx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 11:00:00" "原因B"),
            (newStatusRow "data\未変換.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        $state = getConversionState -path $path
        $state.Done | Should Be 1
        $state.Pending | Should Be 1
        $state.Failed | Should Be 2
        $state.FailedRows.Count | Should Be 2
        $state.FailedRows[0].相対パス | Should Be "data\新しい失敗.docx"
        $state.FailedRows[0].エラー | Should Be "原因B"
        $state.FailedRows[1].相対パス | Should Be "data\古い失敗.xlsx"
    }

    It "変換一覧が無ければ FailedRows は空" {
        $state = getConversionState -path "$TestDrive\none.tsv"
        $state.Exists | Should Be $false
        @($state.FailedRows).Count | Should Be 0
    }
}

Describe "testIndexName" {
    It "使える名前なら空文字列を返す" {
        testIndexName "営業部 2025" | Should Be ""
    }

    It "空なら入力を促す" {
        testIndexName "" | Should Match "入力してください"
    }

    It "前後に空白があれば使えない" {
        testIndexName " 営業" | Should Match "前後に空白"
        testIndexName "営業 " | Should Match "前後に空白"
    }

    It "ファイル名に使えない文字があれば使えない" {
        testIndexName "営業\部" | Should Match "使えない文字"
        testIndexName "営業:部" | Should Match "使えない文字"
    }

    It "末尾が . なら使えない" {
        testIndexName "営業." | Should Match "最後に \."
    }

    It "Windows の予約語は使えない（大文字・小文字を区別しない）" {
        testIndexName "con" | Should Match "使えない名前"
        testIndexName "LPT1" | Should Match "使えない名前"
    }

    It "ほかのインデックスと同じ名前は使えない（大文字・小文字を区別しない）" {
        testIndexName "Sales" @("sales", "tech") | Should Match "ほかのインデックスが使っています"
    }

    It "長すぎる名前は使えない" {
        testIndexName ("あ" * 256) | Should Match "長すぎます"
    }
}

Describe "getIndexStats" {
    It "インデックス名ごとに件数と最終変換日時を集計する" {
        $path = "$TestDrive\stats.tsv"
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "営業\b.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 11:00:00" "原因"),
            (newStatusRow "営業\c.docx" "2025/01/10 12:34:56" "1" $stateNew),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateDone "2" "2026/09/17 09:00:00")
        ) $path

        $stats = getIndexStats (readStatusFile $path).Rows
        $stats["営業"].Total | Should Be 3
        $stats["営業"].Done | Should Be 1
        $stats["営業"].Failed | Should Be 1
        $stats["営業"].Pending | Should Be 1
        $stats["営業"].LastConverted | Should Be "2026/09/19 11:00:00"
        $stats["技術"].Total | Should Be 1
        $stats["技術"].LastConverted | Should Be "2026/09/17 09:00:00"
    }

    It "行が無ければ空を返す" {
        (getIndexStats $null).Count | Should Be 0
    }

    It "getConversionState からも集計を取れる" {
        $path = "$TestDrive\stats_state.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $path

        (getConversionState -path $path).IndexStats["営業"].Total | Should Be 1
    }
}

Describe "renameIndex" {
    It "インデックスのフォルダと変換一覧の記録の名前を変え、中身はそのまま残す" {
        $dir = "$TestDrive\rename\index"
        $path = "$TestDrive\rename\変換一覧.tsv"
        New-Item -ItemType Directory -Path "$dir\営業\a.xlsx" -Force | Out-Null
        Set-Content -LiteralPath "$dir\営業\a.xlsx\Sheet1.tsv" -Value "本文" -Encoding UTF8
        New-Item -ItemType Directory -Path "$dir\技術" -Force | Out-Null
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        renameIndex "営業" "営業部" $dir $path

        Test-Path "$dir\営業" | Should Be $false
        Get-Content -LiteralPath "$dir\営業部\a.xlsx\Sheet1.tsv" | Should Be "本文"

        $status = readStatusFile $path
        $status.Folders[0].Name | Should Be "営業部"
        $status.Folders[0].Path | Should Be "C:\data"
        $status.Rows.ContainsKey("営業部\a.xlsx") | Should Be $true
        $status.Rows["営業部\a.xlsx"].状態 | Should Be $stateDone
        # ほかのインデックスはそのまま
        $status.Folders[1].Name | Should Be "技術"
        $status.Rows.ContainsKey("技術\d.pptx") | Should Be $true
    }

    It "インデックスのフォルダがまだ無くても、変換一覧の記録は変える" {
        $dir = "$TestDrive\rename2\index"
        $path = "$TestDrive\rename2\変換一覧.tsv"
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        renameIndex "営業" "営業部" $dir $path

        (readStatusFile $path).Rows.ContainsKey("営業部\a.xlsx") | Should Be $true
    }

    It "同じ名前のフォルダが既にあれば例外にする" {
        $dir = "$TestDrive\rename3\index"
        $path = "$TestDrive\rename3\変換一覧.tsv"
        New-Item -ItemType Directory -Path "$dir\営業" -Force | Out-Null
        New-Item -ItemType Directory -Path "$dir\技術" -Force | Out-Null

        { renameIndex "営業" "技術" $dir $path } | Should Throw
    }
}

Describe "removeIndex" {
    It "インデックスのフォルダと変換一覧の記録を削除し、ほかのインデックスは残す" {
        $dir = "$TestDrive\remove\index"
        $path = "$TestDrive\remove\変換一覧.tsv"
        New-Item -ItemType Directory -Path "$dir\営業\a.xlsx" -Force | Out-Null
        Set-Content -LiteralPath "$dir\営業\a.xlsx\Sheet1.tsv" -Value "本文" -Encoding UTF8
        New-Item -ItemType Directory -Path "$dir\技術" -Force | Out-Null
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        removeIndex "営業" $dir $path

        Test-Path "$dir\営業" | Should Be $false
        Test-Path "$dir\技術" | Should Be $true

        $status = readStatusFile $path
        $status.Folders.Count | Should Be 1
        $status.Folders[0].Name | Should Be "技術"
        $status.Rows.Count | Should Be 1
        $status.Rows.ContainsKey("技術\d.pptx") | Should Be $true
    }

    It "名前が空なら何もしない" {
        $path = "$TestDrive\remove2\変換一覧.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        removeIndex "" "$TestDrive\remove2\index" $path

        (readStatusFile $path).Rows.Count | Should Be 1
    }
}

Describe "getIndexNameMap / resolveSourcePath" {
    It "インデックス名から変換対象フォルダを引き、元のファイルのパスを返す" {
        $path = "$TestDrive\status_map.tsv"
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "D:\"; Name = "D" }
        ) @() $path
        $map = getIndexNameMap $path
        $map["見積"] | Should Be "C:\data\見積"

        $root = "C:\win_grep\work\index"
        $maps = @{ $root = $map; "C:\win_grep\work" = $map }
        $hit = [pscustomobject]@{ Root = $root; RelDir = "見積\2024"; Book = "A社.xlsx" }
        resolveSourcePath $hit $maps | Should Be "C:\data\見積\2024\A社.xlsx"
        $hit = [pscustomobject]@{ Root = $root; RelDir = "見積\2024\見積\2"; Book = "A社.xlsx" }
        resolveSourcePath $hit $maps | Should Be "C:\data\見積\2024\見積\2\A社.xlsx"
        $hit = [pscustomobject]@{ Root = $root; RelDir = "d"; Book = "直下.xlsx" }
        resolveSourcePath $hit $maps | Should Be "D:\直下.xlsx"
        $hit = [pscustomobject]@{ Root = "$TestDrive\none_index"; RelDir = "不明\x"; Book = "a.xlsx" }
        resolveSourcePath $hit @{} | Should Be $null
    }

    It "設定のインデックス名の場所を、元のフォルダ.txt・変換一覧より優先する" {
        # 「名前」と「今の置き場所」を設定で分けて持つため、フォルダを移したら設定の場所だけを見る
        $dir = "$TestDrive\優先\index"
        $settings = "$TestDrive\優先\setting.config"
        $status = "$TestDrive\優先\status.tsv"
        [void](New-Item -ItemType Directory -Path "$dir\見積" -Force)
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\作った時の場所\見積"; Name = "見積" }) $dir
        writeStatusFile @([pscustomobject]@{ Path = "C:\変換一覧の場所\見積"; Name = "見積" }) @() $status
        writeTargetFolders @([pscustomobject]@{ Name = "見積"; Path = "\server\今の場所\見積"; Enabled = $true }) $settings

        $map = getSourceFolderMap $dir $status $settings
        $map["見積"] | Should Be "\server\今の場所\見積"

        # 変換しないインデックス（indexSources）も同じように優先する
        setIndexSourceFolder "営業" "E:\今の営業" $settings
        (getSourceFolderMap $dir $status $settings)["営業"] | Should Be "E:\今の営業"
    }

    It "既定のインデックスは変換一覧の記録を使う" {
        [System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
        $status = "$TestDrive\status_default.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\新\見積"; Name = "見積" }) @() $status
        $map = getSourceFolderMap (Resolve-Path -LiteralPath $indexDir).ProviderPath $status "$TestDrive\設定なし.config"
        $map["見積"] | Should Be "C:\新\見積"
    }
}

Describe "writeSourceFolderFile / readSourceFolderFile / getSourceLocation" {
    It "インデックス名と変換対象フォルダの対応を、各インデックスのフォルダに書き出して読み込む（説明の行は無視する）" {
        $dir = "$TestDrive\copied[1]\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        [void][System.IO.Directory]::CreateDirectory("$dir\D")
        writeSourceFolderFile @(
            [pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "D:\"; Name = "D" }
        ) $dir
        # インデックスのフォルダ直下（全インデックス分）には書かない
        Test-Path -LiteralPath (Join-Path $dir ${sourceFolderFileName}) | Should Be $false
        (readSourceFolderFile "$dir\見積")["見積"] | Should Be "C:\data\見積"
        (readSourceFolderFile "$dir\D")["d"] | Should Be "D:\"
        (readSourceFolderFile "$TestDrive\none_dir").Count | Should Be 0
    }

    It "別の場所にコピーしたインデックスでも、元のフォルダ.txt から元の場所が分かる" {
        $dir = "$TestDrive\別PC\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir
        $hit = [pscustomobject]@{ Root = $dir; RelDir = "見積\2024"; Book = "A社.xlsx" }
        $location = getSourceLocation $hit
        $location.Known | Should Be $true
        $location.Folder | Should Be "C:\data\見積"
        $location.Rest | Should Be "2024"
        resolveSourcePath $hit @{} | Should Be "C:\data\見積\2024\A社.xlsx"
    }

    It "インデックスのフォルダの中に 元のフォルダ.txt を書き、そのフォルダだけをコピーしても元の場所が分かる" {
        $dir = "$TestDrive\作った PC\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        [void][System.IO.Directory]::CreateDirectory("$dir\営業")
        writeSourceFolderFile @(
            [pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" }
        ) $dir
        (readSourceFolderFile "$dir\見積").Count | Should Be 1
        (readSourceFolderFile "$dir\見積")["見積"] | Should Be "C:\data\見積"

        # <インデックス名> のフォルダだけを別の場所（ほかの PC の work\index 直下など）へコピーした場合
        $other = "$TestDrive\別 PC\index"
        [void][System.IO.Directory]::CreateDirectory($other)
        Copy-Item -LiteralPath "$dir\見積" -Destination "$other\見積" -Recurse
        $hit = [pscustomobject]@{ Root = $other; RelDir = "見積\2024"; Book = "A社.xlsx" }
        resolveSourcePath $hit @{} | Should Be "C:\data\見積\2024\A社.xlsx"
    }

    It "インデックスのフォルダがまだ無ければ、その中には書かない" {
        $dir = "$TestDrive\未変換\index"
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir
        Test-Path -LiteralPath "$dir\見積" | Should Be $false
        (readSourceFolderFile $dir).Count | Should Be 0
    }

    It "以前の版が書いた、インデックスのフォルダ直下の 元のフォルダ.txt は各フォルダへ移して消す" {
        $dir = "$TestDrive\移行\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        [void][System.IO.Directory]::CreateDirectory("$dir\やめた")
        # 以前の版と同じ形式（全インデックス分を直下に 1 ファイル）
        writeListFile (Join-Path $dir ${sourceFolderFileName}) @(
            "# 説明",
            "見積`tC:\旧\見積",
            "やめた`tC:\data\やめた")

        # 変換対象フォルダから外したインデックス（やめた）の記録も残す
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir

        Test-Path -LiteralPath (Join-Path $dir ${sourceFolderFileName}) | Should Be $false
        (readSourceFolderFile "$dir\見積")["見積"] | Should Be "C:\data\見積"
        (readSourceFolderFile "$dir\やめた")["やめた"] | Should Be "C:\data\やめた"
    }

    It "インデックス名のフォルダを検索対象にした場合は、そのフォルダの 元のフォルダ.txt を使う" {
        $dir = "$TestDrive\別PC2\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir
        $hit = [pscustomobject]@{ Root = "$dir\見積"; RelDir = "2024"; Book = "A社.xlsx" }
        resolveSourcePath $hit @{} | Should Be "C:\data\見積\2024\A社.xlsx"
        $hit = [pscustomobject]@{ Root = "$dir\見積"; RelDir = ""; Book = "直下.xlsx" }
        resolveSourcePath $hit @{} | Should Be "C:\data\見積\直下.xlsx"
    }

    It "元の場所が分からなければ Known = false・Folder = 空とし、インデックス名と相対フォルダを返す" {
        $hit = [pscustomobject]@{ Root = "$TestDrive\記録なし\index\"; RelDir = "営業\2024"; Book = "a.xlsx" }
        $location = getSourceLocation $hit
        $location.Known | Should Be $false
        $location.Name | Should Be "営業"
        $location.Folder | Should Be ""
        $location.Rest | Should Be "2024"
        resolveSourcePath $hit @{} | Should Be $null
    }

    It "読んだ対応はキャッシュに入れ、次からはファイルを読まない" {
        $dir = "$TestDrive\cache\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir
        $maps = @{}
        $hit = [pscustomobject]@{ Root = $dir; RelDir = "見積"; Book = "a.xlsx" }
        resolveSourcePath $hit $maps | Should Be "C:\data\見積\a.xlsx"
        Remove-Item -LiteralPath (Join-Path "$dir\見積" ${sourceFolderFileName})
        resolveSourcePath $hit $maps | Should Be "C:\data\見積\a.xlsx"
    }
}

Describe "joinSourcePath" {
    It "フォルダ・相対フォルダ・ファイル名をつなぐ" {
        joinSourcePath "C:\data\" "2024\見積" "a.xlsx" | Should Be "C:\data\2024\見積\a.xlsx"
        joinSourcePath "D:\" "" "a.xlsx" | Should Be "D:\a.xlsx"
        joinSourcePath "C:\data" "2024" | Should Be "C:\data\2024"
    }
}

Describe "getFolderLeafName" {
    It "フォルダ名を返す（ドライブ直下はドライブ名、UNC は共有名、末尾の \ は無視）" {
        getFolderLeafName "C:\data\見積\" | Should Be "見積"
        getFolderLeafName "D:\" | Should Be "D"
        getFolderLeafName "\server\share" | Should Be "share"
    }
}

Describe "indexSources / setIndexSourceFolder" {
    It "インデックス名に対する元のフォルダを記録し、読み込める" {
        $path = "$TestDrive\sources[1].config"
        setIndexSourceFolder "営業" "\server\営業\" $path
        setIndexSourceFolder "見積" "E:\見積" $path
        $sources = @(readIndexSources $path)
        $sources.Count | Should Be 2
        $sources[0].Name | Should Be "営業"
        $sources[0].Path | Should Be "\server\営業"
        $sources[1].Path | Should Be "E:\見積"

        # 同じ名前をもう一度記録すると、場所を書き換える（1 つの名前につき 1 か所）
        setIndexSourceFolder "営業" "D:\新しい営業" $path
        $sources = @(readIndexSources $path)
        $sources.Count | Should Be 2
        @($sources | Where-Object { $_.Name -eq "営業" })[0].Path | Should Be "D:\新しい営業"
    }

    It "変換対象フォルダにある名前なら、そのフォルダの場所を書き換える" {
        $path = "$TestDrive\sources_target.config"
        writeTargetFolders @(
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $false }
        ) $path
        setIndexSourceFolder "見積" "\server\移動先\見積" $path

        $folders = @(getTargetFolders $path)
        $folders[0].Name | Should Be "見積"
        $folders[0].Path | Should Be "\server\移動先\見積"
        $folders[0].Enabled | Should Be $true
        $folders[1].Path | Should Be "C:\data\営業"
        # 変換対象フォルダを書き換えたので、indexSources には入れない
        @(readIndexSources $path).Count | Should Be 0
    }

    It "名前・フォルダが空なら何もしない" {
        $path = "$TestDrive\sources_empty.config"
        setIndexSourceFolder "" "C:\a" $path
        setIndexSourceFolder "営業" "  " $path
        @(readIndexSources $path).Count | Should Be 0
    }
}

Describe "findMovedSource" {
    $moved = "$TestDrive\移動先\見積 [新]"
    [System.IO.Directory]::CreateDirectory("$moved\2024\A社") | Out-Null
    [System.IO.File]::WriteAllText("$moved\2024\A社\見積.xlsx", "")

    It "インデックスの元のフォルダに当たるフォルダを選んだ場合は、そのフォルダを Root にする" {
        $found = findMovedSource $moved "2024\A社" "見積.xlsx"
        $found.Path | Should Be "$moved\2024\A社\見積.xlsx"
        $found.Root | Should Be $moved
    }

    It "ファイルのあるフォルダを選んだ場合は、相対フォルダの分だけ上を Root にする" {
        $found = findMovedSource "$moved\2024\A社\" "2024\A社" "見積.xlsx"
        $found.Path | Should Be "$moved\2024\A社\見積.xlsx"
        $found.Root | Should Be $moved
    }

    It "途中のフォルダを選んだ場合も Root は同じになる" {
        (findMovedSource "$moved\2024" "2024\A社" "見積.xlsx").Root | Should Be $moved
    }

    It "相対フォルダが無ければ、選んだフォルダが Root になる" {
        [System.IO.File]::WriteAllText("$moved\直下.xlsx", "")
        $found = findMovedSource $moved "" "直下.xlsx"
        $found.Path | Should Be "$moved\直下.xlsx"
        $found.Root | Should Be $moved
    }

    It "見つからなければ `$null" {
        findMovedSource "$TestDrive\移動先" "2024\B社" "見積.xlsx" | Should Be $null
    }
}

Describe "getSearchIndexes" {
    It "インデックスのフォルダが無ければ空" {
        @(getSearchIndexes "$TestDrive\無いフォルダ\index").Count | Should Be 0
    }

    It "work\index 直下のフォルダをインデックス 1 件として返し、元のフォルダも返す" {
        $dir = "$TestDrive\一覧\index"
        $settings = "$TestDrive\一覧\setting.config"
        foreach ($name in @("見積", "営業")) {
            [void](New-Item -ItemType Directory -Path "$dir\$name" -Force)
        }
        writeTargetFolders @(
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $true }) $settings

        $indexes = @(getSearchIndexes $dir "$TestDrive\変換一覧なし.tsv" $settings)
        $indexes.Count | Should Be 2
        # ［1 インデックス作成］の一覧と同じ並び
        $indexes[0].Name | Should Be "見積"
        $indexes[0].SourcePath | Should Be "C:\data\見積"
        $indexes[0].Path | Should Be (Resolve-Path -LiteralPath "$dir\見積").ProviderPath
        $indexes[1].Name | Should Be "営業"
        $indexes[1].SourcePath | Should Be "C:\data\営業"
    }

    It "一覧に無いインデックス（コピーしたものなど）も名前順で後ろに並べる" {
        $dir = "$TestDrive\コピー\index"
        $settings = "$TestDrive\コピー\setting.config"
        foreach ($name in @("報告書", "あとから", "見積")) {
            [void](New-Item -ItemType Directory -Path "$dir\$name" -Force)
        }
        writeTargetFolders @([pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true }) $settings

        $indexes = @(getSearchIndexes $dir "$TestDrive\変換一覧なし.tsv" $settings)
        @($indexes | ForEach-Object { $_.Name }) -join "," | Should Be "見積,あとから,報告書"
        # 元のフォルダが分からないものは空
        $indexes[1].SourcePath | Should Be ""
    }

    It "一覧にも変換一覧にも無いインデックスは、そのフォルダの 元のフォルダ.txt から元のフォルダを読む" {
        $dir = "$TestDrive\コピー2\index"
        $settings = "$TestDrive\コピー2\setting.config"
        [void](New-Item -ItemType Directory -Path "$dir\営業" -Force)
        # ほかの PC で作ったインデックスをフォルダごとコピーした状態
        writeSourceFolderFile @([pscustomobject]@{ Path = "\\server\営業"; Name = "営業" }) $dir

        $indexes = @(getSearchIndexes $dir "$TestDrive\変換一覧なし.tsv" $settings)
        $indexes.Count | Should Be 1
        $indexes[0].Name | Should Be "営業"
        $indexes[0].SourcePath | Should Be "\\server\営業"
    }
}

Describe "readSearchExcludes / writeSearchExcludes" {
    It "設定が無ければ空（すべて検索する）" {
        @(readSearchExcludes "$TestDrive\none_excludes.json").Count | Should Be 0
    }

    It "チェックを外したフォルダを保存し、ほかの設定は変えない" {
        $path = "$TestDrive\excludes.json"
        writeSearchOption @{ UseRegex = $true } $path
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "D:\index1\見積\2024\"; Subfolders = $true },
            [pscustomobject]@{ Path = "D:\index1\見積"; Subfolders = $false },
            [pscustomobject]@{ Path = ""; Subfolders = $true }) $path
        $excludes = @(readSearchExcludes $path)
        $excludes.Count | Should Be 2
        $excludes[0].Path | Should Be "D:\index1\見積\2024"
        $excludes[0].Subfolders | Should Be $true
        $excludes[1].Subfolders | Should Be $false
        (readSearchOption $path).UseRegex | Should Be $true
    }

    It "空で保存すると空になる" {
        $path = "$TestDrive\excludes_empty.json"
        writeSearchExcludes @([pscustomobject]@{ Path = "D:\a"; Subfolders = $true }) $path
        writeSearchExcludes @() $path
        @(readSearchExcludes $path).Count | Should Be 0
    }
}

Describe "getIndexNameMap（見出し行まで読む）" {
    It "変換対象フォルダの行だけを読み、見出し行の後は読まない" {
        $path = "$TestDrive\name_map.tsv"
        writeListFile $path @(
            "${statusFolderKey}`tC:\data\見積`t見積",
            "${statusFolderKey}`tC:\old",
            ($statusColumns -join "`t"),
            "${statusFolderKey}`tC:\x`tx")
        $map = getIndexNameMap $path
        $map.Count | Should Be 1
        $map["見積"] | Should Be "C:\data\見積"
        (getIndexNameMap "$TestDrive\none_status.tsv").Count | Should Be 0
    }
}

Describe "getIndexTsvCounts / testIndexComplete" {
    function newTestIndex {
        # テスト用のインデックス（work\index 相当）を作る
        param ([string]$dir)

        [System.IO.Directory]::CreateDirectory("$dir\営業\2024\A社.xlsx") | Out-Null
        writeListFile "$dir\営業\2024\A社.xlsx\明細.tsv" @("a")
        writeListFile "$dir\営業\2024\A社.xlsx\表紙.tsv" @("b")
        [System.IO.Directory]::CreateDirectory("$dir\営業\空.xlsx") | Out-Null
        [System.IO.Directory]::CreateDirectory("$dir\営業\資料.docx") | Out-Null
        writeListFile "$dir\営業\資料.docx\ページ001.tsv" @("c")
    }

    It "元のファイル1つ分のフォルダごとにTSVの数を数える" {
        $dir = "$TestDrive\index1"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $counts["営業\2024\A社.xlsx"] | Should Be 2
        $counts["営業\資料.docx"] | Should Be 1
    }

    It "TSVの無いフォルダ（内容が空のファイル）は0件として数える" {
        $dir = "$TestDrive\index2"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $counts.ContainsKey("営業\空.xlsx") | Should Be $true
        $counts["営業\空.xlsx"] | Should Be 0
    }

    It "大文字・小文字を区別しない" {
        $dir = "$TestDrive\index3"
        newTestIndex $dir
        (getIndexTsvCounts $dir)["営業\2024\a社.XLSX"] | Should Be 2
    }

    It "インデックスのフォルダの直下のTSV（以前の形式）は数えない" {
        $dir = "$TestDrive\index4"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        writeListFile "$dir\ブック.xlsx_シート.tsv" @("a")
        (getIndexTsvCounts $dir).Count | Should Be 0
    }

    It "フォルダが無ければ空を返す" {
        (getIndexTsvCounts "$TestDrive\none_index").Count | Should Be 0
    }

    It "TSVがそろっていれば「済」のままにする" {
        $dir = "$TestDrive\index5"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "2"
        testIndexComplete $row $row.相対パス $counts | Should Be $true
    }

    It "インデックスのフォルダを直接削除した場合は、そろっていないとする" {
        $dir = "$TestDrive\index6"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\2024\消えたブック.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "3"
        testIndexComplete $row $row.相対パス $counts | Should Be $false
    }

    It "TSVが足りない場合も、そろっていないとする" {
        $dir = "$TestDrive\index7"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "5"
        testIndexComplete $row $row.相対パス $counts | Should Be $false
    }

    It "内容が空のファイル（TSV 0 件）は、フォルダがあればそろっているとする" {
        $dir = "$TestDrive\index8"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\空.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "0"
        testIndexComplete $row $row.相対パス $counts | Should Be $true
    }

    It "0 バイトのTSVがあるフォルダは、壊れているとして作り直す" {
        $dir = "$TestDrive\index9"
        newTestIndex $dir
        # 書き込みの途中で電源が落ちた場合など（空のシート・ページは保存しないため、0 バイトのTSVは異常）
        [System.IO.File]::WriteAllBytes("$dir\営業\2024\A社.xlsx\途中.tsv", (New-Object byte[] 0))
        $counts = getIndexTsvCounts $dir
        $counts["営業\2024\A社.xlsx"] | Should Be ${indexBrokenCount}
        $counts["営業\資料.docx"] | Should Be 1   # ほかのファイルは巻き込まない

        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "2"
        testIndexComplete $row $row.相対パス $counts | Should Be $false
        $other = newStatusRow "営業\資料.docx" "2025/01/10 12:34:56" "100" ${stateDone} "1"
        testIndexComplete $other $other.相対パス $counts | Should Be $true
    }

    It "0 バイトのTSVが先に見つかっても、後のTSVで数え直さない" {
        $dir = "$TestDrive\index10"
        [System.IO.Directory]::CreateDirectory("$dir\営業\B社.xlsx") | Out-Null
        [System.IO.File]::WriteAllBytes("$dir\営業\B社.xlsx\001_途中.tsv", (New-Object byte[] 0))
        writeListFile "$dir\営業\B社.xlsx\002_あと.tsv" @("a")
        (getIndexTsvCounts $dir)["営業\B社.xlsx"] | Should Be ${indexBrokenCount}
    }

    It "TSVの数を記録していない行・数えられなかった場合は確認しない" {
        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} ""
        testIndexComplete $row $row.相対パス (getIndexTsvCounts "$TestDrive\none_index2") | Should Be $true
        $done = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "2"
        testIndexComplete $done $done.相対パス $null | Should Be $true
    }
}

Describe "writeConvertProgress / readConvertProgress / removeConvertProgress" {
    It "段階・件数・内容を往復できる" {
        $path = "$TestDrive\進捗1.txt"
        writeConvertProgress ${convertPhaseRun} 12 34 5 "営業\見積.xlsx" $path
        $progress = readConvertProgress $path
        $progress.Phase | Should Be ${convertPhaseRun}
        $progress.Processed | Should Be 12
        $progress.Remaining | Should Be 34
        $progress.Failed | Should Be 5
        $progress.Detail | Should Be "営業\見積.xlsx"
    }

    It "タブ・改行はスペースにする（1行に保つ）" {
        $path = "$TestDrive\進捗2.txt"
        writeConvertProgress ${convertPhaseScan} 0 0 0 "あ`tい`r`nう" $path
        (readConvertProgress $path).Detail | Should Be "あ い う"
    }

    It "ファイルが無い・壊れていれば null" {
        readConvertProgress "$TestDrive\進捗なし.txt" | Should BeNullOrEmpty
        $path = "$TestDrive\進捗3.txt"
        writeListFile $path @("変換`tあ`tい`tう`tえ")   # 件数が数値でない
        readConvertProgress $path | Should BeNullOrEmpty
        writeListFile $path @("変換`t1`t2")              # 列が足りない（書き込みの途中）
        readConvertProgress $path | Should BeNullOrEmpty
    }

    It "画面が読んでいる間も書ける（共有して開く）" {
        $path = "$TestDrive\進捗4.txt"
        writeConvertProgress ${convertPhaseRun} 1 2 0 "はじめ" $path
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        try {
            { writeConvertProgress ${convertPhaseRun} 2 1 0 "つぎ" $path } | Should Not Throw
        } finally {
            $stream.Dispose()
        }
        (readConvertProgress $path).Detail | Should Be "つぎ"
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\進捗5.txt"
        writeConvertProgress ${convertPhaseFinish} 0 0 0 "" $path
        removeConvertProgress $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeConvertProgress $path } | Should Not Throw
    }
}

Describe "writeConvertPlan / readConvertPlan / removeConvertPlan" {
    It "インデックスごとの件数を往復できる（件数は数値で返る）" {
        $path = "$TestDrive\予定1.tsv"
        $rows = @(
            (newConvertPlanRow "営業" "C:\data\営業" ${planKindConvert} 1234 12 5 7 0 0 3),
            (newConvertPlanRow "技術" "\\server\share\技術" ${planKindConvert} 20 0 0 0 0 0 0))
        writeConvertPlan $rows $path
        $plan = readConvertPlan $path
        $plan.Count | Should Be 2
        $plan[0].インデックス名 | Should Be "営業"
        $plan[0].元のフォルダ | Should Be "C:\data\営業"
        $plan[0].区分 | Should Be ${planKindConvert}
        ($plan[0].ファイル数 + 1) | Should Be 1235   # 文字列ではなく数値で返る
        $plan[0].変換対象 | Should Be 12
        $plan[0].新規 | Should Be 5
        $plan[0].更新あり | Should Be 7
        $plan[0].前回失敗 | Should Be 3
        $plan[1].変換対象 | Should Be 0
    }

    It "チェックなし・フォルダなしの区分も往復できる（件数は 0）" {
        $path = "$TestDrive\予定2.tsv"
        writeConvertPlan @(
            (newConvertPlanRow "外した" "D:\過去" ${planKindUnchecked}),
            (newConvertPlanRow "無い" "E:\USB" ${planKindMissing})) $path
        $plan = readConvertPlan $path
        $plan[0].区分 | Should Be ${planKindUnchecked}
        $plan[0].ファイル数 | Should Be 0
        $plan[1].区分 | Should Be ${planKindMissing}
    }

    It "インデックスが1件も無くても読める（空の配列）" {
        $path = "$TestDrive\予定3.tsv"
        writeConvertPlan @() $path
        (readConvertPlan $path).Count | Should Be 0
    }

    It "ファイルが無い・列が合わなければ null（画面は次の機会に読み直す）" {
        readConvertPlan "$TestDrive\予定なし.tsv" | Should BeNullOrEmpty
        $path = "$TestDrive\予定4.tsv"
        writeListFile $path @("べつの見出し")
        readConvertPlan $path | Should BeNullOrEmpty
    }

    It "画面が読んでいる間も書ける（共有して開く）" {
        $path = "$TestDrive\予定5.tsv"
        writeConvertPlan @((newConvertPlanRow "営業" "C:\data" ${planKindConvert} 1 1 1 0 0 0 0)) $path
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        try {
            { writeConvertPlan @((newConvertPlanRow "営業" "C:\data" ${planKindConvert} 2 2 2 0 0 0 0)) $path } | Should Not Throw
        } finally {
            $stream.Dispose()
        }
        (readConvertPlan $path)[0].変換対象 | Should Be 2
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\予定6.tsv"
        writeConvertPlan @() $path
        removeConvertPlan $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeConvertPlan $path } | Should Not Throw
    }
}

Describe "writeConvertStartRequest / readConvertStartRequest / removeConvertStartRequest" {
    It "前回失敗したファイルも再変換するかを伝えられる" {
        $path = "$TestDrive\開始要求1"
        writeConvertStartRequest $true $path
        (readConvertStartRequest $path).RetryFailed | Should Be $true
        writeConvertStartRequest $false $path
        (readConvertStartRequest $path).RetryFailed | Should Be $false
    }

    It "まだ返事が無ければ null（変換側は待ち続ける）" {
        readConvertStartRequest "$TestDrive\開始要求なし" | Should BeNullOrEmpty
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\開始要求2"
        writeConvertStartRequest $false $path
        removeConvertStartRequest $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeConvertStartRequest $path } | Should Not Throw
    }
}

Describe "removeDirectoryRetry" {
    It "フォルダを中身ごと削除する" {
        $dir = "$TestDrive\消すフォルダ"
        [System.IO.Directory]::CreateDirectory("$dir\中") | Out-Null
        writeListFile "$dir\中\a.tsv" @("a")
        removeDirectoryRetry $dir
        Test-Path -LiteralPath $dir | Should Be $false
    }

    It "フォルダが無ければ何もしない" {
        { removeDirectoryRetry "$TestDrive\無いフォルダ" } | Should Not Throw
    }
}

Describe "publishIndexFiles" {
    It "作業フォルダのTSVを、元のファイルのフォルダに入れる" {
        $from = "$TestDrive\pub1\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        writeListFile "$from\明細.tsv" @("a")
        writeListFile "$from\表紙.tsv" @("b")
        $bookDir = "$TestDrive\pub1\index\営業\A社.xlsx"

        publishIndexFiles $from $bookDir "$TestDrive\pub1\出力\A社.xlsx"

        @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv").Count | Should Be 2
        @(Get-ChildItem -LiteralPath $from -Filter "*.tsv").Count | Should Be 0
        Test-Path -LiteralPath "$TestDrive\pub1\出力\A社.xlsx" | Should Be $false
    }

    It "以前の変換結果は残さず入れ替える（シートの削除・名前変更に追従する）" {
        $from = "$TestDrive\pub2\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        $bookDir = "$TestDrive\pub2\index\営業\A社.xlsx"
        [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
        writeListFile "$bookDir\前のシート.tsv" @("old")
        writeListFile "$from\新しいシート.tsv" @("new")

        publishIndexFiles $from $bookDir "$TestDrive\pub2\出力\A社.xlsx"

        Test-Path -LiteralPath "$bookDir\前のシート.tsv" | Should Be $false
        Test-Path -LiteralPath "$bookDir\新しいシート.tsv" | Should Be $true
    }

    It "TSVが1件も無ければ、空のフォルダにする（内容が空のファイル）" {
        $from = "$TestDrive\pub3\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        $bookDir = "$TestDrive\pub3\index\営業\空.xlsx"

        publishIndexFiles $from $bookDir "$TestDrive\pub3\出力\空.xlsx"

        Test-Path -LiteralPath $bookDir -PathType Container | Should Be $true
        @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv").Count | Should Be 0
    }

    It "前回の出力用フォルダが残っていても入れ替えられる" {
        $from = "$TestDrive\pub4\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        writeListFile "$from\明細.tsv" @("a")
        $staging = "$TestDrive\pub4\出力\A社.xlsx"
        [System.IO.Directory]::CreateDirectory($staging) | Out-Null
        writeListFile "$staging\前回の残り.tsv" @("old")
        $bookDir = "$TestDrive\pub4\index\営業\A社.xlsx"

        publishIndexFiles $from $bookDir $staging

        @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv" | ForEach-Object { $_.Name }) | Should Be "明細.tsv"
    }
}

Describe "newAppMutex" {
    It "同じ処理・同じフォルダでは2つ目を取得できない" {
        $first = newAppMutex "test" "$TestDrive\tool"
        try {
            $first.Acquired | Should Be $true
            $second = newAppMutex "test" "$TestDrive\tool"
            try {
                $second.Acquired | Should Be $false
            } finally {
                $second.Mutex.Dispose()
            }
        } finally {
            $first.Mutex.ReleaseMutex()
            $first.Mutex.Dispose()
        }
    }

    It "処理の種類・フォルダが違えば同時に取得できる" {
        $gui = newAppMutex "gui" "$TestDrive\tool"
        $convert = newAppMutex "convert" "$TestDrive\tool"
        $other = newAppMutex "convert" "$TestDrive\tool2"
        try {
            $gui.Acquired | Should Be $true
            $convert.Acquired | Should Be $true
            $other.Acquired | Should Be $true
        } finally {
            foreach ($m in @($gui, $convert, $other)) {
                $m.Mutex.ReleaseMutex()
                $m.Mutex.Dispose()
            }
        }
    }
}

Describe "パス定義" {
    It "リポジトリ直下を基準にする" {
        $rootDir | Should Be (Resolve-Path "$here\..").Path
        $indexDir | Should Be "$rootDir\work\index"
        $publishDir | Should Be "$rootDir\work\変換出力\$PID"
        $resultFile | Should Be "$rootDir\work\検索結果.txt"
        $settingsFile | Should Be "$rootDir\setting.config"
    }
}

Describe "実行時コンパイル（csc.exe）を使わない" {
    # 画面・共通・変換の各スクリプトが Add-Type -TypeDefinition（実行時コンパイル）を使わないこと。
    # 画面で使う型は PowerShell class に移した（csc.exe の親子関係・一時 DLL を出さないため）
    It "scripts に Add-Type -TypeDefinition が無い" {
        foreach ($file in (Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.ps1")) {
            $source = Get-Content $file.FullName -Raw -Encoding UTF8
            ($source -match "Add-Type\s+-TypeDefinition") | Should Be $false
        }
    }
}

Describe "スクリプトの構文" {
    Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.ps1" | ForEach-Object {
        $script = $_

        It "$($script.Name) に構文エラーが無い" {
            $errors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$null, [ref]$errors) | Out-Null
            $errors.Count | Should Be 0
        }
    }
}

Describe "画面定義（XAML）" {
    $xamlNs = "http://schemas.microsoft.com/winfx/2006/xaml"

    Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.xaml" | ForEach-Object {
        $file = $_

        It "$($file.Name) が XML として読める" {
            { [xml](Get-Content $file.FullName -Raw -Encoding UTF8) } | Should Not Throw
        }
    }

    It "フォルダ選択の画面に、config_gui.ps1 が使う x:Name がすべてある" {
        [xml]$xaml = Get-Content "$here\..\scripts\config_gui_folder_select.xaml" -Raw -Encoding UTF8
        $names = @($xaml.SelectNodes("//*") | ForEach-Object { $_.GetAttribute("Name", $xamlNs) } | Where-Object { $_ -ne "" })
        foreach ($name in @(
                "DescriptionText", "BackButton", "ForwardButton", "UpButton", "AddressBox",
                "FolderTree", "EntryList", "EntryPlaceholder", "StatusText", "FolderBox", "OkButton", "ErrorText")) {
            $names -contains $name | Should Be $true
        }
    }
}
