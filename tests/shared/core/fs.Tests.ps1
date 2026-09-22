# ファイルの読み書き（shared\core\fs.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "toSafeFileName" -Tag Io {
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

Describe "toLongPath / fromLongPath" -Tag Io {
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

Describe "copyFileShared" -Tag Io {
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

Describe "長いパス（260文字超）" -Tag Io {
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

Describe "readListFile / writeListFile" -Tag Io {
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

Describe "formatFileTime" -Tag Io {
    It "秒までの日時にする" {
        formatFileTime (New-Object DateTime 2025, 1, 2, 3, 4, 5, 678) | Should Be "2025/01/02 03:04:05"
    }
}

Describe "removeDirectoryRetry" -Tag Io {
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

Describe "newAppMutex" -Tag Io {
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
        $indexer = newAppMutex "indexer" "$TestDrive\tool"
        $other = newAppMutex "indexer" "$TestDrive\tool2"
        try {
            $gui.Acquired | Should Be $true
            $indexer.Acquired | Should Be $true
            $other.Acquired | Should Be $true
        } finally {
            foreach ($m in @($gui, $indexer, $other)) {
                $m.Mutex.ReleaseMutex()
                $m.Mutex.Dispose()
            }
        }
    }
}
