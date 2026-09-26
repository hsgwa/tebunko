# ファイルの読み書き（shared\core\fs.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "toSafeFileName" -Tag Io {
    It "<name>" -TestCases @(
        @{ name = "ファイル名に使えない文字を全角に変換する"; text = 'a<b>c\d*e:f?g|h'; expected = 'a＜b＞c￥d＊e：f？g｜h' }
        @{ name = "スラッシュを全角に変換する"; text = 'a/b'; expected = 'a／b' }
        @{ name = "ダブルクォートを全角に変換する"; text = 'a"b'; expected = 'a”b' }
        @{ name = "使える文字はそのまま返す"; text = 'シート1 (2)'; expected = 'シート1 (2)' }
    ) {
        param($name, $text, $expected)
        toSafeFileName $text | Should Be $expected
    }
}

Describe "toLongPath / fromLongPath" -Tag Io {
    It "<name>。外すと元に戻り、付いていればそのまま、付いていなければ外してもそのまま" -TestCases @(
        @{ name = "ドライブのパスに \\?\ を付ける"; short = "C:\data\a.xlsx"; long = "\\?\C:\data\a.xlsx" }
        @{ name = "ネットワークのパスは \\?\UNC\ にする"; short = "\\server\share\a.xlsx"; long = "\\?\UNC\server\share\a.xlsx" }
    ) {
        param($name, $short, $long)
        toLongPath $short | Should Be $long
        fromLongPath $long | Should Be $short
        toLongPath $long | Should Be $long
        fromLongPath $short | Should Be $short
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

    It "長いパスのフォルダでも、集約ファイルを作り・列挙・検索でき、相対パスは \\?\ の無い形になる" {
        $root = "$TestDrive\indexLong"
        $dir = "$root\$deepRel"
        [void][System.IO.Directory]::CreateDirectory((toLongPath "$dir\book.xlsx"))
        [System.IO.File]::WriteAllText((toLongPath "$dir\book.xlsx\Sheet1.tsv"), "hello`r`n", $utf8Bom)
        [void][System.IO.Directory]::CreateDirectory("$root\short.xlsx")
        [System.IO.File]::WriteAllText("$root\short.xlsx\Sheet1.tsv", "hello`r`n", $utf8Bom)
        try {
            foreach ($folder in (findIndexFoldersWithBooks $root)) {
                [void](updateIndexFolderPack $folder)
            }
            $index = getIndexPackFiles @($root)
            $index.Folders[0].Count | Should Be 2
            $rel = @($index.Packs | ForEach-Object { $_.RelPath } | Sort-Object)
            $rel | Should Be @("$deepRel\content.xlsx.001.tsv", "content.xlsx.001.tsv")

            $hits = @((searchPackIndex "hello" $index.Packs).Hits)
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

    It "ほかから共有せずに開かれていて読めなければ、空の一覧ではなく例外にする（`$ErrorActionPreference によらない）" {
        $path = "$TestDrive\読めない一覧\list.txt"
        writeListFile $path @("a")
        $stream = [System.IO.File]::Open($path, "Open", "ReadWrite", "None")
        try {
            & {
                $ErrorActionPreference = "Continue"
                { readListFile $path } | Should Throw
            }
        } finally {
            $stream.Dispose()
        }
    }

    It "行を渡さなければ空のファイルを作る" {
        $path = "$TestDrive\空の一覧\list.txt"
        writeListFile $path $null
        Test-Path -LiteralPath $path | Should Be $true
        @(readListFile $path).Count | Should Be 0
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

    It "中のファイルがほかから開かれていて消せなければ、決めた回数だけ試してから例外にする" {
        Mock Start-Sleep {}
        $dir = "$TestDrive\使用中のフォルダ"
        writeListFile "$dir\a.tsv" @("a")
        $stream = [System.IO.File]::Open("$dir\a.tsv", "Open", "Read", "None")
        try {
            { removeDirectoryRetry $dir 3 1 } | Should Throw
        } finally {
            $stream.Dispose()
        }
        Assert-MockCalled Start-Sleep -Times 2 -Exactly -Scope It
        Test-Path -LiteralPath "$dir\a.tsv" | Should Be $true
    }
}

Describe "writeTextLinesAtomic" -Tag Io {
    It "新しいファイルを作り、一時ファイルを残さない" {
        $path = "$TestDrive\atomic\新規.txt"
        writeTextLinesAtomic $path @("一", "二")
        @(readListFile $path) -join "," | Should Be "一,二"
        Test-Path -LiteralPath "${path}.tmp" | Should Be $false
    }

    It "既存のファイルを置き換える" {
        $path = "$TestDrive\atomic\置換.txt"
        writeTextLinesAtomic $path @("前")
        writeTextLinesAtomic $path @("後")
        @(readListFile $path) -join "," | Should Be "後"
        Test-Path -LiteralPath "${path}.tmp" | Should Be $false
    }

    It "置き換えられなければ 5 回まで試してから例外にし、元のファイルは壊さない" {
        Mock Start-Sleep {}
        $path = "$TestDrive\atomic\使用中.txt"
        writeTextLinesAtomic $path @("元の内容")
        $stream = [System.IO.File]::Open($path, "Open", "Read", "None")
        try {
            { writeTextLinesAtomic $path @("新しい内容") } | Should Throw
        } finally {
            $stream.Dispose()
        }
        Assert-MockCalled Start-Sleep -Times 4 -Exactly -Scope It
        @(readListFile $path) -join "," | Should Be "元の内容"
    }
}

Describe "getFolderKey" -Tag Unit {
    It "SHA-256 の 16 進 64 文字を返す" {
        # "c:\tool" の SHA-256（小文字にしてから UTF-8 で計算する）
        getFolderKey "C:\Tool" | Should Be "DA4B936E296325CC586C02ADB4518ABB2C4021761A7AFCE2D670263F384C89F5"
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
