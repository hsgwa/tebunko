# ファイルの読み書き（shared\core\fs.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "toSafeFileName" -Tag Io {
    It "<name>" -TestCases @(
        @{ name = "ファイル名に使えない文字を全角に変換する"; text = 'a<b>c\d*e:f?g|h'; expected = 'a＜b＞c￥d＊e：f？g｜h' }
        @{ name = "スラッシュを全角に変換する"; text = 'a/b'; expected = 'a／b' }
        @{ name = "ダブルクォートを全角に変換する"; text = 'a"b'; expected = 'a”b' }
        @{ name = "使える文字はそのまま返す"; text = 'シート1 (2)'; expected = 'シート1 (2)' }
    ) {
        param($name, $text, $expected)
        toSafeFileName $text | Should -Be $expected
    }
}

Describe "toLongPath / fromLongPath" -Tag Io {
    It "<name>。外すと元に戻り、付いていればそのまま、付いていなければ外してもそのまま" -TestCases @(
        @{ name = "ドライブのパスに \\?\ を付ける"; short = "C:\data\a.xlsx"; long = "\\?\C:\data\a.xlsx" }
        @{ name = "ネットワークのパスは \\?\UNC\ にする"; short = "\\server\share\a.xlsx"; long = "\\?\UNC\server\share\a.xlsx" }
    ) {
        param($name, $short, $long)
        toLongPath $short | Should -Be $long
        fromLongPath $long | Should -Be $short
        toLongPath $long | Should -Be $long
        fromLongPath $short | Should -Be $short
    }

    It "/ は \ にする" {
        toLongPath "C:/data/a.xlsx" | Should -Be "\\?\C:\data\a.xlsx"
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
        [System.IO.File]::ReadAllText("$TestDrive\コピー.xlsx") | Should -Be "abc"
    }

    It "読み取り専用のファイルも、通常の属性のコピーを作り、既にあれば上書きする" {
        $source = "$TestDrive\読み取り専用.docx"
        $dest = "$TestDrive\上書き.docx"
        [System.IO.File]::WriteAllText($source, "new", [System.Text.Encoding]::ASCII)
        [System.IO.File]::SetAttributes($source, [System.IO.FileAttributes]::ReadOnly)
        [System.IO.File]::WriteAllText($dest, "old-content", [System.Text.Encoding]::ASCII)
        try {
            copyFileShared $source $dest
            [System.IO.File]::ReadAllText($dest) | Should -Be "new"
            ([System.IO.File]::GetAttributes($dest) -band [System.IO.FileAttributes]::ReadOnly) | Should -Be 0
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
            [System.IO.File]::ReadAllText("$TestDrive\long_copy.xlsx") | Should -Be "long"
        } finally {
            Remove-Item -LiteralPath (toLongPath "$TestDrive\copyLong") -Recurse -Force
        }
    }
}

Describe "長いパス（260文字超）" -Tag Io {
    # フォルダのパスが約248文字、ファイルのパスが260文字を超えると、\\?\ を付けないと扱えない
    BeforeAll {
        $deepRel = ("a" * 100) + "\" + ("b" * 100)
    }

    It "prettyTsv で長いパスに保存できる" {
        $dir = "$TestDrive\prettyLong\$deepRel"
        [void][System.IO.Directory]::CreateDirectory((toLongPath $dir))
        $in = "$TestDrive\long_sheet.tmp"
        [System.IO.File]::WriteAllText($in, "x`r`n", [System.Text.Encoding]::Unicode)
        try {
            $out = "$dir\book.xlsx_Sheet1.tsv"
            $out.Length | Should -BeGreaterThan 260
            prettyTsv $in $out | Should -Be $true
            [System.IO.File]::ReadAllText((toLongPath $out)) | Should -Be "x`r`n"
        } finally {
            Remove-Item -LiteralPath (toLongPath "$TestDrive\prettyLong") -Recurse -Force
        }
    }

    It "長いパスのフォルダでも、本文インデックスのファイルを作り・列挙・検索でき、相対パスは \\?\ の無い形になる" {
        $root = "$TestDrive\indexLong"
        $dir = "$root\$deepRel"
        [void][System.IO.Directory]::CreateDirectory((toLongPath "$dir\book.xlsx"))
        [System.IO.File]::WriteAllText((toLongPath "$dir\book.xlsx\Sheet1.tsv"), "hello`r`n", $utf8Bom)
        [void][System.IO.Directory]::CreateDirectory("$root\short.xlsx")
        [System.IO.File]::WriteAllText("$root\short.xlsx\Sheet1.tsv", "hello`r`n", $utf8Bom)
        try {
            foreach ($folder in (findIndexFoldersWithBooks $root)) {
                [void](updateFolderContentIndex $folder)
            }
            $index = getContentIndexFiles @($root)
            $index.Folders[0].Count | Should -Be 2
            $rel = @($index.ContentIndexFiles | ForEach-Object { $_.RelPath } | Sort-Object)
            $rel | Should -Be @("$deepRel\content_index.xlsx.001.tsv", "content_index.xlsx.001.tsv")

            $hits = @((searchContentIndex "hello" $index.ContentIndexFiles).Hits)
            $hits.Count | Should -Be 2
            @($hits | Where-Object { $_.RelDir -eq $deepRel }).Count | Should -Be 1

            (getIndexSummary @($root)).Count | Should -Be 2
            testIndexExists @($root) | Should -Be $true
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
        $lines.Count | Should -Be 2
        $lines[0] | Should -Be "a\[確定]見積.xlsx"
        $lines[1] | Should -Be " b.xls"
    }

    It "ファイルが無ければ空配列を返す" {
        @(readListFile "$TestDrive\none_list.txt").Count | Should -Be 0
    }

    It "ほかから共有せずに開かれていて読めなければ、空の一覧ではなく例外にする（`$ErrorActionPreference によらない）" {
        $path = "$TestDrive\読めない一覧\list.txt"
        writeListFile $path @("a")
        $stream = [System.IO.File]::Open($path, "Open", "ReadWrite", "None")
        try {
            & {
                $ErrorActionPreference = "Continue"
                { readListFile $path } | Should -Throw
            }
        } finally {
            $stream.Dispose()
        }
    }

    It "行を渡さなければ空のファイルを作る" {
        $path = "$TestDrive\空の一覧\list.txt"
        writeListFile $path $null
        Test-Path -LiteralPath $path | Should -Be $true
        @(readListFile $path).Count | Should -Be 0
    }
}

Describe "formatFileTime" -Tag Io {
    It "秒までの日時にする" {
        formatFileTime (New-Object DateTime 2025, 1, 2, 3, 4, 5, 678) | Should -Be "2025/01/02 03:04:05"
    }
}

Describe "getFolderNameProblemText" -Tag Unit {
    It "「<name>」が使えるか: <usable>" -TestCases @(
        @{ name = "営業部 2025"; usable = $true }
        @{ name = "CONSOLE"; usable = $true }
        @{ name = ("あ" * 255); usable = $true }
        @{ name = ""; usable = $false }
        @{ name = " 営業"; usable = $false }
        @{ name = "営業　"; usable = $false }
        @{ name = ("あ" * 256); usable = $false }
        @{ name = ".."; usable = $false }
        @{ name = "営業."; usable = $false }
        @{ name = "a\b"; usable = $false }
        @{ name = "a/b"; usable = $false }
        @{ name = "a:b"; usable = $false }
        @{ name = "con"; usable = $false }
        @{ name = "NUL.txt"; usable = $false }
        @{ name = "LPT1"; usable = $false }
    ) {
        $problem = getFolderNameProblemText $name
        ($problem -eq "") | Should -Be $usable
    }

    It "予約名を許すとき、「<name>」が使えるか: <usable>（予約名以外は同じ決まり）" -TestCases @(
        @{ name = "aux.backup"; usable = $true }
        @{ name = "CON.txt"; usable = $true }
        @{ name = "con"; usable = $true }
        @{ name = "LPT1"; usable = $true }
        @{ name = "営業部 2025"; usable = $true }
        @{ name = ""; usable = $false }
        @{ name = ".."; usable = $false }
        @{ name = " con"; usable = $false }
        @{ name = "con."; usable = $false }
        @{ name = "a\b"; usable = $false }
        @{ name = "a:b"; usable = $false }
        @{ name = ("あ" * 256); usable = $false }
    ) {
        $problem = getFolderNameProblemText $name $true
        ($problem -eq "") | Should -Be $usable
    }
}

Describe "testPlainFolderName" -Tag Unit {
    It "「<name>」が使えるか: <usable>" -TestCases @(
        @{ name = "営業部"; usable = $true }
        @{ name = "aux.backup"; usable = $true }
        @{ name = "CON"; usable = $true }
        @{ name = "営業."; usable = $true }
        @{ name = "営業 "; usable = $true }
        @{ name = ("あ" * 255); usable = $true }
        @{ name = ""; usable = $false }
        @{ name = "."; usable = $false }
        @{ name = ".."; usable = $false }
        @{ name = ("あ" * 256); usable = $false }
        @{ name = "a\b"; usable = $false }
        @{ name = "a/b"; usable = $false }
        @{ name = "a:b"; usable = $false }
    ) {
        (testPlainFolderName $name) | Should -Be $usable
    }
}

Describe "testPathUnder" -Tag Unit {
    It "<path> は <root> の下: <expected>" -TestCases @(
        @{ path = "C:\置き場\営業"; root = "C:\置き場"; expected = $true }
        @{ path = "C:\置き場\営業\中\a.tsv"; root = "C:\置き場"; expected = $true }
        @{ path = "C:\置き場\営業"; root = "C:\置き場\"; expected = $true }
        @{ path = "c:\置き場\営業"; root = "C:\置き場"; expected = $true }
        @{ path = "C:\置き場"; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\"; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\..\外\営業"; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\営業\..\..\外"; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\営業\..\経理"; root = "C:\置き場"; expected = $true }
        @{ path = "C:\置き場2\営業"; root = "C:\置き場"; expected = $false }
        @{ path = "C:\外\営業"; root = "C:\置き場"; expected = $false }
        @{ path = "D:\置き場\営業"; root = "C:\置き場"; expected = $false }
        @{ path = "\\?\C:\置き場\営業"; root = "C:\置き場"; expected = $true }
        @{ path = "\\?\C:\置き場\..\外"; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\営業"; root = "\\?\C:\置き場"; expected = $true }
        @{ path = ""; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\営業"; root = ""; expected = $false }
        # ネットワークのパス（架空の名前。ネットワークには触れない）
        @{ path = "\\server\share\置き場\営業"; root = "\\server\share\置き場"; expected = $true }
        @{ path = "\\server\share\置き場\..\外"; root = "\\server\share\置き場"; expected = $false }
        @{ path = "\\server\share\置き場"; root = "\\server\share\置き場"; expected = $false }
        @{ path = "\\server\share2\置き場\営業"; root = "\\server\share\置き場"; expected = $false }
        @{ path = "\\?\UNC\server\share\置き場\営業"; root = "\\server\share\置き場"; expected = $true }
        @{ path = "\\?\UNC\server\share\置き場\..\外"; root = "\\server\share\置き場"; expected = $false }
        @{ path = "\\server\share\置き場\営業"; root = "\\?\UNC\server\share\置き場"; expected = $true }
        # 区切りが混ざったパス
        @{ path = "C:/置き場/営業"; root = "C:\置き場"; expected = $true }
        @{ path = "C:\置き場/営業\中"; root = "C:/置き場"; expected = $true }
        @{ path = "C:\置き場/..\外"; root = "C:\置き場"; expected = $false }
        @{ path = "C:/置き場/../外/営業"; root = "C:\置き場"; expected = $false }
        # 最後が . や空白だけの名前（Windows は . や空白を取って解くため、置き場そのもの・別のフォルダになりうる）
        @{ path = "C:\置き場\."; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\ "; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\営業."; root = "C:\置き場"; expected = $true }
        @{ path = "C:\置き場\営業 "; root = "C:\置き場"; expected = $true }
        @{ path = "C:\置き場\.\.."; root = "C:\置き場"; expected = $false }
        # : を含む名前は解けない（代替データストリームの書き方）ため、置き場の下とは認めない
        @{ path = "C:\置き場\営業:ads"; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\営業\a:b"; root = "C:\置き場"; expected = $false }
        @{ path = "C:\置き場\営業\..\..\外:x"; root = "C:\置き場"; expected = $false }
        # 置き場がドライブの直下
        @{ path = "C:\営業"; root = "C:\"; expected = $true }
        @{ path = "C:\営業\中"; root = "C:\"; expected = $true }
        @{ path = "C:\"; root = "C:\"; expected = $false }
        @{ path = "D:\営業"; root = "C:\"; expected = $false }
        @{ path = "C:\..\..\営業"; root = "C:\"; expected = $true }
    ) {
        testPathUnder $path $root | Should -Be $expected
    }

    It "260 文字を超えるパスでも判定できる" {
        $long = "C:\置き場\" + ("あ" * 100) + "\" + ("い" * 100) + "\" + ("う" * 100) + "\a.tsv"
        testPathUnder $long "C:\置き場" | Should -Be $true
        testPathUnder ($long + "\..\..\..\..\..\外") "C:\置き場" | Should -Be $false
    }
}

Describe "removeDirectoryRetry" -Tag Io {
    It "フォルダを中身ごと削除する" {
        $dir = "$TestDrive\消すフォルダ"
        [System.IO.Directory]::CreateDirectory("$dir\中") | Out-Null
        writeListFile "$dir\中\a.tsv" @("a")
        removeDirectoryRetry $dir
        Test-Path -LiteralPath $dir | Should -Be $false
    }

    It "フォルダが無ければ何もしない" {
        { removeDirectoryRetry "$TestDrive\無いフォルダ" } | Should -Not -Throw
    }

    It "中のファイルがほかから開かれていて消せなければ、決めた回数だけ試してから例外にする" {
        Mock Start-Sleep {}
        $dir = "$TestDrive\使用中のフォルダ"
        writeListFile "$dir\a.tsv" @("a")
        $stream = [System.IO.File]::Open("$dir\a.tsv", "Open", "Read", "None")
        try {
            { removeDirectoryRetry $dir 3 1 } | Should -Throw
        } finally {
            $stream.Dispose()
        }
        Should -Invoke Start-Sleep -Times 2 -Exactly -Scope It
        Test-Path -LiteralPath "$dir\a.tsv" | Should -Be $true
    }

    It "中にあるジャンクションの先は消さず、ジャンクションだけを外す" {
        $dir = "$TestDrive\消すフォルダ_リンク"
        $outside = "$TestDrive\外のフォルダ_リンク"
        writeListFile "$outside\外.tsv" @("外")
        writeListFile "$dir\中\a.tsv" @("a")
        if (!(newTestJunction "$dir\中\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        removeDirectoryRetry $dir
        Test-Path -LiteralPath $dir | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
    }

    It "消すフォルダそのものがジャンクションなら、ジャンクションだけを外し、先は消さない" {
        $outside = "$TestDrive\外のフォルダ_根"
        writeListFile "$outside\外.tsv" @("外")
        if (!(newTestJunction "$TestDrive\根のリンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        removeDirectoryRetry "$TestDrive\根のリンク"
        Test-Path -LiteralPath "$TestDrive\根のリンク" | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
    }
}

Describe "removeTreeNoFollow" -Tag Io {
    It "フォルダを中身ごと消す（読み取り専用のファイル・フォルダも）" {
        $dir = "$TestDrive\tree_plain"
        writeListFile "$dir\中\深い\a.tsv" @("a")
        writeListFile "$dir\b.tsv" @("b")
        [System.IO.File]::SetAttributes("$dir\b.tsv", [System.IO.FileAttributes]::ReadOnly)
        [System.IO.File]::SetAttributes("$dir\中\深い", [System.IO.FileAttributes]::ReadOnly -bor [System.IO.FileAttributes]::Directory)
        removeTreeNoFollow $dir
        Test-Path -LiteralPath $dir | Should -Be $false
    }

    It "260 文字を超えるパスも消す" {
        $dir = "$TestDrive\tree_long"
        $deep = "$dir\" + ("深いフォルダ" * 20) + "\" + ("もっと深いフォルダ" * 15)
        [System.IO.Directory]::CreateDirectory((toLongPath $deep)) | Out-Null
        [System.IO.File]::WriteAllText((toLongPath "$deep\a.tsv"), "a")
        $deep.Length | Should -BeGreaterThan 260
        removeTreeNoFollow $dir
        [System.IO.Directory]::Exists((toLongPath $dir)) | Should -Be $false
    }

    It "フォルダが無ければ何もしない" {
        { removeTreeNoFollow "$TestDrive\tree_none" } | Should -Not -Throw
    }

    It "中の何階層も下にあるジャンクションの先を、ファイルごと消さない" {
        $dir = "$TestDrive\tree_junction"
        $outside = "$TestDrive\tree_junction_outside"
        writeListFile "$outside\外.tsv" @("外")
        writeListFile "$outside\下\外の下.tsv" @("外の下")
        writeListFile "$dir\a\b\c.tsv" @("c")
        if (!(newTestJunction "$dir\a\b\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        removeTreeNoFollow $dir
        Test-Path -LiteralPath $dir | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
        readListFile "$outside\下\外の下.tsv" | Should -Be @("外の下")
    }

    It "読み取り専用の先のファイルも、ジャンクションを通しては消さない" {
        $dir = "$TestDrive\tree_junction_ro"
        $outside = "$TestDrive\tree_junction_ro_outside"
        writeListFile "$outside\外.tsv" @("外")
        [System.IO.File]::SetAttributes("$outside\外.tsv", [System.IO.FileAttributes]::ReadOnly)
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        if (!(newTestJunction "$dir\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        removeTreeNoFollow $dir
        Test-Path -LiteralPath $dir | Should -Be $false
        [System.IO.File]::Exists("$outside\外.tsv") | Should -Be $true
    }

    It "消すフォルダそのものがジャンクションなら、ジャンクションだけを外し、先は消さない" {
        $outside = "$TestDrive\tree_root_outside"
        writeListFile "$outside\外.tsv" @("外")
        if (!(newTestJunction "$TestDrive\tree_root_link" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        removeTreeNoFollow "$TestDrive\tree_root_link"
        Test-Path -LiteralPath "$TestDrive\tree_root_link" | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
    }

    It "読み取り専用のジャンクションも、ジャンクションだけを外し、先は消さない" {
        $dir = "$TestDrive\tree_junction_attr"
        $outside = "$TestDrive\tree_junction_attr_outside"
        writeListFile "$outside\外.tsv" @("外")
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        if (!(newTestJunction "$dir\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        [System.IO.File]::SetAttributes("$dir\リンク", [System.IO.File]::GetAttributes("$dir\リンク") -bor [System.IO.FileAttributes]::ReadOnly)
        # リンク自身に読み取り専用が付いていて、先には付いていない
        ([System.IO.File]::GetAttributes("$dir\リンク") -band [System.IO.FileAttributes]::ReadOnly) | Should -Be ([int][System.IO.FileAttributes]::ReadOnly)
        ([System.IO.File]::GetAttributes($outside) -band [System.IO.FileAttributes]::ReadOnly) | Should -Be 0
        removeTreeNoFollow $dir
        Test-Path -LiteralPath $dir | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
        ([System.IO.File]::GetAttributes($outside) -band [System.IO.FileAttributes]::ReadOnly) | Should -Be 0
    }

    It "消すパスの末尾に \ が付いていても、ジャンクションだけを外し、先は消さない" {
        $outside = "$TestDrive\tree_slash_outside"
        writeListFile "$outside\外.tsv" @("外")
        if (!(newTestJunction "$TestDrive\tree_slash_link" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        removeTreeNoFollow "$TestDrive\tree_slash_link\"
        Test-Path -LiteralPath "$TestDrive\tree_slash_link" | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
    }

    It "親を指すジャンクション（輪）があっても、時間内に終わり、全部消える" {
        $dir = "$TestDrive\tree_loop"
        writeListFile "$dir\a\a.tsv" @("a")
        if (!(newTestJunction "$dir\a\輪" $dir)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        # 終わらない作りになったときに、テスト全体が止まらないよう、別のスレッドで上限の時間を決めて待つ
        $worker = [powershell]::Create()
        try {
            [void]$worker.AddScript({ param ($fsScript, $path) . $fsScript; removeTreeNoFollow $path }).AddArgument("${scriptsDir}\shared\core\fs.ps1").AddArgument($dir)
            $handle = $worker.BeginInvoke()
            $finished = $handle.AsyncWaitHandle.WaitOne(30000)
            if (!$finished) { $worker.Stop() }
            $finished | Should -Be $true
            $worker.HadErrors | Should -Be $false
        } finally {
            $worker.Dispose()
        }
        Test-Path -LiteralPath $dir | Should -Be $false
    }

    It "フォルダのシンボリックリンク・ファイルのシンボリックリンクは、リンクだけを消し、先は残す" {
        $dir = "$TestDrive\tree_symlink"
        $outside = "$TestDrive\tree_symlink_outside"
        writeListFile "$outside\外.tsv" @("外")
        writeListFile "$outside\ファイル.txt" @("ファイル")
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        if (!(newTestSymlink "$dir\フォルダ" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        if (!(newTestSymlink "$dir\ファイル" "$outside\ファイル.txt" -File)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        testPathIsLink "$dir\ファイル" | Should -Be $true
        removeTreeNoFollow $dir
        Test-Path -LiteralPath $dir | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
        readListFile "$outside\ファイル.txt" | Should -Be @("ファイル")
    }
}

Describe "testLinkItem" -Tag Unit {
    It "<name> なら <expected>" -TestCases @(
        @{ name = "普通のファイル（属性だけ）"; attributes = "Archive"; linkType = $null; expected = $false }
        @{ name = "普通のフォルダ"; attributes = "Directory"; linkType = $null; expected = $false }
        @{ name = "ReparsePoint だけで LinkType が空（OneDrive のオンデマンドのファイル）"; attributes = "Archive, ReparsePoint"; linkType = ""; expected = $false }
        @{ name = "ジャンクション"; attributes = "Directory, ReparsePoint"; linkType = "Junction"; expected = $true }
        @{ name = "シンボリックリンク"; attributes = "Directory, ReparsePoint"; linkType = "SymbolicLink"; expected = $true }
        @{ name = "ReparsePoint が付かず LinkType だけがある（属性を信用する）"; attributes = "Directory"; linkType = "Junction"; expected = $false }
        @{ name = "フォルダで ReparsePoint があり LinkType が空（種類が確かめられない）"; attributes = "Directory, ReparsePoint"; linkType = ""; expected = $true }
        @{ name = "フォルダで ReparsePoint があり LinkType が `$null（種類が確かめられない）"; attributes = "Directory, ReparsePoint"; linkType = $null; expected = $true }
        @{ name = "フォルダで ReparsePoint があり LinkType が知らない値"; attributes = "Directory, ReparsePoint"; linkType = "AppExecLink"; expected = $true }
    ) {
        param ($name, $attributes, $linkType, $expected)
        $item = [pscustomobject]@{ Attributes = [System.IO.FileAttributes]$attributes; LinkType = $linkType }
        testLinkItem $item | Should -Be $expected
    }

    It "属性が読めない（空）ときは、リンクとして扱う" {
        testLinkItem ([pscustomobject]@{ LinkType = $null }) | Should -Be $true
        testLinkItem $null | Should -Be $true
    }
}

Describe "getLinkKind" -Tag Unit {
    It "<name> は <expected>" -TestCases @(
        @{ name = "普通のファイル"; attributes = "Archive"; linkType = $null; expected = "Plain" }
        @{ name = "普通のフォルダ"; attributes = "Directory"; linkType = $null; expected = "Plain" }
        @{ name = "ReparsePoint のファイル（LinkType が空）"; attributes = "Archive, ReparsePoint"; linkType = ""; expected = "Plain" }
        @{ name = "ジャンクション"; attributes = "Directory, ReparsePoint"; linkType = "Junction"; expected = "Link" }
        @{ name = "シンボリックリンク"; attributes = "Directory, ReparsePoint"; linkType = "SymbolicLink"; expected = "Link" }
        @{ name = "ReparsePoint のフォルダで LinkType が空"; attributes = "Directory, ReparsePoint"; linkType = ""; expected = "Unknown" }
        @{ name = "ReparsePoint のフォルダで LinkType が `$null"; attributes = "Directory, ReparsePoint"; linkType = $null; expected = "Unknown" }
        @{ name = "ReparsePoint のフォルダで LinkType が知らない値"; attributes = "Directory, ReparsePoint"; linkType = "AppExecLink"; expected = "Unknown" }
    ) {
        param ($name, $attributes, $linkType, $expected)
        $item = [pscustomobject]@{ Attributes = [System.IO.FileAttributes]$attributes; LinkType = $linkType }
        getLinkKind $item | Should -Be $expected
    }

    It "属性が読めない（空）ときは Unknown" {
        getLinkKind ([pscustomobject]@{ LinkType = $null }) | Should -Be "Unknown"
    }
}

Describe "removeTreeNoFollow（種類が確かめられないフォルダ）" -Tag Io {
    It "リンクかどうか確かめられないフォルダがあれば、辿らず・消さずに、何も消さないまま理由を付けて失敗する" {
        writeListFile "$TestDrive\unknown_kind\中\a.tsv" @("a")
        writeListFile "$TestDrive\unknown_kind\中\下\b.tsv" @("b")
        writeListFile "$TestDrive\unknown_kind\あ.tsv" @("あ")
        writeListFile "$TestDrive\unknown_kind\兄弟\c.tsv" @("c")
        Mock getLinkKind { return "Unknown" } -ParameterFilter { $item.Name -eq "中" }

        { removeTreeNoFollow "$TestDrive\unknown_kind" } | Should -Throw "*確かめられない*"

        readListFile "$TestDrive\unknown_kind\中\a.tsv" | Should -Be @("a")
        readListFile "$TestDrive\unknown_kind\中\下\b.tsv" | Should -Be @("b")
        readListFile "$TestDrive\unknown_kind\あ.tsv" | Should -Be @("あ")
        readListFile "$TestDrive\unknown_kind\兄弟\c.tsv" | Should -Be @("c")
    }

    It "根が確かめられないフォルダなら、何も消さずに失敗する" {
        writeListFile "$TestDrive\unknown_root\a.tsv" @("a")
        Mock getLinkKind { return "Unknown" } -ParameterFilter { $item.Name -eq "unknown_root" }

        { removeTreeNoFollow "$TestDrive\unknown_root" } | Should -Throw "*確かめられない*"

        readListFile "$TestDrive\unknown_root\a.tsv" | Should -Be @("a")
    }
}

Describe "removeTreeNoFollow（数えたあとでリンクに替わったフォルダ）" -Tag Io {
    It "ファイルを消す直前の確かめでリンクだと分かったら、そのフォルダのファイルを消さずに失敗する" {
        writeListFile "$TestDrive\swapped\中\a.tsv" @("a")
        writeListFile "$TestDrive\swapped\中\下\b.tsv" @("b")
        writeListFile "$TestDrive\swapped\あ.tsv" @("あ")
        $script:swappedCalls = 0
        # 数える段で 2 回（上のフォルダの列挙・降りる直前）は Plain、消す段の 3 回目から Link
        Mock getLinkKind {
            $script:swappedCalls++
            if ($script:swappedCalls -ge 3) {
                return "Link"
            }
            return "Plain"
        } -ParameterFilter { $item.Name -eq "中" }

        { removeTreeNoFollow "$TestDrive\swapped" } | Should -Throw "*リンクに替わった*"

        readListFile "$TestDrive\swapped\中\a.tsv" | Should -Be @("a")
        readListFile "$TestDrive\swapped\中\下\b.tsv" | Should -Be @("b")
    }
}

Describe "trimTrailingSeparator" -Tag Unit {
    It "末尾の区切りだけを落とす。根・区切りだけ・空のパスはそのまま返す: <path>" -TestCases @(
        @{ path = "C:\a\"; expected = "C:\a" }
        @{ path = "C:\a\\"; expected = "C:\a" }
        @{ path = "C:\a/"; expected = "C:\a" }
        @{ path = "C:\a/b/"; expected = "C:\a/b" }
        @{ path = "C:\a"; expected = "C:\a" }
        @{ path = "C:\"; expected = "C:\" }
        @{ path = "C:"; expected = "C:" }
        @{ path = "\\?\C:\"; expected = "\\?\C:\" }
        @{ path = "\\?\"; expected = "\\?\" }
        @{ path = "\"; expected = "\" }
        @{ path = "/"; expected = "/" }
        @{ path = ""; expected = "" }
        @{ path = "\\?\C:\a\"; expected = "\\?\C:\a" }
    ) {
        trimTrailingSeparator $path | Should -Be $expected
    }
}

Describe "testPathIsLink / testPathHasLink（末尾の区切り）" -Tag Io {
    It "ジャンクションのパスの末尾に \ が付いていても、リンクと判定する" {
        if (!(newTestJunction "$TestDrive\islink_slash" "$TestDrive\islink_slash_target")) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        testPathIsLink "$TestDrive\islink_slash\" | Should -Be $true
    }

    It "ドライブのルートは、末尾の \ を落とさずにリンクでないと判定する" {
        testPathIsLink ([System.IO.Path]::GetPathRoot($TestDrive)) | Should -Be $false
    }
}

Describe "testPathIsLink / testPathHasLink" -Tag Io {
    It "普通のフォルダ・ファイル・無いパス・空のパスはリンクでない" {
        writeListFile "$TestDrive\islink_plain\a.tsv" @("a")
        testPathIsLink "$TestDrive\islink_plain" | Should -Be $false
        testPathIsLink "$TestDrive\islink_plain\a.tsv" | Should -Be $false
        testPathIsLink "$TestDrive\islink_plain\無い" | Should -Be $false
        testPathIsLink "" | Should -Be $false
    }

    It "ジャンクションはリンク。ジャンクションの下の普通のフォルダはリンクでない" {
        if (!(newTestJunction "$TestDrive\islink_junction" "$TestDrive\islink_target\下")) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        testPathIsLink "$TestDrive\islink_junction" | Should -Be $true
        testPathIsLink "$TestDrive\islink_target\下" | Should -Be $false
    }

    It "testPathHasLink: base の下の途中（path を含む）にリンクがあれば真、無ければ偽" {
        $base = "$TestDrive\haslink"
        writeListFile "$base\a\b\c.tsv" @("c")
        if (!(newTestJunction "$base\a\リンク" "$TestDrive\haslink_outside\下")) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        testPathHasLink $base "$base\a\b" | Should -Be $false
        testPathHasLink $base "$base\a\リンク" | Should -Be $true
        testPathHasLink $base "$base\a\リンク\下" | Should -Be $true
        testPathHasLink $base "$base\a\リンク\まだ無い" | Should -Be $true
        testPathHasLink $base "$base\a\b\まだ無い\下" | Should -Be $false
        testPathHasLink $base $base | Should -Be $false
        testPathHasLink "$base\a\" "$base\a\b" | Should -Be $false
    }

    It "testPathHasLink: base そのものがリンクでも、base は見ないので、その下に何も無ければ偽" {
        if (!(newTestJunction "$TestDrive\haslink_base" "$TestDrive\haslink_base_target\置き場")) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        writeListFile "$TestDrive\haslink_base\x\y.tsv" @("y")
        testPathHasLink "$TestDrive\haslink_base" "$TestDrive\haslink_base\x\y.tsv" | Should -Be $false
    }

    It "testPathHasLink: path が base の下になければ、確かめられないので真" {
        testPathHasLink "$TestDrive\haslink_a" "$TestDrive\haslink_b\x" | Should -Be $true
        testPathHasLink "$TestDrive\haslink_a" "$TestDrive\haslink_a2\x" | Should -Be $true
    }
}

Describe "moveDirectoryRetry" -Tag Io {
    It "フォルダを中身ごと移動する" {
        $source = "$TestDrive\移動元フォルダ"
        $destination = "$TestDrive\移動先フォルダ"
        [System.IO.Directory]::CreateDirectory("$source\中") | Out-Null
        writeListFile "$source\中\a.tsv" @("a")
        moveDirectoryRetry $source $destination
        Test-Path -LiteralPath $source | Should -Be $false
        Test-Path -LiteralPath "$destination\中\a.tsv" | Should -Be $true
    }

    It "中のファイルがほかから開かれていて移動できなければ、決めた回数だけ試してから例外にし、移動元が残る" {
        Mock Start-Sleep {}
        $source = "$TestDrive\使用中の移動元"
        $destination = "$TestDrive\使用中の移動先"
        writeListFile "$source\a.tsv" @("a")
        $stream = [System.IO.File]::Open("$source\a.tsv", "Open", "Read", "None")
        try {
            { moveDirectoryRetry $source $destination 3 1 } | Should -Throw
        } finally {
            $stream.Dispose()
        }
        Should -Invoke Start-Sleep -Times 2 -Exactly -Scope It
        Test-Path -LiteralPath $source | Should -Be $true
        Test-Path -LiteralPath $destination | Should -Be $false
    }

    It "待っている間に掴んでいたファイルが閉じられれば、次の試しで移動できる" {
        $source = "$TestDrive\途中で解放される移動元"
        $destination = "$TestDrive\途中で解放される移動先"
        writeListFile "$source\a.tsv" @("a")
        $stream = [System.IO.File]::Open("$source\a.tsv", "Open", "Read", "None")
        Mock Start-Sleep { $stream.Dispose() }
        moveDirectoryRetry $source $destination 3 1
        Test-Path -LiteralPath $source | Should -Be $false
        Test-Path -LiteralPath "$destination\a.tsv" | Should -Be $true
        Should -Invoke Start-Sleep -Times 1 -Exactly -Scope It
    }

    It "試すたびに待つ時間が倍になる" {
        Mock Start-Sleep {}
        $source = "$TestDrive\待ち時間が倍になる移動元"
        $destination = "$TestDrive\待ち時間が倍になる移動先"
        writeListFile "$source\a.tsv" @("a")
        $stream = [System.IO.File]::Open("$source\a.tsv", "Open", "Read", "None")
        try {
            { moveDirectoryRetry $source $destination 4 10 } | Should -Throw
        } finally {
            $stream.Dispose()
        }
        Should -Invoke Start-Sleep -Times 1 -Exactly -Scope It -ParameterFilter { $Milliseconds -eq 10 }
        Should -Invoke Start-Sleep -Times 1 -Exactly -Scope It -ParameterFilter { $Milliseconds -eq 20 }
        Should -Invoke Start-Sleep -Times 1 -Exactly -Scope It -ParameterFilter { $Milliseconds -eq 40 }
    }
}

Describe "testAnyEntry" -Tag Io {
    It "<name>" -TestCases @(
        @{ name = "条件なし: 1 つでもあれば true"; pattern = "*"; predicate = $null; expected = $true }
        @{ name = "条件なし: 無ければ false"; pattern = "*.xlsx"; predicate = $null; expected = $false }
        @{ name = "条件あり: 合うものがあれば true"; pattern = "*"; predicate = { param ($p) $p.EndsWith("b.tsv") }; expected = $true }
        @{ name = "条件あり: 合うものが無ければ false"; pattern = "*"; predicate = { param ($p) $p.EndsWith("c.tsv") }; expected = $false }
    ) {
        param ($name, $pattern, $predicate, $expected)
        $dir = "$TestDrive\any_$([Guid]::NewGuid().ToString('N'))"
        [System.IO.Directory]::CreateDirectory("$dir\営業\見積") | Out-Null
        [System.IO.File]::WriteAllText("$dir\営業\見積\a.tsv", "a")
        [System.IO.File]::WriteAllText("$dir\営業\見積\b.tsv", "b")

        testAnyEntry ([System.IO.Directory]::EnumerateFiles((toLongPath $dir), $pattern, [System.IO.SearchOption]::AllDirectories)) $predicate | Should -Be $expected

        # 途中でやめても、調べていた下のフォルダを掴んだまま残さない（すぐに移動できる）
        { [System.IO.Directory]::Move("$dir\営業", "$dir\moved") } | Should -Not -Throw
    }
}

Describe "selectFirstEntries" -Tag Io {
    It "先頭から count 件（足りなければあるだけ）を返し、途中でやめても下のフォルダを掴んだまま残さない: <count> 件" -TestCases @(
        @{ count = 0; expected = 0 }
        @{ count = 2; expected = 2 }
        @{ count = 10; expected = 3 }
    ) {
        param ($count, $expected)
        $dir = "$TestDrive\first_$([Guid]::NewGuid().ToString('N'))"
        [System.IO.Directory]::CreateDirectory("$dir\営業\見積") | Out-Null
        foreach ($name in @("a.tsv", "b.tsv", "c.tsv")) {
            [System.IO.File]::WriteAllText("$dir\営業\見積\$name", "x")
        }

        $entries = selectFirstEntries ([System.IO.Directory]::EnumerateFiles((toLongPath $dir), "*", [System.IO.SearchOption]::AllDirectories)) $count
        $entries.Count | Should -Be $expected
        @($entries | Where-Object { $_ -notlike "*\営業\見積\*.tsv" }).Count | Should -Be 0

        { [System.IO.Directory]::Move("$dir\営業", "$dir\moved") } | Should -Not -Throw
    }
}

Describe "writeTextLinesAtomic" -Tag Io {
    It "新しいファイルを作り、一時ファイルを残さない" {
        $path = "$TestDrive\atomic\新規.txt"
        writeTextLinesAtomic $path @("一", "二")
        @(readListFile $path) -join "," | Should -Be "一,二"
        Test-Path -LiteralPath "${path}.tmp" | Should -Be $false
    }

    It "既存のファイルを置き換える" {
        $path = "$TestDrive\atomic\置換.txt"
        writeTextLinesAtomic $path @("前")
        writeTextLinesAtomic $path @("後")
        @(readListFile $path) -join "," | Should -Be "後"
        Test-Path -LiteralPath "${path}.tmp" | Should -Be $false
    }

    It "置き換えられなければ 5 回まで試してから例外にし、元のファイルは壊さない" {
        Mock Start-Sleep {}
        $path = "$TestDrive\atomic\使用中.txt"
        writeTextLinesAtomic $path @("元の内容")
        $stream = [System.IO.File]::Open($path, "Open", "Read", "None")
        try {
            { writeTextLinesAtomic $path @("新しい内容") } | Should -Throw
        } finally {
            $stream.Dispose()
        }
        Should -Invoke Start-Sleep -Times 4 -Exactly -Scope It
        @(readListFile $path) -join "," | Should -Be "元の内容"
    }
}

Describe "getPathErrorKind" -Tag Unit {
    BeforeAll {
        function newWin32Exception {
            # Win32 のエラー番号を HResult に持つ IOException を作る（0x8007<エラー番号の16進>）
            param ([string]$message, [int]$win32Code)
            $hresult = [int]([long]0x80070000 + $win32Code)
            return New-Object System.IO.IOException($message, $hresult)
        }
    }

    It "<name>" -TestCases @(
        @{ name = "FileNotFoundException は見つからない"; kind = "FileNotFound"; expected = "Missing" }
        @{ name = "DirectoryNotFoundException は見つからない"; kind = "DirectoryNotFound"; expected = "Missing" }
        @{ name = "ERROR_BAD_NETPATH(53) は接続できない"; kind = 53; expected = "Unreachable" }
        @{ name = "ERROR_SEM_TIMEOUT(121) は接続できない"; kind = 121; expected = "Unreachable" }
        @{ name = "ERROR_HOST_UNREACHABLE(1232) は接続できない"; kind = 1232; expected = "Unreachable" }
        @{ name = "アクセス拒否(5) はその他"; kind = "AccessDenied"; expected = "Other" }
        @{ name = "分類の無い例外はその他"; kind = "Other"; expected = "Other" }
    ) {
        param ($name, $kind, $expected)
        $exception = switch ($kind) {
            "FileNotFound" { New-Object System.IO.FileNotFoundException("無い", "a.txt") }
            "DirectoryNotFound" { New-Object System.IO.DirectoryNotFoundException("無い") }
            "AccessDenied" { New-Object System.UnauthorizedAccessException("拒否") }
            "Other" { New-Object System.Exception("不明") }
            default { newWin32Exception "届かない" $kind }
        }
        getPathErrorKind $exception | Should -Be $expected
        # MethodInvocationException に包まれた形でも同じに分かる
        $wrapped = New-Object System.Management.Automation.MethodInvocationException("呼び出しに失敗しました", $exception)
        getPathErrorKind $wrapped | Should -Be $expected
    }
}

Describe "getPathState" -Tag Io {
    BeforeAll {
        $dir = "$TestDrive\pathState\見積"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::WriteAllText("$dir\a.xlsx", "x")
    }

    It "ファイルが見つかれば State=Found・IsDirectory=`$false" {
        $state = getPathState "$dir\a.xlsx"
        $state.State | Should -Be "Found"
        $state.IsDirectory | Should -Be $false
        $state.Message | Should -Be ""
    }

    It "フォルダが見つかれば State=Found・IsDirectory=`$true" {
        $state = getPathState $dir
        $state.State | Should -Be "Found"
        $state.IsDirectory | Should -Be $true
    }

    It "無いファイルは State=Missing（上のフォルダはある）" {
        (getPathState "$dir\無い.xlsx").State | Should -Be "Missing"
    }

    It "上のフォルダも無ければ State=Missing" {
        (getPathState "$TestDrive\無いフォルダ\無い.xlsx").State | Should -Be "Missing"
    }

    It "260 文字を超えるパスも \\?\ を付けて調べられる" {
        $deep = "$TestDrive\pathStateLong\" + ("見積" * 100)
        [void][System.IO.Directory]::CreateDirectory((toLongPath $deep))
        try {
            ($deep.Length) | Should -BeGreaterThan 260
            (getPathState $deep).State | Should -Be "Found"
        } finally {
            removeDirectoryRetry "$TestDrive\pathStateLong"
        }
    }
}

Describe "setNotContentIndexed" -Tag Io {
    BeforeAll {
        function script:hasAttr {
            param ([string]$path)
            $attrs = [System.IO.File]::GetAttributes((toLongPath $path))
            return [bool]($attrs -band [System.IO.FileAttributes]::NotContentIndexed)
        }
    }

    It "-Recurse ありで、フォルダ自身・下のフォルダ・ファイル（長いパスを含む）に付く" {
        $dir = "$TestDrive\nci\付ける"
        $sub1 = "$dir\下"
        $sub2 = "$sub1\" + ("永" * 200)
        $long = "$sub2\long.tsv"
        [void][System.IO.Directory]::CreateDirectory((toLongPath $sub2))
        [System.IO.File]::WriteAllText((toLongPath $long), "x")
        ($long.Length) | Should -BeGreaterThan 260
        writeListFile "$dir\a.tsv" @("a")

        try {
            $result = setNotContentIndexed $dir -Recurse
            $result.Ok | Should -Be $true
            $result.Changed | Should -Be 5   # dir・a.tsv・下・長い名前のフォルダ・long.tsv
            (hasAttr $dir) | Should -Be $true
            (hasAttr "$dir\a.tsv") | Should -Be $true
            (hasAttr $sub1) | Should -Be $true
            (hasAttr $sub2) | Should -Be $true
            (hasAttr $long) | Should -Be $true
        } finally {
            # Pester の TestDrive の後始末は \\?\ を付けずに消すため、260 文字超のパスを自分で先に消しておく
            Remove-Item -LiteralPath (toLongPath $dir) -Recurse -Force
        }
    }

    It "-Recurse なしでは、フォルダ自身と直下のファイルだけに付き、下のフォルダには付かない" {
        $dir = "$TestDrive\nci\直下だけ"
        [System.IO.Directory]::CreateDirectory("$dir\下") | Out-Null
        writeListFile "$dir\a.tsv" @("a")
        writeListFile "$dir\下\b.tsv" @("b")

        $result = setNotContentIndexed $dir
        $result.Ok | Should -Be $true
        $result.Changed | Should -Be 2   # dir・a.tsv だけ
        (hasAttr $dir) | Should -Be $true
        (hasAttr "$dir\a.tsv") | Should -Be $true
        (hasAttr "$dir\下") | Should -Be $false
        (hasAttr "$dir\下\b.tsv") | Should -Be $false
    }

    It "既に付いているものは Changed に数えない（2 回目は 0）" {
        $dir = "$TestDrive\nci\2回目"
        writeListFile "$dir\a.tsv" @("a")

        (setNotContentIndexed $dir -Recurse).Changed | Should -Be 2
        (setNotContentIndexed $dir -Recurse).Changed | Should -Be 0
    }

    It "根に付いていれば、下に付いていないものがあっても変えない（下はたどらない）" {
        $dir = "$TestDrive\nci\根に付いている"
        writeListFile "$dir\a.tsv" @("a")
        setNotContentIndexed $dir -Recurse | Out-Null
        (hasAttr $dir) | Should -Be $true

        # 新しく作った（WriteAllLines で作った）ものは根の属性を継ぐため、Move で入れて属性を継がせない
        # （実装が「入れた直後にそれぞれの場所で付ける」ことに頼らず、根の属性だけで判断することを確かめる）
        $src = "$TestDrive\nci\移す元\b.tsv"
        writeListFile $src @("b")
        [System.IO.File]::Move($src, "$dir\b.tsv")
        (hasAttr "$dir\b.tsv") | Should -Be $false

        $result = setNotContentIndexed $dir -Recurse
        $result.Ok | Should -Be $true
        $result.Changed | Should -Be 0
        (hasAttr "$dir\b.tsv") | Should -Be $false
    }

    It "途中で失敗した回は根に付けない（次回また下からたどれるようにする）" {
        $dir = "$TestDrive\nci\失敗した回"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $locked = New-Object System.IO.DirectoryInfo "$dir\読めない"
        $locked.Create()
        $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $deny = New-Object System.Security.AccessControl.FileSystemAccessRule($user, "ListDirectory", "Deny")
        $acl = $locked.GetAccessControl("Access")
        $acl.AddAccessRule($deny)
        $locked.SetAccessControl($acl)
        try {
            $result = setNotContentIndexed $dir -Recurse
            $result.Ok | Should -Be $false
            (hasAttr $dir) | Should -Be $false
        } finally {
            $acl = $locked.GetAccessControl("Access")
            [void]$acl.RemoveAccessRule($deny)
            $locked.SetAccessControl($acl)
        }
    }

    It "隣のフォルダには付かない" {
        $dir = "$TestDrive\nci\隣"
        [System.IO.Directory]::CreateDirectory("$dir\対象") | Out-Null
        [System.IO.Directory]::CreateDirectory("$dir\隣のフォルダ") | Out-Null

        setNotContentIndexed "$dir\対象" -Recurse | Out-Null
        (hasAttr "$dir\隣のフォルダ") | Should -Be $false
    }

    It "ほかの属性（読み取り専用など）は変えない" {
        $dir = "$TestDrive\nci\属性"
        $path = "$dir\a.tsv"
        writeListFile $path @("a")
        [System.IO.File]::SetAttributes($path, [System.IO.FileAttributes]::ReadOnly)
        try {
            setNotContentIndexed $dir -Recurse | Out-Null
            $attrs = [System.IO.File]::GetAttributes($path)
            ($attrs -band [System.IO.FileAttributes]::ReadOnly) | Should -Not -Be 0
            ($attrs -band [System.IO.FileAttributes]::NotContentIndexed) | Should -Not -Be 0
        } finally {
            [System.IO.File]::SetAttributes($path, [System.IO.FileAttributes]::Normal)
        }
    }

    It "無いパスは例外にならず Ok=`$false・Reason を返す" {
        $result = setNotContentIndexed "$TestDrive\nci\無いフォルダ"
        $result.Ok | Should -Be $false
        $result.Changed | Should -Be 0
        $result.Reason | Should -Not -BeNullOrEmpty
    }

    It "列挙の途中で読めなくなるフォルダがあっても例外にならず Ok=`$false・Reason を返す" {
        $dir = "$TestDrive\nci\読めない親"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $locked = New-Object System.IO.DirectoryInfo "$dir\読めない"
        $locked.Create()
        $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $deny = New-Object System.Security.AccessControl.FileSystemAccessRule($user, "ListDirectory", "Deny")
        $acl = $locked.GetAccessControl("Access")
        $acl.AddAccessRule($deny)
        $locked.SetAccessControl($acl)
        try {
            $result = setNotContentIndexed $dir -Recurse
            $result.Ok | Should -Be $false
            $result.Reason | Should -Not -BeNullOrEmpty
        } finally {
            $acl = $locked.GetAccessControl("Access")
            [void]$acl.RemoveAccessRule($deny)
            $locked.SetAccessControl($acl)
        }
    }
}

Describe "getFilesNoFollow" -Tag Io {
    It "下のファイル・フォルダを集める（root は Directories に入れない）。Links は空、HasError は `$false" {
        $dir = "$TestDrive\nofollow\普通"
        writeListFile "$dir\a.tsv" @("a")
        writeListFile "$dir\下\b.tsv" @("b")
        writeListFile "$dir\下\さらに下\c.tsv" @("c")

        $tree = getFilesNoFollow $dir
        @($tree.Files | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @("a.tsv", "b.tsv", "c.tsv")
        @($tree.Directories | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @("さらに下", "下")
        @($tree.Links).Count | Should -Be 0
        $tree.HasError | Should -Be $false
    }

    It "末尾に \ を付けた root でも同じ結果になる" {
        $dir = "$TestDrive\nofollow\末尾"
        writeListFile "$dir\a.tsv" @("a")

        @((getFilesNoFollow "$dir\").Files).Count | Should -Be 1
    }

    It "Hidden のものは、既定では入れず、IncludeHidden を付けると入れる" {
        $dir = "$TestDrive\nofollow\隠し"
        writeListFile "$dir\a.tsv" @("a")
        writeListFile "$dir\隠し.tsv" @("h")
        writeListFile "$dir\隠しフォルダ\d.tsv" @("d")
        [System.IO.File]::SetAttributes("$dir\隠し.tsv", [System.IO.FileAttributes]::Hidden)
        [System.IO.File]::SetAttributes("$dir\隠しフォルダ", [System.IO.FileAttributes]::Directory -bor [System.IO.FileAttributes]::Hidden)

        @((getFilesNoFollow $dir).Files | ForEach-Object { $_.Name }) | Should -Be @("a.tsv")
        @((getFilesNoFollow $dir -IncludeHidden).Files | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @("a.tsv", "d.tsv", "隠し.tsv")
    }

    It "System だけのファイル・フォルダは、既定でも入れる（Get-ChildItem -Recurse が返す範囲と同じ）" {
        $dir = "$TestDrive\nofollow\システム"
        writeListFile "$dir\a.tsv" @("a")
        writeListFile "$dir\sys.tsv" @("s")
        writeListFile "$dir\sysフォルダ\d.tsv" @("d")
        $system = [System.IO.FileAttributes]::System
        [System.IO.File]::SetAttributes("$dir\sys.tsv", $system)
        [System.IO.File]::SetAttributes("$dir\sysフォルダ", [System.IO.FileAttributes]::Directory -bor $system)

        $expected = @(Get-ChildItem -LiteralPath $dir -Recurse -File | ForEach-Object { $_.Name } | Sort-Object)
        $expected | Should -Be @("a.tsv", "d.tsv", "sys.tsv")
        @((getFilesNoFollow $dir).Files | ForEach-Object { $_.Name } | Sort-Object) | Should -Be $expected
    }

    It "並びは Get-ChildItem -Recurse と同じ（そのフォルダの中身を並べてから、サブフォルダの下へ進む）" {
        $dir = "$TestDrive\nofollow\並び"
        writeListFile "$dir\z.tsv" @("z")
        writeListFile "$dir\b\b2.tsv" @("b")
        writeListFile "$dir\b\b1\深.tsv" @("b")
        writeListFile "$dir\b\b1\b0\最深.tsv" @("b")
        writeListFile "$dir\a\a1.tsv" @("a")
        writeListFile "$dir\c.tsv" @("c")

        $tree = getFilesNoFollow $dir
        @($tree.Files | ForEach-Object { fromLongPath $_.FullName }) | Should -Be @(Get-ChildItem -LiteralPath $dir -Recurse -File | ForEach-Object { $_.FullName })
        @($tree.Directories | ForEach-Object { fromLongPath $_.FullName }) | Should -Be @(Get-ChildItem -LiteralPath $dir -Recurse -Directory | ForEach-Object { $_.FullName })
    }

    It "260 文字を超える長いパスの下も列挙する" {
        $dir = "$TestDrive\nofollowLong"
        $deep = "$dir\" + ("a" * 100) + "\" + ("b" * 100)
        [void][System.IO.Directory]::CreateDirectory((toLongPath $deep))
        [System.IO.File]::WriteAllText((toLongPath "$deep\深い.tsv"), "x`r`n", $utf8Bom)
        writeListFile "$dir\浅い.tsv" @("y")
        try {
            "$deep\深い.tsv".Length | Should -BeGreaterThan 260
            $tree = getFilesNoFollow $dir
            $tree.HasError | Should -Be $false
            @($tree.Files).Count | Should -Be 2
            @($tree.Files | ForEach-Object { $_.Name }) | Should -Contain "浅い.tsv"
            @($tree.Files | ForEach-Object { $_.Name }) | Should -Contain "深い.tsv"
        } finally {
            removeTreeNoFollow $dir
        }
    }

    It "ジャンクションの先へは降りず、Links に入れる（Files・Directories には入れない）" {
        $dir = "$TestDrive\nofollow\リンク"
        $outside = "$TestDrive\nofollow\外のリンク"
        writeListFile "$outside\外.tsv" @("外")
        writeListFile "$dir\a.tsv" @("a")
        if (!(newTestJunction "$dir\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $tree = getFilesNoFollow $dir
        @($tree.Files | ForEach-Object { $_.Name }) | Should -Be @("a.tsv")
        @($tree.Directories).Count | Should -Be 0
        @($tree.Links) | Should -Be @("$dir\リンク")
        $tree.HasError | Should -Be $false
    }

    It "Hidden のジャンクションは、既定では Files・Links のどちらにも入れず、IncludeHidden を付けると Links に入れる（中へは降りない）" {
        $dir = "$TestDrive\nofollow\隠しリンク"
        $outside = "$TestDrive\nofollow\隠しリンクの外"
        writeListFile "$outside\外.tsv" @("外")
        writeListFile "$dir\a.tsv" @("a")
        if (!(newTestJunction "$dir\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        [System.IO.File]::SetAttributes("$dir\リンク", [System.IO.File]::GetAttributes("$dir\リンク") -bor [System.IO.FileAttributes]::Hidden)

        $default = getFilesNoFollow $dir
        @($default.Files | ForEach-Object { $_.Name }) | Should -Be @("a.tsv")
        @($default.Links).Count | Should -Be 0
        $all = getFilesNoFollow $dir -IncludeHidden
        @($all.Files | ForEach-Object { $_.Name }) | Should -Be @("a.tsv")
        @($all.Links) | Should -Be @("$dir\リンク")
    }

    It "親を指すジャンクション（輪）があっても止まらず、Links に入れる" {
        $dir = "$TestDrive\nofollow\輪"
        writeListFile "$dir\a\a.tsv" @("a")
        if (!(newTestJunction "$dir\a\輪" $dir)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $tree = getFilesNoFollow $dir
        @($tree.Files | ForEach-Object { $_.Name }) | Should -Be @("a.tsv")
        @($tree.Links) | Should -Be @("$dir\a\輪")
        $tree.HasError | Should -Be $false
    }

    It "ファイルのシンボリックリンクは Files に入れず、Links に入れる（作れなければ飛ばす）" {
        $dir = "$TestDrive\nofollow\ファイルのリンク"
        writeListFile "$dir\a.tsv" @("a")
        writeListFile "$TestDrive\nofollow\外のファイル\外.tsv" @("外")
        if (!(newTestSymlink "$dir\リンク.tsv" "$TestDrive\nofollow\外のファイル\外.tsv" -file)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $tree = getFilesNoFollow $dir
        @($tree.Files | ForEach-Object { $_.Name }) | Should -Be @("a.tsv")
        @($tree.Links) | Should -Be @("$dir\リンク.tsv")
    }

    It "root そのものがジャンクションでも見ない（先の中を列挙する）" {
        $outside = "$TestDrive\nofollow\外の根"
        writeListFile "$outside\外.tsv" @("外")
        if (!(newTestJunction "$TestDrive\nofollow\根のリンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $tree = getFilesNoFollow "$TestDrive\nofollow\根のリンク"
        @($tree.Files | ForEach-Object { $_.Name }) | Should -Be @("外.tsv")
        @($tree.Links).Count | Should -Be 0
    }

    It "リンクかどうか確かめられないフォルダは、既定では降り、SkipUnknown を付けると降りずに Links に入れる" {
        $dir = "$TestDrive\nofollow\不明"
        $outside = "$TestDrive\nofollow\不明の中"
        writeListFile "$outside\中.tsv" @("中")
        if (!(newTestJunction "$dir\不明" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        # ReparsePoint が付き LinkType が取れないフォルダ（OneDrive のオンデマンドのフォルダなど）の見本として、種類を Unknown にする
        Mock getLinkKind { return "Unknown" }

        @((getFilesNoFollow $dir).Files | ForEach-Object { $_.Name }) | Should -Be @("中.tsv")
        $skipped = getFilesNoFollow $dir -SkipUnknown
        @($skipped.Files).Count | Should -Be 0
        @($skipped.Links) | Should -Be @("$dir\不明")
    }

    It "読めないフォルダがあっても例外にせず、HasError と Reason を返し、読めたところは返す" {
        $dir = "$TestDrive\nofollow\読めない親"
        writeListFile "$dir\a.tsv" @("a")
        $locked = New-Object System.IO.DirectoryInfo "$dir\読めない"
        $locked.Create()
        $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $deny = New-Object System.Security.AccessControl.FileSystemAccessRule($user, "ListDirectory", "Deny")
        $acl = $locked.GetAccessControl("Access")
        $acl.AddAccessRule($deny)
        $locked.SetAccessControl($acl)
        try {
            $tree = getFilesNoFollow $dir
            $tree.HasError | Should -Be $true
            $tree.Reason | Should -Not -BeNullOrEmpty
            @($tree.Files | ForEach-Object { $_.Name }) | Should -Be @("a.tsv")
        } finally {
            $acl = $locked.GetAccessControl("Access")
            [void]$acl.RemoveAccessRule($deny)
            $locked.SetAccessControl($acl)
        }
    }
}

Describe "setNotContentIndexed（リンク）" -Tag Io {
    BeforeAll {
        function script:hasLinkAttr {
            param ([string]$path)
            $attrs = [System.IO.File]::GetAttributes((toLongPath $path))
            return [bool]($attrs -band [System.IO.FileAttributes]::NotContentIndexed)
        }
    }

    It "-Recurse で、中のジャンクションの先のフォルダ・ファイルにも、ジャンクション自身にも付けない" {
        $dir = "$TestDrive\nci-link\付ける"
        $outside = "$TestDrive\nci-link\外"
        writeListFile "$outside\外.tsv" @("外")
        writeListFile "$outside\下\外の下.tsv" @("外")
        writeListFile "$dir\a.tsv" @("a")
        if (!(newTestJunction "$dir\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $result = setNotContentIndexed $dir -Recurse
        $result.Ok | Should -Be $true
        $result.Changed | Should -Be 2   # dir・a.tsv だけ
        (hasLinkAttr $dir) | Should -Be $true
        (hasLinkAttr "$dir\a.tsv") | Should -Be $true
        (hasLinkAttr "$dir\リンク") | Should -Be $false
        (hasLinkAttr $outside) | Should -Be $false
        (hasLinkAttr "$outside\外.tsv") | Should -Be $false
        (hasLinkAttr "$outside\下") | Should -Be $false
        (hasLinkAttr "$outside\下\外の下.tsv") | Should -Be $false
    }

    It "-Recurse で読めないフォルダがあると Ok=`$false にし、根にも属性を付けない" {
        $dir = "$TestDrive\nci-link\読めない"
        writeListFile "$dir\a.tsv" @("a")
        $locked = New-Object System.IO.DirectoryInfo "$dir\読めない"
        $locked.Create()
        $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $deny = New-Object System.Security.AccessControl.FileSystemAccessRule($user, "ListDirectory", "Deny")
        $acl = $locked.GetAccessControl("Access")
        $acl.AddAccessRule($deny)
        $locked.SetAccessControl($acl)
        try {
            $result = setNotContentIndexed $dir -Recurse
            $result.Ok | Should -Be $false
            $result.Changed | Should -Be 0
            $result.Reason | Should -Not -BeNullOrEmpty
            (hasLinkAttr $dir) | Should -Be $false
            (hasLinkAttr "$dir\a.tsv") | Should -Be $false
        } finally {
            $acl = $locked.GetAccessControl("Access")
            [void]$acl.RemoveAccessRule($deny)
            $locked.SetAccessControl($acl)
        }
    }

    It "親を指すジャンクション（輪）があっても終わり、輪の先に属性が付き続けない" {
        $dir = "$TestDrive\nci-link\輪"
        writeListFile "$dir\a\a.tsv" @("a")
        if (!(newTestJunction "$dir\a\輪" $dir)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $result = setNotContentIndexed $dir -Recurse
        $result.Ok | Should -Be $true
        $result.Changed | Should -Be 3   # dir・a・a.tsv
        (hasLinkAttr "$dir\a\輪") | Should -Be $false
    }

    It "Hidden のファイル・フォルダには、これまでどおり付ける" {
        $dir = "$TestDrive\nci-link\隠し"
        writeListFile "$dir\隠し.tsv" @("h")
        writeListFile "$dir\隠しフォルダ\d.tsv" @("d")
        [System.IO.File]::SetAttributes("$dir\隠し.tsv", [System.IO.FileAttributes]::Hidden)
        [System.IO.File]::SetAttributes("$dir\隠しフォルダ", [System.IO.FileAttributes]::Directory -bor [System.IO.FileAttributes]::Hidden)

        (setNotContentIndexed $dir -Recurse).Ok | Should -Be $true
        (hasLinkAttr "$dir\隠し.tsv") | Should -Be $true
        (hasLinkAttr "$dir\隠しフォルダ") | Should -Be $true
        (hasLinkAttr "$dir\隠しフォルダ\d.tsv") | Should -Be $true
    }
}

Describe "testNotContentIndexed" -Tag Io {
    It "属性が付いていれば `$true" {
        $dir = "$TestDrive\nci-test\付いている"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        setNotContentIndexed $dir | Out-Null

        testNotContentIndexed $dir | Should -Be $true
    }

    It "属性が付いていなければ `$false" {
        $dir = "$TestDrive\nci-test\付いていない"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null

        testNotContentIndexed $dir | Should -Be $false
    }

    It "無いパスは例外にならず `$false" {
        testNotContentIndexed "$TestDrive\nci-test\無いフォルダ" | Should -Be $false
    }
}

Describe "getFolderKey" -Tag Unit {
    It "SHA-256 の 16 進 64 文字を返す" {
        # "c:\tool" の SHA-256（小文字にしてから UTF-8 で計算する）
        getFolderKey "C:\Tool" | Should -Be "DA4B936E296325CC586C02ADB4518ABB2C4021761A7AFCE2D670263F384C89F5"
    }
}

Describe "newAppMutex" -Tag Io {
    It "同じ処理・同じフォルダでは2つ目を取得できない" {
        $first = newAppMutex "test" "$TestDrive\tool"
        try {
            $first.Acquired | Should -Be $true
            $second = newAppMutex "test" "$TestDrive\tool"
            try {
                $second.Acquired | Should -Be $false
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
            $gui.Acquired | Should -Be $true
            $indexer.Acquired | Should -Be $true
            $other.Acquired | Should -Be $true
        } finally {
            foreach ($m in @($gui, $indexer, $other)) {
                $m.Mutex.ReleaseMutex()
                $m.Mutex.Dispose()
            }
        }
    }
}

Describe "writeTextLinesAtomic（文字コード）" -Tag Io {
    It "既定は BOM 付き UTF-8、encoding を渡せば BOM なし UTF-8 で書く" {
        $path = "$TestDrive\atomic\文字コード.txt"
        writeTextLinesAtomic $path @("あ")
        [System.IO.File]::ReadAllBytes($path)[0..2] -join "," | Should -Be "239,187,191"
        writeTextLinesAtomic $path @("あ") (New-Object System.Text.UTF8Encoding($false))
        [System.IO.File]::ReadAllBytes($path)[0] | Should -Not -Be 239
        [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8).Trim() | Should -Be "あ"
    }
}

Describe "invokeWithNamedMutex" -Tag Io {
    BeforeAll {
        # 別のスレッド（新しいスレッドで動くランスペース）でミューテックスを取る。
        # release が $true なら合図があるまで持ち続け、$false なら取ったまま終わる（abandoned になる）
        function startMutexHolder {
            param ([string]$name, [bool]$release)
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread
            $runspace.Open()
            $ps = [powershell]::Create()
            $ps.Runspace = $runspace
            $got = New-Object System.Threading.ManualResetEvent($false)
            $stop = New-Object System.Threading.ManualResetEvent($false)
            [void]$ps.AddScript({
                param ($name, $release, $got, $stop)
                $m = New-Object System.Threading.Mutex($false, $name)
                [void]$m.WaitOne()
                [void]$got.Set()
                if ($release) {
                    [void]$stop.WaitOne()
                    $m.ReleaseMutex()
                }
            }).AddArgument($name).AddArgument($release).AddArgument($got).AddArgument($stop)
            $handle = $ps.BeginInvoke()
            $got.WaitOne(10000) | Should -Be $true
            return @{ Ps = $ps; Runspace = $runspace; Handle = $handle; Stop = $stop }
        }
        function closeMutexHolder {
            param ($holder)
            [void]$holder.Stop.Set()
            [void]$holder.Ps.EndInvoke($holder.Handle)
            $holder.Ps.Dispose()
            $holder.Runspace.Dispose()
        }
    }

    It "action の出力を返し、終わったら手放す（続けて取れる）" {
        $name = "Local\tebunko_test_mutex_$([guid]::NewGuid().ToString('N'))"
        invokeWithNamedMutex $name 1000 { 42 } | Should -Be 42
        invokeWithNamedMutex $name 1000 { 43 } | Should -Be 43
    }

    It "action が例外でも手放す" {
        $name = "Local\tebunko_test_mutex_$([guid]::NewGuid().ToString('N'))"
        { invokeWithNamedMutex $name 1000 { throw "失敗" } } | Should -Throw "失敗"
        invokeWithNamedMutex $name 1000 { "取れた" } | Should -Be "取れた"
    }

    It "同じスレッドの入れ子は通す" {
        $name = "Local\tebunko_test_mutex_$([guid]::NewGuid().ToString('N'))"
        $inner = { "内側" }
        invokeWithNamedMutex $name 1000 { invokeWithNamedMutex $name 1000 $inner } | Should -Be "内側"
    }

    It "別のスレッドが持ったままだと、決めた時間で分かる例外にする" {
        $name = "Local\tebunko_test_mutex_$([guid]::NewGuid().ToString('N'))"
        $holder = startMutexHolder $name $true
        try {
            { invokeWithNamedMutex $name 200 { "実行されない" } } | Should -Throw "*排他の待ちが時間切れになりました*"
        } finally {
            closeMutexHolder $holder
        }
        invokeWithNamedMutex $name 1000 { "取れた" } | Should -Be "取れた"
    }

    It "持ったまま終わったスレッドの後は、abandoned として続ける" {
        $name = "Local\tebunko_test_mutex_$([guid]::NewGuid().ToString('N'))"
        $holder = startMutexHolder $name $false
        [void]$holder.Ps.EndInvoke($holder.Handle)
        $holder.Ps.Dispose()
        $holder.Runspace.Dispose()   # スレッドが終わる
        Start-Sleep -Milliseconds 300
        invokeWithNamedMutex $name 2000 { "続けた" } | Should -Be "続けた"
    }
}

Describe "testCloudOnlyAttributes" -Tag Unit {
    # 属性は数値で比べる。0x400000 = RECALL_ON_DATA_ACCESS・0x40000 = RECALL_ON_OPEN・0x1000 = OFFLINE
    # （ピン留め 0x80000・ピン留めの解除 0x100000・再解析ポイントだけでは、クラウドにだけあるとは言えない）
    It "<name>" -TestCases @(
        @{ name = "通常のファイルは違う"; attributes = 0x80; expected = $false }
        @{ name = "Archive は違う"; attributes = 0x20; expected = $false }
        @{ name = "ReadOnly は違う"; attributes = 0x1; expected = $false }
        @{ name = "ReparsePoint だけでは違う"; attributes = 0x400; expected = $false }
        @{ name = "ピン留め（0x80000）は違う"; attributes = (0x80000 -bor 0x20); expected = $false }
        @{ name = "ピン留めの解除（0x100000）は違う"; attributes = (0x100000 -bor 0x20); expected = $false }
        @{ name = "RECALL_ON_DATA_ACCESS（0x400000）はクラウドにだけある"; attributes = 0x400000; expected = $true }
        @{ name = "RECALL_ON_OPEN（0x40000）はクラウドにだけある"; attributes = 0x40000; expected = $true }
        @{ name = "OFFLINE（0x1000）はクラウドにだけある"; attributes = 0x1000; expected = $true }
        @{ name = "RECALL_ON_DATA_ACCESS に Archive と ReparsePoint が付いていても"; attributes = (0x400000 -bor 0x20 -bor 0x400); expected = $true }
        @{ name = "RECALL_ON_OPEN に Archive と ReparsePoint が付いていても"; attributes = (0x40000 -bor 0x20 -bor 0x400); expected = $true }
        @{ name = "OFFLINE に Archive と ReparsePoint が付いていても"; attributes = (0x1000 -bor 0x20 -bor 0x400); expected = $true }
        @{ name = "ピン留めとクラウドにだけある印が両方付いていればクラウドにだけある"; attributes = (0x80000 -bor 0x400000); expected = $true }
    ) {
        param ($name, $attributes, $expected)
        testCloudOnlyAttributes $attributes | Should -Be $expected
    }

    It "ファイル情報の Attributes（列挙型）をそのまま渡せる" {
        $path = Join-Path $TestDrive "cloud_attr.txt"
        [System.IO.File]::WriteAllText($path, "x")
        testCloudOnlyAttributes ([int](Get-Item -LiteralPath $path).Attributes) | Should -Be $false
        [System.IO.File]::SetAttributes($path, [System.IO.FileAttributes]::Offline)
        try {
            testCloudOnlyAttributes ([int](Get-Item -LiteralPath $path).Attributes) | Should -Be $true
        } finally {
            [System.IO.File]::SetAttributes($path, [System.IO.FileAttributes]::Normal)
        }
    }
}
