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
            $index.Folders[0].Count | Should -Be 2
            $rel = @($index.Packs | ForEach-Object { $_.RelPath } | Sort-Object)
            $rel | Should -Be @("$deepRel\content_index.xlsx.001.tsv", "content_index.xlsx.001.tsv")

            $hits = @((searchPackIndex "hello" $index.Packs).Hits)
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
