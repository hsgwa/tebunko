# ［新規で開く］のためのコピーの作成と片付け（tebunko\core\open_copy.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:getTmpRoot([string]$base) {
        return "$TestDrive\$base\tmp"
    }

    function script:newCopyDir([string]$base) {
        $dir = "$(getTmpRoot $base)\$(getMachineKey)\open_copy"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        return $dir
    }

    function script:newOldCopy([string]$copyDir, [string]$sub, [string]$name, [datetime]$created) {
        # 作ったのが古いコピー（1 つ下のフォルダごと）を置く。中のファイルも同じ日時にする
        $dir = Join-Path $copyDir $sub
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $dir $name), "copy")
        (Get-Item -LiteralPath (Join-Path $dir $name)).LastWriteTime = $created
        (Get-Item -LiteralPath $dir).LastWriteTime = $created
        (Get-Item -LiteralPath $dir).CreationTime = $created
        return $dir
    }

    function script:newJunction([string]$link, [string]$target) {
        # ジャンクションを作る（指す先は、呼ぶ側が $TestDrive の下にする）。作れない環境では $false
        [System.IO.Directory]::CreateDirectory($target) | Out-Null
        try {
            New-Item -ItemType Junction -Path $link -Target $target -ErrorAction Stop | Out-Null
            return $true
        } catch {
            return $false
        }
    }
}

Describe "getOpenCopyDir" -Tag Unit {
    It "ワークスペースの tmp の下の、PC の鍵のフォルダの open_copy を返す（フォルダは作らない）" {
        $workspace = newTestWorkspace @{ TmpRoot = "$TestDrive\dir\tmp" }
        getOpenCopyDir $workspace | Should -Be "$TestDrive\dir\tmp\$(getMachineKey)\open_copy"
        Test-Path -LiteralPath "$TestDrive\dir" | Should -Be $false
    }
}

Describe "testOpenCopyDir" -Tag Unit {
    BeforeAll {
        $script:key = getMachineKey
    }

    It "<name> は <expected>" -TestCases @(
        @{ name = "tmp\<PC の鍵>\open_copy"; dir = "C:\work\tmp\<key>\open_copy"; tmp = "C:\work\tmp"; expected = $true }
        @{ name = "末尾の区切りつき"; dir = "C:\work\tmp\<key>\open_copy\"; tmp = "C:\work\tmp"; expected = $true }
        @{ name = "大文字・小文字の違い"; dir = "C:\WORK\TMP\<key>\OPEN_COPY"; tmp = "C:\work\tmp"; expected = $true }
        @{ name = "空"; dir = ""; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "空白"; dir = "   "; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "tmp の指定が空"; dir = "C:\work\tmp\<key>\open_copy"; tmp = ""; expected = $false }
        @{ name = "ドライブの直下"; dir = "C:\"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "ドライブだけ"; dir = "C:"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "ルート"; dir = "\"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "相対パス"; dir = "work\tmp\<key>\open_copy"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "ドライブ無しの絶対パス"; dir = "\work\tmp\<key>\open_copy"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "open_copy でない名前"; dir = "C:\work\tmp\<key>"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "名前が似ているだけ"; dir = "C:\work\tmp\<key>\open_copy2"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "別のワークスペースの open_copy"; dir = "C:\other\tmp\<key>\open_copy"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "別の PC の鍵の open_copy"; dir = "C:\work\tmp\other\open_copy"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "tmp の直下の open_copy"; dir = "C:\work\tmp\open_copy"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = ".. を含む"; dir = "C:\work\tmp\x\..\<key>\open_copy"; tmp = "C:\work\tmp"; expected = $false }
        @{ name = "元のフォルダ"; dir = "C:\data\営業"; tmp = "C:\work\tmp"; expected = $false }
    ) {
        param ($name, $dir, $tmp, $expected)
        testOpenCopyDir $dir.Replace("<key>", $script:key) $tmp | Should -Be $expected
    }
}

Describe "testOpenCopyDir（リンク）" -Tag Io {
    It "open_copy 自身・その親（PC の鍵）・tmp がジャンクションなら、使わない（<name>）" -TestCases @(
        @{ name = "open_copy"; level = 0 }
        @{ name = "PC の鍵"; level = 1 }
        @{ name = "tmp"; level = 2 }
    ) {
        param ($name, $level)
        $base = "$TestDrive\link$level"
        $key = getMachineKey
        # 指す先は $TestDrive の下に作る（実体のフォルダ）
        $real = "$base\real"
        [System.IO.Directory]::CreateDirectory("$real\$key\open_copy") | Out-Null
        switch ($level) {
            0 {
                [System.IO.Directory]::CreateDirectory("$base\tmp\$key") | Out-Null
                $made = newJunction "$base\tmp\$key\open_copy" "$real\$key\open_copy"
            }
            1 {
                [System.IO.Directory]::CreateDirectory("$base\tmp") | Out-Null
                $made = newJunction "$base\tmp\$key" "$real\$key"
            }
            2 {
                $made = newJunction "$base\tmp" $real
            }
        }
        if (!$made) {
            Set-ItResult -Skipped -Because "この環境ではジャンクションを作れない"
            return
        }

        testOpenCopyDir "$base\tmp\$key\open_copy" "$base\tmp" | Should -Be $false
        # 比べとして、リンクでない同じ形は使える
        $plain = newCopyDir "plain$level"
        testOpenCopyDir $plain (getTmpRoot "plain$level") | Should -Be $true
    }
}

Describe "getOpenCopyName" -Tag Unit {
    It "<book> の名前は <expected>" -TestCases @(
        @{ book = "仕様書.pdf"; expected = "仕様書（コピー 20261012-1530）.pdf" }
        @{ book = "a.txt. "; expected = "a（コピー 20261012-1530）.txt" }
        @{ book = "x\..\..\悪い.txt"; expected = "悪い（コピー 20261012-1530）.txt" }
        @{ book = "x/悪い.txt"; expected = "悪い（コピー 20261012-1530）.txt" }
        @{ book = "a.tar.gz"; expected = "a.tar（コピー 20261012-1530）.gz" }
        @{ book = "拡張子なし"; expected = "拡張子なし（コピー 20261012-1530）" }
        @{ book = ".hidden"; expected = "（コピー 20261012-1530）.hidden" }
        @{ book = ""; expected = "" }
    ) {
        param ($book, $expected)
        getOpenCopyName $book ([datetime]"2026-10-12 15:30:45") | Should -Be $expected
    }
}

Describe "getOpenCopyBlocker" -Tag Unit {
    BeforeAll {
        $script:when = [datetime]"2026-10-12 15:30:00"
    }

    It "<name> は <expected>" -TestCases @(
        @{ name = "ふつうの名前"; book = "仕様書.pdf"; expected = "" }
        @{ name = "名前が空"; book = ""; expected = "NoName" }
        @{ name = "[ を含む"; book = "報告[1].txt"; expected = "Bracket" }
        @{ name = "] だけを含む"; book = "報告].txt"; expected = "Bracket" }
        @{ name = "長すぎる"; book = ("あ" * 250 + ".txt"); expected = "TooLong" }
    ) {
        param ($name, $book, $expected)
        getOpenCopyBlocker $book "C:\work\tmp\key\open_copy" $script:when | Should -Be $expected
    }
}

Describe "copySourceForOpen" -Tag Io {
    BeforeAll {
        $script:when = [datetime]"2026-10-12 15:30:00"
    }

    It "元のファイルの名前に（コピー 日時）を付けた名前で、同じ中身のコピーを open_copy の下に作り、元のファイルは変えない" {
        $copyDir = newCopyDir "copy1"
        [System.IO.File]::WriteAllText("$TestDrive\copy1\議事メモ.txt", "本文あり")
        $before = Get-Item -LiteralPath "$TestDrive\copy1\議事メモ.txt"
        $writeTime = $before.LastWriteTimeUtc
        $attributes = $before.Attributes

        $path = copySourceForOpen "$TestDrive\copy1\議事メモ.txt" "議事メモ.txt" $copyDir (getTmpRoot "copy1") $script:when

        $path | Should -BeLike "$copyDir\*\議事メモ（コピー 20261012-1530）.txt"
        [System.IO.File]::ReadAllText($path) | Should -Be "本文あり"
        $after = Get-Item -LiteralPath "$TestDrive\copy1\議事メモ.txt"
        $after.LastWriteTimeUtc | Should -Be $writeTime
        $after.Attributes | Should -Be $attributes
        [System.IO.File]::ReadAllText("$TestDrive\copy1\議事メモ.txt") | Should -Be "本文あり"
    }

    It "元のファイルをほかのアプリが開いたまま（書き込み中）でも、コピーを作れる" {
        $copyDir = newCopyDir "copy1b"
        [System.IO.File]::WriteAllText("$TestDrive\copy1b\a.txt", "使用中")
        $stream = [System.IO.File]::Open("$TestDrive\copy1b\a.txt", [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::ReadWrite)
        try {
            $path = copySourceForOpen "$TestDrive\copy1b\a.txt" "a.txt" $copyDir (getTmpRoot "copy1b")
        } finally {
            $stream.Dispose()
        }

        $path | Should -Not -Be ""
        [System.IO.File]::ReadAllText($path) | Should -Be "使用中"
    }

    It "元のファイルが読み取り専用でも、コピーは通常の属性で作る" {
        $copyDir = newCopyDir "copy2"
        [System.IO.File]::WriteAllText("$TestDrive\copy2\a.txt", "x")
        [System.IO.File]::SetAttributes("$TestDrive\copy2\a.txt", [System.IO.FileAttributes]::ReadOnly)

        $path = copySourceForOpen "$TestDrive\copy2\a.txt" "a.txt" $copyDir (getTmpRoot "copy2")

        ([System.IO.File]::GetAttributes($path) -band [System.IO.FileAttributes]::ReadOnly) | Should -Be 0
        [System.IO.File]::SetAttributes("$TestDrive\copy2\a.txt", [System.IO.FileAttributes]::Normal)
    }

    It "同じファイルを続けて開いても、別のフォルダに作る（名前が重ならない）" {
        $copyDir = newCopyDir "copy3"
        [System.IO.File]::WriteAllText("$TestDrive\copy3\a.txt", "x")

        $first = copySourceForOpen "$TestDrive\copy3\a.txt" "a.txt" $copyDir (getTmpRoot "copy3")
        $second = copySourceForOpen "$TestDrive\copy3\a.txt" "a.txt" $copyDir (getTmpRoot "copy3")

        $first | Should -Not -Be $second
        Test-Path -LiteralPath $first | Should -Be $true
        Test-Path -LiteralPath $second | Should -Be $true
    }

    It "元のファイルが無ければ空文字を返す" {
        $copyDir = newCopyDir "copy4"
        copySourceForOpen "$TestDrive\copy4\無い.txt" "無い.txt" $copyDir (getTmpRoot "copy4") | Should -Be ""
    }

    It "置き場所が open_copy でない・空・相対パスのときは、何も作らず空文字を返す（<name>）" -TestCases @(
        @{ name = "別の名前"; sub = "copy5\other" }
        @{ name = "空"; sub = "" }
    ) {
        param ($name, $sub)
        [System.IO.Directory]::CreateDirectory("$TestDrive\copy5") | Out-Null
        [System.IO.File]::WriteAllText("$TestDrive\copy5\a.txt", "x")
        $dir = if ($sub) { "$TestDrive\$sub" } else { "" }

        copySourceForOpen "$TestDrive\copy5\a.txt" "a.txt" $dir (getTmpRoot "copy5") | Should -Be ""
        Test-Path -LiteralPath "$TestDrive\copy5\other" | Should -Be $false
    }

    It "名前が空・[ ] を含む・パスが長すぎるときは、コピーを作らず空文字を返す" {
        $copyDir = newCopyDir "copy6"
        [System.IO.File]::WriteAllText("$TestDrive\copy6\a.txt", "x")
        copySourceForOpen "$TestDrive\copy6\a.txt" "" $copyDir (getTmpRoot "copy6") | Should -Be ""
        copySourceForOpen "$TestDrive\copy6\a.txt" "報告[1].txt" $copyDir (getTmpRoot "copy6") | Should -Be ""
        copySourceForOpen "$TestDrive\copy6\a.txt" ("あ" * 250 + ".txt") $copyDir (getTmpRoot "copy6") | Should -Be ""
        @(Get-ChildItem -LiteralPath $copyDir -Force).Count | Should -Be 0
    }
}

Describe "clearOldOpenCopies" -Tag Io {
    BeforeAll {
        $script:now = [datetime]"2026-10-12 12:00:00"
    }

    It "1 日より古いコピーだけを消し、新しいコピーは残す" {
        $copyDir = newCopyDir "clear1"
        $old = newOldCopy $copyDir "old1" "a.txt" $script:now.AddDays(-2)
        $new = newOldCopy $copyDir "new1" "b.txt" $script:now.AddHours(-3)

        clearOldOpenCopies $copyDir (getTmpRoot "clear1") $script:now | Should -Be 1

        Test-Path -LiteralPath $old | Should -Be $false
        Test-Path -LiteralPath "$new\b.txt" | Should -Be $true
        Test-Path -LiteralPath $copyDir | Should -Be $true
    }

    It "フォルダは古くても、中のファイルを最近保存したコピーは残す" {
        $copyDir = newCopyDir "clear1b"
        $edited = newOldCopy $copyDir "edited" "a.txt" $script:now.AddDays(-5)
        (Get-Item -LiteralPath "$edited\a.txt").LastWriteTime = $script:now.AddHours(-1)
        (Get-Item -LiteralPath $edited).LastWriteTime = $script:now.AddDays(-5)

        clearOldOpenCopies $copyDir (getTmpRoot "clear1b") $script:now | Should -Be 0
        Test-Path -LiteralPath "$edited\a.txt" | Should -Be $true
    }

    It "現在より先の日時（時計のずれ）のコピーは、新しい扱いで残す" {
        $copyDir = newCopyDir "clear1c"
        $future = newOldCopy $copyDir "future" "a.txt" $script:now.AddDays(3)

        clearOldOpenCopies $copyDir (getTmpRoot "clear1c") $script:now | Should -Be 0
        Test-Path -LiteralPath "$future\a.txt" | Should -Be $true
    }

    It "読み取り専用のコピーも消す" {
        $copyDir = newCopyDir "clear2"
        $old = newOldCopy $copyDir "old1" "a.txt" $script:now.AddDays(-2)
        [System.IO.File]::SetAttributes("$old\a.txt", [System.IO.FileAttributes]::ReadOnly)

        clearOldOpenCopies $copyDir (getTmpRoot "clear2") $script:now | Should -Be 1
        Test-Path -LiteralPath $old | Should -Be $false
    }

    It "開いている（消せない）コピーは残す" {
        $copyDir = newCopyDir "clear3"
        $old = newOldCopy $copyDir "old1" "a.txt" $script:now.AddDays(-2)
        $other = newOldCopy $copyDir "old2" "b.txt" $script:now.AddDays(-2)
        $stream = [System.IO.File]::Open("$old\a.txt", [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        try {
            clearOldOpenCopies $copyDir (getTmpRoot "clear3") $script:now | Should -Be 1
        } finally {
            $stream.Dispose()
        }

        Test-Path -LiteralPath "$old\a.txt" | Should -Be $true
        Test-Path -LiteralPath $other | Should -Be $false
    }

    It "フォルダが無ければ何もしない" {
        clearOldOpenCopies "$(getTmpRoot "clear4")\$(getMachineKey)\open_copy" (getTmpRoot "clear4") $script:now | Should -Be 0
    }

    It "open_copy の中の、ファイルでないもの（下のフォルダ）は消さず、直下のファイルも触らない" {
        $copyDir = newCopyDir "clear5"
        $old = newOldCopy $copyDir "old1" "a.txt" $script:now.AddDays(-2)
        [System.IO.Directory]::CreateDirectory("$old\下") | Out-Null
        [System.IO.File]::WriteAllText("$copyDir\直下.txt", "x")

        clearOldOpenCopies $copyDir (getTmpRoot "clear5") $script:now | Should -Be 0

        Test-Path -LiteralPath "$old\下" | Should -Be $true
        Test-Path -LiteralPath "$copyDir\直下.txt" | Should -Be $true
    }

    It "open_copy の外・空・ルート・相対パスを渡されたら、何も消さない（<name>）" -TestCases @(
        @{ name = "名前が open_copy でない場所"; sub = "clear6\tmp\key" }
        @{ name = "ワークスペースの tmp"; sub = "clear6\tmp" }
        @{ name = "別の場所の open_copy"; sub = "clear6\other\open_copy" }
        @{ name = "空"; sub = "" }
        @{ name = "ドライブの直下"; sub = "<root>" }
        @{ name = "相対パス"; sub = "<relative>" }
    ) {
        param ($name, $sub)
        $keep = newOldCopy "$TestDrive\clear6\tmp\key\open_copy" "old1" "a.txt" $script:now.AddDays(-2)
        $keepOther = newOldCopy "$TestDrive\clear6\other\open_copy" "old1" "a.txt" $script:now.AddDays(-2)
        $dir = if ($sub -eq "<root>") { [System.IO.Path]::GetPathRoot($TestDrive) } elseif ($sub -eq "<relative>") { "clear6\tmp\key\open_copy" } elseif ($sub) { "$TestDrive\$sub" } else { "" }

        clearOldOpenCopies $dir "$TestDrive\clear6\tmp" $script:now | Should -Be 0

        Test-Path -LiteralPath "$keep\a.txt" | Should -Be $true
        Test-Path -LiteralPath "$keepOther\a.txt" | Should -Be $true
    }

    It "tmp の指定と違う場所の open_copy は、名前が合っていても消さない" {
        $copyDir = newCopyDir "clear7"
        $old = newOldCopy $copyDir "old1" "a.txt" $script:now.AddDays(-2)

        clearOldOpenCopies $copyDir "$TestDrive\clear7\別の tmp" $script:now | Should -Be 0
        Test-Path -LiteralPath "$old\a.txt" | Should -Be $true
    }

    It "リンク（ジャンクション）を辿らず、指す先のファイルを消さない（<name>）" -TestCases @(
        @{ name = "open_copy 自身"; level = 0 }
        @{ name = "PC の鍵"; level = 1 }
        @{ name = "tmp"; level = 2 }
        @{ name = "open_copy の 1 つ下"; level = 3 }
    ) {
        param ($name, $level)
        $base = "$TestDrive\clearlink$level"
        $key = getMachineKey
        # 指す先は、$TestDrive の下の「古いコピーが入った実体」
        $real = "$base\real"
        $oldInReal = newOldCopy "$real\$key\open_copy" "old1" "a.txt" $script:now.AddDays(-2)
        switch ($level) {
            0 {
                [System.IO.Directory]::CreateDirectory("$base\tmp\$key") | Out-Null
                $made = newJunction "$base\tmp\$key\open_copy" "$real\$key\open_copy"
                $target = "$oldInReal\a.txt"
            }
            1 {
                [System.IO.Directory]::CreateDirectory("$base\tmp") | Out-Null
                $made = newJunction "$base\tmp\$key" "$real\$key"
                $target = "$oldInReal\a.txt"
            }
            2 {
                $made = newJunction "$base\tmp" $real
                $target = "$oldInReal\a.txt"
            }
            3 {
                [System.IO.Directory]::CreateDirectory("$base\tmp\$key\open_copy") | Out-Null
                $outside = "$base\outside"
                [System.IO.Directory]::CreateDirectory($outside) | Out-Null
                [System.IO.File]::WriteAllText("$outside\a.txt", "外")
                (Get-Item -LiteralPath "$outside\a.txt").LastWriteTime = $script:now.AddDays(-9)
                (Get-Item -LiteralPath $outside).CreationTime = $script:now.AddDays(-9)
                $made = newJunction "$base\tmp\$key\open_copy\old1" $outside
                $target = "$outside\a.txt"
            }
        }
        if (!$made) {
            Set-ItResult -Skipped -Because "この環境ではジャンクションを作れない"
            return
        }

        clearOldOpenCopies "$base\tmp\$key\open_copy" "$base\tmp" $script:now | Should -Be 0

        Test-Path -LiteralPath $target | Should -Be $true
    }
}
