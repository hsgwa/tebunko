# ［新規で開く］のためのコピーの作成と片付け（tebunko\core\open_copy.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:newCopyDir([string]$base) {
        $dir = "$TestDrive\$base\tmp\key\open_copy"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        return $dir
    }

    function script:newOldCopy([string]$copyDir, [string]$sub, [string]$name, [datetime]$created) {
        # 作ったのが古いコピー（1 つ下のフォルダごと）を置く
        $dir = Join-Path $copyDir $sub
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $dir $name), "copy")
        (Get-Item -LiteralPath $dir).CreationTime = $created
        return $dir
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
    It "<name> は <expected>" -TestCases @(
        @{ name = "open_copy"; dir = "C:\work\tmp\key\open_copy"; expected = $true }
        @{ name = "末尾の区切りつき"; dir = "C:\work\tmp\key\open_copy\"; expected = $true }
        @{ name = "空"; dir = ""; expected = $false }
        @{ name = "空白"; dir = "   "; expected = $false }
        @{ name = "ドライブの直下"; dir = "C:\"; expected = $false }
        @{ name = "ドライブだけ"; dir = "C:"; expected = $false }
        @{ name = "ルート"; dir = "\"; expected = $false }
        @{ name = "open_copy でない名前"; dir = "C:\work\tmp\key"; expected = $false }
        @{ name = "名前が似ているだけ"; dir = "C:\work\tmp\key\open_copy2"; expected = $false }
        @{ name = "元のフォルダ"; dir = "C:\data\営業"; expected = $false }
    ) {
        param ($name, $dir, $expected)
        testOpenCopyDir $dir | Should -Be $expected
    }
}

Describe "getOpenCopyName" -Tag Unit {
    It "<book> の名前は <expected>" -TestCases @(
        @{ book = "仕様書.pdf"; expected = "仕様書.pdf" }
        @{ book = "a.txt. "; expected = "a.txt" }
        @{ book = "x\..\..\悪い.txt"; expected = "悪い.txt" }
        @{ book = "x/悪い.txt"; expected = "悪い.txt" }
        @{ book = ""; expected = "" }
    ) {
        param ($book, $expected)
        getOpenCopyName $book | Should -Be $expected
    }
}

Describe "copySourceForOpen" -Tag Io {
    It "元のファイルと同じ名前・中身のコピーを open_copy の下に作り、元のファイルは変えない" {
        $copyDir = newCopyDir "copy1"
        [System.IO.File]::WriteAllText("$TestDrive\copy1\議事メモ.txt", "本文あり")

        $path = copySourceForOpen "$TestDrive\copy1\議事メモ.txt" "議事メモ.txt" $copyDir

        $path | Should -BeLike "$copyDir\*\議事メモ.txt"
        [System.IO.File]::ReadAllText($path) | Should -Be "本文あり"
        [System.IO.File]::ReadAllText("$TestDrive\copy1\議事メモ.txt") | Should -Be "本文あり"
    }

    It "元のファイルが読み取り専用でも、コピーは通常の属性で作る" {
        $copyDir = newCopyDir "copy2"
        [System.IO.File]::WriteAllText("$TestDrive\copy2\a.txt", "x")
        [System.IO.File]::SetAttributes("$TestDrive\copy2\a.txt", [System.IO.FileAttributes]::ReadOnly)

        $path = copySourceForOpen "$TestDrive\copy2\a.txt" "a.txt" $copyDir

        ([System.IO.File]::GetAttributes($path) -band [System.IO.FileAttributes]::ReadOnly) | Should -Be 0
        [System.IO.File]::SetAttributes("$TestDrive\copy2\a.txt", [System.IO.FileAttributes]::Normal)
    }

    It "同じファイルを続けて開いても、別のフォルダに作る（名前が重ならない）" {
        $copyDir = newCopyDir "copy3"
        [System.IO.File]::WriteAllText("$TestDrive\copy3\a.txt", "x")

        $first = copySourceForOpen "$TestDrive\copy3\a.txt" "a.txt" $copyDir
        $second = copySourceForOpen "$TestDrive\copy3\a.txt" "a.txt" $copyDir

        $first | Should -Not -Be $second
        Test-Path -LiteralPath $first | Should -Be $true
        Test-Path -LiteralPath $second | Should -Be $true
    }

    It "元のファイルが無ければ空文字を返す" {
        $copyDir = newCopyDir "copy4"
        copySourceForOpen "$TestDrive\copy4\無い.txt" "無い.txt" $copyDir | Should -Be ""
    }

    It "置き場所が open_copy でない・空のときは、何も作らず空文字を返す（<name>）" -TestCases @(
        @{ name = "別の名前"; sub = "copy5\other" }
        @{ name = "空"; sub = "" }
    ) {
        param ($name, $sub)
        [System.IO.Directory]::CreateDirectory("$TestDrive\copy5") | Out-Null
        [System.IO.File]::WriteAllText("$TestDrive\copy5\a.txt", "x")
        $dir = if ($sub) { "$TestDrive\$sub" } else { "" }

        copySourceForOpen "$TestDrive\copy5\a.txt" "a.txt" $dir | Should -Be ""
        Test-Path -LiteralPath "$TestDrive\copy5\other" | Should -Be $false
    }

    It "名前が空・パスが長すぎるときは、空文字を返す" {
        $copyDir = newCopyDir "copy6"
        [System.IO.File]::WriteAllText("$TestDrive\copy6\a.txt", "x")
        copySourceForOpen "$TestDrive\copy6\a.txt" "" $copyDir | Should -Be ""
        copySourceForOpen "$TestDrive\copy6\a.txt" ("あ" * 250 + ".txt") $copyDir | Should -Be ""
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

        clearOldOpenCopies $copyDir $script:now | Should -Be 1

        Test-Path -LiteralPath $old | Should -Be $false
        Test-Path -LiteralPath "$new\b.txt" | Should -Be $true
        Test-Path -LiteralPath $copyDir | Should -Be $true
    }

    It "読み取り専用のコピーも消す" {
        $copyDir = newCopyDir "clear2"
        $old = newOldCopy $copyDir "old1" "a.txt" $script:now.AddDays(-2)
        [System.IO.File]::SetAttributes("$old\a.txt", [System.IO.FileAttributes]::ReadOnly)

        clearOldOpenCopies $copyDir $script:now | Should -Be 1
        Test-Path -LiteralPath $old | Should -Be $false
    }

    It "開いている（消せない）コピーは残す" {
        $copyDir = newCopyDir "clear3"
        $old = newOldCopy $copyDir "old1" "a.txt" $script:now.AddDays(-2)
        $other = newOldCopy $copyDir "old2" "b.txt" $script:now.AddDays(-2)
        $stream = [System.IO.File]::Open("$old\a.txt", [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        try {
            clearOldOpenCopies $copyDir $script:now | Should -Be 1
        } finally {
            $stream.Dispose()
        }

        Test-Path -LiteralPath "$old\a.txt" | Should -Be $true
        Test-Path -LiteralPath $other | Should -Be $false
    }

    It "フォルダが無ければ何もしない" {
        clearOldOpenCopies "$TestDrive\clear4\tmp\key\open_copy" $script:now | Should -Be 0
    }

    It "open_copy の中の、ファイルでないもの（下のフォルダ）は消さず、直下のファイルも触らない" {
        $copyDir = newCopyDir "clear5"
        $old = newOldCopy $copyDir "old1" "a.txt" $script:now.AddDays(-2)
        [System.IO.Directory]::CreateDirectory("$old\下") | Out-Null
        [System.IO.File]::WriteAllText("$copyDir\直下.txt", "x")

        clearOldOpenCopies $copyDir $script:now | Should -Be 0

        Test-Path -LiteralPath "$old\下" | Should -Be $true
        Test-Path -LiteralPath "$copyDir\直下.txt" | Should -Be $true
    }

    It "open_copy の外・空・ルートを渡されたら、何も消さない（<name>）" -TestCases @(
        @{ name = "名前が open_copy でない場所"; sub = "clear6\tmp\key" }
        @{ name = "ワークスペースの tmp"; sub = "clear6\tmp" }
        @{ name = "空"; sub = "" }
        @{ name = "ドライブの直下"; sub = "<root>" }
    ) {
        param ($name, $sub)
        $keep = newOldCopy "$TestDrive\clear6\tmp\key" "old1" "a.txt" $script:now.AddDays(-2)
        $dir = if ($sub -eq "<root>") { [System.IO.Path]::GetPathRoot($TestDrive) } elseif ($sub) { "$TestDrive\$sub" } else { "" }

        clearOldOpenCopies $dir $script:now | Should -Be 0

        Test-Path -LiteralPath "$keep\a.txt" | Should -Be $true
    }
}
