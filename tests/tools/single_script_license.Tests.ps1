# tools\single_script_license.ps1（単一 .ps1 に同梱するライセンス文を、# のコメントにして入れ・元に戻す道具）のテスト。
# 作る道具の通し（本物の LICENSE-*.txt が入るか）は tests\tools\new_single_script.Tests.ps1。設計は docs/design/structure/single-script.md
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "${scriptsDir}\..\tools\single_script_license.ps1"

    function writeLicenseFile {
        param ([string]$dir, [string]$name, [string]$text)
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $dir $name), $text, (New-Object System.Text.UTF8Encoding($false)))
    }
}

Describe "single_script_license.ps1 の改行のそろえ方（比べるときの元の形）" -Tag Unit {
    It "<name>: 手で書いた期待値と同じになる" -TestCases @(
        @{ name = "LF"; text = "a`nb`n"; expected = "a`r`nb`r`n" }
        @{ name = "CRLF"; text = "a`r`nb`r`n"; expected = "a`r`nb`r`n" }
        @{ name = "LF と CRLF が混ざる"; text = "a`nb`r`nc`n"; expected = "a`r`nb`r`nc`r`n" }
        @{ name = "末尾の改行なし"; text = "a`nb"; expected = "a`r`nb`r`n" }
        @{ name = "末尾の空行が 2 つ"; text = "a`n`n`n"; expected = "a`r`n`r`n`r`n" }
        @{ name = "行末の空白"; text = "a  `nb`t`n"; expected = "a  `r`nb`t`r`n" }
    ) {
        param ($name, $text, $expected)
        (normalizeLicenseText $text) | Should -BeExactly $expected
    }
}

Describe "single_script_license.ps1 のコメントへの変換と、元への復元" -Tag Unit {
    It "<name>: コメントにして取り出すと、元の文字列（改行は CRLF・末尾は改行）と同じになる" -TestCases @(
        @{ name = "LF"; text = "ISC License`nline 2`nline 3`n" }
        @{ name = "CRLF"; text = "ISC License`r`nline 2`r`nline 3`r`n" }
        @{ name = "最後に改行が無い"; text = "ISC License`nline 2" }
        @{ name = "途中の空行"; text = "ISC License`n`nline 3`n" }
        @{ name = "最後に空行が続く"; text = "ISC License`n`n" }
        @{ name = "空白だけの行"; text = "ISC License`n   `nline 3`n" }
        @{ name = "タブの字下げ"; text = "ISC License`n`tindented`nline 3`n" }
        @{ name = "行末の空白"; text = "ISC License  `nline 2`t`n" }
        @{ name = "# で始まる行"; text = "# ISC License`n#`n##x`n" }
        @{ name = "日本語"; text = "ライセンス`n本文`n" }
    ) {
        param ($name, $text)
        $comment = convertToLicenseComment "LICENSE-X.txt" $text
        (restoreLicenseFromComment $comment "LICENSE-X.txt") | Should -BeExactly (normalizeLicenseText $text)
    }

    It "行の形: 空行は # だけ、それ以外は # と 1 つの空白を足すだけ。ファイル名の印で挟む" {
        $comment = convertToLicenseComment "LICENSE-X.txt" "a`r`n`r`n  b`r`n"
        $comment | Should -BeExactly ((@(
                    "# ---- 第三者の部品のライセンス文: LICENSE-X.txt（ここから）----"
                    "# a"
                    "#"
                    "#   b"
                    "# ---- 第三者の部品のライセンス文: LICENSE-X.txt（ここまで）----") -join "`r`n"))
    }

    It "複数のライセンス文をつないでも、それぞれ元の文字列で取り出せ、順に並ぶ" {
        $first = "ISC License`nfirst text`n"
        $second = "MIT License`r`nsecond text`r`n`r`nend"
        $script = (@(
                (convertToLicenseComment "LICENSE-A.txt" $first)
                (convertToLicenseComment "LICENSE-B.txt" $second)
                "param ()") -join "`r`n") + "`r`n"
        (restoreLicenseFromComment $script "LICENSE-A.txt") | Should -BeExactly (normalizeLicenseText $first)
        (restoreLicenseFromComment $script "LICENSE-B.txt") | Should -BeExactly (normalizeLicenseText $second)
        $script.IndexOf("LICENSE-A.txt（ここから）") | Should -BeLessThan $script.IndexOf("LICENSE-B.txt（ここから）")
    }

    It "<name>: 空（空白だけ）、または改行を伴わない CR だけの行区切りがあるライセンス文は止める" -TestCases @(
        @{ name = "空文字列"; text = ""; message = "*ライセンス文が空です*" }
        @{ name = "改行だけ"; text = "`r`n`r`n"; message = "*ライセンス文が空です*" }
        @{ name = "空白・タブ・改行だけ"; text = " `t`r`n  `n"; message = "*ライセンス文が空です*" }
        @{ name = "CR だけの区切り"; text = "ISC License`rsecond line`n"; message = "*CR だけの行区切り*" }
        @{ name = "末尾が CR だけ"; text = "ISC License`r"; message = "*CR だけの行区切り*" }
    ) {
        param ($name, $text, $message)
        { convertToLicenseComment "LICENSE-X.txt" $text } | Should -Throw $message
    }

    It "元に空行がある文面で、U+00AD と # だけの行が混ざっても、空行として戻さず `$null を返す" {
        $comment = convertToLicenseComment "LICENSE-X.txt" "ISC License`n`nline 3`n"
        $broken = $comment.Replace("`r`n#`r`n", "`r`n$([char]0x00AD)#`r`n")
        [string]::Equals($broken, $comment, [System.StringComparison]::Ordinal) | Should -Be $false
        (restoreLicenseFromComment $broken "LICENSE-X.txt") | Should -BeNullOrEmpty
    }

    It "<name>: 印が無い・終わりの印が無い・コメントでない行があるときは、`$null を返す（道具の検査が止まる形になる）" -TestCases @(
        @{ name = "# が外れた行"; edit = { param ($c) $c.Replace("# line 2", "line 2") } }
        @{ name = "印の間の空行"; edit = { param ($c) $c.Replace("# line 2", "") } }
        @{ name = "# だけでなく 1 文字の行"; edit = { param ($c) $c.Replace("# line 2", "x") } }
        @{ name = "頭の 2 文字が # でないが、外すと元と同じになる行"; edit = { param ($c) $c.Replace("# line 2", "xxline 2") } }
        @{ name = "見えない文字（U+00AD）で始まる行"; edit = { param ($c) $c.Replace("# line 2", "$([char]0x00AD)# line 2") } }
        @{ name = "印が無い"; edit = { param ($c) "ISC License" } }
        @{ name = "終わりの印が無い"; edit = { param ($c) ($c -split "\r?\n" | Select-Object -SkipLast 1) -join "`r`n" } }
    ) {
        param ($name, $edit)
        $original = "ISC License`nline 2`nline 3`n"
        $broken = & $edit (convertToLicenseComment "LICENSE-X.txt" $original)
        (restoreLicenseFromComment $broken "LICENSE-X.txt") | Should -BeNullOrEmpty
    }

    It "<name>: 行が欠けた・書き換わったときは、取り出した中身が元の文字列と違う" -TestCases @(
        @{ name = "1 行欠けた"; edit = { param ($c) ($c -split "\r?\n" | Where-Object { $_ -ne "# line 2" }) -join "`r`n" } }
        @{ name = "1 行が書き換わった"; edit = { param ($c) $c.Replace("# line 2", "# line 2 changed") } }
    ) {
        param ($name, $edit)
        $original = "ISC License`nline 2`nline 3`n"
        $broken = & $edit (convertToLicenseComment "LICENSE-X.txt" $original)
        (restoreLicenseFromComment $broken "LICENSE-X.txt") | Should -Not -BeExactly (normalizeLicenseText $original)
    }
}

Describe "single_script_license.ps1 のライセンス文のファイルの一覧" -Tag Unit {
    It "複数あるときは、名前の文字コード順（言語・地域に依らない）に並べ、LICENSE-*.txt 以外は入れない" {
        $dir = Join-Path $TestDrive "order"
        foreach ($name in @("LICENSE-b.txt", "LICENSE-Z.txt", "LICENSE-a.txt", "LICENSE-A2.txt", "README.txt", "LICENSE-x.md")) {
            writeLicenseFile $dir $name "text of $name`n"
        }
        # 呼ぶ側（new_single_script.ps1）と同じ受け方。配列が入れ子にならず、1 件ずつの FileInfo になる
        $files = @(getLicenseFiles $dir)
        $files.Count | Should -Be 4
        $files[0] | Should -BeOfType ([System.IO.FileInfo])
        @($files | ForEach-Object { $_.Name }) | Should -Be @("LICENSE-A2.txt", "LICENSE-Z.txt", "LICENSE-a.txt", "LICENSE-b.txt")
    }

    It "1 件のときも、配列として 1 件で受けられる" {
        $dir = Join-Path $TestDrive "single"
        writeLicenseFile $dir "LICENSE-a.txt" "text`n"
        $files = @(getLicenseFiles $dir)
        $files.Count | Should -Be 1
        $files[0] | Should -BeOfType ([System.IO.FileInfo])
        $files[0].Name | Should -BeExactly "LICENSE-a.txt"
    }

    It "<name>: 止める" -TestCases @(
        @{ name = "フォルダが無い"; setup = { param ($d) }; message = "*フォルダがありません*" }
        @{ name = "LICENSE-*.txt が 0 件"; setup = { param ($d) writeLicenseFile $d "README.txt" "x" }; message = "*ライセンス文がありません*" }
        @{ name = "空のファイル"; setup = { param ($d) writeLicenseFile $d "LICENSE-a.txt" "text`n"; writeLicenseFile $d "LICENSE-b.txt" "" }; message = "*ライセンス文が空です: LICENSE-b.txt*" }
        @{ name = "空白だけのファイル"; setup = { param ($d) writeLicenseFile $d "LICENSE-a.txt" " `t`r`n`r`n" }; message = "*ライセンス文が空です: LICENSE-a.txt*" }
        @{ name = "CR だけの行区切りがあるファイル"; setup = { param ($d) writeLicenseFile $d "LICENSE-a.txt" "line 1`rline 2`n" }; message = "*CR だけの行区切り*" }
    ) {
        param ($name, $setup, $message)
        $dir = Join-Path $TestDrive ("stop-" + [guid]::NewGuid().ToString("N"))
        & $setup $dir
        { getLicenseFiles $dir } | Should -Throw $message
    }
}
