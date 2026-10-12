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

    It "<name>: 空（空白だけ）のライセンス文は止める" -TestCases @(
        @{ name = "空文字列"; text = "" }
        @{ name = "改行だけ"; text = "`r`n`r`n" }
        @{ name = "空白・タブ・改行だけ"; text = " `t`r`n  `n" }
    ) {
        param ($name, $text)
        { convertToLicenseComment "LICENSE-X.txt" $text } | Should -Throw "*ライセンス文が空です*"
    }

    It "<name>: 印が無い・欠けている・中身が違うときは、元の文字列と一致しない（道具の検査が止まる形になる）" -TestCases @(
        @{ name = "印が無い"; edit = { param ($c) "ISC License" } }
        @{ name = "終わりの印が無い"; edit = { param ($c) ($c -split "\r?\n" | Select-Object -SkipLast 1) -join "`r`n" } }
        @{ name = "1 行欠けた"; edit = { param ($c) ($c -split "\r?\n" | Where-Object { $_ -ne "# line 2" }) -join "`r`n" } }
        @{ name = "1 行が書き換わった"; edit = { param ($c) $c.Replace("# line 2", "# line 2 changed") } }
    ) {
        param ($name, $edit)
        $original = "ISC License`nline 2`nline 3`n"
        $broken = & $edit (convertToLicenseComment "LICENSE-X.txt" $original)
        $restored = restoreLicenseFromComment $broken "LICENSE-X.txt"
        $restored | Should -Not -BeExactly (normalizeLicenseText $original)
    }
}

Describe "single_script_license.ps1 のライセンス文のファイルの一覧" -Tag Unit {
    It "複数あるときは、名前の文字コード順（言語・地域に依らない）に並べ、LICENSE-*.txt 以外は入れない" {
        $dir = Join-Path $TestDrive "order"
        foreach ($name in @("LICENSE-b.txt", "LICENSE-Z.txt", "LICENSE-a.txt", "LICENSE-A2.txt", "README.txt", "LICENSE-x.md")) {
            writeLicenseFile $dir $name "text of $name`n"
        }
        $names = @(getLicenseFiles $dir | ForEach-Object { $_.Name })
        $names | Should -Be @("LICENSE-A2.txt", "LICENSE-Z.txt", "LICENSE-a.txt", "LICENSE-b.txt")
    }

    It "<name>: 止める" -TestCases @(
        @{ name = "フォルダが無い"; setup = { param ($d) }; message = "*フォルダがありません*" }
        @{ name = "LICENSE-*.txt が 0 件"; setup = { param ($d) writeLicenseFile $d "README.txt" "x" }; message = "*ライセンス文がありません*" }
        @{ name = "空のファイル"; setup = { param ($d) writeLicenseFile $d "LICENSE-a.txt" "text`n"; writeLicenseFile $d "LICENSE-b.txt" "" }; message = "*ファイルが空です*" }
        @{ name = "空白だけのファイル"; setup = { param ($d) writeLicenseFile $d "LICENSE-a.txt" " `t`r`n`r`n" }; message = "*ファイルが空です*" }
    ) {
        param ($name, $setup, $message)
        $dir = Join-Path $TestDrive ("stop-" + [guid]::NewGuid().ToString("N"))
        & $setup $dir
        { getLicenseFiles $dir } | Should -Throw $message
    }
}
