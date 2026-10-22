# CSV/TSV の書き出し（shared\core\csv_writer.ps1）のテスト
BeforeDiscovery {
    $lineCases = @(
        @{ name = "区切りも引用符も無いセルは囲まない"; cells = @("a", "b"); delimiter = ","; expected = "a,b" }
        @{ name = "コンマを含むセルだけ囲む"; cells = @("a,b", "c"); delimiter = ","; expected = '"a,b",c' }
        @{ name = "引用符を含むセルは囲み、中を二重にする"; cells = @('a"b'); delimiter = ","; expected = '"a""b"' }
        @{ name = "CR を含むセルは囲む"; cells = @("a`rb"); delimiter = ","; expected = "`"a`rb`"" }
        @{ name = "LF を含むセルは囲む"; cells = @("a`nb"); delimiter = ","; expected = "`"a`nb`"" }
        @{ name = "空のセルは空"; cells = @("", "x", ""); delimiter = ","; expected = ",x," }
        @{ name = "`$null のセルは空"; cells = @($null, "x"); delimiter = ","; expected = ",x" }
        @{ name = "先頭が = でも加工しない"; cells = @("=1+1", "@a"); delimiter = ","; expected = "=1+1,@a" }
        @{ name = "セルが 0 個なら空文字"; cells = @(); delimiter = ","; expected = "" }
        @{ name = "TSV はタブを含むセルを囲む"; cells = @("a`tb", "c"); delimiter = "`t"; expected = "`"a`tb`"`tc" }
        @{ name = "TSV はコンマだけのセルを囲まない"; cells = @("a,b", "c"); delimiter = "`t"; expected = "a,b`tc" }
        @{ name = "TSV でも引用符・改行は囲む"; cells = @('a"b', "x`ny"); delimiter = "`t"; expected = "`"a`"`"b`"`t`"x`ny`"" }
    )
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function writeBody {
        param ([string]$path, [object[]]$rows, [string]$delimiter)
        $text = ""
        foreach ($row in $rows) {
            $text += (toCsvLine $row $delimiter) + "`r`n"
        }
        [System.IO.File]::WriteAllBytes($path, (New-Object System.Text.UTF8Encoding($false)).GetBytes($text))
    }
}

Describe "toCsvLine" -Tag Unit {
    It "<name>" -TestCases $lineCases {
        toCsvLine $cells $delimiter | Should -BeExactly $expected
    }
}

Describe "writeCsvFile" -Tag Io {
    BeforeEach {
        $script:body = Join-Path $TestDrive "body.tmp"
        $script:out = Join-Path $TestDrive "out.csv"
        Remove-Item -LiteralPath $script:out -Force -ErrorAction SilentlyContinue
        $script:headers = @("名前", "メモ", "値")
        $script:rows = @(
            @("山田", "1行目`r`n2行目", "a,b"),
            @("佐藤", 'say "hi"', "x`ty"),
            @("", "", "=1+1")
        )
    }

    It "<label> BOM 付き UTF-8・CRLF で、読み戻した値が元と同じ" -TestCases @(
        @{ label = "コンマ"; delimiter = "," }
        @{ label = "タブ"; delimiter = "`t" }
    ) {
        writeBody $script:body $script:rows $delimiter
        writeCsvFile $script:body $script:out $script:headers $delimiter

        $bytes = [System.IO.File]::ReadAllBytes($script:out)
        $bytes[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
        $text = [System.IO.File]::ReadAllText($script:out, (New-Object System.Text.UTF8Encoding($true)))
        $text.StartsWith("$(toCsvLine $script:headers $delimiter)`r`n") | Should -BeTrue
        $text.EndsWith("`r`n") | Should -BeTrue
        $text.EndsWith("`r`n`r`n") | Should -BeFalse
        ([regex]::Matches($text, "(?<!`r)`n")).Count | Should -Be 0

        $read = @(Import-Csv -LiteralPath $script:out -Delimiter $delimiter -Encoding UTF8)
        $read.Count | Should -Be 3
        for ($i = 0; $i -lt 3; $i++) {
            $read[$i].名前 | Should -BeExactly $script:rows[$i][0]
            $read[$i].メモ | Should -BeExactly $script:rows[$i][1]
            $read[$i].値 | Should -BeExactly $script:rows[$i][2]
        }
    }

    It "本文が空でも見出し行だけのファイルができる" {
        [System.IO.File]::WriteAllBytes($script:body, [byte[]]@())
        writeCsvFile $script:body $script:out $script:headers
        $text = [System.IO.File]::ReadAllText($script:out, (New-Object System.Text.UTF8Encoding($true)))
        $text | Should -BeExactly "名前,メモ,値`r`n"
    }

    It "保存先が共有なしで開かれていると IOException になる" {
        writeBody $script:body $script:rows ","
        $lock = New-Object System.IO.FileStream($script:out, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $err = try { writeCsvFile $script:body $script:out $script:headers } catch { $_.Exception }
            # PowerShell は .NET メソッドの例外を包むので、元の例外は InnerException にある
            $err.InnerException | Should -BeOfType ([System.IO.IOException])
        } finally {
            $lock.Dispose()
        }
    }

    It "読み取り専用ファイルには UnauthorizedAccessException になる" {
        writeBody $script:body $script:rows ","
        [System.IO.File]::WriteAllText($script:out, "x")
        Set-ItemProperty -LiteralPath $script:out -Name IsReadOnly -Value $true
        try {
            $err = try { writeCsvFile $script:body $script:out $script:headers } catch { $_.Exception }
            $err.InnerException | Should -BeOfType ([System.UnauthorizedAccessException])
        } finally {
            Set-ItemProperty -LiteralPath $script:out -Name IsReadOnly -Value $false
        }
    }
}
