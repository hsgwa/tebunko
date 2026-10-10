# TSV のファイルの読み書き（shared\core\tsv_file.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "prettyTsv" -Tag Io {
    It "UTF-16のTSVを整形してBOM付きUTF-8で保存する（[ ] を含むパスも可）" {
        $in = "$TestDrive\sheet1.tmp"
        $out = "$TestDrive\見積[確定].xlsx_一覧.tsv"
        [System.IO.File]::WriteAllText($in, "`"a`nb`"`tc`t`r`n`t`r`n", [System.Text.Encoding]::Unicode)

        prettyTsv $in $out | Should -Be $true
        $bytes = [System.IO.File]::ReadAllBytes($out)
        $bytes[0..2] -join "," | Should -Be "239,187,191"
        [System.IO.File]::ReadAllText($out) | Should -Be "`"a${cellNewLine}b`"`tc`r`n"
    }

    It "出力範囲の左上の位置を指定できる" {
        $in = "$TestDrive\sheet2.tmp"
        $out = "$TestDrive\book.xlsx_B2.tsv"
        [System.IO.File]::WriteAllText($in, "x`r`n", [System.Text.Encoding]::Unicode)

        prettyTsv $in $out 2 2 | Should -Be $true
        [System.IO.File]::ReadAllText($out) | Should -Be "`r`n`tx`r`n"
    }

    It "内容が空なら保存せず `$false を返す" {
        $in = "$TestDrive\empty.tmp"
        $out = "$TestDrive\empty.tsv"
        [System.IO.File]::WriteAllText($in, "`t`t`r`n`r`n", [System.Text.Encoding]::Unicode)

        prettyTsv $in $out | Should -Be $false
        Test-Path -LiteralPath $out | Should -Be $false
    }
}
