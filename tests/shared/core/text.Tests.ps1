# TSV とセルの文字列（shared\core\text.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "replaceCellNewLine" -Tag Io {
    It "ダブルクォート内の改行（LF・CR・CRLF）をセル内改行の文字に置き換え、外の改行は残す" {
        replaceCellNewLine "`"a`nb`rc`r`nd`"`te`r`n" | Should -Be "`"a${cellNewLine}b${cellNewLine}c${cellNewLine}d`"`te`r`n"
    }
}

Describe "formatTsv" -Tag Io {
    It "行末の空セルと末尾の空行を取り除き、途中の空行は残す" {
        $lines = (formatTsv "a`tb`t`t`r`n`t`r`nc  `r`n`t`r`n") -split "`r`n"
        $lines.Count | Should -Be 3
        $lines[0] | Should -Be "a`tb"
        $lines[1] | Should -Be ""
        $lines[2] | Should -Be "c  "
    }

    It "セル内改行を含む行を1行にまとめる" {
        $lines = (formatTsv "`"x`r`ny`"`tz`r`nw`r`n") -split "`r`n"
        $lines.Count | Should -Be 2
        $lines[0] | Should -Be "`"x${cellNewLine}y`"`tz"
        $lines[1] | Should -Be "w"
    }

    It "出力範囲の左上の位置に合わせて、先頭に空行・空セルを補う（D5 → 5行目の4列目）" {
        $lines = (formatTsv "a`t`tb`r`n`r`nc`r`n" 5 4) -split "`r`n"
        $lines.Count | Should -Be 7
        $lines[0..3] -join "|" | Should -Be "|||"
        $lines[4] | Should -Be "`t`t`ta`t`tb"
        $lines[5] | Should -Be ""
        $lines[6] | Should -Be "`t`t`tc"
    }

    It "空白だけなら空文字を返す" {
        formatTsv "`t `t`r`n`r`n" 3 2 | Should -Be ""
    }
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

Describe "countTsvFields" -Tag Io {
    It "タブ区切りのセル数を返す" {
        countTsvFields "a`t`tb" | Should -Be 3
        countTsvFields "" | Should -Be 1
    }

    It "先頭のセルが空でも数え落とさない" {
        countTsvFields "`tb`tc" | Should -Be 3
        countTsvFields "`t`"b`tc`"" | Should -Be 2
        countTsvFields "`t" | Should -Be 2
    }

    It '" で始まるセル内のタブ・改行は区切りとしない' {
        countTsvFields "a`t`"左`t右`"`t`"`"`"x`"`"`"" | Should -Be 3
        countTsvFields "`"1行目`n2行目`"`tb" | Should -Be 2
    }

    It '" で始まらないセルの " は囲みとしない' {
        countTsvFields "a`"b`tc`"d" | Should -Be 2
    }
}

Describe "toColumnName" -Tag Io {
    It "列番号を列名に変換する" {
        toColumnName 1 | Should -Be "A"
        toColumnName 26 | Should -Be "Z"
        toColumnName 27 | Should -Be "AA"
        toColumnName 702 | Should -Be "ZZ"
        toColumnName 703 | Should -Be "AAA"
        toColumnName 16384 | Should -Be "XFD"
    }
}

Describe "splitTsvCells" -Tag Io {
    It 'ダブルクォートで囲まれたセルはタブを含んでも1セルとし、囲みを外して "" を " に戻す' {
        $cells = splitTsvCells "a`t`"b`tc`"`t`"d`"`"e`"`t"
        $cells.Count | Should -Be 4
        $cells[0] | Should -Be "a"
        $cells[1] | Should -Be "b`tc"
        $cells[2] | Should -Be "d`"e"
        $cells[3] | Should -Be ""
    }

    It "1セルでも配列で返す" {
        (splitTsvCells "a").Count | Should -Be 1
    }

    It "先頭のセルが空でも次のセルを落とさない" {
        $cells = splitTsvCells "`tりんご`t`"x`ty`""
        $cells.Count | Should -Be 3
        $cells[0] | Should -Be ""
        $cells[1] | Should -Be "りんご"
        $cells[2] | Should -Be "x`ty"
    }
}
