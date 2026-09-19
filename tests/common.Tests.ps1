# Pester 3.4 以降で実行: Invoke-Pester .\tests
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here\..\scripts\common.ps1"

Describe "toSafeFileName" {
    It "ファイル名に使えない文字を全角に変換する" {
        toSafeFileName 'a<b>c\d*e:f?g|h' | Should Be 'a＜b＞c￥d＊e：f？g｜h'
    }

    It "スラッシュを全角に変換する" {
        toSafeFileName 'a/b' | Should Be 'a／b'
    }

    It "ダブルクォートを全角に変換する" {
        toSafeFileName 'a"b' | Should Be 'a”b'
    }

    It "使える文字はそのまま返す" {
        toSafeFileName 'シート1 (2)' | Should Be 'シート1 (2)'
    }
}

Describe "replaceNewLineToSpace" {
    It "ダブルクォート内の改行を取り除く" {
        replaceNewLineToSpace "`"a`r`nb`"`tc`r`n" | Should Be "`"ab`"`tc`r`n"
    }

    It "ダブルクォート外の改行は残す" {
        replaceNewLineToSpace "a`r`nb" | Should Be "a`r`nb"
    }
}

Describe "formatTsv" {
    It "空行と行末の空白を取り除く" {
        $lines = @((formatTsv "a`tb`t`t`r`n`t`r`nc  `r`n") -split "\r?\n" | Where-Object { $_ -ne "" })
        $lines.Count | Should Be 2
        $lines[0] | Should Be "a`tb"
        $lines[1] | Should Be "c"
    }

    It "セル内改行を含む行を1行にまとめる" {
        $lines = @((formatTsv "`"x`r`ny`"`tz`r`n") -split "\r?\n" | Where-Object { $_ -ne "" })
        $lines.Count | Should Be 1
        $lines[0] | Should Be "`"xy`"`tz"
    }
}

Describe "prettyTsv" {
    It "UTF-16のTSVを整形してBOM付きUTF-8で保存する（[ ] を含むパスも可）" {
        $in = "$TestDrive\sheet1.tmp"
        $out = "$TestDrive\見積[確定].xlsx_一覧.tsv"
        [System.IO.File]::WriteAllText($in, "`"a`nb`"`tc`t`r`n`t`r`n", [System.Text.Encoding]::Unicode)

        prettyTsv $in $out | Should Be $true
        $bytes = [System.IO.File]::ReadAllBytes($out)
        $bytes[0..2] -join "," | Should Be "239,187,191"
        [System.IO.File]::ReadAllText($out) | Should Be "`"ab`"`tc`r`n"
    }

    It "内容が空なら保存せず `$false を返す" {
        $in = "$TestDrive\empty.tmp"
        $out = "$TestDrive\empty.tsv"
        [System.IO.File]::WriteAllText($in, "`t`t`r`n`r`n", [System.Text.Encoding]::Unicode)

        prettyTsv $in $out | Should Be $false
        Test-Path -LiteralPath $out | Should Be $false
    }
}

Describe "readListFile / writeListFile" {
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
}

Describe "splitIndexFileName" {
    It "ブック名とシート名に分解する" {
        $name = splitIndexFileName "book.xlsx_Sheet1.tsv"
        $name.book | Should Be "book.xlsx"
        $name.sheet | Should Be "Sheet1"
    }

    It "シート名に _ を含んでも分解できる" {
        $name = splitIndexFileName "売上.xls_2024_上期.tsv"
        $name.book | Should Be "売上.xls"
        $name.sheet | Should Be "2024_上期"
    }

    It "xlsm も扱える" {
        (splitIndexFileName "macro.xlsm_A.tsv").book | Should Be "macro.xlsm"
    }

    It "形式外のファイル名はそのままブック名として返す" {
        $name = splitIndexFileName "other.tsv"
        $name.book | Should Be "other.tsv"
        $name.sheet | Should Be ""
    }
}

Describe "toResultLine" {
    It "ブック名・シート名・該当行をタブ区切りにする" {
        toResultLine "book.xlsx_Sheet1.tsv" "時刻`t12:34:56" | Should Be "book.xlsx`tSheet1`t時刻`t12:34:56"
    }
}

Describe "readConfigLines" {
    It "ファイルが無ければ空ファイルを作成して例外を投げる" {
        $path = "$TestDrive\new\設定.txt"
        { readConfigLines $path } | Should Throw "作成しました"
        Test-Path $path | Should Be $true
    }

    It "空行を除き前後の空白を取り除いて返す" {
        $path = "$TestDrive\words.txt"
        Set-Content $path -Value @("  検索1 ", "", "検索2", "   ") -Encoding UTF8
        $lines = readConfigLines $path
        $lines.Count | Should Be 2
        $lines[0] | Should Be "検索1"
        $lines[1] | Should Be "検索2"
    }
}

Describe "getTargetFolder" {
    It "1行ならその値を返す" {
        $path = "$TestDrive\target1.txt"
        Set-Content $path -Value @("C:\データ\Excel", "") -Encoding UTF8
        getTargetFolder $path | Should Be "C:\データ\Excel"
    }

    It "2行以上なら例外を投げる" {
        $path = "$TestDrive\target2.txt"
        Set-Content $path -Value @("C:\a", "C:\b") -Encoding UTF8
        { getTargetFolder $path } | Should Throw "１行だけ"
    }

    It "空なら例外を投げる" {
        $path = "$TestDrive\target0.txt"
        Set-Content $path -Value @("") -Encoding UTF8
        { getTargetFolder $path } | Should Throw "１行だけ"
    }
}

Describe "getIndexFolders" {
    It "設定ファイルが無ければ work\index を返す" {
        getIndexFolders "$TestDrive\none.txt" | Should Be $indexDir
    }

    It "設定ファイルが空なら work\index を返す" {
        $path = "$TestDrive\index_empty.txt"
        Set-Content $path -Value @("") -Encoding UTF8
        getIndexFolders $path | Should Be $indexDir
    }

    It "設定ファイルの各行を返す" {
        $path = "$TestDrive\index.txt"
        Set-Content $path -Value @("D:\index1", "D:\index2") -Encoding UTF8
        $folders = getIndexFolders $path
        $folders.Count | Should Be 2
        $folders[1] | Should Be "D:\index2"
    }
}

Describe "パス定義" {
    It "リポジトリ直下を基準にする" {
        $rootDir | Should Be (Resolve-Path "$here\..").Path
        $indexDir | Should Be "$rootDir\work\index"
        $resultFile | Should Be "$rootDir\output\検索結果.txt"
    }
}

Describe "スクリプトの構文" {
    Get-ChildItem "$here\..\scripts\*.ps1" | ForEach-Object {
        $script = $_

        It "$($script.Name) に構文エラーが無い" {
            $errors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$null, [ref]$errors) | Out-Null
            $errors.Count | Should Be 0
        }
    }
}
