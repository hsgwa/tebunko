# 1 ファイル（テキスト）の抽出（tebunko\indexer\extract_text.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\extract_text.ps1"

    function newSourceFile([string]$name, [byte[]]$bytes) {
        $path = Join-Path $TestDrive $name
        [System.IO.File]::WriteAllBytes($path, $bytes)
        return $path
    }
}

Describe "extractTextFile" -Tag Io {
    BeforeEach {
        $outDir = Join-Path $TestDrive "out-$(New-Guid)"
        [System.IO.Directory]::CreateDirectory($outDir) | Out-Null
    }

    It "本文の TSV を 1 つ書き出し、途中の空の行も捨てない" {
        $source = newSourceFile "a.txt" ([System.Text.Encoding]::UTF8.GetBytes("1行目`r`n`r`n3行目`r`n"))
        (extractTextFile $source $outDir) | Should -Be 1
        $path = Join-Path $outDir "本文.tsv"
        Test-Path -LiteralPath $path | Should -Be $true
        $lines = [System.IO.File]::ReadAllLines($path)
        ($lines -join "|") | Should -Be "1行目||3行目"
    }

    It "空のファイルは TSV を書き出さず 0 を返す" {
        $source = newSourceFile "empty.txt" ([byte[]]@())
        (extractTextFile $source $outDir) | Should -Be 0
        (Get-ChildItem -LiteralPath $outDir -File).Count | Should -Be 0
    }

    It "空の行だけのファイルは TSV を書き出さず 0 を返す" {
        $source = newSourceFile "blank.txt" ([System.Text.Encoding]::UTF8.GetBytes("`r`n`r`n"))
        (extractTextFile $source $outDir) | Should -Be 0
        (Get-ChildItem -LiteralPath $outDir -File).Count | Should -Be 0
    }

    It "大きさの上限（差し替えた小さい値）を超えるファイルは、決めた文言で失敗にする" {
        $source = newSourceFile "big.txt" ([System.Text.Encoding]::UTF8.GetBytes("12345678"))
        { extractTextFile $source $outDir 4 } | Should -Throw "ファイルサイズが大きすぎるため取り込めません。"
    }

    It "バイナリと判定したファイルは、決めた文言で失敗にする" {
        $bytes = New-Object byte[] 40
        for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [byte]($i % 5 + 1) }  # NUL を含まないが有効な文字コードでもない並び
        # 偏りの無い NUL を混ぜて、確実にバイナリと判定させる
        $bytes[0] = 0; $bytes[3] = 0
        $source = newSourceFile "bin.txt" $bytes
        { extractTextFile $source $outDir } | Should -Throw "テキストファイルではないため取り込めません。"
    }
}
