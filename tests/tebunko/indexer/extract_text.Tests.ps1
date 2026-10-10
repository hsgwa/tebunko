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
        $path = Join-Path $outDir "doc_body.tsv"
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
        { extractTextFile $source $outDir 4 } | Should -Throw "ファイルサイズが大きすぎるため更新できません。"
    }

    It "バイナリと判定したファイルは、決めた文言で失敗にする" {
        $bytes = New-Object byte[] 40
        for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [byte]($i % 5 + 1) }  # NUL を含まないが有効な文字コードでもない並び
        # 偏りの無い NUL を混ぜて、確実にバイナリと判定させる
        $bytes[0] = 0; $bytes[3] = 0
        $source = newSourceFile "bin.txt" $bytes
        { extractTextFile $source $outDir } | Should -Throw "テキストファイルではないため更新できません。"
    }

    It "取り込み後の出力フォルダには TSV だけがあり、コピー（source.*）が残らない" {
        $source = newSourceFile "keep.txt" ([System.Text.Encoding]::UTF8.GetBytes("あ`r`nい"))
        (extractTextFile $source $outDir) | Should -Be 1
        (@(Get-ChildItem -LiteralPath $outDir -File | ForEach-Object { $_.Name }) -join ",") | Should -Be "doc_body.tsv"
    }

    It "<name> でも、コピーが残らない" -TestCases @(
        @{ name = "空のファイル"; bytes = [byte[]]@(); max = 1000; message = $null }
        @{ name = "空の行だけのファイル"; bytes = [System.Text.Encoding]::UTF8.GetBytes("`r`n`r`n"); max = 1000; message = $null }
        @{ name = "バイナリ（読み取り中の例外）"; bytes = [byte[]]((0..39 | ForEach-Object { if ($_ -eq 0 -or $_ -eq 3) { 0 } else { $_ % 5 + 1 } })); max = 1000; message = "テキストファイルではないため更新できません。" }
    ) {
        $source = newSourceFile "case.txt" $bytes
        if ($message) {
            { extractTextFile $source $outDir $max } | Should -Throw $message
        } else {
            (extractTextFile $source $outDir $max) | Should -Be 0
        }
        (Get-ChildItem -LiteralPath $outDir -File).Count | Should -Be 0
    }

    It "大きさの上限を超えるときは、コピーを作らずに失敗する（出力フォルダが空のまま）" {
        $source = newSourceFile "big2.txt" ([System.Text.Encoding]::UTF8.GetBytes("12345678"))
        { extractTextFile $source $outDir 4 } | Should -Throw "ファイルサイズが大きすぎるため更新できません。"
        (Get-ChildItem -LiteralPath $outDir -File).Count | Should -Be 0
    }

    It "別のストリームが書き込みで開いている元のファイルも取り込める" {
        $source = newSourceFile "busy.txt" ([System.Text.Encoding]::UTF8.GetBytes("開いたまま"))
        $stream = [System.IO.FileStream]::new($source, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::ReadWrite)
        try {
            (extractTextFile $source $outDir) | Should -Be 1
        } finally {
            $stream.Dispose()
        }
        [System.IO.File]::ReadAllLines((Join-Path $outDir "doc_body.tsv"))[0] | Should -Be "開いたまま"
    }

    It "元のファイルの更新日時・大きさ・中身は、取り込みの前後で変わらない" {
        $source = newSourceFile "same.txt" ([System.Text.Encoding]::UTF8.GetBytes("そのまま"))
        [System.IO.File]::SetLastWriteTimeUtc($source, [datetime]::new(2020, 1, 2, 3, 4, 5, [System.DateTimeKind]::Utc))
        $before = @{ time = [System.IO.File]::GetLastWriteTimeUtc($source); length = ([System.IO.FileInfo]$source).Length; hash = (Get-FileHash -LiteralPath $source).Hash }
        (extractTextFile $source $outDir) | Should -Be 1
        [System.IO.File]::GetLastWriteTimeUtc($source) | Should -Be $before.time
        ([System.IO.FileInfo]$source).Length | Should -Be $before.length
        (Get-FileHash -LiteralPath $source).Hash | Should -Be $before.hash
    }
}
