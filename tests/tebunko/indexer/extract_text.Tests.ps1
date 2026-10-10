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

    It "<name> は、コピーが残らず（出力フォルダが空）、結果は <expected>" -TestCases @(
        @{ name = "空のファイル"; bytes = [byte[]]@(); max = 1000; message = $null; expected = "0" }
        @{ name = "空の行だけのファイル"; bytes = [System.Text.Encoding]::UTF8.GetBytes("`r`n`r`n"); max = 1000; message = $null; expected = "0" }
        @{ name = "大きさの上限を超えるファイル"; bytes = [System.Text.Encoding]::UTF8.GetBytes("12345678"); max = 4; message = "ファイルサイズが大きすぎるため更新できません。"; expected = "失敗" }
        @{ name = "バイナリ（コピーを読む途中の例外）"; bytes = [byte[]]((0..39 | ForEach-Object { if ($_ -eq 0 -or $_ -eq 3) { 0 } else { $_ % 5 + 1 } })); max = 1000; message = "テキストファイルではないため更新できません。"; expected = "失敗" }
    ) {
        $source = newSourceFile "case.txt" $bytes
        if ($message) {
            { extractTextFile $source $outDir $max } | Should -Throw $message
        } else {
            (extractTextFile $source $outDir $max) | Should -Be 0
        }
        (Get-ChildItem -LiteralPath $outDir -File).Count | Should -Be 0
    }

    It "大きさの上限を超えるときは、コピーを作らない" {
        Mock copyFileShared { }
        $source = newSourceFile "big2.txt" ([System.Text.Encoding]::UTF8.GetBytes("12345678"))
        { extractTextFile $source $outDir 4 } | Should -Throw "ファイルサイズが大きすぎるため更新できません。"
        Should -Invoke copyFileShared -Times 0 -Exactly
    }

    It "取り込み後の出力フォルダには本文の TSV だけがあり、コピーが残らない" {
        $source = newSourceFile "keep.txt" ([System.Text.Encoding]::UTF8.GetBytes("あ`r`nい"))
        (extractTextFile $source $outDir) | Should -Be 1
        (@(Get-ChildItem -LiteralPath $outDir -File | ForEach-Object { $_.Name }) -join ",") | Should -Be "doc_body.tsv"
    }

    It "元のファイルの拡張子が .tsv でも、出力フォルダには本文の TSV だけがあり、中身は元のとおり" {
        $source = newSourceFile "table.tsv" ([System.Text.Encoding]::UTF8.GetBytes("列1`t列2"))
        (extractTextFile $source $outDir) | Should -Be 1
        (@(Get-ChildItem -LiteralPath $outDir -File | ForEach-Object { $_.Name }) -join ",") | Should -Be "doc_body.tsv"
        [System.IO.File]::ReadAllLines((Join-Path $outDir "doc_body.tsv"))[0] | Should -Be "列1`t列2"
    }

    It "コピーを消せなかったときは、黙って進まず取り込みの失敗にする" {
        Mock Remove-Item { throw "拒否" }
        $source = newSourceFile "stuck.txt" ([System.Text.Encoding]::UTF8.GetBytes("あ"))
        { extractTextFile $source $outDir } | Should -Throw "作業領域のコピーを消せませんでした*"
        Test-Path -LiteralPath (Join-Path $outDir "doc_body.tsv") | Should -Be $false
    }

    It "元のファイルが無いときは、ファイルが見つからない例外になる（コピーは作らない）" {
        $caught = $null
        try { extractTextFile (Join-Path $TestDrive "none.txt") $outDir } catch { $caught = $_.Exception }
        $caught | Should -Not -BeNullOrEmpty
        $inner = $(if ($caught.InnerException) { $caught.InnerException } else { $caught })
        $inner | Should -BeOfType ([System.IO.FileNotFoundException])
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
