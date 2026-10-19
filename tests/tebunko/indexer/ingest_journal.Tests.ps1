# 取り込みのジャーナル（tebunko\indexer\ingest_journal.ps1 の IngestJournal・前の版の ingesting.txt の読み取り）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\ingest_journal.ps1"
}

Describe "IngestJournal" -Tag Io {
    It "追記した行は、閉じる前でも別の読み手から読め、前回のジャーナルとして順に読める（BOM は新しいファイルのときだけ・CRLF）" {
        $path = "$TestDrive\journal[1]\store\journal\ingest.tsv"
        $journal = [IngestJournal]::new($path)
        @($journal.ReadPrevious()).Count | Should -Be 0

        $journal.AddRecord("版`t2")
        $journal.AddRecord("始`t営業\a [確定].xlsx`t1")
        # 開いたままでも、書いた行は読める（強制終了しても残る）
        @([IngestJournal]::new($path).ReadPrevious()) | Should -Be @("版`t2", "始`t営業\a [確定].xlsx`t1")
        $journal.Close()

        $again = [IngestJournal]::new($path)
        $again.AddRecord("終")
        $again.Close()
        $bytes = [System.IO.File]::ReadAllBytes($path)
        ($bytes[0..2] -join ",") | Should -Be "239,187,191"
        [System.Text.Encoding]::UTF8.GetString($bytes).Split([string[]]@("`r`n"), [System.StringSplitOptions]::None).Count | Should -Be 4
        @([IngestJournal]::new($path).ReadPrevious()) | Should -Be @("版`t2", "始`t営業\a [確定].xlsx`t1", "終")
    }

    It "最後の行が改行で終わっていない（書いている途中で止まった）ときは、その行だけ捨てる。LF の行も読める" {
        $path = "$TestDrive\journal_cut.tsv"
        [System.IO.File]::WriteAllText($path, "版`t2`n始`t営業\a.xlsx`t1`n始`t営業\b", (New-Object System.Text.UTF8Encoding($true)))
        @([IngestJournal]::new($path).ReadPrevious()) | Should -Be @("版`t2", "始`t営業\a.xlsx`t1")

        [System.IO.File]::WriteAllText($path, "途中", (New-Object System.Text.UTF8Encoding($true)))
        @([IngestJournal]::new($path).ReadPrevious()).Count | Should -Be 0
    }

    It "Reset は、開いている書き込みを閉じて、渡した行だけにする。そのあとも追記できる" {
        $path = "$TestDrive\journal_reset.tsv"
        $journal = [IngestJournal]::new($path)
        $journal.AddRecord("版`t2")
        $journal.AddRecord("始`t営業\a.xlsx`t1")
        $journal.Reset([string[]]@("版`t2", "回`t2026/10/17 10:00:00`t3", "持`t営業\b.xlsx`t1"))
        @($journal.ReadPrevious()) | Should -Be @("版`t2", "回`t2026/10/17 10:00:00`t3", "持`t営業\b.xlsx`t1")

        $journal.AddRecord("始`t営業\b.xlsx`t2")
        $journal.Close()
        @([IngestJournal]::new($path).ReadPrevious())[-1] | Should -Be "始`t営業\b.xlsx`t2"
    }
}

Describe "前の版の取り込み中のファイルの記録（ingesting.txt）" -Tag Io {
    It "回数と相対パスを読み、壊れた行・使えない相対パスは読み飛ばす。読んだあとは removeLegacyIngesting で消せる（無くても消せる）" {
        $dir = "$TestDrive\legacy[1]"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $path = getLegacyIngestingPath $dir
        $path | Should -Be "$dir\ingesting.txt"
        writeListFile $path @("2`tフォルダ1\a [確定]\見積.xlsx", "x`t営業\a.xlsx", "0`t営業\a.xlsx", "1`t", "営業\a.xlsx", "", "1`t..\外\a.xlsx", "1`t営業", "3`t営業\残る.xlsx", "1`taux.backup\c.xlsx")

        $legacy = @(readLegacyIngesting $path)
        @($legacy | ForEach-Object { "$($_.RelPath)=$($_.Count)" }) | Should -Be @("フォルダ1\a [確定]\見積.xlsx=2", "営業\残る.xlsx=3", "aux.backup\c.xlsx=1")

        removeLegacyIngesting $path
        Test-Path -LiteralPath $path | Should -Be $false
        @(readLegacyIngesting $path).Count | Should -Be 0
        { removeLegacyIngesting $path } | Should -Not -Throw
    }
}

Describe "同じフォルダを 2 回書き出しても結果が同じ（仕事のやり直しが冪等）" -Tag Io {
    It "TSV を置き直して書き出し直すと、本文インデックスのファイルの中身が最初の書き出しと同じになる" {
        $work = Join-Path $TestDrive "idempotent"
        $index = "$work\index"
        $folder = "$index\営業"
        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $pending[$folder] = @()
        $tsv = "$folder\A.xlsx\$(toIndexFileName "S")"

        newTsv $tsv @("A の単価")
        [void](publishIndexFolders $pending $index)
        $file = "$folder\content_index.xlsx.001.tsv"
        $first = [System.IO.File]::ReadAllBytes($file)

        # 書き出しの途中で止まった跡と同じ状態（同じ仕事をやり直して、TSV がもう一度置かれた）から、もう一度書き出す
        newTsv $tsv @("A の単価")
        [void](publishIndexFolders $pending $index)
        $second = [System.IO.File]::ReadAllBytes($file)

        ($second -join ",") | Should -Be ($first -join ",")
        Test-Path -LiteralPath "$folder\A.xlsx" | Should -Be $false
    }
}
