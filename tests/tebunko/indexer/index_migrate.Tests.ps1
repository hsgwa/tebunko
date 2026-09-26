# TSV のインデックスへの取り込みと、作業フォルダ・外したフォルダのインデックスの後始末（tebunko\indexer\index_migrate.ps1）のテスト。
# 作業フォルダ（$tmpDir）・出力用のフォルダ（$publishDir）・インデックスのフォルダ（$indexDir）は indexer.ps1 が決めるため、
# テストごとに TestDrive の下に差し替える。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"

    # 存在しないプロセスID（Windows のプロセスIDは 4 の倍数のため、奇数は使われない）
    $deadPid = 999999999
}

Describe "publishTsv" -Tag Io {
    It "作業フォルダのTSVを、そのファイルのインデックスのフォルダへ移す" {
        $tmpDir = "$TestDrive\publish\tmp"
        $publishDir = "$TestDrive\publish\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        $bookDir = "$TestDrive\publish\index\営業\見積.xlsx"
        newTsv "$tmpDir\Sheet1.tsv" @("a`tb")
        newTsv "$tmpDir\Sheet2.tsv" @("c")
        newTsv "$bookDir\古いシート.tsv" @("old")

        publishTsv $bookDir

        @(Get-ChildItem -LiteralPath $bookDir | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @("Sheet1.tsv", "Sheet2.tsv")
        @(Get-ChildItem -LiteralPath $tmpDir).Count | Should -Be 0
        Test-Path -LiteralPath "$publishDir\見積.xlsx" | Should -Be $false
    }

    It "ファイル名に [ ] があっても取り込み、TSV 以外のファイルは取り込まない" {
        $tmpDir = "$TestDrive\publish2\tmp"
        $publishDir = "$TestDrive\publish2\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        $bookDir = "$TestDrive\publish2\index\営業\[確定]見積.xlsx"
        newTsv "$tmpDir\[表]売上.tsv" @("a")
        newTsv "$tmpDir\作業.txt" @("x")

        publishTsv $bookDir

        @(Get-ChildItem -LiteralPath $bookDir | ForEach-Object { $_.Name }) | Should -Be @("[表]売上.tsv")
        Test-Path -LiteralPath "$tmpDir\作業.txt" | Should -Be $true
    }

    It "取り込んだ TSV が 0 件なら、前回のインデックスを空にする（文字の無いファイルになった）" {
        $tmpDir = "$TestDrive\publish3\tmp"
        $publishDir = "$TestDrive\publish3\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        $bookDir = "$TestDrive\publish3\index\営業\空.xlsx"
        [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
        newTsv "$bookDir\Sheet1.tsv" @("old")

        publishTsv $bookDir

        Test-Path -LiteralPath $bookDir -PathType Container | Should -Be $true
        @(Get-ChildItem -LiteralPath $bookDir).Count | Should -Be 0
    }

    It "260 文字を超えるパスのインデックスにも取り込む" {
        $tmpDir = "$TestDrive\publish4\tmp"
        $publishDir = "$TestDrive\publish4\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        $bookDir = "$TestDrive\publish4\index\" + ("深いフォルダ" * 20) + "\" + ("もっと深いフォルダ" * 15) + "\見積.xlsx"
        newTsv "$tmpDir\Sheet1.tsv" @("a")
        try {
            $bookDir.Length | Should -BeGreaterThan 260
            publishTsv $bookDir
            [System.IO.File]::Exists((toLongPath "$bookDir\Sheet1.tsv")) | Should -Be $true
        } finally {
            # TestDrive の後片付けは長いパスを消せないため、ここで消す
            removeDirectoryRetry "$TestDrive\publish4"
        }
    }
}

Describe "clearTmpDir" -Tag Io {
    It "作業フォルダ直下のファイルだけを消す" {
        $tmpDir = "$TestDrive\clear"
        newTsv "$tmpDir\a.tsv" @("a")
        newTsv "$tmpDir\sub\b.tsv" @("b")

        clearTmpDir

        Test-Path -LiteralPath "$tmpDir\a.tsv" | Should -Be $false
        Test-Path -LiteralPath "$tmpDir\sub\b.tsv" | Should -Be $true
    }
}

Describe "removeTmpDir" -Tag Io {
    It "作業フォルダと出力用のフォルダを削除する" {
        $tmpDir = "$TestDrive\remove\tmp"
        $publishDir = "$TestDrive\remove\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        newTsv "$tmpDir\a.tsv" @("a")
        newTsv "$publishDir\b.xlsx\b.tsv" @("b")

        removeTmpDir

        Test-Path -LiteralPath $tmpDir | Should -Be $false
        Test-Path -LiteralPath $publishDir | Should -Be $false
    }

    It "削除できなくても止まらず、次のフォルダも削除する" {
        $tmpDir = "$TestDrive\remove_fail\tmp"
        $publishDir = "$TestDrive\remove_fail\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        newTsv "$tmpDir\使用中.tsv" @("a")
        newTsv "$publishDir\b.xlsx\b.tsv" @("b")
        # 1 つ目（作業フォルダ）の中のファイルをほかから開いておき、削除に失敗させる
        $stream = [System.IO.File]::Open("$tmpDir\使用中.tsv", "Open", "Read", "None")
        try {
            { removeTmpDir } | Should -Not -Throw
            Test-Path -LiteralPath "$tmpDir\使用中.tsv" | Should -Be $true
        } finally {
            $stream.Dispose()
        }
        Test-Path -LiteralPath $publishDir | Should -Be $false
    }
}

Describe "removeStaleProcessDirs" -Tag Io {
    It "終了したプロセスのフォルダだけを削除し、実行中のプロセス・自分・プロセスID以外の名前のフォルダは残す" {
        $parent = "$TestDrive\stale"
        foreach ($name in @([string]$deadPid, "4", [string]$PID, "abc", "1234567890")) {
            [System.IO.Directory]::CreateDirectory("$parent\$name") | Out-Null
        }
        newTsv "$parent\$deadPid\a.tsv" @("a")

        removeStaleProcessDirs $parent

        Test-Path -LiteralPath "$parent\$deadPid" | Should -Be $false
        Test-Path -LiteralPath "$parent\4" | Should -Be $true  # System プロセス（実行中）
        Test-Path -LiteralPath "$parent\$PID" | Should -Be $true
        Test-Path -LiteralPath "$parent\abc" | Should -Be $true
        Test-Path -LiteralPath "$parent\1234567890" | Should -Be $true  # 10 桁はプロセスIDとみなさない
    }

    It "削除できないフォルダは次回に回す" {
        $parent = "$TestDrive\stale_locked"
        [System.IO.Directory]::CreateDirectory("$parent\$deadPid") | Out-Null
        Mock Remove-Item { throw "使用中" }

        { removeStaleProcessDirs $parent } | Should -Not -Throw
        Test-Path -LiteralPath "$parent\$deadPid" | Should -Be $true
    }
}

Describe "removeStaleTmpDirs" -Tag Io {
    It "作業フォルダと出力用のフォルダの、それぞれの親フォルダを片付ける" {
        $tmpDir = "$TestDrive\stale_both\temp\$PID"
        $publishDir = "$TestDrive\stale_both\出力\$PID"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_both\temp\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_both\出力\$deadPid") | Out-Null

        removeStaleTmpDirs

        Test-Path -LiteralPath "$TestDrive\stale_both\temp\$deadPid" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\stale_both\出力\$deadPid" | Should -Be $false
    }
}

Describe "removeDroppedFolders" -Tag Io {
    It "クロール対象から削除されたフォルダのインデックスだけを削除する" {
        $indexDir = "$TestDrive\dropped\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("a")
        newTsv "$indexDir\技術\仕様.docx\ページ001.tsv" @("b")
        $folders = @([pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" })
        $previous = @(
            [pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" },
            [pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" },
            [pscustomobject]@{ Path = "C:\data\消えた"; Name = "消えた" }  # インデックスのフォルダが無い
        )

        removeDroppedFolders $folders $previous

        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should -Be $true
        Test-Path -LiteralPath "$indexDir\技術" | Should -Be $false
    }

    It "インデックス名に [ ] があり、中に 260 文字を超えるパスがあっても削除する" {
        $indexDir = "$TestDrive\dropped2\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        $deep = "$indexDir\[旧]営業\" + ("深いフォルダ" * 20) + "\" + ("もっと深いフォルダ" * 15)
        newTsv (toLongPath "$deep\見積.xlsx\Sheet1.tsv") @("a")
        newTsv "$indexDir\[旧]営業2\見積.xlsx\Sheet1.tsv" @("b")  # 名前の先頭が同じだけの別のインデックスは残す

        removeDroppedFolders @([pscustomobject]@{ Path = "C:\data\営業2"; Name = "[旧]営業2" }) @(
            [pscustomobject]@{ Path = "C:\data\旧営業"; Name = "[旧]営業" },
            [pscustomobject]@{ Path = "C:\data\営業2"; Name = "[旧]営業2" }
        )

        [System.IO.Directory]::Exists((toLongPath "$indexDir\[旧]営業")) | Should -Be $false
        Test-Path -LiteralPath "$indexDir\[旧]営業2\見積.xlsx\Sheet1.tsv" | Should -Be $true
    }

    It "前回のクロール対象フォルダが無ければ何も削除しない" {
        $indexDir = "$TestDrive\dropped3\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("a")

        removeDroppedFolders @() $null

        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should -Be $true
    }

    It "インデックス名の大文字・小文字だけが違うフォルダは、同じフォルダとして残す" {
        $indexDir = "$TestDrive\dropped4\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv "$indexDir\Sales\見積.xlsx\Sheet1.tsv" @("a")

        removeDroppedFolders @([pscustomobject]@{ Path = "C:\data\sales"; Name = "sales" }) @([pscustomobject]@{ Path = "C:\data\Sales"; Name = "Sales" })

        Test-Path -LiteralPath "$indexDir\Sales\見積.xlsx\Sheet1.tsv" | Should -Be $true
    }
}
