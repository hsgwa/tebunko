# TSV のインデックスへの取り込みと、作業フォルダ・以前の形式の後始末（tebunko_grep\indexer\index_migrate.ps1）のテスト。
# 作業フォルダ（$tmpDir）・出力用のフォルダ（$publishDir）・インデックスのフォルダ（$indexDir）は indexer.ps1 が決めるため、
# テストごとに TestDrive の下に差し替える。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\indexer\index_migrate.ps1"

# 存在しないプロセスID（Windows のプロセスIDは 4 の倍数のため、奇数は使われない）
$deadPid = 999999999

Describe "publishTsv" -Tag Io {
    It "作業フォルダのTSVを、そのファイルのインデックスのフォルダへ移す" {
        $tmpDir = "$TestDrive\publish\tmp"
        $publishDir = "$TestDrive\publish\出力"
        $bookDir = "$TestDrive\publish\index\営業\見積.xlsx"
        newTsv "$tmpDir\Sheet1.tsv" @("a`tb")
        newTsv "$tmpDir\Sheet2.tsv" @("c")
        newTsv "$bookDir\古いシート.tsv" @("old")

        publishTsv $bookDir

        @(Get-ChildItem -LiteralPath $bookDir | ForEach-Object { $_.Name } | Sort-Object) | Should Be @("Sheet1.tsv", "Sheet2.tsv")
        @(Get-ChildItem -LiteralPath $tmpDir).Count | Should Be 0
        Test-Path -LiteralPath "$publishDir\見積.xlsx" | Should Be $false
    }

    It "ファイル名に [ ] があっても取り込み、TSV 以外のファイルは取り込まない" {
        $tmpDir = "$TestDrive\publish2\tmp"
        $publishDir = "$TestDrive\publish2\出力"
        $bookDir = "$TestDrive\publish2\index\営業\[確定]見積.xlsx"
        newTsv "$tmpDir\[表]売上.tsv" @("a")
        newTsv "$tmpDir\作業.txt" @("x")

        publishTsv $bookDir

        @(Get-ChildItem -LiteralPath $bookDir | ForEach-Object { $_.Name }) | Should Be @("[表]売上.tsv")
        Test-Path -LiteralPath "$tmpDir\作業.txt" | Should Be $true
    }

    It "取り込んだ TSV が 0 件なら、前回のインデックスを空にする（文字の無いファイルになった）" {
        $tmpDir = "$TestDrive\publish3\tmp"
        $publishDir = "$TestDrive\publish3\出力"
        $bookDir = "$TestDrive\publish3\index\営業\空.xlsx"
        [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
        newTsv "$bookDir\Sheet1.tsv" @("old")

        publishTsv $bookDir

        Test-Path -LiteralPath $bookDir -PathType Container | Should Be $true
        @(Get-ChildItem -LiteralPath $bookDir).Count | Should Be 0
    }

    It "260 文字を超えるパスのインデックスにも取り込む" {
        $tmpDir = "$TestDrive\publish4\tmp"
        $publishDir = "$TestDrive\publish4\出力"
        $bookDir = "$TestDrive\publish4\index\" + ("深いフォルダ" * 20) + "\" + ("もっと深いフォルダ" * 15) + "\見積.xlsx"
        newTsv "$tmpDir\Sheet1.tsv" @("a")
        try {
            $bookDir.Length | Should BeGreaterThan 260
            publishTsv $bookDir
            [System.IO.File]::Exists((toLongPath "$bookDir\Sheet1.tsv")) | Should Be $true
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

        Test-Path -LiteralPath "$tmpDir\a.tsv" | Should Be $false
        Test-Path -LiteralPath "$tmpDir\sub\b.tsv" | Should Be $true
    }
}

Describe "removeTmpDir" -Tag Io {
    It "作業フォルダと出力用のフォルダを削除する" {
        $tmpDir = "$TestDrive\remove\tmp"
        $publishDir = "$TestDrive\remove\出力"
        newTsv "$tmpDir\a.tsv" @("a")
        newTsv "$publishDir\b.xlsx\b.tsv" @("b")

        removeTmpDir

        Test-Path -LiteralPath $tmpDir | Should Be $false
        Test-Path -LiteralPath $publishDir | Should Be $false
    }

    It "削除できなくても止まらず、次のフォルダも削除する" {
        $tmpDir = "$TestDrive\remove_fail\tmp"
        $publishDir = "$TestDrive\remove_fail\出力"
        newTsv "$tmpDir\使用中.tsv" @("a")
        newTsv "$publishDir\b.xlsx\b.tsv" @("b")
        # 1 つ目（作業フォルダ）の中のファイルをほかから開いておき、削除に失敗させる
        $stream = [System.IO.File]::Open("$tmpDir\使用中.tsv", "Open", "Read", "None")
        try {
            { removeTmpDir } | Should Not Throw
            Test-Path -LiteralPath "$tmpDir\使用中.tsv" | Should Be $true
        } finally {
            $stream.Dispose()
        }
        Test-Path -LiteralPath $publishDir | Should Be $false
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

        Test-Path -LiteralPath "$parent\$deadPid" | Should Be $false
        Test-Path -LiteralPath "$parent\4" | Should Be $true  # System プロセス（実行中）
        Test-Path -LiteralPath "$parent\$PID" | Should Be $true
        Test-Path -LiteralPath "$parent\abc" | Should Be $true
        Test-Path -LiteralPath "$parent\1234567890" | Should Be $true  # 10 桁はプロセスIDとみなさない
    }

    It "削除できないフォルダは次回に回す" {
        $parent = "$TestDrive\stale_locked"
        [System.IO.Directory]::CreateDirectory("$parent\$deadPid") | Out-Null
        Mock Remove-Item { throw "使用中" }

        { removeStaleProcessDirs $parent } | Should Not Throw
        Test-Path -LiteralPath "$parent\$deadPid" | Should Be $true
    }
}

Describe "removeStaleTmpDirs" -Tag Io {
    It "作業フォルダと出力用のフォルダの、それぞれの親フォルダを片付ける" {
        $tmpDir = "$TestDrive\stale_both\temp\$PID"
        $publishDir = "$TestDrive\stale_both\出力\$PID"
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_both\temp\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_both\出力\$deadPid") | Out-Null

        removeStaleTmpDirs

        Test-Path -LiteralPath "$TestDrive\stale_both\temp\$deadPid" | Should Be $false
        Test-Path -LiteralPath "$TestDrive\stale_both\出力\$deadPid" | Should Be $false
    }
}

Describe "moveLegacyIndex" -Tag Io {
    function newStatus([object[]]$folders, [string[]]$relPaths) {
        $rows = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($relPath in $relPaths) {
            $rows[$relPath] = newStatusRow $relPath
        }
        return @{ Folders = $folders; Rows = $rows }
    }

    $folders = @(
        [pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" },
        [pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" }
    )

    It "取り込み一覧のインデックス名の無いフォルダのインデックスを、そのインデックス名の下へ移し、相対パスも付け替える" {
        $indexDir = "$TestDrive\legacy1\index"
        newTsv "$indexDir\見積.xlsx\Sheet1.tsv" @("a")
        newTsv "$indexDir\sub\報告.docx\ページ001.tsv" @("b")
        $status = newStatus @([pscustomobject]@{ Path = "C:\data\技術"; Name = "" }) @("見積.xlsx", "sub\報告.docx")

        $rows = moveLegacyIndex $folders $status $true

        Test-Path -LiteralPath "$indexDir\技術\見積.xlsx\Sheet1.tsv" | Should Be $true
        Test-Path -LiteralPath "$indexDir\技術\sub\報告.docx\ページ001.tsv" | Should Be $true
        Test-Path -LiteralPath "$indexDir\見積.xlsx" | Should Be $false
        Test-Path -LiteralPath "${indexDir}_移行中" | Should Be $false
        @($rows.Keys | Sort-Object) | Should Be @("技術\sub\報告.docx", "技術\見積.xlsx")
        $rows["技術\見積.xlsx"].相対パス | Should Be "技術\見積.xlsx"
    }

    It "取り込み一覧のフォルダの書き方（大文字・小文字・末尾の \）が違っても、同じフォルダとして移す" {
        $indexDir = "$TestDrive\legacy_case\index"
        newTsv "$indexDir\見積.xlsx\Sheet1.tsv" @("a")
        $status = newStatus @([pscustomobject]@{ Path = "c:\DATA\技術\"; Name = "" }) @("見積.xlsx")

        $rows = moveLegacyIndex $folders $status $true

        Test-Path -LiteralPath "$indexDir\技術\見積.xlsx\Sheet1.tsv" | Should Be $true
        @($rows.Keys) | Should Be @("技術\見積.xlsx")
    }

    It "以前の形式のインデックスに、インデックス名と同じ名前のフォルダがあっても移せる" {
        $indexDir = "$TestDrive\legacy_same\index"
        newTsv "$indexDir\技術\仕様.docx\ページ001.tsv" @("a")  # 元のフォルダの下の「技術」フォルダ
        $status = newStatus @([pscustomobject]@{ Path = "C:\data\技術"; Name = "" }) @("技術\仕様.docx")

        $rows = moveLegacyIndex $folders $status $true

        Test-Path -LiteralPath "$indexDir\技術\技術\仕様.docx\ページ001.tsv" | Should Be $true
        @($rows.Keys) | Should Be @("技術\技術\仕様.docx")
    }

    It "前回の移行が途中で止まり _移行中 が残っていれば、その続きから移す（取り込み一覧あり）" {
        # 1 回目の移動の後に止まり、次の実行で空の work\index が作られた状態
        $indexDir = "$TestDrive\legacy_resume\index"
        newTsv "${indexDir}_移行中\見積.xlsx\Sheet1.tsv" @("a")
        [System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
        $status = newStatus @([pscustomobject]@{ Path = "C:\data\技術"; Name = "" }) @("見積.xlsx")

        $rows = moveLegacyIndex $folders $status $true

        Test-Path -LiteralPath "$indexDir\技術\見積.xlsx\Sheet1.tsv" | Should Be $true
        Test-Path -LiteralPath "${indexDir}_移行中" | Should Be $false
        @($rows.Keys) | Should Be @("技術\見積.xlsx")
    }

    It "前回の移行が途中で止まり _移行中 が残っていれば、その続きから移す（取り込み一覧なし）" {
        $indexDir = "$TestDrive\legacy_resume2\index"
        newTsv "${indexDir}_移行中\見積.xlsx_Sheet1.tsv" @("a")
        [System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
        $status = newStatus @() @()

        $rows = moveLegacyIndex $folders $status $false

        $rows.Count | Should Be 0
        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx_Sheet1.tsv" | Should Be $true
        Test-Path -LiteralPath "${indexDir}_移行中" | Should Be $false
    }

    It "取り込み一覧が無く、クロール対象フォルダも無ければ、直下に何があっても移さない" {
        $indexDir = "$TestDrive\legacy_nofolder\index"
        newTsv "$indexDir\見積.xlsx_Sheet1.tsv" @("a")
        $status = newStatus @() @()

        $rows = moveLegacyIndex @() $status $false

        $rows.Count | Should Be 0
        Test-Path -LiteralPath "$indexDir\見積.xlsx_Sheet1.tsv" | Should Be $true
    }

    It "以前の形式のフォルダがクロール対象から外れていれば、移さずに知らせ、前回の行は使わない" {
        $indexDir = "$TestDrive\legacy2\index"
        newTsv "$indexDir\見積.xlsx\Sheet1.tsv" @("a")
        $status = newStatus @([pscustomobject]@{ Path = "C:\data\外した"; Name = "" }) @("見積.xlsx")

        $rows = moveLegacyIndex $folders $status $true

        $rows.Count | Should Be 0
        Test-Path -LiteralPath "$indexDir\見積.xlsx\Sheet1.tsv" | Should Be $true
    }

    It "取り込み一覧が無く、work\index 直下にインデックス名以外のものがあれば、1件目のフォルダのインデックスとみなす" {
        $indexDir = "$TestDrive\legacy3\index"
        newTsv "$indexDir\見積.xlsx_Sheet1.tsv" @("a")
        $status = newStatus @() @()

        $rows = moveLegacyIndex $folders $status $false

        $rows.Count | Should Be 0
        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx_Sheet1.tsv" | Should Be $true
    }

    It "取り込み一覧が無くても、直下がインデックス名のフォルダと 元のフォルダ.txt だけなら移さない" {
        $indexDir = "$TestDrive\legacy4\index"
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("a")
        newTsv "$indexDir\${sourceFolderFileName}" @("# 説明")
        $status = newStatus @() @("営業\見積.xlsx")

        $rows = moveLegacyIndex $folders $status $false

        @($rows.Keys) | Should Be @("営業\見積.xlsx")
        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should Be $true
    }

    It "取り込み一覧があり、以前の形式のフォルダが無ければ、前回の行をそのまま返す" {
        $indexDir = "$TestDrive\legacy5\index"
        newTsv "$indexDir\ばらばら.tsv" @("a")
        $status = newStatus @($folders) @("営業\見積.xlsx", "技術\仕様.docx")

        $rows = moveLegacyIndex $folders $status $true

        $rows.Count | Should Be 2
        Test-Path -LiteralPath "$indexDir\ばらばら.tsv" | Should Be $true
    }
}

Describe "removeDroppedFolders" -Tag Io {
    It "クロール対象から削除されたフォルダのインデックスだけを削除する" {
        $indexDir = "$TestDrive\dropped\index"
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("a")
        newTsv "$indexDir\技術\仕様.docx\ページ001.tsv" @("b")
        $folders = @([pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" })
        $previous = @(
            [pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" },
            [pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" },
            [pscustomobject]@{ Path = "C:\data\消えた"; Name = "消えた" },  # インデックスのフォルダが無い
            [pscustomobject]@{ Path = "C:\data\以前"; Name = "" }           # 以前の形式（インデックス名なし）は対象外
        )

        removeDroppedFolders $folders $previous

        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should Be $true
        Test-Path -LiteralPath "$indexDir\技術" | Should Be $false
    }

    It "インデックス名に [ ] があり、中に 260 文字を超えるパスがあっても削除する" {
        $indexDir = "$TestDrive\dropped2\index"
        $deep = "$indexDir\[旧]営業\" + ("深いフォルダ" * 20) + "\" + ("もっと深いフォルダ" * 15)
        newTsv (toLongPath "$deep\見積.xlsx\Sheet1.tsv") @("a")
        newTsv "$indexDir\[旧]営業2\見積.xlsx\Sheet1.tsv" @("b")  # 名前の先頭が同じだけの別のインデックスは残す

        removeDroppedFolders @([pscustomobject]@{ Path = "C:\data\営業2"; Name = "[旧]営業2" }) @(
            [pscustomobject]@{ Path = "C:\data\旧営業"; Name = "[旧]営業" },
            [pscustomobject]@{ Path = "C:\data\営業2"; Name = "[旧]営業2" }
        )

        [System.IO.Directory]::Exists((toLongPath "$indexDir\[旧]営業")) | Should Be $false
        Test-Path -LiteralPath "$indexDir\[旧]営業2\見積.xlsx\Sheet1.tsv" | Should Be $true
    }

    It "前回のクロール対象フォルダが無ければ何も削除しない" {
        $indexDir = "$TestDrive\dropped3\index"
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("a")

        removeDroppedFolders @() $null

        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should Be $true
    }

    It "インデックス名の大文字・小文字だけが違うフォルダは、同じフォルダとして残す" {
        $indexDir = "$TestDrive\dropped4\index"
        newTsv "$indexDir\Sales\見積.xlsx\Sheet1.tsv" @("a")

        removeDroppedFolders @([pscustomobject]@{ Path = "C:\data\sales"; Name = "sales" }) @([pscustomobject]@{ Path = "C:\data\Sales"; Name = "Sales" })

        Test-Path -LiteralPath "$indexDir\Sales\見積.xlsx\Sheet1.tsv" | Should Be $true
    }
}

Describe "migrateFlatIndex" -Tag Io {
    It "以前の形式のTSVを <ファイル名>\<場所>.tsv へ移し、今の形式・分けられない名前のものは触らない" {
        $indexDir = "$TestDrive\flat\index"
        newTsv "$indexDir\営業\A社.xlsx_Sheet1.tsv" @("new")
        newTsv "$indexDir\営業\A社.xlsx\Sheet1.tsv" @("old")         # 移し先に同じ名前があれば置き換える
        newTsv "$indexDir\営業\sub\報告.docx_ページ001.tsv" @("b")
        newTsv "$indexDir\営業\B社.xlsx\Sheet1.tsv" @("current")      # 今の形式
        newTsv "$indexDir\営業\memo_1.tsv" @("memo")                  # ファイル名と場所に分けられない

        migrateFlatIndex

        Test-Path -LiteralPath "$indexDir\営業\A社.xlsx_Sheet1.tsv" | Should Be $false
        (Get-Content -LiteralPath "$indexDir\営業\A社.xlsx\Sheet1.tsv" -Encoding UTF8) | Should Be "new"
        Test-Path -LiteralPath "$indexDir\営業\sub\報告.docx\ページ001.tsv" | Should Be $true
        Test-Path -LiteralPath "$indexDir\営業\B社.xlsx\Sheet1.tsv" | Should Be $true
        Test-Path -LiteralPath "$indexDir\営業\memo_1.tsv" | Should Be $true
    }

    It "ファイル名に [ ] ・ _ があるもの、符号化した場所のものも、元のファイル名と場所に分けて移す" {
        $indexDir = "$TestDrive\flat2\index"
        newTsv "$indexDir\[確定]見積.xlsx_Sheet1.tsv" @("a")
        newTsv "$indexDir\A_B社.xlsx_Sheet1.tsv" @("b")
        newTsv "$indexDir\資料.pptx_スライド003%5Fノート.tsv" @("c")

        migrateFlatIndex

        Test-Path -LiteralPath "$indexDir\[確定]見積.xlsx\Sheet1.tsv" | Should Be $true
        Test-Path -LiteralPath "$indexDir\A_B社.xlsx\Sheet1.tsv" | Should Be $true
        Test-Path -LiteralPath "$indexDir\資料.pptx\$(toIndexFileName "スライド003_ノート")" | Should Be $true
        @(Get-ChildItem -LiteralPath $indexDir -File).Count | Should Be 0
    }

    It "260 文字を超えるパスにある以前の形式のTSVも移す" {
        $indexDir = "$TestDrive\flat3\index"
        $deep = "$indexDir\営業\" + ("深いフォルダ" * 20) + "\" + ("もっと深いフォルダ" * 15)
        newTsv (toLongPath "$deep\見積.xlsx_Sheet1.tsv") @("a")
        try {
            migrateFlatIndex
            [System.IO.File]::Exists((toLongPath "$deep\見積.xlsx\Sheet1.tsv")) | Should Be $true
            [System.IO.File]::Exists((toLongPath "$deep\見積.xlsx_Sheet1.tsv")) | Should Be $false
        } finally {
            # TestDrive の後片付けは長いパスを消せないため、ここで消す
            removeDirectoryRetry "$TestDrive\flat3"
        }
    }

    It "移せないTSVは数えて残し、止まらない" {
        $indexDir = "$TestDrive\flat_fail\index"
        newTsv "$indexDir\C社.xlsx_Sheet1.tsv" @("a")
        # 移し先のフォルダと同じ名前のファイルがあると、フォルダを作れない
        [System.IO.File]::WriteAllText("$indexDir\C社.xlsx", "")

        { migrateFlatIndex } | Should Not Throw
        Test-Path -LiteralPath "$indexDir\C社.xlsx_Sheet1.tsv" | Should Be $true
    }

    It "インデックスのフォルダが無くても止まらない" {
        $indexDir = "$TestDrive\flat_none\index"
        { migrateFlatIndex } | Should Not Throw
    }
}
