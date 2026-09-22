# 制限モードのインデックス作成（tebunko_grep\restricted\restricted_indexer.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\office\office_text.ps1"
. "${scriptsDir}\shared\office\office_reader_clm.ps1"
. "${scriptsDir}\tebunko_grep\indexer\indexer_plan.ps1"
. "${scriptsDir}\tebunko_grep\indexer\index_migrate.ps1"
. "${scriptsDir}\tebunko_grep\restricted\restricted_indexer.ps1"

function newIndexerFixture {
    # 取り込むファイルを入れたフォルダと、work にあたるフォルダを作り、その中のパスを返す
    param (
        [string]$name
    )

    $root = Join-Path $TestDrive $name
    $source = Join-Path $root "元"
    New-Item -ItemType Directory -Path $source -Force | Out-Null
    Copy-Item -LiteralPath "${testDataDir}\office\Word\基本.docx" -Destination (Join-Path $source "文書.docx")
    Copy-Item -LiteralPath "${testDataDir}\office\PowerPoint\基本.pptx" -Destination (Join-Path $source "資料.pptx")
    # 制限モードでは取り込まないファイル（Excel・ZIP でないファイル）
    [System.IO.File]::WriteAllText((Join-Path $source "表.xlsx"), "dummy")
    [System.IO.File]::WriteAllText((Join-Path $source "中身はdoc.docx"), "not a zip")
    return @{
        Source   = $source
        Work     = (Join-Path $root "work")
        Settings = (Join-Path $root "setting.config")
    }
}

function invokeFixtureIndexing {
    # 取り込み先を fixture のフォルダに向けて invokeRestrictedIndexing を呼ぶ
    param (
        [hashtable]$fixture,
        [bool]$retryFailed = $false
    )

    ${workDir}             = $fixture.Work
    ${indexDir}            = Join-Path $fixture.Work "index"
    ${tmpDir}              = Join-Path $fixture.Work "tmp\$PID"
    ${publishDir}          = Join-Path $fixture.Work "取り込み出力\$PID"
    ${statusFile}          = Join-Path $fixture.Work "取り込み一覧.tsv"
    ${ingestingFile}       = Join-Path $fixture.Work "取り込み中.txt"
    ${settingsFile}        = $fixture.Settings
    ${restrictedLockFile}  = Join-Path $fixture.Work "インデックス作成中.lock"
    return (invokeRestrictedIndexing $retryFailed)
}

function newFixtureFolder {
    param (
        [hashtable]$fixture,
        [string]$name = "元",
        [bool]$enabled = $true
    )

    return (New-Object PSObject -Property ([ordered]@{ Name = $name; Path = $fixture.Source; Enabled = $enabled }))
}

function getFixtureStatus {
    param (
        [hashtable]$fixture
    )

    return (readStatusFile (Join-Path $fixture.Work "取り込み一覧.tsv"))
}

Describe "testRestrictedIngestable" -Tag Unit {
    It "Word・PowerPoint の新形式だけ取り込む（大文字の拡張子も）" {
        @("a\文書.docx", "b.DOCX", "c.docm", "d.pptx", "e.PPTM") | ForEach-Object {
            testRestrictedIngestable $_ | Should Be $true
        }
        @("a.xlsx", "b.xlsm", "c.doc", "d.ppt", "e.pdf", "f") | ForEach-Object {
            testRestrictedIngestable $_ | Should Be $false
        }
    }
}

Describe "readZipSignature" -Tag Io {
    It "ZIP（Office の新形式）なら真、そうでなければ偽" {
        readZipSignature "${testDataDir}\office\Word\基本.docx" | Should Be $true
        $path = Join-Path $TestDrive "sig.docx"
        [System.IO.File]::WriteAllText($path, "not a zip")
        readZipSignature $path | Should Be $false
        [System.IO.File]::WriteAllBytes($path, @())
        readZipSignature $path | Should Be $false
    }
}

Describe "getRestrictedTsvCounts" -Tag Io {
    It "フォルダごとの TSV の数を返し、0 バイトの TSV があるフォルダは壊れているものとして扱う" {
        $dir = Join-Path $TestDrive "counts"
        foreach ($rel in @("元\a.docx\ページ001.tsv", "元\a.docx\ページ002.tsv", "元\b.pptx\スライド001.tsv")) {
            $path = Join-Path $dir $rel
            New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
            Set-Content -LiteralPath $path -Value "x" -Encoding UTF8
        }
        New-Item -ItemType Directory -Path (Join-Path $dir "元\空.docx") -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $dir "元\b.pptx\スライド002.tsv") -Value ([byte[]]@()) -Encoding Byte
        # インデックスのフォルダの直下の TSV（以前の形式）は数えない
        Set-Content -LiteralPath (Join-Path $dir "直下.tsv") -Value "x" -Encoding UTF8

        $counts = getRestrictedTsvCounts $dir
        $counts["元\a.docx"] | Should Be 2
        $counts["元\A.DOCX"] | Should Be 2  # 大文字・小文字を区別しない
        $counts["元\b.pptx"] | Should Be ${indexBrokenCount}
        $counts["元\空.docx"] | Should Be 0
        $counts.ContainsKey("元\無い.docx") | Should Be $false
    }

    It "インデックスのフォルダが無ければ空を返す" {
        (getRestrictedTsvCounts (Join-Path $TestDrive "無い")).Count | Should Be 0
    }
}

Describe "enterIndexingLock / exitIndexingLock" -Tag Io {
    It "実行中のインデックス作成があれば、その理由を返す" {
        $path = Join-Path $TestDrive "lock1.lock"
        enterIndexingLock $path | Should Be ""
        @(Get-Content -LiteralPath $path)[0].Split("`t")[0] | Should Be ([string]$PID)
        $reason = enterIndexingLock $path
        $reason | Should Match "ほかのインデックス作成が実行中です"
        $reason | Should Match ([regex]::Escape($path))
        exitIndexingLock $path
        Test-Path -LiteralPath $path | Should Be $false
    }

    It "強制終了で残ったロック（終わっているプロセス・壊れた中身）は取り直す" {
        $path = Join-Path $TestDrive "lock2.lock"
        # 使われていないプロセス ID を探す
        $dead = 99999
        while (Get-Process -Id $dead -ErrorAction SilentlyContinue) { $dead-- }
        Set-Content -LiteralPath $path -Value "$dead`t$($env:COMPUTERNAME)`t2026/09/20 10:00:00" -Encoding UTF8
        enterIndexingLock $path | Should Be ""
        Set-Content -LiteralPath $path -Value "こわれた行" -Encoding UTF8
        enterIndexingLock $path | Should Be ""
        exitIndexingLock $path
    }

    It "ほかの PC のロックは、そのままでは取れない" {
        $path = Join-Path $TestDrive "lock3.lock"
        Set-Content -LiteralPath $path -Value "1234`tPC-OTHER`t2026/09/20 10:00:00" -Encoding UTF8
        enterIndexingLock $path | Should Match "PC-OTHER"
        exitIndexingLock $path
    }
}

Describe "invokeRestrictedIndexing" -Tag Io {
    It "クロール対象フォルダが無ければ、続けられないエラーにする" {
        $fixture = newIndexerFixture "idx_none"
        ${settingsFile} = $fixture.Settings
        { invokeFixtureIndexing $fixture } | Should Throw "クロール対象フォルダがありません"
    }

    It "Word・PowerPoint を取り込み、Excel・ZIP でないファイルは未取り込みのまま残す" {
        $fixture = newIndexerFixture "idx_run"
        writeTargetFolders @((New-Object PSObject -Property ([ordered]@{ Name = "元"; Path = $fixture.Source; Enabled = $true }))) $fixture.Settings

        $result = invokeFixtureIndexing $fixture
        $result.Success | Should Be 2
        @($result.Failed).Count | Should Be 0
        @($result.Skipped).Count | Should Be 2
        $result.Targets | Should Be 4

        # いつものインデクサと同じ形（work\index\<インデックス名>\<相対パス>\<場所>.tsv）に作る
        $indexDir = Join-Path $fixture.Work "index"
        Test-Path -LiteralPath (Join-Path $indexDir "元\文書.docx\ページ001.tsv") | Should Be $true
        Test-Path -LiteralPath (Join-Path $indexDir "元\資料.pptx\スライド001.tsv") | Should Be $true
        Test-Path -LiteralPath (Join-Path $indexDir "元\表.xlsx") | Should Be $false

        $status = getFixtureStatus $fixture
        $status.Rows["元\文書.docx"].状態 | Should Be ${stateDone}
        $status.Rows["元\文書.docx"].TSV数 | Should Be "3"
        $status.Rows["元\文書.docx"].抽出版 | Should Be "2"
        $status.Rows["元\資料.pptx"].状態 | Should Be ${stateDone}
        # 取り込めないファイルは「失敗」にしない（いつもの画面が使える PC で取り込めるようにするため）
        $status.Rows["元\表.xlsx"].状態 | Should Be ${stateNew}
        $status.Rows["元\中身はdoc.docx"].状態 | Should Be ${stateNew}
        # クロール対象フォルダとインデックス名を記録する
        $status.Folders[0].Name | Should Be "元"
        Test-Path -LiteralPath (Join-Path $indexDir "元\元のフォルダ.txt") | Should Be $true
        # 作業フォルダとロックは残さない
        Test-Path -LiteralPath (Join-Path $fixture.Work "インデックス作成中.lock") | Should Be $false
        Test-Path -LiteralPath (Join-Path $fixture.Work "tmp\$PID") | Should Be $false
    }

    It "2 回目は取り込み済みのファイルを取り込み直さない" {
        $fixture = newIndexerFixture "idx_again"
        writeTargetFolders @((New-Object PSObject -Property ([ordered]@{ Name = "元"; Path = $fixture.Source; Enabled = $true }))) $fixture.Settings
        invokeFixtureIndexing $fixture | Out-Null

        $result = invokeFixtureIndexing $fixture
        $result.Success | Should Be 0
        # 取り込めないファイルは未取り込みのままのため、毎回対象には挙がる（読まずに飛ばす）
        @($result.Skipped).Count | Should Be 2
        (getFixtureStatus $fixture).Rows["元\文書.docx"].状態 | Should Be ${stateDone}
    }

    It "インデックス（TSV）を直接削除されたファイルは、取り込み一覧が「済」でも取り込み直す" {
        $fixture = newIndexerFixture "idx_lost"
        writeTargetFolders @((newFixtureFolder $fixture)) $fixture.Settings
        invokeFixtureIndexing $fixture | Out-Null
        Remove-Item -LiteralPath (Join-Path $fixture.Work "index\元\資料.pptx") -Recurse -Force

        (invokeFixtureIndexing $fixture).Success | Should Be 1
        Test-Path -LiteralPath (Join-Path $fixture.Work "index\元\資料.pptx\スライド001.tsv") | Should Be $true
    }

    It "元のファイルが無くなったら、そのインデックスと記録を消す" {
        $fixture = newIndexerFixture "idx_dropped"
        writeTargetFolders @((New-Object PSObject -Property ([ordered]@{ Name = "元"; Path = $fixture.Source; Enabled = $true }))) $fixture.Settings
        invokeFixtureIndexing $fixture | Out-Null
        Remove-Item -LiteralPath (Join-Path $fixture.Source "文書.docx") -Force

        $result = invokeFixtureIndexing $fixture
        $result.Dropped | Should Be 0  # クロールの時点で一覧から外れる（取り込み対象にも挙がらない）
        Test-Path -LiteralPath (Join-Path $fixture.Work "index\元\文書.docx") | Should Be $false
        (getFixtureStatus $fixture).Rows.ContainsKey("元\文書.docx") | Should Be $false
    }

    It "チェックを外したフォルダは取り込まず、インデックスも記録も残す" {
        $fixture = newIndexerFixture "idx_unchecked"
        writeTargetFolders @((New-Object PSObject -Property ([ordered]@{ Name = "元"; Path = $fixture.Source; Enabled = $true }))) $fixture.Settings
        invokeFixtureIndexing $fixture | Out-Null
        writeTargetFolders @((New-Object PSObject -Property ([ordered]@{ Name = "元"; Path = $fixture.Source; Enabled = $false }))) $fixture.Settings

        { invokeFixtureIndexing $fixture } | Should Throw "チェックの付いたクロール対象フォルダがありません"
        Test-Path -LiteralPath (Join-Path $fixture.Work "index\元\文書.docx\ページ001.tsv") | Should Be $true
        (getFixtureStatus $fixture).Rows["元\文書.docx"].状態 | Should Be ${stateDone}
    }

    It "以前の形式のインデックスが残っていれば、いつもの画面で作り直してもらう" {
        $fixture = newIndexerFixture "idx_legacy"
        writeTargetFolders @((New-Object PSObject -Property ([ordered]@{ Name = "元"; Path = $fixture.Source; Enabled = $true }))) $fixture.Settings
        New-Item -ItemType Directory -Path $fixture.Work -Force | Out-Null
        # インデックス名の無いクロール対象フォルダの行（以前の形式）
        $status = Join-Path $fixture.Work "取り込み一覧.tsv"
        Set-Content -LiteralPath $status -Value @("クロール対象フォルダ`t$($fixture.Source)", "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー`t抽出版") -Encoding UTF8
        { invokeFixtureIndexing $fixture } | Should Throw "以前の版の形式のインデックスが残っています"
    }

    It "チェックの無いフォルダ・見つからないフォルダの行は、取り込まずに一覧に残す" {
        $fixture = newIndexerFixture "idx_mixed"
        $other = Join-Path $TestDrive "idx_mixed\別"
        New-Item -ItemType Directory -Path $other -Force | Out-Null
        Copy-Item -LiteralPath "${testDataDir}\office\Word\基本.docx" -Destination (Join-Path $other "別文書.docx")
        # インデックス名を空にして渡し、取り込み時に割り当てさせる（assignIndexNames → writeTargetFolders）
        $folders = @(
            (New-Object PSObject -Property ([ordered]@{ Name = ""; Path = $fixture.Source; Enabled = $true })),
            (New-Object PSObject -Property ([ordered]@{ Name = ""; Path = $other; Enabled = $true })),
            (New-Object PSObject -Property ([ordered]@{ Name = "無い"; Path = (Join-Path $TestDrive "idx_mixed\無い"); Enabled = $true }))
        )
        writeTargetFolders $folders $fixture.Settings
        invokeFixtureIndexing $fixture | Out-Null
        # 名前が割り当てられて設定に書かれる
        @(getTargetFolders $fixture.Settings | ForEach-Object { $_.Name }) -join "," | Should Be "元,別,無い"

        # 2 回目は 別 のチェックを外す（インデックスも一覧の行も残す）
        $folders = @(getTargetFolders $fixture.Settings)
        $folders[1].Enabled = $false
        writeTargetFolders $folders $fixture.Settings
        invokeFixtureIndexing $fixture | Out-Null
        $status = getFixtureStatus $fixture
        $status.Rows["別\別文書.docx"].状態 | Should Be ${stateDone}
        Test-Path -LiteralPath (Join-Path $fixture.Work "index\別\別文書.docx\ページ001.tsv") | Should Be $true
    }

    It "取り込めなかったファイルは失敗にし、-RetryFailed で取り込み直す" {
        $fixture = newIndexerFixture "idx_failed"
        # ZIP だが中身が docx でないファイル（読み取りで失敗する）
        Copy-Item -LiteralPath "${testDataDir}\office\PowerPoint\基本.pptx" -Destination (Join-Path $fixture.Source "中身はpptx.docx")
        writeTargetFolders @((newFixtureFolder $fixture)) $fixture.Settings

        $result = invokeFixtureIndexing $fixture
        @($result.Failed).Count | Should Be 1
        $result.Failed[0].RelPath | Should Be "元\中身はpptx.docx"
        $row = (getFixtureStatus $fixture).Rows["元\中身はpptx.docx"]
        $row.状態 | Should Be ${stateFailed}
        $row.エラー | Should Not BeNullOrEmpty
        # 更新が無ければ、次からは取り込み対象にしない
        (invokeFixtureIndexing $fixture).Targets | Should Be 2
        # -RetryFailed では取り込み直す（また失敗する）
        $retried = invokeFixtureIndexing $fixture $true
        @($retried.Failed).Count | Should Be 1
    }

    It "取り込み中に強制終了されたファイルは最後に回し、続けて止まったら失敗にする" {
        $fixture = newIndexerFixture "idx_interrupted"
        writeTargetFolders @((newFixtureFolder $fixture)) $fixture.Settings
        $ingesting = Join-Path $fixture.Work "取り込み中.txt"
        New-Item -ItemType Directory -Path $fixture.Work -Force | Out-Null

        # 1 回目の強制終了: 最後に回して取り込む（印は消える）
        writeIngestingFile "元\文書.docx" 1 $ingesting
        (invokeFixtureIndexing $fixture).Success | Should Be 2
        Test-Path -LiteralPath $ingesting | Should Be $false

        # 2 回続けて強制終了したファイルは取り込まずに失敗にする
        writeIngestingFile "元\資料.pptx" 2 $ingesting
        Remove-Item -LiteralPath (Join-Path $fixture.Work "index\元\資料.pptx") -Recurse -Force
        $result = invokeFixtureIndexing $fixture
        $result.Targets | Should Be 2  # 資料.pptx は対象から外れ、取り込めない 2 件（Excel・ZIP でないファイル）が残る
        $row = (getFixtureStatus $fixture).Rows["元\資料.pptx"]
        $row.状態 | Should Be ${stateFailed}
        $row.エラー | Should Match "2 回続けて強制終了された"
        Test-Path -LiteralPath $ingesting | Should Be $false

        # 取り込み対象に無いファイルの印は、そのまま消す
        writeIngestingFile "元\無い.docx" 1 $ingesting
        invokeFixtureIndexing $fixture | Out-Null
        Test-Path -LiteralPath $ingesting | Should Be $false
    }

    It "ほかのインデックス作成が実行中なら始めない" {
        $fixture = newIndexerFixture "idx_locked"
        writeTargetFolders @((New-Object PSObject -Property ([ordered]@{ Name = "元"; Path = $fixture.Source; Enabled = $true }))) $fixture.Settings
        New-Item -ItemType Directory -Path $fixture.Work -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $fixture.Work "インデックス作成中.lock") -Value "1234`tPC-OTHER`t2026/09/20 10:00:00" -Encoding UTF8
        { invokeFixtureIndexing $fixture } | Should Throw "ほかのインデックス作成が実行中です"
    }
}
