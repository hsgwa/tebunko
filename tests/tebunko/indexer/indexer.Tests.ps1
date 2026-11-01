# インデクサの起動口（tebunko\indexer.ps1）と本体（indexer\indexer_run.ps1 の invokeIndexer）のテスト。
# indexer.ps1 は indexer_lib.ps1（lib.ps1 を含む）を読み込み、置き場所（リポジトリ直下の setting.config・work\）を決める。
# リポジトリの設定・インデックスを書き換えないよう、settings.ps1 で ${settingsFile} を決める直前に止めて、
# ツールのフォルダ（${rootDir}）をテスト用のフォルダ（TestDrive）に差し替えてから続けさせる（Set-PSBreakpoint の -Action）。
# 同じやり方で、取り込みの途中に中止・画面の返事・元のファイルの削除を起こす。
# 画面とのやり取りは受け渡しの口（newIndexerChannel）で行う。途中に割り込むテストは、取り込みのスレッドを使わない
# （Workers = 0。ブレークポイントはテストと同じスレッドでしか止まらないため）。取り込みのスレッドを使う場合は別に確かめる。
# Excel は COM が要るため使わない。Word・PowerPoint の新形式（.docx・.pptx）はファイルを直接読むため、そのまま取り込む。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    # findLine・newRoot・writeTestSettings・runIndexer・readTestStatus・readTestError・readTestProgress は
    # index_compat.Tests.ps1 と共通のため tests\helpers\indexer.ps1 にある（中身は変えていない）
    . "$PSScriptRoot\..\..\helpers\indexer.ps1"

    # ジャーナルを読むため（クラスはこのファイルの外の関数から見えないので、読み込んでおく）
    . "${scriptsDir}\tebunko\indexer\ingest_journal.ps1"

    $planPath    ="${scriptsDir}\tebunko\indexer\indexer_plan.ps1"
    $reporterPath = "${scriptsDir}\tebunko\indexer\indexing_reporter.ps1"
    $docxSource  = "${testDataDir}\office\Word\形式\大文字拡張子.DOCX"
    $pptxSource  = "${testDataDir}\office\PowerPoint\基本.pptx"

    function newSourceFolder {
        # 取り込むフォルダを作り、Word・PowerPoint のファイルと、壊れた PowerPoint のファイルを置く
        param ([string]$name)
        $dir = Join-Path $TestDrive $name
        [System.IO.Directory]::CreateDirectory("$dir\資料") | Out-Null
        Copy-Item -LiteralPath $docxSource -Destination "$dir\議事録.docx"
        Copy-Item -LiteralPath $pptxSource -Destination "$dir\資料\提案.pptx"
        [System.IO.File]::WriteAllText("$dir\壊れた.pptx", "PowerPoint ではない内容")
        return $dir
    }

    function writeTestJournal {
        # 前回の取り込みが残したジャーナル（work\store\journal\ingest.tsv）を、版と回の行に続けて作る（取り込み一覧に行が無い回）
        param ([string]$root, [string[]]$records, [int]$rows = 0)
        $path = "$root\work\store\journal\ingest.tsv"
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
        $lines = @("版`t2", "回`t2026/10/17 10:00:00`t$rows") + $records
        [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
        return $path
    }

    function assertGramMatchesFullScan {
        # インデックス name の gram 索引（目録）があり、本文から取った語の検索結果が、gram 索引で絞ったときと全件を照合したときで同じになる
        param ([string]$root, [string]$name)
        $index = "$root\work\store\content"
        $gram = "$root\work\store\gram"
        [System.IO.File]::Exists("$gram\$name\toc.tsv") | Should -BeTrue
        (readGramToc "$gram\$name").Segments.Count | Should -BeGreaterThan 0
        $all = getContentIndexFiles @($index)
        $compared = 0
        $narrowedWords = 0
        foreach ($item in @($all.ContentIndexFiles)) {
            $body = getContentIndexBodyText (readContentIndexText $item.Path)
            foreach ($word in @(($body -split "\s+") | Where-Object { $_.Length -ge 2 } | Select-Object -First 3)) {
                $narrowed = getGramContentIndexFiles $word $true @($index) $index $gram
                if ($null -ne $narrowed) {
                    # 絞れたとき（$null でない）は、1 つ以上のファイルを外している
                    @($narrowed.ContentIndexFiles).Count | Should -BeLessThan @($all.ContentIndexFiles).Count
                    $narrowedWords++
                }
                $files = @(if ($null -eq $narrowed) { $all.ContentIndexFiles } else { $narrowed.ContentIndexFiles })
                $viaGram = searchContentIndex $word $files $true 0
                $full = searchContentIndex $word $all.ContentIndexFiles $true 0
                (@($viaGram.Hits | ForEach-Object { "$($_.RelPath)|$($_.Book)|$($_.Location)|$($_.LineNumber)" }) -join "`n") | Should -Be (@($full.Hits | ForEach-Object { "$($_.RelPath)|$($_.Book)|$($_.Location)|$($_.LineNumber)" }) -join "`n")
                $full.Hits.Count | Should -BeGreaterThan 0
                $compared++
            }
        }
        $compared | Should -BeGreaterThan 0
        $narrowedWords | Should -BeGreaterThan 0
    }
}

Describe "indexer.ps1（続けられないエラー）" -Tag Io {
    # folders の path は TestDrive からの相対パス（表を作る探索のときは TestDrive が無いため、テストの中で組み立てる）
    It "<name>" -TestCases @(
        @{ name = "クロール対象フォルダが無ければ、エラーを書いて 1 で終わる"; folders = @(); message = "クロール対象フォルダがありません" }
        @{ name = "チェックの付いたフォルダが無ければ、エラーを書いて 1 で終わる"; folders = @(@{ name = "営業"; path = "無し"; enabled = $false })
           message = "チェックの付いたクロール対象フォルダがありません" }
    ) {
        param ($name, $folders, $message)
        $folders = @($folders | ForEach-Object { @{ name = $_.name; path = (Join-Path $TestDrive $_.path); enabled = $_.enabled } })
        $root = newRoot
        writeTestSettings $root $folders
        runIndexer $root | Should -Be 1
        readTestError | Should -Match $message
    }

    It "受け渡しの口を渡さなければ（コンソールから実行したとき）、自分で口を作って動かし、終了コードで終わる" {
        $root = newRoot
        writeTestSettings $root @()
        $global:indexerTestRoot = $root
        $point = Set-PSBreakpoint -Script $settingsScriptPath -Line (findLine $settingsScriptPath '^\$\{settingsFile\}\s*=') -Action {
            Set-Variable -Name rootDir -Value $global:indexerTestRoot -Scope 1
        }
        try {
            $global:LASTEXITCODE = 0
            & $indexerPath *> $null
            $LASTEXITCODE | Should -Be 1
        } finally {
            Remove-PSBreakpoint -Breakpoint $point
            Remove-Variable -Name indexerTestRoot -Scope Global -ErrorAction SilentlyContinue
        }
        [System.IO.File]::ReadAllText("$root\work\indexing_log.txt") | Should -Match "クロール対象フォルダがありません"
    }

    It "画面なしで壊れた setting.config のまま起動すると、退避して既定の設定で動き、ログに知らせを残して 1 で終わる" {
        $root = newRoot
        [System.IO.File]::WriteAllText("$root\setting.config", '{ "targetFolders": [', ${utf8Bom})
        # 退避すると既定のワークスペースになるため、利用者の既定のワークスペース（tebunko_ws） に書かないよう、既定の場所もテスト用に差し替える
        # （環境変数 TEBUNKO_DEFAULT_WORKSPACE。テストの間だけ入れ替え、終わったら元に戻す）
        $global:indexerTestRoot = $root
        $points = @(
            (Set-PSBreakpoint -Script $settingsScriptPath -Line (findLine $settingsScriptPath '^\$\{settingsFile\}\s*=') -Action {
                Set-Variable -Name rootDir -Value $global:indexerTestRoot -Scope 1
            })
        )
        $previousWorkspace = $env:TEBUNKO_DEFAULT_WORKSPACE
        $env:TEBUNKO_DEFAULT_WORKSPACE = "$root\profile\Documents\tebunko_ws"
        try {
            $global:LASTEXITCODE = 0
            & $indexerPath *> $null
            $LASTEXITCODE | Should -Be 1
        } finally {
            $env:TEBUNKO_DEFAULT_WORKSPACE = $previousWorkspace
            foreach ($point in $points) { Remove-PSBreakpoint -Breakpoint $point }
            Remove-Variable -Name indexerTestRoot -Scope Global -ErrorAction SilentlyContinue
        }
        $broken = @(Get-ChildItem -LiteralPath $root -Filter "setting.config.broken-*")
        $broken.Count | Should -Be 1
        [System.IO.File]::ReadAllText($broken[0].FullName, [System.Text.Encoding]::UTF8) | Should -Be '{ "targetFolders": ['
        $log = [System.IO.File]::ReadAllText("$root\profile\Documents\tebunko_ws\indexing_log.txt")
        $log | Should -Match "設定ファイルが壊れていたため、既定の設定で起動しました"
        $log | Should -Match ([regex]::Escape($broken[0].FullName))
        $log | Should -Match "クロール対象フォルダがありません"
    }

    It "ほかのインデックス作成が動いていれば（<name>）、そのインデックス作成のログに触らずに 1 で終わる" -TestCases @(
        @{ name = "同じセッションのミューテックス"; kind = "mutex" }
        @{ name = "別のセッション・PC・パスの書き方の代わり（ミューテックスは持たず、鍵のファイルだけ）"; kind = "file" }
    ) {
        param ($name, $kind)
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $TestDrive; enabled = $true })
        # 実行中のインデックス作成が画面とやり取りしているファイル
        $files = @("indexing_log.txt")
        foreach ($fileName in $files) {
            [System.IO.File]::WriteAllText("$root\work\$fileName", "実行中")
        }
        # ほかの PC が取り込み中の出力用のフォルダ（PID はこの PC で動いていないもの）。片付けで消さない
        $foreign = "$root\work\publish\999999\import"
        [System.IO.Directory]::CreateDirectory($foreign) | Out-Null
        if ($kind -eq "mutex") {
            $other = newAppMutex "indexer" "$root\work" "tebunko"
        } else {
            $other = openLockFile "$root\work\indexing.lock" 1
        }
        try {
            runIndexer $root | Should -Be 1
        } finally {
            if ($kind -eq "mutex") {
                $other.Mutex.Dispose()
            } else {
                $other.Dispose()
            }
        }
        foreach ($fileName in $files) {
            [System.IO.File]::ReadAllText("$root\work\$fileName") | Should -Be "実行中"
        }
        [System.IO.Directory]::Exists($foreign) | Should -Be $true
        # 鍵が取れなかったときは、前回の分へ移さない
        [System.IO.File]::Exists("$root\work\indexing_log_previous.txt") | Should -Be $false
        readTestError | Should -Match "ほかのインデックス作成が実行中です"
    }
}

Describe "indexer.ps1（前の版が作った 64MB を超える本文インデックスのファイル）" -Tag Slow {
    # 別のプロセスの上限は下げられないため、実際に 64MB を超えるファイルを作る（中身は流して書く）。毎回のテストでは流さない

    It "同じ拡張子の小さい本文インデックスのファイルが残るフォルダでも、取り込む対象が無い更新で外れ、中の元のファイルが同じ回に取り込み直されて検索に戻る" {
        $root = newRoot
        $dir = newSourceFolder "上限超え"
        writeTestSettings $root @(@{ name = ""; path = $dir; enabled = $true })
        runIndexer $root | Should -Be 0
        $folder = "$root\work\store\content\上限超え"
        $file = "$folder\content_index.docx.001.tsv"
        [System.IO.File]::Exists($file) | Should -Be $true

        # 前の版が作った形: 議事録.docx を入れた、検索の上限（64MB）を超えるファイルと、別の docx を入れた小さいファイル
        writeContentIndexFile $file (convertToContentIndexText @(@{ Name = "議事録.docx"; Places = @(@{ Place = "ページ001"; Text = "先頭`n" }) }))
        $stream = New-Object System.IO.FileStream($file, [System.IO.FileMode]::Append)
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UnicodeEncoding($false, $false)))
        try {
            $line = ("あ" * 999) + "`n"
            for ($i = 0; $i -lt 36000; $i++) { $writer.Write($line) }
        } finally {
            $writer.Dispose()
        }
        writeContentIndexFile "$folder\content_index.docx.002.tsv" (convertToContentIndexText @(@{ Name = "別.docx"; Places = @(@{ Place = "ページ001"; Text = "別の文書`n" }) }))
        (Get-Item -LiteralPath $file).Length | Should -BeGreaterThan ${contentIndexSearchMaxBytes}

        runIndexer $root | Should -Be 0

        $status = readTestStatus $root
        $status.Rows["上限超え\議事録.docx"].状態 | Should -Be ${stateDone}
        foreach ($item in @(Get-ChildItem -LiteralPath $folder -Filter "content_index.docx.*.tsv" -File)) {
            $item.Length | Should -BeLessThan 1MB
        }
        $names = @(Get-ChildItem -LiteralPath $folder -Filter "content_index.docx.*.tsv" -File | ForEach-Object { splitContentIndexTextByBook (readContentIndexText $_.FullName) } | ForEach-Object { $_.Name })
        $names | Should -Contain "議事録.docx"
        $names | Should -Contain "別.docx"
        readTestError | Should -BeNullOrEmpty
    }
}

Describe "indexer.ps1（取り込み）" -Tag Io {
    BeforeAll {
        $source = newSourceFolder "営業"
        $unchecked = newSourceFolder "経理"
        $missing = Join-Path $TestDrive "無くなったフォルダ"
    }

    It "チェックの付いたフォルダを取り込み、名前を設定に保存し、失敗したファイルを一覧に残す" {
        $root = newRoot
        writeTestSettings $root @(
            @{ name = ""; path = $source; enabled = $true },
            @{ name = ""; path = $unchecked; enabled = $false },
            @{ name = ""; path = $missing; enabled = $true })

        runIndexer $root | Should -Be 0

        $status = readTestStatus $root
        $status.Rows.Count | Should -Be 3
        $status.Rows["営業\議事録.docx"].状態 | Should -Be ${stateDone}
        $status.Rows["営業\資料\提案.pptx"].状態 | Should -Be ${stateDone}
        $status.Rows["営業\壊れた.pptx"].状態 | Should -Be ${stateFailed}
        $status.Rows["営業\壊れた.pptx"].エラー | Should -Match "PowerPoint"
        # 取り込んだ TSV はフォルダの本文インデックスのファイルに入れ、元のファイルごとのフォルダは残さない
        [System.IO.File]::Exists("$root\work\store\content\営業\content_index.docx.001.tsv") | Should -Be $true
        [System.IO.Directory]::Exists("$root\work\store\content\営業\議事録.docx") | Should -Be $false
        Test-Path -LiteralPath "$root\work\store\content\営業\source_folder.txt" | Should -Be $true
        # 名前の無かったフォルダには名前を割り当てて保存する
        @(getTargetFolders "$root\setting.config" | ForEach-Object { $_.Name }) -join "," | Should -Be "営業,経理,無くなったフォルダ"
        # 画面が終わり方を読めるよう、最後の進み具合を残す
        $progress = readTestProgress
        $progress.Phase | Should -Be ${indexingPhaseFinish}
        $progress.Failed | Should -Be 1
        readTestError | Should -BeNullOrEmpty
        Test-Path -LiteralPath "$root\work\ingesting.txt" | Should -Be $false
        # 正常に終わった回のジャーナルは「終」で終わる
        @([IngestJournal]::new("$root\work\store\journal\ingest.tsv").ReadPrevious())[-1] | Should -Be "終"
        # 取り込みの作業フォルダ（work\tmp\<PC の鍵>\<PID>）も削除する
        Test-Path -LiteralPath (Join-Path "$root\work\tmp\$(getMachineKey)" "$PID") | Should -Be $false
        # 取り込み中は、前の版の片付けだけに使う場所（%TEMP%）ではなくワークスペースの tmp の下を使う
        $global:capturedTmpDir | Should -Not -BeNullOrEmpty
        $global:capturedTmpDir.StartsWith("$root\work\tmp", [System.StringComparison]::OrdinalIgnoreCase) | Should -Be $true
        # ${legacyTmpParent} の差し替えが取り込み中も効いていた（本物の %TEMP%\tebunko を指していない）ことを確かめる
        $global:capturedLegacyTmpParent | Should -Be (Join-Path $root "legacy_tmp")
        Test-Path -LiteralPath (Join-Path "$root\legacy_tmp" "$PID") | Should -Be $false
        # 検索を絞る gram 索引ができていて、絞った検索の結果が全件を照合した結果と同じになる
        assertGramMatchesFullScan $root "営業"
        # ワークスペースの直下・content_index\<名前>\ の直下のファイル名はすべて英語にする（再発防止）
        foreach ($entry in @(Get-ChildItem -LiteralPath "$root\work")) {
            $entry.Name | Should -Match "^[\x20-\x7E]+$"
        }
        foreach ($entry in @(Get-ChildItem -LiteralPath "$root\work\store\content\営業" -File)) {
            $entry.Name | Should -Match "^[\x20-\x7E]+$"
        }
    }

    It "subfolders を false にして更新すると、サブフォルダのファイルが取り込み一覧と本文インデックスから外れ、true に戻すと取り込み直す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        runIndexer $root | Should -Be 0
        (readTestStatus $root).Rows.ContainsKey("営業\資料\提案.pptx") | Should -Be $true
        $contentText = { (@(Get-ChildItem -LiteralPath "$root\work\store\content\営業" -Recurse -Filter "content_index.*.tsv" | ForEach-Object { Get-Content -Raw -LiteralPath $_.FullName }) -join "") }
        # 提案.pptx の中の語（TC21-01）が、サブフォルダ側のファイルの目印
        (& $contentText) | Should -Match "TC21-01"

        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true; subfolders = $false })
        runIndexer $root | Should -Be 0
        $rows = (readTestStatus $root).Rows
        $rows.ContainsKey("営業\資料\提案.pptx") | Should -Be $false
        $rows["営業\議事録.docx"].状態 | Should -Be ${stateDone}
        (& $contentText) | Should -Match "議事録"
        (& $contentText) | Should -Not -Match "TC21-01"

        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true; subfolders = $true })
        runIndexer $root | Should -Be 0
        (readTestStatus $root).Rows["営業\資料\提案.pptx"].状態 | Should -Be ${stateDone}
        (& $contentText) | Should -Match "TC21-01"
    }

    It "作成中は、取り込み一覧に 1 件ごとに結果が追記され、画面が読む「取り込み済み＋失敗」の件数が途中で増える" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        $global:testRowTotals = New-Object System.Collections.ArrayList
        # 1 件を追記する直前の、取り込み済みと失敗の件数を控える（始めは対象の行が「未取り込み」で並んでいる）
        $watch = @{ Script = $runPath; Pattern = '^\s+\$ledger\.AddRow\(\$row\)'; Action = { $s = getIndexingState; [void]$global:testRowTotals.Add($s.Done + $s.Failed) } }
        try {
            runIndexer $root @{} @($watch) | Should -Be 0
            @($global:testRowTotals) -join "," | Should -Be "0,1,2"
        } finally {
            Remove-Variable -Name testRowTotals -Scope Global -ErrorAction SilentlyContinue
        }
        (readTestStatus $root).Rows.Count | Should -Be 3
    }

    It "2 回目は更新の無いファイルを取り込まず、前回失敗したファイルもスキップする" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        runIndexer $root | Should -Be 0
        $first = (readTestStatus $root).Rows["営業\壊れた.pptx"].取り込み日時
        $firstLog = [System.IO.File]::ReadAllText("$root\work\indexing_log.txt")
        $firstLog | Should -Match "^﻿?==== tebunko の診断情報 ===="
        [System.IO.File]::Exists("$root\work\indexing_log_previous.txt") | Should -Be $false

        runIndexer $root | Should -Be 0
        # 前回の記録を 1 世代残し、今回の記録の先頭にも版と環境が入る
        [System.IO.File]::ReadAllText("$root\work\indexing_log_previous.txt") | Should -Be $firstLog
        [System.IO.File]::ReadAllText("$root\work\indexing_log.txt") | Should -Match "^﻿?==== tebunko の診断情報 ===="
        $progress = readTestProgress
        $progress.Detail | Should -Be "更新が必要なファイルはありませんでした"
        $status = readTestStatus $root
        $status.Rows.Count | Should -Be 3
        $status.Rows["営業\壊れた.pptx"].状態 | Should -Be ${stateFailed}
        $status.Rows["営業\壊れた.pptx"].取り込み日時 | Should -Be $first
    }

    It "フォルダが見つからないとき（ネットワークの切断など）は、前回の結果とインデックスを残す" {
        $gone = newSourceFolder "一時"
        $root = newRoot
        writeTestSettings $root @(@{ name = "一時"; path = $gone; enabled = $true })
        runIndexer $root | Should -Be 0
        Remove-Item -LiteralPath $gone -Recurse -Force

        runIndexer $root | Should -Be 0
        $status = readTestStatus $root
        $status.Rows.Count | Should -Be 3
        $status.Rows["一時\議事録.docx"].状態 | Should -Be ${stateDone}
        [System.IO.File]::Exists("$root\work\store\content\一時\content_index.docx.001.tsv") | Should -Be $true
    }

    It "ワークスペースのパスに [ ] があれば、テキストファイルを含むすべての取り込みを失敗として記録し、%TEMP% には書き込まない" {
        # テキストファイルも中間 TSV を $tmpDir に作るため、docx・pptx と同じくスキップの対象になることを確かめる
        $bracketSource = newSourceFolder "角かっこ用"
        [System.IO.File]::WriteAllText("$bracketSource\メモ.txt", "テキストファイルの内容", [System.Text.Encoding]::UTF8)
        $root = Join-Path $TestDrive "[共有]tool"
        [System.IO.Directory]::CreateDirectory("$root\work") | Out-Null
        writeTestSettings $root @(@{ name = ""; path = $bracketSource; enabled = $true })

        runIndexer $root | Should -Be 0

        $status = readTestStatus $root
        $status.Rows.Count | Should -Be 4
        foreach ($row in $status.Rows.Values) {
            $row.状態 | Should -Be ${stateFailed}
            $row.エラー | Should -Match "\[ \]"
        }
        $progress = readTestProgress
        $progress.Failed | Should -Be 4
        # 取り込み中も、置けない代わりに %TEMP% には書き込まない
        $global:capturedTmpDir | Should -BeNullOrEmpty
        Test-Path -LiteralPath (Join-Path "$root\legacy_tmp" "$PID") | Should -Be $false
        (Get-Content -LiteralPath "$root\work\indexing_log.txt" -Raw) | Should -Match "スキップ"
    }

    It "設定で取り込まないにしたテキストの拡張子は取り込まず、前に取り込んだものはインデックスから外す（戻すと取り込む）" {
$textSource = Join-Path $TestDrive "取り込まない拡張子"
[System.IO.Directory]::CreateDirectory($textSource) | Out-Null
[System.IO.File]::WriteAllText("$textSource\メモ.txt", "メモの内容", [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText("$textSource\記録.log", "記録の内容", [System.Text.Encoding]::UTF8)
$root = Join-Path $TestDrive "取り込まない拡張子用"
[System.IO.Directory]::CreateDirectory("$root\work") | Out-Null
$folders = @(@{ name = "資料"; path = $textSource; enabled = $true })
$settingsPath = "$root\setting.config"

# .log を取り込まない: メモ.txt だけ取り込む
writeTestSettings $root $folders
writeIndexTextExcludes "資料" @(".log") $settingsPath
runIndexer $root | Should -Be 0
@((readTestStatus $root).Rows.Keys | ForEach-Object { [System.IO.Path]::GetFileName($_) }) | Should -Be @("メモ.txt")

# 取り込むに戻す: .log を新しく取り込む
writeIndexTextExcludes "資料" @() $settingsPath
runIndexer $root | Should -Be 0
@((readTestStatus $root).Rows.Keys | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Sort-Object).Count | Should -Be 2

# また取り込まないにする: 取り込み一覧とインデックスから外れ、ログが「取り込まないにした拡張子」になる
writeIndexTextExcludes "資料" @(".log") $settingsPath
runIndexer $root | Should -Be 0
@((readTestStatus $root).Rows.Keys | ForEach-Object { [System.IO.Path]::GetFileName($_) }) | Should -Be @("メモ.txt")
(Get-Content -LiteralPath "$root\work\indexing_log.txt" -Raw) | Should -Match "対象外にした拡張子のファイル 1 件は、インデックスから除きます"
}


It "前回取り込み中に強制終了したファイルは最後に回して取り込む" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        # 「始」だけで「終」の無いジャーナル（取り込みの途中で強制終了した跡）
        writeTestJournal $root @("始`t営業\議事録.docx`t1") | Out-Null

        runIndexer $root | Should -Be 0
        (readTestStatus $root).Rows["営業\議事録.docx"].状態 | Should -Be ${stateDone}
        @([IngestJournal]::new("$root\work\store\journal\ingest.tsv").ReadPrevious())[-1] | Should -Be "終"
        assertGramMatchesFullScan $root "営業"
    }

    It "前の版の ingesting.txt（強制終了したファイルの記録）も読んで、最後に回して取り込み、読んだあとは消す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        writeListFile "$root\work\ingesting.txt" @("1`t営業\議事録.docx")

        runIndexer $root | Should -Be 0
        (readTestStatus $root).Rows["営業\議事録.docx"].状態 | Should -Be ${stateDone}
        Test-Path -LiteralPath "$root\work\ingesting.txt" | Should -Be $false
    }

    It "続けて強制終了したファイルは取り込まずに失敗とする" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        writeTestJournal $root @("始`t営業\議事録.docx`t2") | Out-Null

        runIndexer $root | Should -Be 0
        $row = (readTestStatus $root).Rows["営業\議事録.docx"]
        $row.状態 | Should -Be ${stateFailed}
        $row.エラー | Should -Match "強制終了"
    }

    It "Office に加えてテキスト（Shift_JIS の .txt）も取り込み、その語で検索すると正しい行番号でヒットする" {
        $root = newRoot
        $dir = Join-Path $TestDrive "テキスト混在"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        Copy-Item -LiteralPath $docxSource -Destination "$dir\議事録.docx"
        $sjisBytes = [System.Text.Encoding]::GetEncoding(932).GetBytes("1行目`r`n検索語がある2行目`r`n3行目`r`n")
        [System.IO.File]::WriteAllBytes("$dir\メモ.txt", $sjisBytes)
        writeTestSettings $root @(@{ name = ""; path = $dir; enabled = $true })

        runIndexer $root | Should -Be 0

        $status = readTestStatus $root
        $status.Rows["テキスト混在\メモ.txt"].状態 | Should -Be ${stateDone}
        $status.Rows["テキスト混在\議事録.docx"].状態 | Should -Be ${stateDone}

        $contentIndexFiles = (getContentIndexFiles @("$root\work\store\content")).ContentIndexFiles
        $hits = @((searchContentIndex "検索語" $contentIndexFiles $true).Hits)
        $hits.Count | Should -Be 1
        $hits[0].Book | Should -Be "メモ.txt"
        $hits[0].LineNumber | Should -Be 2
        $hits[0].Line | Should -Be "検索語がある2行目"
    }
}

Describe "indexer.ps1（利用者のPowerPointが起動している場合）" -Tag Office {
    # 利用者のPowerPointが起動していることにするため、自分のセッションで実際にPowerPointを起動する
    # （getApp は自分のセッションの POWERPNT を数えるため、モックではなく実際に起動して確かめる。
    #   実際の PowerPoint が要るため Office タグにする（既定・CI では実行しない。tests/run.ps1 の $defaultExclude）
    It "PowerPointが要るファイルは取り込まずに後回しにする。閉じれば次回のインデックス作成で取り込む" {
        $dir = Join-Path $TestDrive "後回し"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        Copy-Item -LiteralPath $docxSource -Destination "$dir\議事録.docx"
        Copy-Item -LiteralPath "${testDataDir}\office\PowerPoint\形式\旧形式.ppt" -Destination "$dir\旧形式.ppt"
        $root = newRoot
        writeTestSettings $root @(@{ name = "後回し"; path = $dir; enabled = $true })

        # 自分で起動したことが分かるよう、起動前後のプロセスIDの差から新しいプロセスを控える（後で終わるのを待つため）
        $before = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $userPpt = New-Object -ComObject PowerPoint.Application
        $after = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $userPptId = @($after | Where-Object { $before -notcontains $_ }) | Select-Object -First 1
        try {
            runIndexer $root | Should -Be 0

            $status = readTestStatus $root
            # PowerPointが要るファイルは「未取り込み」のまま残し（「失敗」にすると次の更新で飛ばされる）、理由を「エラー」列に書く
            $status.Rows["後回し\旧形式.ppt"].状態 | Should -Be ${stateNew}
            $status.Rows["後回し\旧形式.ppt"].エラー | Should -Be ${postponedReasonRunning}
            # Office を使わないファイルはそのまま取り込む
            $status.Rows["後回し\議事録.docx"].状態 | Should -Be ${stateDone}

            $progress = readTestProgress
            $progress.Processed | Should -Be 2
            $progress.Failed | Should -Be 1
            $progress.Remaining | Should -Be 0
            readTestError | Should -BeNullOrEmpty
            # 利用者のPowerPointは、インデックス作成の間も強制終了されていない（プロセスが残っていることで確かめる）
            $userPptId | Should -Not -BeNullOrEmpty
            (Get-Process -Id $userPptId -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
        } finally {
            # 完全に終わるまで待ってから次のインデックス作成に進む（次回は自分のセッションに残っていないようにする）
            $userPpt.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($userPpt)
            if ($userPptId) {
                $deadline = [datetime]::Now.AddSeconds(15)
                while ((Get-Process -Id $userPptId -ErrorAction SilentlyContinue) -and [datetime]::Now -lt $deadline) {
                    Start-Sleep -Milliseconds 200
                }
            }
        }

        # 利用者がPowerPointを閉じたので、次のインデックス作成で取り込む
        runIndexer $root | Should -Be 0
        $second = (readTestStatus $root).Rows["後回し\旧形式.ppt"]
        $second.状態 | Should -Be ${stateDone}
        $second.エラー | Should -BeNullOrEmpty
    }
}

Describe "indexer.ps1（後回しの司令の流れ。実際のPowerPointは使わない）" -Tag Io {
    BeforeAll {
        . "${scriptsDir}\shared\office\office_app.ps1"
        . "${scriptsDir}\tebunko\indexer\extract_office.ps1"
    }

    It "利用者のPowerPointが使用中の例外なら、後回しにして起動し直しを早めない。次回はそのファイルを取り込む" {
        $dir = Join-Path $TestDrive "後回し2"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        Copy-Item -LiteralPath $docxSource -Destination "$dir\議事録.docx"
        Copy-Item -LiteralPath "${testDataDir}\office\PowerPoint\形式\旧形式.ppt" -Destination "$dir\旧形式.ppt"
        $root = newRoot
        writeTestSettings $root @(@{ name = "後回し2"; path = $dir; enabled = $true })

        # ingestFile を .ppt だけ差し替える（.docx は本物のまま。getApp が投げるのと同じ型・文言の例外にする）。
        # 1 回目は例外（利用者の PowerPoint が使用中）、閉じたことにする 2 回目は成功にする
        $global:testOfficeAppInUseMessage = ${officeAppInUseMessage}
        $global:testUserPptOpen = $true
        Mock ingestFile {
            if ($global:testUserPptOpen) {
                throw (New-Object System.InvalidOperationException "PowerPoint$global:testOfficeAppInUseMessage")
            }
            return 1
        } -ParameterFilter { $sourcePath -like "*.ppt" }
        Mock stopAllApps {}

        try {
            runIndexer $root @{ Workers = 0 } | Should -Be 0

            $status = readTestStatus $root
            # 失敗として数え、取り込み一覧は「未取り込み」のまま理由を「エラー」列に書く
            $status.Rows["後回し2\旧形式.ppt"].状態 | Should -Be ${stateNew}
            $status.Rows["後回し2\旧形式.ppt"].エラー | Should -Be ${postponedReasonRunning}
            $status.Rows["後回し2\議事録.docx"].状態 | Should -Be ${stateDone}
            $progress = readTestProgress
            $progress.Failed | Should -Be 1
            $progress.Remaining | Should -Be 0
            # 失敗の件数（getIndexingState）に入る。待ちの件数には入らない
            $indexingState = getIndexingState ([datetime]::MaxValue) "$root\work\store\ingest_status.tsv"
            $indexingState.Failed | Should -Be 1
            $indexingState.Pending | Should -Be 0
            (readStatusFile "$root\work\store\ingest_status.tsv").Rows.Values | Where-Object { $_.エラー } | ForEach-Object { $_.エラー } | Should -Be ${postponedReasonRunning}
            (Get-Content -LiteralPath "$root\work\indexing_log.txt" -Raw) | Should -Match "PowerPoint を閉じてからもう一度更新すると、そのまま取り込みます"
            (Get-Content -LiteralPath "$root\work\indexing_log.txt" -Raw) | Should -Not -Match "失敗分も再取り込みする"
            # stopAllApps は後片付け（finally）で 1 回だけ呼ばれる。後回しを取り込んだ件数に数えると
            # 100 件ごとの起動し直しの判定が早まって途中でも呼ばれるが、ここでは増えない
            # （RestartInterval を差し替えられる runIngestWorker 単体の It で、数えないことを詳しく確かめている）
            Should -Invoke stopAllApps -Times 1 -Exactly -Scope It

            # 利用者が PowerPoint を閉じたことにすると、後回しにしたファイルを取り込む
            $global:testUserPptOpen = $false
            runIndexer $root | Should -Be 0
            # 再取り込みを選ばなくても取り込み、エラー列は空になる
            $second = (readTestStatus $root).Rows["後回し2\旧形式.ppt"]
            $second.状態 | Should -Be ${stateDone}
            $second.エラー | Should -BeNullOrEmpty
        } finally {
            Remove-Variable -Name testOfficeAppInUseMessage, testUserPptOpen -Scope Global -ErrorAction SilentlyContinue
        }
    }

    It "取り込みのワーカーを諦めたら、そのワーカーのファイルを未取り込みのまま後回しにし、0 で終わる。取り込み一覧・インデックスは変えない" {
        $dir = Join-Path $TestDrive "諦め"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        foreach ($name in "a", "b", "c") {
            Copy-Item -LiteralPath $docxSource -Destination "$dir\$name.docx"
        }
        $root = newRoot
        writeTestSettings $root @(@{ name = "諦め"; path = $dir; enabled = $true })
        # 最初のファイルを渡す直前に、そのワーカーを諦めたことにする（本物では、制限時間を過ぎたときの監視が立てる）
        $global:testAbandonDone = $false
        $hook = @{ Script = $runPath; Pattern = '^\s+\$dispatcher\.Submit\(\$step\)'; Action = {
                if (-not $global:testAbandonDone) {
                    $global:testAbandonDone = $true
                    foreach ($slot in $dispatcher.Pool.Slots) { $slot.State.Abandoned = $true }
                }
            }
        }
        try {
            runIndexer $root @{ Workers = 2 } @($hook) | Should -Be 0
        } finally {
            Remove-Variable -Name testAbandonDone -Scope Global -ErrorAction SilentlyContinue
        }

        $status = readTestStatus $root
        foreach ($name in "a", "b", "c") {
            $status.Rows["諦め\$name.docx"].状態 | Should -Be ${stateNew}
            $status.Rows["諦め\$name.docx"].エラー | Should -Be ${postponedReasonInUse}
        }
        (readTestProgress).Failed | Should -Be 3
        # 書き出していない（取り込んだ TSV もインデックスも無い）。取り込み中の記録・書き出しの置き場にも残さない
        (findIndexFoldersWithBooks "$root\work\store\content").Count | Should -Be 0
        @([IngestJournal]::new("$root\work\store\journal\ingest.tsv").ReadPrevious())[-1] | Should -Be "終"
        @(Get-ChildItem -LiteralPath "$root\work\publish" -Recurse -File -Force -ErrorAction SilentlyContinue).Count | Should -Be 0
        (Get-Content -LiteralPath "$root\work\indexing_log.txt" -Raw) | Should -Match "PowerPoint を使っている方がいたため取り込めませんでした"
    }

    It "司令のスレッドで取り込む形では、終わりの待ちに、中止の合図（Channel.Stop）を見る判定を渡す（立てる前は偽・立てたあとは真）" {
        $dir = Join-Path $TestDrive "待ち配線"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        Copy-Item -LiteralPath $docxSource -Destination "$dir\議事録.docx"
        $root = newRoot
        writeTestSettings $root @(@{ name = "待ち配線"; path = $dir; enabled = $true })
        $global:waitAnswers = New-Object System.Collections.ArrayList
        Mock stopAllApps {}
        Mock waitKeptApps {
            [void]$global:waitAnswers.Add((& $shouldStop))
            # 渡された判定が閉じ込めた $channel に、中止の合図を立てる
            & $shouldStop.Module { $channel.Stop = $true }
            [void]$global:waitAnswers.Add((& $shouldStop))
        }
        try {
            # 待ちの中で中止の合図を立てるため、終了コードは見ない
            [void](runIndexer $root @{ Workers = 0 })
            @($global:waitAnswers) | Should -Be @($false, $true)
        } finally {
            Remove-Variable -Name waitAnswers -Scope Global -ErrorAction SilentlyContinue
        }
    }
}

Describe "indexer.ps1（クラウドにだけあるファイル）" -Tag Io {
    It "既定ではダウンロードせずに残し（一覧に行を作らない・件数とパスをログに残す）、ダウンロードして取り込む指定なら取り込んで検索できる" {
        # クラウドにだけあるファイルは、OFFLINE 属性で代わりに作る（本物の OneDrive には触れず、ダウンロードも起きない）
        $dir = Join-Path $TestDrive "雲の資料"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $cloudFile = Join-Path $dir "議事録.docx"
        Copy-Item -LiteralPath $docxSource -Destination $cloudFile
        [System.IO.File]::SetAttributes($cloudFile, [System.IO.FileAttributes]::Offline)
        try {
            $root = newRoot
            writeTestSettings $root @(@{ name = "雲"; path = $dir; enabled = $true })

            runIndexer $root | Should -Be 0
            $channel = $script:lastChannel
            $channel.CloudSkipped | Should -Be 1
            (readTestStatus $root).Rows.Count | Should -Be 0
            [System.IO.File]::Exists("$root\work\store\content\雲\content_index.docx.001.tsv") | Should -Be $false
            $log = [System.IO.File]::ReadAllText("$root\work\indexing_log.txt")
            $log | Should -Match "クラウドにだけあるファイル 1 件は、ダウンロードせずに残します"
            $log | Should -Match ([regex]::Escape("雲\議事録.docx"))
            (readTestProgress).Detail | Should -Be "クラウドにだけあるファイルは、ダウンロードせずに残しました"

            # 次のインデックス作成でも確かめる。ダウンロードして取り込む指定なら取り込む
            runIndexer $root @{ IncludeCloud = $true } | Should -Be 0
            $script:lastChannel.CloudSkipped | Should -Be 0
            (readTestStatus $root).Rows["雲\議事録.docx"].状態 | Should -Be ${stateDone}
            [System.IO.File]::Exists("$root\work\store\content\雲\content_index.docx.001.tsv") | Should -Be $true
        } finally {
            [System.IO.File]::SetAttributes($cloudFile, [System.IO.FileAttributes]::Normal)
        }
    }

    # 前回未完了・前回失敗の行を持つ、クラウドにだけあるファイル（OFFLINE 属性）を作る。戻りは @{ Dir; File; Row; BookDir }
    # （本物の OneDrive には触れず、ダウンロードも起きない）
    BeforeAll {
    function newCloudCase {
        param ([string]$name, [string]$root, [string]$state)
        $dir = Join-Path $TestDrive $name
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $file = Join-Path $dir "議事録.docx"
        Copy-Item -LiteralPath $docxSource -Destination $file
        [System.IO.File]::SetAttributes($file, [System.IO.FileAttributes]::Offline)
        writeTestSettings $root @(@{ name = "雲"; path = $dir; enabled = $true })
        $info = Get-Item -LiteralPath $file
        $row = newStatusRow "雲\議事録.docx" (formatFileTime $info.LastWriteTime) ([string]$info.Length) $state "" "" "" ""
        writeStatusFile @(@{ Path = $dir; Name = "雲" }) @($row) "$root\work\store\ingest_status.tsv"
        $bookDir = "$root\work\store\content\雲\議事録.docx"
        [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
        [System.IO.File]::WriteAllText("$bookDir\途中.tsv", "x")
        # 前回、このファイルを取り込み始めて、書き出す前に止まった跡（取り込み一覧に前回の行が 1 行ある）
        writeTestJournal $root @("始`t雲\議事録.docx`t1") 1 | Out-Null
        return @{ File = $file; BookDir = $bookDir }
    }
    }

    It "前回未完了でクラウドにだけあるファイルは、<name>" -TestCases @(
        @{ name = "確認で取りやめたら、前回の行と取り込み済みの中身を残す（フォルダは実行の始めに本文インデックスのファイルへまとめる）"; mode = "cancel"; exitCode = 2; rowKept = $true; dirKept = $true }
        @{ name = "確認でダウンロードせずに進めたら、行を作らず、取り込み途中のインデックスのフォルダを消す"; mode = "approve"; exitCode = 0; rowKept = $false; dirKept = $false }
        @{ name = "画面なしの実行でも、行を作らず、取り込み途中のインデックスのフォルダを消す"; mode = "none"; exitCode = 0; rowKept = $false; dirKept = $false }
    ) {
        param ($name, $mode, $exitCode, $rowKept, $dirKept)
        $root = newRoot
        $case = newCloudCase "雲の未完了_$mode" $root ${stateNew}
        try {
            $options = @{}
            $breaks = @()
            if ($mode -ne "none") {
                $options.ConfirmTargets = $true
                $answer = @{ Script = $reporterPath; Pattern = '^\s+if \(!\$ch\.Answered\.WaitOne' }
                $answer.Action = if ($mode -eq "cancel") { { answerIndexingPlan $channel $null } } else { { answerIndexingPlan $channel @{ RetryFailed = $false; IncludeCloud = $false } } }
                $breaks = @($answer)
            }
            runIndexer $root $options $breaks | Should -Be $exitCode
            $rows = (readTestStatus $root).Rows
            $rows.ContainsKey("雲\議事録.docx") | Should -Be $rowKept
            if ($rowKept) { $rows["雲\議事録.docx"].状態 | Should -Be ${stateNew} }
            # 取り込み途中のインデックス（TSV）は、実行の始めに本文インデックスのファイルへまとめるため、フォルダそのものはどの場合も無くなる。残すなら中身がそのファイルに入り、消すなら入らない
            [System.IO.Directory]::Exists($case.BookDir) | Should -Be $false
            [System.IO.File]::Exists("$root\work\store\content\雲\content_index.docx.001.tsv") | Should -Be $dirKept
        } finally {
            [System.IO.File]::SetAttributes($case.File, [System.IO.FileAttributes]::Normal)
        }
    }

    It "前回失敗でクラウドにだけあるファイルは、失敗分を更新し直す=<retry>・ダウンロードする=<include> のとき、<expect>" -TestCases @(
        @{ retry = $false; include = $false; skipped = 0; done = $false; expect = "残す（クラウドのせいで残した件数には数えない）" }
        @{ retry = $false; include = $true; skipped = 0; done = $false; expect = "更新し直さないので残す" }
        @{ retry = $true; include = $false; skipped = 1; done = $false; expect = "ダウンロードせずに残す（件数に数える）" }
        @{ retry = $true; include = $true; skipped = 0; done = $true; expect = "ダウンロードして取り込む" }
    ) {
        param ($retry, $include, $skipped, $done, $expect)
        $root = newRoot
        $case = newCloudCase "雲の失敗_${retry}_${include}" $root ${stateFailed}
        try {
            runIndexer $root @{ RetryFailed = $retry; IncludeCloud = $include } | Should -Be 0
            $script:lastChannel.CloudSkipped | Should -Be $skipped
            (readTestStatus $root).Rows["雲\議事録.docx"].状態 | Should -Be $(if ($done) { ${stateDone} } else { ${stateFailed} })
        } finally {
            [System.IO.File]::SetAttributes($case.File, [System.IO.FileAttributes]::Normal)
        }
    }

    It "画面なしの実行は、設定の cloudFiles が <mode> のとき、クラウドにだけあるファイルを<expect>" -TestCases @(
        @{ mode = "ask"; imported = $false; expect = "ダウンロードせずに残す" }
        @{ mode = "download"; imported = $true; expect = "ダウンロードして取り込む" }
    ) {
        param ($mode, $imported, $expect)
        $root = newRoot
        $dir = Join-Path $TestDrive "雲の画面なし_$mode"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $file = Join-Path $dir "議事録.docx"
        Copy-Item -LiteralPath $docxSource -Destination $file
        [System.IO.File]::SetAttributes($file, [System.IO.FileAttributes]::Offline)
        writeTestSettings $root @(@{ name = "雲"; path = $dir; enabled = $true })
        writeCloudFiles $mode "$root\setting.config"
        $global:indexerTestRoot = $root
        $point = Set-PSBreakpoint -Script $settingsScriptPath -Line (findLine $settingsScriptPath '^\$\{settingsFile\}\s*=') -Action {
            Set-Variable -Name rootDir -Value $global:indexerTestRoot -Scope 1
        }
        try {
            $global:LASTEXITCODE = 0
            & $indexerPath *> $null
            $LASTEXITCODE | Should -Be 0
            (readTestStatus $root).Rows.ContainsKey("雲\議事録.docx") | Should -Be $imported
        } finally {
            Remove-PSBreakpoint -Breakpoint $point
            Remove-Variable -Name indexerTestRoot -Scope Global -ErrorAction SilentlyContinue
            [System.IO.File]::SetAttributes($file, [System.IO.FileAttributes]::Normal)
        }
    }
}

Describe "indexer.ps1（画面の確認・中止）" -Tag Io {
    BeforeAll {
        $source = newSourceFolder "総務"
        # 確認待ち（IndexingReporter.WaitForApproval）の中で、画面の返事を置く
        $approvalLine = @{ Script = $reporterPath; Pattern = '^\s+if \(!\$ch\.Answered\.WaitOne' }
    }

    It "確認を待つとき（ConfirmTargets）、取りやめたら、取り込まずに 2 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        $cancel = $approvalLine.Clone()
        $cancel.Action = { answerIndexingPlan $channel $null }

        runIndexer $root @{ ConfirmTargets = $true } @($cancel) | Should -Be 2
        # 1 件も取り込んでいないため、取り込み対象にした行は記録しない
        (readTestStatus $root).Rows.Count | Should -Be 0
        Test-Path -LiteralPath "$root\work\store\content\総務\議事録.docx" | Should -Be $false
        (readTestProgress).Detail | Should -Be "インデックス作成を取りやめました"
    }

    It "取りやめても、段階ごとの時間の行がログに出る" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        $cancel = $approvalLine.Clone()
        # 返事を待った時間（画面の確認待ち）が、全体に含めずに別に出ることを確かめるため、返事を少し遅らせる
        $cancel.Action = { Start-Sleep -Milliseconds 300; answerIndexingPlan $channel $null }

        runIndexer $root @{ ConfirmTargets = $true } @($cancel) | Should -Be 2
        $log = [System.IO.File]::ReadAllText("$root\work\indexing_log.txt")
        $log | Should -Match "段階ごとの時間（秒）: "
        $log | Should -Match "画面の確認待ち [0-9.]+（全体に含めない）"
        $stages = $script:lastChannel.StageTimings.Stages
        $stages.Wait | Should -BeGreaterThan 0.2
    }

    It "確認を待つとき（ConfirmTargets）、取りやめたら、更新されたファイルの行も前回の記録のまま残す（次回も更新ありになる）" {
        $edited = newSourceFolder "企画"
        $root = newRoot
        writeTestSettings $root @(@{ name = "企画"; path = $edited; enabled = $true })
        runIndexer $root | Should -Be 0
        $before = (readTestStatus $root).Rows["企画\議事録.docx"]
        (Get-Item -LiteralPath "$edited\議事録.docx").LastWriteTime = (Get-Date).AddDays(1)
        $cancel = $approvalLine.Clone()
        $cancel.Action = { answerIndexingPlan $channel $null }

        runIndexer $root @{ ConfirmTargets = $true } @($cancel) | Should -Be 2
        $row = (readTestStatus $root).Rows["企画\議事録.docx"]
        $row.状態 | Should -Be ${stateDone}
        $row.更新日時 | Should -Be $before.更新日時
    }

    It "確認を待つとき（ConfirmTargets）、画面が開始を選んだら取り込む（失敗分の再取り込みも画面の返事に従う）" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        runIndexer $root | Should -Be 0
        $approve = $approvalLine.Clone()
        $approve.Action = { answerIndexingPlan $channel @{ RetryFailed = $true } }

        runIndexer $root @{ ConfirmTargets = $true } @($approve) | Should -Be 0
        $progress = readTestProgress
        $progress.Processed | Should -Be 1
        $progress.Failed | Should -Be 1
    }

    Context "無くなったファイルのインデックスの削除（［更新を開始］のあと）" {
        BeforeEach {
            $script:vanishName = "無くなる$([guid]::NewGuid().ToString('N').Substring(0, 6))"
            $script:vanishSource = newSourceFolder $script:vanishName
            $script:vanishRoot = newRoot
            writeTestSettings $script:vanishRoot @(@{ name = $script:vanishName; path = $script:vanishSource; enabled = $true })
            runIndexer $script:vanishRoot | Should -Be 0
            $script:vanishText = { (@(Get-ChildItem -LiteralPath "$($script:vanishRoot)\work\store\content\$($script:vanishName)" -Recurse -Filter "content_index.*.tsv" | ForEach-Object { Get-Content -Raw -LiteralPath $_.FullName }) -join "") }
            (& $script:vanishText) | Should -Match "TC21-01"
        }

        It "無くなったファイルがあるとき、確認で取りやめると、取り込み一覧の行も本文インデックスも残る" {
            Remove-Item -LiteralPath "$($script:vanishSource)\資料\提案.pptx"
            $cancel = $approvalLine.Clone()
            $cancel.Action = { answerIndexingPlan $channel $null }

            runIndexer $script:vanishRoot @{ ConfirmTargets = $true } @($cancel) | Should -Be 2

            (readTestStatus $script:vanishRoot).Rows.ContainsKey("$($script:vanishName)\資料\提案.pptx") | Should -BeTrue
            (& $script:vanishText) | Should -Match "TC21-01"
        }

        It "直下だけに設定を変えて取りやめると、サブフォルダ側のインデックスは残り、前回の設定（ingestedSubfolders）は変わらない" {
            writeTestSettings $script:vanishRoot @(@{ name = $script:vanishName; path = $script:vanishSource; enabled = $true; subfolders = $false })
            $cancel = $approvalLine.Clone()
            $cancel.Action = { answerIndexingPlan $channel $null }

            runIndexer $script:vanishRoot @{ ConfirmTargets = $true } @($cancel) | Should -Be 2

            (readTestStatus $script:vanishRoot).Rows.ContainsKey("$($script:vanishName)\資料\提案.pptx") | Should -BeTrue
            (& $script:vanishText) | Should -Match "TC21-01"
            @(getTargetFolders "$($script:vanishRoot)\setting.config")[0].IngestedSubfolders | Should -BeTrue
        }

        It "直下だけに設定を変えて［更新を開始］すると、サブフォルダ側が消え、前回の設定が false になる（確認には『サブフォルダ側の N 件』を出す）" {
            writeTestSettings $script:vanishRoot @(@{ name = $script:vanishName; path = $script:vanishSource; enabled = $true; subfolders = $false })
            $global:planRemovedCount = $null
            $approve = $approvalLine.Clone()
            $approve.Action = { $global:planRemovedCount = @($channel.Plan)[0].SubfolderRemovedCount; answerIndexingPlan $channel @{ RetryFailed = $false } }
            try {
                runIndexer $script:vanishRoot @{ ConfirmTargets = $true } @($approve) | Should -Be 0
                $global:planRemovedCount | Should -Be 1
            } finally {
                Remove-Variable -Name planRemovedCount -Scope Global -ErrorAction SilentlyContinue
            }

            (readTestStatus $script:vanishRoot).Rows.ContainsKey("$($script:vanishName)\資料\提案.pptx") | Should -BeFalse
            (& $script:vanishText) | Should -Not -Match "TC21-01"
            @(getTargetFolders "$($script:vanishRoot)\setting.config")[0].IngestedSubfolders | Should -BeFalse
        }

        It "含める設定のまま更新して終わると、前回の設定は true で書かれる。確認を出さない実行でも書く" {
            runIndexer $script:vanishRoot | Should -Be 0
            $folder = @(getTargetFolders "$($script:vanishRoot)\setting.config")[0]
            $folder.IngestedSubfolders | Should -BeTrue
            $folder.Subfolders | Should -BeTrue
        }

        It "［更新を開始］すると、無くなったファイルのインデックスが消えて検索に出ない" {
            Remove-Item -LiteralPath "$($script:vanishSource)\資料\提案.pptx"
            $approve = $approvalLine.Clone()
            $approve.Action = { answerIndexingPlan $channel @{ RetryFailed = $false } }

            runIndexer $script:vanishRoot @{ ConfirmTargets = $true } @($approve) | Should -Be 0

            (readTestStatus $script:vanishRoot).Rows.ContainsKey("$($script:vanishName)\資料\提案.pptx") | Should -BeFalse
            (& $script:vanishText) | Should -Not -Match "TC21-01"
        }

        It "確認を出さない実行でも、無くなったファイルのインデックスは消える" {
            Remove-Item -LiteralPath "$($script:vanishSource)\資料\提案.pptx"

            runIndexer $script:vanishRoot | Should -Be 0

            (readTestStatus $script:vanishRoot).Rows.ContainsKey("$($script:vanishName)\資料\提案.pptx") | Should -BeFalse
            (& $script:vanishText) | Should -Not -Match "TC21-01"
        }

        It "［更新を開始］のあと、取り込み一覧を書き直す前に、外す予定の記録が書かれている（強制終了しても次の回に外せる）" {
            Remove-Item -LiteralPath "$($script:vanishSource)\資料\提案.pptx"
            $global:pendingRecorded = -1
            $probe = @{ Script = $runPath; Pattern = '^\s+writeStatusKeepingDropped \$ledger \$folders \$rows\.ToArray\(\)'; Action = { $global:pendingRecorded = @(readPendingRemovals).Count } }
            $approve = $approvalLine.Clone()
            $approve.Action = { answerIndexingPlan $channel @{ RetryFailed = $false } }
            try {
                runIndexer $script:vanishRoot @{ ConfirmTargets = $true } @($approve, $probe) | Should -Be 0
                $global:pendingRecorded | Should -BeGreaterThan 0
            } finally {
                Remove-Variable -Name pendingRecorded -Scope Global -ErrorAction SilentlyContinue
            }
        }
    }

    It "行の［更新］（選んだ名前だけ）のとき、ほかのフォルダの前回の設定は書かない。チェックなしのフォルダも書かない" {
        $picked = newSourceFolder "選ぶ"
        $other = newSourceFolder "選ばない"
        $off = newSourceFolder "外す"
        $root = newRoot
        writeTestSettings $root @(
            @{ name = "選ぶ"; path = $picked; enabled = $true; subfolders = $false },
            @{ name = "選ばない"; path = $other; enabled = $true; subfolders = $false },
            @{ name = "外す"; path = $off; enabled = $false; subfolders = $false })
        $approve = $approvalLine.Clone()
        $approve.Action = { answerIndexingPlan $channel @{ RetryFailed = $false } }

        runIndexer $root @{ ConfirmTargets = $true; OnlyNames = @("選ぶ") } @($approve) | Should -Be 0

        $folders = @(getTargetFolders "$root\setting.config")
        ($folders | Where-Object { $_.Name -eq "選ぶ" }).IngestedSubfolders | Should -BeFalse
        ($folders | Where-Object { $_.Name -eq "選ばない" }).IngestedSubfolders | Should -BeTrue
        ($folders | Where-Object { $_.Name -eq "外す" }).IngestedSubfolders | Should -BeTrue
    }

    It "取り込みの途中で中止を求められたら、残りを未取り込みのまま 2 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        # 1 件目を記録した直後に、画面が中止要求を作る
        $stop = @{ Script = $runPath; Pattern = '^\s+\$ledger\.AddRow\(\$row\)'; Action = { $channel.Stop = $true } }

        runIndexer $root @{} @($stop) | Should -Be 2
        $rows = @((readTestStatus $root).Rows.Values)
        @($rows | Where-Object { $_.状態 -eq ${stateNew} }).Count | Should -Be 2
        $progress = readTestProgress
        $progress.Processed | Should -Be 1
        $progress.Remaining | Should -Be 2
        # 取り込んだ TSV は元のファイルごとのフォルダに残さず（本文インデックスのファイルに入れる）、途中なので対応済みにはしない
        (findIndexFoldersWithBooks "$root\work\store\content").Count | Should -Be 0
    }

    It "取り込みの途中で中止を求められても、取り込めた分は済のまま検索でき、止めた・始めていない分は未取り込みで、次の回に取り込む" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        # 済になった行を記録した直後に、画面が中止要求を作る（取り込めないファイルがあれば、その行は失敗として記録される）
        $global:testStopDoneState = ${stateDone}
        $stop = @{ Script = $runPath; Pattern = '^\s+\$ledger\.AddRow\(\$row\)'; Action = { if ($row.状態 -eq $global:testStopDoneState) { $channel.Stop = $true } } }

        try {
            runIndexer $root @{} @($stop) | Should -Be 2
        } finally {
            Remove-Variable -Name testStopDoneState -Scope Global -ErrorAction SilentlyContinue
        }
        $rows = @((readTestStatus $root).Rows.Values)
        @($rows | Where-Object { $_.状態 -eq ${stateDone} }).Count | Should -Be 1
        @($rows | Where-Object { $_.状態 -eq ${stateNew} }).Count | Should -BeGreaterThan 0
        # 止めた分の取り込み途中の TSV は、本文インデックスに入らない（元のファイルごとのフォルダも残らない）
        (findIndexFoldersWithBooks "$root\work\store\content").Count | Should -Be 0
        $found = getContentIndexFiles @("$root\work\store\content")
        @($found.ContentIndexFiles).Count | Should -BeGreaterThan 0

        # 次の回で、残りを取り込んで済にする
        runIndexer $root | Should -Be 0
        $after = @((readTestStatus $root).Rows.Values)
        @($after | Where-Object { $_.状態 -eq ${stateNew} }).Count | Should -Be 0
        @($after | Where-Object { $_.状態 -eq ${stateDone} }).Count | Should -BeGreaterThan 1
        # 止めた回は gram 索引を作らず（検索は全件を照合）、次の回で作る
        assertGramMatchesFullScan $root "総務"
    }

    It "取り込みのスレッドにファイルを渡した直後に中止を求められたら、そのファイルを取り込まず（失敗にせず）未取り込みのまま、すぐ 2 で終わる" {
        $dir = Join-Path $TestDrive "渡した直後に中止"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        foreach ($name in "a", "b", "c") {
            Copy-Item -LiteralPath $docxSource -Destination "$dir\$name.docx"
        }
        $root = newRoot
        writeTestSettings $root @(@{ name = "渡した直後に中止"; path = $dir; enabled = $true })
        # 最初のファイルを取り込みのスレッドに渡す直前に、画面が中止要求を作る（スレッドは取り出すとき、中止を見て取り込まずに返す）
        $hook = @{ Script = $runPath; Pattern = '^\s+\$dispatcher\.Submit\(\$step\)'; Action = { $channel.Stop = $true } }
        $watch = [System.Diagnostics.Stopwatch]::StartNew()

        runIndexer $root @{ Workers = 1 } @($hook) | Should -Be 2
        $watch.Stop()
        $watch.Elapsed.TotalMinutes | Should -BeLessThan 2

        $status = readTestStatus $root
        foreach ($name in "a", "b", "c") {
            $status.Rows["渡した直後に中止\$name.docx"].状態 | Should -Be ${stateNew}
        }
        # 止めたファイルは失敗として数えない
        (readTestProgress).Failed | Should -Be 0
        (findIndexFoldersWithBooks "$root\work\store\content").Count | Should -Be 0
        @([IngestJournal]::new("$root\work\store\journal\ingest.tsv").ReadPrevious())[-1] | Should -Be "終"
        $log = Get-Content -LiteralPath "$root\work\indexing_log.txt" -Raw
        $log | Should -Match "中止のため、取り込みませんでした（次回取り込みします）"
        $log | Should -Match "中止の要求を受けたため、インデックス作成を中止します"
    }

    It "一覧化（クロール）の前に中止を求められたら、確認を出さずに 2 で終わり、取り込み一覧は前回のままで、無くなったファイルのインデックスも消さない" {
        $a = newSourceFolder "総務"
        $b = newSourceFolder "企画"
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $a; enabled = $true }, @{ name = "企画"; path = $b; enabled = $true })
        runIndexer $root | Should -Be 0
        $before = readTestStatus $root
        $beforeFiles = @(Get-ChildItem -LiteralPath "$root\work\store\content" -Recurse -File | ForEach-Object { $_.FullName } | Sort-Object)
        # 元のファイルが無くなったことにする（止めなければ、この行とインデックスが消える）
        Remove-Item -LiteralPath "$a\議事録.docx"
        $stop = @{ Script = $runPath; Pattern = '^\s+\$planner\.Crawl\('; Action = { $channel.Stop = $true } }

        runIndexer $root @{ ConfirmTargets = $true } @($stop) | Should -Be 2
        $after = readTestStatus $root
        $after.Rows.Count | Should -Be $before.Rows.Count
        $after.Rows.ContainsKey("総務\議事録.docx") | Should -Be $true
        @(Get-ChildItem -LiteralPath "$root\work\store\content" -Recurse -File | ForEach-Object { $_.FullName } | Sort-Object) | Should -Be $beforeFiles
        # 確認の返事を待っていない
        $script:lastChannel.Answered.WaitOne(0) | Should -Be $false
    }

    It "古い形式のインデックスがあるとき、作り直しの途中で中止したら、古い形式の分を消さず、目録の行は「作成中」のまま。最後まで作ると消して「作り終えた」にする" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        $ws = [Workspace]::new("$root\work")
        newTsv "$($ws.OldIndexDir)\総務\content_index.docx.001.tsv" @("古い中身")
        newTsv $ws.OldStatusFile @("古い取り込み一覧")
        $stop = @{ Script = $runPath; Pattern = '^\s+\$ledger\.AddRow\(\$row\)'; Action = { $channel.Stop = $true } }

        runIndexer $root @{} @($stop) | Should -Be 2

        Test-Path -LiteralPath "$($ws.OldIndexDir)\総務\content_index.docx.001.tsv" | Should -Be $true
        Test-Path -LiteralPath $ws.OldStatusFile | Should -Be $true
        (readIndexFormat $ws.FormatFile).Indexes["総務"] | Should -Be ${catalogRowBuilding}

        runIndexer $root | Should -Be 0

        Test-Path -LiteralPath $ws.OldIndexDir | Should -Be $false
        Test-Path -LiteralPath $ws.OldStatusFile | Should -Be $false
        (readIndexFormat $ws.FormatFile).Indexes["総務"] | Should -Be ${catalogRowFinished}
    }

    It "古い形式のインデックスがあるとき、取り込めないファイルがあっても、最後まで作ると古い形式の分を消して「作り終えた」にする（失敗は一覧に残す）" {
        $root = newRoot
        # $source には、取り込めない壊れた PowerPoint のファイルが入っている。作り直しの 1 回目で失敗する
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        $ws = [Workspace]::new("$root\work")
        newTsv "$($ws.OldIndexDir)\総務\content_index.docx.001.tsv" @("古い中身")
        newTsv $ws.OldStatusFile @("古い取り込み一覧")

        runIndexer $root | Should -Be 0

        (readTestStatus $root).Rows["総務\壊れた.pptx"].状態 | Should -Be ${stateFailed}
        Test-Path -LiteralPath $ws.OldIndexDir | Should -Be $false
        Test-Path -LiteralPath $ws.OldStatusFile | Should -Be $false
        (readIndexFormat $ws.FormatFile).Indexes["総務"] | Should -Be ${catalogRowFinished}
    }

    It "古い形式のインデックスが 2 つあるとき、1 つ目を取り込み終えたあと 2 つ目の途中で<Kind>と、どちらも「作り終えた」にせず、古い形式の分を消さない（終了コード <Code>）" -TestCases @(
        @{ Kind = "中止する"; Code = 2; Stop = $true }
        @{ Kind = "フォルダが見えなくなる"; Code = 1; Stop = $false }
    ) {
        param ($Kind, $Code, $Stop)
        $a = newSourceFolder "総務"
        $b = newSourceFolder "企画"
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $a; enabled = $true }, @{ name = "企画"; path = $b; enabled = $true })
        $ws = [Workspace]::new("$root\work")
        newTsv "$($ws.OldIndexDir)\総務\content_index.docx.001.tsv" @("古い中身")
        newTsv "$($ws.OldIndexDir)\企画\content_index.docx.001.tsv" @("古い中身")
        newTsv $ws.OldStatusFile @("古い取り込み一覧")
        # 2 つ目（企画）のファイルのところで止める。並べた順に取り込むが、並列のときは 1 つ目を取り込み終えているとは限らない（どちらでも、作り終えた扱いにはならない）
        if ($Stop) {
            # 中止は、2 つ目の最初のファイルを一覧に足したあとに求める
            $hook = @{ Script = $runPath; Pattern = '^\s+\$ledger\.AddRow\(\$row\)'; Action = {
                if ($row.相対パス -like "企画\*") { $channel.Stop = $true }
            } }
        } else {
            # フォルダが見えなくなるのは、2 つ目の最初のファイルを取り込む直前
            $global:indexerTestVictim = $b
            $hook = @{ Script = $dispatcherPath; Pattern = '^\s+\$row = \$this\.Targets\[\$i\]'; Action = {
                if ($this.Targets[$i].相対パス -like "企画\*" -and (Test-Path -LiteralPath $global:indexerTestVictim)) {
                    Remove-Item -LiteralPath $global:indexerTestVictim -Recurse -Force
                }
            } }
        }
        try {
            runIndexer $root @{} @($hook) | Should -Be $Code
        } finally {
            Remove-Variable -Name indexerTestVictim -Scope Global -ErrorAction SilentlyContinue
        }

        foreach ($name in "総務", "企画") {
            (readIndexFormat $ws.FormatFile).Indexes[$name] | Should -Be ${catalogRowBuilding}
            Test-Path -LiteralPath "$($ws.OldIndexDir)\$name\content_index.docx.001.tsv" | Should -Be $true
        }
        Test-Path -LiteralPath $ws.OldStatusFile | Should -Be $true
    }
}

Describe "indexer.ps1（設定から外れたインデックス）" -Tag Io {
    BeforeAll {
        $approvalLine = @{ Script = $reporterPath; Pattern = '^\s+if \(!\$ch\.Answered\.WaitOne' }
        function newDroppedRoot {
            # 「総務」「企画」を取り込んだあと、設定から「企画」を外した状態にする。Before には「企画」の中身（更新日時・TSV）を控える
            $a = newSourceFolder "総務"
            $b = newSourceFolder "企画"
            $root = newRoot
            writeTestSettings $root @(@{ name = "総務"; path = $a; enabled = $true }, @{ name = "企画"; path = $b; enabled = $true })
            runIndexer $root | Should -Be 0
            $contentIndexFile = @(Get-ChildItem -LiteralPath "$root\work\store\content\企画" -Filter "*.tsv" -File)[0]
            writeTestSettings $root @(@{ name = "総務"; path = $a; enabled = $true })
            return @{ Root = $root; Source = $a; ContentIndexFile = $contentIndexFile.FullName; Time = $contentIndexFile.LastWriteTimeUtc; Hash = (Get-FileHash -LiteralPath $contentIndexFile.FullName).Hash }
        }
        function assertDroppedKept {
            param ($state)
            $dir = "$($state.Root)\work\store\content\企画"
            Test-Path -LiteralPath $state.ContentIndexFile | Should -Be $true
            (Get-Item -LiteralPath $state.ContentIndexFile).LastWriteTimeUtc | Should -Be $state.Time
            (Get-FileHash -LiteralPath $state.ContentIndexFile).Hash | Should -Be $state.Hash
            $status = readTestStatus $state.Root
            @($status.Folders | ForEach-Object { $_.Name }) | Should -Contain "企画"
            @($status.Rows.Keys | Where-Object { $_ -like "企画\*" }).Count | Should -BeGreaterThan 0
            # 検索の経路（content_index_search.ps1）で、残したインデックスの中の語が今までどおり当たる
            $found = getContentIndexFiles @("$($state.Root)\work\store\content")
            $regex = New-Object regex ([regex]::Escape("大文字拡張子"))
            @((searchContentIndexFiles $found.ContentIndexFiles 0 $found.ContentIndexFiles.Count $regex -1 $regex "lines").Hits).Count | Should -BeGreaterThan 0
        }
    }

    It "確認なし（画面なしの実行）では、設定から外れたインデックスを消さず、黄色のログで知らせる（取り込むファイルが <added> 件）" -TestCases @(
        @{ added = 0 }
        @{ added = 1 }
    ) {
        param ($added)
        $state = newDroppedRoot
        if ($added -gt 0) {
            # 取り込むファイルがある回は、終わりの書き出し（取り込み一覧を書き直す所）を通る
            Copy-Item -LiteralPath $docxSource -Destination "$($state.Source)\追加.docx"
        }
        runIndexer $state.Root | Should -Be 0
        assertDroppedKept $state
        [System.IO.File]::ReadAllText("$($state.Root)\work\indexing_log.txt") | Should -Match "設定に無いインデックス（企画: .*）は削除せずに残しました"
    }

    It "確認で取りやめたら、設定から外れたインデックスを消さず、次の回にまた見つける" {
        $state = newDroppedRoot
        $cancel = $approvalLine.Clone()
        $cancel.Action = { answerIndexingPlan $channel $null }
        runIndexer $state.Root @{ ConfirmTargets = $true } @($cancel) | Should -Be 2
        assertDroppedKept $state
        $seen = New-Object System.Collections.Generic.List[string]
        $global:droppedSeenKinds = $seen
        $approve = $approvalLine.Clone()
        $approve.Action = { $global:droppedSeenKinds.Add((@($channel.Plan | Where-Object { $_.区分 -eq ${planKindDropped} } | ForEach-Object { $_.インデックス名 }) -join ",")); answerIndexingPlan $channel $null }
        runIndexer $state.Root @{ ConfirmTargets = $true } @($approve) | Should -Be 2
        Remove-Variable -Name droppedSeenKinds -Scope Global -ErrorAction SilentlyContinue
        $seen[0] | Should -Be "企画"
    }

    It "確認で［更新を開始］を押したら、設定から外れたインデックスを消し、取り込み一覧からも外す。確認に出す予定には「削除予定」の行がある" {
        $state = newDroppedRoot
        $global:droppedPlanKinds = $null
        $approve = $approvalLine.Clone()
        $approve.Action = { $global:droppedPlanKinds = @($channel.Plan | ForEach-Object { "$($_.インデックス名):$($_.区分)" }); answerIndexingPlan $channel @{ RetryFailed = $false } }
        runIndexer $state.Root @{ ConfirmTargets = $true } @($approve) | Should -Be 0
        $kinds = $global:droppedPlanKinds
        Remove-Variable -Name droppedPlanKinds -Scope Global -ErrorAction SilentlyContinue
        $kinds | Should -Contain "企画:${planKindDropped}"
        Test-Path -LiteralPath "$($state.Root)\work\store\content\企画" | Should -Be $false
        $status = readTestStatus $state.Root
        @($status.Folders | ForEach-Object { $_.Name }) | Should -Not -Contain "企画"
        @($status.Rows.Keys | Where-Object { $_ -like "企画\*" }).Count | Should -Be 0
        Test-Path -LiteralPath "$($state.Root)\work\store\content\総務" | Should -Be $true
    }
}

Describe "indexer.ps1（書き換えられた取り込み一覧）" -Tag Io {
    It "置き場の外を指す行・名前だけの行・使えない名前のフォルダ行があっても、外を消さず最後まで動き、黄色のログで知らせる" {
        $a = newSourceFolder "総務"
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $a; enabled = $true })
        runIndexer $root | Should -Be 0
        [System.IO.Directory]::CreateDirectory("$root\work\outside") | Out-Null
        [System.IO.File]::WriteAllText("$root\work\outside\keep.txt", "残る")
        $statusFile = "$root\work\store\ingest_status.tsv"
        $key = ${statusFolderKey}
        $extra = @(
            "総務\..\..\outside\keep.txt`t2024-01-01 00:00:00`t1`tok`t1`t2024-01-01 00:00:00`t",
            "..\outside\keep.txt`t2024-01-01 00:00:00`t1`tok`t1`t2024-01-01 00:00:00`t",
            "総務`t2024-01-01 00:00:00`t1`tok`t1`t2024-01-01 00:00:00`t",
            "${key}`t${a}`t旧営業.",
            "${key}`t${a}`t..\outside"
        ) -join "`r`n"
        $existing = [System.IO.File]::ReadAllText($statusFile)
        if (!$existing.EndsWith("`n")) { $existing += "`r`n" }
        [System.IO.File]::WriteAllText($statusFile, $existing + $extra + "`r`n", (New-Object System.Text.UTF8Encoding($true)))

        { runIndexer $root | Should -Be 0 } | Should -Not -Throw

        Test-Path -LiteralPath "$root\work\outside\keep.txt" | Should -Be $true
        Test-Path -LiteralPath "$root\work\store\content\総務" | Should -Be $true
        $log = [System.IO.File]::ReadAllText("$root\work\indexing_log.txt")
        $log | Should -Match "取り込み一覧に読めない行が \d+ 行ありました"
        $status = readTestStatus $root
        @($status.Folders | ForEach-Object { $_.Name }) | Should -Contain "総務"
        @($status.Rows.Keys | Where-Object { $_ -like "*outside*" }).Count | Should -Be 0
    }
}

Describe "indexer.ps1（選んだインデックスだけを更新する。OnlyNames）" -Tag Io {
    BeforeAll {
        $approvalLine = @{ Script = $reporterPath; Pattern = '^\s+if \(!\$ch\.Answered\.WaitOne' }
        $global:seenPlan = $null
    }

    BeforeEach {
        # 総務・企画の 2 つをチェック付きで取り込んだあと、両方のファイルを更新した状態にする
        $script:a = newSourceFolder "総務"
        $script:b = newSourceFolder "企画"
        $script:c = newSourceFolder "休み"
        $script:root = newRoot
        writeTestSettings $script:root @(
            @{ name = "総務"; path = $script:a; enabled = $true }
            @{ name = "企画"; path = $script:b; enabled = $true }
            @{ name = "休み"; path = $script:c; enabled = $false }
        )
        runIndexer $script:root | Should -Be 0
        foreach ($dir in @($script:a, $script:b)) {
            (Get-Item -LiteralPath "$dir\議事録.docx").LastWriteTime = (Get-Date).AddDays(1)
        }
        $script:before = (readTestStatus $script:root).Rows
    }

    It "選んだインデックスだけを取り込み、選ばなかったインデックスの記録・確認の表には触れない" {
        $approve = $approvalLine.Clone()
        $approve.Action = { $global:seenPlan = @($channel.Plan); answerIndexingPlan $channel @{ RetryFailed = $false } }

        runIndexer $script:root @{ ConfirmTargets = $true; OnlyNames = @("企画") } @($approve) | Should -Be 0

        $rows = (readTestStatus $script:root).Rows
        $rows["企画\議事録.docx"].更新日時 | Should -Not -Be $script:before["企画\議事録.docx"].更新日時
        # 選ばなかった総務は、更新されたファイルがあっても前回のまま（状態・日時・失敗の記録も含め行ごと同じ）
        foreach ($key in @($script:before.Keys | Where-Object { $_.StartsWith("総務\") })) {
            $rows[$key].更新日時 | Should -Be $script:before[$key].更新日時
            $rows[$key].状態 | Should -Be $script:before[$key].状態
            $rows[$key].エラー | Should -Be $script:before[$key].エラー
        }
        # 確認の表は、選んだインデックスだけ（選ばなかった行を「チェックなし」などと出さない）
        @($global:seenPlan | ForEach-Object { $_.インデックス名 }) | Should -Be @("企画")
        $script:lastChannel.OnlySkipped.Count | Should -Be 0
    }

    It "選んだ名前が設定に無い・チェックが付いていないときは、結果に残し、ほかの選んだものは更新する" {
        runIndexer $script:root @{ OnlyNames = @("企画", "休み", "無い名前") } | Should -Be 0

        $skipped = @($script:lastChannel.OnlySkipped)
        @($skipped | ForEach-Object { "$($_.Name):$($_.Reason)" }) | Should -Be @("休み:チェックが付いていません", "無い名前:設定にありません")
        $rows = (readTestStatus $script:root).Rows
        $rows["企画\議事録.docx"].更新日時 | Should -Not -Be $script:before["企画\議事録.docx"].更新日時
        $rows["総務\議事録.docx"].更新日時 | Should -Be $script:before["総務\議事録.docx"].更新日時
    }

    It "選んだ名前がどれも更新できなければ、エラーにして何も更新しない" {
        runIndexer $script:root @{ OnlyNames = @("休み") } | Should -Be 1

        readTestError | Should -Match "更新できるものがありません"
        (readTestStatus $script:root).Rows["総務\議事録.docx"].更新日時 | Should -Be $script:before["総務\議事録.docx"].更新日時
    }

    It "OnlyNames が空なら、チェックの付いたものすべてを更新する（これまでどおり）" {
        runIndexer $script:root @{ OnlyNames = @() } | Should -Be 0

        $rows = (readTestStatus $script:root).Rows
        $rows["企画\議事録.docx"].更新日時 | Should -Not -Be $script:before["企画\議事録.docx"].更新日時
        $rows["総務\議事録.docx"].更新日時 | Should -Not -Be $script:before["総務\議事録.docx"].更新日時
    }
}

Describe "indexer.ps1（制限時間）" -Tag Io {
    It "制限時間を過ぎて失敗したファイルは、制限時間で中止したことをエラーに書く" {
        $source = newSourceFolder "監査"
        $root = newRoot
        writeTestSettings $root @(@{ name = "監査"; path = $source; enabled = $true })
        # 取り込みの直前に、見張り（OfficeWatchdog）が制限時間を過ぎたと判断した状態にする
        $timeout = @{ Script = $runPath; Pattern = '^\s+\$result\.TsvCount = invokeTimedIngest'; Action = { $officeWatchdog.State.TimedOut = $true } }

        runIndexer $root @{} @($timeout) | Should -Be 0
        $status = readTestStatus $root
        $status.Rows["監査\壊れた.pptx"].状態 | Should -Be ${stateFailed}
        $status.Rows["監査\壊れた.pptx"].エラー | Should -Match "分以内に更新が終わらなかった"
        # 取り込めたファイルは、制限時間の印が立っていても成功のまま
        $status.Rows["監査\議事録.docx"].状態 | Should -Be ${stateDone}
    }
}

Describe "indexer.ps1（取り込み中に元のファイルが無くなる）" -Tag Io {
    # 取り込みを始める直前（1 件目）に、元のファイル・フォルダを消す
    BeforeAll {
        $beforeIngest = '^\s+\$row = \$this\.Targets\[\$i\]'
    }

    It "元のファイルが無くなっていたら、取り込まずに一覧から除く" {
        $source = newSourceFolder "人事"
        $root = newRoot
        writeTestSettings $root @(@{ name = "人事"; path = $source; enabled = $true })
        $global:indexerTestVictim = "$source\壊れた.pptx"
        $remove = @{ Script = $dispatcherPath; Pattern = $beforeIngest; Action = {
            if (Test-Path -LiteralPath $global:indexerTestVictim) { Remove-Item -LiteralPath $global:indexerTestVictim -Force }
        } }
        try {
            runIndexer $root @{} @($remove) | Should -Be 0
        } finally {
            Remove-Variable -Name indexerTestVictim -Scope Global
        }
        $status = readTestStatus $root
        $status.Rows.Count | Should -Be 2
        $status.Rows.ContainsKey("人事\壊れた.pptx") | Should -Be $false
    }

    It "前回の作成で残った TSV（元のファイルごとのフォルダ）は、次の作成の始めに本文インデックスのファイルへ入れる" {
        $source = newSourceFolder "総務3"
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務3"; path = $source; enabled = $true })
        runIndexer $root | Should -Be 0
        writeListFile "$root\work\store\content\総務3\残った.xlsx\S.tsv" @("残っていた中身")
        # 取り込みは終わったが、書き出す前に止まった跡（ジャーナルに「始」があり「出」が無い）
        $rowCount = @((readTestStatus $root).Rows.Keys).Count
        writeTestJournal $root @("始`t総務3\残った.xlsx`t1", "終") $rowCount | Out-Null

        runIndexer $root | Should -Be 0
        [System.IO.Directory]::Exists("$root\work\store\content\総務3\残った.xlsx") | Should -Be $false
        $contentIndexFiles = findContentIndexFiles "$root\work\store\content" "総務3" $false
        (searchContentIndex "残っていた中身" $contentIndexFiles $true).Hits.Count | Should -Be 1
    }

    It "クロール対象フォルダごと見えなくなったら、残りを未取り込みのまま 1 で終わる" {
        $source = newSourceFolder "法務"
        $root = newRoot
        writeTestSettings $root @(@{ name = "法務"; path = $source; enabled = $true })
        $global:indexerTestVictim = $source
        $remove = @{ Script = $dispatcherPath; Pattern = $beforeIngest; Action = {
            if (Test-Path -LiteralPath $global:indexerTestVictim) { Remove-Item -LiteralPath $global:indexerTestVictim -Recurse -Force }
        } }
        try {
            runIndexer $root @{} @($remove) | Should -Be 1
        } finally {
            Remove-Variable -Name indexerTestVictim -Scope Global
        }
        readTestError | Should -Match "クロール対象フォルダが見つからなくなった"
        $rows = @((readTestStatus $root).Rows.Values)
        $rows.Count | Should -Be 3
        @($rows | Where-Object { $_.状態 -eq ${stateNew} }).Count | Should -Be 3
    }
}

Describe "indexer.ps1（まれな状況）" -Tag Io {
    BeforeAll {
        $source = newSourceFolder "法務"
    }

    It "既定のワークスペースにほかのファイルがあれば、ワークスペースに何も書かずに 1 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "法務"; path = $source; enabled = $true })
        # 既定のワークスペースが使えないと判定された状態にする（空でないフォルダにエラーのファイルやログを書かない）
        $block = @{ Script = $runPath; Pattern = '^\s+if \(\$workspaceBlock\) \{'; Action = {
                Set-Variable -Name workspaceBlock -Value "「C:\Users\test\Documents\tebunko_ws」は空のフォルダではありません。" -Scope 1
            }
        }

        runIndexer $root @{} @($block) | Should -Be 1

        @(Get-ChildItem -LiteralPath "$root\work" -Force).Count | Should -Be 0
    }

}

Describe "indexer.ps1（前の版のインデックス）" -Tag Io {
    function script:newLegacyWorkspace {
        # 前の版のワークスペース（index\・system_index\・状態ファイル・取り込み一覧の「済」の行）を root\work に作る（しるしあり）。
        # 取り込み一覧の行は、元のファイルの実際の更新日時・サイズ・今の抽出版にする。
        # そうしないと getIngestDecision の sameFile が偽になり、本文インデックスが無いこと（lost）とは別の理由
        # （updated・outdated）で取り込み直す判定になってしまい、このテストが確かめたいことと違う経路を通ってしまう
        param ([string]$root, [string]$name, [string]$sourcePath)
        $work = "$root\work"
        newTsv "$work\index\$name\議事録.docx\S.tsv" @("旧いインデックス")
        newTsv "$work\index\$name\元のフォルダ.txt" @("# 前の版の記録", "$name`tC:\旧い場所")
        newTsv "$work\system_index\$name\システムインデックス.txt" @("x00000000")
        newTsv "$work\system_index_state.tsv" @("対応済み`t$name`t")
        $docxFile = Get-Item -LiteralPath "$sourcePath\議事録.docx"
        $updated = formatFileTime $docxFile.LastWriteTime
        $size = [string]$docxFile.Length
        $extractVersion = [string]$extractVersions[".docx"]
        writeStatusFile @([pscustomobject]@{ Path = $sourcePath; Name = $name }) @(
            (newStatusRow "$name\議事録.docx" $updated $size ${stateDone} "1" "2026/09/01 10:00:00" "" $extractVersion)
        ) "$work\store\ingest_status.tsv"
    }

    It "前の版のインデックスがあれば知らせ、前のシステムインデックスを消して、取り込み直す（前の版の index には触らない）" {
        $source = newSourceFolder "旧版"
        $root = newRoot
        writeTestSettings $root @(@{ name = "旧版"; path = $source; enabled = $true })
        newLegacyWorkspace $root "旧版" $source
        $before = (readTestStatus $root).Rows["旧版\議事録.docx"].取り込み日時

        runIndexer $root | Should -Be 0

        $log = [System.IO.File]::ReadAllText("$root\work\indexing_log.txt")
        $log | Should -Match "前の版のインデックス"
        # 「済」の行が取り込み直された理由が、本文インデックスが無いこと（lost）であることをログで確かめる
        $log | Should -Match "インデックスが無い・壊れている 1 件"
        [System.IO.File]::Exists("$root\work\store\content\旧版\content_index.docx.001.tsv") | Should -Be $true
        # 前の版のシステムインデックスと状態ファイルは消える
        Test-Path -LiteralPath "$root\work\system_index" | Should -Be $false
        Test-Path -LiteralPath "$root\work\system_index_state.tsv" | Should -Be $false
        # 取り込み一覧の「済」の行（前の版の記録）も、本文インデックスが無いため取り込み直す
        $row = (readTestStatus $root).Rows["旧版\議事録.docx"]
        $row.状態 | Should -Be ${stateDone}
        $row.取り込み日時 | Should -Not -Be $before
        # 前の版の index\ は読まず・消さない
        Test-Path -LiteralPath "$root\work\index\旧版\元のフォルダ.txt" | Should -Be $true
        Test-Path -LiteralPath "$root\work\index\旧版\議事録.docx\S.tsv" | Should -Be $true
    }

    It "前の版のシステムインデックスを消せなくても止まらず、ログに残して取り込む。閉じてもう一度で消える" {
        $source = newSourceFolder "旧版2"
        $root = newRoot
        writeTestSettings $root @(@{ name = "旧版2"; path = $source; enabled = $true })
        newLegacyWorkspace $root "旧版2" $source
        $legacyTxt = "$root\work\system_index\旧版2\システムインデックス.txt"
        $stream = [System.IO.File]::Open($legacyTxt, "Open", "Read", "None")
        try {
            runIndexer $root | Should -Be 0
        } finally {
            $stream.Dispose()
        }
        [System.IO.File]::ReadAllText("$root\work\indexing_log.txt") | Should -Match "消せませんでした"
        [System.IO.File]::Exists("$root\work\store\content\旧版2\content_index.docx.001.tsv") | Should -Be $true
        Test-Path -LiteralPath $legacyTxt | Should -Be $true

        runIndexer $root | Should -Be 0
        Test-Path -LiteralPath $legacyTxt | Should -Be $false
    }
}

Describe "indexer.ps1（取り込みのスレッド）" -Tag Io {
    BeforeAll {
        $source = newSourceFolder "並列"
        [System.IO.Directory]::CreateDirectory("$source\資料\下") | Out-Null
        Copy-Item -LiteralPath $docxSource -Destination "$source\資料\下\報告.docx"
        Copy-Item -LiteralPath $pptxSource -Destination "$source\資料\報告2.pptx"

        function script:readContentIndexLines([string]$root) {
            # 本文インデックスのファイルの中身（ファイル名・場所・行）を、パスの順に並べて返す
            $lines = New-Object System.Collections.Generic.List[string]
            foreach ($file in @([System.IO.Directory]::GetFiles("$root\work\store\content", ${contentIndexFilePattern}, "AllDirectories") | Sort-Object)) {
                $lines.Add($file.Substring("$root\work\store\content".Length))
                $lines.AddRange([string[]]@([System.IO.File]::ReadAllLines($file)))
            }
            return , $lines.ToArray()
        }
    }

    It "段階ごとの時間と種類ごとの時間が、スレッドを使っても使わなくても、ログと受け渡しの口（StageTimings）に出る（Workers = <workers>）" -TestCases @(
        @{ workers = 0 }
        @{ workers = 3 }
    ) {
        param ($workers)
        $root = newRoot
        writeTestSettings $root @(@{ name = "並列"; path = $source; enabled = $true })
        runIndexer $root @{ Workers = $workers } | Should -Be 0

        $timings = $script:lastChannel.StageTimings
        $timings.Stages.Total | Should -BeGreaterThan 0
        $timings.Stages.Ingest | Should -BeGreaterThan 0
        $timings.Stages.Gram | Should -BeGreaterThan 0
        ($timings.Stages.Startup + $timings.Stages.Crawl + $timings.Stages.Ingest + $timings.Stages.FinishFlush + $timings.Stages.Gram + $timings.Stages.Other) | Should -BeGreaterThan ($timings.Stages.Total - 0.01)
        ($timings.Stages.Startup + $timings.Stages.Crawl + $timings.Stages.Ingest + $timings.Stages.FinishFlush + $timings.Stages.Gram + $timings.Stages.Other) | Should -BeLessThan ($timings.Stages.Total + 0.01)
        $timings.Kinds[".docx"].Ok | Should -Be 2
        $timings.Kinds[".docx"].Count | Should -Be 2
        $timings.Kinds[".pptx"].Ok | Should -Be 2
        $timings.Kinds[".pptx"].Failed | Should -Be 1
        $timings.Kinds[".pptx"].Count | Should -Be 3
        $timings.Kinds[".docx"].AvgMs | Should -BeGreaterThan 0
        $log = [System.IO.File]::ReadAllText("$root\work\indexing_log.txt")
        $log | Should -Match "段階ごとの時間（秒）: 起動 .* / クロール .* / 取り込み .*（うちまとめ .*） / 仕上げのまとめ .* / gram 索引 .* / その他 .* / 全体 "
        $log | Should -Match "  \.docx 2 件: 抽出 "
        $log | Should -Match "  \.pptx 3 件: 抽出 .*（失敗 1・"
    }

    It "取り込みのスレッドで取り込んでも、取り込み一覧・本文インデックスのファイルはスレッドを使わないときと同じになる" {
        $single = newRoot
        writeTestSettings $single @(@{ name = "並列"; path = $source; enabled = $true })
        runIndexer $single @{ Workers = 0 } | Should -Be 0
        $parallel = newRoot
        writeTestSettings $parallel @(@{ name = "並列"; path = $source; enabled = $true })
        runIndexer $parallel @{ Workers = 3 } | Should -Be 0

        $expected = readTestStatus $single
        $actual = readTestStatus $parallel
        $actual.Rows.Count | Should -Be $expected.Rows.Count
        foreach ($key in $expected.Rows.Keys) {
            $actual.Rows[$key].状態 | Should -Be $expected.Rows[$key].状態
            $actual.Rows[$key].TSV数 | Should -Be $expected.Rows[$key].TSV数
        }
        (readContentIndexLines $parallel) -join "`n" | Should -BeExactly ((readContentIndexLines $single) -join "`n")
        # 取り込んだ TSV（元のファイルごとのフォルダ）・取り込み中の記録・一時フォルダは残さない
        (findIndexFoldersWithBooks "$parallel\work\store\content").Count | Should -Be 0
        @([IngestJournal]::new("$parallel\work\store\journal\ingest.tsv").ReadPrevious())[-1] | Should -Be "終"
        @(Get-ChildItem -LiteralPath "$parallel\work\publish" -Force -ErrorAction SilentlyContinue).Count | Should -Be 0
        Test-Path -LiteralPath (Join-Path "$parallel\work\tmp\$(getMachineKey)" "$PID") | Should -Be $false
        # 取り込み中は、前の版の片付けだけに使う場所（%TEMP%）ではなくワークスペースの tmp の下を使う
        $global:capturedTmpDir | Should -Not -BeNullOrEmpty
        $global:capturedTmpDir.StartsWith("$parallel\work\tmp", [System.StringComparison]::OrdinalIgnoreCase) | Should -Be $true
        # ${legacyTmpParent} の差し替えが取り込み中も効いていた（本物の %TEMP%\tebunko を指していない）ことを確かめる
        $global:capturedLegacyTmpParent | Should -Be (Join-Path $parallel "legacy_tmp")
        Test-Path -LiteralPath (Join-Path "$parallel\legacy_tmp" "$PID") | Should -Be $false
        $progress = readTestProgress
        $progress.Processed | Should -Be 5
        $progress.Failed | Should -Be 1
        (Get-Content -LiteralPath "$parallel\work\indexing_log.txt" -Raw) | Should -Match "3 個のスレッド、Excel・Word・PowerPoint のファイルはそれぞれ 1 つのスレッド"
    }

    It "取り込みのスレッドが読み込む部品が見つからなければ、続けられないエラーで 1 を返す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "並列"; path = $source; enabled = $true })
        # getPartLoad indexerLib が部品を見つけられないようにする（単一 .ps1 版の分岐に入れる。
        # bundledParts はこの breakpoint が止まった関数（invokeIndexerBody）のスコープだけに置き、他のテストに残さない）
        $broken = @{ Script = $runPath; Pattern = '^\s+\$dispatcher\.Start\('; Action = {
                Set-Variable -Name bundledParts -Value @{} -Scope 1
            }
        }
        runIndexer $root @{ Workers = 2 } @($broken) | Should -Be 1
        readTestError | Should -Match "indexerLib の部品が入っていません"
    }
}

Describe "getIngestWorkerCount" -Tag Unit {
    # インデックス作成の本体は lib.ps1 に入っていない（indexer_lib.ps1 から読み込む）
    BeforeAll {
        . $runPath
    }

    It "指定があればその数、無ければ設定、設定が 0 ならコア数から決める（ファイルの数より多くしない）" {
        getIngestWorkerCount 0 10 3 8 | Should -Be 0
        getIngestWorkerCount 2 10 3 8 | Should -Be 2
        getIngestWorkerCount -1 10 3 8 | Should -Be 3
        getIngestWorkerCount -1 10 9 8 | Should -Be 4
        getIngestWorkerCount -1 10 0 8 | Should -Be 4
        getIngestWorkerCount -1 10 0 2 | Should -Be 1
        getIngestWorkerCount -1 2 0 8 | Should -Be 2
        getIngestWorkerCount 3 0 0 8 | Should -Be 0
    }
}

Describe "取り込みのスレッドのスクリプト（ingestWorkerScript）" -Tag Io {
    # 取り込みのスレッドで動くスクリプトを、このスレッドで直接動かして確かめる（スレッドの中の動きはブレークポイントで止められないため）。
    # 本物は IngestDispatcher が indexerLib の Prelude をつないでから動かすため（newWorkerTmpDir など indexer_lib.ps1 の関数を使う）、
    # ここでは直接呼ぶため indexer_lib.ps1 をまとめて読み込んでおく
    BeforeAll {
        . "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"
    }

    It "取り込み待ちの列のファイルを取り込んで結果の列に入れ、列が閉じられたら Office を片づけて終わる" {
        $root = Join-Path $TestDrive "worker_direct"
        foreach ($dir in "store\content", "tmp", "publish") {
            [System.IO.Directory]::CreateDirectory("$root\$dir") | Out-Null
        }
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $results = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $tasks.Add(@{ RelPath = "営業\議事録.docx"; SourcePath = $docxSource })
        $tasks.Add(@{ RelPath = "営業\無い.docx"; SourcePath = (Join-Path $TestDrive "無い.docx") })
        $tasks.CompleteAdding()
        $settings = @{
            Lib = getPartLoad indexerLib
            WorkDir = $root; TmpDir = "$root\tmp"; PublishDir = "$root\publish"
            FileTimeoutMinutes = 10; RestartInterval = 1
            OfficePids = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,object]'
        }
        $priority = [System.Threading.Thread]::CurrentThread.Priority
        try {
            & ${ingestWorkerScript} $settings $tasks $results 7
        } finally {
            [System.Threading.Thread]::CurrentThread.Priority = $priority
        }

        $results.Count | Should -Be 2
        $done = $results.Take()
        $done.RelPath | Should -Be "営業\議事録.docx"
        $done.Ok | Should -Be $true
        $done.TsvCount | Should -BeGreaterThan 0
        # 取り込んだ TSV は、渡したインデックスのフォルダの、元のファイルごとのフォルダに置く
        @([System.IO.Directory]::GetFiles("$root\store\content\営業\議事録.docx", "*.tsv")).Count | Should -Be $done.TsvCount
        # 一時フォルダはスレッドごとに分ける
        [System.IO.Directory]::Exists("$root\tmp\w7") | Should -Be $true
        $failed = $results.Take()
        $failed.Ok | Should -Be $false
        $failed.Message | Should -Not -BeNullOrEmpty
    }

    It "終わりの待ちには、設定の Channel の Stop を見る判定を渡す（立てる前は偽・立てたあとは真）" {
        $root = Join-Path $TestDrive "worker_wait"
        foreach ($dir in "store\content", "tmp", "publish") {
            [System.IO.Directory]::CreateDirectory("$root\$dir") | Out-Null
        }
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $results = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $tasks.CompleteAdding()
        $global:waitChannel = [pscustomobject]@{ Stop = $false }
        $global:waitAnswers = New-Object System.Collections.ArrayList
        $settings = @{
            Lib = getPartLoad indexerLib
            WorkDir = $root; TmpDir = "$root\tmp"; PublishDir = "$root\publish"
            FileTimeoutMinutes = 10; RestartInterval = 1
            OfficePids = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,object]'
            Channel = $global:waitChannel
        }
        # 見張りは本物のまま動かす（Pids が空なので何も止めない。ingestWorkerScript が Start と Stop を呼ぶ）
        Mock stopAllApps {}
        Mock waitKeptApps {
            [void]$global:waitAnswers.Add((& $shouldStop))
            $global:waitChannel.Stop = $true
            [void]$global:waitAnswers.Add((& $shouldStop))
        }
        $priority = [System.Threading.Thread]::CurrentThread.Priority
        try {
            & ${ingestWorkerScript} $settings $tasks $results 8
            @($global:waitAnswers) | Should -Be @($false, $true)
        } finally {
            [System.Threading.Thread]::CurrentThread.Priority = $priority
            Remove-Variable -Name waitChannel, waitAnswers -Scope Global -ErrorAction SilentlyContinue
        }
    }
}

Describe "取り込みのスレッドのスクリプト（諦めたワーカーの配線）" -Tag Io {
    BeforeAll {
        . "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"
    }

    It "ワーカーの状態と利用者が使った印の入れ物を、見張りに渡す。諦めていなければ Office を片づける。諦めていれば片づけない（<title>）" -TestCases @(
        @{ title = "諦めていない"; abandoned = $false; expectedStops = 1 }
        @{ title = "諦めた"; abandoned = $true; expectedStops = 0 }
    ) {
        param ($title, $abandoned, $expectedStops)
        $root = Join-Path $TestDrive "worker_wiring_$abandoned"
        foreach ($dir in "store\content", "tmp", "publish") {
            [System.IO.Directory]::CreateDirectory("$root\$dir") | Out-Null
        }
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $results = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $tasks.CompleteAdding()
        $workerState = [hashtable]::Synchronized(@{ Abandoned = $abandoned; Handled = $false })
        $users = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,object]'
        $settings = @{
            Lib = getPartLoad indexerLib
            WorkDir = $root; TmpDir = "$root\tmp"; PublishDir = "$root\publish"
            FileTimeoutMinutes = 10; RestartInterval = 1
            OfficePids = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,object]'
            Channel = [pscustomobject]@{ Stop = $false }
            WorkerState = $workerState; UserOfficePids = $users
        }
        Mock stopAllApps {}
        Mock waitKeptApps {}
        $priority = [System.Threading.Thread]::CurrentThread.Priority
        $previousAbandon = $script:officeWatchdog.State.Abandon
        $previousUsers = $script:officeWatchdog.State.Users
        try {
            & ${ingestWorkerScript} $settings $tasks $results 9
            # 見張りの State が、司令と同じ入れ物を持つ（別の入れ物にすると、印も諦めた合図も司令に届かない）
            [object]::ReferenceEquals($script:officeWatchdog.State.Abandon, $workerState) | Should -Be $true
            [object]::ReferenceEquals($script:officeWatchdog.State.Users, $users) | Should -Be $true
        } finally {
            [System.Threading.Thread]::CurrentThread.Priority = $priority
            # 後のテストに諦めたワーカーの状態を残さない
            $script:workerState = $null
            $script:officeWatchdog.State.Abandon = $previousAbandon
            $script:officeWatchdog.State.Users = $previousUsers
        }
        Should -Invoke stopAllApps -Times $expectedStops -Exactly -Scope It
        Should -Invoke waitKeptApps -Times $expectedStops -Exactly -Scope It
    }
}

Describe "runIngestWorker・invokeIngestTask（ワーカー）" -Tag Io {
    BeforeAll {
        . "${scriptsDir}\shared\office\office_app.ps1"
        . "${scriptsDir}\tebunko\indexer\extract_office.ps1"
        . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"
        . $runPath
    }

    AfterEach { $script:officeUnavailable = $false }

    It "列の順に取り込み、結果を 1 件に 1 つ返して、列が閉じられたら終わる" {
        Mock invokeIngestTask { param ($task) @{ RelPath = $task.RelPath; Ok = $true; Reroute = $false; TimedOut = $false } }
        Mock stopAllApps { }
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        foreach ($path in "a.xlsx", "b.xlsx", "c.xlsx") { $tasks.Add(@{ RelPath = $path }) }
        $tasks.CompleteAdding()
        $results = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        runIngestWorker $tasks $results 10 2
        @($results.ToArray() | ForEach-Object { $_.RelPath }) -join "," | Should -Be "a.xlsx,b.xlsx,c.xlsx"
        # Office を持つスレッドは、決まった数を取り込むたびに Office を起動し直す
        Should -Invoke stopAllApps -Times 1 -Exactly -Scope It
    }

    It "読み取りのスレッドは Office を起動し直さない" {
        Mock invokeIngestTask { param ($task) @{ RelPath = $task.RelPath; Ok = $true; Reroute = $false; TimedOut = $false } }
        Mock stopAllApps { }
        $script:officeUnavailable = $true
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        foreach ($path in "a.docx", "b.docx") { $tasks.Add(@{ RelPath = $path }) }
        $tasks.CompleteAdding()
        runIngestWorker $tasks (New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]') 10 1
        Should -Invoke stopAllApps -Times 0 -Exactly -Scope It
    }

    It "Office を使わずに読んだファイルは、起動し直しの数に数えない（Office を使うファイルの数で決める）" {
        Mock invokeIngestTask {
            param ($task)
            $script:officeUnavailable = ($task.RelPath -like "*.pdf")
            @{ RelPath = $task.RelPath; Ok = $true; Reroute = $false; TimedOut = $false }
        }
        Mock stopAllApps { }
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        foreach ($path in "a.pdf", "a.xlsx", "b.pdf", "b.xlsx") { $tasks.Add(@{ RelPath = $path }) }
        $tasks.CompleteAdding()
        runIngestWorker $tasks (New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]') 10 2
        # 4 件すべてを数えると 2 回（2 件目と 4 件目）になる。Office を使う 2 件だけなら 1 回
        Should -Invoke stopAllApps -Times 1 -Exactly -Scope It
    }

    It "後回し（Postponed）は取り込んだ件数に数えず、起動し直しを早めない" {
        Mock invokeIngestTask {
            param ($task)
            @{ RelPath = $task.RelPath; Ok = $true; Reroute = $false; Postponed = ($task.RelPath -like "後回し*"); TimedOut = $false }
        }
        Mock stopAllApps { }
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        foreach ($path in "後回し1.pptx", "後回し2.pptx", "a.xlsx", "b.xlsx") { $tasks.Add(@{ RelPath = $path }) }
        $tasks.CompleteAdding()
        $results = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        runIngestWorker $tasks $results 10 2
        @($results.ToArray()).Count | Should -Be 4
        # 後回しを数えると、起動し直し（2件ごと）が早まって2回になる
        Should -Invoke stopAllApps -Times 1 -Exactly -Scope It
    }

    It "中止（Stopped）で返したファイルは、取り込んだ件数に数えず、Office を起動し直さない。列の残りも Stopped で返す" {
        Mock invokeIngestTask {
            param ($task, $timeout, $channel)
            @{ RelPath = $task.RelPath; Ok = $false; Reroute = $false; TimedOut = $false; Stopped = ($task.RelPath -ne "a.xlsx") }
        }
        Mock stopAllApps { }
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        foreach ($path in "a.xlsx", "b.xlsx", "c.xlsx", "d.xlsx") { $tasks.Add(@{ RelPath = $path }) }
        $tasks.CompleteAdding()
        $results = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $channel = [hashtable]::Synchronized(@{ Stop = $true })
        runIngestWorker $tasks $results 10 2 $channel
        @($results.ToArray() | Where-Object { $_.Stopped }).Count | Should -Be 3
        # 数えると 2 件目で起動し直しが起きる
        Should -Invoke stopAllApps -Times 0 -Exactly -Scope It
        Should -Invoke invokeIngestTask -Times 4 -Exactly -Scope It -ParameterFilter { $channel.Stop }
    }
}

Describe "invokeIngestTask（Office が要る）" -Tag Io {
    # invokeIngestTask そのものを確かめるため、invokeIngestTask を Mock する上の Describe と分ける
    BeforeAll {
        . "${scriptsDir}\shared\office\office_reader.ps1"
        . "${scriptsDir}\shared\office\office_embedded.ps1"
        . "${scriptsDir}\shared\office\office_app.ps1"
        . "${scriptsDir}\tebunko\indexer\extract_office.ps1"
        . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"
        . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"
        . "${scriptsDir}\tebunko\indexer\indexer_decide.ps1"
        . $runPath
        . "$PSScriptRoot\..\..\helpers\cfb.ps1"
    }

    AfterEach { $script:officeUnavailable = $false }

    It "読み取りのスレッドでIRMのファイルは、ingestFile を Mock せずに、回さず失敗として返す" {
        ${tmpDir} = Join-Path $TestDrive "irm_reader_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $irmSource = Join-Path $TestDrive "irm-reader.docx"
        newCompoundFile $irmSource @(@(([char]6 + "DataSpaces"), "DRMEncryptedDataSpace", "DRMEncryptedTransform"))
        Mock getApp { throw "Office は使わない" }
        $script:officeUnavailable = $true

        $result = invokeIngestTask @{ RelPath = "資料\irm-reader.docx"; SourcePath = $irmSource } 10
        $result.Ok | Should -Be $false
        $result.Reroute | Should -Be $false
        $result.Postponed | Should -Be $false
        $result.Message | Should -Be "IRM・秘密度ラベルで暗号化されているため更新できません。"
        Should -Invoke getApp -Times 0 -Exactly -Scope It
    }

    It "「Office が要る」の例外なら、失敗にせず回し直し（Reroute）として返し、Office も終了しない" {
        ${tmpDir} = Join-Path $TestDrive "reroute_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        # Mock の中からは、この Describe の BeforeAll で読み込んだ値が見えないため、global に置いて渡す
        $global:testOfficeRequiredMessage = ${officeRequiredMessage}
        Mock ingestFile { throw (New-Object System.OperationCanceledException $global:testOfficeRequiredMessage) }
        Mock stopApp { }
        $script:officeUnavailable = $true
        $result = invokeIngestTask @{ RelPath = "営業\中身が旧形式.docx"; SourcePath = "C:\data\中身が旧形式.docx" } 10
        Remove-Variable -Name testOfficeRequiredMessage -Scope Global
        $result.Reroute | Should -Be $true
        $result.Ok | Should -Be $false
        $result.Message | Should -BeNullOrEmpty
        Should -Invoke stopApp -Times 0 -Exactly -Scope It
    }

    It "利用者のPowerPointが使用中の例外なら、失敗にせず後回し（Postponed）として返し、Office も終了しない" {
        ${tmpDir} = Join-Path $TestDrive "postponed_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $global:testOfficeAppInUseMessage = ${officeAppInUseMessage}
        Mock ingestFile { throw (New-Object System.InvalidOperationException "PowerPoint$global:testOfficeAppInUseMessage") }
        Mock stopApp { }
        $result = invokeIngestTask @{ RelPath = "資料\発表.ppt"; SourcePath = "C:\data\発表.ppt" } 10
        Remove-Variable -Name testOfficeAppInUseMessage -Scope Global
        $result.Postponed | Should -Be $true
        $result.Ok | Should -Be $false
        $result.Message | Should -BeNullOrEmpty
        Should -Invoke stopApp -Times 0 -Exactly -Scope It
    }

    It "暗号化の判定でOfficeにまだ触れていない失敗（throwProtectionFailure）なら、失敗にするがOfficeは終了しない" {
        ${tmpDir} = Join-Path $TestDrive "protection_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        Mock ingestFile { throwProtectionFailure "テスト用の失敗" }
        Mock stopApp { }
        $result = invokeIngestTask @{ RelPath = "資料\irm.docx"; SourcePath = "C:\data\irm.docx" } 10
        $result.Ok | Should -Be $false
        $result.Message | Should -Be "テスト用の失敗"
        Should -Invoke stopApp -Times 0 -Exactly -Scope It
    }

    It "ふつうの失敗では、今までどおりOfficeを終了する" {
        ${tmpDir} = Join-Path $TestDrive "ordinary_fail_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        Mock ingestFile { throw "ふつうの失敗です。" }
        Mock stopApp { }
        $result = invokeIngestTask @{ RelPath = "資料\ふつう.docx"; SourcePath = "C:\data\ふつう.docx" } 10
        $result.Ok | Should -Be $false
        Should -Invoke stopApp -Times 1 -Exactly -Scope It
    }

    It "ファイルごとに task.Office で Office の使用を切り替える。Deadline は Office を使うファイルの間だけ入れ、使わないファイルでは MaxValue のまま（持っている Office を止めない）" {
        ${tmpDir} = Join-Path $TestDrive "office_switch_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $global:testSwitchWatch = $script:officeWatchdog
        $global:testSwitchSeen = New-Object System.Collections.Generic.List[object]
        Mock ingestFile {
            $global:testSwitchSeen.Add(@{ Unavailable = $script:officeUnavailable; Deadline = $global:testSwitchWatch.State.Deadline })
            throw "記録だけして失敗にする"
        }
        Mock stopApp { }
        foreach ($office in $false, $true, $false, $true) {
            [void](invokeIngestTask @{ RelPath = "資料\a.docx"; SourcePath = "C:\data\a.docx"; Office = $office } 10)
        }
        $seen = $global:testSwitchSeen.ToArray()
        Remove-Variable -Name testSwitchSeen, testSwitchWatch -Scope Global
        @($seen | ForEach-Object { $_.Unavailable }) | Should -Be @($true, $false, $true, $false)
        @($seen | ForEach-Object { $_.Deadline -eq [datetime]::MaxValue }) | Should -Be @($true, $false, $true, $false)
        $script:officeWatchdog.State.Deadline | Should -Be ([datetime]::MaxValue)
        # Office を使わないファイルの失敗では、Office を終了しない（使うファイルの 2 回だけ）
        Should -Invoke stopApp -Times 2 -Exactly -Scope It
    }

    It "Office を使わないファイルは、取り込み中に中止の合図が立っても読み終わるまで待ち、Stopped にも後回しにもしない" {
        ${tmpDir} = Join-Path $TestDrive "no_office_wait_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $channel = [hashtable]::Synchronized(@{ Stop = $false })
        $global:testWaitChannel = $channel
        Mock ingestFile { $global:testWaitChannel.Stop = $true; 3 }
        Mock publishTsv { }
        Mock getBookDir { "dir" }
        Mock getExtractVersion { "1" }
        Mock stopApp { }
        $result = invokeIngestTask @{ RelPath = "資料\a.pdf"; SourcePath = "C:\data\a.pdf"; Office = $false } 10 $channel
        Remove-Variable -Name testWaitChannel -Scope Global
        $result.Ok | Should -Be $true
        $result.Stopped | Should -Be $false
        $result.Postponed | Should -Be $false
        $result.TimedOut | Should -Be $false
        $result.TsvCount | Should -Be 3
        Should -Invoke stopApp -Times 0 -Exactly -Scope It
    }

    It "中止の合図が立っているときは、ingestFile を呼ばずに Stopped を返す（失敗にしない）" {
        ${tmpDir} = Join-Path $TestDrive "stop_before_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        Mock ingestFile { }
        Mock stopApp { }
        $channel = [hashtable]::Synchronized(@{ Stop = $true })
        $result = invokeIngestTask @{ RelPath = "資料\待ち.docx"; SourcePath = "C:\data\待ち.docx" } 10 $channel
        $result.Stopped | Should -Be $true
        $result.Ok | Should -Be $false
        $result.Message | Should -BeNullOrEmpty
        Should -Invoke ingestFile -Times 0 -Exactly -Scope It
    }

    It "取り込み中に中止で見張りがプロセスを止めた（Cancelled）ときは、失敗にせず Stopped を返す。中止の合図だけが立った例外のときも同じ" -TestCases @(
        @{ via = "Cancelled" }
        @{ via = "Stop" }
    ) {
        param ($via)
        ${tmpDir} = Join-Path $TestDrive "stop_during_tmp_$via"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $channel = [hashtable]::Synchronized(@{ Stop = $false })
        $global:testCancelWatch = $script:officeWatchdog
        $global:testCancelChannel = $channel
        $global:testCancelVia = $via
        Mock ingestFile {
            if ($global:testCancelVia -eq "Cancelled") { $global:testCancelWatch.State.Cancelled = $true } else { $global:testCancelChannel.Stop = $true }
            throw "プロセスを止められた"
        }
        Mock stopApp { }
        try {
            $result = invokeIngestTask @{ RelPath = "資料\途中.docx"; SourcePath = "C:\data\途中.docx" } 10 $channel
        } finally {
            $global:testCancelWatch.State.Cancelled = $false
            Remove-Variable -Name testCancelWatch, testCancelChannel, testCancelVia -Scope Global -ErrorAction SilentlyContinue
        }
        $result.Stopped | Should -Be $true
        $result.Ok | Should -Be $false
        $result.Message | Should -BeNullOrEmpty
    }

    It "元のファイルの 1 行が検索の上限を超えるときは、置き場へ移さず、決めた文言の失敗として返す" {
        . "${scriptsDir}\tebunko\indexer\extract_text.ps1"
        $workspace = newTestWorkspace @{} "$TestDrive\fit_task_ws"
        ${tmpDir} = Join-Path $TestDrive "fit_task_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $source = Join-Path $TestDrive "fit_task.txt"
        [System.IO.File]::WriteAllText($source, ("前の行`r`n" + ("あ" * 1000) + "`r`n後の行"))
        ${contentIndexSearchMaxBytes} = 1KB

        $result = invokeIngestTask @{ RelPath = "営業\fit_task.txt"; SourcePath = $source } 10

        $result.Ok | Should -Be $false
        $result.Message | Should -Be "ファイルの文字が多すぎるため更新できません（上限 64 MB）。"
        [System.IO.Directory]::Exists("$($workspace.IndexDir)\営業\fit_task.txt") | Should -Be $false
    }

    It "元のファイルの文字が検索の上限を超えても、1 行ずつが上限以内なら失敗にせず置き場へ移す（書き出しで塊に分ける）" {
        . "${scriptsDir}\tebunko\indexer\extract_text.ps1"
        $workspace = newTestWorkspace @{} "$TestDrive\fit_task_split_ws"
        ${tmpDir} = Join-Path $TestDrive "fit_task_split_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $source = Join-Path $TestDrive "fit_task_split.txt"
        [System.IO.File]::WriteAllText($source, ((1..300 | ForEach-Object { "長い行の文字 $_" }) -join "`r`n"))
        ${contentIndexSearchMaxBytes} = 1KB

        $result = invokeIngestTask @{ RelPath = "営業\fit_task_split.txt"; SourcePath = $source } 10

        $result.Ok | Should -Be $true
        [System.IO.Directory]::Exists("$($workspace.IndexDir)\営業\fit_task_split.txt") | Should -Be $true
    }

    It "本文（word/document.xml）が部品のサイズの上限を超えた .docx は、Office を使わず簡潔なメッセージで失敗にする" {
        ${tmpDir} = Join-Path $TestDrive "big_for_indexer_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $orig = $script:zipPartMaxBytes
        $script:zipPartMaxBytes = 10
        try {
            $path = Join-Path $TestDrive "big_for_indexer.docx"
            $stream = [System.IO.File]::Create($path)
            $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
            $writer = New-Object System.IO.StreamWriter($zip.CreateEntry("word/document.xml").Open(), (New-Object System.Text.UTF8Encoding($false)))
            $writer.Write('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + ("あ" * 20) + "</w:body></w:document>")
            $writer.Dispose()
            $zip.Dispose()
            $stream.Dispose()

            Mock getApp { throw "Office は使わない" }
            $result = invokeIngestTask @{ RelPath = "資料\big_for_indexer.docx"; SourcePath = $path } 10
            $result.Ok | Should -Be $false
            $result.Reroute | Should -Be $false
            $result.Message | Should -Be "ファイルサイズが大きすぎるため更新できません。"
            Should -Invoke getApp -Times 0 -Exactly -Scope It
            # 原因を調べられるよう、画面・取り込み一覧には出さない部品名・大きさ・部品ごとか合計かを、インデックス作成のログにだけ書く
            $result.Log | Should -Match "word/document\.xml"
            $result.Log | Should -Match "部品ごと"
        } finally {
            $script:zipPartMaxBytes = $orig
        }
    }

    It "ふつうの失敗（ZipSizeLimitException ではない例外）では、サイズの上限の詳細をログに書かず、今までどおりの表示にする" {
        ${tmpDir} = Join-Path $TestDrive "ordinary_fail_log_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        Mock ingestFile { throw "ふつうの失敗です。" }
        Mock stopApp { }
        $result = invokeIngestTask @{ RelPath = "資料\ふつう2.docx"; SourcePath = "C:\data\ふつう2.docx" } 10
        $result.Ok | Should -Be $false
        $result.Message | Should -Be "ふつうの失敗です。"
        $result.Log | Should -Not -Match "サイズの上限"
    }

    It "<app> の取り込みのあとに利用者が開いたファイルがあれば、閉じずに渡し、次の取り込みは新しく起動する" -TestCases @(
        @{ app = "Excel"; ext = "xlsx"; items = "Workbooks"; shown = $true }
        @{ app = "Word"; ext = "docx"; items = "Documents"; shown = $true }
    ) {
        param ($app, $ext, $items, $shown)
        ${tmpDir} = Join-Path $TestDrive "handover_tmp_$app"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $script:officeOwnDir = ${tmpDir}
        $script:onOfficeHandOver = ${officeHandOverNotice}
        $global:handoverApp = $app
        $global:handoverFake = New-Object psobject -Property @{
            Visible = $false; UserControl = $false; DisplayAlerts = $false; EnableEvents = $false; ScreenUpdating = $false
            AskToUpdateLinks = $false; AutomationSecurity = 3
        }
        $global:handoverFake | Add-Member -MemberType NoteProperty -Name $items -Value @([pscustomobject]@{ FullName = "C:\利用者\山田の資料.$ext" })
        $global:handoverFake | Add-Member -MemberType ScriptMethod -Name Quit -Value { throw "Quit は呼ばない" }
        $global:handoverStarts = New-Object System.Collections.ArrayList
        Mock ingestFile {
            [void]$global:handoverStarts.Add(-not $script:apps.ContainsKey($global:handoverApp))
            $script:apps[$global:handoverApp] = @{ Com = $global:handoverFake; Pid = 4242; Shared = $false }
            return 1
        }
        Mock publishTsv { }
        Mock getBookDir { "x" }
        Mock getExtractVersion { "1" }
        try {
            $first = invokeIngestTask @{ RelPath = "資料\a.$ext"; SourcePath = "C:\data\a.$ext" } 10
            $first.Ok | Should -Be $true
            $first.Log | Should -Match "開かれたファイルがあるため $app を利用者に渡しました"
            $script:apps.ContainsKey($app) | Should -Be $false
            $global:handoverFake.Visible | Should -Be $shown
            $global:handoverFake.AutomationSecurity | Should -Be 1

            # 渡したあとの次のファイルは、新しいアプリを起動する（渡したアプリを使い続けない）
            $global:handoverFake.$items = @()
            [void](invokeIngestTask @{ RelPath = "資料\b.$ext"; SourcePath = "C:\data\b.$ext" } 10)
            @($global:handoverStarts) | Should -Be @($true, $true)
        } finally {
            $script:apps.Remove($app)
            $script:officeOwnDir = $null
            $script:onOfficeHandOver = $null
            Remove-Variable -Name handoverFake, handoverStarts, handoverApp -Scope Global
        }
    }

    It "PowerPoint の取り込みのあとに利用者が開いたファイルがあれば、閉じずに渡し、次のファイルは後回しにする（渡した PowerPoint が利用者のものとして残るため）" {
        ${tmpDir} = Join-Path $TestDrive "handover_tmp_PowerPoint"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $script:officeOwnDir = ${tmpDir}
        $script:onOfficeHandOver = ${officeHandOverNotice}
        $global:handoverFake = New-Object psobject -Property @{
            Visible = $false; DisplayAlerts = 1; AutomationSecurity = 3
            Presentations = @([pscustomobject]@{ FullName = "C:\利用者\山田の資料.pptx" })
        }
        $global:handoverFake | Add-Member -MemberType ScriptMethod -Name Quit -Value { throw "Quit は呼ばない" }
        # 1 回目の起動の前後は、まだ PowerPoint がいない → 1 つ増える。そのあとは、渡した PowerPoint が残っている
        $global:handoverPids = New-Object System.Collections.Queue
        foreach ($ids in @(@(), @(), @(4242))) { $global:handoverPids.Enqueue($ids) }
        Mock getOwnSessionProcessIds {
            if ($global:handoverPids.Count -gt 0) { return , @($global:handoverPids.Dequeue()) }
            return , @(4242)
        }
        Mock New-Object { $global:handoverFake } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        Mock ingestFile { [void](getApp "PowerPoint"); return 1 }
        Mock publishTsv { }
        Mock getBookDir { "x" }
        Mock getExtractVersion { "1" }
        try {
            $first = invokeIngestTask @{ RelPath = "資料\a.pptx"; SourcePath = "C:\data\a.pptx" } 10
            $first.Ok | Should -Be $true
            $first.Log | Should -Match "開かれたファイルがあるため PowerPoint を利用者に渡しました"
            $script:apps.ContainsKey("PowerPoint") | Should -Be $false
            # PowerPoint は窓を隠していないため、窓は触らない。マクロの設定は戻る
            $global:handoverFake.Visible | Should -Be $false
            $global:handoverFake.AutomationSecurity | Should -Be 1

            # 渡した PowerPoint が残っているため、次のファイルは起動せず後回しにする
            $second = invokeIngestTask @{ RelPath = "資料\b.pptx"; SourcePath = "C:\data\b.pptx" } 10
            $second.Postponed | Should -Be $true
            $second.Ok | Should -Be $false
            $script:apps.ContainsKey("PowerPoint") | Should -Be $false
            Should -Invoke New-Object -Times 1 -Exactly -Scope It -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
        } finally {
            $script:apps.Remove("PowerPoint")
            $script:officeOwnDir = $null
            $script:onOfficeHandOver = $null
            Remove-Variable -Name handoverFake, handoverPids -Scope Global
        }
    }

    It "渡すときに窓や設定を戻しきれなければ、ログは「窓や設定を戻しきれませんでした」の文言になり、Excel は終了させず残す" {
        ${tmpDir} = Join-Path $TestDrive "handover_fail_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $script:officeOwnDir = ${tmpDir}
        $script:onOfficeHandOver = ${officeHandOverNotice}
        $script:officeKeptApps = @()
        $global:handoverFake = New-Object psobject -Property @{
            UserControl = $false; DisplayAlerts = $false; EnableEvents = $false; ScreenUpdating = $false
            AskToUpdateLinks = $false; AutomationSecurity = 3
            Workbooks = @([pscustomobject]@{ FullName = "C:\利用者\山田の資料.xlsx" })
        }
        $global:handoverFake | Add-Member -MemberType ScriptProperty -Name Visible -Value { $false } -SecondValue { throw "設定できません" }
        $global:handoverFake | Add-Member -MemberType ScriptMethod -Name Quit -Value { throw "Quit は呼ばない" }
        Mock ingestFile {
            $script:apps["Excel"] = @{ Com = $global:handoverFake; Pid = 4242; Shared = $false }
            return 1
        }
        Mock publishTsv { }
        Mock getBookDir { "x" }
        Mock getExtractVersion { "1" }
        Mock Stop-Process { }
        try {
            $result = invokeIngestTask @{ RelPath = "資料\a.xlsx"; SourcePath = "C:\data\a.xlsx" } 10
            $result.Ok | Should -Be $true
            $result.Log | Should -Match "窓や設定を戻しきれませんでした"
            $result.Log | Should -Not -Match "(?m)利用者に渡しました$"
            $script:apps.ContainsKey("Excel") | Should -Be $false
            @($script:officeKeptApps).Count | Should -Be 1
            Should -Invoke Stop-Process -Times 0 -Exactly -Scope It
        } finally {
            $script:apps.Remove("Excel")
            $script:officeKeptApps = @()
            $script:officeOwnDir = $null
            $script:onOfficeHandOver = $null
            Remove-Variable -Name handoverFake -Scope Global
        }
    }

    It "Word のファイルの一覧を読めなければ、ログは「一覧を読めなかったため、終了させずに残しています」の文言になり、Word は終了させず残す" {
        ${tmpDir} = Join-Path $TestDrive "handover_keep_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
        $script:officeOwnDir = ${tmpDir}
        $script:onOfficeHandOver = ${officeHandOverNotice}
        $script:officeKeptApps = @()
        $global:handoverFake = New-Object psobject -Property @{ Visible = $false; DisplayAlerts = 0; AutomationSecurity = 3 }
        $global:handoverFake | Add-Member -MemberType ScriptMethod -Name Quit -Value { throw "Quit は呼ばない" }
        # 一覧を読めない（COM が呼び出しを拒んだ）ことにする
        Mock getWorkbookSplit { throw "呼び出しを拒まれました" }
        Mock ingestFile {
            $script:apps["Word"] = @{ Com = $global:handoverFake; Pid = 4242; Shared = $false }
            return 1
        }
        Mock publishTsv { }
        Mock getBookDir { "x" }
        Mock getExtractVersion { "1" }
        Mock Stop-Process { }
        try {
            $result = invokeIngestTask @{ RelPath = "資料\a.docx"; SourcePath = "C:\data\a.docx" } 10
            $result.Ok | Should -Be $true
            $result.Log | Should -Match "Word のファイルの一覧を読めなかったため、終了させずに残しています"
            $result.Log | Should -Not -Match "利用者に渡しました"
            $script:apps.ContainsKey("Word") | Should -Be $false
            @($script:officeKeptApps).Count | Should -Be 1
            Should -Invoke Stop-Process -Times 0 -Exactly -Scope It
        } finally {
            $script:apps.Remove("Word")
            $script:officeKeptApps = @()
            $script:officeOwnDir = $null
            $script:onOfficeHandOver = $null
            Remove-Variable -Name handoverFake -Scope Global
        }
    }
}

Describe "諦めたワーカー（制限時間を過ぎても利用者が使っている PowerPoint を止めなかったとき）" -Tag Io {
    # 実際の PowerPoint は使わない。監視が立てる合図（ワーカーの Abandoned）を、テストが立てる
    BeforeAll {
        . "${scriptsDir}\shared\office\office_reader.ps1"
        . "${scriptsDir}\shared\office\office_embedded.ps1"
        . "${scriptsDir}\shared\office\office_app.ps1"
        . "${scriptsDir}\tebunko\indexer\extract_office.ps1"
        . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"
        . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"
        . "${scriptsDir}\tebunko\indexer\indexer_decide.ps1"
        . $runPath
    }

    BeforeEach {
        $global:testWorkerState = [hashtable]::Synchronized(@{ Abandoned = $false; Handled = $false })
        $script:workerState = $global:testWorkerState
        ${tmpDir} = Join-Path $TestDrive "abandoned_tmp"
        [System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null
    }
    AfterEach {
        $script:workerState = $null
        $script:officeUnavailable = $false
        Remove-Variable -Name testWorkerState -Scope Global -ErrorAction SilentlyContinue
    }

    It "<title>" -TestCases @(
        @{ title = "諦めたワーカーのファイルは、取り込まずに後回しにする"; when = "before" }
        @{ title = "取り込みが戻ったときに諦めていたら、書き出さずに後回しにする（司令と同じフォルダを書き換えない）"; when = "after" }
        @{ title = "取り込みが失敗して戻ったときに諦めていても、失敗にせず、アプリも終了させない"; when = "throw" }
    ) {
        param ($title, $when)
        $global:testWorkerWhen = $when
        Mock ingestFile {
            if ($global:testWorkerWhen -eq "throw") { $global:testWorkerState.Abandoned = $true; throw "固まった呼び出しが、失敗で戻った" }
            if ($global:testWorkerWhen -eq "after") { $global:testWorkerState.Abandoned = $true }
            return 1
        }
        Mock publishTsv { }
        Mock stopApp { }
        Mock handOverForeignApp { }
        if ($when -eq "before") { $global:testWorkerState.Abandoned = $true }
        try {
            $result = invokeIngestTask @{ RelPath = "資料\発表.pptx"; SourcePath = "C:\data\発表.pptx" } 10
        } finally {
            Remove-Variable -Name testWorkerWhen -Scope Global -ErrorAction SilentlyContinue
        }
        $result.Postponed | Should -Be $true
        $result.Ok | Should -Be $false
        $result.Message | Should -BeNullOrEmpty
        Should -Invoke publishTsv -Times 0 -Exactly -Scope It
        Should -Invoke stopApp -Times 0 -Exactly -Scope It
        Should -Invoke handOverForeignApp -Times 0 -Exactly -Scope It
        Should -Invoke ingestFile -Times $(if ($when -eq "before") { 0 } else { 1 }) -Exactly -Scope It
    }

    It "諦めていないワーカーは、今までどおり書き出して成功にする" {
        Mock ingestFile { return 1 }
        Mock publishTsv { }
        Mock getExtractVersion { return "1" }
        Mock handOverForeignApp { }
        $script:officeUnavailable = $true
        $result = invokeIngestTask @{ RelPath = "資料\発表.pptx"; SourcePath = "C:\data\発表.pptx" } 10
        $result.Ok | Should -Be $true
        $result.Postponed | Should -Be $false
        Should -Invoke publishTsv -Times 1 -Exactly -Scope It
    }
}

Describe "completeIngestJob（取り込みの 1 件の終わり方）" -Tag Io {
    # 取り込みのワーカーの数が 0（司令のスレッドで取り込む形）と同じ形で、司令の結果の受け方だけを確かめる。
    # 取り込み一覧（StatusLedger）・書き出し待ち（PendingPublish）・進み具合（IndexingReporter）は本物、本文インデックスへの書き出しだけ Mock にする
    BeforeAll {
        . "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"
        . $runPath
        $script:jobCount = 0

        # 1 件を取り込み中にした状態の run を作る（folder の Busy は 1。Undispatched のときは Pending が 1 のまま）
        function newJobRun {
            param ([bool]$undispatched = $false)

            $script:jobCount++
            $dir = Join-Path $TestDrive "job$($script:jobCount)"
            [System.IO.Directory]::CreateDirectory($dir) | Out-Null
            $ledger = [StatusLedger]::new([pscustomobject]@{ StatusFile = "$dir\ingest_status.tsv" })
            $pending = [PendingPublish]::new()
            $folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\a.docx"))
            $pending.MarkFolder($folder)
            $pending.AddPending($folder)
            if (!$undispatched) {
                $pending.Dispatch($folder)
            }
            $channel = newIndexerChannel
            $options = @{ LookAhead = 200; WorkerScript = { }; PostponedReasonInUse = "使っていた"; PostponedReasonRunning = "起動していた" }
            $dispatcher = [IngestDispatcher]::new(@(), [string[]]@(), [string[]]@(), @{}, @{}, $options)
            return @{
                Ledger = $ledger; Pending = $pending; Reporter = [IndexingReporter]::new($channel); Dispatcher = $dispatcher; Journal = [IngestJournal]::new("$dir\journal.tsv")
                Previous = (New-Object 'System.Collections.Generic.Dictionary[string,object]'); Total = 3; SuccessCount = 0; PostponedCount = 0; CurrentPath = ""
                Channel = $channel; Dir = $dir; Folder = $folder
            }
        }

        function newJobEntry {
            param ($run)

            return @{ Row = (newStatusRow "営業\a.docx" "2026/10/01 10:00:00" "1" ${stateNew}); Number = 1; Folder = $run.Folder }
        }
    }

    BeforeEach {
        Mock writeIndexerLog { }
        Mock publishIndexFolders { return @() }
    }

    It "成功は、取り込み一覧に「取り込み済み」で足し、書き出し待ちのフォルダを本文インデックスへ書き出し、成功の数と進み具合を更新する" {
        $run = newJobRun
        $received = @{ Kind = "Completed"; Entry = (newJobEntry $run); Result = @{ RelPath = "営業\a.docx"; Ok = $true; TsvCount = 4; ExtractVersion = "7"; Log = "" } }
        completeIngestJob $run $received
        $run.SuccessCount | Should -Be 1
        $run.PostponedCount | Should -Be 0
        $row = $run.Ledger.ReadStatus().Rows["営業\a.docx"]
        $row.状態 | Should -Be ${stateDone}
        $row.TSV数 | Should -Be "4"
        $row.抽出版 | Should -Be "7"
        Should -Invoke publishIndexFolders -Times 1 -Exactly -Scope It
        $progress = readIndexingProgress $run.Channel
        $progress.Processed | Should -Be 1
        $progress.Remaining | Should -Be 2
        $progress.Failed | Should -Be 0
    }

    It "失敗は、取り込み一覧に「失敗」と原因を足し、今回の失敗に数える。書き出し待ちには入れない" {
        $run = newJobRun
        $received = @{ Kind = "Completed"; Entry = (newJobEntry $run); Result = @{ RelPath = "営業\a.docx"; Ok = $false; Message = "壊れている"; Log = "" } }
        completeIngestJob $run $received
        $run.SuccessCount | Should -Be 0
        $row = $run.Ledger.ReadStatus().Rows["営業\a.docx"]
        $row.状態 | Should -Be ${stateFailed}
        $row.エラー | Should -Be "壊れている"
        $run.Ledger.Failures.Count | Should -Be 1
        (readIndexingProgress $run.Channel).Failed | Should -Be 1
    }

    It "後回し（<Kind>）は、理由を失敗として見せ、前の行の TSV 数・取り込み日時・抽出版を残す。後回しの数に数える。渡していた仕事（Undispatched が偽）だけ、取り込み一覧に「未取り込み」の行を足す" -TestCases @(
        @{ Kind = "Postponed"; Undispatched = $false }
        @{ Kind = "RerouteAbandoned"; Undispatched = $false }
        @{ Kind = "Postponed"; Undispatched = $true }
    ) {
        param ($Kind, $Undispatched)
        $run = newJobRun $Undispatched
        $run.Previous["営業\a.docx"] = (newStatusRow "営業\a.docx" "2026/09/01 10:00:00" "1" ${stateDone} "5" "2026/09/02 09:00:00" "" "6")
        $received = @{ Kind = $Kind; Entry = (newJobEntry $run); Reason = "使っていた"; Undispatched = $Undispatched }
        completeIngestJob $run $received
        $run.PostponedCount | Should -Be 1
        $run.SuccessCount | Should -Be 0
        $entryRow = $received.Entry.Row
        $entryRow.状態 | Should -Be ${stateNew}
        $entryRow.エラー | Should -Be "使っていた"
        $entryRow.TSV数 | Should -Be "5"
        $entryRow.取り込み日時 | Should -Be "2026/09/02 09:00:00"
        $entryRow.抽出版 | Should -Be "6"
        $run.Ledger.Failures.Count | Should -Be 1
        $run.Ledger.Failures[0].Message | Should -Be "使っていた"
        # 取り込み一覧の行は「未取り込み」（「失敗」にすると次の更新で飛ばされる）。渡さずに残したものは足さない
        $run.Ledger.ReadStatus().Rows.ContainsKey("営業\a.docx") | Should -Be (!$Undispatched)
        if (!$Undispatched) {
            $run.Ledger.ReadStatus().Rows["営業\a.docx"].状態 | Should -Be ${stateNew}
        }
        # 取り込み中（渡していない）の数を戻すため、書き出し待ちのフォルダは揃ったものとして書き出す
        Should -Invoke publishIndexFolders -Times 1 -Exactly -Scope It
    }

    It "中止（Stopped）は、失敗には足さず、取り込み一覧に今の行を足す（ジャーナルが、この仕事は終わったと読める）。最初の 1 件だけ中止のログを出す" -TestCases @(
        @{ First = $true; Lines = 3 }
        @{ First = $false; Lines = 1 }
    ) {
        param ($First, $Lines)
        $run = newJobRun
        $received = @{ Kind = "Stopped"; Entry = (newJobEntry $run); RelPath = "営業\a.docx"; First = $First }
        completeIngestJob $run $received
        $run.SuccessCount | Should -Be 0
        $run.PostponedCount | Should -Be 0
        $run.Ledger.Failures.Count | Should -Be 0
        $run.Ledger.ReadStatus().Rows["営業\a.docx"].状態 | Should -Be ${stateNew}
        # 最初の 1 件は（空行・中止の知らせ・このファイル）の 3 回、ほかは（このファイル）の 1 回ログを書く
        Should -Invoke writeIndexerLog -Times $Lines -Exactly -Scope It
    }
}

Describe "取り込みのジャーナルの復旧と書き出し（flushPendingPublish・recoverIngestJournal・writeStatusWithJournal）" -Tag Io {
    # 本文インデックスへの書き出し（publishIndexFolders）だけ Mock にして、書き出せたとき・書き出せなかったときのジャーナルの記録と、
    # 始めの復旧が次の回へ持ち越すフォルダを確かめる。ジャーナル・取り込み一覧・書き出し待ちは本物
    BeforeAll {
        . "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"
        . $runPath
        $script:recoverCount = 0

        function newRecoverEnv {
            # 前回のジャーナル（records）を書き、書き出し前のフォルダ（営業\a.xlsx の TSV を入れる元のファイルごとのフォルダ）を作った状態の部品を返す
            param (
                [string[]]$records = $null,
                [string[]]$folders = @("営業")
            )

            $script:recoverCount++
            $ws = [Workspace]::new((Join-Path $TestDrive "recover$($script:recoverCount)"))
            [System.IO.Directory]::CreateDirectory($ws.IndexDir) | Out-Null
            foreach ($name in $folders) {
                [System.IO.Directory]::CreateDirectory("$($ws.IndexDir)\$name\a.xlsx") | Out-Null
                [System.IO.File]::WriteAllText("$($ws.IndexDir)\$name\a.xlsx\S.tsv", "x")
            }
            $journal = [IngestJournal]::new($ws.JournalFile)
            if ($null -ne $records) {
                $journal.Reset([string[]](@("版`t2", "回`t2026/10/17 10:00:00`t0") + $records))
            }
            $channel = newIndexerChannel
            return @{
                Workspace = $ws; Journal = $journal; Pending = [PendingPublish]::new(); Channel = $channel
                Ledger = [StatusLedger]::new([pscustomobject]@{ StatusFile = "$($ws.Dir)\ingest_status.tsv" })
                Reporter = [IndexingReporter]::new($channel)
            }
        }

        function readJournalLines {
            param ($rc)

            $rc.Journal.Close()
            return @($rc.Journal.ReadPrevious())
        }
    }

    BeforeEach {
        Mock writeIndexerLog { }
        Mock publishIndexFolders { return @{ Count = 1; Skipped = @(); Published = @($pending.Keys) } }
    }

    It "flushPendingPublish は、書き出せたフォルダに「出」を足し、書き出せなかったフォルダ（Failed）を空のままにする" {
        $rc = newRecoverEnv
        $workspace = $rc.Workspace
        $rc.Pending.MarkFolder("$($workspace.IndexDir)\営業")
        $failed = New-Object System.Collections.Generic.List[string]
        flushPendingPublish $rc.Pending $rc.Journal -All -Failed $failed
        $failed.Count | Should -Be 0
        (readJournalLines $rc) | Should -Be @("出`t営業")
    }

    It "flushPendingPublish は、<name>のとき、書き出せなかったフォルダだけを Failed に足し、「出」はできたフォルダにだけ足す" -TestCases @(
        @{ name = "2 フォルダのうち 1 つの本文インデックスを書けなかった（Skipped）"; fail = "skip"; expectedFailed = @("営業"); expectedJournal = @("出`t開発") }
        @{ name = "書き出しが例外になった（全部が書き出せていない）"; fail = "throw"; expectedFailed = @("営業", "開発"); expectedJournal = @() }
        @{ name = "2 フォルダのうち 1 つを、置き場の外を指すパスで飛ばした（Skipped に出ず、Published にも無い）"; fail = "link"; expectedFailed = @("営業"); expectedJournal = @("出`t開発") }
        @{ name = "Failed を渡さずに呼び、本文インデックスを書けなかった（落ちず、「出」は書けたフォルダにだけ足す）"; fail = "skip"; noList = $true; expectedFailed = @(); expectedJournal = @("出`t開発") }
        @{ name = "Failed を渡さずに呼び、書き出しが例外になった（落ちず、「出」は足さない）"; fail = "throw"; noList = $true; expectedFailed = @(); expectedJournal = @() }
    ) {
        param ($name, $fail, $expectedFailed, $expectedJournal, $noList = $false)
        $rc = newRecoverEnv -folders @("営業", "開発")
        $workspace = $rc.Workspace
        if ($fail -eq "skip") {
            Mock publishIndexFolders { return @{ Count = 1; Skipped = @(@{ Rel = "営業"; Stage = "本文インデックス"; Reason = "書けない" }); Published = @($pending.Keys | Where-Object { $_ -notlike "*\営業" }) } }
        } elseif ($fail -eq "link") {
            Mock publishIndexFolders { return @{ Count = 1; Skipped = @(); Published = @($pending.Keys | Where-Object { $_ -notlike "*\営業" }) } }
        } else {
            Mock publishIndexFolders { throw "壊れた" }
        }
        $rc.Pending.MarkFolder("$($workspace.IndexDir)\営業")
        $rc.Pending.MarkFolder("$($workspace.IndexDir)\開発")
        $failed = New-Object System.Collections.Generic.List[string]
        if ($noList) {
            { flushPendingPublish $rc.Pending $rc.Journal -All } | Should -Not -Throw
        } else {
            flushPendingPublish $rc.Pending $rc.Journal -All -Failed $failed
        }
        @($failed | Sort-Object) | Should -Be @($expectedFailed | Sort-Object)
        @(readJournalLines $rc | Sort-Object) | Should -Be @($expectedJournal | Sort-Object)
    }

    It "前の回が終まで書いたが、営業の出が無い（書き出せなかった）とき、始めに書き出し直し、出を足して、次の回へ持ち越すものは無い。済んだ仕事は止まった扱いにしない" {
        $rc = newRecoverEnv @("始`t営業\a.xlsx`t1", "終")
        $workspace = $rc.Workspace
        $result = recoverIngestJournal $rc.Ledger $rc.Journal $rc.Pending $rc.Channel $rc.Reporter
        @($result.Resume).Count | Should -Be 0
        @($result.Unpublished).Count | Should -Be 0
        Should -Invoke publishIndexFolders -Times 1 -Exactly -Scope It
        (readJournalLines $rc)[-1] | Should -Be "出`t営業"
    }

    It "始めの書き出しが失敗したときは、そのフォルダを次の回へ持ち越す（Unpublished）。済んだ仕事は止まった扱いにしない" {
        Mock publishIndexFolders { return @{ Count = 0; Skipped = @(@{ Rel = "営業"; Stage = "本文インデックス"; Reason = "書けない" }); Published = @() } }
        $rc = newRecoverEnv @("始`t営業\a.xlsx`t1", "終")
        $workspace = $rc.Workspace
        $result = recoverIngestJournal $rc.Ledger $rc.Journal $rc.Pending $rc.Channel $rc.Reporter
        @($result.Resume).Count | Should -Be 0
        @($result.Unpublished) | Should -Be @("営業")
    }

    It "始めの書き出しが 2 フォルダのうち 1 つだけ失敗したときは、失敗した 1 つだけを次の回へ持ち越し、書き出せた 1 つは「出」を足す" {
        Mock publishIndexFolders { return @{ Count = 1; Skipped = @(@{ Rel = "開発"; Stage = "本文インデックス"; Reason = "書けない" }); Published = @($pending.Keys | Where-Object { $_ -notlike "*\開発" }) } }
        $rc = newRecoverEnv @("始`t営業\a.xlsx`t1", "始`t開発\b.xlsx`t1", "終") -folders @("営業", "開発")
        $workspace = $rc.Workspace
        $result = recoverIngestJournal $rc.Ledger $rc.Journal $rc.Pending $rc.Channel $rc.Reporter
        @($result.Unpublished) | Should -Be @("開発")
        (readJournalLines $rc)[-1] | Should -Be "出`t営業"
        @(readJournalLines $rc) | Should -Not -Contain "出`t開発"
    }

    It "中止を求められていれば、始めの書き出しはせず、書き出していないフォルダを次の回へ持ち越す" {
        $rc = newRecoverEnv @("未`t営業")
        $workspace = $rc.Workspace
        $rc.Channel.Stop = $true
        $result = recoverIngestJournal $rc.Ledger $rc.Journal $rc.Pending $rc.Channel $rc.Reporter
        @($result.Unpublished) | Should -Be @("営業")
        Should -Invoke publishIndexFolders -Times 0 -Exactly -Scope It
    }

    It "取り込み中に止まった仕事（始だけ）は、書き出しが失敗しても、やり直す仕事として返す" {
        Mock publishIndexFolders { return @{ Count = 0; Skipped = @(@{ Rel = "営業"; Stage = "本文インデックス"; Reason = "書けない" }); Published = @() } }
        $rc = newRecoverEnv @("始`t営業\a.xlsx`t1")
        $workspace = $rc.Workspace
        $result = recoverIngestJournal $rc.Ledger $rc.Journal $rc.Pending $rc.Channel $rc.Reporter
        @($result.Resume | ForEach-Object { "$($_.RelPath)=$($_.Count)" }) | Should -Be @("営業\a.xlsx=1")
        @($result.Unpublished) | Should -Be @("営業")
    }

    It "ジャーナルが読めない（無い）ときは、今ある書き出し前のフォルダ（<count> 件。前の版で止まった跡を含む）だけを書き出す" -TestCases @(
        @{ count = 0; folders = @() }
        @{ count = 2; folders = @("営業", "開発") }
    ) {
        param ($count, $folders)
        $rc = newRecoverEnv -folders $folders
        $workspace = $rc.Workspace
        $script:expectedKeys = (@($folders | ForEach-Object { "$($workspace.IndexDir)\$_" }) | Sort-Object) -join "|"
        $result = recoverIngestJournal $rc.Ledger $rc.Journal $rc.Pending $rc.Channel $rc.Reporter
        @($result.Unpublished).Count | Should -Be 0
        if ($count -eq 0) {
            Should -Invoke publishIndexFolders -Times 0 -Exactly -Scope It
        } else {
            Should -Invoke publishIndexFolders -Times 1 -Exactly -Scope It -ParameterFilter { ((@($pending.Keys) | Sort-Object) -join "|") -eq $script:expectedKeys }
        }
    }

    It "上限を超える本文インデックスのファイルがあるフォルダは、書き出し前の TSV が無くても、始めの書き出しに足す（上限を超えないフォルダは足さない）" {
        ${contentIndexSearchMaxBytes} = 1KB
        $rc = newRecoverEnv -folders @()
        $workspace = $rc.Workspace
        foreach ($name in "営業", "総務") {
            [System.IO.Directory]::CreateDirectory("$($workspace.IndexDir)\$name") | Out-Null
        }
        $big = ((1..200 | ForEach-Object { "巨大な行の単価 $_" }) -join "`n") + "`n"
        writeContentIndexFile "$($workspace.IndexDir)\営業\content_index.xlsx.001.tsv" (convertToContentIndexText @(@{ Name = "A.xlsx"; Places = @(@{ Place = "S"; Text = $big }) }))
        writeContentIndexFile "$($workspace.IndexDir)\総務\content_index.xlsx.001.tsv" (convertToContentIndexText @(@{ Name = "A.xlsx"; Places = @(@{ Place = "S"; Text = "x`n" }) }))
        $script:expectedKeys = "$($workspace.IndexDir)\営業"
        $result = recoverIngestJournal $rc.Ledger $rc.Journal $rc.Pending $rc.Channel $rc.Reporter
        @($result.Unpublished).Count | Should -Be 0
        Should -Invoke publishIndexFolders -Times 1 -Exactly -Scope It -ParameterFilter { ((@($pending.Keys) | Sort-Object) -join "|") -eq $script:expectedKeys }
    }

    It "applyStartupRequeued: 取り込み直しの元のファイルの行をすぐ「未取り込み」・エラー空にして保存し、印を空にする（行が無いものは無視する）" {
        $rc = newRecoverEnv -folders @()
        $ledger = [StatusLedger]::new([pscustomobject]@{ StatusFile = "$TestDrive\apply_status.tsv" })
        $folders = @([pscustomobject]@{ Path = "C:\共有\営業部"; Name = "営業" })
        writeStatusFile $folders @(
            (newStatusRow "営業\A.xlsx" "2024/01/01 00:00:00" "100" ${stateDone} "1" "2024/01/01 00:00:00" "" "1")
            (newStatusRow "営業\B.xlsx" "2024/01/01 00:00:00" "100" ${stateDone} "1" "2024/01/01 00:00:00" "" "1")
        ) "$TestDrive\apply_status.tsv"
        $status = readStatusFile "$TestDrive\apply_status.tsv"
        $pending = [PendingPublish]::new()
        $pending.AddRequeued([string[]]@("営業\A.xlsx", "営業\無い.xlsx"))

        applyStartupRequeued $ledger $pending $status

        $pending.Requeued.Count | Should -Be 0
        $status.Rows["営業\A.xlsx"].状態 | Should -Be ${stateNew}
        $saved = readStatusFile "$TestDrive\apply_status.tsv"
        $saved.Rows["営業\A.xlsx"].状態 | Should -Be ${stateNew}
        $saved.Rows["営業\B.xlsx"].状態 | Should -Be ${stateDone}
        $saved.Folders.Count | Should -Be 1
    }

    It "selectFinishStatusRows: 書き出しで外した上限超えのファイルの中の元のファイルは、この回に済になっていても「未取り込み」・エラー空に戻し、無くなった行は除く" {
        $ledger = [StatusLedger]::new([pscustomobject]@{ StatusFile = "$TestDrive\select_status.tsv" })
        $pending = [PendingPublish]::new()
        $pending.AddRequeued([string[]]@("営業\A.xlsx", ""))
        [void]$ledger.DroppedRows.Add("営業\C.xlsx")
        $rows = @(
            (newStatusRow "営業\A.xlsx" "2024/01/01 00:00:00" "100" ${stateDone} "1" "2024/01/01 00:00:00" "" "1")
            (newStatusRow "営業\B.xlsx" "2024/01/01 00:00:00" "100" ${stateDone} "1" "2024/01/01 00:00:00" "" "1")
            (newStatusRow "営業\C.xlsx" "2024/01/01 00:00:00" "100" ${stateDone} "1" "2024/01/01 00:00:00" "" "1")
        )
        $rows[0].エラー = "前のエラー"
        $selected = selectFinishStatusRows $rows $ledger $pending
        @($selected).Count | Should -Be 2
        $selected[0].相対パス | Should -Be "営業\A.xlsx"
        $selected[0].状態 | Should -Be ${stateNew}
        $selected[0].エラー | Should -Be ""
        $selected[1].状態 | Should -Be ${stateDone}
        $rows[0].状態 | Should -Be ${stateDone}
        $pending.Requeued.Count | Should -Be 1
    }

    It "前の版の ingesting.txt があれば、止まった仕事に足す。ジャーナルにも同じ仕事があるときは重ねない" {
        $rc = newRecoverEnv @("始`t営業\a.xlsx`t1")
        $workspace = $rc.Workspace
        writeListFile "$($workspace.Dir)\ingesting.txt" @("1`t営業\a.xlsx", "1`t営業\b.xlsx")
        $result = recoverIngestJournal $rc.Ledger $rc.Journal $rc.Pending $rc.Channel $rc.Reporter
        @($result.Resume | ForEach-Object { $_.RelPath } | Sort-Object) | Should -Be @("営業\a.xlsx", "営業\b.xlsx")
    }

    It "writeStatusWithJournal は、書き直しの前に「書」を足し、-Finish のときだけ書き直しのあとに「終」を足す" -TestCases @(
        @{ finish = $true; expected = @("書", "終") }
        @{ finish = $false; expected = @("書") }
    ) {
        param ($finish, $expected)
        Mock writeStatusKeepingDropped { }
        $rc = newRecoverEnv
        $workspace = $rc.Workspace
        if ($finish) {
            writeStatusWithJournal $rc.Journal $rc.Ledger @() @() @() @{} -Finish
        } else {
            writeStatusWithJournal $rc.Journal $rc.Ledger @() @() @() @{}
        }
        Should -Invoke writeStatusKeepingDropped -Times 1 -Exactly -Scope It
        (readJournalLines $rc) | Should -Be $expected
    }
}

Describe "indexer.ps1（無くなった元のファイルを、書き出せなかった回の次に外し直す）" -Tag Io {
    BeforeAll {
        function getContentIndexFile {
            param ([string]$root, [string]$name)
            return "$root\work\store\content\$name\content_index.docx.001.tsv"
        }
        function getAllContentIndexText {
            # name のインデックスの本文インデックスのファイルすべての中身をつなぐ
            param ([string]$root, [string]$name)
            return (@(Get-ChildItem -LiteralPath "$root\work\store\content\$name" -Recurse -Filter "content_index*.tsv" -File | ForEach-Object { readContentIndexText $_.FullName }) -join "`n")
        }
        function newIngestedRoot {
            # 「総務」を取り込んだ状態にして、設定のルートを返す
            param ([string]$folderName)
            $dir = newSourceFolder $folderName
            $root = newRoot
            writeTestSettings $root @(@{ name = $folderName; path = $dir; enabled = $true })
            runIndexer $root | Should -Be 0
            return @{ Root = $root; Dir = $dir; Name = $folderName }
        }
    }

    It "元のファイルを消した回に本文インデックスのファイルが壊れて書き出せなくても、一覧からは消え、記録に残り、直った次の回に外れて記録が消える" {
        $env = newIngestedRoot "削除記録"
        $root = $env.Root
        $recordPath = "$root\work\store\pending_remove.tsv"
        $file = getContentIndexFile $root $env.Name
        $good = [System.IO.File]::ReadAllBytes($file)
        (readContentIndexText $file) | Should -BeLike "*議事録.docx*"
        Remove-Item -LiteralPath "$($env.Dir)\議事録.docx"
        [System.IO.File]::WriteAllText($file, "壊れた中身`n")

        runIndexer $root | Should -Be 0

        (readTestStatus $root).Rows.ContainsKey("$($env.Name)\議事録.docx") | Should -Be $false
        @(readPendingRemovals $recordPath | ForEach-Object { "$($_.Folder)|$($_.Name)" }) | Should -Be @("$($env.Name)|議事録.docx")

        [System.IO.File]::WriteAllBytes($file, $good)
        runIndexer $root | Should -Be 0

        (getAllContentIndexText $root $env.Name) | Should -Not -BeLike "*議事録.docx*"
        (getAllContentIndexText $root $env.Name) | Should -BeLike "*提案.pptx*"
        Test-Path -LiteralPath $recordPath | Should -Be $false
    }

    It "前の回の記録を置いた状態から流すと、その名前が本文インデックスから外れ、置き場に無いフォルダの行は捨てられ、記録が無くなる" {
        $env = newIngestedRoot "強制終了後"
        $root = $env.Root
        $recordPath = "$root\work\store\pending_remove.tsv"
        # 一覧から行を消して、本文インデックスから外す前に強制終了した状態にする
        Remove-Item -LiteralPath "$($env.Dir)\議事録.docx"
        $statusPath = [Workspace]::new("$root\work").StatusFile
        writeListFile $statusPath @(Get-Content -LiteralPath $statusPath -Encoding UTF8 | Where-Object { $_ -notlike "*議事録.docx*" })
        writeListFile $recordPath @("$($env.Name)`t議事録.docx", "消えたインデックス`t別.docx")

        runIndexer $root | Should -Be 0

        (getAllContentIndexText $root $env.Name) | Should -Not -BeLike "*議事録.docx*"
        Test-Path -LiteralPath $recordPath | Should -Be $false
    }

    It "消したファイルが次の回までに同じ名前で戻って取り込み直されたとき、記録から戻した名前と重なっても新しい本文が残る" {
        $env = newIngestedRoot "戻る"
        $root = $env.Root
        $recordPath = "$root\work\store\pending_remove.tsv"
        $file = getContentIndexFile $root $env.Name
        $good = [System.IO.File]::ReadAllBytes($file)
        $docx = "$($env.Dir)\議事録.docx"
        Remove-Item -LiteralPath $docx
        [System.IO.File]::WriteAllText($file, "壊れた中身`n")
        runIndexer $root | Should -Be 0
        Test-Path -LiteralPath $recordPath | Should -Be $true

        [System.IO.File]::WriteAllBytes($file, $good)
        Copy-Item -LiteralPath $docxSource -Destination $docx

        runIndexer $root | Should -Be 0

        (getAllContentIndexText $root $env.Name) | Should -BeLike "*議事録.docx*"
        (readTestStatus $root).Rows["$($env.Name)\議事録.docx"].状態 | Should -Be ${stateDone}
        Test-Path -LiteralPath $recordPath | Should -Be $false
    }
}
