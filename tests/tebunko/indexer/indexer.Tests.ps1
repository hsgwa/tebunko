# インデクサの起動口（tebunko\indexer.ps1）と本体（indexer\indexer_run.ps1 の invokeIndexer）のテスト。
# indexer.ps1 は indexer_lib.ps1（lib.ps1 を含む）を読み込み、置き場所（リポジトリ直下の setting.config・work\）を決める。
# リポジトリの設定・インデックスを書き換えないよう、data_dir.ps1 で ${dataDir} を決める直前に止めて、
# ツールのフォルダ（${rootDir}）をテスト用のフォルダ（TestDrive）に差し替えてから続けさせる（Set-PSBreakpoint の -Action）。
# 同じやり方で、取り込みの途中に中止・画面の返事・元のファイルの削除を起こす。
# 画面とのやり取りは受け渡しの口（newIndexerChannel）で行う。途中に割り込むテストは、取り込みのスレッドを使わない
# （Workers = 0。ブレークポイントはテストと同じスレッドでしか止まらないため）。取り込みのスレッドを使う場合は別に確かめる。
# Excel は COM が要るため使わない。Word・PowerPoint の新形式（.docx・.pptx）はファイルを直接読むため、そのまま取り込む。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    $indexerPath = "${scriptsDir}\tebunko\indexer.ps1"
    $runPath     = "${scriptsDir}\tebunko\indexer\indexer_run.ps1"
    $dataDirPath = "${scriptsDir}\shared\core\data_dir.ps1"
    $planPath    = "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"
    $docxSource  = "${testDataDir}\office\Word\形式\大文字拡張子.DOCX"
    $pptxSource  = "${testDataDir}\office\PowerPoint\基本.pptx"

    function findLine {
        # ファイルの中で pattern に一致する最初の行の番号を返す（テストが行番号を直接書かないようにする）
        param ([string]$path, [string]$pattern)
        $lines = [System.IO.File]::ReadAllLines($path)
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match $pattern) { return $i + 1 }
        }
        throw "${path} に ${pattern} がありません"
    }

    $script:rootCount = 0
    function newRoot {
        # テストごとに別のツールの置き場所（setting.config・work\ を置くフォルダ）を作る
        $script:rootCount++
        $root = Join-Path $TestDrive "tool$($script:rootCount)"
        [System.IO.Directory]::CreateDirectory("$root\work") | Out-Null
        return $root
    }

    function writeTestSettings {
        # テスト用の置き場所に setting.config を書く。folders は @{ name; path; enabled } の配列
        param ([string]$root, [object[]]$folders)
        $settings = newSettings
        $settings.targetFolders = @($folders)
        # 既定のワークスペース（%USERPROFILE%\Documents\tebunko_ws）には利用者のインデックスがあるため、テスト用の work を指す
        $settings.workspaceFolder = "$root\work"
        writeSettings $settings "$root\setting.config"
    }

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

    # 最後に動かしたインデックス作成の受け渡しの口（進み具合・エラーを確かめる）
    $script:lastChannel = $null

    function runIndexer {
        # テスト用の置き場所（root）で indexer.ps1 を動かし、終了コードを返す。
        #   options: @{ RetryFailed; ConfirmTargets; Workers（既定 0 = 取り込みのスレッドを使わない） }
        #   breaks : 途中で動かす処理 @{ Script; Pattern; Action }（Pattern に一致する行に来るたびに Action を動かす）
        param (
            [string]$root,
            [hashtable]$options = @{},
            [object[]]$breaks = @()
        )

        $workers = if ($options.ContainsKey("Workers")) { $options.Workers } else { 0 }
        $channel = newIndexerChannel ([bool]$options.RetryFailed) ([bool]$options.ConfirmTargets) $workers
        $script:lastChannel = $channel
        $global:indexerTestRoot = $root
        $points = New-Object System.Collections.Generic.List[object]
        try {
            # ${dataDir} を決める行で、その前に ${rootDir} を差し替える（テスト用のフォルダには書き込めるため、setting.config・work もそこになる）。
            # Action は止まった場所の子のスコープで動く
            $points.Add((Set-PSBreakpoint -Script $dataDirPath -Line (findLine $dataDirPath '^\$\{dataDir\}\s*=') -Action {
                Set-Variable -Name rootDir -Value $global:indexerTestRoot -Scope 1
            }))
            foreach ($break in $breaks) {
                $points.Add((Set-PSBreakpoint -Script $break.Script -Line (findLine $break.Script $break.Pattern) -Action $break.Action))
            }
            & $indexerPath -Channel $channel *> $null
            return $channel.ExitCode
        } finally {
            foreach ($point in $points) { Remove-PSBreakpoint -Breakpoint $point }
            Remove-Variable -Name indexerTestRoot -Scope Global -ErrorAction SilentlyContinue
        }
    }

    function readTestStatus {
        param ([string]$root)
        return (readStatusFile "$root\work\取り込み一覧.tsv")
    }

    function readTestSystemState {
        param ([string]$root)
        return (readSystemIndexState "$root\work\システムインデックスの状態.tsv")
    }

    function readTestError {
        return $script:lastChannel.Error
    }

    function readTestProgress {
        return (readIndexingProgress $script:lastChannel)
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
        $point = Set-PSBreakpoint -Script $dataDirPath -Line (findLine $dataDirPath '^\$\{dataDir\}\s*=') -Action {
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
        [System.IO.File]::ReadAllText("$root\work\インデックス作成ログ.txt") | Should -Match "クロール対象フォルダがありません"
    }

    It "画面なしで壊れた setting.config のまま起動すると、退避して既定の設定で動き、ログに知らせを残して 1 で終わる" {
        $root = newRoot
        [System.IO.File]::WriteAllText("$root\setting.config", '{ "targetFolders": [', ${utf8Bom})
        # 退避すると既定のワークスペースになるため、利用者の本物の tebunko_ws に書かないよう、既定の場所（プロファイルの位置）もテスト用に差し替える
        $global:indexerTestRoot = $root
        $points = @(
            (Set-PSBreakpoint -Script $dataDirPath -Line (findLine $dataDirPath '^\$\{dataDir\}\s*=') -Action {
                Set-Variable -Name rootDir -Value $global:indexerTestRoot -Scope 1
            })
            (Set-PSBreakpoint -Script "${scriptsDir}\tebunko\core\settings.ps1" -Line (findLine "${scriptsDir}\tebunko\core\settings.ps1" 'return Join-Path \$profileDir') -Action {
                Set-Variable -Name profileDir -Value "$($global:indexerTestRoot)\profile" -Scope 1
            })
        )
        try {
            $global:LASTEXITCODE = 0
            & $indexerPath *> $null
            $LASTEXITCODE | Should -Be 1
        } finally {
            foreach ($point in $points) { Remove-PSBreakpoint -Breakpoint $point }
            Remove-Variable -Name indexerTestRoot -Scope Global -ErrorAction SilentlyContinue
        }
        $broken = @(Get-ChildItem -LiteralPath $root -Filter "setting.config.broken-*")
        $broken.Count | Should -Be 1
        [System.IO.File]::ReadAllText($broken[0].FullName, [System.Text.Encoding]::UTF8) | Should -Be '{ "targetFolders": ['
        $log = [System.IO.File]::ReadAllText("$root\profile\Documents\tebunko_ws\インデックス作成ログ.txt")
        $log | Should -Match "設定ファイルが壊れていたため、既定の設定で起動しました"
        $log | Should -Match ([regex]::Escape($broken[0].FullName))
        $log | Should -Match "クロール対象フォルダがありません"
    }

    It "同じ置き場所でほかのインデックス作成が動いていれば、そのインデックス作成のログに触らずに 1 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $TestDrive; enabled = $true })
        # 実行中のインデックス作成が画面とやり取りしているファイル
        $files = @("インデックス作成ログ.txt")
        foreach ($name in $files) {
            [System.IO.File]::WriteAllText("$root\work\$name", "実行中")
        }
        $other = newAppMutex "indexer" "$root\work"
        try {
            runIndexer $root | Should -Be 1
        } finally {
            $other.Mutex.Dispose()
        }
        foreach ($name in $files) {
            [System.IO.File]::ReadAllText("$root\work\$name") | Should -Be "実行中"
        }
        readTestError | Should -Match "ほかのインデックス作成が実行中です"
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
        # 取り込んだ TSV はフォルダの集約ファイルに入れ、元のファイルごとのフォルダは残さない
        [System.IO.File]::Exists("$root\work\content_index\営業\content_index.docx.001.tsv") | Should -Be $true
        [System.IO.Directory]::Exists("$root\work\content_index\営業\議事録.docx") | Should -Be $false
        Test-Path -LiteralPath "$root\work\content_index\営業\元のフォルダ.txt" | Should -Be $true
        # フォルダごとのシステムインデックスを作り、インデックスを対応済みにする
        [System.IO.File]::Exists("$root\work\system_index\営業\${systemIndexFileName}") | Should -Be $true
        (readTestSystemState $root).Covered.Contains("営業") | Should -Be $true
        # 名前の無かったフォルダには名前を割り当てて保存する
        @(getTargetFolders "$root\setting.config" | ForEach-Object { $_.Name }) -join "," | Should -Be "営業,経理,無くなったフォルダ"
        # 画面が終わり方を読めるよう、最後の進み具合を残す
        $progress = readTestProgress
        $progress.Phase | Should -Be ${indexingPhaseFinish}
        $progress.Failed | Should -Be 1
        readTestError | Should -BeNullOrEmpty
        Test-Path -LiteralPath "$root\work\取り込み中.txt" | Should -Be $false
    }

    It "2 回目は更新の無いファイルを取り込まず、前回失敗したファイルもスキップする" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        runIndexer $root | Should -Be 0
        $first = (readTestStatus $root).Rows["営業\壊れた.pptx"].取り込み日時

        runIndexer $root | Should -Be 0
        $progress = readTestProgress
        $progress.Detail | Should -Be "取り込みが必要なファイルはありませんでした"
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
        [System.IO.File]::Exists("$root\work\content_index\一時\content_index.docx.001.tsv") | Should -Be $true
    }

    It "前回取り込み中に強制終了したファイルは最後に回して取り込む" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        writeListFile "$root\work\取り込み中.txt" @("1`t営業\議事録.docx")

        runIndexer $root | Should -Be 0
        (readTestStatus $root).Rows["営業\議事録.docx"].状態 | Should -Be ${stateDone}
        Test-Path -LiteralPath "$root\work\取り込み中.txt" | Should -Be $false
    }

    It "続けて強制終了したファイルは取り込まずに失敗とする" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        writeListFile "$root\work\取り込み中.txt" @("2`t営業\議事録.docx")

        runIndexer $root | Should -Be 0
        $row = (readTestStatus $root).Rows["営業\議事録.docx"]
        $row.状態 | Should -Be ${stateFailed}
        $row.エラー | Should -Match "強制終了"
        Test-Path -LiteralPath "$root\work\取り込み中.txt" | Should -Be $false
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
            # PowerPointが要るファイルは「未取り込み」のまま残す（失敗にはしない）
            $status.Rows["後回し\旧形式.ppt"].状態 | Should -Be ${stateNew}
            # Office を使わないファイルはそのまま取り込む
            $status.Rows["後回し\議事録.docx"].状態 | Should -Be ${stateDone}

            $progress = readTestProgress
            $progress.Processed | Should -Be 1
            $progress.Failed | Should -Be 0
            $progress.Remaining | Should -Be 1
            readTestError | Should -BeNullOrEmpty
            $script:lastChannel.Postponed | Should -Be 1
            $script:lastChannel.Notice | Should -Match "PowerPoint が起動していたため、1 件を取り込まずに残しました"
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
        (readTestStatus $root).Rows["後回し\旧形式.ppt"].状態 | Should -Be ${stateDone}
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
            $status.Rows["後回し2\旧形式.ppt"].状態 | Should -Be ${stateNew}
            $status.Rows["後回し2\議事録.docx"].状態 | Should -Be ${stateDone}
            $script:lastChannel.Postponed | Should -Be 1
            $script:lastChannel.Notice | Should -Match "1 件を取り込まずに残しました"
            # stopAllApps は後片付け（finally）で 1 回だけ呼ばれる。後回しを取り込んだ件数に数えると
            # 100 件ごとの起動し直しの判定が早まって途中でも呼ばれるが、ここでは増えない
            # （RestartInterval を差し替えられる runIngestWorker 単体の It で、数えないことを詳しく確かめている）
            Should -Invoke stopAllApps -Times 1 -Exactly -Scope It

            # 利用者が PowerPoint を閉じたことにすると、後回しにしたファイルを取り込む
            $global:testUserPptOpen = $false
            runIndexer $root | Should -Be 0
            (readTestStatus $root).Rows["後回し2\旧形式.ppt"].状態 | Should -Be ${stateDone}
        } finally {
            Remove-Variable -Name testOfficeAppInUseMessage, testUserPptOpen -Scope Global -ErrorAction SilentlyContinue
        }
    }
}

Describe "indexer.ps1（画面の確認・中止）" -Tag Io {
    BeforeAll {
        $source = newSourceFolder "総務"
        # 確認待ち（waitForIndexingApproval）の中で、画面の返事を置く
        $approvalLine = @{ Script = $planPath; Pattern = '^\s+if \(!\$channel\.Answered\.WaitOne' }
    }

    It "確認を待つとき（ConfirmTargets）、取りやめたら、取り込まずに 2 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        $cancel = $approvalLine.Clone()
        $cancel.Action = { answerIndexingPlan $channel $null }

        runIndexer $root @{ ConfirmTargets = $true } @($cancel) | Should -Be 2
        # 1 件も取り込んでいないため、取り込み対象にした行は記録しない
        (readTestStatus $root).Rows.Count | Should -Be 0
        Test-Path -LiteralPath "$root\work\content_index\総務\議事録.docx" | Should -Be $false
        (readTestProgress).Detail | Should -Be "インデックス作成を取りやめました"
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

    It "取り込みの途中で中止を求められたら、残りを未取り込みのまま 2 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        # 1 件目を記録した直後に、画面が中止要求を作る
        $stop = @{ Script = $runPath; Pattern = '^\s+addStatusRow \$row'; Action = { $channel.Stop = $true } }

        runIndexer $root @{} @($stop) | Should -Be 2
        $rows = @((readTestStatus $root).Rows.Values)
        @($rows | Where-Object { $_.状態 -eq ${stateNew} }).Count | Should -Be 2
        $progress = readTestProgress
        $progress.Processed | Should -Be 1
        $progress.Remaining | Should -Be 2
        # 取り込んだ TSV は元のファイルごとのフォルダに残さず（集約ファイルに入れる）、途中なので対応済みにはしない
        (findIndexFoldersWithBooks "$root\work\content_index").Count | Should -Be 0
        (readTestSystemState $root).Covered.Count | Should -Be 0
    }
}

Describe "indexer.ps1（制限時間）" -Tag Io {
    It "制限時間を過ぎて失敗したファイルは、制限時間で中止したことをエラーに書く" {
        $source = newSourceFolder "監査"
        $root = newRoot
        writeTestSettings $root @(@{ name = "監査"; path = $source; enabled = $true })
        # 取り込みの直前に、見張り（startWatchdog）が制限時間を過ぎたと判断した状態にする
        $timeout = @{ Script = $runPath; Pattern = '^\s+\$result\.TsvCount = ingestFile'; Action = { $watchdog.TimedOut = $true } }

        runIndexer $root @{} @($timeout) | Should -Be 0
        $status = readTestStatus $root
        $status.Rows["監査\壊れた.pptx"].状態 | Should -Be ${stateFailed}
        $status.Rows["監査\壊れた.pptx"].エラー | Should -Match "分以内に取り込みが終わらなかった"
        # 取り込めたファイルは、制限時間の印が立っていても成功のまま
        $status.Rows["監査\議事録.docx"].状態 | Should -Be ${stateDone}
    }
}

Describe "indexer.ps1（取り込み中に元のファイルが無くなる）" -Tag Io {
    # 取り込みを始める直前（1 件目）に、元のファイル・フォルダを消す
    BeforeAll {
        $beforeIngest = '^\s+\$row = \$targets\[\$i\]'
    }

    It "元のファイルが無くなっていたら、取り込まずに一覧から除く" {
        $source = newSourceFolder "人事"
        $root = newRoot
        writeTestSettings $root @(@{ name = "人事"; path = $source; enabled = $true })
        $global:indexerTestVictim = "$source\壊れた.pptx"
        $remove = @{ Script = $runPath; Pattern = $beforeIngest; Action = {
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

    It "前回の作成で残った TSV（元のファイルごとのフォルダ）は、次の作成の始めに集約ファイルへ入れる" {
        $source = newSourceFolder "総務3"
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務3"; path = $source; enabled = $true })
        runIndexer $root | Should -Be 0
        writeListFile "$root\work\content_index\総務3\残った.xlsx\S.tsv" @("残っていた中身")

        runIndexer $root | Should -Be 0
        [System.IO.Directory]::Exists("$root\work\content_index\総務3\残った.xlsx") | Should -Be $false
        $packs = getPackFiles "$root\work\content_index" "総務3" $false
        (searchPackIndex "残っていた中身" $packs $true).Hits.Count | Should -Be 1
    }

    It "クロール対象フォルダごと見えなくなったら、残りを未取り込みのまま 1 で終わる" {
        $source = newSourceFolder "法務"
        $root = newRoot
        writeTestSettings $root @(@{ name = "法務"; path = $source; enabled = $true })
        $global:indexerTestVictim = $source
        $remove = @{ Script = $runPath; Pattern = $beforeIngest; Action = {
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
        newTsv "$work\システムインデックスの状態.tsv" @("対応済み`t$name`t")
        $docxFile = Get-Item -LiteralPath "$sourcePath\議事録.docx"
        $updated = formatFileTime $docxFile.LastWriteTime
        $size = [string]$docxFile.Length
        $extractVersion = [string]$extractVersions[".docx"]
        writeStatusFile @([pscustomobject]@{ Path = $sourcePath; Name = $name }) @(
            (newStatusRow "$name\議事録.docx" $updated $size ${stateDone} "1" "2026/09/01 10:00:00" "" $extractVersion)
        ) "$work\取り込み一覧.tsv"
    }

    It "前の版のインデックスがあれば知らせて片付け、取り込み直す（前の版の index には触らない）" {
        $source = newSourceFolder "旧版"
        $root = newRoot
        writeTestSettings $root @(@{ name = "旧版"; path = $source; enabled = $true })
        newLegacyWorkspace $root "旧版" $source
        $before = (readTestStatus $root).Rows["旧版\議事録.docx"].取り込み日時

        runIndexer $root | Should -Be 0

        $log = [System.IO.File]::ReadAllText("$root\work\インデックス作成ログ.txt")
        $log | Should -Match "前の版のインデックス"
        # 「済」の行が取り込み直された理由が、本文インデックスが無いこと（lost）であることをログで確かめる
        $log | Should -Match "インデックスが無い・壊れている 1 件"
        [System.IO.File]::Exists("$root\work\content_index\旧版\content_index.docx.001.tsv") | Should -Be $true
        [System.IO.File]::Exists("$root\work\system_index\旧版\system_index.txt") | Should -Be $true
        @([System.IO.Directory]::GetFiles("$root\work\system_index", "システムインデックス*.txt", "AllDirectories")).Count | Should -Be 0
        $stateText = [System.IO.File]::ReadAllText("$root\work\システムインデックスの状態.tsv")
        $stateText | Should -Not -Match "システムインデックス"
        (readTestSystemState $root).Covered.Contains("旧版") | Should -Be $true
        # 取り込み一覧の「済」の行（前の版の記録）も、本文インデックスが無いため取り込み直す
        $row = (readTestStatus $root).Rows["旧版\議事録.docx"]
        $row.状態 | Should -Be ${stateDone}
        $row.取り込み日時 | Should -Not -Be $before
        # 前の版の index\ は読まず・消さない
        Test-Path -LiteralPath "$root\work\index\旧版\元のフォルダ.txt" | Should -Be $true
        Test-Path -LiteralPath "$root\work\index\旧版\議事録.docx\S.tsv" | Should -Be $true
    }

    It "前の版のシステムインデックスを片付けられなければ、content_index を作らずに 1 で終わる。閉じてもう一度で取り込み直す" {
        $source = newSourceFolder "旧版2"
        $root = newRoot
        writeTestSettings $root @(@{ name = "旧版2"; path = $source; enabled = $true })
        newLegacyWorkspace $root "旧版2" $source
        $legacyTxt = "$root\work\system_index\旧版2\システムインデックス.txt"
        $stream = [System.IO.File]::Open($legacyTxt, "Open", "Read", "None")
        try {
            runIndexer $root | Should -Be 1
        } finally {
            $stream.Dispose()
        }
        readTestError | Should -Match "片付けられませんでした"
        Test-Path -LiteralPath "$root\work\content_index" | Should -Be $false
        Test-Path -LiteralPath $legacyTxt | Should -Be $true

        runIndexer $root | Should -Be 0
        [System.IO.File]::Exists("$root\work\content_index\旧版2\content_index.docx.001.tsv") | Should -Be $true
        Test-Path -LiteralPath $legacyTxt | Should -Be $false
    }
}

Describe "indexer.ps1（取り込みのスレッド）" -Tag Io {
    BeforeAll {
        $source = newSourceFolder "並列"
        [System.IO.Directory]::CreateDirectory("$source\資料\下") | Out-Null
        Copy-Item -LiteralPath $docxSource -Destination "$source\資料\下\報告.docx"
        Copy-Item -LiteralPath $pptxSource -Destination "$source\資料\報告2.pptx"

        function script:readPackLines([string]$root) {
            # 集約ファイルの中身（ファイル名・場所・行）を、パスの順に並べて返す
            $lines = New-Object System.Collections.Generic.List[string]
            foreach ($file in @([System.IO.Directory]::GetFiles("$root\work\content_index", ${packFilePattern}, "AllDirectories") | Sort-Object)) {
                $lines.Add($file.Substring("$root\work\content_index".Length))
                $lines.AddRange([string[]]@([System.IO.File]::ReadAllLines($file)))
            }
            return , $lines.ToArray()
        }
    }

    It "取り込みのスレッドで取り込んでも、取り込み一覧・集約ファイルはスレッドを使わないときと同じになる" {
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
        (readPackLines $parallel) -join "`n" | Should -BeExactly ((readPackLines $single) -join "`n")
        # 取り込んだ TSV（元のファイルごとのフォルダ）・取り込み中の記録・一時フォルダは残さない
        (findIndexFoldersWithBooks "$parallel\work\content_index").Count | Should -Be 0
        Test-Path -LiteralPath "$parallel\work\取り込み中.txt" | Should -Be $false
        @(Get-ChildItem -LiteralPath "$parallel\work\取り込み出力" -Force -ErrorAction SilentlyContinue).Count | Should -Be 0
        $progress = readTestProgress
        $progress.Processed | Should -Be 5
        $progress.Failed | Should -Be 1
        (Get-Content -LiteralPath "$parallel\work\インデックス作成ログ.txt" -Raw) | Should -Match "3 個のスレッドで並べて取り込みます"
    }

    It "取り込みのスレッドが始められなければ、続けられないエラーで 1 を返す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "並列"; path = $source; enabled = $true })
        # 取り込みのスレッドが読み込む部品の場所を、無い場所にする
        $broken = @{ Script = $runPath; Pattern = '^\s+\$pool = newIngestPool'; Action = {
                Set-Variable -Name indexerLibPath -Value (Join-Path $TestDrive "無い.ps1") -Scope 1
            }
        }
        runIndexer $root @{ Workers = 2 } @($broken) | Should -Be 1
        readTestError | Should -Match "取り込みのスレッドが止まりました"
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
    # 取り込みのスレッドで動くスクリプトを、このスレッドで直接動かして確かめる（スレッドの中の動きはブレークポイントで止められないため）
    BeforeAll {
        . $runPath
        . "${scriptsDir}\shared\office\office_app.ps1"
    }

    It "取り込み待ちの列のファイルを取り込んで結果の列に入れ、列が閉じられたら Office を片づけて終わる" {
        $root = Join-Path $TestDrive "worker_direct"
        foreach ($dir in "content_index", "tmp", "publish") {
            [System.IO.Directory]::CreateDirectory("$root\$dir") | Out-Null
        }
        $tasks = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $results = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $tasks.Add(@{ RelPath = "営業\議事録.docx"; SourcePath = $docxSource })
        $tasks.Add(@{ RelPath = "営業\無い.docx"; SourcePath = (Join-Path $TestDrive "無い.docx") })
        $tasks.CompleteAdding()
        $settings = @{
            Lib = "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"
            WorkDir = $root; TmpDir = "$root\tmp"; PublishDir = "$root\publish"
            FileTimeoutMinutes = 10; RestartInterval = 1
            OfficePids = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,string]'
            Lane = ${laneReader}
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
        @([System.IO.Directory]::GetFiles("$root\content_index\営業\議事録.docx", "*.tsv")).Count | Should -Be $done.TsvCount
        # 一時フォルダはスレッドごとに分ける
        [System.IO.Directory]::Exists("$root\tmp\w7") | Should -Be $true
        $failed = $results.Take()
        $failed.Ok | Should -Be $false
        $failed.Message | Should -Not -BeNullOrEmpty
    }
}

Describe "runIngestWorker・invokeIngestTask（レーン）" -Tag Io {
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
}

Describe "invokeIngestTask（Office が要る）" -Tag Io {
    # invokeIngestTask そのものを確かめるため、invokeIngestTask を Mock する上の Describe と分ける
    BeforeAll {
        . "${scriptsDir}\shared\office\office_app.ps1"
        . "${scriptsDir}\tebunko\indexer\extract_office.ps1"
        . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"
        . $runPath
    }

    AfterEach { $script:officeUnavailable = $false }

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
}

Describe "getIngestLaneCapacity" -Tag Unit {
    BeforeAll {
        . $runPath
    }
    It "Office のレーンは取り込み中と次の 1 件、読み取りのレーンはスレッドの数の 2 倍まで渡す" {
        getIngestLaneCapacity ${laneExcel} 3 | Should -Be 2
        getIngestLaneCapacity ${lanePowerPoint} 3 | Should -Be 2
        getIngestLaneCapacity ${laneReader} 3 | Should -Be 6
        getIngestLaneCapacity ${laneReader} 0 | Should -Be 2
    }
}
