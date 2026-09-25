# インデクサの起動口（tebunko_grep\indexer.ps1）のテスト。
# indexer.ps1 は lib.ps1 を読み込み、置き場所（リポジトリ直下の setting.config・work\）を決める。
# リポジトリの設定・インデックスを書き換えないよう、data_dir.ps1 で ${dataDir} を決める直前に止めて、
# ツールのフォルダ（${rootDir}）をテスト用のフォルダ（TestDrive）に差し替えてから続けさせる（Set-PSBreakpoint の -Action）。
# 同じやり方で、取り込みの途中に中止要求・画面の返事・元のファイルの削除を起こす。
# Excel は COM が要るため使わない。Word・PowerPoint の新形式（.docx・.pptx）はファイルを直接読むため、そのまま取り込む。
. "$PSScriptRoot\..\..\helpers\load.ps1"

$indexerPath = "${scriptsDir}\tebunko_grep\indexer.ps1"
$indexerDir  = "${scriptsDir}\tebunko_grep"
$dataDirPath = "${scriptsDir}\shared\core\data_dir.ps1"
$planPath    = "${scriptsDir}\tebunko_grep\indexer\indexer_plan.ps1"
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
    # 既定のワークスペース（%USERPROFILE%\Documents\tebunko）は開発の PC ではほかのファイルがあり使えないため、テスト用の work を指す
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

function invokeIndexer {
    # テスト用の置き場所（root）で indexer.ps1 を動かし、終了コードを返す。
    #   breaks: 途中で動かす処理 @{ Script; Pattern; Action }（Pattern に一致する行に来るたびに Action を動かす）
    param (
        [string]$root,
        [hashtable]$arguments = @{},
        [object[]]$breaks = @()
    )

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
        $global:LASTEXITCODE = 0
        & $indexerPath @arguments *> $null
        return $LASTEXITCODE
    } finally {
        foreach ($point in $points) { Remove-PSBreakpoint -Breakpoint $point }
        Remove-Variable -Name indexerTestRoot -Scope Global -ErrorAction SilentlyContinue
        # インデクサは置き場所ごとのミューテックスを持ったまま終わる（本来はプロセスの終了で解放される）。
        # 同じ置き場所で続けて動かせるよう、参照の無くなったミューテックスを GC で閉じる
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function readTestStatus {
    param ([string]$root)
    return (readStatusFile "$root\work\取り込み一覧.tsv")
}

function readTestError {
    param ([string]$root)
    $path = "$root\work\インデックス作成エラー.txt"
    if (!(Test-Path -LiteralPath $path)) { return $null }
    return [System.IO.File]::ReadAllText($path)
}

Describe "indexer.ps1（続けられないエラー）" -Tag Io {
    It "クロール対象フォルダが無ければ、エラーを書いて 1 で終わる" {
        $root = newRoot
        writeTestSettings $root @()
        invokeIndexer $root | Should Be 1
        readTestError $root | Should Match "クロール対象フォルダがありません"
    }

    It "チェックの付いたフォルダが無ければ、エラーを書いて 1 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = (Join-Path $TestDrive "無し"); enabled = $false })
        invokeIndexer $root | Should Be 1
        readTestError $root | Should Match "チェックの付いたクロール対象フォルダがありません"
    }

    It "同じ置き場所でほかのインデックス作成が動いていれば、そのインデックス作成の受け渡しのファイル・ログに触らずに 1 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $TestDrive; enabled = $true })
        # 実行中のインデックス作成が画面とやり取りしているファイル
        $files = @("インデックス作成中止要求", "取り込み予定.tsv", "インデックス作成開始要求", "インデックス作成進捗.txt", "インデックス作成ログ.txt")
        foreach ($name in $files) {
            [System.IO.File]::WriteAllText("$root\work\$name", "実行中")
        }
        $other = newAppMutex "indexer" "$root\work"
        try {
            invokeIndexer $root | Should Be 1
        } finally {
            $other.Mutex.Dispose()
        }
        foreach ($name in $files) {
            [System.IO.File]::ReadAllText("$root\work\$name") | Should Be "実行中"
        }
        readTestError $root | Should Match "ほかのインデックス作成が実行中です"
    }

    It "前回の中止要求・エラー・取り込み予定・開始要求は、始める前に消す" {
        $root = newRoot
        writeTestSettings $root @()
        foreach ($name in @("インデックス作成中止要求", "取り込み予定.tsv", "インデックス作成開始要求", "変換対象一覧.txt")) {
            [System.IO.File]::WriteAllText("$root\work\$name", "前回")
        }
        invokeIndexer $root | Should Be 1
        Test-Path -LiteralPath "$root\work\インデックス作成中止要求" | Should Be $false
        Test-Path -LiteralPath "$root\work\取り込み予定.tsv" | Should Be $false
        Test-Path -LiteralPath "$root\work\インデックス作成開始要求" | Should Be $false
    }
}

Describe "indexer.ps1（取り込み）" -Tag Io {
    $source = newSourceFolder "営業"
    $unchecked = newSourceFolder "経理"
    $missing = Join-Path $TestDrive "無くなったフォルダ"

    It "チェックの付いたフォルダを取り込み、名前を設定に保存し、失敗したファイルを一覧に残す" {
        $root = newRoot
        writeTestSettings $root @(
            @{ name = ""; path = $source; enabled = $true },
            @{ name = ""; path = $unchecked; enabled = $false },
            @{ name = ""; path = $missing; enabled = $true })
        # 以前の版の途中状態ファイルは消す
        [System.IO.File]::WriteAllText("$root\work\変換失敗一覧.txt", "前回")

        invokeIndexer $root | Should Be 0

        $status = readTestStatus $root
        $status.Rows.Count | Should Be 3
        $status.Rows["営業\議事録.docx"].状態 | Should Be ${stateDone}
        $status.Rows["営業\資料\提案.pptx"].状態 | Should Be ${stateDone}
        $status.Rows["営業\壊れた.pptx"].状態 | Should Be ${stateFailed}
        $status.Rows["営業\壊れた.pptx"].エラー | Should Match "PowerPoint"
        # 取り込んだ TSV はフォルダの集約ファイルに入れ、元のファイルごとのフォルダは残さない
        [System.IO.File]::Exists("$root\work\index\営業\content.docx.001.tsv") | Should Be $true
        [System.IO.Directory]::Exists("$root\work\index\営業\議事録.docx") | Should Be $false
        Test-Path -LiteralPath "$root\work\index\営業\元のフォルダ.txt" | Should Be $true
        Test-Path -LiteralPath "$root\work\変換失敗一覧.txt" | Should Be $false
        # 名前の無かったフォルダには名前を割り当てて保存する
        @(getTargetFolders "$root\setting.config" | ForEach-Object { $_.Name }) -join "," | Should Be "営業,経理,無くなったフォルダ"
        # 画面が終わり方を読めるよう、最後の進み具合を残す
        $progress = readIndexingProgress "$root\work\インデックス作成進捗.txt"
        $progress.Phase | Should Be ${indexingPhaseFinish}
        $progress.Failed | Should Be 1
        readTestError $root | Should BeNullOrEmpty
        Test-Path -LiteralPath "$root\work\取り込み中.txt" | Should Be $false
    }

    It "設定 workspaceFolder があれば、そのフォルダ（無ければ作る）にインデックス・取り込み一覧を作り、ツールのフォルダの work には書かない" {
        $root = newRoot
        $work = Join-Path $TestDrive "別のドライブのつもり$($script:rootCount)\データ"
        $settings = newSettings
        $settings.targetFolders = @(@{ name = "営業"; path = $source; enabled = $true })
        $settings.workspaceFolder = $work
        writeSettings $settings "$root\setting.config"

        invokeIndexer $root | Should Be 0

        Test-Path -LiteralPath "$work\取り込み一覧.tsv" | Should Be $true
        [System.IO.File]::Exists("$work\index\営業\content.docx.001.tsv") | Should Be $true
        Test-Path -LiteralPath "$work\インデックス作成ログ.txt" | Should Be $true
        @(Get-ChildItem -LiteralPath "$root\work").Count | Should Be 0
    }

    It "2 回目は更新の無いファイルを取り込まず、前回失敗したファイルもスキップする" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        invokeIndexer $root | Should Be 0
        $first = (readTestStatus $root).Rows["営業\壊れた.pptx"].取り込み日時

        invokeIndexer $root | Should Be 0
        $progress = readIndexingProgress "$root\work\インデックス作成進捗.txt"
        $progress.Detail | Should Be "取り込みが必要なファイルはありませんでした"
        $status = readTestStatus $root
        $status.Rows.Count | Should Be 3
        $status.Rows["営業\壊れた.pptx"].状態 | Should Be ${stateFailed}
        $status.Rows["営業\壊れた.pptx"].取り込み日時 | Should Be $first
    }

    It "-RetryFailed なら、前回失敗して更新の無いファイルも取り込み直す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        invokeIndexer $root | Should Be 0

        invokeIndexer $root @{ RetryFailed = $true } | Should Be 0
        $progress = readIndexingProgress "$root\work\インデックス作成進捗.txt"
        $progress.Processed | Should Be 1
        $progress.Failed | Should Be 1
        (readTestStatus $root).Rows["営業\壊れた.pptx"].状態 | Should Be ${stateFailed}
    }

    It "チェックを外したフォルダは取り込まず、前回の結果を一覧に残す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        invokeIndexer $root | Should Be 0
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $false }, @{ name = "経理"; path = $unchecked; enabled = $true })

        invokeIndexer $root | Should Be 0
        $status = readTestStatus $root
        $status.Rows["営業\議事録.docx"].状態 | Should Be ${stateDone}
        $status.Rows["経理\議事録.docx"].状態 | Should Be ${stateDone}
    }

    It "フォルダが見つからないとき（ネットワークの切断など）は、前回の結果とインデックスを残す" {
        $gone = newSourceFolder "一時"
        $root = newRoot
        writeTestSettings $root @(@{ name = "一時"; path = $gone; enabled = $true })
        invokeIndexer $root | Should Be 0
        Remove-Item -LiteralPath $gone -Recurse -Force

        invokeIndexer $root | Should Be 0
        $status = readTestStatus $root
        $status.Rows.Count | Should Be 3
        $status.Rows["一時\議事録.docx"].状態 | Should Be ${stateDone}
        [System.IO.File]::Exists("$root\work\index\一時\content.docx.001.tsv") | Should Be $true
    }

    It "前回取り込み中に強制終了したファイルは最後に回して取り込む" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        writeListFile "$root\work\取り込み中.txt" @("1`t営業\議事録.docx")

        invokeIndexer $root | Should Be 0
        (readTestStatus $root).Rows["営業\議事録.docx"].状態 | Should Be ${stateDone}
        Test-Path -LiteralPath "$root\work\取り込み中.txt" | Should Be $false
    }

    It "続けて強制終了したファイルは取り込まずに失敗とする" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        writeListFile "$root\work\取り込み中.txt" @("2`t営業\議事録.docx")

        invokeIndexer $root | Should Be 0
        $row = (readTestStatus $root).Rows["営業\議事録.docx"]
        $row.状態 | Should Be ${stateFailed}
        $row.エラー | Should Match "強制終了"
        Test-Path -LiteralPath "$root\work\取り込み中.txt" | Should Be $false
    }

    It "強制終了の記録が取り込み対象に無いファイルなら、記録を消して続ける" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true })
        writeListFile "$root\work\取り込み中.txt" @("1`t営業\無いファイル.docx")

        invokeIndexer $root | Should Be 0
        (readTestStatus $root).Rows["営業\議事録.docx"].状態 | Should Be ${stateDone}
        Test-Path -LiteralPath "$root\work\取り込み中.txt" | Should Be $false
    }
}

Describe "indexer.ps1（画面の確認・中止）" -Tag Io {
    $source = newSourceFolder "総務"
    # 確認待ち（waitForIndexingApproval）の中で、画面の返事を置く
    $approvalLine = @{ Script = $planPath; Pattern = 'if \(Test-Path -LiteralPath \$\{stopRequestFile\}\)' }

    It "-ConfirmTargets で取りやめたら、取り込まずに 2 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        $cancel = $approvalLine.Clone()
        $cancel.Action = { [System.IO.File]::WriteAllText(${stopRequestFile}, "") }

        invokeIndexer $root @{ ConfirmTargets = $true } @($cancel) | Should Be 2
        # 1 件も取り込んでいないため、取り込み対象にした行は記録しない
        (readTestStatus $root).Rows.Count | Should Be 0
        Test-Path -LiteralPath "$root\work\index\総務\議事録.docx" | Should Be $false
        (readIndexingProgress "$root\work\インデックス作成進捗.txt").Detail | Should Be "インデックス作成を取りやめました"
        Test-Path -LiteralPath "$root\work\取り込み予定.tsv" | Should Be $false
    }

    It "-ConfirmTargets で取りやめたら、取り込み済みの行は前回のまま残す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        invokeIndexer $root | Should Be 0
        $before = (readTestStatus $root).Rows["総務\壊れた.pptx"].取り込み日時
        $cancel = $approvalLine.Clone()
        $cancel.Action = { [System.IO.File]::WriteAllText(${stopRequestFile}, "") }

        invokeIndexer $root @{ ConfirmTargets = $true; RetryFailed = $true } @($cancel) | Should Be 2
        $status = readTestStatus $root
        $status.Rows.Count | Should Be 3
        $status.Rows["総務\壊れた.pptx"].取り込み日時 | Should Be $before
    }

    It "-ConfirmTargets で取りやめたら、更新されたファイルの行も前回の記録のまま残す（次回も更新ありになる）" {
        $edited = newSourceFolder "企画"
        $root = newRoot
        writeTestSettings $root @(@{ name = "企画"; path = $edited; enabled = $true })
        invokeIndexer $root | Should Be 0
        $before = (readTestStatus $root).Rows["企画\議事録.docx"]
        (Get-Item -LiteralPath "$edited\議事録.docx").LastWriteTime = (Get-Date).AddDays(1)
        $cancel = $approvalLine.Clone()
        $cancel.Action = { [System.IO.File]::WriteAllText(${stopRequestFile}, "") }

        invokeIndexer $root @{ ConfirmTargets = $true } @($cancel) | Should Be 2
        $row = (readTestStatus $root).Rows["企画\議事録.docx"]
        $row.状態 | Should Be ${stateDone}
        $row.更新日時 | Should Be $before.更新日時
    }

    It "-ConfirmTargets で画面が開始を選んだら取り込む（失敗分の再取り込みも画面の返事に従う）" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        invokeIndexer $root | Should Be 0
        $approve = $approvalLine.Clone()
        $approve.Action = { writeIndexingStartRequest $true }

        invokeIndexer $root @{ ConfirmTargets = $true } @($approve) | Should Be 0
        $progress = readIndexingProgress "$root\work\インデックス作成進捗.txt"
        $progress.Processed | Should Be 1
        $progress.Failed | Should Be 1
        Test-Path -LiteralPath "$root\work\インデックス作成開始要求" | Should Be $false
    }

    It "取り込みの途中で中止を求められたら、残りを未取り込みのまま 2 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務"; path = $source; enabled = $true })
        # 1 件目を記録した直後に、画面が中止要求を作る
        $stop = @{ Script = $indexerPath; Pattern = '^\s+addStatusRow \$row'; Action = { [System.IO.File]::WriteAllText(${stopRequestFile}, "") } }

        invokeIndexer $root @{} @($stop) | Should Be 2
        $rows = @((readTestStatus $root).Rows.Values)
        @($rows | Where-Object { $_.状態 -eq ${stateNew} }).Count | Should Be 2
        $progress = readIndexingProgress "$root\work\インデックス作成進捗.txt"
        $progress.Processed | Should Be 1
        $progress.Remaining | Should Be 2
        Test-Path -LiteralPath "$root\work\インデックス作成中止要求" | Should Be $false
    }
}

Describe "indexer.ps1（制限時間）" -Tag Io {
    It "制限時間を過ぎて失敗したファイルは、制限時間で中止したことをエラーに書く" {
        $source = newSourceFolder "監査"
        $root = newRoot
        writeTestSettings $root @(@{ name = "監査"; path = $source; enabled = $true })
        # 取り込みの直前に、見張り（startWatchdog）が制限時間を過ぎたと判断した状態にする
        $timeout = @{ Script = $indexerPath; Pattern = '^\s+\$tsvCount = ingestFile'; Action = { $watchdog.TimedOut = $true } }

        invokeIndexer $root @{} @($timeout) | Should Be 0
        $status = readTestStatus $root
        $status.Rows["監査\壊れた.pptx"].状態 | Should Be ${stateFailed}
        $status.Rows["監査\壊れた.pptx"].エラー | Should Match "分以内に取り込みが終わらなかった"
        # 取り込めたファイルは、制限時間の印が立っていても成功のまま
        $status.Rows["監査\議事録.docx"].状態 | Should Be ${stateDone}
    }
}

Describe "indexer.ps1（取り込み中に元のファイルが無くなる）" -Tag Io {
    # 取り込みを始める直前（1 件目）に、元のファイル・フォルダを消す
    $beforeIngest = '^\s+\$row = \$targets\[\$i\]'

    It "元のファイルが無くなっていたら、取り込まずに一覧から除く" {
        $source = newSourceFolder "人事"
        $root = newRoot
        writeTestSettings $root @(@{ name = "人事"; path = $source; enabled = $true })
        $global:indexerTestVictim = "$source\壊れた.pptx"
        $remove = @{ Script = $indexerPath; Pattern = $beforeIngest; Action = {
            if (Test-Path -LiteralPath $global:indexerTestVictim) { Remove-Item -LiteralPath $global:indexerTestVictim -Force }
        } }
        try {
            invokeIndexer $root @{} @($remove) | Should Be 0
        } finally {
            Remove-Variable -Name indexerTestVictim -Scope Global
        }
        $status = readTestStatus $root
        $status.Rows.Count | Should Be 2
        $status.Rows.ContainsKey("人事\壊れた.pptx") | Should Be $false
    }

    It "元のファイルが無くなったら、次のインデックス作成で集約ファイルから外す" {
        $source = newSourceFolder "総務2"
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務2"; path = $source; enabled = $true })
        invokeIndexer $root | Should Be 0
        [System.IO.File]::Exists("$root\work\index\総務2\content.docx.001.tsv") | Should Be $true
        Remove-Item -LiteralPath "$source\議事録.docx" -Force

        invokeIndexer $root | Should Be 0
        # docx は議事録.docx だけだったため、集約ファイルごと無くなる
        [System.IO.File]::Exists("$root\work\index\総務2\content.docx.001.tsv") | Should Be $false
        (readTestStatus $root).Rows.ContainsKey("総務2\議事録.docx") | Should Be $false
    }

    It "前回の作成で残った TSV（元のファイルごとのフォルダ）は、次の作成の始めに集約ファイルへ入れる" {
        $source = newSourceFolder "総務3"
        $root = newRoot
        writeTestSettings $root @(@{ name = "総務3"; path = $source; enabled = $true })
        invokeIndexer $root | Should Be 0
        writeListFile "$root\work\index\総務3\残った.xlsx\S.tsv" @("残っていた中身")

        invokeIndexer $root | Should Be 0
        [System.IO.Directory]::Exists("$root\work\index\総務3\残った.xlsx") | Should Be $false
        $packs = getPackFiles "$root\work\index" "総務3" $false
        (searchPackIndex "残っていた中身" $packs $true).Hits.Count | Should Be 1
    }

    It "クロール対象フォルダごと見えなくなったら、残りを未取り込みのまま 1 で終わる" {
        $source = newSourceFolder "法務"
        $root = newRoot
        writeTestSettings $root @(@{ name = "法務"; path = $source; enabled = $true })
        $global:indexerTestVictim = $source
        $remove = @{ Script = $indexerPath; Pattern = $beforeIngest; Action = {
            if (Test-Path -LiteralPath $global:indexerTestVictim) { Remove-Item -LiteralPath $global:indexerTestVictim -Recurse -Force }
        } }
        try {
            invokeIndexer $root @{} @($remove) | Should Be 1
        } finally {
            Remove-Variable -Name indexerTestVictim -Scope Global
        }
        readTestError $root | Should Match "クロール対象フォルダが見つからなくなった"
        $rows = @((readTestStatus $root).Rows.Values)
        $rows.Count | Should Be 3
        @($rows | Where-Object { $_.状態 -eq ${stateNew} }).Count | Should Be 3
    }
}

Describe "indexer.ps1（システムインデックス）" -Tag Io {
    $source = newSourceFolder "広報"

    function script:readTestSystemState {
        param ([string]$root)
        return (readSystemIndexState "$root\work\システムインデックスの状態.tsv")
    }

    function script:readTestLog {
        param ([string]$root)
        return [System.IO.File]::ReadAllText("$root\work\インデックス作成ログ.txt")
    }

    It "取り込むと、フォルダごとのシステムインデックスを作り、インデックスを対応済みにする" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "広報"; path = $source; enabled = $true })

        invokeIndexer $root | Should Be 0

        $txt = "$root\work\system_index\広報\${systemIndexFileName}"
        [System.IO.File]::Exists($txt) | Should Be $true
        [System.IO.File]::Exists("$root\work\system_index\広報\資料\${systemIndexFileName}") | Should Be $true
        $state = readTestSystemState $root
        $state.Covered.Contains("広報") | Should Be $true
        # 反映待ちの日時は txt の更新日時（Windows Search に反映されたかの判定に使う）
        $state.Pending["広報\${systemIndexFileName}"] | Should Be ([System.IO.File]::GetLastWriteTimeUtc($txt).Ticks)
    }

    It "取り込むファイルが無くても、無くなったシステムインデックスは作り直す（この版に上げた直後など）" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "広報"; path = $source; enabled = $true })
        invokeIndexer $root | Should Be 0
        removeDirectoryRetry "$root\work\system_index"
        [System.IO.File]::Delete("$root\work\システムインデックスの状態.tsv")

        invokeIndexer $root | Should Be 0

        (readTestLog $root) | Should Match "取り込みが必要なファイルはありません"
        [System.IO.File]::Exists("$root\work\system_index\広報\${systemIndexFileName}") | Should Be $true
        (readTestSystemState $root).Covered.Contains("広報") | Should Be $true
    }

    It "取り込みの途中で中止しても、取り込んだ分は集約ファイルとシステムインデックスに入れ、対応済みにはしない" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "広報"; path = $source; enabled = $true })
        # 取り込めたファイル（TSV を入れ替えたファイル）を記録した直後に中止する
        $stop = @{ Script = $indexerPath; Pattern = '^\s+addStatusRow \$row'; Action = { if ($row.状態 -eq ${stateDone}) { [System.IO.File]::WriteAllText(${stopRequestFile}, "") } } }

        invokeIndexer $root @{} @($stop) | Should Be 2

        # 取り込んだ TSV は残さない（集約ファイルに入れる）
        (findIndexFoldersWithBooks "$root\work\index").Count | Should Be 0
        @([System.IO.Directory]::GetFiles("$root\work\index", ${packFilePattern}, "AllDirectories")).Count | Should Be 1
        @([System.IO.Directory]::GetFiles("$root\work\system_index", "*.txt", "AllDirectories")).Count | Should Be 1
        $state = readTestSystemState $root
        $state.Covered.Count | Should Be 0
        # 書いた txt は反映待ち（txt の更新日時）
        @($state.Pending.Values | Where-Object { $_ -eq 0 }).Count | Should Be 0
    }
    It "システムインデックスを作れなくても、インデックス作成は終わり、理由をログに書く" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "広報"; path = $source; enabled = $true })
        # system_index という名前のファイルがあると、フォルダを作れない
        [System.IO.File]::WriteAllText("$root\work\system_index", "")

        invokeIndexer $root | Should Be 0

        (readTestLog $root) | Should Match "システムインデックスを作れませんでした"
        (readTestStatus $root).Rows["広報\議事録.docx"].状態 | Should Be ${stateDone}
        (readTestSystemState $root).Covered.Count | Should Be 0

        # 取り込むファイルが無いときも同じ
        invokeIndexer $root | Should Be 0
        (readTestLog $root) | Should Match "システムインデックスを作れませんでした"
    }

    It "状態ファイルに書けなくても取り込みを続け、インデックス作成の終わりに作り直す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "広報"; path = $source; enabled = $true })
        # 1 件目の TSV を入れ替える直前に状態ファイルをほかから開き、記録した直後に閉じる
        $lock = @{ Script = $indexerPath; Pattern = '^\s+publishTsv \(getBookDir'; Action = {
                if (!$global:systemStateLock) {
                    [System.IO.Directory]::CreateDirectory(${workDir}) | Out-Null
                    $global:systemStateLock = [System.IO.FileStream]::new(${systemIndexStateFile}, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
                }
            }
        }
        $unlock = @{ Script = $indexerPath; Pattern = '^\s+addStatusRow \$row'; Action = {
                if ($global:systemStateLock) { $global:systemStateLock.Dispose() }
            }
        }
        try {
            invokeIndexer $root @{} @($lock, $unlock) | Should Be 0
        } finally {
            if ($global:systemStateLock) { $global:systemStateLock.Dispose() }
            Remove-Variable -Name systemStateLock -Scope Global -ErrorAction SilentlyContinue
        }

        (readTestLog $root) | Should Match "システムインデックスの状態を書き込めませんでした"
        (readTestSystemState $root).Covered.Contains("広報") | Should Be $true
        [System.IO.File]::Exists("$root\work\system_index\広報\${systemIndexFileName}") | Should Be $true
    }
}

Describe "indexer.ps1（まれな状況）" -Tag Io {
    $source = newSourceFolder "法務"

    It "既定のワークスペースにほかのファイルがあれば、ワークスペースに何も書かずに 1 で終わる" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "法務"; path = $source; enabled = $true })
        # 既定のワークスペースが使えないと判定された状態にする（空でないフォルダにエラーのファイルやログを書かない）
        $block = @{ Script = $indexerPath; Pattern = '^if \(\$workspaceBlock\) \{'; Action = {
                Set-Variable -Name workspaceBlock -Value "「C:\Users\test\Documents\tebunko」は空のフォルダではありません。" -Scope 1
            }
        }

        invokeIndexer $root @{} @($block) | Should Be 1

        @(Get-ChildItem -LiteralPath "$root\work" -Force).Count | Should Be 0
    }

    It "インデックスのフォルダを調べられなくても、確認を省いて取り込む" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "法務"; path = $source; enabled = $true })
        $unreadable = @{ Script = $indexerPath; Pattern = '^if \(\$null -eq \$indexCounts\) \{'; Action = {
                Set-Variable -Name indexCounts -Value $null -Scope 1
            }
        }

        invokeIndexer $root @{} @($unreadable) | Should Be 0

        (readTestLog $root) | Should Match "インデックスのフォルダを調べられないため"
        (readTestStatus $root).Rows["法務\議事録.docx"].状態 | Should Be ${stateDone}
    }

    It "失敗したファイルが表示の上限を超えたら、残りの件数を出す" {
        $root = newRoot
        writeTestSettings $root @(@{ name = "法務"; path = $source; enabled = $true })
        $limit = @{ Script = $indexerPath; Pattern = '^\s+if \(\$failures\.Count -gt \$failureListLimit\) \{'; Action = {
                Set-Variable -Name failureListLimit -Value 0 -Scope 1
            }
        }

        invokeIndexer $root @{} @($limit) | Should Be 0

        (readTestLog $root) | Should Match "ほか 1 件"
    }
}
