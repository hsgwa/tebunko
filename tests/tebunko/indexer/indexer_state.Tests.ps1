# インデックス作成の状態ファイル（tebunko\indexer\indexer_state.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "readStatusFile / writeStatusFile / addStatusRow" -Tag Io {
    It "まとめて書き出した行と、1件ずつ追記した行の整形が同じ（速さのために別々に書いているため）" {
        # writeStatusFile は数万行を速く書くため、toStatusLine と同じ整形をその場に展開している
        $row = newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "10420" $stateFailed "" "2026/09/20 10:00:00" "エラー`tの`r`n説明"
        $path = "$TestDrive\status_same.tsv"
        writeStatusFile @() @($row) $path
        $written = @(readStatusLines $path)[-1]
        $written | Should -Be (toStatusLine $row)
        $written | Should -Be "営業\見積.xlsx`t2025/01/10 12:34:56`t10420`t${stateFailed}`t`t2026/09/20 10:00:00`tエラー の 説明`t"
        $done = newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "10420" $stateDone "3" "2026/09/20 10:00:00" "" "2"
        writeStatusFile @() @($done) $path
        @(readStatusLines $path)[-1] | Should -Be (toStatusLine $done)
    }

    It "書き込んだクロール対象フォルダと行をそのまま読み込める（[ ] や先頭の空白を含むパス）" {
        $path = "$TestDrive\status[1].tsv"
        $rows = @(
            (newStatusRow "a\[確定]見積.xlsx" "2025/01/10 12:34:56" "10420" $stateDone "3" "2026/09/19 10:00:00"),
            (newStatusRow " b.xls" "2025/02/01 08:00:00" "0" $stateNew)
        )
        $folders = @(
            [pscustomobject]@{ Path = "C:\data [1]"; Name = "data [1]" },
            [pscustomobject]@{ Path = "D:\"; Name = "D" }
        )
        writeStatusFile $folders $rows $path

        $status = readStatusFile $path
        $status.Folders.Count | Should -Be 2
        $status.Folders[0].Path | Should -Be "C:\data [1]"
        $status.Folders[0].Name | Should -Be "data [1]"
        $status.Folders[1].Path | Should -Be "D:\"
        $status.Folders[1].Name | Should -Be "D"
        $status.Rows.Count | Should -Be 2
        $row = $status.Rows["a\[確定]見積.xlsx"]
        $row.更新日時 | Should -Be "2025/01/10 12:34:56"
        $row.サイズ | Should -Be "10420"
        $row.状態 | Should -Be $stateDone
        $row.TSV数 | Should -Be "3"
        $status.Rows[" b.xls"].状態 | Should -Be $stateNew
    }

    It "先頭にクロール対象フォルダ（パス・インデックス名）、次に見出しのTSVになる" {
        $path = "$TestDrive\status_lines.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        $lines = [System.IO.File]::ReadAllLines($path)
        $lines[0] | Should -Be "クロール対象フォルダ`tC:\data`tdata"
        $lines[1] | Should -Be "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー`t抽出版"
        $lines[2] | Should -Be "a.xlsx`t2025/01/10 12:34:56`t1`t未取り込み`t`t`t`t"
    }

    It "抽出版を書き込んで読み込める" {
        $path = "$TestDrive\status_version.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "2" "" "" "2") $path
        addStatusRow (newStatusRow "b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "" "" "2") $path
        $status = readStatusFile $path
        $status.Rows["a.xlsx"].抽出版 | Should -Be "2"
        $status.Rows["b.xlsx"].抽出版 | Should -Be "2"
    }

    It "以前の形式（抽出版の列が無い）の行は抽出版を空として読む。今の形式で列が足りない行は無視する" {
        $path = "$TestDrive\status_no_version.tsv"
        [System.IO.File]::WriteAllLines($path, [string[]]@(
            "クロール対象フォルダ`tC:\data`tdata",
            "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー",
            "a.xlsx`t2025/01/10 12:34:56`t1`t済`t1`t`t"
        ), $utf8Bom)
        $status = readStatusFile $path
        $status.Rows["a.xlsx"].状態 | Should -Be $stateDone
        $status.Rows["a.xlsx"].抽出版 | Should -Be ""

        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "" "" "2") $path
        [System.IO.File]::AppendAllText($path, "a.xlsx`t2025/01/10 12:34:56`t1`t失敗`t`t`tエラー`r`n", $utf8Bom)  # 7 列（書き込みの途中）
        (readStatusFile $path).Rows["a.xlsx"].状態 | Should -Be $stateDone
    }

    It "追記した行が前の行より優先される。相対パスの大文字・小文字は区別しない" {
        $path = "$TestDrive\status_append.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "Dir\A.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        addStatusRow (newStatusRow "dir\a.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 10:00:00" "パスワードが違います") $path

        $status = readStatusFile $path
        $status.Rows.Count | Should -Be 1
        $status.Rows["DIR\A.XLSX"].状態 | Should -Be $stateFailed
        $status.Rows["DIR\A.XLSX"].エラー | Should -Be "パスワードが違います"
    }

    It "エラーメッセージのタブ・改行はスペースにして1行に収める" {
        $path = "$TestDrive\status_error.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "" "" $stateFailed "" "" "1行目`r`n2行目`tタブ") $path
        (readStatusFile $path).Rows["a.xlsx"].エラー | Should -Be "1行目 2行目 タブ"
    }

    It "列数の合わない行（書き込み途中で中断した行）は無視する" {
        $path = "$TestDrive\status_broken.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1") $path
        [System.IO.File]::AppendAllText($path, "a.xlsx`t2025/01/10", $utf8Bom)

        $status = readStatusFile $path
        $status.Rows["a.xlsx"].状態 | Should -Be $stateDone
    }

    It "書き直すと既存のファイルを置き換え、一時ファイルは残らない" {
        $path = "$TestDrive\status_replace.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\old"; Name = "old" }) @(newStatusRow "old\a.xlsx") $path
        writeStatusFile @([pscustomobject]@{ Path = "C:\new"; Name = "new" }) @(newStatusRow "new\b.xlsx") $path

        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Path }) | Should -Be @("C:\new")
        @($status.Rows.Keys) | Should -Be @("new\b.xlsx")
        Test-Path -LiteralPath "${path}.tmp" | Should -Be $false
    }

    It "以前の形式（クロール対象フォルダが1つでインデックス名なし）も読める" {
        $path = "$TestDrive\status_legacy.tsv"
        [System.IO.File]::WriteAllLines($path, [string[]]@(
            "クロール対象フォルダ`tC:\old",
            "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー",
            "a.xlsx`t2025/01/10 12:34:56`t1`t済`t1`t`t"
        ), $utf8Bom)

        $status = readStatusFile $path
        $status.Folders[0].Path | Should -Be "C:\old"
        $status.Folders[0].Name | Should -Be ""
        $status.Rows["a.xlsx"].状態 | Should -Be $stateDone
    }

    It "ファイルが無ければ空の一覧を返す" {
        $status = readStatusFile "$TestDrive\none_status.tsv"
        $status.Folders.Count | Should -Be 0
        $status.Rows.Count | Should -Be 0
    }
}

Describe "readIngestingFiles / writeIngestingFiles / removeIngestingFile" -Tag Io {
    It "取り込み中の複数のファイルの相対パスと回数をそのまま読み込める（[ ] や空白を含むパス）" {
        $path = "$TestDrive\ingesting[1].txt"
        writeIngestingFiles @(@{ RelPath = "フォルダ1\a [確定]\見積.xlsx"; Count = 2 }, @{ RelPath = "b.docx"; Count = 1 }) $path

        $ingesting = readIngestingFiles $path
        $ingesting.Count | Should -Be 2
        $ingesting[0].RelPath | Should -Be "フォルダ1\a [確定]\見積.xlsx"
        $ingesting[0].Count | Should -Be 2
        $ingesting[1].RelPath | Should -Be "b.docx"
    }

    It "取り込み中が無ければ記録を消す" {
        $path = "$TestDrive\ingesting_empty.txt"
        writeIngestingFiles @(@{ RelPath = "a.xlsx"; Count = 1 }) $path
        writeIngestingFiles @() $path
        Test-Path -LiteralPath $path | Should -Be $false
        (readIngestingFiles $path).Count | Should -Be 0
    }

    It "削除すると記録なし（空）になる。ファイルが無くても削除でエラーにならない" {
        $path = "$TestDrive\ingesting_remove.txt"
        writeIngestingFiles @(@{ RelPath = "a.xlsx"; Count = 1 }) $path
        removeIngestingFile $path
        Test-Path -LiteralPath $path | Should -Be $false
        (readIngestingFiles $path).Count | Should -Be 0
        { removeIngestingFile "$TestDrive\none_ingesting.txt" } | Should -Not -Throw
    }

    It "壊れた行（回数が数値でない・相対パスが無い・空）は読み飛ばす" {
        $path = "$TestDrive\ingesting_broken.txt"
        writeListFile $path @("x`ta.xlsx", "0`ta.xlsx", "1`t", "a.xlsx", "", "3`t残る.xlsx")
        $ingesting = readIngestingFiles $path
        $ingesting.Count | Should -Be 1
        $ingesting[0].RelPath | Should -Be "残る.xlsx"
        $ingesting[0].Count | Should -Be 3
    }
}

Describe "describeIngestError" -Tag Io {
    BeforeDiscovery {
        function newComError([string]$message, [string]$code) {
            return New-Object System.Runtime.InteropServices.COMException($message, [Convert]::ToInt32($code, 16))
        }

        $password = "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは取り込めません）"
    }

    # expected は返す文言そのもの、pattern は返す文言の形（元のメッセージの前に付ける原因など）
    It "<name>" -TestCases @(
        @{ name = "パスワード付きのファイルは、Officeアプリの分かりにくいメッセージを付けずに原因だけを返す（Excel）"; expected = $password
           exception = (newComError "入力したパスワードが間違っています。CapsLock キーの状態に注意して…" "800A03EC") }
        @{ name = "パスワード付きのファイルは、Officeアプリの分かりにくいメッセージを付けずに原因だけを返す（Word）"; expected = $password
           exception = (newComError "パスワードが正しくありません。文書を開けません。 (C:\Users\a\AppData\...\source.doc)" "800A1520") }
        @{ name = "パスワード付きのファイルは、Officeアプリの分かりにくいメッセージを付けずに原因だけを返す（PowerPoint）"; expected = $password
           exception = (newComError "Presentations.Open : 読み取りパスワードをもう一度入力してください(&P):" "80004005") }
        @{ name = "メソッド呼び出しの例外は中の例外のメッセージを使う"
           exception = (New-Object System.Management.Automation.MethodInvocationException('"7" 個の引数を指定して "Open" を呼び出し中に例外が発生しました',
               (newComError "Excel でファイル 'a.xlsx' を開くことができません。ファイル形式またはファイル拡張子が正しくありません。" "800A03EC")))
           expected = "ファイルが壊れているか、拡張子と中身の形式が一致していません（詳細: Excel でファイル 'a.xlsx' を開くことができません。ファイル形式またはファイル拡張子が正しくありません。）" }
        # throw "文字列" の例外は RuntimeException
        @{ name = "スクリプト自身が throw したメッセージはそのまま返す"
           exception = (New-Object System.Management.Automation.RuntimeException("ファイルが壊れているか、PowerPointのファイルではありません。"))
           expected = "ファイルが壊れているか、PowerPointのファイルではありません。" }
        @{ name = "使用中は原因を付けて元のメッセージを詳細にする"
           exception = (New-Object System.IO.IOException("別のプロセスで使用されているため、アクセスできません。", [Convert]::ToInt32("80070020", 16)))
           expected = "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再取り込みしてください）（詳細: 別のプロセスで使用されているため、アクセスできません。）" }
        @{ name = "アクセス権なしは原因を付けて元のメッセージを詳細にする"
           exception = (New-Object System.UnauthorizedAccessException("アクセスが拒否されました。"))
           expected = "ファイルを読むアクセス権がありません（詳細: アクセスが拒否されました。）" }
        @{ name = "ファイルなしは原因を付けて元のメッセージを詳細にする"
           exception = (New-Object System.IO.FileNotFoundException("見つかりません。")); pattern = "^ファイルが見つかりません（.*（詳細: 見つかりません。）$" }
        @{ name = "パスが長すぎるときは原因を付けて元のメッセージを詳細にする"
           exception = (New-Object System.IO.PathTooLongException("長すぎます。")); expected = "パスが長すぎるため読めません（詳細: 長すぎます。）" }
        @{ name = "Officeアプリの異常終了は HRESULT で判断する"
           exception = (newComError "RPC サーバーを利用できません。" "800706BA"); pattern = "^Officeアプリが異常終了したか、内部でエラーが発生しました（.*（詳細: RPC サーバーを利用できません。）$" }
        @{ name = "Officeアプリの応答なしは HRESULT で判断する"
           exception = (newComError "呼び出し先が呼び出しを拒否しました。" "80010001"); pattern = "^Officeアプリが応答しませんでした" }
        @{ name = "Officeアプリの起動失敗は HRESULT で判断する"
           exception = (newComError "クラスが登録されていません" "80040154"); pattern = "^Officeアプリ（Excel・Word・PowerPoint）を起動できませんでした" }
        @{ name = "メモリ不足（巨大なシート）は原因を付けて元のメッセージを詳細にする"
           exception = (New-Object System.Management.Automation.MethodInvocationException('"1" 個の引数を指定して "ReadAllText" を呼び出し中に例外が発生しました',
               (New-Object System.OutOfMemoryException("Exception of type 'System.OutOfMemoryException' was thrown."))))
           pattern = "^シート・文書が大きすぎて取り込めません（メモリが不足しました）（詳細: " }
        @{ name = "原因が分からないものは元のメッセージ（改行は詰める）"
           exception = (newComError "予期しない`r`nエラーです。" "800A03EC"); expected = "予期しない エラーです。" }
        @{ name = "メッセージが無ければエラーコードを返す"
           exception = (New-Object System.Exception(" ")); pattern = "^エラーコード 0x[0-9A-F]{8}$" }
    ) {
        param ($name, $exception, $expected, $pattern)
        if ($pattern) {
            describeIngestError $exception | Should -Match $pattern
        } else {
            describeIngestError $exception | Should -Be $expected
        }
    }
}

Describe "getIndexingState" -Tag Io {
    It "失敗したファイルの行を、取り込み日時の新しい順で FailedRows に返す" {
        $path = "$TestDrive\status_state.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(
            (newStatusRow "data\済.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/19 10:00:00"),
            (newStatusRow "data\古い失敗.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/18 09:00:00" "原因A"),
            (newStatusRow "data\新しい失敗.docx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 11:00:00" "原因B"),
            (newStatusRow "data\未取り込み.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        $state = getIndexingState -path $path
        $state.Done | Should -Be 1
        $state.Pending | Should -Be 1
        $state.Failed | Should -Be 2
        $state.FailedRows.Count | Should -Be 2
        $state.FailedRows[0].相対パス | Should -Be "data\新しい失敗.docx"
        $state.FailedRows[0].エラー | Should -Be "原因B"
        $state.FailedRows[1].相対パス | Should -Be "data\古い失敗.xlsx"
    }

    It "取り込み一覧が無ければ FailedRows は空" {
        $state = getIndexingState -path "$TestDrive\none.tsv"
        $state.Exists | Should -Be $false
        @($state.FailedRows).Count | Should -Be 0
    }
}

Describe "newIndexerChannel / writeIndexingProgress / readIndexingProgress" -Tag Unit {
    It "画面が決めた条件を入れ、スレッドをまたいで使える形にする" {
        $channel = newIndexerChannel $true $true 2
        $channel.RetryFailed | Should -Be $true
        $channel.ConfirmTargets | Should -Be $true
        $channel.Workers | Should -Be 2
        $channel.IsSynchronized | Should -Be $true
        $channel.Stop | Should -Be $false
        $channel.ExitCode | Should -BeNullOrEmpty
        $channel.OfficePids.Count | Should -Be 0
    }

    It "段階・件数・内容を往復できる。タブ・改行はスペースにする" {
        $channel = newIndexerChannel
        readIndexingProgress $channel | Should -BeNullOrEmpty
        writeIndexingProgress ${indexingPhaseIngest} 12 34 5 "営業\見積`t.xlsx" $channel
        $progress = readIndexingProgress $channel
        $progress.Phase | Should -Be ${indexingPhaseIngest}
        $progress.Processed | Should -Be 12
        $progress.Remaining | Should -Be 34
        $progress.Failed | Should -Be 5
        $progress.Detail | Should -Be "営業\見積 .xlsx"
    }

    It "受け渡しの口が無ければ何もしない" {
        { writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "a" $null } | Should -Not -Throw
        readIndexingProgress $null | Should -BeNullOrEmpty
    }
}

Describe "requestIndexingStop / answerIndexingPlan" -Tag Unit {
    It "中止を求めると、確認を待っていても返事を待つのをやめる" {
        $channel = newIndexerChannel
        requestIndexingStop $channel
        $channel.Stop | Should -Be $true
        $channel.Answered.WaitOne(0) | Should -Be $true
    }

    It "取り込む返事を渡す（中止にはしない）" {
        $channel = newIndexerChannel
        answerIndexingPlan $channel @{ RetryFailed = $true }
        $channel.Answer.RetryFailed | Should -Be $true
        $channel.Stop | Should -Be $false
        $channel.Answered.WaitOne(0) | Should -Be $true
    }

    It "取りやめの返事（`$null）は中止にする" {
        $channel = newIndexerChannel
        answerIndexingPlan $channel $null
        $channel.Stop | Should -Be $true
    }
}

Describe "testIndexerRunning" -Tag Io {
    It "インデックス作成の鍵をほかが持っていれば `$true、持っていなければ `$false（調べた後は鍵を持たない）" {
        $work = Join-Path $TestDrive "running_work"
        testIndexerRunning $work | Should -Be $false
        testIndexerRunning $work | Should -Be $false
        $other = newAppMutex "indexer" $work
        try {
            # 同じスレッドが持つ鍵も「動いている」と数える（画面のインデクサのスレッドが持っているとき）
            testIndexerRunning $work | Should -Be $true
        } finally {
            $other.Mutex.ReleaseMutex()
            $other.Mutex.Dispose()
        }
        testIndexerRunning $work | Should -Be $false
    }
}

Describe "writeIndexerLog" -Tag Unit {
    AfterEach {
        $script:indexerLog = $null
        $script:indexerEcho = $false
    }

    It "ログを開いていればログに書く。開いていなければ何もしない" {
        { writeIndexerLog "何もしない" } | Should -Not -Throw
        $script:indexerLog = New-Object System.IO.StringWriter
        writeIndexerLog "1 行目"
        writeIndexerLog "2 行目" "Yellow"
        $script:indexerLog.ToString() | Should -Be "1 行目`r`n2 行目`r`n"
    }

    It "コンソールに出すときは色を付ける" {
        Mock Write-Host {}
        $script:indexerEcho = $true
        writeIndexerLog "注意" "Yellow"
        writeIndexerLog "普通"
        Should -Invoke Write-Host -Times 1 -Exactly -Scope It -ParameterFilter { "$Object" -eq "注意" -and $ForegroundColor -eq "Yellow" }
        Should -Invoke Write-Host -Times 1 -Exactly -Scope It -ParameterFilter { "$Object" -eq "普通" -and !$ForegroundColor }
    }

    It "ログに書けなくても止めない" {
        $writer = New-Object System.IO.StringWriter
        $writer.Dispose()
        $script:indexerLog = $writer
        { writeIndexerLog "書けない" } | Should -Not -Throw
    }
}

Describe "getIndexingState（指定した時刻以降に取り込んだ件数）" -Tag Io {
    BeforeAll {
        $path = "$TestDrive\status_since.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(
            (newStatusRow "data\前.xlsx" "" "1" $stateDone "1" "2026/09/19 09:59:59"),
            (newStatusRow "data\同じ秒.xlsx" "" "1" $stateDone "1" "2026/09/19 10:00:00"),
            (newStatusRow "data\後.docx" "" "1" $stateFailed "" "2026/09/19 10:00:01" "原因"),
            (newStatusRow "data\日時なし.pptx" "" "1" $stateNew),
            (newStatusRow "data\日時が壊れている.xlsx" "" "1" $stateDone "1" "2026-09-19 11:00")
        ) $path
    }

    It "取り込み日時が指定した時刻（秒未満は切り捨て）以降の行を数える。日時の無い・読めない行は数えない" {
        $state = getIndexingState ([datetime]"2026/09/19 10:00:00.700") $path
        $state.IngestedSince | Should -Be 2
        $state.Total | Should -Be 5
        $state.Done | Should -Be 3
    }

    It "時刻を指定しなければ数えない" {
        (getIndexingState -path $path).IngestedSince | Should -Be 0
    }
}

Describe "renameStatusIndexName / removeStatusIndexName（以前の形式の取り込み一覧）" -Tag Io {
    # 抽出版の列が無い以前の形式の取り込み一覧も readStatusFile は読めるため、名前の変更・削除も同じように行う
    BeforeAll {
        function writeLegacyStatus([string]$path) {
            [System.IO.File]::WriteAllLines($path, [string[]]@(
                "クロール対象フォルダ`tC:\data`t営業",
                "クロール対象フォルダ`tD:\tech`t技術",
                "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー",
                "営業\a.xlsx`t2025/01/10 12:34:56`t1`t済`t1`t2026/09/18 10:00:00`t",
                "技術\b.docx`t2025/01/10 12:34:56`t1`t済`t1`t2026/09/18 10:00:00`t"
            ), $utf8Bom)
        }
    }

    It "名前を変えると、以前の形式の行の相対パスも新しい名前にする" {
        $path = "$TestDrive\legacy_rename.tsv"
        writeLegacyStatus $path
        renameStatusIndexName "営業" "営業部" $path
        $status = readStatusFile $path
        $status.Folders[0].Name | Should -Be "営業部"
        @($status.Rows.Keys | Sort-Object) -join "," | Should -Be "営業部\a.xlsx,技術\b.docx"
        $status.Rows["営業部\a.xlsx"].状態 | Should -Be $stateDone
    }

    It "削除すると、以前の形式の行も取り除く" {
        $path = "$TestDrive\legacy_remove.tsv"
        writeLegacyStatus $path
        removeStatusIndexName "営業" $path
        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Name }) -join "," | Should -Be "技術"
        @($status.Rows.Keys) -join "," | Should -Be "技術\b.docx"
    }
}
