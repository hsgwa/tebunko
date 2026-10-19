# インデックス作成の状態ファイル（tebunko\indexer\indexer_state.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    # ZipSizeLimitException（writeZipSizeLimitLog のテストで使う）は office_reader.ps1 にあり、
    # 本番の読み込み口（lib.ps1）には含まれない（indexer_lib.ps1 が別に読み込む）ため、ここで直接読み込む
    . "${scriptsDir}\shared\office\office_reader.ps1"
}

Describe "describeIngestError" -Tag Io {
    BeforeDiscovery {
        function newComError([string]$message, [string]$code) {
            return New-Object System.Runtime.InteropServices.COMException($message, [Convert]::ToInt32($code, 16))
        }

        $password = "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは更新できません）"
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
           expected = "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再度更新してください）（詳細: 別のプロセスで使用されているため、アクセスできません。）" }
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
           pattern = "^シート・文書が大きすぎて更新できません（メモリが不足しました）（詳細: " }
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

    It "「未取り込み」でエラー（理由）が書いてある行（利用者のPowerPointのために取り込めなかったファイル）は、待ちではなく失敗に数え、FailedRows に理由付きで返す。エラーが空の「未取り込み」は待ちのまま" {
        $path = "$TestDrive\status_state_postponed.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(
            (newStatusRow "data\済.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/19 10:00:00"),
            (newStatusRow "data\PowerPoint.ppt" "2025/01/10 12:34:56" "1" $stateNew "" "" ${postponedReasonRunning}),
            (newStatusRow "data\未取り込み.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        $state = getIndexingState -path $path
        $state.Done | Should -Be 1
        $state.Pending | Should -Be 1
        $state.Failed | Should -Be 1
        $state.FailedRows.Count | Should -Be 1
        $state.FailedRows[0].相対パス | Should -Be "data\PowerPoint.ppt"
        $state.FailedRows[0].エラー | Should -Be ${postponedReasonRunning}
        $state.IndexStats["data"].Failed | Should -Be 1
        $state.IndexStats["data"].Pending | Should -Be 1
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
        $channel.UserOfficePids.Count | Should -Be 0
    }

    It "クラウドにだけあるファイルをダウンロードして取り込むか（IncludeCloud）は既定で取り込まない。残した件数（CloudSkipped）は 0" {
        $channel = newIndexerChannel
        $channel.IncludeCloud | Should -Be $false
        $channel.CloudSkipped | Should -Be 0
        (newIndexerChannel -includeCloud $true).IncludeCloud | Should -Be $true
    }

    It "選んだインデックス名（OnlyNames）は、空の名前を除いて入れる。無ければ空" {
        (newIndexerChannel).OnlyNames.Count | Should -Be 0
        $channel = newIndexerChannel $false $false -1 @("営業", "", "技術")
        @($channel.OnlyNames) | Should -Be @("営業", "技術")
        $channel.OnlySkipped.Count | Should -Be 0
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

    It "ダウンロードして取り込む返事（IncludeCloud）もそのまま渡す" {
        $channel = newIndexerChannel
        answerIndexingPlan $channel @{ RetryFailed = $false; IncludeCloud = $true }
        $channel.Answer.IncludeCloud | Should -Be $true
        $channel.Stop | Should -Be $false
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
        $other = newAppMutex "indexer" $work "tebunko"
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

Describe "addStatusRow" -Tag Io {
    It "追記した行にも抽出版を書き込み、読み込める" {
        $path = "$TestDrive\status_version.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "data\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "2" "" "" "2") $path
        addStatusRow (newStatusRow "data\b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "" "" "2") $path
        $status = readStatusFile $path
        $status.Rows["data\a.xlsx"].抽出版 | Should -Be "2"
        $status.Rows["data\b.xlsx"].抽出版 | Should -Be "2"
    }

    It "追記した行が前の行より優先される。相対パスの大文字・小文字は区別しない" {
        $path = "$TestDrive\status_append.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "data\Dir\A.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        addStatusRow (newStatusRow "data\dir\a.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 10:00:00" "パスワードが違います") $path

        $status = readStatusFile $path
        $status.Rows.Count | Should -Be 1
        $status.Rows["DATA\DIR\A.XLSX"].状態 | Should -Be $stateFailed
        $status.Rows["DATA\DIR\A.XLSX"].エラー | Should -Be "パスワードが違います"
    }
}

Describe "writeZipSizeLimitLog" -Tag Unit {
    BeforeAll {
        # [ZipSizeLimitException] の型リテラルは、Pester の It ブロックからは直接解決できない
        # （BeforeAll で dot-source したクラスでも、実行スコープが分かれるため）。
        # BeforeAll で定義したこの関数の中でだけ型リテラルを使い、It からは関数越しに呼ぶ
        function newZipSizeLimitException([string]$message, [string]$partName, [long]$measuredBytes, [string]$limitKind) {
            return [ZipSizeLimitException]::new($message, $partName, $measuredBytes, $limitKind)
        }
    }

    AfterEach {
        $script:indexerLog = $null
    }

    It "<limitKind>: 部品名・大きさ・上限の種類をログに書く（Output は文字数）" -TestCases @(
        @{ limitKind = "Part"; kindText = "部品ごと"; sizeText = "123456789 バイト、約117.7MB" }
        @{ limitKind = "Total"; kindText = "1ファイルの合計"; sizeText = "123456789 バイト、約117.7MB" }
        @{ limitKind = "Output"; kindText = "埋め込みから出す文字の合計"; sizeText = "123456789 文字、約117.7M文字" }
    ) {
        param ($limitKind, $kindText, $sizeText)
        $script:indexerLog = New-Object System.IO.StringWriter
        $exception = newZipSizeLimitException "ファイルサイズが大きすぎるため更新できません。" "xl/charts/chart1.xml" 123456789 $limitKind
        writeZipSizeLimitLog $exception
        $script:indexerLog.ToString() | Should -Be "    サイズの上限（${kindText}）を超えました: xl/charts/chart1.xml（${sizeText}）`r`n"
    }

    It "ZipSizeLimitException でなければ何もしない" {
        $script:indexerLog = New-Object System.IO.StringWriter
        writeZipSizeLimitLog (New-Object System.Exception("ふつうの失敗です。"))
        $script:indexerLog.ToString() | Should -Be ""
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

Describe "renameStatusIndexName / removeStatusIndexName" -Tag Io {
    BeforeAll {
        function writeTwoIndexStatus([string]$path) {
            writeStatusFile @(
                [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
                [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
            ) @(
                (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00" "" "2"),
                (newStatusRow "技術\b.docx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00" "" "2")
            ) $path
        }
    }

    It "名前を変えると、クロール対象フォルダの行と、そのインデックスの行の相対パスを新しい名前にする" {
        $path = "$TestDrive\status_rename.tsv"
        writeTwoIndexStatus $path
        renameStatusIndexName "営業" "営業部" $path
        $status = readStatusFile $path
        $status.Folders[0].Name | Should -Be "営業部"
        @($status.Rows.Keys | Sort-Object) -join "," | Should -Be "営業部\a.xlsx,技術\b.docx"
        $status.Rows["営業部\a.xlsx"].状態 | Should -Be $stateDone
        $status.Rows["営業部\a.xlsx"].抽出版 | Should -Be "2"
    }

    It "削除すると、クロール対象フォルダの行と、そのインデックスの行を取り除く" {
        $path = "$TestDrive\status_remove.tsv"
        writeTwoIndexStatus $path
        removeStatusIndexName "営業" $path
        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Name }) -join "," | Should -Be "技術"
        @($status.Rows.Keys) -join "," | Should -Be "技術\b.docx"
    }
}

Describe "StatusLedger" -Tag Io {
    BeforeAll {
        # 場所を暗黙に使わないため、コンストラクタで渡す保存先（Workspace の代わりに、同じプロパティを持つ値で差し替える）
        function newTestWorkspace([string]$dir) {
            [System.IO.Directory]::CreateDirectory($dir) | Out-Null
            return [pscustomobject]@{ StatusFile = "$dir\ingest_status.tsv" }
        }
    }

    It "ReadStatus / WriteStatus / AddRow は、渡した保存先（Workspace の StatusFile）を読み書きする" {
        $ws = newTestWorkspace "$TestDrive\ledger1"
        $ledger = [StatusLedger]::new($ws)
        $ledger.WriteStatus(@([pscustomobject]@{ Path = "C:\data"; Name = "data" }), @((newStatusRow "data\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)))
        $ledger.AddRow((newStatusRow "data\b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"))

        $status = $ledger.ReadStatus()
        $status.Rows.Count | Should -Be 2
        $status.Rows["data\b.xlsx"].状態 | Should -Be $stateDone
        # 生の関数で読んでも同じ場所を指す
        (readStatusFile $ws.StatusFile).Rows.Count | Should -Be 2
    }

    It "ReadRelPaths は、渡した保存先（StatusFile）のデータ行の相対パスを、追記した行も数えてファイルの順に返す" {
        $ws = newTestWorkspace "$TestDrive\ledger2"
        $ledger = [StatusLedger]::new($ws)
        $ledger.ReadRelPaths().Count | Should -Be 0
        $ledger.WriteStatus(@([pscustomobject]@{ Path = "C:\data"; Name = "data" }), @((newStatusRow "data\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)))
        $ledger.AddRow((newStatusRow "data\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"))
        @($ledger.ReadRelPaths()) | Should -Be @("data\a.xlsx", "data\a.xlsx")
    }

    It "RenameIndexName / RemoveIndexName は、renameStatusIndexName / removeStatusIndexName と同じ結果になる" {
        $ws = newTestWorkspace "$TestDrive\ledger3"
        $ledger = [StatusLedger]::new($ws)
        $ledger.WriteStatus(@(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ), @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "技術\b.docx" "2025/01/10 12:34:56" "1" $stateDone "1")
        ))

        $ledger.RenameIndexName("営業", "営業部")
        $status = $ledger.ReadStatus()
        $status.Folders[0].Name | Should -Be "営業部"
        $status.Rows.ContainsKey("営業部\a.xlsx") | Should -Be $true

        $ledger.RemoveIndexName("技術")
        $status = $ledger.ReadStatus()
        @($status.Folders | ForEach-Object { $_.Name }) | Should -Be @("営業部")
    }

    It "AddFailure / AddDropped は、今回の失敗・元ファイルが無くなった行を記録する（新しく作ったときは空）" {
        $ws = newTestWorkspace "$TestDrive\ledger4"
        $ledger = [StatusLedger]::new($ws)
        $ledger.Failures.Count | Should -Be 0
        $ledger.DroppedRows.Count | Should -Be 0

        $ledger.AddFailure("営業\壊れた.pptx", "取り込みに失敗しました")
        $ledger.Failures.Count | Should -Be 1
        $ledger.Failures[0].RelPath | Should -Be "営業\壊れた.pptx"
        $ledger.Failures[0].Message | Should -Be "取り込みに失敗しました"

        $ledger.AddDropped("営業\消えた.xlsx")
        $ledger.DroppedRows.Contains("営業\消えた.xlsx") | Should -Be $true
        # 相対パスの大文字・小文字は区別しない（取り込み一覧と同じ比べ方）
        $ledger.DroppedRows.Contains("営業\消えた.XLSX") | Should -Be $true
    }
}

Describe "newIngestPlanRow（クラウドにだけあるファイル）" -Tag Unit {
    It "クラウドの件数・失敗の件数・合計サイズを入れる。渡さなければ 0" {
        $row = newIngestPlanRow "営業" "C:\data" ${planKindIngest} 10 2 1 1 0 0 0 3 1 3000 500
        $row.クラウド | Should -Be 3
        $row.クラウド失敗 | Should -Be 1
        $row.クラウド容量 | Should -Be 3000
        $row.クラウド失敗容量 | Should -Be 500
        $empty = newIngestPlanRow "営業" "C:\data"
        $empty.クラウド | Should -Be 0
        $empty.クラウド容量 | Should -Be 0
    }

    It "列は ingestPlanColumns と同じ" {
        $row = newIngestPlanRow "営業" "C:\data"
        @($row.PSObject.Properties.Name) | Should -Be @(${ingestPlanColumns})
    }
}

Describe "readPendingRemovals / writePendingRemovals" -Tag Io {
    It "書いた記録を読み戻せる。記録が無ければ空" {
        $path = "$TestDrive\store\pending_remove.tsv"
        @(readPendingRemovals $path).Count | Should -Be 0

        writePendingRemovals @(@{ Folder = "営業\2024"; Name = "A社 [確定].xlsx" }, @{ Folder = "経理"; Name = "b.docx" }) $path

        $read = @(readPendingRemovals $path)
        $read.Count | Should -Be 2
        ($read | Where-Object { $_.Folder -eq "営業\2024" }).Name | Should -Be "A社 [確定].xlsx"
        [System.IO.File]::Exists("$path.tmp") | Should -Be $false
    }

    It "0 件を書くと記録を消す（無くても失敗しない）" {
        $path = "$TestDrive\store2\pending_remove.tsv"
        writePendingRemovals @(@{ Folder = "営業"; Name = "a.xlsx" }) $path
        [System.IO.File]::Exists($path) | Should -Be $true

        writePendingRemovals @() $path
        [System.IO.File]::Exists($path) | Should -Be $false
        { writePendingRemovals @() $path } | Should -Not -Throw
    }

    It "置き場の外を指す行（..）は読み飛ばす" {
        $path = "$TestDrive\store3\pending_remove.tsv"
        writeListFile $path @("営業`tok.xlsx", "..\外`tx.xlsx", "営業\..\..`ty.xlsx")

        $read = @(readPendingRemovals $path)
        $read.Count | Should -Be 1
        $read[0].Name | Should -Be "ok.xlsx"
    }
}