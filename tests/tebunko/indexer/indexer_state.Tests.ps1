# インデックス作成の状態ファイル（tebunko\indexer\indexer_state.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "readStatusFile / writeStatusFile / addStatusRow" -Tag Io {
    It "まとめて書き出した行と、1件ずつ追記した行の整形が同じ（速さのために別々に書いているため）" {
        # writeStatusFile は数万行を速く書くため、toStatusLine と同じ整形をその場に展開している
        $row = newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "10420" $stateFailed "" "2026/09/20 10:00:00" "エラー`tの`r`n説明"
        $path = "$TestDrive\status_same.tsv"
        writeStatusFile @() @($row) $path
        $written = @(readStatusLines $path)[-1]
        $written | Should Be (toStatusLine $row)
        $written | Should Be "営業\見積.xlsx`t2025/01/10 12:34:56`t10420`t${stateFailed}`t`t2026/09/20 10:00:00`tエラー の 説明`t"
        $done = newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "10420" $stateDone "3" "2026/09/20 10:00:00" "" "2"
        writeStatusFile @() @($done) $path
        @(readStatusLines $path)[-1] | Should Be (toStatusLine $done)
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
        $status.Folders.Count | Should Be 2
        $status.Folders[0].Path | Should Be "C:\data [1]"
        $status.Folders[0].Name | Should Be "data [1]"
        $status.Folders[1].Path | Should Be "D:\"
        $status.Folders[1].Name | Should Be "D"
        $status.Rows.Count | Should Be 2
        $row = $status.Rows["a\[確定]見積.xlsx"]
        $row.更新日時 | Should Be "2025/01/10 12:34:56"
        $row.サイズ | Should Be "10420"
        $row.状態 | Should Be $stateDone
        $row.TSV数 | Should Be "3"
        $status.Rows[" b.xls"].状態 | Should Be $stateNew
    }

    It "先頭にクロール対象フォルダ（パス・インデックス名）、次に見出しのTSVになる" {
        $path = "$TestDrive\status_lines.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        $lines = [System.IO.File]::ReadAllLines($path)
        $lines[0] | Should Be "クロール対象フォルダ`tC:\data`tdata"
        $lines[1] | Should Be "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー`t抽出版"
        $lines[2] | Should Be "a.xlsx`t2025/01/10 12:34:56`t1`t未取り込み`t`t`t`t"
    }

    It "抽出版を書き込んで読み込める" {
        $path = "$TestDrive\status_version.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "2" "" "" "2") $path
        addStatusRow (newStatusRow "b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "" "" "2") $path
        $status = readStatusFile $path
        $status.Rows["a.xlsx"].抽出版 | Should Be "2"
        $status.Rows["b.xlsx"].抽出版 | Should Be "2"
    }

    It "以前の形式（抽出版の列が無い）の行は抽出版を空として読む。今の形式で列が足りない行は無視する" {
        $path = "$TestDrive\status_no_version.tsv"
        [System.IO.File]::WriteAllLines($path, [string[]]@(
            "クロール対象フォルダ`tC:\data`tdata",
            "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー",
            "a.xlsx`t2025/01/10 12:34:56`t1`t済`t1`t`t"
        ), $utf8Bom)
        $status = readStatusFile $path
        $status.Rows["a.xlsx"].状態 | Should Be $stateDone
        $status.Rows["a.xlsx"].抽出版 | Should Be ""

        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "" "" "2") $path
        [System.IO.File]::AppendAllText($path, "a.xlsx`t2025/01/10 12:34:56`t1`t失敗`t`t`tエラー`r`n", $utf8Bom)  # 7 列（書き込みの途中）
        (readStatusFile $path).Rows["a.xlsx"].状態 | Should Be $stateDone
    }

    It "追記した行が前の行より優先される。相対パスの大文字・小文字は区別しない" {
        $path = "$TestDrive\status_append.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "Dir\A.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        addStatusRow (newStatusRow "dir\a.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 10:00:00" "パスワードが違います") $path

        $status = readStatusFile $path
        $status.Rows.Count | Should Be 1
        $status.Rows["DIR\A.XLSX"].状態 | Should Be $stateFailed
        $status.Rows["DIR\A.XLSX"].エラー | Should Be "パスワードが違います"
    }

    It "エラーメッセージのタブ・改行はスペースにして1行に収める" {
        $path = "$TestDrive\status_error.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "" "" $stateFailed "" "" "1行目`r`n2行目`tタブ") $path
        (readStatusFile $path).Rows["a.xlsx"].エラー | Should Be "1行目 2行目 タブ"
    }

    It "列数の合わない行（書き込み途中で中断した行）は無視する" {
        $path = "$TestDrive\status_broken.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1") $path
        [System.IO.File]::AppendAllText($path, "a.xlsx`t2025/01/10", $utf8Bom)

        $status = readStatusFile $path
        $status.Rows["a.xlsx"].状態 | Should Be $stateDone
    }

    It "書き直すと既存のファイルを置き換え、一時ファイルは残らない" {
        $path = "$TestDrive\status_replace.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\old"; Name = "old" }) @(newStatusRow "old\a.xlsx") $path
        writeStatusFile @([pscustomobject]@{ Path = "C:\new"; Name = "new" }) @(newStatusRow "new\b.xlsx") $path

        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Path }) | Should Be @("C:\new")
        @($status.Rows.Keys) | Should Be @("new\b.xlsx")
        Test-Path -LiteralPath "${path}.tmp" | Should Be $false
    }

    It "以前の形式（クロール対象フォルダが1つでインデックス名なし）も読める" {
        $path = "$TestDrive\status_legacy.tsv"
        [System.IO.File]::WriteAllLines($path, [string[]]@(
            "クロール対象フォルダ`tC:\old",
            "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー",
            "a.xlsx`t2025/01/10 12:34:56`t1`t済`t1`t`t"
        ), $utf8Bom)

        $status = readStatusFile $path
        $status.Folders[0].Path | Should Be "C:\old"
        $status.Folders[0].Name | Should Be ""
        $status.Rows["a.xlsx"].状態 | Should Be $stateDone
    }

    It "ファイルが無ければ空の一覧を返す" {
        $status = readStatusFile "$TestDrive\none_status.tsv"
        $status.Folders.Count | Should Be 0
        $status.Rows.Count | Should Be 0
    }
}

Describe "readIngestingFile / writeIngestingFile / removeIngestingFile" -Tag Io {
    It "書き込んだ相対パスと回数をそのまま読み込める（[ ] や空白を含むパス）" {
        $path = "$TestDrive\ingesting[1].txt"
        writeIngestingFile "フォルダ1\a [確定]\見積.xlsx" 2 $path

        $ingesting = readIngestingFile $path
        $ingesting.RelPath | Should Be "フォルダ1\a [確定]\見積.xlsx"
        $ingesting.Count | Should Be 2
    }

    It "削除すると記録なし（`$null）になる" {
        $path = "$TestDrive\ingesting_remove.txt"
        writeIngestingFile "a.xlsx" 1 $path
        removeIngestingFile $path

        Test-Path -LiteralPath $path | Should Be $false
        readIngestingFile $path | Should Be $null
    }

    It "ファイルが無くても削除でエラーにならない" {
        { removeIngestingFile "$TestDrive\none_ingesting.txt" } | Should Not Throw
    }

    It "壊れた記録（回数が数値でない・相対パスが無い・空）は `$null を返す" {
        $path = "$TestDrive\ingesting_broken.txt"
        foreach ($content in @("x`ta.xlsx", "0`ta.xlsx", "1`t", "a.xlsx", "")) {
            [System.IO.File]::WriteAllText($path, $content, $utf8Bom)
            readIngestingFile $path | Should Be $null
        }
    }
}

Describe "describeIngestError" -Tag Io {
    function newComError([string]$message, [string]$code) {
        return New-Object System.Runtime.InteropServices.COMException($message, [Convert]::ToInt32($code, 16))
    }

    It "パスワード付きのファイルは、Officeアプリの分かりにくいメッセージを付けずに原因だけを返す" {
        $expected = "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは取り込めません）"
        describeIngestError (newComError "入力したパスワードが間違っています。CapsLock キーの状態に注意して…" "800A03EC") | Should Be $expected
        describeIngestError (newComError "パスワードが正しくありません。文書を開けません。 (C:\Users\a\AppData\...\source.doc)" "800A1520") | Should Be $expected
        describeIngestError (newComError "Presentations.Open : 読み取りパスワードをもう一度入力してください(&P):" "80004005") | Should Be $expected
    }

    It "メソッド呼び出しの例外は中の例外のメッセージを使う" {
        $inner = newComError "Excel でファイル 'a.xlsx' を開くことができません。ファイル形式またはファイル拡張子が正しくありません。" "800A03EC"
        $outer = New-Object System.Management.Automation.MethodInvocationException('"7" 個の引数を指定して "Open" を呼び出し中に例外が発生しました', $inner)
        describeIngestError $outer | Should Be "ファイルが壊れているか、拡張子と中身の形式が一致していません（詳細: Excel でファイル 'a.xlsx' を開くことができません。ファイル形式またはファイル拡張子が正しくありません。）"
    }

    It "スクリプト自身が throw したメッセージはそのまま返す" {
        $exception = $null
        try { throw "ファイルが壊れているか、PowerPointのファイルではありません。" } catch { $exception = $_.Exception }
        describeIngestError $exception | Should Be "ファイルが壊れているか、PowerPointのファイルではありません。"
    }

    It "使用中・アクセス権なし・ファイルなしは原因を付けて元のメッセージを詳細にする" {
        $locked = New-Object System.IO.IOException("別のプロセスで使用されているため、アクセスできません。", [Convert]::ToInt32("80070020", 16))
        describeIngestError $locked | Should Be "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再取り込みしてください）（詳細: 別のプロセスで使用されているため、アクセスできません。）"
        describeIngestError (New-Object System.UnauthorizedAccessException("アクセスが拒否されました。")) | Should Match "^ファイルを読むアクセス権がありません（詳細: アクセスが拒否されました。）$"
        describeIngestError (New-Object System.IO.FileNotFoundException("見つかりません。")) | Should Match "^ファイルが見つかりません（"
    }

    It "Officeアプリの異常終了・応答なし・起動失敗は HRESULT で判断する" {
        describeIngestError (newComError "RPC サーバーを利用できません。" "800706BA") | Should Match "^Officeアプリが異常終了したか、内部でエラーが発生しました（.*（詳細: RPC サーバーを利用できません。）$"
        describeIngestError (newComError "呼び出し先が呼び出しを拒否しました。" "80010001") | Should Match "^Officeアプリが応答しませんでした"
        describeIngestError (newComError "クラスが登録されていません" "80040154") | Should Match "^Officeアプリ（Excel・Word・PowerPoint）を起動できませんでした"
    }

    It "メモリ不足（巨大なシート）は原因を付けて元のメッセージを詳細にする" {
        $inner = New-Object System.OutOfMemoryException("Exception of type 'System.OutOfMemoryException' was thrown.")
        $outer = New-Object System.Management.Automation.MethodInvocationException('"1" 個の引数を指定して "ReadAllText" を呼び出し中に例外が発生しました', $inner)
        describeIngestError $outer | Should Match "^シート・文書が大きすぎて取り込めません（メモリが不足しました）（詳細: "
    }

    It "原因が分からないものは元のメッセージ（改行は詰める）、メッセージが無ければエラーコードを返す" {
        describeIngestError (newComError "予期しない`r`nエラーです。" "800A03EC") | Should Be "予期しない エラーです。"
        describeIngestError (New-Object System.Exception(" ")) | Should Match "^エラーコード 0x[0-9A-F]{8}$"
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
        $state.Done | Should Be 1
        $state.Pending | Should Be 1
        $state.Failed | Should Be 2
        $state.FailedRows.Count | Should Be 2
        $state.FailedRows[0].相対パス | Should Be "data\新しい失敗.docx"
        $state.FailedRows[0].エラー | Should Be "原因B"
        $state.FailedRows[1].相対パス | Should Be "data\古い失敗.xlsx"
    }

    It "取り込み一覧が無ければ FailedRows は空" {
        $state = getIndexingState -path "$TestDrive\none.tsv"
        $state.Exists | Should Be $false
        @($state.FailedRows).Count | Should Be 0
    }
}

Describe "writeIndexingProgress / readIndexingProgress / removeIndexingProgress" -Tag Io {
    It "段階・件数・内容を往復できる" {
        $path = "$TestDrive\進捗1.txt"
        writeIndexingProgress ${indexingPhaseIngest} 12 34 5 "営業\見積.xlsx" $path
        $progress = readIndexingProgress $path
        $progress.Phase | Should Be ${indexingPhaseIngest}
        $progress.Processed | Should Be 12
        $progress.Remaining | Should Be 34
        $progress.Failed | Should Be 5
        $progress.Detail | Should Be "営業\見積.xlsx"
    }

    It "タブ・改行はスペースにする（1行に保つ）" {
        $path = "$TestDrive\進捗2.txt"
        writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "あ`tい`r`nう" $path
        (readIndexingProgress $path).Detail | Should Be "あ い う"
    }

    It "ファイルが無い・壊れていれば null" {
        readIndexingProgress "$TestDrive\進捗なし.txt" | Should BeNullOrEmpty
        $path = "$TestDrive\進捗3.txt"
        writeListFile $path @("見積`tあ`tい`tう`tえ")   # 件数が数値でない
        readIndexingProgress $path | Should BeNullOrEmpty
        writeListFile $path @("見積`t1`t2")              # 列が足りない（書き込みの途中）
        readIndexingProgress $path | Should BeNullOrEmpty
    }

    It "画面が読んでいる間も書ける（共有して開く）" {
        $path = "$TestDrive\進捗4.txt"
        writeIndexingProgress ${indexingPhaseIngest} 1 2 0 "はじめ" $path
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        try {
            { writeIndexingProgress ${indexingPhaseIngest} 2 1 0 "つぎ" $path } | Should Not Throw
        } finally {
            $stream.Dispose()
        }
        (readIndexingProgress $path).Detail | Should Be "つぎ"
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\進捗5.txt"
        writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "" $path
        removeIndexingProgress $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeIndexingProgress $path } | Should Not Throw
    }
}

Describe "writeIngestPlan / readIngestPlan / removeIngestPlan" -Tag Io {
    It "インデックスごとの件数を往復できる（件数は数値で返る）" {
        $path = "$TestDrive\予定1.tsv"
        $rows = @(
            (newIngestPlanRow "営業" "C:\data\営業" ${planKindIngest} 1234 12 5 7 0 0 3),
            (newIngestPlanRow "技術" "\\server\share\技術" ${planKindIngest} 20 0 0 0 0 0 0))
        writeIngestPlan $rows $path
        $plan = readIngestPlan $path
        $plan.Count | Should Be 2
        $plan[0].インデックス名 | Should Be "営業"
        $plan[0].元のフォルダ | Should Be "C:\data\営業"
        $plan[0].区分 | Should Be ${planKindIngest}
        ($plan[0].ファイル数 + 1) | Should Be 1235   # 文字列ではなく数値で返る
        $plan[0].取り込み対象 | Should Be 12
        $plan[0].新規 | Should Be 5
        $plan[0].更新あり | Should Be 7
        $plan[0].前回失敗 | Should Be 3
        $plan[1].取り込み対象 | Should Be 0
    }

    It "チェックなし・フォルダなしの区分も往復できる（件数は 0）" {
        $path = "$TestDrive\予定2.tsv"
        writeIngestPlan @(
            (newIngestPlanRow "外した" "D:\過去" ${planKindUnchecked}),
            (newIngestPlanRow "無い" "E:\USB" ${planKindMissing})) $path
        $plan = readIngestPlan $path
        $plan[0].区分 | Should Be ${planKindUnchecked}
        $plan[0].ファイル数 | Should Be 0
        $plan[1].区分 | Should Be ${planKindMissing}
    }

    It "インデックスが1件も無くても読める（空の配列）" {
        $path = "$TestDrive\予定3.tsv"
        writeIngestPlan @() $path
        (readIngestPlan $path).Count | Should Be 0
    }

    It "ファイルが無い・列が合わなければ null（画面は次の機会に読み直す）" {
        readIngestPlan "$TestDrive\予定なし.tsv" | Should BeNullOrEmpty
        $path = "$TestDrive\予定4.tsv"
        writeListFile $path @("べつの見出し")
        readIngestPlan $path | Should BeNullOrEmpty
    }

    It "画面が読んでいる間も書ける（共有して開く）" {
        $path = "$TestDrive\予定5.tsv"
        writeIngestPlan @((newIngestPlanRow "営業" "C:\data" ${planKindIngest} 1 1 1 0 0 0 0)) $path
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        try {
            { writeIngestPlan @((newIngestPlanRow "営業" "C:\data" ${planKindIngest} 2 2 2 0 0 0 0)) $path } | Should Not Throw
        } finally {
            $stream.Dispose()
        }
        (readIngestPlan $path)[0].取り込み対象 | Should Be 2
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\予定6.tsv"
        writeIngestPlan @() $path
        removeIngestPlan $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeIngestPlan $path } | Should Not Throw
    }
}

Describe "writeIndexingStartRequest / readIndexingStartRequest / removeIndexingStartRequest" -Tag Io {
    It "前回失敗したファイルも再取り込みするかを伝えられる" {
        $path = "$TestDrive\開始要求1"
        writeIndexingStartRequest $true $path
        (readIndexingStartRequest $path).RetryFailed | Should Be $true
        writeIndexingStartRequest $false $path
        (readIndexingStartRequest $path).RetryFailed | Should Be $false
    }

    It "まだ返事が無ければ null（インデクサは待ち続ける）" {
        readIndexingStartRequest "$TestDrive\開始要求なし" | Should BeNullOrEmpty
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\開始要求2"
        writeIndexingStartRequest $false $path
        removeIndexingStartRequest $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeIndexingStartRequest $path } | Should Not Throw
    }
}

Describe "ほかから共有せずに開かれているときの読み込み" -Tag Io {
    # インデクサが書き込み・置き換えをしている最中に画面が読む場合。例外にせず $null を返し、次の機会に読み直す
    function lockFile([string]$path) {
        return [System.IO.File]::Open($path, "Open", "ReadWrite", "None")
    }

    It "進み具合・取り込み予定・開始要求は `$null を返す" {
        $progress = "$TestDrive\lock_progress.txt"
        writeIndexingProgress ${indexingPhaseIngest} 1 2 0 "a" $progress
        $plan = "$TestDrive\lock_plan.tsv"
        writeIngestPlan @() $plan
        $request = "$TestDrive\lock_request"
        writeIndexingStartRequest $true $request

        foreach ($case in @(@($progress, { readIndexingProgress $progress }), @($plan, { readIngestPlan $plan }), @($request, { readIndexingStartRequest $request }))) {
            $stream = lockFile $case[0]
            try {
                & $case[1] | Should Be $null
            } finally {
                $stream.Dispose()
            }
        }
        # 開放されれば読める
        (readIndexingStartRequest $request).RetryFailed | Should Be $true
    }
}

Describe "describeIngestError（パスが長すぎる）" -Tag Io {
    It "原因を付けて元のメッセージを詳細にする" {
        describeIngestError (New-Object System.IO.PathTooLongException("長すぎます。")) | Should Be "パスが長すぎるため読めません（詳細: 長すぎます。）"
    }
}

Describe "getIndexingState（指定した時刻以降に取り込んだ件数）" -Tag Io {
    $path = "$TestDrive\status_since.tsv"
    writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(
        (newStatusRow "data\前.xlsx" "" "1" $stateDone "1" "2026/09/19 09:59:59"),
        (newStatusRow "data\同じ秒.xlsx" "" "1" $stateDone "1" "2026/09/19 10:00:00"),
        (newStatusRow "data\後.docx" "" "1" $stateFailed "" "2026/09/19 10:00:01" "原因"),
        (newStatusRow "data\日時なし.pptx" "" "1" $stateNew),
        (newStatusRow "data\日時が壊れている.xlsx" "" "1" $stateDone "1" "2026-09-19 11:00")
    ) $path

    It "取り込み日時が指定した時刻（秒未満は切り捨て）以降の行を数える。日時の無い・読めない行は数えない" {
        $state = getIndexingState ([datetime]"2026/09/19 10:00:00.700") $path
        $state.IngestedSince | Should Be 2
        $state.Total | Should Be 5
        $state.Done | Should Be 3
    }

    It "時刻を指定しなければ数えない" {
        (getIndexingState -path $path).IngestedSince | Should Be 0
    }
}

Describe "readStatusLines" -Tag Io {
    It "ファイルが無ければ空の配列" {
        @(readStatusLines "$TestDrive\無い一覧.tsv").Count | Should Be 0
    }
}

Describe "renameStatusIndexName / removeStatusIndexName（以前の形式の取り込み一覧）" -Tag Io {
    # 抽出版の列が無い以前の形式の取り込み一覧も readStatusFile は読めるため、名前の変更・削除も同じように行う
    function writeLegacyStatus([string]$path) {
        [System.IO.File]::WriteAllLines($path, [string[]]@(
            "クロール対象フォルダ`tC:\data`t営業",
            "クロール対象フォルダ`tD:\tech`t技術",
            "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー",
            "営業\a.xlsx`t2025/01/10 12:34:56`t1`t済`t1`t2026/09/18 10:00:00`t",
            "技術\b.docx`t2025/01/10 12:34:56`t1`t済`t1`t2026/09/18 10:00:00`t"
        ), $utf8Bom)
    }

    It "名前を変えると、以前の形式の行の相対パスも新しい名前にする" {
        $path = "$TestDrive\legacy_rename.tsv"
        writeLegacyStatus $path
        renameStatusIndexName "営業" "営業部" $path
        $status = readStatusFile $path
        $status.Folders[0].Name | Should Be "営業部"
        @($status.Rows.Keys | Sort-Object) -join "," | Should Be "営業部\a.xlsx,技術\b.docx"
        $status.Rows["営業部\a.xlsx"].状態 | Should Be $stateDone
    }

    It "削除すると、以前の形式の行も取り除く" {
        $path = "$TestDrive\legacy_remove.tsv"
        writeLegacyStatus $path
        removeStatusIndexName "営業" $path
        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Name }) -join "," | Should Be "技術"
        @($status.Rows.Keys) -join "," | Should Be "技術\b.docx"
    }
}
