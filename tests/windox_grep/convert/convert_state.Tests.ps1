# 変換の状態ファイル（windox_grep\convert\convert_state.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "readStatusFile / writeStatusFile / addStatusRow" -Tag Io {
    It "まとめて書き出した行と、1件ずつ追記した行の整形が同じ（速さのために別々に書いているため）" {
        # writeStatusFile は数万行を速く書くため、toStatusLine と同じ整形をその場に展開している
        $row = newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "10420" $stateFailed "" "2026/09/20 10:00:00" "エラー`tの`r`n説明"
        $path = "$TestDrive\status_same.tsv"
        writeStatusFile @() @($row) $path
        $written = @(readStatusLines $path)[-1]
        $written | Should Be (toStatusLine $row)
        $written | Should Be "営業\見積.xlsx`t2025/01/10 12:34:56`t10420`t${stateFailed}`t`t2026/09/20 10:00:00`tエラー の 説明"
    }

    It "書き込んだ変換対象フォルダと行をそのまま読み込める（[ ] や先頭の空白を含むパス）" {
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

    It "先頭に変換対象フォルダ（パス・インデックス名）、次に見出しのTSVになる" {
        $path = "$TestDrive\status_lines.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        $lines = [System.IO.File]::ReadAllLines($path)
        $lines[0] | Should Be "変換対象フォルダ`tC:\data`tdata"
        $lines[1] | Should Be "相対パス`t更新日時`tサイズ`t状態`tTSV数`t変換日時`tエラー"
        $lines[2] | Should Be "a.xlsx`t2025/01/10 12:34:56`t1`t未変換`t`t`t"
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

    It "以前の形式（変換対象フォルダが1つでインデックス名なし）も読める" {
        $path = "$TestDrive\status_legacy.tsv"
        [System.IO.File]::WriteAllLines($path, [string[]]@(
            "変換対象フォルダ`tC:\old",
            "相対パス`t更新日時`tサイズ`t状態`tTSV数`t変換日時`tエラー",
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

Describe "readConvertingFile / writeConvertingFile / removeConvertingFile" -Tag Io {
    It "書き込んだ相対パスと回数をそのまま読み込める（[ ] や空白を含むパス）" {
        $path = "$TestDrive\converting[1].txt"
        writeConvertingFile "フォルダ1\a [確定]\見積.xlsx" 2 $path

        $converting = readConvertingFile $path
        $converting.RelPath | Should Be "フォルダ1\a [確定]\見積.xlsx"
        $converting.Count | Should Be 2
    }

    It "削除すると記録なし（`$null）になる" {
        $path = "$TestDrive\converting_remove.txt"
        writeConvertingFile "a.xlsx" 1 $path
        removeConvertingFile $path

        Test-Path -LiteralPath $path | Should Be $false
        readConvertingFile $path | Should Be $null
    }

    It "ファイルが無くても削除でエラーにならない" {
        { removeConvertingFile "$TestDrive\none_converting.txt" } | Should Not Throw
    }

    It "壊れた記録（回数が数値でない・相対パスが無い・空）は `$null を返す" {
        $path = "$TestDrive\converting_broken.txt"
        foreach ($content in @("x`ta.xlsx", "0`ta.xlsx", "1`t", "a.xlsx", "")) {
            [System.IO.File]::WriteAllText($path, $content, $utf8Bom)
            readConvertingFile $path | Should Be $null
        }
    }
}

Describe "describeConvertError" -Tag Io {
    function newComError([string]$message, [string]$code) {
        return New-Object System.Runtime.InteropServices.COMException($message, [Convert]::ToInt32($code, 16))
    }

    It "パスワード付きのファイルは、Officeアプリの分かりにくいメッセージを付けずに原因だけを返す" {
        $expected = "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは変換できません）"
        describeConvertError (newComError "入力したパスワードが間違っています。CapsLock キーの状態に注意して…" "800A03EC") | Should Be $expected
        describeConvertError (newComError "パスワードが正しくありません。文書を開けません。 (C:\Users\a\AppData\...\source.doc)" "800A1520") | Should Be $expected
        describeConvertError (newComError "Presentations.Open : 読み取りパスワードをもう一度入力してください(&P):" "80004005") | Should Be $expected
    }

    It "メソッド呼び出しの例外は中の例外のメッセージを使う" {
        $inner = newComError "Excel でファイル 'a.xlsx' を開くことができません。ファイル形式またはファイル拡張子が正しくありません。" "800A03EC"
        $outer = New-Object System.Management.Automation.MethodInvocationException('"7" 個の引数を指定して "Open" を呼び出し中に例外が発生しました', $inner)
        describeConvertError $outer | Should Be "ファイルが壊れているか、拡張子と中身の形式が一致していません（詳細: Excel でファイル 'a.xlsx' を開くことができません。ファイル形式またはファイル拡張子が正しくありません。）"
    }

    It "スクリプト自身が throw したメッセージはそのまま返す" {
        $exception = $null
        try { throw "ファイルが壊れているか、PowerPointのファイルではありません。" } catch { $exception = $_.Exception }
        describeConvertError $exception | Should Be "ファイルが壊れているか、PowerPointのファイルではありません。"
    }

    It "使用中・アクセス権なし・ファイルなしは原因を付けて元のメッセージを詳細にする" {
        $locked = New-Object System.IO.IOException("別のプロセスで使用されているため、アクセスできません。", [Convert]::ToInt32("80070020", 16))
        describeConvertError $locked | Should Be "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再変換してください）（詳細: 別のプロセスで使用されているため、アクセスできません。）"
        describeConvertError (New-Object System.UnauthorizedAccessException("アクセスが拒否されました。")) | Should Match "^ファイルを読むアクセス権がありません（詳細: アクセスが拒否されました。）$"
        describeConvertError (New-Object System.IO.FileNotFoundException("見つかりません。")) | Should Match "^ファイルが見つかりません（"
    }

    It "Officeアプリの異常終了・応答なし・起動失敗は HRESULT で判断する" {
        describeConvertError (newComError "RPC サーバーを利用できません。" "800706BA") | Should Match "^Officeアプリが異常終了したか、内部でエラーが発生しました（.*（詳細: RPC サーバーを利用できません。）$"
        describeConvertError (newComError "呼び出し先が呼び出しを拒否しました。" "80010001") | Should Match "^Officeアプリが応答しませんでした"
        describeConvertError (newComError "クラスが登録されていません" "80040154") | Should Match "^Officeアプリ（Excel・Word・PowerPoint）を起動できませんでした"
    }

    It "メモリ不足（巨大なシート）は原因を付けて元のメッセージを詳細にする" {
        $inner = New-Object System.OutOfMemoryException("Exception of type 'System.OutOfMemoryException' was thrown.")
        $outer = New-Object System.Management.Automation.MethodInvocationException('"1" 個の引数を指定して "ReadAllText" を呼び出し中に例外が発生しました', $inner)
        describeConvertError $outer | Should Match "^シート・文書が大きすぎて変換できません（メモリが不足しました）（詳細: "
    }

    It "原因が分からないものは元のメッセージ（改行は詰める）、メッセージが無ければエラーコードを返す" {
        describeConvertError (newComError "予期しない`r`nエラーです。" "800A03EC") | Should Be "予期しない エラーです。"
        describeConvertError (New-Object System.Exception(" ")) | Should Match "^エラーコード 0x[0-9A-F]{8}$"
    }
}

Describe "getConversionState" -Tag Io {
    It "失敗したファイルの行を、変換日時の新しい順で FailedRows に返す" {
        $path = "$TestDrive\status_state.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(
            (newStatusRow "data\済.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/19 10:00:00"),
            (newStatusRow "data\古い失敗.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/18 09:00:00" "原因A"),
            (newStatusRow "data\新しい失敗.docx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 11:00:00" "原因B"),
            (newStatusRow "data\未変換.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        $state = getConversionState -path $path
        $state.Done | Should Be 1
        $state.Pending | Should Be 1
        $state.Failed | Should Be 2
        $state.FailedRows.Count | Should Be 2
        $state.FailedRows[0].相対パス | Should Be "data\新しい失敗.docx"
        $state.FailedRows[0].エラー | Should Be "原因B"
        $state.FailedRows[1].相対パス | Should Be "data\古い失敗.xlsx"
    }

    It "変換一覧が無ければ FailedRows は空" {
        $state = getConversionState -path "$TestDrive\none.tsv"
        $state.Exists | Should Be $false
        @($state.FailedRows).Count | Should Be 0
    }
}

Describe "writeConvertProgress / readConvertProgress / removeConvertProgress" -Tag Io {
    It "段階・件数・内容を往復できる" {
        $path = "$TestDrive\進捗1.txt"
        writeConvertProgress ${convertPhaseRun} 12 34 5 "営業\見積.xlsx" $path
        $progress = readConvertProgress $path
        $progress.Phase | Should Be ${convertPhaseRun}
        $progress.Processed | Should Be 12
        $progress.Remaining | Should Be 34
        $progress.Failed | Should Be 5
        $progress.Detail | Should Be "営業\見積.xlsx"
    }

    It "タブ・改行はスペースにする（1行に保つ）" {
        $path = "$TestDrive\進捗2.txt"
        writeConvertProgress ${convertPhaseScan} 0 0 0 "あ`tい`r`nう" $path
        (readConvertProgress $path).Detail | Should Be "あ い う"
    }

    It "ファイルが無い・壊れていれば null" {
        readConvertProgress "$TestDrive\進捗なし.txt" | Should BeNullOrEmpty
        $path = "$TestDrive\進捗3.txt"
        writeListFile $path @("変換`tあ`tい`tう`tえ")   # 件数が数値でない
        readConvertProgress $path | Should BeNullOrEmpty
        writeListFile $path @("変換`t1`t2")              # 列が足りない（書き込みの途中）
        readConvertProgress $path | Should BeNullOrEmpty
    }

    It "画面が読んでいる間も書ける（共有して開く）" {
        $path = "$TestDrive\進捗4.txt"
        writeConvertProgress ${convertPhaseRun} 1 2 0 "はじめ" $path
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        try {
            { writeConvertProgress ${convertPhaseRun} 2 1 0 "つぎ" $path } | Should Not Throw
        } finally {
            $stream.Dispose()
        }
        (readConvertProgress $path).Detail | Should Be "つぎ"
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\進捗5.txt"
        writeConvertProgress ${convertPhaseFinish} 0 0 0 "" $path
        removeConvertProgress $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeConvertProgress $path } | Should Not Throw
    }
}

Describe "writeConvertPlan / readConvertPlan / removeConvertPlan" -Tag Io {
    It "インデックスごとの件数を往復できる（件数は数値で返る）" {
        $path = "$TestDrive\予定1.tsv"
        $rows = @(
            (newConvertPlanRow "営業" "C:\data\営業" ${planKindConvert} 1234 12 5 7 0 0 3),
            (newConvertPlanRow "技術" "\\server\share\技術" ${planKindConvert} 20 0 0 0 0 0 0))
        writeConvertPlan $rows $path
        $plan = readConvertPlan $path
        $plan.Count | Should Be 2
        $plan[0].インデックス名 | Should Be "営業"
        $plan[0].元のフォルダ | Should Be "C:\data\営業"
        $plan[0].区分 | Should Be ${planKindConvert}
        ($plan[0].ファイル数 + 1) | Should Be 1235   # 文字列ではなく数値で返る
        $plan[0].変換対象 | Should Be 12
        $plan[0].新規 | Should Be 5
        $plan[0].更新あり | Should Be 7
        $plan[0].前回失敗 | Should Be 3
        $plan[1].変換対象 | Should Be 0
    }

    It "チェックなし・フォルダなしの区分も往復できる（件数は 0）" {
        $path = "$TestDrive\予定2.tsv"
        writeConvertPlan @(
            (newConvertPlanRow "外した" "D:\過去" ${planKindUnchecked}),
            (newConvertPlanRow "無い" "E:\USB" ${planKindMissing})) $path
        $plan = readConvertPlan $path
        $plan[0].区分 | Should Be ${planKindUnchecked}
        $plan[0].ファイル数 | Should Be 0
        $plan[1].区分 | Should Be ${planKindMissing}
    }

    It "インデックスが1件も無くても読める（空の配列）" {
        $path = "$TestDrive\予定3.tsv"
        writeConvertPlan @() $path
        (readConvertPlan $path).Count | Should Be 0
    }

    It "ファイルが無い・列が合わなければ null（画面は次の機会に読み直す）" {
        readConvertPlan "$TestDrive\予定なし.tsv" | Should BeNullOrEmpty
        $path = "$TestDrive\予定4.tsv"
        writeListFile $path @("べつの見出し")
        readConvertPlan $path | Should BeNullOrEmpty
    }

    It "画面が読んでいる間も書ける（共有して開く）" {
        $path = "$TestDrive\予定5.tsv"
        writeConvertPlan @((newConvertPlanRow "営業" "C:\data" ${planKindConvert} 1 1 1 0 0 0 0)) $path
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        try {
            { writeConvertPlan @((newConvertPlanRow "営業" "C:\data" ${planKindConvert} 2 2 2 0 0 0 0)) $path } | Should Not Throw
        } finally {
            $stream.Dispose()
        }
        (readConvertPlan $path)[0].変換対象 | Should Be 2
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\予定6.tsv"
        writeConvertPlan @() $path
        removeConvertPlan $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeConvertPlan $path } | Should Not Throw
    }
}

Describe "writeConvertStartRequest / readConvertStartRequest / removeConvertStartRequest" -Tag Io {
    It "前回失敗したファイルも再変換するかを伝えられる" {
        $path = "$TestDrive\開始要求1"
        writeConvertStartRequest $true $path
        (readConvertStartRequest $path).RetryFailed | Should Be $true
        writeConvertStartRequest $false $path
        (readConvertStartRequest $path).RetryFailed | Should Be $false
    }

    It "まだ返事が無ければ null（変換側は待ち続ける）" {
        readConvertStartRequest "$TestDrive\開始要求なし" | Should BeNullOrEmpty
    }

    It "削除できる（無ければ何もしない）" {
        $path = "$TestDrive\開始要求2"
        writeConvertStartRequest $false $path
        removeConvertStartRequest $path
        Test-Path -LiteralPath $path | Should Be $false
        { removeConvertStartRequest $path } | Should Not Throw
    }
}
