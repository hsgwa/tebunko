# インデックス作成の状態（取り込み一覧・取り込み中のファイル）の読み書きと、画面とインデクサの受け渡しの口・ログ。

function newStatusRow {
    # 取り込み一覧の1行（1ファイル）を作る
    param (
        [string]$relPath,
        [string]$updated = "",
        [string]$size = "",
        [string]$state = ${stateNew},
        [string]$tsvCount = "",
        [string]$ingested = "",
        [string]$errorMessage = "",
        [string]$extractVersion = ""
    )

    return [pscustomobject]@{
        相対パス = $relPath
        更新日時 = $updated
        サイズ   = $size
        状態     = $state
        TSV数    = $tsvCount
        取り込み日時 = $ingested
        エラー   = $errorMessage
        抽出版   = $extractVersion
    }
}


# 1行1件を保つため、タブ・改行はスペースにする（この文字を含む値だけ置き換える）
${statusLineBreaks} = [char[]]@("`t", "`r", "`n")


function toStatusLine {
    # 取り込み一覧の1行を文字列にする（1件ずつの追記用。addStatusRow）。
    # エラーメッセージ等のタブ・改行は、1行1件を保つためスペースにする。
    # 数万行をまとめて書き出す writeStatusFile では、同じ整形をその場に展開している（関数の呼び出しだけで
    # 5万行あたり数秒かかるため）。整形を変えるときは両方を直す（テストで同じ結果になることを確かめている）
    param (
        $row
    )

    $errorText = [string]$row.エラー
    if ($errorText.IndexOfAny(${statusLineBreaks}) -ge 0) {
        $errorText = $errorText -replace "[\t\r\n]+", " "
    }
    return [string]::Join("`t", @(
        [string]$row.相対パス, [string]$row.更新日時, [string]$row.サイズ, [string]$row.状態,
        [string]$row.TSV数, [string]$row.取り込み日時, $errorText, [string]$row.抽出版))
}

function describeIngestError {
    # 取り込みで発生した例外から、取り込み一覧のエラー列・画面に表示する失敗の理由を返す。
    # Officeアプリのメッセージは分かりにくい（パスワード付きでも「入力したパスワードが間違っています」等）ため、
    # よくある原因は「原因（詳細: 元のメッセージ）」の形に言い換える
    param (
        [System.Exception]$exception
    )

    # メソッド呼び出しの例外（「"7" 個の引数を指定して "Open" を呼び出し中に例外が発生しました」等）は、中の例外が本来の理由
    $base = $exception.GetBaseException()
    $message = ($base.Message -replace "\s+", " ").Trim()
    $code = "{0:X8}" -f $base.HResult
    if ($message -eq "") {
        $message = "エラーコード 0x${code}"
    }

    # スクリプト自身が throw したメッセージは、利用者向けに書いてあるためそのまま使う
    if ($base.GetType() -eq [System.Management.Automation.RuntimeException]) {
        return $message
    }

    if ($message -match "パスワード|password") {
        # 元のメッセージ（パスワードが間違っています等）は、パスワードを入力していない利用者には誤解を招くため付けない
        return "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは取り込めません）"
    }

    $cause = $null
    if ($base -is [System.IO.FileNotFoundException] -or $base -is [System.IO.DirectoryNotFoundException]) {
        $cause = "ファイルが見つかりません（取り込み中に移動・削除・名前変更された可能性があります）"
    } elseif ($base -is [System.UnauthorizedAccessException] -or $code -eq "80070005") {
        $cause = "ファイルを読むアクセス権がありません"
    } elseif ($base -is [System.IO.PathTooLongException]) {
        $cause = "パスが長すぎるため読めません"
    } elseif ($base -is [System.OutOfMemoryException]) {
        # 巨大なシート（テキストにして約 1GB 超）は、整形（prettyTsv）で一度に読み込めずメモリ不足になる
        $cause = "シート・文書が大きすぎて取り込めません（メモリが不足しました）"
    } elseif (@("80070020", "80070021") -contains $code) {
        # 共有違反・ロック違反
        $cause = "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再取り込みしてください）"
    } elseif (@("80040154", "80080005", "800401F3") -contains $code) {
        # クラス未登録・サーバーの起動失敗・ProgID 不正
        $cause = "Officeアプリ（Excel・Word・PowerPoint）を起動できませんでした（インストール・ライセンス認証の状態を確認してください）"
    } elseif (@("800706BA", "800706BE", "80010105", "80010108") -contains $code) {
        # RPC サーバーを利用できない・呼び出し失敗・サーバーで例外・切断
        $cause = "Officeアプリが異常終了したか、内部でエラーが発生しました（ファイルが壊れている、または大きすぎる可能性があります）"
    } elseif (@("80010001", "8001010A") -contains $code) {
        # 呼び出しの拒否・処理中のため後で再試行
        $cause = "Officeアプリが応答しませんでした（Officeアプリでダイアログを表示中などの可能性があります）"
    } elseif ($base -is [System.IO.InvalidDataException] -or $base -is [System.Xml.XmlException] -or
              $message -match "ファイル形式|破損|壊れ|corrupt|file format") {
        $cause = "ファイルが壊れているか、拡張子と中身の形式が一致していません"
    }

    if ($cause) {
        return "${cause}（詳細: ${message}）"
    }
    return $message
}

function readStatusFile {
    # 取り込み一覧を読み込み、@{ Folders; Rows } を返す。ファイルが無ければ空。
    #   Folders: クロール対象フォルダ @{ Path; Name（インデックス名） } の配列
    #   Rows   : 相対パス（"インデックス名\フォルダからの相対パス"。大文字・小文字を区別しない）→ 行
    # インデックス作成中は1件ごとに行を追記するため、同じ相対パスの行は後の行を優先する。列数の合わない行（書き込み途中で中断した行など）は無視する
    param (
        [string]$path = $workspace.StatusFile
    )

    $rows = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $folders = New-Object System.Collections.Generic.List[object]
    $result = @{ Folders = $folders; Rows = $rows }
    if (!(Test-Path -LiteralPath $path)) {
        return $result
    }

    # インデックス作成中に画面などから読んでも、インデクサの追記・置き換えを妨げないよう共有を許して開く
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        $lines = $reader.ReadToEnd() -split "\r?\n"
    } finally {
        $reader.Dispose()
    }

    $columnCount = ${statusColumns}.Count
    foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3) {
            $folders.Add([pscustomobject]@{ Path = $fields[1]; Name = $fields[2] })
            continue
        }
        if ($fields.Count -ne $columnCount -or $fields[0] -eq "" -or $fields[0] -eq ${statusColumns}[0]) {
            continue
        }
        # 数万行を読むため、1行ごとの関数呼び出し（newStatusRow）は使わずにその場で作る（列は $statusColumns と同じ）
        $rows[$fields[0]] = [pscustomobject]@{
            相対パス = $fields[0]
            更新日時 = $fields[1]
            サイズ   = $fields[2]
            状態     = $fields[3]
            TSV数    = $fields[4]
            取り込み日時 = $fields[5]
            エラー   = $fields[6]
            抽出版   = $fields[7]
        }
    }
    return $result
}

function writeStatusFile {
    # 取り込み一覧を書き出す（1ファイル1行）。途中で中断しても壊れないよう、一時ファイルに書いてから置き換える
    #   folders: クロール対象フォルダ @{ Path; Name } の配列。"クロール対象フォルダ<TAB>パス<TAB>インデックス名" の行にする
    param (
        [object[]]$folders,
        [object[]]$rows,
        [string]$path = $workspace.StatusFile
    )

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($folder in @($folders | Where-Object { $_ })) {
        $lines.Add("${statusFolderKey}`t$($folder.Path)`t$($folder.Name)")
    }
    $lines.Add((${statusColumns} -join "`t"))
    # 数万行を書き出すため、1行ごとの関数呼び出し（toStatusLine）・パイプライン（Where-Object）は使わない
    # （5万行で約 10 秒 → 約 0.2 秒。この間はインデックス作成の進み具合が止まって見えるため速くする）
    foreach ($row in $rows) {
        if ($null -eq $row) {
            continue
        }
        $errorText = [string]$row.エラー
        if ($errorText.IndexOfAny(${statusLineBreaks}) -ge 0) {
            $errorText = $errorText -replace "[\t\r\n]+", " "
        }
        $lines.Add([string]::Join("`t", @(
            [string]$row.相対パス, [string]$row.更新日時, [string]$row.サイズ, [string]$row.状態,
            [string]$row.TSV数, [string]$row.取り込み日時, $errorText, [string]$row.抽出版)))
    }

    writeTextLinesAtomic $path $lines
}

function addStatusRow {
    # 取り込み一覧の末尾に1行追記する（readStatusFile では後の行が優先される）
    param (
        $row,
        [string]$path = $workspace.StatusFile
    )

    [System.IO.File]::AppendAllText($path, "$(toStatusLine $row)`r`n", ${utf8Bom})
}

function readIngestingFiles {
    # 取り込み中のファイルの記録を読み、@{ RelPath = 相対パス; Count = 続けて取り込みを始めて終わらなかった回数 } の配列を返す。
    # 取り込みを複数のスレッドで行うため、1 行に 1 ファイル。記録が無ければ空。壊れた行は読み飛ばす
    param (
        [string]$path = $workspace.IngestingFile
    )

    $result = New-Object System.Collections.Generic.List[hashtable]
    foreach ($line in @(readListFile $path)) {
        $fields = $line.Split("`t")
        $count = 0
        if ($fields.Count -ne 2 -or -not [int]::TryParse($fields[0], [ref]$count) -or $count -lt 1 -or $fields[1] -eq "") {
            continue
        }
        $result.Add(@{ RelPath = $fields[1]; Count = $count })
    }
    return , $result.ToArray()
}

function writeIngestingFiles {
    # 取り込み中のファイル（@{ RelPath; Count } の並び）を "回数<TAB>相対パス" で記録する。無ければ記録を消す
    param (
        [object[]]$entries,
        [string]$path = $workspace.IngestingFile
    )

    $lines = @($entries | Where-Object { $_ } | ForEach-Object { "$($_.Count)`t$($_.RelPath)" })
    if ($lines.Count -eq 0) {
        removeIngestingFile $path
        return
    }
    writeListFile $path $lines
}

function removeIngestingFile {
    param (
        [string]$path = $workspace.IngestingFile
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

# ----------------------------------------------------------------------------
# 画面（または indexer.ps1）とインデクサの受け渡しの口（docs/00_共通_4_プロセスとスレッド.md 7.5）
# ----------------------------------------------------------------------------

function newIndexerChannel {
    # 受け渡しの口を作る。画面とインデクサのスレッドの両方から読み書きするため Synchronized にする。
    #   画面が書く      : RetryFailed・ConfirmTargets・Workers（取り込みのスレッドの数。0 は司令のスレッドで取り込む、-1 は設定・コア数から決める）・Stop・Answer
    #   インデクサが書く: Progress（readIndexingProgress の形）・Plan（取り込み予定）・Error・ExitCode（0 完了 / 1 エラー / 2 中止）
    #   OfficePids: インデックス作成が起動した Office の PID → プロセス名（閉じるときに止まらなければ、この PID だけを止める）
    param (
        [bool]$retryFailed = $false,
        [bool]$confirmTargets = $false,
        [int]$workers = -1
    )

    return [hashtable]::Synchronized(@{
        RetryFailed = $retryFailed; ConfirmTargets = $confirmTargets; Workers = $workers
        Progress = $null; Stop = $false
        Plan = $null; Answer = $null; Answered = New-Object System.Threading.ManualResetEvent($false)
        Error = ""; ExitCode = $null
        OfficePids = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,string]'
    })
}

# いま動いているインデックス作成の受け渡しの口（invokeIndexer が入れる。$null なら進み具合を伝えない）
$script:indexerChannel = $null

function writeIndexingProgress {
    # インデックス作成の進み具合を受け渡しの口に入れる（画面が 1 秒ごとに読む）
    param (
        [string]$phase,
        [int]$processed = 0,
        [int]$remaining = 0,
        [int]$failed = 0,
        [string]$detail = "",
        $channel = $script:indexerChannel
    )

    if ($null -ne $channel) {
        $channel.Progress = @{ Phase = $phase; Processed = $processed; Remaining = $remaining; Failed = $failed; Detail = ($detail -replace "[\t\r\n]+", " ") }
    }
}

function readIndexingProgress {
    # インデックス作成の進み具合（@{ Phase; Processed; Remaining; Failed; Detail }）を返す。まだ無ければ $null
    param (
        $channel
    )

    if ($null -eq $channel) {
        return $null
    }
    return $channel.Progress
}

function requestIndexingStop {
    # 中止を求める。取り込みはファイルの切れ目で止まる。確認を待っていれば、取りやめの返事にする
    param (
        $channel
    )

    $channel.Stop = $true
    [void]$channel.Answered.Set()
}

function answerIndexingPlan {
    # 確認の返事を返す。取り込む → @{ RetryFailed } ／ 取りやめ → $null
    param (
        $channel,
        $answer
    )

    $channel.Answer = $answer
    if ($null -eq $answer) {
        $channel.Stop = $true
    }
    [void]$channel.Answered.Set()
}

function testIndexerRunning {
    # この work でインデックス作成が動いているか（画面のスレッド・画面を使わない indexer.ps1 のどちらでも）。
    # インデックス作成が持つ鍵（newAppMutex "indexer"）を取れるかで調べ、取れたらすぐ放す
    param (
        [string]$dir = $workspace.Dir
    )

    $mutex = newAppMutex "indexer" $dir
    try {
        return !$mutex.Acquired
    } finally {
        if ($mutex.Acquired) {
            $mutex.Mutex.ReleaseMutex()
        }
        $mutex.Mutex.Dispose()
    }
}

# インデックス作成のログ（invokeIndexer が開く TextWriter。取り込みのスレッドでは 1 ファイル分を貯める StringWriter）
$script:indexerLog = $null
# ログをコンソールにも出すか（画面を使わずに indexer.ps1 を実行したとき）
$script:indexerEcho = $false

function writeIndexerLog {
    # インデックス作成の表示内容をログ（インデックス作成ログ.txt）に書く。color はコンソールに出すときの色
    param (
        [string]$text = "",
        [string]$color = ""
    )

    if ($null -ne $script:indexerLog) {
        try {
            $script:indexerLog.WriteLine($text)
        } catch {
            # ログのためだけの書き込みのため、書けなくてもインデックス作成は続ける
        }
    }
    if ($script:indexerEcho) {
        if ($color) {
            Write-Host $text -ForegroundColor $color
        } else {
            Write-Host $text
        }
    }
}

function newIngestPlanRow {
    # 取り込み予定（インデックス1件分）の行を作る
    param (
        [string]$name,
        [string]$path,
        [string]$kind = ${planKindIngest},
        [int]$total = 0,    # 見つかった Office ファイルの数
        [int]$targets = 0,  # 今回取り込むファイルの数（新規＋更新あり＋前回未完了＋インデックスなし）
        [int]$new = 0,
        [int]$updated = 0,
        [int]$pending = 0,
        [int]$lost = 0,     # インデックス（TSV）が無くなった・壊れているため取り込み直す
        [int]$failed = 0    # 前回失敗し、その後更新されていない（再取り込みするかは画面で選ぶ）
    )

    return [pscustomobject]@{
        インデックス名 = $name
        元のフォルダ   = $path
        区分           = $kind
        ファイル数     = $total
        取り込み対象       = $targets
        新規           = $new
        更新あり       = $updated
        前回未完了     = $pending
        インデックスなし   = $lost
        前回失敗       = $failed
    }
}

function readStatusLines {
    # 取り込み一覧を1行ずつ読む（インデックス作成中でも読めるよう共有を許して開く）。ファイルが無ければ空
    param (
        [string]$path = $workspace.StatusFile
    )

    if (!(Test-Path -LiteralPath $path)) {
        return @()
    }
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        $text = $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
    return @($text -split "\r?\n" | Where-Object { $_ -ne "" })
}

function renameStatusIndexName {
    # 取り込み一覧に記録したインデックス名を書き換える（クロール対象フォルダの行と、各行の相対パスの先頭）。
    # 行の順序と内容をそのまま保つため、行ごとに書き換えて置き換える
    param (
        [string]$oldName,
        [string]$newName,
        [string]$path = $workspace.StatusFile
    )

    $lines = @(readStatusLines $path)
    if ($lines.Count -eq 0) {
        return
    }

    $result = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3 -and [string]::Equals($fields[2], $oldName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $fields[2] = $newName
            $result.Add($fields -join "`t")
            continue
        }
        if ($fields.Count -eq ${statusColumns}.Count -and $fields[0] -ne "") {
            $split = splitIndexRelPath $fields[0]
            if ($split.Rest -ne "" -and [string]::Equals($split.Name, $oldName, [System.StringComparison]::OrdinalIgnoreCase)) {
                $fields[0] = "${newName}\$($split.Rest)"
                $result.Add($fields -join "`t")
                continue
            }
        }
        $result.Add($line)
    }
    writeTextLinesAtomic $path $result
}

function removeStatusIndexName {
    # 取り込み一覧から、あるインデックスの記録（クロール対象フォルダの行と、そのインデックスの各行）を取り除く
    param (
        [string]$name,
        [string]$path = $workspace.StatusFile
    )

    $lines = @(readStatusLines $path)
    if ($lines.Count -eq 0) {
        return
    }

    $result = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3 -and [string]::Equals($fields[2], $name, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        if ($fields.Count -eq ${statusColumns}.Count -and $fields[0] -ne "") {
            $split = splitIndexRelPath $fields[0]
            if ($split.Rest -ne "" -and [string]::Equals($split.Name, $name, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
        }
        $result.Add($line)
    }
    writeTextLinesAtomic $path $result
}

function getIndexingState {
    # 取り込み一覧からインデックス作成の状態を返す:
    #   @{ Exists; Folders（クロール対象フォルダ @{ Path; Name } の配列）; Total; Pending（未取り込み）; Failed; Done; IngestedSince（since 以降に取り込んだ件数）; Updated（取り込み一覧の更新日時）;
    #      FailedRows（失敗したファイルの行。取り込み日時の新しい順）; IndexStats（インデックス名ごとの集計。getIndexStats） }
    param (
        [datetime]$since = [datetime]::MaxValue,
        [string]$path = $workspace.StatusFile
    )

    $state = @{ Exists = $false; Folders = @(); Total = 0; Pending = 0; Failed = 0; Done = 0; IngestedSince = 0; Updated = $null; FailedRows = @(); IndexStats = (getIndexStats $null) }
    if (!(Test-Path -LiteralPath $path)) {
        return $state
    }

    $state.Exists = $true
    $state.Updated = (Get-Item -LiteralPath $path).LastWriteTime
    $status = readStatusFile $path
    $state.Folders = $status.Folders.ToArray()
    $state.Total = $status.Rows.Count
    $state.IndexStats = getIndexStats $status.Rows
    # since を指定しないとき（画面の集計）は、行ごとの日時の解析（数万行では数秒かかる）を省く
    $countSince = ($since -ne [datetime]::MaxValue)
    $sinceSecond = if ($countSince) { $since.AddTicks(-($since.Ticks % [timespan]::TicksPerSecond)) } else { $since }
    $failedRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $status.Rows.Values) {
        if ($row.状態 -eq ${stateNew}) {
            $state.Pending++
        } elseif ($row.状態 -eq ${stateFailed}) {
            $state.Failed++
            $failedRows.Add($row)
        } elseif ($row.状態 -eq ${stateDone}) {
            $state.Done++
        }

        $ingested = [datetime]::MinValue
        if ($countSince -and $row.取り込み日時 -and [datetime]::TryParseExact($row.取り込み日時, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ingested) -and $ingested -ge $sinceSecond) {
            $state.IngestedSince++
        }
    }
    # 取り込み日時は "yyyy/MM/dd HH:mm:ss" のため、文字列の順で新しい順に並ぶ
    $state.FailedRows = @($failedRows | Sort-Object -Property @{ Expression = { [string]$_.取り込み日時 }; Descending = $true }, @{ Expression = { $_.相対パス } })
    return $state
}
