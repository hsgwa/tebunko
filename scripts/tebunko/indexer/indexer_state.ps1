# インデックス作成の状態（取り込み一覧）の読み書きと、画面とインデクサの受け渡しの口・ログ。

# 1ファイルの取り込みの制限時間（分）。超えたらOfficeアプリを強制終了し、そのファイルは失敗とする（クラウドからのダウンロードも含む。確認の文言も同じ値を使う）
${ingestFileTimeoutMinutes} = 10

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
        return "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは更新できません）"
    }

    $cause = $null
    if ($base -is [System.IO.FileNotFoundException] -or $base -is [System.IO.DirectoryNotFoundException]) {
        $cause = "ファイルが見つかりません（更新中に移動・削除・名前変更された可能性があります）"
    } elseif ($base -is [System.UnauthorizedAccessException] -or $code -eq "80070005") {
        $cause = "ファイルを読むアクセス権がありません"
    } elseif ($base -is [System.IO.PathTooLongException]) {
        $cause = "パスが長すぎるため読めません"
    } elseif ($base -is [System.OutOfMemoryException]) {
        # 巨大なシート（テキストにして約 1GB 超）は、整形（prettyTsv）で一度に読み込めずメモリ不足になる
        $cause = "シート・文書が大きすぎて更新できません（メモリが不足しました）"
    } elseif (@("80070020", "80070021") -contains $code) {
        # 共有違反・ロック違反
        $cause = "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再度更新してください）"
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

function addStatusRow {
    # 取り込み一覧の末尾に1行追記する（readStatusFile では後の行が優先される）
    param (
        $row,
        [string]$path = $workspace.StatusFile
    )

    [System.IO.File]::AppendAllText($path, "$(toStatusLine $row)`r`n", ${utf8Bom})
}

# ----------------------------------------------------------------------------
# 画面（または indexer.ps1）とインデクサの受け渡しの口（docs/design/structure/threads.md「画面とインデクサの受け渡し」）
# ----------------------------------------------------------------------------

function newIndexerChannel {
    # 受け渡しの口を作る。画面とインデクサのスレッドの両方から読み書きするため Synchronized にする。
    #   画面が書く      : RetryFailed・IncludeCloud（クラウドにだけあるファイルもダウンロードして取り込むか）・ConfirmTargets・Workers（取り込みのスレッドの数。0 は司令のスレッドで取り込む、-1 は設定・コア数から決める）・Stop・Answer・
    #                     OnlyNames（更新するインデックス名の配列。空なら、チェックの付いたものすべて）
    #   インデクサが書く: Progress（readIndexingProgress の形）・Plan（取り込み予定）・Error・ExitCode（0 完了 / 1 エラー / 2 中止）・
    #                     CloudSkipped（クラウドにだけあるため、ダウンロードせずに残した件数）・
    #                     OnlySkipped（OnlyNames のうち、更新できなかった名前の @{ Name; Reason } の配列）
    #                     StageTimings（段階ごとの時間と、取り込みの種類ごとの時間。IngestTiming.ToHashtable の @{ Stages; Kinds }。回の終わりに入る）
    #   OfficePids: インデックス作成が起動した Office の PID → @{ ProcessName; StartTicks }（閉じるときに止まらなければ、名前と起動時刻が合う PID だけを止める）
    #   UserOfficePids: 取り込み中に見える窓があった（利用者が使った）PowerPoint の PID → @{ ProcessName; StartTicks }（制限時間を過ぎても、閉じるときも止めない）
    param (
        [bool]$retryFailed = $false,
        [bool]$confirmTargets = $false,
        [int]$workers = -1,
        [string[]]$onlyNames = @(),
        [bool]$includeCloud = $false
    )

    return [hashtable]::Synchronized(@{
        RetryFailed = $retryFailed; IncludeCloud = $includeCloud; ConfirmTargets = $confirmTargets; Workers = $workers
        OnlyNames = @($onlyNames | Where-Object { $_ }); OnlySkipped = @()
        Progress = $null; Stop = $false
        Plan = $null; Answer = $null; Answered = New-Object System.Threading.ManualResetEvent($false)
        Error = ""; ExitCode = $null; CloudSkipped = 0; StageTimings = $null
        OfficePids = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,object]'
        UserOfficePids = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,object]'
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
    # 中止を求める。一覧化・取り込み中は、切れ目まで待たずにできたところまでで止まる（取り込み中の Office は止める）。確認を待っていれば、取りやめの返事にする
    param (
        $channel
    )

    $channel.Stop = $true
    [void]$channel.Answered.Set()
}

function answerIndexingPlan {
    # 確認の返事を返す。取り込む → @{ RetryFailed; IncludeCloud } ／ 取りやめ → $null
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

    $mutex = newAppMutex "indexer" $dir ${appId}
    try {
        return !$mutex.Acquired
    } finally {
        if ($mutex.Acquired) {
            $mutex.Mutex.ReleaseMutex()
        }
        $mutex.Mutex.Dispose()
    }
}

function writeZipSizeLimitLog {
    # .docx・.pptx・.xlsx の部品がサイズの上限を超えたときの詳細をログに書く（画面には出さない）。
    # 原因を調べられるよう、部品名・大きさ・部品ごとか合計かはここだけに書く。
    # 呼び出し元（extract_office.ps1・indexer_run.ps1・office_reader.ps1 のグラフ・SmartArt 読み取り）は、
    # 捕まえた例外をそのまま渡せばよい。ZipSizeLimitException かどうかはここで確かめ、違えば何もしない
    # （param の型を [ZipSizeLimitException] にしないのは、indexer_state.ps1 が office_reader.ps1 より先に
    # 読み込まれ、その時点ではクラスがまだ無いため）
    param (
        [System.Exception]$exception
    )

    if ($exception -isnot [ZipSizeLimitException]) {
        return
    }

    $kindText = switch ($exception.LimitKind) { "Total" { "1ファイルの合計" } "Output" { "埋め込みから出す文字の合計" } default { "部品ごと" } }
    # Output の MeasuredBytes は、バイトでなく出す行の文字数（単位を「文字」にする）
    $mb = [math]::Round($exception.MeasuredBytes / 1MB, 1)
    $sizeText = $(if ($exception.LimitKind -eq "Output") { "$($exception.MeasuredBytes) 文字、約${mb}M文字" } else { "$($exception.MeasuredBytes) バイト、約${mb}MB" })
    writeIndexerLog "    サイズの上限（${kindText}）を超えました: $($exception.PartName)（${sizeText}）" "Yellow"
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
        [int]$failed = 0,   # 前回失敗し、その後更新されていない（再取り込みするかは画面で選ぶ）
        [int]$cloud = 0,        # クラウドにだけあるため、取り込み対象に数えていないファイル（新規・更新あり・前回未完了など）
        [int]$cloudFailed = 0,  # クラウドにだけあり、前回失敗したファイル
        [long]$cloudBytes = 0,  # クラウドにだけあり、取り込み対象のファイルの合計サイズ（バイト）
        [long]$cloudFailedBytes = 0  # クラウドにだけあり、前回失敗したファイルの合計サイズ（バイト）
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
        クラウド       = $cloud
        クラウド失敗   = $cloudFailed
        クラウド容量   = $cloudBytes
        クラウド失敗容量 = $cloudFailedBytes
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
    #      FailedRows（失敗したファイルの行。取り込み日時の新しい順）; IndexStats（インデックス名ごとの集計。getIndexStats）;
    #      Format（格納形式の状態。getWorkspaceIndexFormat の結果。画面の裏の仕事が入れる。ここでは $null） }
    param (
        [datetime]$since = [datetime]::MaxValue,
        [string]$path = $workspace.StatusFile
    )

    $state = @{ Exists = $false; Folders = @(); Total = 0; Pending = 0; Failed = 0; Done = 0; IngestedSince = 0; Updated = $null; FailedRows = @(); IndexStats = (getIndexStats $null); Format = $null }
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
        if (testStatusRowFailed $row) {
            $state.Failed++
            $failedRows.Add($row)
        } elseif ($row.状態 -eq ${stateNew}) {
            $state.Pending++
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

# 取り込み一覧の読み書きと、今回の取り込みでの失敗・元ファイルが無くなった行の記録をまとめる（StatusLedger）。
# 列と状態の定義（core/paths.ps1 の ${statusColumns} など）は判断層（indexer_decide.ps1）・index_store.ps1 なども読むため、
# 今の場所（paths.ps1 の変数）のまま動かさない（docs/design/structure/classes.md）。
# 保存先（取り込み一覧）は、コンストラクタで受け取った $workspace（Workspace のインスタンス。
# 場所を暗黙に使わないため、メソッドの中で ${workspace} を読まない）の StatusFile から決める
class StatusLedger {
    hidden $Workspace
    # 今回の取り込みで失敗したファイル（@{ RelPath; Message } の並び）
    [System.Collections.Generic.List[object]]$Failures
    # 取り込みの直前に元のファイルが無くなっていた相対パス（取り込み一覧から除く）
    [System.Collections.Generic.HashSet[string]]$DroppedRows

    StatusLedger($workspace) {
        $this.Workspace = $workspace
        $this.Failures = New-Object System.Collections.Generic.List[object]
        $this.DroppedRows = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }

    [hashtable] ReadStatus() {
        # 取り込み一覧を読み込む（readStatusFile と同じ形。@{ Folders; Rows; Rejected; LegacyFolders }）
        return readStatusFile $this.Workspace.StatusFile
    }

    [void] WriteStatus($folders, $rows) {
        # 取り込み一覧を書き出す（1ファイル1行）
        writeStatusFile $folders $rows $this.Workspace.StatusFile
    }

    [void] AddRow($row) {
        # 取り込み一覧の末尾に1行追記する
        addStatusRow $row $this.Workspace.StatusFile
    }

    [string[]] ReadRelPaths() {
        # 取り込み一覧のデータ行の相対パスを、ファイルの順に返す（同じ相対パスの追記も数える。ジャーナルの「この回の行」の数え方）
        return [string[]]@(readStatusRelPaths $this.Workspace.StatusFile)
    }

    [void] RenameIndexName([string]$oldName, [string]$newName) {
        # 取り込み一覧に記録したインデックス名を書き換える（クロール対象フォルダの行と、各行の相対パスの先頭）
        renameStatusIndexName $oldName $newName $this.Workspace.StatusFile
    }

    [void] RemoveIndexName([string]$name) {
        # 取り込み一覧から、あるインデックスの記録（クロール対象フォルダの行と、そのインデックスの各行）を取り除く
        removeStatusIndexName $name $this.Workspace.StatusFile
    }

    [void] AddFailure([string]$relPath, [string]$message) {
        # 今回の取り込みで失敗したファイルを記録する
        $this.Failures.Add(@{ RelPath = $relPath; Message = $message })
    }

    [void] AddDropped([string]$relPath) {
        # 取り込みの直前に元のファイルが無くなっていた（取り込み一覧・インデックスから除く）ファイルを記録する
        [void]$this.DroppedRows.Add($relPath)
    }
}

function readPendingRemovals {
    # 本文インデックスから外し終えていない、無くなった元のファイルの記録を読み、@{ Folder; Name } の配列を返す。無ければ空。
    # 置き場の外を指す・使えない相対パスの行は読み飛ばす（testIndexRelPath）
    param (
        [string]$path = $workspace.PendingRemoveFile
    )

    $checkedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    $result = New-Object System.Collections.Generic.List[hashtable]
    foreach ($entry in @(convertFromPendingRemoveLines @(readListFile $path))) {
        if (testIndexRelPath "$($entry.Folder)\$($entry.Name)" $false $checkedNames) {
            continue
        }
        $result.Add($entry)
    }
    return $result.ToArray()
}

function writePendingRemovals {
    # 外し終えていない、無くなった元のファイルの記録を書く（一時ファイルに書いてから置き換える）。0 件なら記録を消す
    param (
        [object[]]$entries,
        [string]$path = $workspace.PendingRemoveFile
    )

    $lines = @(convertToPendingRemoveLines $entries)
    if ($lines.Count -eq 0) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
        }
        return
    }
    writeTextLinesAtomic $path $lines
}