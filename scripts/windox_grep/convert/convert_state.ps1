# 変換の状態を表すファイル（変換一覧・変換中・進捗・予定・開始要求）の読み書き。

function newStatusRow {
    # 変換一覧の1行（1ファイル）を作る
    param (
        [string]$relPath,
        [string]$updated = "",
        [string]$size = "",
        [string]$state = ${stateNew},
        [string]$tsvCount = "",
        [string]$converted = "",
        [string]$errorMessage = ""
    )

    return [pscustomobject]@{
        相対パス = $relPath
        更新日時 = $updated
        サイズ   = $size
        状態     = $state
        TSV数    = $tsvCount
        変換日時 = $converted
        エラー   = $errorMessage
    }
}


# 1行1件を保つため、タブ・改行はスペースにする（この文字を含む値だけ置き換える）
${statusLineBreaks} = [char[]]@("`t", "`r", "`n")


function toStatusLine {
    # 変換一覧の1行を文字列にする（1件ずつの追記用。addStatusRow）。
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
        [string]$row.TSV数, [string]$row.変換日時, $errorText))
}

function describeConvertError {
    # 変換で発生した例外から、変換一覧のエラー列・画面に表示する失敗の理由を返す。
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
        return "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは変換できません）"
    }

    $cause = $null
    if ($base -is [System.IO.FileNotFoundException] -or $base -is [System.IO.DirectoryNotFoundException]) {
        $cause = "ファイルが見つかりません（変換中に移動・削除・名前変更された可能性があります）"
    } elseif ($base -is [System.UnauthorizedAccessException] -or $code -eq "80070005") {
        $cause = "ファイルを読むアクセス権がありません"
    } elseif ($base -is [System.IO.PathTooLongException]) {
        $cause = "パスが長すぎるため読めません"
    } elseif ($base -is [System.OutOfMemoryException]) {
        # 巨大なシート（テキストにして約 1GB 超）は、整形（prettyTsv）で一度に読み込めずメモリ不足になる
        $cause = "シート・文書が大きすぎて変換できません（メモリが不足しました）"
    } elseif (@("80070020", "80070021") -contains $code) {
        # 共有違反・ロック違反
        $cause = "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再変換してください）"
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
    # 変換一覧を読み込み、@{ Folders; Rows } を返す。ファイルが無ければ空。
    #   Folders: 変換対象フォルダ @{ Path; Name（インデックス名） } の配列。以前の形式（インデックス名なし）は Name が空
    #   Rows   : 相対パス（"インデックス名\フォルダからの相対パス"。大文字・小文字を区別しない）→ 行
    # 変換中は1件ごとに行を追記するため、同じ相対パスの行は後の行を優先する。列数の合わない行（書き込み途中で中断した行など）は無視する
    param (
        [string]$path = ${statusFile}
    )

    $rows = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $folders = New-Object System.Collections.Generic.List[object]
    $result = @{ Folders = $folders; Rows = $rows }
    if (!(Test-Path -LiteralPath $path)) {
        return $result
    }

    # 変換中に画面などから読んでも、変換側の追記・置き換えを妨げないよう共有を許して開く
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        $lines = $reader.ReadToEnd() -split "\r?\n"
    } finally {
        $reader.Dispose()
    }

    $header = ${statusColumns} -join "`t"
    foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and ($fields.Count -eq 2 -or $fields.Count -eq 3)) {
            $folders.Add([pscustomobject]@{ Path = $fields[1]; Name = $(if ($fields.Count -eq 3) { $fields[2] } else { "" }) })
            continue
        }
        if ($line -eq $header -or $fields.Count -ne ${statusColumns}.Count -or $fields[0] -eq "") {
            continue
        }
        # 数万行を読むため、1行ごとの関数呼び出し（newStatusRow）は使わずにその場で作る（列は $statusColumns と同じ）
        $rows[$fields[0]] = [pscustomobject]@{
            相対パス = $fields[0]
            更新日時 = $fields[1]
            サイズ   = $fields[2]
            状態     = $fields[3]
            TSV数    = $fields[4]
            変換日時 = $fields[5]
            エラー   = $fields[6]
        }
    }
    return $result
}

function writeStatusFile {
    # 変換一覧を書き出す（1ファイル1行）。途中で中断しても壊れないよう、一時ファイルに書いてから置き換える
    #   folders: 変換対象フォルダ @{ Path; Name } の配列。"変換対象フォルダ<TAB>パス<TAB>インデックス名" の行にする
    param (
        [object[]]$folders,
        [object[]]$rows,
        [string]$path = ${statusFile}
    )

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($folder in @($folders | Where-Object { $_ })) {
        $lines.Add("${statusFolderKey}`t$($folder.Path)`t$($folder.Name)")
    }
    $lines.Add((${statusColumns} -join "`t"))
    # 数万行を書き出すため、1行ごとの関数呼び出し（toStatusLine）・パイプライン（Where-Object）は使わない
    # （5万行で約 10 秒 → 約 0.2 秒。この間は変換の進み具合が止まって見えるため速くする）
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
            [string]$row.TSV数, [string]$row.変換日時, $errorText)))
    }

    writeTextLinesAtomic $path $lines
}

function addStatusRow {
    # 変換一覧の末尾に1行追記する（readStatusFile では後の行が優先される）
    param (
        $row,
        [string]$path = ${statusFile}
    )

    [System.IO.File]::AppendAllText($path, "$(toStatusLine $row)`r`n", ${utf8Bom})
}

function readConvertingFile {
    # 変換中のファイルの記録を読み、@{ RelPath = 相対パス; Count = 続けて変換を始めて終わらなかった回数 } を返す。
    # 記録が無い・壊れている場合は $null
    param (
        [string]$path = ${convertingFile}
    )

    $lines = @(readListFile $path)
    if ($lines.Count -eq 0) {
        return $null
    }
    $fields = $lines[0].Split("`t")
    $count = 0
    if ($fields.Count -ne 2 -or -not [int]::TryParse($fields[0], [ref]$count) -or $count -lt 1 -or $fields[1] -eq "") {
        return $null
    }
    return @{ RelPath = $fields[1]; Count = $count }
}

function writeConvertingFile {
    # 変換を始めるファイルを "回数<TAB>相対パス" で記録する。変換が終われば removeConvertingFile で消す
    param (
        [string]$relPath,
        [int]$count,
        [string]$path = ${convertingFile}
    )

    writeListFile $path @("${count}`t${relPath}")
}

function removeConvertingFile {
    param (
        [string]$path = ${convertingFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function writeConvertProgress {
    # 変換の進み具合を1行で書く（画面が読む）。書き込みは1ファイルにつき1回で、変換の速さに影響しない大きさにする。
    #   "<段階><TAB><処理済み><TAB><残り><TAB><失敗><TAB><いま行っていること>"
    # 画面が読んでいる最中でも書けるよう、共有を許して開く
    param (
        [string]$phase,
        [int]$processed = 0,
        [int]$remaining = 0,
        [int]$failed = 0,
        [string]$detail = "",
        [string]$path = ${convertProgressFile}
    )

    $line = "{0}`t{1}`t{2}`t{3}`t{4}" -f $phase, $processed, $remaining, $failed, ($detail -replace "[\t\r\n]+", " ")
    try {
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        try {
            $bytes = ${utf8Bom}.GetPreamble() + ${utf8Bom}.GetBytes($line)
            $stream.Write($bytes, 0, $bytes.Length)
        } finally {
            $stream.Dispose()
        }
    } catch {
        # 進み具合の表示のためだけのファイルのため、書けなくても変換は続ける
    }
}

function readConvertProgress {
    # 変換の進み具合を読む（無い・壊れていれば $null）。画面が毎秒呼ぶため、1行だけ読む
    param (
        [string]$path = ${convertProgressFile}
    )

    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    # 変換側が書いている最中でも読めるよう、共有を許して開く
    $text = ""
    try {
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
        try {
            $text = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }
    } catch {
        return $null  # 書き込みと重なった等。次の機会に読む
    }

    $fields = (($text -split "\r?\n")[0]).Split("`t")
    if ($fields.Count -lt 5) {
        return $null  # 書き込みの途中
    }
    $numbers = @(0, 0, 0)
    for ($i = 0; $i -lt 3; $i++) {
        $value = 0
        if (-not [int]::TryParse($fields[$i + 1], [ref]$value)) {
            return $null
        }
        $numbers[$i] = $value
    }
    return @{ Phase = $fields[0]; Processed = $numbers[0]; Remaining = $numbers[1]; Failed = $numbers[2]; Detail = $fields[4] }
}

function removeConvertProgress {
    param (
        [string]$path = ${convertProgressFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function newConvertPlanRow {
    # 変換予定（インデックス1件分）の行を作る
    param (
        [string]$name,
        [string]$path,
        [string]$kind = ${planKindConvert},
        [int]$total = 0,    # 見つかった Office ファイルの数
        [int]$targets = 0,  # 今回変換するファイルの数（新規＋更新あり＋前回未完了＋変換結果なし）
        [int]$new = 0,
        [int]$updated = 0,
        [int]$pending = 0,
        [int]$lost = 0,     # 変換結果（TSV）が無くなった・壊れているため変換し直す
        [int]$failed = 0    # 前回失敗し、その後更新されていない（再変換するかは画面で選ぶ）
    )

    return [pscustomobject]@{
        インデックス名 = $name
        元のフォルダ   = $path
        区分           = $kind
        ファイル数     = $total
        変換対象       = $targets
        新規           = $new
        更新あり       = $updated
        前回未完了     = $pending
        変換結果なし   = $lost
        前回失敗       = $failed
    }
}

function writeConvertPlan {
    # 変換予定を書き出す（インデックス1件1行）。画面は数える前から読むため、
    # 途中の状態を読ませないよう一時ファイルに書いてから置き換える
    param (
        [object[]]$rows,
        [string]$path = ${convertPlanFile}
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add((${convertPlanColumns} -join "`t"))
    foreach ($row in @($rows | Where-Object { $_ })) {
        $values = foreach ($column in ${convertPlanColumns}) { ([string]$row.$column) -replace "[\t\r\n]+", " " }
        $lines.Add([string]::Join("`t", @($values)))
    }
    writeTextLinesAtomic $path $lines
}

function readConvertPlan {
    # 変換予定を読む。ファイルが無い・列が合わない場合は $null（画面は次の機会に読み直す）。
    # 件数の列は数値にして返す
    param (
        [string]$path = ${convertPlanFile}
    )

    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    # 変換側が置き換えている最中でも読めるよう、共有を許して開く
    $text = ""
    try {
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
        try {
            $text = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }
    } catch {
        return $null  # 置き換えと重なった等。次の機会に読む
    }

    $lines = @($text -split "\r?\n")
    if ($lines.Count -eq 0 -or $lines[0] -ne (${convertPlanColumns} -join "`t")) {
        return $null
    }
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $fields = $lines[$i].Split("`t")
        if ($fields.Count -ne ${convertPlanColumns}.Count) {
            continue  # 空行・書き込みの途中
        }
        $row = [ordered]@{}
        for ($c = 0; $c -lt ${convertPlanColumns}.Count; $c++) {
            $column = ${convertPlanColumns}[$c]
            $value = $fields[$c]
            if ($c -ge 3) {
                # 件数の列。数値にできない場合は 0 とする
                $number = 0
                [void][int]::TryParse($value, [ref]$number)
                $value = $number
            }
            $row[$column] = $value
        }
        $rows.Add([pscustomobject]$row)
    }
    return , @($rows.ToArray())
}

function removeConvertPlan {
    param (
        [string]$path = ${convertPlanFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function writeConvertStartRequest {
    # 画面が「変換する」を選んだことを変換側に伝える（前回失敗したファイルも再変換するかも伝える）。
    # 変換側が読んでいる途中の内容を見ないよう、一時ファイルに書いてから置き換える
    param (
        [bool]$retryFailed = $false,
        [string]$path = ${convertStartRequestFile}
    )

    $lines = @()
    if ($retryFailed) {
        $lines = @(${retryFailedMark})
    }
    writeTextLinesAtomic $path $lines
}

function readConvertStartRequest {
    # 画面からの「変換する」の返事を読む。まだ無ければ $null
    param (
        [string]$path = ${convertStartRequestFile}
    )

    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    $lines = @()
    try {
        $lines = @(readListFile $path)
    } catch {
        return $null  # 置き換えと重なった等。次の機会に読む
    }
    return @{ RetryFailed = (@($lines | Where-Object { $_.Trim() -eq ${retryFailedMark} }).Count -gt 0) }
}

function removeConvertStartRequest {
    param (
        [string]$path = ${convertStartRequestFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function readStatusLines {
    # 変換一覧を1行ずつ読む（変換中でも読めるよう共有を許して開く）。ファイルが無ければ空
    param (
        [string]$path = ${statusFile}
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
    # 変換一覧に記録したインデックス名を書き換える（変換対象フォルダの行と、各行の相対パスの先頭）。
    # 行の順序と内容をそのまま保つため、行ごとに書き換えて置き換える
    param (
        [string]$oldName,
        [string]$newName,
        [string]$path = ${statusFile}
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
    # 変換一覧から、あるインデックスの記録（変換対象フォルダの行と、そのインデックスの各行）を取り除く
    param (
        [string]$name,
        [string]$path = ${statusFile}
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

function getConversionState {
    # 変換一覧から変換の状態を返す:
    #   @{ Exists; Folders（変換対象フォルダ @{ Path; Name } の配列）; Total; Pending（未変換）; Failed; Done; ConvertedSince（since 以降に変換した件数）; Updated（変換一覧の更新日時）;
    #      FailedRows（失敗したファイルの行。変換日時の新しい順）; IndexStats（インデックス名ごとの集計。getIndexStats） }
    param (
        [datetime]$since = [datetime]::MaxValue,
        [string]$path = ${statusFile}
    )

    $state = @{ Exists = $false; Folders = @(); Total = 0; Pending = 0; Failed = 0; Done = 0; ConvertedSince = 0; Updated = $null; FailedRows = @(); IndexStats = (getIndexStats $null) }
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

        $converted = [datetime]::MinValue
        if ($countSince -and $row.変換日時 -and [datetime]::TryParseExact($row.変換日時, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$converted) -and $converted -ge $sinceSecond) {
            $state.ConvertedSince++
        }
    }
    # 変換日時は "yyyy/MM/dd HH:mm:ss" のため、文字列の順で新しい順に並ぶ
    $state.FailedRows = @($failedRows | Sort-Object -Property @{ Expression = { [string]$_.変換日時 }; Descending = $true }, @{ Expression = { $_.相対パス } })
    return $state
}
