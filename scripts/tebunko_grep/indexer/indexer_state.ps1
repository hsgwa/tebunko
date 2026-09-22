# インデックス作成の状態を表すファイル（取り込み一覧・取り込み中・進捗・予定・開始要求）の読み書き。

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

    if (!${fullLanguage}) {
        # 制限言語モード（制限モードのインデックス作成）では [pscustomobject] を作れないため、同じ項目のハッシュテーブルにする
        return @{ 相対パス = $relPath; 更新日時 = $updated; サイズ = $size; 状態 = $state; TSV数 = $tsvCount; 取り込み日時 = $ingested; エラー = $errorMessage; 抽出版 = $extractVersion }
    }
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
    #   Folders: クロール対象フォルダ @{ Path; Name（インデックス名） } の配列。以前の形式（インデックス名なし）は Name が空
    #   Rows   : 相対パス（"インデックス名\フォルダからの相対パス"。大文字・小文字を区別しない）→ 行
    # インデックス作成中は1件ごとに行を追記するため、同じ相対パスの行は後の行を優先する。列数の合わない行（書き込み途中で中断した行など）は無視する
    param (
        [string]$path = ${statusFile}
    )

    if (!${fullLanguage}) {
        return (readStatusFileClm $path)
    }
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

    # 以前の形式（抽出版の列が無い）は見出しの列数で見分け、その見出しの後の行は以前の列数で読む。
    # インデックス作成中の追記は、見出しを今の形式で書き直した後に行うため、1 つのファイルで形式が混ざることは無い
    $columnCount = ${statusColumns}.Count
    foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and ($fields.Count -eq 2 -or $fields.Count -eq 3)) {
            $folders.Add([pscustomobject]@{ Path = $fields[1]; Name = $(if ($fields.Count -eq 3) { $fields[2] } else { "" }) })
            continue
        }
        if ($fields[0] -eq ${statusColumns}[0]) {
            $columnCount = $fields.Count  # 見出し
            continue
        }
        if ($fields.Count -ne $columnCount -or $fields[0] -eq "") {
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
            抽出版   = $(if ($fields.Count -gt 7) { $fields[7] } else { "" })
        }
    }
    return $result
}

function readStatusFileClm {
    # readStatusFile の制限言語モードの書き方（FileStream・Dictionary・[pscustomobject] を使えない）。
    # Rows はハッシュテーブル（大文字・小文字を区別しない）、行もハッシュテーブル（項目は newStatusRow と同じ）、Folders は配列にする
    param (
        [string]$path
    )

    $rows = @{}
    $folders = @()
    if (!(Test-Path -LiteralPath $path)) {
        return @{ Folders = $folders; Rows = $rows }
    }
    $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
    if ($null -eq $text) {
        $text = ""
    }
    $columnCount = ${statusColumns}.Count
    foreach ($line in ($text -split "\r?\n")) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and ($fields.Count -eq 2 -or $fields.Count -eq 3)) {
            $folders += New-Object PSObject -Property ([ordered]@{ Path = $fields[1]; Name = $(if ($fields.Count -eq 3) { $fields[2] } else { "" }) })
            continue
        }
        if ($fields[0] -eq ${statusColumns}[0]) {
            $columnCount = $fields.Count
            continue
        }
        if ($fields.Count -ne $columnCount -or $fields[0] -eq "") {
            continue
        }
        $rows[$fields[0]] = @{
            相対パス = $fields[0]; 更新日時 = $fields[1]; サイズ = $fields[2]; 状態 = $fields[3]; TSV数 = $fields[4]
            取り込み日時 = $fields[5]; エラー = $fields[6]; 抽出版 = $(if ($fields.Count -gt 7) { $fields[7] } else { "" })
        }
    }
    return @{ Folders = $folders; Rows = $rows }
}

function writeStatusFile {
    # 取り込み一覧を書き出す（1ファイル1行）。途中で中断しても壊れないよう、一時ファイルに書いてから置き換える
    #   folders: クロール対象フォルダ @{ Path; Name } の配列。"クロール対象フォルダ<TAB>パス<TAB>インデックス名" の行にする
    param (
        [object[]]$folders,
        [object[]]$rows,
        [string]$path = ${statusFile}
    )

    $head = @(foreach ($folder in @($folders | Where-Object { $_ })) {
            "${statusFolderKey}`t$($folder.Path)`t$($folder.Name)"
        }) + @(${statusColumns} -join "`t")
    # 数万行を書き出すため、1行ごとの関数呼び出し（toStatusLine）・パイプライン（Where-Object）は使わない
    # （5万行で約 10 秒 → 約 0.2 秒。この間はインデックス作成の進み具合が止まって見えるため速くする）。
    # 制限言語モードでも動くよう、List ではなく foreach の出力を受けて配列にする
    $body = @(foreach ($row in $rows) {
            if ($null -eq $row) {
                continue
            }
            $errorText = [string]$row.エラー
            if ($errorText.IndexOfAny(${statusLineBreaks}) -ge 0) {
                $errorText = $errorText -replace "[\t\r\n]+", " "
            }
            [string]::Join("`t", @(
                    [string]$row.相対パス, [string]$row.更新日時, [string]$row.サイズ, [string]$row.状態,
                    [string]$row.TSV数, [string]$row.取り込み日時, $errorText, [string]$row.抽出版))
        })

    writeTextLinesAtomic $path ($head + $body)
}

function addStatusRow {
    # 取り込み一覧の末尾に1行追記する（readStatusFile では後の行が優先される）
    param (
        $row,
        [string]$path = ${statusFile}
    )

    if (!${fullLanguage}) {
        # 制限言語モード: Add-Content -Encoding UTF8 は、既にあるファイルには BOM を付けずに足す（AppendAllText と同じ中身）
        Add-Content -LiteralPath $path -Value (toStatusLine $row) -Encoding UTF8
        return
    }
    [System.IO.File]::AppendAllText($path, "$(toStatusLine $row)`r`n", ${utf8Bom})
}

function readIngestingFile {
    # 取り込み中のファイルの記録を読み、@{ RelPath = 相対パス; Count = 続けて取り込みを始めて終わらなかった回数 } を返す。
    # 記録が無い・壊れている場合は $null
    param (
        [string]$path = ${ingestingFile}
    )

    $lines = @(readListFile $path)
    if ($lines.Count -eq 0) {
        return $null
    }
    $fields = $lines[0].Split("`t")
    # 回数は [int]::TryParse と同じく読む（[ref] は制限言語モードで使えないため、形を確かめてから [int] にする）
    if ($fields.Count -ne 2 -or $fields[0] -notmatch '^\s*[+-]?[0-9]{1,9}\s*$' -or $fields[1] -eq "") {
        return $null
    }
    $count = [int]$fields[0]
    if ($count -lt 1) {
        return $null
    }
    return @{ RelPath = $fields[1]; Count = $count }
}

function writeIngestingFile {
    # 取り込みを始めるファイルを "回数<TAB>相対パス" で記録する。取り込みが終われば removeIngestingFile で消す
    param (
        [string]$relPath,
        [int]$count,
        [string]$path = ${ingestingFile}
    )

    writeListFile $path @("${count}`t${relPath}")
}

function removeIngestingFile {
    param (
        [string]$path = ${ingestingFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function writeIndexingProgress {
    # インデックス作成の進み具合を1行で書く（画面が読む）。書き込みは1ファイルにつき1回で、インデックス作成の速さに影響しない大きさにする。
    #   "<段階><TAB><処理済み><TAB><残り><TAB><失敗><TAB><いま行っていること>"
    # 画面が読んでいる最中でも書けるよう、共有を許して開く
    param (
        [string]$phase,
        [int]$processed = 0,
        [int]$remaining = 0,
        [int]$failed = 0,
        [string]$detail = "",
        [string]$path = ${indexingProgressFile}
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
        # 進み具合の表示のためだけのファイルのため、書けなくてもインデックス作成は続ける
    }
}

function readIndexingProgress {
    # インデックス作成の進み具合を読む（無い・壊れていれば $null）。画面が毎秒呼ぶため、1行だけ読む
    param (
        [string]$path = ${indexingProgressFile}
    )

    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    # インデクサが書いている最中でも読めるよう、共有を許して開く
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

function removeIndexingProgress {
    param (
        [string]$path = ${indexingProgressFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
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

    # インデックス 1 件に 1 回だけ作るため、制限言語モードでも作れる New-Object PSObject にする
    return New-Object PSObject -Property ([ordered]@{
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
    })
}

function writeIngestPlan {
    # 取り込み予定を書き出す（インデックス1件1行）。画面は数える前から読むため、
    # 途中の状態を読ませないよう一時ファイルに書いてから置き換える
    param (
        [object[]]$rows,
        [string]$path = ${ingestPlanFile}
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add((${ingestPlanColumns} -join "`t"))
    foreach ($row in @($rows | Where-Object { $_ })) {
        $values = foreach ($column in ${ingestPlanColumns}) { ([string]$row.$column) -replace "[\t\r\n]+", " " }
        $lines.Add([string]::Join("`t", @($values)))
    }
    writeTextLinesAtomic $path $lines
}

function readIngestPlan {
    # 取り込み予定を読む。ファイルが無い・列が合わない場合は $null（画面は次の機会に読み直す）。
    # 件数の列は数値にして返す
    param (
        [string]$path = ${ingestPlanFile}
    )

    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    # インデクサが置き換えている最中でも読めるよう、共有を許して開く
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
    if ($lines.Count -eq 0 -or $lines[0] -ne (${ingestPlanColumns} -join "`t")) {
        return $null
    }
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $fields = $lines[$i].Split("`t")
        if ($fields.Count -ne ${ingestPlanColumns}.Count) {
            continue  # 空行・書き込みの途中
        }
        $row = [ordered]@{}
        for ($c = 0; $c -lt ${ingestPlanColumns}.Count; $c++) {
            $column = ${ingestPlanColumns}[$c]
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

function removeIngestPlan {
    param (
        [string]$path = ${ingestPlanFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function writeIndexingStartRequest {
    # 画面が「取り込む」を選んだことをインデクサに伝える（前回失敗したファイルも再取り込みするかも伝える）。
    # インデクサが読んでいる途中の内容を見ないよう、一時ファイルに書いてから置き換える
    param (
        [bool]$retryFailed = $false,
        [string]$path = ${indexingStartRequestFile}
    )

    $lines = @()
    if ($retryFailed) {
        $lines = @(${retryFailedMark})
    }
    writeTextLinesAtomic $path $lines
}

function readIndexingStartRequest {
    # 画面からの「取り込む」の返事を読む。まだ無ければ $null
    param (
        [string]$path = ${indexingStartRequestFile}
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

function removeIndexingStartRequest {
    param (
        [string]$path = ${indexingStartRequestFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function readStatusLines {
    # 取り込み一覧を1行ずつ読む（インデックス作成中でも読めるよう共有を許して開く）。ファイルが無ければ空
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
    # 取り込み一覧に記録したインデックス名を書き換える（クロール対象フォルダの行と、各行の相対パスの先頭）。
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

    # 制限モード（制限言語モード）からも使うため、List ではなく配列に集める
    $result = @(foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3 -and [string]::Equals($fields[2], $oldName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $fields[2] = $newName
            $fields -join "`t"
            continue
        }
        # 以前の形式（抽出版の列が無い）の行も readStatusFile は読むため、同じように扱う
        if ($fields.Count -ge ${statusColumns}.Count - 1 -and $fields.Count -le ${statusColumns}.Count -and $fields[0] -ne "") {
            $split = splitIndexRelPath $fields[0]
            if ($split.Rest -ne "" -and [string]::Equals($split.Name, $oldName, [System.StringComparison]::OrdinalIgnoreCase)) {
                $fields[0] = "${newName}\$($split.Rest)"
                $fields -join "`t"
                continue
            }
        }
        $line
    })
    writeTextLinesAtomic $path $result
}

function removeStatusIndexName {
    # 取り込み一覧から、あるインデックスの記録（クロール対象フォルダの行と、そのインデックスの各行）を取り除く
    param (
        [string]$name,
        [string]$path = ${statusFile}
    )

    $lines = @(readStatusLines $path)
    if ($lines.Count -eq 0) {
        return
    }

    # 制限モード（制限言語モード）からも使うため、List ではなく配列に集める
    $result = @(foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3 -and [string]::Equals($fields[2], $name, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        # 以前の形式（抽出版の列が無い）の行も readStatusFile は読むため、同じように扱う
        if ($fields.Count -ge ${statusColumns}.Count - 1 -and $fields.Count -le ${statusColumns}.Count -and $fields[0] -ne "") {
            $split = splitIndexRelPath $fields[0]
            if ($split.Rest -ne "" -and [string]::Equals($split.Name, $name, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
        }
        $line
    })
    writeTextLinesAtomic $path $result
}

function getIndexingState {
    # 取り込み一覧からインデックス作成の状態を返す:
    #   @{ Exists; Folders（クロール対象フォルダ @{ Path; Name } の配列）; Total; Pending（未取り込み）; Failed; Done; IngestedSince（since 以降に取り込んだ件数）; Updated（取り込み一覧の更新日時）;
    #      FailedRows（失敗したファイルの行。取り込み日時の新しい順）; IndexStats（インデックス名ごとの集計。getIndexStats） }
    param (
        [datetime]$since = [datetime]::MaxValue,
        [string]$path = ${statusFile}
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
