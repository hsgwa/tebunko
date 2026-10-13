# 取り込み一覧（status ファイル）の読み書き（状態層）。列と状態の定義は paths.ps1。

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

function getIndexNameMap {
    # 取り込み一覧に記録したインデックス名 → クロール対象フォルダのパス（大文字・小文字を区別しない）。
    # クロール対象フォルダの行は先頭にあるため、見出し行まで読んで打ち切る（取り込み一覧が大きくても時間がかからないように）
    param (
        [string]$path = $workspace.StatusFile
    )

    $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if (!(Test-Path -LiteralPath $path)) {
        return , $map
    }
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        $header = ${statusColumns} -join "`t"
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line -eq $header) {
                break
            }
            $fields = $line.Split("`t")
            if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3 -and $fields[2]) {
                $map[$fields[2]] = $fields[1]
            }
        }
    } finally {
        $reader.Dispose()
    }
    return , $map
}
