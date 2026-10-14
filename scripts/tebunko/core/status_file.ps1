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
    # 取り込み一覧を読み込み、@{ Folders; Rows; Rejected; LegacyFolders } を返す。ファイルが無ければ空。
    #   Folders : クロール対象フォルダ @{ Path; Name（インデックス名） } の配列
    #   LegacyFolders: 名前が今の決まり（予約名を許す getFolderNameProblemText）に合わないが、フォルダ名として外へ出ない（前の版で作られた）クロール対象フォルダ。
    #             Folders と同じ形。名前でパスを組んで使わない（最後の . や空白は、パスにすると別のフォルダを指す）
    #   Rows    : 相対パス（"インデックス名\フォルダからの相対パス"。大文字・小文字を区別しない）→ 行
    #   Rejected: 名前（同じ決まり）・相対パス（testIndexRelPath）が使えないため読み飛ばした行の数
    #             （読み飛ばした行は、次に取り込み一覧を書き直すときに消える）
    # 既にある名前は、予約名（aux.backup など）でもそのまま読む（空にしたり読み飛ばしたりすると、別の名前のフォルダになり取り込み直しになる）。
    # インデックス作成中は1件ごとに行を追記するため、同じ相対パスの行は後の行を優先する。列数の合わない行（書き込み途中で中断した行など）は無視する
    param (
        [string]$path = $workspace.StatusFile
    )

    $rows = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $folders = New-Object System.Collections.Generic.List[object]
    $legacyFolders = New-Object System.Collections.Generic.List[object]
    $rejected = 0
    $result = @{ Folders = $folders; Rows = $rows; Rejected = 0; LegacyFolders = $legacyFolders }
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
    # 共有のワークスペースでは他の人が書き換えられるため、名前・相対パスが使えない行（.. で外へ出るものなど）は読み飛ばす。
    # 行の確かめは testIndexRelPath（名前だけの行は通さない）と同じ決めを、数万行を読むためその場で書く（下の行の組み立てと同じ理由）。
    # 名前が今の決まり（予約名を許す getFolderNameProblemText）に合わなくても、フォルダ名として外へ出ないもの（前の版で作られた名前）は、
    # LegacyFolders に分けて返す（設定から外れたインデックスとして、確認のうえ削除できるように。名前でパスを組まずに消す: removeDroppedFolders）
    $checkedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    $rejectedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3) {
            if (!$checkedNames.Contains($fields[2]) -and (getFolderNameProblemText $fields[2] $true)) {
                if (testPlainFolderName $fields[2]) {
                    $legacyFolders.Add([pscustomobject]@{ Path = $fields[1]; Name = $fields[2] })
                } else {
                    $rejected++
                }
                continue
            }
            [void]$checkedNames.Add($fields[2])
            $folders.Add([pscustomobject]@{ Path = $fields[1]; Name = $fields[2] })
            continue
        }
        if ($fields.Count -ne $columnCount -or $fields[0] -eq "" -or $fields[0] -eq ${statusColumns}[0]) {
            continue
        }
        $relPath = $fields[0]
        $separator = $relPath.IndexOf("\")
        $usable = $separator -gt 0
        if ($usable) {
            $rowName = $relPath.Substring(0, $separator)
            if (!$checkedNames.Contains($rowName)) {
                if ($rejectedNames.Contains($rowName)) {
                    $usable = $false
                } elseif (getFolderNameProblemText $rowName $true) {
                    # 通らない名前も覚え、同じ名前の 2 行目からは調べ直さない
                    [void]$rejectedNames.Add($rowName)
                    $usable = $false
                } else {
                    [void]$checkedNames.Add($rowName)
                }
            }
        }
        if ($usable) {
            foreach ($segment in $relPath.Substring($separator + 1).Split("\")) {
                if ($segment -eq "" -or $segment -eq "." -or $segment -eq ".." -or $segment.Length -gt ${maxFileNameLength} -or
                    $segment.IndexOfAny(${invalidFileNameChars}) -ge 0) {
                    $usable = $false
                    break
                }
            }
        }
        if (!$usable) {
            $rejected++
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
    $result.Rejected = $rejected
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
