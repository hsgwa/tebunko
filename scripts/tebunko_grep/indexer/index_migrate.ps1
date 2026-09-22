# TSV のインデックスへの取り込みと、作業フォルダ・以前の形式の後始末。

function publishTsv {
    # 作業フォルダのTSVを、そのファイルのインデックスのフォルダへ移動する。
    # 途中で強制終了されても一部のシートだけのインデックスが残らないよう、
    # 出力用のフォルダ（work\取り込み出力\<PID>）に集めてからフォルダごと入れ替える（publishIndexFiles）
    param (
        [string]$bookDir
    )

    publishIndexFiles $tmpDir $bookDir (Join-Path ${publishDir} (getPathLeaf $bookDir))
}

function clearTmpDir {
    Get-ChildItem -LiteralPath (toLongPath $tmpDir) -File | Remove-Item -Force
}

function removeTmpDir {
    # 作業フォルダ（%TEMP%\tebunko_grep\<PID>）と出力用のフォルダ（work\取り込み出力\<PID>）を削除する。終了時に呼ぶ
    foreach ($dir in @(${tmpDir}, ${publishDir})) {
        try {
            removeDirectoryRetry $dir
        } catch {
            Write-Host "    作業フォルダを削除できませんでした: ${dir}" -ForegroundColor Yellow
        }
    }
}

function removeStaleTmpDirs {
    # 強制終了などで残った、ほかの（終了済みの）プロセスの作業フォルダ
    # （%TEMP%\tebunko_grep\<PID>・work\取り込み出力\<PID>）を削除する
    foreach ($parent in @((Split-Path ${tmpDir} -Parent), (Split-Path ${publishDir} -Parent))) {
        removeStaleProcessDirs $parent
    }
}

function removeStaleProcessDirs {
    # プロセスIDの名前のフォルダのうち、そのプロセスが既に終わっているものを削除する
    param (
        [string]$parent
    )

    if (!(Test-Path -LiteralPath $parent)) {
        return
    }
    foreach ($dir in @(Get-ChildItem -LiteralPath $parent -Directory | Where-Object { $_.Name -match "^\d{1,9}$" -and [int]$_.Name -ne $PID })) {
        if (Get-Process -Id ([int]$dir.Name) -ErrorAction SilentlyContinue) {
            continue  # 実行中のインデックス作成（またはPIDを再利用した別のプロセス）のものは残す
        }
        try {
            Remove-Item -LiteralPath (toLongPath $dir.FullName) -Recurse -Force
        } catch {
            # 使用中などで削除できなければ、次回に回す
        }
    }
}

function moveLegacyIndex {
    # 以前の形式（work\index 直下にクロール対象フォルダ1つ分のインデックスがある）を、そのフォルダのインデックス名の下へ移す。
    # 前回の取り込み一覧の行（相対パス → 行）を、移した後の相対パス（"インデックス名\…"）で返す
    param (
        [object[]]$folders,  # assignIndexNames の結果
        $status,             # readStatusFile の結果
        [bool]$statusExists
    )

    $legacyPath = $null
    # 前回の移行が途中（work\index を別名にした直後）で止まると、_移行中 が残り、work\index は空で作り直されている
    $movingDir = "${indexDir}_移行中"
    $resuming = [System.IO.Directory]::Exists($movingDir)
    $legacyFolder = @($status.Folders | Where-Object { -not $_.Name }) | Select-Object -First 1
    if ($legacyFolder) {
        $legacyPath = normalizeFolderPath $legacyFolder.Path
    } elseif (-not $statusExists) {
        # 取り込み一覧が無い（取り込み一覧を使う前の版）: work\index 直下が各フォルダのインデックス名のフォルダだけでなければ、
        # 以前の形式で、クロール対象フォルダの1件目のフォルダのインデックスとみなす
        $names = @($folders | ForEach-Object { $_.Name })
        $others = @(Get-ChildItem -LiteralPath $indexDir -Force | Where-Object {
            ($_.PSIsContainer -and $names -notcontains $_.Name) -or (-not $_.PSIsContainer -and $_.Name -ne ${sourceFolderFileName})
        })
        if ($others.Count -gt 0 -or $resuming) {
            $legacyPath = $folders[0].Path
        }
    }
    if (!$legacyPath) {
        return , $status.Rows
    }

    $rows = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $folder = @($folders | Where-Object { $_.Path -eq $legacyPath }) | Select-Object -First 1
    if (!$folder) {
        Write-Host "work\index 直下に、クロール対象から外したフォルダ（${legacyPath}）の以前の形式のインデックスがあります。不要なら削除してください。" -ForegroundColor Yellow
        return , $rows
    }

    # work\index を丸ごと work\index\<インデックス名> に移す（同じ名前のサブフォルダがあっても衝突しないよう、いったん別名にする）。
    # 前回の移行が途中で止まっていれば、残った _移行中 を移すところから続ける
    if (-not $resuming) {
        [System.IO.Directory]::Move($indexDir, $movingDir)
    }
    [System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
    [System.IO.Directory]::Move($movingDir, (Join-Path $indexDir $folder.Name))
    Write-Host "以前の形式のインデックスを work\index\$($folder.Name) に移しました。（$($folder.Path) のインデックス）"

    foreach ($row in $status.Rows.Values) {
        $row.相対パス = "$($folder.Name)\$($row.相対パス)"
        $rows[$row.相対パス] = $row
    }
    return , $rows
}

function removeDroppedFolders {
    # クロール対象フォルダから削除されたフォルダのインデックス（work\index\<インデックス名>）を削除する。
    # チェックを外しただけのフォルダは削除しない
    param (
        [object[]]$folders,          # assignIndexNames の結果
        [object[]]$previousFolders   # readStatusFile の Folders
    )

    # インデックス名で比べる。フォルダを移動して登録し直した場合は、同じ名前を引き継ぐため削除しない（assignIndexNames）
    $current = @($folders | ForEach-Object { $_.Name })
    foreach ($previous in @($previousFolders | Where-Object { $_.Name -and $current -notcontains $_.Name })) {
        $dir = Join-Path $indexDir $previous.Name
        if (Test-Path -LiteralPath $dir) {
            # 中に長いパス（260文字超）のTSVがあっても削除できるよう \\?\ 付きで削除する
            Remove-Item -LiteralPath (toLongPath $dir) -Recurse -Force
        }
        Write-Host "クロール対象から削除されたフォルダ（$($previous.Path)）のインデックスを削除しました。"
    }
}

function migrateFlatIndex {
    # 以前の形式のTSV（<ファイル名>_<場所>.tsv）を、今の形式（<ファイル名>\<場所>.tsv）へ移す。
    # 取り込み直さず、名前を変えるだけ（更新日時もそのまま）。
    # 今の形式のTSVの名前は場所だけ（_ は符号化されている）のため、_ を含む名前が以前の形式
    $moved = 0
    $failed = 0
    foreach ($file in @(Get-ChildItem -LiteralPath (toLongPath $indexDir) -Filter "*.tsv" -File -Recurse -ErrorAction SilentlyContinue)) {
        if ($file.Name.IndexOf("_") -lt 0) {
            continue
        }
        $name = splitIndexFileName $file.Name
        if ($name.book -eq $file.Name) {
            continue  # ファイル名と場所に分けられないものは触らない
        }

        $bookDir = Join-Path ([System.IO.Path]::GetDirectoryName((fromLongPath $file.FullName))) $name.book
        try {
            $dest = Join-Path $bookDir (toIndexFileName $name.sheet)
            [System.IO.Directory]::CreateDirectory((toLongPath $bookDir)) | Out-Null
            if (Test-Path -LiteralPath (toLongPath $dest)) {
                Remove-Item -LiteralPath (toLongPath $dest) -Force
            }
            [System.IO.File]::Move($file.FullName, (toLongPath $dest))
            $moved++
        } catch {
            $failed++
        }
    }

    if ($moved -gt 0) {
        Write-Host "以前の形式のTSV ${moved} 件を、元のファイル名のフォルダへ移しました。（取り込み直しません）"
    }
    if ($failed -gt 0) {
        Write-Host "以前の形式のTSV ${failed} 件は移せませんでした。該当のファイルは次のインデックス作成で作り直します。" -ForegroundColor Yellow
    }
}
