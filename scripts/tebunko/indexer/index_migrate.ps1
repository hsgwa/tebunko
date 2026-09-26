# TSV のインデックスへの取り込みと、作業フォルダ・外したフォルダのインデックスの後始末。

function publishTsv {
    # 作業フォルダのTSVを、そのファイルのインデックスのフォルダへ移動する。
    # 途中で強制終了されても一部のシートだけのインデックスが残らないよう、
    # 出力用のフォルダ（work\取り込み出力\<PID>）に集めてからフォルダごと入れ替える（publishIndexFiles）
    param (
        [string]$bookDir
    )

    publishIndexFiles $tmpDir $bookDir (Join-Path $workspace.PublishDir ([System.IO.Path]::GetFileName($bookDir)))
}

function clearTmpDir {
    Get-ChildItem -LiteralPath (toLongPath $tmpDir) -File | Remove-Item -Force
}

function removeTmpDir {
    # 作業フォルダ（%TEMP%\tebunko\<PID>）と出力用のフォルダ（work\取り込み出力\<PID>）を削除する。終了時に呼ぶ
    foreach ($dir in @(${tmpDir}, $workspace.PublishDir)) {
        try {
            removeDirectoryRetry $dir
        } catch {
            writeIndexerLog "    作業フォルダを削除できませんでした: ${dir}" "Yellow"
        }
    }
}

function removeStaleTmpDirs {
    # 強制終了などで残った、ほかの（終了済みの）プロセスの作業フォルダ
    # （%TEMP%\tebunko\<PID>・work\取り込み出力\<PID>）を削除する
    foreach ($parent in @((Split-Path ${tmpDir} -Parent), (Split-Path $workspace.PublishDir -Parent))) {
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
        $dir = Join-Path $workspace.IndexDir $previous.Name
        if (Test-Path -LiteralPath $dir) {
            # 中に長いパス（260文字超）のTSVがあっても削除できるよう \\?\ 付きで削除する
            Remove-Item -LiteralPath (toLongPath $dir) -Recurse -Force
        }
        writeIndexerLog "クロール対象から削除されたフォルダ（$($previous.Path)）のインデックスを削除しました。"
    }
}
