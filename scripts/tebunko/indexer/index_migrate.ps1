# TSV のインデックスへの取り込みと、作業フォルダ・外したフォルダのインデックスの後始末。

function publishTsv {
    # 作業フォルダのTSVを、そのファイルのインデックスのフォルダへ移動する。
    # 途中で強制終了されても一部のシートだけのインデックスが残らないよう、
    # 出力用のフォルダ（work\publish\<PID>）に集めてからフォルダごと入れ替える（publishIndexFiles）
    param (
        [string]$bookDir
    )

    publishIndexFiles $tmpDir $bookDir (Join-Path $workspace.PublishDir ([System.IO.Path]::GetFileName($bookDir)))
}

function clearTmpDir {
    Get-ChildItem -LiteralPath (toLongPath $tmpDir) -File | Remove-Item -Force
}

function removeEmptyDir {
    # フォルダが空のときだけ、再帰しない削除で消す。中にほかのプロセスのフォルダなどが残っている・
    # フォルダが既に無いときは、何もせずそのままにする
    param (
        [string]$dir
    )
    try {
        [System.IO.Directory]::Delete($dir, $false)
    } catch {
        # 空でない・既に無い・アクセス権が無い（共有フォルダで他の利用者のフォルダなど）
    }
}

function removeTmpDir {
    # 作業フォルダ（work\tmp\<PC の鍵>\<PID>）と出力用のフォルダ（work\publish\<PID>）を削除する。終了時に呼ぶ。
    # ${tmpDir} が決まっていなくても（置けなかった・決める前）、出力用のフォルダは消す
    # （$workspace.PublishDir は ${tmpDir} が置けないときも作られるため）
    foreach ($dir in @(${tmpDir}, $workspace.PublishDir) | Where-Object { $_ }) {
        try {
            removeDirectoryRetry $dir
        } catch {
            writeIndexerLog "    作業フォルダを削除できませんでした: ${dir}" "Yellow"
        }
    }
    # 空になった tmp\<PC の鍵>・tmp も消す（ワークスペースの tmp を使えたときだけ。
    # ほかのプロセスのフォルダが残っているときは消えない）
    if (${tmpDir} -and ${tmpDir}.StartsWith($workspace.TmpRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        removeEmptyDir (Split-Path ${tmpDir} -Parent)
        removeEmptyDir $workspace.TmpRoot
    }
}

function removeStaleTmpDirs {
    # 強制終了などで残った、ほかの（終了済みの）プロセスの作業フォルダ
    # （ワークスペースの tmp\<PC の鍵>\<PID>・work\publish\<PID>・前の版までの %TEMP%\tebunko\<PID>）を削除する。
    # 空になった親（tmp\<PC の鍵>・publish・${legacyTmpParent}。tmp\<PC の鍵> はさらに tmp 自体も）も消す
    foreach ($parent in @((Split-Path (getWorkspaceTmpDir $workspace) -Parent), (Split-Path $workspace.PublishDir -Parent), ${legacyTmpParent})) {
        removeStaleProcessDirs $parent
        removeEmptyDir $parent
    }
    removeEmptyDir $workspace.TmpRoot
}

function initTmpDir {
    # インデックス作成の始めに、取り込みの作業フォルダの片付けと用意をまとめて行う: @{ Dir; Reason（置けなかったときだけ） }
    #   1. 強制終了などで残った、ほかの（終了済みの）プロセスの作業フォルダを片付ける（決めた場所を巻き込まないよう、場所を決める前に行う）
    #   2. selectTmpDir で置き場所を決める
    #   3. 置けたら、フォルダを作って tmp・tmp\<PC の鍵>・<PID> に NotContentIndexed を付ける
    #   4. 置けなければ（Dir が空）、フォルダは作らず、理由と「取り込みをすべてスキップする」ことをログに 1 行書く
    #      （どのファイルも中間 TSV などをこのフォルダに作るため、テキストファイルを含めすべての取り込みが対象になる）
    #      （1 ファイルごとのスキップは invokeIngestTask が取り込みの失敗として記録する）
    removeStaleTmpDirs
    $selected = selectTmpDir $workspace
    if (!$selected.Dir) {
        writeIndexerLog "$(getTmpDirUnavailableMessage $selected.Reason)取り込みをすべてスキップします。" "Yellow"
        return $selected
    }
    [System.IO.Directory]::CreateDirectory($selected.Dir) | Out-Null
    foreach ($dir in @($workspace.TmpRoot, (Split-Path $selected.Dir -Parent), $selected.Dir)) {
        [void](setNotContentIndexed $dir)
    }
    return $selected
}

function newWorkerTmpDir {
    # 取り込みのスレッドの作業フォルダ（parentDir\w<番号>）を作り、パスを返す。
    # 親（parentDir。司令のスレッドの作業フォルダ）に NotContentIndexed が付いていれば、同じ属性を付ける。
    # 親が無い（置けなかった）ときは、何も作らず空文字を返す
    param (
        [string]$parentDir,
        [int]$number
    )

    if (!$parentDir) {
        return ""
    }
    $dir = Join-Path $parentDir "w$number"
    [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    if (testNotContentIndexed $parentDir) {
        [void](setNotContentIndexed $dir)
    }
    return $dir
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
    # クロール対象フォルダから削除されたフォルダのインデックス（work\content_index\<インデックス名>）を削除する。
    # チェックを外しただけのフォルダは削除しない
    param (
        [object[]]$folders,          # assignIndexNames の結果
        [object[]]$previousFolders,  # readStatusFile の Folders
        $ws = $workspace             # 対象のワークスペース（テストが差し替える）
    )

    # インデックス名で比べる。フォルダを移動して登録し直した場合は、同じ名前を引き継ぐため削除しない（assignIndexNames）
    $current = @($folders | ForEach-Object { $_.Name })
    foreach ($previous in @($previousFolders | Where-Object { $_.Name -and $current -notcontains $_.Name })) {
        $dir = Join-Path $ws.IndexDir $previous.Name
        if (Test-Path -LiteralPath $dir) {
            # 中に長いパス（260文字超）のTSVがあっても削除できるよう \\?\ 付きで削除する
            Remove-Item -LiteralPath (toLongPath $dir) -Recurse -Force
        }
        writeIndexerLog "クロール対象から削除されたフォルダ（$($previous.Path)）のインデックスを削除しました。"
    }
}
