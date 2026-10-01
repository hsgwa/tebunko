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

function removeEmptyDir {
    # フォルダが空のときだけ、再帰しない削除で消す。中にほかのプロセスのフォルダなどが残っている・
    # フォルダが既に無いときは、何もせずそのままにする
    param (
        [string]$dir
    )
    try {
        [System.IO.Directory]::Delete($dir, $false)
    } catch [System.IO.IOException] {
        # 空でない、または既に無い
    }
}

function removeTmpDir {
    # 作業フォルダ（既定は work\tmp\<PC の鍵>\<PID>、代わりの場所は %TEMP%\tebunko\<PID>）と
    # 出力用のフォルダ（work\取り込み出力\<PID>）を削除する。終了時に呼ぶ。
    # ${tmpDir} が決まる前（空・未設定）に呼ばれても、作業フォルダを消さずに何もしない
    # （空の値から親やワークスペースを消さないため）
    if (!${tmpDir}) {
        return
    }
    foreach ($dir in @(${tmpDir}, $workspace.PublishDir)) {
        try {
            removeDirectoryRetry $dir
        } catch {
            writeIndexerLog "    作業フォルダを削除できませんでした: ${dir}" "Yellow"
        }
    }
    # 空になった tmp\<PC の鍵>・tmp も消す（ワークスペースの tmp を使ったときだけ。
    # 代わりの場所（%TEMP%）のときや、ほかのプロセスのフォルダが残っているときは消えない）
    if (${tmpDir}.StartsWith($workspace.TmpRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        removeEmptyDir (Split-Path ${tmpDir} -Parent)
        removeEmptyDir $workspace.TmpRoot
    }
}

function removeStaleTmpDirs {
    # 強制終了などで残った、ほかの（終了済みの）プロセスの作業フォルダ
    # （ワークスペースの tmp\<PC の鍵>\<PID>・work\取り込み出力\<PID>・前の版までの %TEMP%\tebunko\<PID>）を削除する。
    # 空になった親（tmp\<PC の鍵>・取り込み出力・${legacyTmpParent}。tmp\<PC の鍵> はさらに tmp 自体も）も消す
    foreach ($parent in @((Split-Path (getWorkspaceTmpDir $workspace) -Parent), (Split-Path $workspace.PublishDir -Parent), ${legacyTmpParent})) {
        removeStaleProcessDirs $parent
        removeEmptyDir $parent
    }
    removeEmptyDir $workspace.TmpRoot
}

function initTmpDir {
    # インデックス作成の始めに、取り込みの作業フォルダの片付けと用意をまとめて行う: @{ Dir; Reason（代わりの場所にしたときだけ） }
    #   1. 強制終了などで残った、ほかの（終了済みの）プロセスの作業フォルダを片付ける（決めた場所を巻き込まないよう、場所を決める前に行う）
    #   2. selectTmpDir で置き場所を決めてフォルダを作る
    #   3. 代わりの場所（%TEMP%）でなければ、tmp・tmp\<PC の鍵>・<PID> に NotContentIndexed を付ける
    #      （%TEMP% は Windows の既定で検索の対象から外れているため付けない）
    #   4. 代わりの場所にしたら、理由をインデックス作成ログに 1 行書く
    removeStaleTmpDirs
    $selected = selectTmpDir $workspace
    [System.IO.Directory]::CreateDirectory($selected.Dir) | Out-Null
    if (!$selected.Reason) {
        foreach ($dir in @($workspace.TmpRoot, (Split-Path $selected.Dir -Parent), $selected.Dir)) {
            [void](setNotContentIndexed $dir)
        }
    } elseif ($selected.Reason -eq "Brackets") {
        writeIndexerLog "ワークスペースのパスに [ ] が含まれるため、取り込みの作業フォルダを %TEMP% に置きます。" "Yellow"
    } else {
        writeIndexerLog "ワークスペースのパスが長いため、取り込みの作業フォルダを %TEMP% に置きます。" "Yellow"
    }
    return $selected
}

function newWorkerTmpDir {
    # 取り込みのスレッドの作業フォルダ（parentDir\w<番号>）を作り、パスを返す。
    # 親（parentDir。司令のスレッドの作業フォルダ）に NotContentIndexed が付いていれば、同じ属性を付ける
    # （代わりの場所（%TEMP%）には付けていないため、付いていなければ何もしない）
    param (
        [string]$parentDir,
        [int]$number
    )

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
