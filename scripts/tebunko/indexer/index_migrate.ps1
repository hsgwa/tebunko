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
    # 作業フォルダ直下のファイルを消す。作業フォルダそのものがリンク（ジャンクション・シンボリックリンク）のときは、
    # 先のファイルを消さないよう、何も消さずにログに出す
    if (testPathIsLink $tmpDir) {
        writeIndexerLog "    作業フォルダがリンク（ジャンクション・シンボリックリンク）のため、中のファイルを消しませんでした: ${tmpDir}" "Yellow"
        return
    }
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
    # tmp\<PC の鍵> の途中にリンク（ジャンクション・シンボリックリンク）があるときは、先のフォルダを消さないよう、片付けない
    $keyDir = Split-Path (getWorkspaceTmpDir $workspace) -Parent
    foreach ($parent in @($keyDir, (Split-Path $workspace.PublishDir -Parent), ${legacyTmpParent})) {
        if (($parent -eq $keyDir) -and (testPathHasLink $workspace.TmpRoot $keyDir)) {
            writeIndexerLog "    作業領域の途中にリンク（ジャンクション・シンボリックリンク）があるため、残った作業フォルダを片付けませんでした: ${keyDir}" "Yellow"
            continue
        }
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
    # 作業領域（tmp\<PC の鍵>\<PID>）の途中にリンク（ジャンクション・シンボリックリンク）があれば、先に書かないよう置けない扱いにする
    # （tmp そのものは見ない。利用者が tmp を別の場所へのリンクにした構成も、今までどおり使える）。
    # 同じ名前（前のプロセスの残り・先に置かれたもの）が既にあれば、リンクは辿らずに消してから作り直す
    $parentLinked = testPathHasLink $workspace.TmpRoot (Split-Path $selected.Dir -Parent)
    if (!$parentLinked) {
        try {
            removeDirectoryRetry $selected.Dir
        } catch {
            # 消せなくても、下の確かめでリンクが残っていなければ、そのまま使う
        }
        [System.IO.Directory]::CreateDirectory($selected.Dir) | Out-Null
    }
    if ($parentLinked -or (testPathHasLink $workspace.TmpRoot $selected.Dir)) {
        $unavailable = @{ Dir = ""; Reason = "Link" }
        writeIndexerLog "$(getTmpDirUnavailableMessage $unavailable.Reason)取り込みをすべてスキップします。" "Yellow"
        return $unavailable
    }
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
    if (testPathIsLink $parentDir) {
        writeIndexerLog "    作業フォルダがリンク（ジャンクション・シンボリックリンク）のため、取り込みの作業フォルダを置けません: ${parentDir}" "Yellow"
        return ""
    }
    $dir = Join-Path $parentDir "w$number"
    # 同じ名前（前の残り・先に置かれたリンク）が既にあれば、リンクは辿らずに消してから作り直す
    try {
        removeDirectoryRetry $dir
    } catch {
        # 消せなくても、下の確かめでリンクが残っていなければ、そのまま使う
    }
    [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    if (testPathIsLink $dir) {
        return ""
    }
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
            # 中に長いパスがあっても消え、中のリンクの先は消さない。使用中などで消せなければ 1 回だけ試して次回に回す
            removeDirectoryRetry $dir.FullName 1
        } catch {
            # 使用中などで削除できなければ、次回に回す
        }
    }
}

function findDroppedIndexes {
    # 前回の取り込み一覧にあって、今回のクロール対象フォルダに無いインデックス（設定から外れたもの）を返す。見つけるだけで、消さない
    # （消すかどうかは利用者が確認ダイアログで決める。removeDroppedFolders）。チェックを外しただけのフォルダは、今回の一覧にあるため含まない
    param (
        [object[]]$folders,          # assignIndexNames の結果
        [object[]]$previousFolders   # readStatusFile の Folders
    )

    # インデックス名で比べる。フォルダを移動して登録し直した場合は、同じ名前を引き継ぐため含めない（assignIndexNames）
    $current = @($folders | ForEach-Object { $_.Name })
    return @($previousFolders | Where-Object { $_ -and $_.Name -and $current -notcontains $_.Name })
}

function getDroppedStatusRows {
    # 消さずに残す外れたインデックスの、前回の取り込み一覧の行（次の回にまた見つけられるよう、書き直す一覧に残す）
    param (
        [object[]]$dropped,
        $previousRows  # readStatusFile の Rows（相対パス → 行）
    )

    $result = New-Object System.Collections.Generic.List[object]
    if ($null -eq $previousRows) {
        return $result.ToArray()
    }
    foreach ($item in @($dropped | Where-Object { $_ })) {
        $prefix = "$($item.Name)\"
        foreach ($key in @($previousRows.Keys)) {
            if ($key.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $result.Add($previousRows[$key])
            }
        }
    }
    return $result.ToArray()
}

function writeStatusKeepingDropped {
    # 取り込み一覧を書く。消さずに残す外れたインデックスのフォルダの行と各行も、前回のまま足して書く（次の回にまた見つけられるように）
    param (
        $ledger,
        [object[]]$folders,      # 今回のクロール対象フォルダ
        [object[]]$rows,         # 今回の各行
        [object[]]$keptDropped,  # 消さずに残す外れたインデックス
        $previousRows            # readStatusFile の Rows（前回の行）
    )

    $ledger.WriteStatus((@($folders) + @($keptDropped)), (@($rows) + @(getDroppedStatusRows $keptDropped $previousRows)))
}

function removeDroppedFolders {
    # 設定から外れたインデックス（findDroppedIndexes の結果。work\content_index\<インデックス名>）を削除する。
    # 利用者が確認ダイアログで［更新を開始］を押した回でだけ呼ぶ
    param (
        [object[]]$dropped,
        $ws = $workspace             # 対象のワークスペース（テストが差し替える）
    )

    foreach ($previous in @($dropped | Where-Object { $_ })) {
        $dir = Join-Path $ws.IndexDir $previous.Name
        if (Test-Path -LiteralPath $dir) {
            # 中に長いパス（260文字超）のTSVがあっても削除でき、中のリンク（ジャンクション・シンボリックリンク）の先は消さない
            removeDirectoryRetry $dir
        }
        writeIndexerLog "クロール対象から削除されたフォルダ（$($previous.Path)）のインデックスを削除しました。"
    }
}
