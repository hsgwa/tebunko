# TSV のインデックスへの取り込みと、作業フォルダ・外したフォルダのインデックスの後始末。

function publishTsv {
    # 作業フォルダのTSVを、そのファイルのインデックスのフォルダへ移動する。
    # 途中で強制終了されても一部のシートだけのインデックスが残らないよう、
    # 出力用のフォルダ（work\publish\<PID>）に集めてからフォルダごと入れ替える（publishIndexFiles）。
    # 移す前に、この元のファイルだけの本文インデックスのファイルが検索の上限を超えないかを確かめ、超えるなら例外にする（取り込みの失敗として取り込み一覧に残る。
    # 超えたまま置くと、同じ本文インデックスのファイルの小さな元のファイルまで検索に出なくなる）
    param (
        [string]$bookDir
    )

    assertContentIndexBookFits $tmpDir ([System.IO.Path]::GetFileName($bookDir))
    publishIndexFiles $tmpDir $bookDir (Join-Path $workspace.PublishDir ([System.IO.Path]::GetFileName($bookDir))) $workspace.IndexDir $workspace.PublishDir
}

function clearTmpDir {
    # 作業フォルダ直下のファイルを消す。作業フォルダそのものがリンク（ジャンクション・シンボリックリンク）のときは、
    # 先のファイルを消さないよう、何も消さずにログに出す
    if (testPathIsLink $tmpDir) {
        writeIndexerLog "    作業フォルダがリンク（ジャンクション・シンボリックリンク）のため、中のファイルを消しませんでした: ${tmpDir}" "Yellow"
        return
    }
    # Remove-Item でなく .NET で消す（消す入口を .NET にそろえる。読み取り専用の属性は外して消す）
    foreach ($file in @([System.IO.Directory]::EnumerateFiles((toLongPath $tmpDir)))) {
        removeEntryOnly (New-Object System.IO.FileInfo $file)
    }
}

function removeEmptyDir {
    # フォルダが空のときだけ、再帰しない削除で消す。中にほかのプロセスのフォルダなどが残っている・
    # フォルダが既に無いときは、何もせずそのままにする。
    # リンク（ジャンクション・シンボリックリンク）は、リンク自体を外さないよう、先が空でも消さない
    # （利用者が tmp や出力用のフォルダの親をリンクにした構成を、そのまま使えるようにするため）
    param (
        [string]$dir
    )
    if (testPathIsLink $dir) {
        return
    }
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
    # 消すのは、作業フォルダは tmp の下、出力用のフォルダは publish の下で、途中にリンクが無いときだけ（removeDirectoryUnder）
    $roots = @{ }
    if (${tmpDir}) { $roots[${tmpDir}] = $workspace.TmpRoot }
    $roots[$workspace.PublishDir] = Split-Path $workspace.PublishDir -Parent
    foreach ($dir in @(${tmpDir}, $workspace.PublishDir) | Where-Object { $_ }) {
        try {
            if (!(removeDirectoryUnder $dir $roots[$dir])) {
                writeIndexerLog "    作業フォルダを削除できませんでした: ${dir}（置き場の外か、途中にリンクがある）" "Yellow"
            }
        } catch {
            writeIndexerLog "    作業フォルダを削除できませんでした: ${dir}（$($_.Exception.Message)）" "Yellow"
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
    # work\publish\<PID> は PC ごとに分かれていない（PID が自分の PC で動いていなければ消す）。共有フォルダのワークスペースで
    # ほかの PC が取り込み中のものを消さないのは、この関数がインデックス作成の鍵（enterIndexerLock。別の PC とも排他）の中でだけ呼ばれるため
    # 途中にリンク（ジャンクション・シンボリックリンク）があるときは、先のフォルダを消さないよう、片付けない
    # （removeStaleProcessDirs が、親ごとの根の下で途中にリンクが無いことを確かめる）
    $keyDir = Split-Path (getWorkspaceTmpDir $workspace) -Parent
    # 出力用のフォルダは <ワークスペース>\publish\<PID>。親（publish）の根はワークスペース（その 1 つ上）
    $publishParent = Split-Path $workspace.PublishDir -Parent
    $parents = @(
        @{ Parent = $keyDir; Root = $workspace.TmpRoot },
        @{ Parent = $publishParent; Root = (Split-Path $publishParent -Parent) },
        @{ Parent = ${legacyTmpParent}; Root = (Split-Path ${legacyTmpParent} -Parent) }
    )
    foreach ($entry in $parents) {
        removeStaleProcessDirs $entry.Parent $entry.Root
        removeEmptyDir $entry.Parent
    }
    removeEmptyDir $workspace.TmpRoot
}

function newDirectoryReplacingExisting {
    # フォルダを作る。同じ名前（前の残り・先に置かれたリンク）が既にあれば、リンクは辿らずに消してから作り直す。
    # 作ったものがリンクでなければ $true（置けた）、リンクのまま残った（消せなかった）ときは $false
    # 消す・作るのは、root の下で、途中にリンクが無いときだけ（removeDirectoryUnder・newDirectoryUnder）。偽なら何もせず $false
    param (
        [string]$dir,
        [string]$root
    )

    try {
        if (!(removeDirectoryUnder $dir $root)) {
            return $false
        }
    } catch {
        # 消せなくても、下の確かめでリンクが残っていなければ、そのまま使う
    }
    if (!(newDirectoryUnder $dir $root)) {
        return $false
    }
    return !(testPathIsLink $dir)
}

function initTmpDir {
    # インデックス作成の始めに、取り込みの作業フォルダの片付けと用意をまとめて行う: @{ Dir; Reason（置けなかったときだけ） }
    #   1. 強制終了などで残った、ほかの（終了済みの）プロセスの作業フォルダを片付ける（決めた場所を巻き込まないよう、場所を決める前に行う）
    #   2. selectTmpDir で置き場所を決める
    #   3. 置けたら、フォルダを作る
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
    $placed = $false
    if (!(testPathHasLink $workspace.TmpRoot (Split-Path $selected.Dir -Parent))) {
        $placed = newDirectoryReplacingExisting $selected.Dir $workspace.TmpRoot
    }
    if (!$placed) {
        $unavailable = @{ Dir = ""; Reason = "Link" }
        writeIndexerLog "$(getTmpDirUnavailableMessage $unavailable.Reason)取り込みをすべてスキップします。" "Yellow"
        return $unavailable
    }
    return $selected
}

function newWorkerTmpDir {
    # 取り込みのスレッドの作業フォルダ（parentDir\w<番号>）を作り、パスを返す。
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
    if (!(newDirectoryReplacingExisting $dir $workspace.TmpRoot)) {
        return ""
    }
    return $dir
}

function removeStaleProcessDirs {
    # プロセスIDの名前のフォルダのうち、そのプロセスが既に終わっているものを削除する。
    # 消すのは、root の下で、途中（parent を含む）にリンクが無いものだけ（removeDirectoryUnder）。そうでなければ黄色のログを出して残す
    param (
        [string]$parent,
        [string]$root
    )

    if (![System.IO.Directory]::Exists((toLongPath $parent))) {
        return
    }
    foreach ($dir in @(Get-ChildItem -LiteralPath $parent -Directory | Where-Object { $_.Name -match "^\d{1,9}$" -and [int]$_.Name -ne $PID })) {
        if (Get-Process -Id ([int]$dir.Name) -ErrorAction SilentlyContinue) {
            continue  # 実行中のインデックス作成（またはPIDを再利用した別のプロセス）のものは残す
        }
        try {
            # 中に長いパスがあっても消え、中のリンクの先は消さない。使用中などで消せなければ 1 回だけ試して次回に回す
            if (!(removeDirectoryUnder $dir.FullName $root 1)) {
                writeIndexerLog "    終了したプロセスの作業フォルダを削除できませんでした: $($dir.FullName)（置き場の外か、途中にリンクがある）" "Yellow"
            }
        } catch {
            # 使用中・リンクかどうか確かめられないフォルダがあるなどで削除できなければ、原因をログに出して次回に回す
            writeIndexerLog "    終了したプロセスの作業フォルダを削除できませんでした: $($dir.FullName)（$($_.Exception.Message)）" "Yellow"
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

    # 取り込み一覧の名前は共有のワークスペースで書き換えられうる。名前でパスを組んで消さず、置き場（IndexDir）の直下に
    # 実在するフォルダを列挙して、名前が同じものを消す（最後の . や空白のある名前は、パスにすると別のフォルダを指すため。
    # 置き場の外を指す名前は、列挙したものに当たらず、何も消えない（実在の aux.backup・CON も、名前が合えば消す）。例外にすると、書き換えられた 1 行で
    # 毎回インデックス作成が止まるため、ログに出して続ける）
    # 置き場の有無は getPathState で見る（Directory.Exists は、届かない・読めないときも偽になり、どの名前も「既にありません」として記録が消えるため）。
    # 無い、またはフォルダでない（ファイルがある）と確かめられたときだけ空として続け、確かめられなければ例外にする
    $existing = @()
    $indexDirState = getPathState $ws.IndexDir
    if ($indexDirState.State -eq ${pathStateFound}) {
        if ($indexDirState.IsDirectory) {
            $existing = @([System.IO.Directory]::GetDirectories((toLongPath $ws.IndexDir)))
        }
    } elseif ($indexDirState.State -ne ${pathStateMissing}) {
        throw "インデックスの置き場があるかどうか確かめられません: $($ws.IndexDir)（$($indexDirState.Message)）"
    }
    foreach ($previous in @($dropped | Where-Object { $_ })) {
        $name = [string]$previous.Name
        $found = @($existing | Where-Object { [string]::Equals([System.IO.Path]::GetFileName($_), $name, [System.StringComparison]::OrdinalIgnoreCase) })
        if ($found.Count -eq 0) {
            if ((testPathSegments @($name) $true) -ne "") {
                writeIndexerLog "クロール対象から削除されたフォルダ（$($previous.Path)）のインデックス名が使えないため、削除しませんでした。" "Yellow"
            } else {
                writeIndexerLog "クロール対象から削除されたフォルダ（$($previous.Path)）のインデックスのフォルダは、既にありません。"
            }
            continue
        }
        foreach ($dir in $found) {
            # 中に長いパス（260文字超）のTSVがあっても削除でき、中のリンク（ジャンクション・シンボリックリンク）の先は消さない。
            # 消えていなければ例外になる（下のログは、消えたことを確かめてから出る）
            if (!(removeDirectoryUnder $dir $ws.IndexDir)) {
                throw "置き場の外か、途中にリンクがあるため、消せません: ${dir}"
            }
        }
        writeIndexerLog "クロール対象から削除されたフォルダ（$($previous.Path)）のインデックスを削除しました。"
    }
}
