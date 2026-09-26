# ワークスペース（インデックス・取り込み一覧・ログを置くフォルダ）の中の、tebunko が使うファイルの場所。
# 場所はフォルダ（Dir）から組み立てるだけで、設定は読まない。どのフォルダを使うかは設定の workspaceFolder で決まる（settings.ps1 の getWorkDir）。
# 今のワークスペースは paths.ps1 の ${workspace} に置く。別のスレッドへは Dir（文字列）を渡し、そこで作り直す。
# クラスのメソッドからはスクリプトの関数・変数が見えないため、ここでは .NET の型だけを使う

class Workspace {
    [string]$Dir
    [string]$IndexDir
    # システムインデックス（本文インデックスの 2-gram を書いた txt。index と同じ相対パスの構成。Windows Search に索引させる）と、その状態
    [string]$SystemIndexDir
    [string]$SystemIndexStateFile
    # 取り込んだTSVをインデックスに入れる直前に集めるフォルダ（publishIndexFiles）。
    # フォルダごと入れ替えるため、インデックスと同じドライブ（ワークスペースの中）に置く。
    # 検索対象に入らないよう index の外にする。インデックス作成を同時に複数実行しても混ざらないよう、プロセスごとに分ける
    [string]$PublishDir
    # 取り込み一覧・出力（自動生成）
    [string]$StatusFile
    [string]$IngestingFile  # 取り込み中のファイル。強制終了で残っていれば、そのファイルの取り込み中に止まった
    [string]$ResultFile
    # インデックス作成の記録。画面とインデクサの受け渡しはメモリ上で行う（indexer_state.ps1 の newIndexerChannel）
    [string]$IndexingLogFile  # インデクサの表示内容の記録（実行ごとに上書き）
    [string]$GuiErrorLogFile  # 画面で起きた予期しないエラーの記録（追記。原因を後から追えるようにする）

    Workspace([string]$dir) {
        $this.Dir = $dir
        $this.IndexDir = "$dir\index"
        $this.SystemIndexDir = "$dir\system_index"
        $this.SystemIndexStateFile = "$dir\システムインデックスの状態.tsv"
        $this.PublishDir = "$dir\取り込み出力\$([System.Diagnostics.Process]::GetCurrentProcess().Id)"
        $this.StatusFile = "$dir\取り込み一覧.tsv"
        $this.IngestingFile = "$dir\取り込み中.txt"
        $this.ResultFile = "$dir\検索結果.txt"
        $this.IndexingLogFile = "$dir\インデックス作成ログ.txt"
        $this.GuiErrorLogFile = "$dir\画面エラー.txt"
    }

    # ワークスペースを移すときに移すもの（tebunko が作るファイル・フォルダ）。利用者のほかのファイルは含めない。
    # 取り込み出力はプロセスごとのフォルダ（PublishDir）の親を移す
    [string[]] Entries() {
        return @($this.IndexDir, $this.SystemIndexDir, $this.SystemIndexStateFile, $this.StatusFile, $this.IngestingFile,
            $this.ResultFile, $this.IndexingLogFile, $this.GuiErrorLogFile, [System.IO.Path]::GetDirectoryName($this.PublishDir))
    }
}

function getWorkspaceEntries {
    # ワークスペース dir にある、tebunko のファイル・フォルダ（Workspace.Entries のうち、あるもの）のフルパスを返す
    param (
        [string]$dir
    )

    return @([Workspace]::new($dir).Entries() | Where-Object {
        $long = toLongPath $_
        [System.IO.Directory]::Exists($long) -or [System.IO.File]::Exists($long)
    })
}

function getWorkspaceMoveConflicts {
    # ワークスペース from の中身を to へ移すとき、to に同じ名前が既にあるもの（移せないもの）の名前を返す
    param (
        [string]$from,
        [string]$to
    )

    return @(getWorkspaceEntries $from | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Where-Object {
        $long = toLongPath (Join-Path $to $_)
        [System.IO.Directory]::Exists($long) -or [System.IO.File]::Exists($long)
    })
}

function moveWorkspaceEntry {
    # ファイル・フォルダを 1 つ移す。同じドライブならそのまま移し、別のドライブのフォルダは写してから元を消す
    # （Directory.Move は別のドライブへ移せないため）。写している途中で失敗したら、写した分を消して例外にする
    param (
        [string]$source,
        [string]$dest
    )

    $longSource = toLongPath $source
    $longDest = toLongPath $dest
    if (![System.IO.Directory]::Exists($longSource)) {
        # File.Move は別のドライブへも移せる
        [System.IO.File]::Move($longSource, $longDest)
        return
    }
    if ([System.IO.Path]::GetPathRoot($source) -eq [System.IO.Path]::GetPathRoot($dest)) {
        [System.IO.Directory]::Move($longSource, $longDest)
        return
    }
    try {
        copyDirectoryTree $longSource $longDest
    } catch {
        removeDirectoryRetry $dest
        throw
    }
    removeDirectoryRetry $source
}

function copyDirectoryTree {
    # フォルダを中身ごと写す（\\?\ 付きのパスを受け取る）
    param (
        [string]$source,
        [string]$dest
    )

    [System.IO.Directory]::CreateDirectory($dest) | Out-Null
    foreach ($file in [System.IO.Directory]::GetFiles($source)) {
        [System.IO.File]::Copy($file, [System.IO.Path]::Combine($dest, [System.IO.Path]::GetFileName($file)))
    }
    foreach ($sub in [System.IO.Directory]::GetDirectories($source)) {
        copyDirectoryTree $sub ([System.IO.Path]::Combine($dest, [System.IO.Path]::GetFileName($sub)))
    }
}

function removeWorkspaceEntries {
    # ワークスペース dir の tebunko のファイル・フォルダ（Workspace.Entries）を削除し、削除した数を返す。ほかのファイルは消さない。
    # ワークスペースに選んだフォルダのインデックスを使わず、消して最初からやり直すときに使う
    param (
        [string]$dir
    )

    $entries = @(getWorkspaceEntries $dir)
    foreach ($entry in $entries) {
        $long = toLongPath $entry
        if ([System.IO.Directory]::Exists($long)) {
            removeDirectoryRetry $entry
        } else {
            [System.IO.File]::Delete($long)
        }
    }
    return $entries.Count
}

function useWorkspaceTargets {
    # ワークスペース dir の取り込み一覧にあるクロール対象フォルダを、インデックスの一覧（設定の targetFolders）にし、その数を返す。
    # ほかの人が作ったワークスペースを使うとき、一覧をそのワークスペースに合わせる（合わせないままインデックス作成をすると、
    # 一覧に無いインデックスは削除されたフォルダのものとして消える。removeDroppedFolders）。取り込み一覧が無ければ一覧は変えない
    param (
        [string]$dir,
        [string]$path = ${settingsFile}
    )

    $map = getIndexNameMap ([Workspace]::new($dir.TrimEnd("\")).StatusFile)
    if ($map.Count -eq 0) {
        return 0
    }
    writeTargetFolders @($map.Keys | ForEach-Object { [pscustomobject]@{ Name = $_; Path = $map[$_]; Enabled = $true } }) $path
    return $map.Count
}

function moveSearchExcludes {
    # 検索対象ツリーでチェックを外したフォルダ（searchExcludes。インデックスの下のフルパス）のうち、
    # ワークスペース from の index の下のものを、to の index の下に付け替えて保存する。付け替えた数を返す
    param (
        [string]$from,
        [string]$to,
        [string]$path = ${settingsFile}
    )

    $fromIndex = [Workspace]::new($from.TrimEnd("\")).IndexDir
    $toIndex = [Workspace]::new($to.TrimEnd("\")).IndexDir
    return invokeSettingsLocked $path {
        $count = 0
        $excludes = @(readSearchExcludes $path | ForEach-Object {
            $folder = $_.Path
            if ($folder.Equals($fromIndex, [System.StringComparison]::OrdinalIgnoreCase) -or
                $folder.StartsWith("$fromIndex\", [System.StringComparison]::OrdinalIgnoreCase)) {
                $folder = $toIndex + $folder.Substring($fromIndex.Length)
                $count++
            }
            [pscustomobject]@{ Path = $folder; Subfolders = $_.Subfolders }
        })
        if ($count -gt 0) {
            writeSearchExcludes $excludes $path
        }
        return $count
    }
}

function moveWorkspace {
    # ワークスペース from の中身（tebunko のファイル・フォルダ）を to へ移し、移した数を返す。ほかのファイルは移さない。
    # to に同じ名前があれば、何も移さずに例外にする。途中で移せなかったら、移した分を from へ戻してから例外にする
    # （どちらかのワークスペースに中身がそろった状態にし、半分ずつに分かれたままにしない）
    param (
        [string]$from,
        [string]$to
    )

    $conflicts = @(getWorkspaceMoveConflicts $from $to)
    if ($conflicts.Count -gt 0) {
        throw "「${to}」には、すでに $($conflicts -join '、') があります。空のフォルダを選んでください。"
    }
    $entries = @(getWorkspaceEntries $from)
    [System.IO.Directory]::CreateDirectory((toLongPath $to)) | Out-Null
    $moved = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($source in $entries) {
            moveWorkspaceEntry $source (Join-Path $to ([System.IO.Path]::GetFileName($source)))
            $moved.Add($source)
        }
    } catch {
        $reason = $_.Exception.Message
        for ($i = $moved.Count - 1; $i -ge 0; $i--) {
            try {
                moveWorkspaceEntry (Join-Path $to ([System.IO.Path]::GetFileName($moved[$i]))) $moved[$i]
            } catch {
                # 戻せなかったものは、移した先に残る（例外のメッセージで両方の場所を伝える）
            }
        }
        throw "ワークスペースの中身を「${to}」へ移せませんでした（${reason}）。ファイルを開いているアプリを閉じてから、もう一度変えてください。中身は「${from}」に残しています。"
    }
    return $moved.Count
}
