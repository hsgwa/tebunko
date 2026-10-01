# ワークスペース（インデックス・取り込み一覧・ログを置くフォルダ）の中の、tebunko が使うファイルの場所。
# 場所はフォルダ（Dir）から組み立てるだけで、設定は読まない。どのフォルダを使うかは設定の workspaceFolder で決まる（settings.ps1 の getWorkDir）。
# 今のワークスペースは paths.ps1 の ${workspace} に置く。別のスレッドへは Dir（文字列）を渡し、そこで作り直す。
# クラスのメソッドからはスクリプトの関数・変数が見えないため、ここでは .NET の型だけを使う

class Workspace {
    [string]$Dir
    [string]$IndexDir
    # 前の版（content_index・system_index に名前をそろえる前）が使っていた本文インデックスのフォルダ。
    # 新しい版はここを読まず、消しもしない（前の版のしるしの調べ・知らせ・［8 設定］で「移す」「消して最初から」の対象にするために持つ）
    [string]$LegacyIndexDir
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
    # 取り込みの作業フォルダの置き場所（下は <PC の鍵>\<PID>\w<番号> と分かれる。selectTmpDir・getWorkspaceTmpDir）。
    # 前の版までの %TEMP%\tebunko\<PID> の代わりに、ワークスペースの中に置く（共有フォルダでも 1 か所にまとまる）
    [string]$TmpRoot

    Workspace([string]$dir) {
        $this.Dir = $dir
        $this.IndexDir = "$dir\content_index"
        $this.LegacyIndexDir = "$dir\index"
        $this.SystemIndexDir = "$dir\system_index"
        $this.SystemIndexStateFile = "$dir\システムインデックスの状態.tsv"
        $this.PublishDir = "$dir\取り込み出力\$([System.Diagnostics.Process]::GetCurrentProcess().Id)"
        $this.StatusFile = "$dir\取り込み一覧.tsv"
        $this.IngestingFile = "$dir\取り込み中.txt"
        $this.ResultFile = "$dir\検索結果.txt"
        $this.IndexingLogFile = "$dir\インデックス作成ログ.txt"
        $this.GuiErrorLogFile = "$dir\画面エラー.txt"
        $this.TmpRoot = "$dir\tmp"
    }

    # ワークスペースを移すときに移すもの（tebunko が作るファイル・フォルダ）。利用者のほかのファイルは含めない。
    # 取り込み出力はプロセスごとのフォルダ（PublishDir）の親を移す。
    # LegacyIndexDir は、あれば前の版のワークスペースとして tebunko のものと扱う（あるものだけが getWorkspaceEntries で拾われる）
    [string[]] Entries() {
        return @($this.IndexDir, $this.LegacyIndexDir, $this.SystemIndexDir, $this.SystemIndexStateFile, $this.StatusFile, $this.IngestingFile,
            $this.ResultFile, $this.IndexingLogFile, $this.GuiErrorLogFile, [System.IO.Path]::GetDirectoryName($this.PublishDir), $this.TmpRoot)
    }
}

# 前の版（content_index・system_index に名前をそろえる前）が system_index に置いていた txt の名前の型。
# 新しい版は system_index.txt を使うため、前の版の txt を見つける・片付けるためだけに使う（読み書きの形式は変えない）
${legacySystemIndexPattern} = "システムインデックス*.txt"

function getMachineKey {
    # この PC を識別する短い鍵（getFolderKey の先頭 8 文字）。
    # 共有フォルダのワークスペースを複数の PC から使うとき、一時フォルダを PC ごとに分けるために使う（getWorkspaceTmpDir）
    return (getFolderKey ([Environment]::MachineName)).Substring(0, 8)
}

function getWorkspaceTmpDir {
    # 取り込みの作業フォルダの候補（$workspace.TmpRoot\<PC の鍵>\<PID>）を返す。副作用は無い（selectTmpDir が使う）
    param (
        [Workspace]$workspace
    )

    return Join-Path $workspace.TmpRoot (Join-Path (getMachineKey) $PID)
}

function selectTmpDir {
    # 取り込みの作業フォルダの置き場所を決める: @{ Dir; Reason（置けないときだけ） }
    #   候補: getWorkspaceTmpDir（ワークスペースの tmp\ の下）
    #   候補のパスに [ ] があれば、置けない（Dir = ""、Reason = Brackets）。Excel は [ ] を含むパスに保存できないため
    #   候補の長さ + 取り込みのスレッドが下に作る最も長い名前の分（${tmpNameReserve}）が $excelMaxPath 以上でも、
    #     置けない（Dir = ""、Reason = TooLong）。Office は長すぎるパスを開けないため
    #   一時ファイルもワークスペースの下にしか置かない（%TEMP% には逃がさない）。
    #   置けないときは、呼び出し側（initTmpDir）が取り込みをすべてスキップする
    #   （どのファイルも中間 TSV などをこのフォルダに作るため、テキストファイルを含めすべての取り込みが対象になる）
    #   どちらでもなければ候補のまま（Reason = ""）
    param (
        [Workspace]$workspace
    )

    $candidate = getWorkspaceTmpDir $workspace
    if ($candidate.IndexOfAny([char[]]@("[", "]")) -ge 0) {
        return @{ Dir = ""; Reason = "Brackets" }
    }
    if (($candidate.Length + ${tmpNameReserve}) -ge $excelMaxPath) {
        return @{ Dir = ""; Reason = "TooLong" }
    }
    return @{ Dir = $candidate; Reason = "" }
}

function getTmpDirUnavailableMessage {
    # 取り込みの作業フォルダを置けない理由（selectTmpDir の Reason）を、利用者向けの1文にする。
    # initTmpDir のログと、invokeIngestTask が1ファイルごとに書くエラーの両方で使う
    param (
        [string]$reason
    )

    if ($reason -eq "Brackets") {
        return "ワークスペースのパスに [ ]（角かっこ）が含まれるため、取り込みの作業フォルダを置けません。"
    }
    if ($reason -eq "TooLong") {
        return "ワークスペースのパスが長すぎるため、取り込みの作業フォルダを置けません。"
    }
    return "取り込みの作業フォルダを置けません。"
}

function getLegacyIndexState {
    # 前の版のワークスペース（dir）の状態を返す: @{ HasLegacyIndex; ContentEmpty; HasLegacySystemIndex }
    #   HasLegacyIndex       : 前の版の index\ があり、直下のどれかのフォルダに 元のフォルダ.txt がある（前の版のしるし）。
    #                          index\ があるだけでは、利用者が選んだフォルダにたまたま index があるときと区別できないため、しるしとしない
    #   ContentEmpty         : content_index\ が無いか、下のフォルダを含めてファイルが 1 つも無い
    #   HasLegacySystemIndex : system_index\ の下に前の名前の txt（システムインデックス*.txt）がある。
    #                          ContentEmpty のときだけ調べる（片付けの条件にしか使わないため）
    param (
        [string]$dir
    )

    $ws = [Workspace]::new($dir)
    # 見つけたところで列挙をやめるため、testAnyEntry で調べる（foreach を break で抜けると、調べていたフォルダを
    # 掴んだまま残り、続くインポートの上書きで content_index\<名前> を移動できなくなる）
    $hasLegacyIndex = $false
    $longLegacy = toLongPath $ws.LegacyIndexDir
    if ([System.IO.Directory]::Exists($longLegacy)) {
        $hasLegacyIndex = testAnyEntry ([System.IO.Directory]::EnumerateDirectories($longLegacy)) {
            param ($sub)
            [System.IO.File]::Exists("$sub\${sourceFolderFileName}")
        }
    }

    $contentEmpty = $true
    $longContent = toLongPath $ws.IndexDir
    if ([System.IO.Directory]::Exists($longContent)) {
        $contentEmpty = !(testAnyEntry ([System.IO.Directory]::EnumerateFiles($longContent, "*", [System.IO.SearchOption]::AllDirectories)))
    }

    $hasLegacySystemIndex = $false
    if ($contentEmpty) {
        $longSystem = toLongPath $ws.SystemIndexDir
        if ([System.IO.Directory]::Exists($longSystem)) {
            $hasLegacySystemIndex = testAnyEntry ([System.IO.Directory]::EnumerateFiles($longSystem, ${legacySystemIndexPattern}, [System.IO.SearchOption]::AllDirectories))
        }
    }

    return @{ HasLegacyIndex = $hasLegacyIndex; ContentEmpty = $contentEmpty; HasLegacySystemIndex = $hasLegacySystemIndex }
}

function testLegacyCleanupNeeded {
    # 取り込み直しを始めるときに、前の版のシステムインデックスを片付けるかどうか（getLegacyIndexState の結果から決まる）。
    # content_index\ が空で、かつ、前の版のしるしがあるか、system_index\ に前の名前の txt が残っているときに片付ける
    # （取り込み直す前に利用者が index\ を消していても、前の名前の txt を残さないため）
    param (
        $state
    )

    return ($state.ContentEmpty -and ($state.HasLegacyIndex -or $state.HasLegacySystemIndex))
}

function getLegacyIndexMessage {
    # 前の版のワークスペース（index\。しるしあり）が見つかったときの知らせ。しるしが無ければ空
    param (
        [string]$dir,
        [bool]$hasLegacyIndex
    )

    if (!$hasLegacyIndex) {
        return ""
    }
    $legacyDir = [Workspace]::new($dir).LegacyIndexDir
    return "前の版のインデックス（「${legacyDir}」）は、この版では使えません。インデックス作成で、元のファイルをすべて取り込み直します。" +
        "取り込み直した後、「${legacyDir}」フォルダは削除してかまいません。"
}

function clearLegacySystemIndex {
    # 前の版のシステムインデックス（system_index\ と、システムインデックスの状態ファイルの中身）を片付ける: @{ Ok; Reason（失敗のときだけ） }
    # 片付けの順番: (1) updateSystemIndexState の排他の中で、状態ファイルの中身を空にする（ファイルは消さない。
    #     画面の検索（fast_search.ps1）やインデックスの削除（index_store.ps1）が同じ排他で書き換えるため、
    #     消した直後に前のキーを書き戻されないようにする）。
    # (2) system_index\ を removeDirectoryRetry で消す（Windows Search が txt を一時的に開くことがあるため）。
    # どちらかに失敗したら Ok = $false（途中で止まっても、次に呼べば同じ状態から続けられる）
    param (
        [string]$dir
    )

    $ws = [Workspace]::new($dir)
    $cleared = updateSystemIndexState {
        param ($state)
        $state.Covered.Clear()
        $state.Pending.Clear()
        $state.Excluded.Clear()
    } $ws.SystemIndexStateFile
    if (!$cleared) {
        return @{ Ok = $false; Reason = "システムインデックスの状態ファイル（$([System.IO.Path]::GetFileName($ws.SystemIndexStateFile))）を開けませんでした" }
    }
    try {
        removeDirectoryRetry $ws.SystemIndexDir
    } catch {
        return @{ Ok = $false; Reason = $_.Exception.Message }
    }
    return @{ Ok = $true; Reason = "" }
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
    return invokeSettingsLocked -path $path -action {
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
