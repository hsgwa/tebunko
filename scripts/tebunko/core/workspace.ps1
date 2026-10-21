# ワークスペース（インデックス・取り込み一覧・ログを置くフォルダ）の中の、tebunko が使うファイルの場所。
# 場所はフォルダ（Dir）から組み立てるだけで、設定は読まない。どのフォルダを使うかは設定の workspaceFolder で決まる（settings.ps1 の getWorkDir）。
# 今のワークスペースは paths.ps1 の ${workspace} に置く。別のスレッドへは Dir（文字列）を渡し、そこで作り直す。
# Workspace のメソッドは場所を組み立てるだけなので、.NET の型だけを使う（クラスのメソッドからスクリプトの関数は呼べるが、
# $script: 変数は呼び出し元によって見える範囲が違うため、クラスの本体では読まない。docs/design/structure/classes.md）

class Workspace {
    [string]$Dir
    # インデックスの置き場（版番号つきの目録 FormatFile と、本文インデックス IndexDir・取り込み一覧 StatusFile をまとめる。store\）。
    # 後の版が足す置き場（gram\・journal\）も、この中に置く
    [string]$StoreDir
    [string]$FormatFile
    [string]$IndexDir
    # gram 索引（本文インデックスのファイルごとの 2-gram の索引。インデックス名ごとのフォルダ。検索で、検索語を含まないファイルを読まずに外す）
    [string]$GramDir
    # 前の版（版番号つきの置き場に替える前）の本文インデックスのフォルダと取り込み一覧。
    # 新しい版は読むだけで書かない。作り直しが最後まで終わったインデックスから消す（index_format_store.ps1）
    [string]$OldIndexDir
    [string]$OldStatusFile
    # 前の版（content_index・system_index に名前をそろえる前）が使っていた本文インデックスのフォルダ。
    # 新しい版はここを読まず、消しもしない（前の版のしるしの調べ・知らせ・［設定］で「移す」「消して最初から」の対象にするために持つ）
    [string]$LegacyIndexDir
    # 前の版が作ったシステムインデックス（本文インデックスの 2-gram を書いた txt）と、その状態ファイル。
    # 新しい版は読まず、使わない。インデックス作成の始めに removeOldSystemIndex が消す（ワークスペースを移す・書き出す・消す対象には含めない）
    [string]$OldSystemIndexDir
    [string]$OldSystemIndexStateFile
    # 取り込んだTSVをインデックスに入れる直前に集めるフォルダ（publishIndexFiles）。
    # フォルダごと入れ替えるため、インデックスと同じドライブ（ワークスペースの中）に置く。
    # 検索対象に入らないよう index の外にする。インデックス作成を同時に複数実行しても混ざらないよう、プロセスごとに分ける
    [string]$PublishDir
    # 取り込み一覧・出力（自動生成）
    [string]$StatusFile
    # 取り込みのジャーナル（どの仕事を始め・どのフォルダを書き出したかの追記ログ。強制終了からの復旧に使う。store\journal\。ingest_journal.ps1）
    [string]$JournalFile
    # 本文インデックスから外し終えていない、無くなった元のファイルの記録（書き出しを飛ばした・落ちた・強制終了した回の次の回に外し直す）
    [string]$PendingRemoveFile
    [string]$ResultFile
    # インデックス作成の記録。画面とインデクサの受け渡しはメモリ上で行う（indexer_state.ps1 の newIndexerChannel）
    [string]$IndexingLogFile  # インデクサの表示内容の記録（実行ごとに上書き）
    [string]$GuiErrorLogFile  # 画面で起きた予期しないエラーの記録（追記。原因を後から追えるようにする。1 MB で下の前回の分に回す）
    [string]$IndexingLogPreviousFile  # 前回のインデックス作成の記録（インデックス作成を始めるとき、今の記録をここへ移す）
    [string]$GuiErrorLogPreviousFile  # 画面エラーの記録の前回の分（1 MB を超えたとき、今の記録をここへ移す）
    # 取り込みの作業フォルダの置き場所（下は <PC の鍵>\<PID>\w<番号> と分かれる。selectTmpDir・getWorkspaceTmpDir）。
    # 前の版までの %TEMP%\tebunko\<PID> の代わりに、ワークスペースの中に置く（共有フォルダでも 1 か所にまとまる）
    [string]$TmpRoot
    # このツールが起動した Office の PID の記録の置き場所（下は <PC の鍵>\<PID>.txt と分かれる。getOfficePidDir・shared/office/office_process.ps1）。
    # 画面を起動したとき、前回残った Office を確認して止めるために使う。tmp の下には置かない（tmp は片付けで消える場所で、記録は画面ごと落ちた後に読むため）
    [string]$OfficePidRoot
    # 起動の段階の記録（環境変数 TEBUNKO_STARTUP_TRACE が 1 のときだけ、画面の起動の終わりに書く。ui\startup_trace.ps1。ローカルのワークスペースだけ）
    [string]$StartupTraceFile

    Workspace([string]$dir) {
        $this.Dir = $dir
        $this.StoreDir = "$dir\store"
        $this.FormatFile = "$dir\store\format.tsv"
        $this.IndexDir = "$dir\store\content"
        $this.GramDir = "$dir\store\gram"
        $this.OldIndexDir = "$dir\content_index"
        $this.OldStatusFile = "$dir\ingest_status.tsv"
        $this.LegacyIndexDir = "$dir\index"
        $this.OldSystemIndexDir = "$dir\system_index"
        $this.OldSystemIndexStateFile = "$dir\system_index_state.tsv"
        $this.PublishDir = "$dir\publish\$([System.Diagnostics.Process]::GetCurrentProcess().Id)"
        $this.StatusFile = "$dir\store\ingest_status.tsv"
        $this.JournalFile = "$dir\store\journal\ingest.tsv"
        $this.PendingRemoveFile = "$dir\store\pending_remove.tsv"
        $this.ResultFile = "$dir\search_results.txt"
        $this.IndexingLogFile = "$dir\indexing_log.txt"
        $this.GuiErrorLogFile = "$dir\gui_error_log.txt"
        $this.IndexingLogPreviousFile = "$dir\indexing_log_previous.txt"
        $this.GuiErrorLogPreviousFile = "$dir\gui_error_log_previous.txt"
        $this.TmpRoot = "$dir\tmp"
        $this.OfficePidRoot = "$dir\office_pids"
        $this.StartupTraceFile = "$dir\startup_trace.tsv"
    }

    # ワークスペースを移すときに移すもの（tebunko が作るファイル・フォルダ）。利用者のほかのファイルは含めない。
    # 取り込みの出力（publish）はプロセスごとのフォルダ（PublishDir）の親を移す。
    # StoreDir の中（IndexDir・StatusFile・FormatFile・JournalFile）は StoreDir ごと移す。
    # OldIndexDir・OldStatusFile・LegacyIndexDir は、あれば前の版のワークスペースとして tebunko のものと扱う（あるものだけが WorkspaceMover.ListEntries で拾われる）
    # 前の版の取り込み中のファイルの記録（ワークスペース直下の ingesting.txt）は含めない。新しい版は始めに 1 度だけ読んで消す一時的な跡で、
    # 移すときに残っても取り込み直しが増えるだけ（取り込み一覧が前回のままなので、止まったファイルは更新ありのまま次の回で取り込まれる）
    [string[]] Entries() {
        return @($this.StoreDir, $this.OldIndexDir, $this.OldStatusFile, $this.LegacyIndexDir,
            $this.ResultFile, $this.IndexingLogFile, $this.GuiErrorLogFile,
            $this.IndexingLogPreviousFile, $this.GuiErrorLogPreviousFile, [System.IO.Path]::GetDirectoryName($this.PublishDir), $this.TmpRoot, $this.OfficePidRoot, $this.StartupTraceFile)
    }
}

function getWorkspaceDirOfIndexDir {
    # 本文インデックスのフォルダ（Workspace.IndexDir）から、そのワークスペースのフォルダを返す。
    # IndexDir の位置を Workspace の外で仮定しないよう、親をたどる段数はここだけで持つ
    param (
        [string]$indexDir
    )

    $storeDir = [System.IO.Path]::GetDirectoryName($indexDir.TrimEnd("\"))
    return [System.IO.Path]::GetDirectoryName($storeDir)
}

function getMachineKey {
    # この PC を識別する短い鍵（getFolderKey の先頭 8 文字）。
    # 共有フォルダのワークスペースを複数の PC から使うとき、一時フォルダを PC ごとに分けるために使う（getWorkspaceTmpDir）
    return (getFolderKey ([Environment]::MachineName)).Substring(0, 8)
}

function getOfficePidDir {
    # 起動した Office の PID の記録を置くフォルダ（$workspace.OfficePidRoot\<PC の鍵>）を返す。文字列を組み立てるだけで、ファイルには触らない。
    # 共有のワークスペースで、ほかの PC の PID と混ざらないよう PC ごとに分ける
    param (
        [Workspace]$workspace
    )

    return Join-Path $workspace.OfficePidRoot (getMachineKey)
}

function getWorkspaceJobQueue {
    # ネットワークの場所に触るかもしれない startJob の列を決める唯一の関数。
    # paths（ワークスペース・インデックスのフォルダ・元のフォルダ・記録の置き場所など）の中に 1 つでもネットワークの場所があれば、
    # 届かない共有で待たされても他の仕事を巻き込まない専用の列（"network"）、無ければ既定の列（"default"）
    param (
        [string[]]$paths
    )

    return $(if (testAnyNetworkPath $paths) { "network" } else { "default" })
}

function assertWorkspaceReachable {
    # インデックスの削除・名前の変更の裏の仕事の先頭で呼ぶ。ワークスペース（のインデックスのフォルダ）に届かないとき、
    # 例外にする（届かないのに「何もせず成功」にして、設定だけ変えてしまわないため）。無いこと（まだインデックスを作っていない）は例外にしない
    param (
        [string]$dir
    )

    $state = getPathState $dir
    if ($state.State -eq ${pathStateUnreachable}) {
        throw "ワークスペースに接続できません：${dir}"
    }
    if ($state.State -eq ${pathStateOther}) {
        throw "ワークスペースを確かめられません：$($state.Message)"
    }
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
        return "ワークスペースのパスに [ ]（角かっこ）が含まれるため、更新の作業フォルダを置けません。"
    }
    if ($reason -eq "TooLong") {
        return "ワークスペースのパスが長すぎるため、更新の作業フォルダを置けません。"
    }
    if ($reason -eq "Link") {
        return "ワークスペースの作業領域の途中にリンク（ジャンクション・シンボリックリンク）があるため、更新の作業フォルダを置けません。"
    }
    return "更新の作業フォルダを置けません。"
}

function getLegacyIndexState {
    # 前の版のワークスペース（dir）の状態を返す: @{ HasLegacyIndex; ContentEmpty }
    #   HasLegacyIndex       : 前の版の index\ があり、直下のどれかのフォルダに legacySourceFolderFileName（元のフォルダ.txt）がある（前の版のしるし）。
    #                          index\ があるだけでは、利用者が選んだフォルダにたまたま index があるときと区別できないため、しるしとしない
    #   ContentEmpty         : 今の本文インデックスの置き場（IndexDir）と前の版の置き場（OldIndexDir）の両方が、無いか、下のフォルダを含めてファイルが 1 つも無い
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
            [System.IO.File]::Exists("$sub\${legacySourceFolderFileName}")
        }
    }

    $contentEmpty = $true
    foreach ($contentDir in @($ws.IndexDir, $ws.OldIndexDir)) {
        $longContent = toLongPath $contentDir
        if ($contentEmpty -and [System.IO.Directory]::Exists($longContent)) {
            $contentEmpty = !(testAnyEntry ([System.IO.Directory]::EnumerateFiles($longContent, "*", [System.IO.SearchOption]::AllDirectories)))
        }
    }

    return @{ HasLegacyIndex = $hasLegacyIndex; ContentEmpty = $contentEmpty }
}

function removeOldSystemIndex {
    # 前の版が作ったシステムインデックス（system_index\ と system_index_state.tsv）を消す: @{ Ok; Reason（失敗のときだけ） }
    # 新しい版は使わないため、画面の起動（裏の仕事）・インデックスの作成と更新の始め・インポートの登録前の 3 か所から呼ぶ（無ければ何もしない）。
    # ワークスペース（dir）の下で、途中にリンクが無いことを確かめてから消す（removeDirectoryUnder・removeFileUnder）。
    # 消せなくても例外にせず Ok = $false を返す（記録先は呼ぶ側が決める: 起動は画面エラーの記録・作成と更新はログ・インポートは警告。残して続け、次の機会にもう一度消す）
    param (
        [string]$dir
    )

    $ws = [Workspace]::new($dir)
    $problems = New-Object System.Collections.Generic.List[string]
    try {
        if (!(removeDirectoryUnder $ws.OldSystemIndexDir $ws.Dir)) {
            $problems.Add("途中にリンクがあるため、消せません: $($ws.OldSystemIndexDir)")
        }
    } catch {
        $problems.Add($_.Exception.Message)
    }
    try {
        if (!(removeFileUnder $ws.OldSystemIndexStateFile $ws.Dir)) {
            $problems.Add("途中にリンクがあるため、消せません: $($ws.OldSystemIndexStateFile)")
        }
    } catch {
        $problems.Add($_.Exception.Message)
    }
    if ($problems.Count -gt 0) {
        return @{ Ok = $false; Reason = ($problems -join " ") }
    }
    return @{ Ok = $true; Reason = "" }
}

function getLegacyIndexMessage {
    # 前の前の版のワークスペース（index\。しるしあり）が見つかったときの知らせ。しるしが無ければ空
    param (
        [string]$dir,
        [bool]$hasLegacyIndex
    )

    if (!$hasLegacyIndex) {
        return ""
    }
    $legacyDir = [Workspace]::new($dir).LegacyIndexDir
    return "前の版のインデックス（「${legacyDir}」）は、この版では使えません。［すべて更新］で、元のファイルをすべて更新し直します。" +
        "更新し直した後、「${legacyDir}」フォルダは削除してかまいません。"
}

function newWorkspace {
    # ワークスペースの場所（Dir。文字列）から Workspace を作る。クラスの本体は別のクラスの型を書かないため、クラスから作るときはこれを呼ぶ
    param (
        [string]$dir
    )

    return [Workspace]::new($dir)
}
