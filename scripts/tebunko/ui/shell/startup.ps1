# startGui の一部: 起動時の読み込み（loadWorkspaceViews・loadStartupData）と、起動時のタブ選びなど。startGui の中で読み込む。

    function loadWorkspaceViews {
        # ワークスペースの中身（インデックスの一覧・取り込みの状態・検索対象のツリー・件数）を画面に読み込む。
        # 起動したとき（loadStartupData）と、設定の画面でワークスペースを変えたとき（settings\settings.ps1 の switchWorkspace）に呼ぶ
        loadTargets
        refreshIndexingState
        loadIndexTree
        # 検索対象のツリーの読み込みを始めたので、［検索］の可否（読み込み中は検索できない）を決め直す
        updateSearchButton
        refreshIndexSummary
    }

    function ensureNetworkDriveCache {
        # インデックスの一覧・設定にネットワークのパスがあるときだけ、ドライブの割り当て（CIM）をネットワークを調べる列で
        # 1 回照会し、画面のスレッドのキャッシュに入れる（画面のスレッドで CIM を照会しないようにする。方針 5）。
        # ローカルだけの利用者には、列も CIM の照会も増えない
        $paths = @($script:targetItems | ForEach-Object { [string]$_.Path }) + @(readIndexSources | ForEach-Object { [string]$_.Path })
        if ((getWorkspaceJobQueue $paths) -ne "network") {
            return
        }
        startJob {
            getDriveTargets
        } @() {
            param ($output, $errorText)
            if (!$errorText -and $output -and $output.Count -gt 0) {
                setDriveTargets $output[0]
            }
        } (getWorkspaceJobQueue $paths)
    }

    function startOldSystemIndexRemoval {
        # 前の版が作ったシステムインデックス（system_index\・system_index_state.tsv）を、裏の仕事で消す（画面のスレッドでは触らない）。
        # 消せなくても起動は続け、画面エラーの記録に 1 行残す。無ければ何もしない。インデックス作成の始めとインポートでも消す（removeOldSystemIndex）
        startJob {
            param ($workspaceDir)
            removeOldSystemIndex $workspaceDir
        } @($workspace.Dir) {
            param ($output, $errorText)
            $reason = if ($errorText) { [string]$errorText } elseif ($null -ne $output -and @($output).Count -gt 0 -and !$output[0].Ok) { [string]$output[0].Reason } else { "" }
            if ($reason) {
                writeErrorLog "前の版が作った検索用のファイルを消せませんでした（起動時）" ([System.Management.Automation.ErrorRecord]::new([System.Exception]::new($reason), "OldSystemIndex", [System.Management.Automation.ErrorCategory]::NotSpecified, $null))
            }
        } (getWorkspaceJobQueue @($workspace.Dir))
    }

    function loadStartupData {
        try {
            loadWorkspaceViews
            # 前の版のインデックス（index\）が見つかれば、その知らせを覚える（裏で調べ、分かったらステータスに出す）
            refreshLegacyIndexMessage
            ensureNetworkDriveCache
        } finally {
            $script:startupLoaded = $true
        }
        # 起動時に出した知らせ（既定のワークスペースが使えない・前の版のインデックスがある）は、読み込みが終わっても消さない
        setStatus $(if ($script:workspaceBlock) { $script:workspaceBlock } else { "" })
        # 前回のインデックス作成が起動したまま残った Office があれば、終了するか確認する（記録の読み取りは裏で行う）
        startLeftoverCheck
        # 前の版が作った検索用のファイル（system_index\）を裏で消す
        startOldSystemIndexRemoval
        # ［新規で開く］のために作ったコピー（open_copy）のうち、1 日より古いものを裏で片付ける（起動時の 1 回だけ。開いているコピーは残る）
        startJob {
            param ($copyDir, $tmpRoot)
            [void](clearOldOpenCopies $copyDir $tmpRoot)
        } @((getOpenCopyDir $workspace), $workspace.TmpRoot) {
            param ($output, $errorText)
        } (getWorkspaceJobQueue @($workspace.Dir, (getOpenCopyDir $workspace)))
        # 強制終了などで残った、検索の溢れ（search_overflow）のうち、動いていないプロセスのものを裏で消す（起動時の 1 回だけ）
        startJob {
            param ($tmpRoot)
            [void](removeStaleSearchOverflow $tmpRoot)
        } @($workspace.TmpRoot) {
            param ($output, $errorText)
        } (getWorkspaceJobQueue @($workspace.Dir, $workspace.TmpRoot))
    }

    # ---- 起動 ----

    # 版・コミットの表示は起動時に 1 回だけ組み立てる（VERSION.txt は配布物にだけあり、開発中は無いので「開発版」になる）
    $versionInfo = readVersionFile (Join-Path ${rootDir} "VERSION.txt")
    $script:aboutView = getAboutView $versionInfo
    if ($null -eq $versionInfo) {
        # 開発版は、.git のファイルから分かるタグ・コミットを添える。ファイルを読むので、裏で調べて、分かったら表示を差し替える
        startJob {
            param ($toolDir)
            readDevVersion $toolDir
        } @(${rootDir}) {
            param ($output, $errorText)
            if (!$errorText -and $null -ne $output -and @($output).Count -gt 0 -and $null -ne $output[0]) {
                $script:aboutView = getAboutView $null $output[0]
            }
        }
    }

    updateSettingsView
    setSearchOptionToUi (readSearchOption)
    setOpenMode (readOpenMode)
    # 検索ワードの注意と［検索］の可否（検索ワードが空なので押せない）。画面を出したときに押せる色で出ないよう、先に決める
    updateWordNotice
    # 一覧を読み込むまで（loadStartupData）は、「インデックスがありません」の案内を出さない
    $ui.IndexGridPlaceholder.Visibility = "Collapsed"

    # 起動時の画面：インデックス作成が中断中、またはインデックスが無ければ［インデックス管理］、それ以外は［検索］
    # （ネットワークのワークスペースは、画面のスレッドで調べず、［検索］を開く）
    $openIndexTab = ($script:indexingState -and $script:indexingState.Pending -gt 0) -or !(testStartupIndexExists $workspace.Dir)
    selectScreen $(if ($openIndexTab) { "IndexTab" } else { "SearchTab" })
    # 既定のワークスペースにほかのファイルが置いてあれば、［設定］を開いて別のフォルダを選んでもらう（画面を出した後に知らせる）
    $script:workspaceBlock = getWorkspaceBlockMessage
    updateWorkspaceBanner
    if ($script:workspaceBlock) {
        selectScreen "SettingsTab"
    }
    setStatus "読み込んでいます…"
    stepSplash 95
