# 画面（WPF）の本体。gui.ps1（起動口）が、起動中の表示・Add-Type・Unblock-File・部品の読み込みを終えた後に呼ぶ。
#
# ui\ 配下のファイルは $PSScriptRoot を使わない（ここは tebunko\ui\ に置くため、$PSScriptRoot は ui\ を指してしまう）。
# tebunko\ のパスが要るときは、呼び出し元（gui.ps1）から渡された $TebunkoDir を使う。

function startGui {
    param (
        [string]$TebunkoDir
    )

    # 検索・画面の裏の仕事も同じプロセスのスレッドで動くため、画面を止める重い GC（全体の GC）をなるべく後回しにする
    # （docs/design/structure/closing.md「GC とメモリ」）
    [System.Runtime.GCSettings]::LatencyMode = [System.Runtime.GCLatencyMode]::SustainedLowLatency

    ${searchLimit} = 10000
    # 選択行のプレビューに出す行数は、プレビューの高さ（ドラッグで変わる）に収まるだけ出す（getPreviewContextLines）
    ${previewRowHeight}     = 22   # プレビューの 1 行の高さの目安。高さから出せる行数を求めるのに使う
    ${previewScrollBarSize} = 18   # 横スクロールバーの高さの目安（ViewportHeight が取れないときに引く）
    ${maxPreviewRows}       = 101  # プレビューに出す行数の上限（選択行＋前後 50 行）
    ${backgroundWorkers} = 2  # startJob のスレッドの数

    # 画面定義（XAML）とアイコンの置き場所
    ${xamlDir}       = "$TebunkoDir\xaml"
    ${sharedXamlDir} = "$TebunkoDir\..\shared\xaml"
    ${iconFile}      = "$TebunkoDir\tebunko.ico"  # タイトルバーとタスクバーに出すアイコン
    # アイコンは Window.Icon（loadWindow）でタイトルバー・タスクバーに出る。
    # ※以前は SetAppId（P/Invoke）でタスクバーのボタンを PowerShell と分けていたが、
    #   実行時コンパイル（csc.exe）を無くすため廃止した（アイコン自体は Window.Icon で出るため残る）。
    ${themeFile} = "${sharedXamlDir}\theme.xaml"  # 画面の見た目（色・文字・コントロールの形）の共通定義

    # ---- 多重起動の防止（ツールの配置フォルダごと） ----
    #
    # すでに開いているときは、その画面のウィンドウを前面に出して終わる（もう一度起動するのは、
    # たいてい「開いたつもりのウィンドウが他のウィンドウの裏にある」ときのため）。
    # 知らせるのは名前付きイベントで行う。ここは C# の型をコンパイルする前のため、.NET の機能だけを使う。

    $instanceKey = getFolderKey ${rootDir}
    $mutexName = "Local\${appId}_gui_" + $instanceKey
    $activateName = "Local\${appId}_gui_activate_" + $instanceKey
    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$createdNew)
    if (!$createdNew) {
        $running = $null
        if ([System.Threading.EventWaitHandle]::TryOpenExisting($activateName, [ref]$running)) {
            [void]$running.Set()
            $running.Close()
            exit
        }
        # 以前の版の画面が開いている等で知らせられないときだけ、メッセージを出す
        closeSplash
        [System.Windows.MessageBox]::Show("すでに開いています。", ${appTitle}, "OK", "Information") | Out-Null
        exit
    }
    $activateEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, $activateName)

    # ---- 壊れた設定ファイルの退避 ----
    # setting.config が JSON として読めないときは、別の名前に退避して既定の設定で起動する（画面が出てから知らせる。$script:settingsRecovery）。
    # 退避するのは、ここ（多重起動でない最初の起動）だけ。ほかの読み込みは、壊れたファイルを動かさずに例外にする（readSettings）
    $script:settingsRecovery = repairBrokenSettings
    # ワークスペース（setting.config の workspaceFolder）は、壊れた設定ファイルの退避の後に決める
    initWorkspace
    stepSplash 40

    $ErrorActionPreference = "Stop"

    # ---- ウィンドウと、画面の部品の対応 ----

    $window = loadWindow "${xamlDir}\tebunko.xaml"
    stepSplash 55

    # タブの中身はタブごとのファイルに分けてある。読み込んでタブに入れ、x:Name の対応表（$ui）を作る。
    # 別ファイルから読み込んだ中身は、そのファイルごとに名前を持つため、$window.FindName では見つからない。
    # タブごとに FindName する（名前の一覧もタブごとに分けておく）
    $tabs = @(
        @{ Tab = "IndexTab"; File = "tab_index.xaml"; Names = @(
            "IndexGrid", "IndexGridPlaceholder", "NewIndexButton", "EditIndexButton", "RemoveIndexButton",
            "ExportIndexButton", "ImportIndexButton",
            "IndexSummaryText", "IndexingStateText", "IndexingButton", "IndexingHint",
            "FailedPanel", "FailedHeading", "FailedGrid",
            "IndexingProgressPanel", "IndexingProgressText", "IndexingProgressEta", "IndexingProgress",
            "IndexingProgressDetail", "IndexingStopButton", "IndexingLogButton") }
        @{ Tab = "SearchTab"; File = "tab_search.xaml"; Names = @(
            "WordBox", "SearchButton", "RegexCheck", "CaseCheck", "ShapeCheck", "CommentCheck", "FileFilterBox", "FileFilterPlaceholder",
            "WordNotice", "SearchTargetText", "GoIndexTabButton", "FastSearchText",
            "IndexTree", "IndexTreePlaceholder", "CheckAllIndexButton", "UncheckAllIndexButton",
            "SummaryText", "SearchProgress", "FilterBox", "FilterPlaceholder", "ExpandAllButton", "CollapseAllButton", "ResultGrid",
            "MenuOpen", "MenuOpenReadOnly", "MenuOpenNew", "MenuOpenFolder", "MenuCopy", "MenuCopyPath",
            "DetailPanel", "DetailTitle", "OpenButton", "OpenModeCombo", "OpenFolderButton",
            "PreviewScroll", "PreviewHeaderScroll", "PreviewHeader", "PreviewRows", "PreviewNote",
            "PreviewPlaceholder", "MenuPreviewCopy", "MenuPreviewCopyRow", "ExportButton") }
        @{ Tab = "SettingsTab"; File = "tab_settings.xaml"; Names = @(
            "WorkspaceText", "WorkspaceNote", "ChangeWorkspaceButton", "ResetWorkspaceButton", "SettingsFileText", "SettingsFileNote") }
        @{ Tab = "KillTab"; File = "tab_kill.xaml"; Names = @(
            "ProcessGrid", "ProcessSummaryText", "RefreshProcessButton",
            "KillAllButton", "KillSelectedButton", "KillBackgroundButton") }
    )

    $ui = @{}
    foreach ($name in @("Tabs", "IndexTab", "SearchTab", "SettingsTab", "KillTab", "IndexTabHeader", "KillTabHeader", "StatusText", "MoreButton", "AboutMenuItem")) {
        $ui[$name] = $window.FindName($name)
    }
    foreach ($tab in $tabs) {
        $content = loadXaml "${xamlDir}\$($tab.File)"
        $ui[$tab.Tab].Content = $content
        foreach ($name in $tab.Names) {
            $ui[$name] = $content.FindName($name)
        }
        stepSplash 70
    }
    $taskbar = $window.TaskbarItemInfo

    ${okBrush}   = themeBrush "Ok"
    ${warnBrush} = themeBrush "Warn"
    ${ngBrush}   = themeBrush "Danger.Text"
    ${infoBrush} = themeBrush "Accent"
    ${grayBrush} = themeBrush "Ink.Muted"

    # ---- 画面の中身（それぞれのファイルにイベントの登録まで入っている。$ui を作った後に読み込む） ----
    . "$TebunkoDir\..\shared\ui\shell.ps1"
    # safe で包んでいない処理（Tabs.SelectionChanged の前半・PreviewKeyDown・Closing・活性化のタイマーなど）が
    # 投げた例外や、XAML の描画中に WPF が投げる例外を、画面のスレッドの Dispatcher で受ける
    [void](registerUnhandledErrorHandler $window.Dispatcher)
    # 画面から頼む短い仕事（startJob）のスレッド。長い仕事（件数の数え上げ等）の間もプレビューが待たないよう 2 つにする。
    # 各スレッドは最初の仕事の前に lib.ps1 を 1 回だけ読み込み、ワークスペースを決める
    # （docs/design/structure/threads.md「スレッドの一覧」）。
    # 列を作る式は 1 か所にまとめ、既定の列はここで作り、ネットワークの列は shell.ps1 に式（factory）だけ渡して
    # 初めて使うときに作らせる（届かない共有が無い利用者には、スレッドも lib.ps1 の読み込みも増えない）
    $backgroundLoad = getPartLoad lib
    $backgroundLoad.Prelude = "$($backgroundLoad.Prelude); initWorkspace"
    $newBackgroundQueue = { [BackgroundQueue]::new(${backgroundWorkers}, $backgroundLoad, $Host) }
    $script:backgroundQueue = & $newBackgroundQueue
    setNetworkQueueFactory $newBackgroundQueue
    . "$TebunkoDir\..\shared\ui\folder_dialog.ps1"
    . "$TebunkoDir\ui\index_view.ps1"
    . "$TebunkoDir\ui\indexing_view.ps1"
    . "$TebunkoDir\ui\search_view.ps1"
    . "$TebunkoDir\ui\preview_view.ps1"
    . "$TebunkoDir\ui\settings_view.ps1"
    . "$TebunkoDir\ui\about_view.ps1"
    stepSplash 80
    . "$TebunkoDir\ui\index_tab.ps1"
    . "$TebunkoDir\ui\indexing_tab.ps1"
    . "$TebunkoDir\ui\result_list.ps1"
    . "$TebunkoDir\ui\search_tab.ps1"
    . "$TebunkoDir\ui\preview.ps1"
    . "$TebunkoDir\ui\open_source.ps1"
    . "$TebunkoDir\ui\index_tree.ps1"
    . "$TebunkoDir\ui\process_tab.ps1"
    . "$TebunkoDir\ui\settings_tab.ps1"
    . "$TebunkoDir\ui\about_dialog.ps1"
    stepSplash 90
    # ============================================================================
    # ウィンドウ全体
    # ============================================================================

    # 起動時の読み込み（loadStartupData）が済んだか。済むまでは、タブの切り替え・ウィンドウの前面化で読み直さない
    # （起動時のタブを選んだとき・ウィンドウを出したときにも呼ばれ、同じ読み込みが重なるため）
    $script:startupLoaded = $false

    $ui.Tabs.Add_SelectionChanged({
        param ($sender, $e)
        # 中の表・一覧の選択変更も伝わってくるため、タブの切り替えだけを扱う
        if ($e.OriginalSource -ne $ui.Tabs -or !$script:startupLoaded) {
            return
        }
        safe {
            if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
                refreshProcesses
                $script:processTimer.Start()
            } else {
                $script:processTimer.Stop()
            }
            if ($ui.Tabs.SelectedItem -eq $ui.IndexTab) {
                refreshIndexingState
            }
        }
    })

    $window.Add_Activated({
        if (!$script:startupLoaded) {
            return
        }
        safe {
            # クロール対象フォルダがほかの画面で変更されていれば読み直す
            if ((getTargetsKey @(getTargetFolders)) -ne $script:savedTargets) {
                loadTargets
                setStatus "インデックス一覧がほかで変更されたため、読み直しました"
            }
            # フォルダの有無は別スレッドで調べる（届かないネットワークのフォルダで画面が固まらないように）
            refreshFolderStatus
            if (!(isIndexing)) {
                refreshIndexingState
                if (shouldRefreshFastSearchStatus) {
                    refreshFastSearchStatus
                }
            }
            updateSearchTarget
            if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
                refreshProcesses
            } else {
                updateKillBadge
            }
        }
    })

    $window.Add_PreviewKeyDown({
        param ($sender, $e)
        $modifiers = [System.Windows.Input.Keyboard]::Modifiers
        if ($e.Key -eq "F" -and $modifiers -eq "Control") {
            $ui.Tabs.SelectedItem = $ui.SearchTab
            $ui.WordBox.Focus() | Out-Null
            $ui.WordBox.SelectAll()
            $e.Handled = $true
        } elseif ($e.Key -eq "F" -and $modifiers -eq ([System.Windows.Input.ModifierKeys]::Control -bor [System.Windows.Input.ModifierKeys]::Shift)) {
            $ui.Tabs.SelectedItem = $ui.SearchTab
            $ui.FilterBox.Focus() | Out-Null
            $e.Handled = $true
        } elseif ($e.Key -eq "F5") {
            safe {
                if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
                    refreshProcesses
                } else {
                    refreshIndexingState
                    refreshIndexSummary
                    loadIndexTree
                }
            }
            $e.Handled = $true
        } elseif ($e.Key -eq "Escape" -and $script:search) {
            cancelSearch
            $e.Handled = $true
        }
    })

    # 閉じるときの順番（docs/design/structure/closing.md「閉じるときの順番」）。インデックス作成は画面のプロセスのスレッドで動くため、
    # 止めてから閉じる。止め終わるまで閉じるのを保留し、closeTimer が終わりを待ってから閉じ直す
    $script:closeWaiting = $false   # インデックス作成が止まるのを待っている
    $script:closeDeadline = $null   # これを過ぎたら、インデックス作成が起動した Office を止める
    $script:closeKilled = $false    # Office を止めた
    $script:closeReady = $false     # 待ち終えた（もう聞かずに閉じる）
    ${closeWaitSeconds} = 60        # インデックス作成が止まるのを待つ時間。過ぎたら Office を止めて、さらに closeKillWaitSeconds 待つ
    ${closeKillWaitSeconds} = 15

    $script:closeTimer = newTimer 500 {
        safe {
            $now = Get-Date
            if (!(isIndexing)) {
                $script:closeTimer.Stop()
                $script:closeReady = $true
                $window.Close()
            } elseif (!$script:closeKilled -and $now -gt $script:closeDeadline) {
                # 取り込み中の Office が応答しない。インデックス作成が起動した Office だけを PID で止める（COM の呼び出しが戻る）
                $script:closeKilled = $true
                [void]$script:indexingSession.KillOffice()
                $script:closeDeadline = $now.AddSeconds(${closeKillWaitSeconds})
            } elseif ($script:closeKilled -and $now -gt $script:closeDeadline) {
                # それでも止まらない。スレッドは finally で止める
                $script:closeTimer.Stop()
                $script:closeReady = $true
                $window.Close()
            }
        }
    }

    function enterIndexingStopFlow {
        # ふだんの「インデックス作成を止めてから閉じる」流れに入る（すでに入っていれば、足りない手順だけ補う）。
        # 1 つずつ try で包み、失敗しても残りを進める（finally の IndexingSession.Close の打ち切りに頼らず、
        # 取り込み中のファイルの区切りで止めるため）。closeDeadline は $null のときだけ決め、
        # 途中の失敗で入り直しても延ばし直さない。closeTimer も動いていなければ動かす（M1: ここが抜けていると、
        # 流れの途中の失敗のあとに closeTimer が動かず、外から止める以外に閉じられなくなる）
        $script:closeWaiting = $true
        try { $script:indexingTimer.Stop() } catch { }
        try { $script:indexingSession.Stop() } catch { }
        try { $ui.IndexingStopButton.IsEnabled = $false } catch { }
        try { $ui.IndexingProgressText.Text = "インデックス作成を止めています…" } catch { }
        try { $ui.IndexingProgressDetail.Text = "取り込み中のファイルが終わると、画面を閉じます。" } catch { }
        try { setStatus "インデックス作成を止めてから閉じます…" } catch { }
        try {
            if ($null -eq $script:closeDeadline) {
                $script:closeDeadline = (Get-Date).AddSeconds(${closeWaitSeconds})
            }
            if (!$script:closeTimer.IsEnabled) {
                $script:closeTimer.Start()
            }
        } catch { }
    }

    $window.Add_Closing({
        param ($sender, $e)
        if ($script:closeReady) {
            return
        }
        if ($script:closeWaiting) {
            # 止まるのを待っている間に、もう一度閉じようとした
            $e.Cancel = $true
            return
        }
        try {
            # インデックス作成は画面のプロセスで動いているため、画面を閉じるときは止める
            if (isIndexing) {
                $answer = showConfirm `
                    -heading "まだインデックス作成の途中です。止めてから閉じますか？" `
                    -facts @(
                        (factNext "いま取り込んでいるファイルが終わったところで止まり、画面を閉じます"),
                        (factKept "ここまで取り込んだ分はそのまま残ります" "次に開いて［インデックス作成を開始］を押すと、続きから再開します")
                    ) `
                    -choices @(@{ Text = "インデックス作成を止めて閉じる"; Value = "stop"; Careful = $true }) `
                    -cancelText "閉じない"
                $e.Cancel = $true
                if ($answer -ne "stop" -or !(isIndexing)) {
                    if ($answer -eq "stop") {
                        # 聞いている間に終わった
                        $script:closeReady = $true
                        $window.Dispatcher.BeginInvoke([action]{ $window.Close() }) | Out-Null
                    }
                    return
                }
                enterIndexingStopFlow
                return
            }
            if ($script:search) {
                cancelSearch
            }
        } catch {
            # 知らせる処理自体が失敗しても、閉じ方の分岐へは必ず進む（R1）
            try { reportUnexpectedError "画面を閉じる途中" $_ } catch { }
            $stillIndexing = $true
            try {
                $stillIndexing = isIndexing
            } catch {
                # isIndexing 自体が失敗したときは、安全な側（止める流れ）に倒す
                $stillIndexing = $true
            }
            if (!$stillIndexing -and !$script:closeWaiting) {
                # インデックス作成中でなく、まだ止める流れにも入っていなければ、閉じるのを止めない
                $e.Cancel = $false
                return
            }
            # 確認のダイアログは出し直さない。すでに流れの途中（closeWaiting が立っている）で
            # 失敗したときも、抜けている手順（closeTimer が動いていないなど）を補う
            $e.Cancel = $true
            enterIndexingStopFlow
        }
    })

    $window.Add_Loaded({
        safe {
            if ($ui.Tabs.SelectedItem -eq $ui.SearchTab) {
                $ui.WordBox.Focus() | Out-Null
            }
            if ($script:settingsRecovery) {
                closeSplash
                showMessage (getSettingsRecoveryMessage $script:settingsRecovery) "OK" "Warning" | Out-Null
            }
            if ($script:workspaceBlock) {
                # 知らせを読む間、起動中の表示が裏に残らないように先に閉じる
                closeSplash
                showMessage $script:workspaceBlock "OK" "Warning" | Out-Null
            }
        }
    })

    # 画面を描き終わったら、起動中の表示を閉じ、一覧の読み込みと別スレッドでの集計を始める。
    # ウィンドウを出す前に行うと、そのぶん画面が出るのが遅れるため（一覧は読み込むまで空で出る）
    $window.Add_ContentRendered({
        closeSplash
        # 描いた画面が映ってから読み込むよう、描画より優先度の低い Background で行う
        $window.Dispatcher.BeginInvoke([action]{ safe { loadStartupData } }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
    })

    function loadWorkspaceViews {
        # ワークスペースの中身（インデックスの一覧・取り込みの状態・検索対象のツリー・件数）を画面に読み込む。
        # 起動したとき（loadStartupData）と、［8 設定］でワークスペースを変えたとき（settings_tab.ps1 の switchWorkspace）に呼ぶ
        loadTargets
        refreshIndexingState
        loadIndexTree
        # 検索対象のツリーを読み込んだので、［検索］の可否を決め直す
        updateSearchButton
        checkFastSearchAvailable
        refreshIndexSummary
    }

    function ensureNetworkDriveCache {
        # インデックスの一覧・設定にネットワークのパスがあるときだけ、ドライブの割り当て（CIM）をネットワークを調べる列で
        # 1 回照会し、画面のスレッドのキャッシュに入れる（画面のスレッドで CIM を照会しないようにする。方針 5）。
        # ローカルだけの利用者には、列も CIM の照会も増えない
        $paths = @($script:targetItems | ForEach-Object { [string]$_.Path }) + @(readIndexSources | ForEach-Object { [string]$_.Path })
        if (!(testAnyNetworkPath $paths)) {
            return
        }
        startJob {
            getDriveTargets
        } @() {
            param ($output, $errorText)
            if (!$errorText -and $output -and $output.Count -gt 0) {
                setDriveTargets $output[0]
            }
        } "network"
    }

    function loadStartupData {
        try {
            loadWorkspaceViews
            updateKillBadge
            ensureNetworkDriveCache
        } finally {
            $script:startupLoaded = $true
        }
        # 起動時に出した知らせ（既定のワークスペースが使えない・前の版のインデックスがある）は、読み込みが終わっても消さない
        setStatus $(if ($script:workspaceBlock) { $script:workspaceBlock } elseif ($script:legacyIndexMessage) { $script:legacyIndexMessage } else { "" })
    }

    # ---- 起動 ----

    # 版・コミットの表示は起動時に 1 回だけ組み立てる（VERSION.txt は配布物にだけあり、開発中は無いので「開発版」になる）
    $script:aboutView = getAboutView (readVersionFile (Join-Path ${rootDir} "VERSION.txt"))

    updateSettingsView
    setSearchOptionToUi (readSearchOption)
    setOpenMode (readOpenMode)
    updateOpenMenu
    # 検索ワードの注意と［検索］の可否（検索ワードが空なので押せない）。画面を出したときに押せる色で出ないよう、先に決める
    updateWordNotice
    # 一覧を読み込むまで（loadStartupData）は、「インデックスがありません」の案内を出さない
    $ui.IndexGridPlaceholder.Visibility = "Collapsed"

    # 起動時のタブ：インデックス作成が中断中、またはインデックスが無ければ［1 インデックス管理］、それ以外は［2 検索］
    $openIndexTab = ($script:indexingState -and $script:indexingState.Pending -gt 0) -or !(testIndexExists)
    $ui.Tabs.SelectedItem = if ($openIndexTab) { $ui.IndexTab } else { $ui.SearchTab }
    # 前の版のインデックス（index\）が見つかれば、［1 インデックス管理］のステータスに知らせを出す
    $script:legacyIndexMessage = getLegacyIndexMessage $workspace.Dir (getLegacyIndexState $workspace.Dir).HasLegacyIndex
    if ($script:legacyIndexMessage) {
        setStatus $script:legacyIndexMessage
    }
    # 既定のワークスペースにほかのファイルが置いてあれば、［8 設定］を開いて別のフォルダを選んでもらう（画面を出した後に知らせる）
    $script:workspaceBlock = getWorkspaceBlockMessage
    if ($script:workspaceBlock) {
        $ui.Tabs.SelectedItem = $ui.SettingsTab
        setStatus $script:workspaceBlock
    }
    setStatus "読み込んでいます…"
    stepSplash 95

    # 多重起動したとき（2つ目のプロセスが $activateEvent を合図）に、この画面を前面へ出す。
    # 画面のスレッドで一定間隔にイベントを確認する（P/Invoke を使わず、WPF の Activate で前面化する）。
    # ※以前は C# の SingleInstance（AttachThreadInput 等の P/Invoke）で行っていたが、
    #   実行時コンパイル（csc.exe）を無くすため、DispatcherTimer＋Window.Activate に置き換えた。
    $activateTimer = New-Object System.Windows.Threading.DispatcherTimer
    $activateTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $activateTimer.Add_Tick({
        if ($activateEvent.WaitOne(0)) {
            if ($window.WindowState -eq [System.Windows.WindowState]::Minimized) {
                $window.WindowState = [System.Windows.WindowState]::Normal
            }
            [void]$window.Activate()
            # ほかのプロセスが前面のときは Activate が無視されることがあるため、最前面を一瞬立ててから戻す
            $window.Topmost = $true
            $window.Topmost = $false
        }
    })
    $activateTimer.Start()

    try {
        [void]$window.ShowDialog()
    } finally {
        # インデックス作成のスレッド、検索の司令のスレッドと照合のプール、画面の裏の仕事のスレッドを片づける
        # （docs/design/structure/closing.md「閉じるときの順番」）。片づける順番はそのまま変えない。
        # 画面の裏の仕事（$script:backgroundQueue・$script:networkQueue）だけ、止まった仕事（届かない共有の
        # Test-Path など、OS の呼び出しで戻らないもの）を待たずに戻る Abandon（前は Close）を使う
        $script:closeTimer.Stop()
        $script:indexingTimer.Stop()
        if ($script:indexingSession) {
            $script:indexingSession.Close()
        }
        $script:searchService.Close()
        $script:jobTimer.Stop()
        $script:backgroundQueue.Abandon()
        if ($script:networkQueue) {
            $script:networkQueue.Abandon()
        }
        $activateTimer.Stop()
        $activateEvent.Close()
        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }
}
