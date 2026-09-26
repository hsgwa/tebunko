# 画面（WPF）
#
# ［1 インデックス管理］［2 検索］［8 設定］［9 プロセス停止］の4タブ。画面の定義は xaml\tebunko.xaml。
# インデックス作成は indexer.ps1 をウィンドウを出さずに起動して進み具合を表示し、検索・プロセス停止は画面内で行う。
#
# このファイルは起動口。画面の中身は ui\ 配下と ..\shared\ui\ 配下に分けてある（下の読み込みの順に意味がある）。

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# ---- 起動中の表示 ----
# スクリプトの読み込みに数秒かかるため、先に小さなウィンドウ（xaml\splash.xaml）を出して、起動していることを知らせる。
# 画面（$window）を描き終わったら閉じる（ContentRendered）。多重起動の判定より前に出すため、2 つ目の起動でも一瞬出る。
# 読み込みの間は画面のスレッドが塞がるため、区切りごとに stepSplash で描画と入力の処理を進める（応答なしにしない）
$script:splash = $null
try {
    $splashStream = [System.IO.File]::OpenRead("$PSScriptRoot\xaml\splash.xaml")
    try {
        $script:splash = [System.Windows.Markup.XamlReader]::Load($splashStream)
    } finally {
        $splashStream.Dispose()
    }
    $script:splash.Show()
} catch {
    # 出せなくても起動は続ける
    $script:splash = $null
}

function stepSplash {
    param (
        [int]$percent
    )

    if ($null -eq $script:splash) {
        return
    }
    $script:splash.FindName("SplashProgress").Value = $percent
    # 自分のスレッドの Dispatcher に優先度 Background の空の仕事を渡し、それより優先度の高い描画・入力を先に処理させる
    $script:splash.Dispatcher.Invoke([action]{ }, [System.Windows.Threading.DispatcherPriority]::Background)
}

function closeSplash {
    if ($null -ne $script:splash) {
        $script:splash.Close()
        $script:splash = $null
    }
}

stepSplash 5

# zip 展開で付く Mark-of-the-Web（外部由来の印）を、scripts 配下から消す。印が残っていると
# RemoteSigned でスクリプトの読み込みがブロックされるため。通常は tebunko.bat が起動前に消すが、
# ショートカットから直接起動したときや、あとでファイルを差し替えたときのために、ここでも消しておく。
# （この gui.ps1 自身が印付きだと、この行に来る前にブロックされる。その場合は tebunko.bat から起動する）
try {
    Get-ChildItem -LiteralPath (Split-Path $PSScriptRoot -Parent) -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
} catch { }

. "$PSScriptRoot\lib.ps1"
stepSplash 40

$ErrorActionPreference = "Stop"

# 検索・画面の裏の仕事も同じプロセスのスレッドで動くため、画面を止める重い GC（全体の GC）をなるべく後回しにする
# （docs/design/architecture/threads.md「GC とメモリ」）
[System.Runtime.GCSettings]::LatencyMode = [System.Runtime.GCLatencyMode]::SustainedLowLatency

${appTitle}    = "tebunko"
${searchLimit} = 10000
# 選択行のプレビューに出す行数は、プレビューの高さ（ドラッグで変わる）に収まるだけ出す（getPreviewContextLines）
${previewRowHeight}     = 22   # プレビューの 1 行の高さの目安。高さから出せる行数を求めるのに使う
${previewScrollBarSize} = 18   # 横スクロールバーの高さの目安（ViewportHeight が取れないときに引く）
${maxPreviewRows}       = 101  # プレビューに出す行数の上限（選択行＋前後 50 行）
${libPath}  = "$PSScriptRoot\lib.ps1"  # 別スレッドで読み込む（検索の司令・startJob のスレッド）
${backgroundWorkers} = 2  # startJob のスレッドの数
# インデクサ（ウィンドウを出さずに別プロセスで起動する。indexing_tab.ps1）。
# ui\ 配下のファイルの中で $PSScriptRoot を使うと ui\ を指してしまうため、パスはここで決める
${indexerScriptPath} = "$PSScriptRoot\indexer.ps1"

# 画面定義（XAML）とアイコンの置き場所
${xamlDir}       = "$PSScriptRoot\xaml"
${sharedXamlDir} = "$PSScriptRoot\..\shared\xaml"
${iconFile}      = "$PSScriptRoot\tebunko.ico"  # タイトルバーとタスクバーに出すアイコン
# アイコンは Window.Icon（loadWindow）でタイトルバー・タスクバーに出る。
# ※以前は SetAppId（P/Invoke）でタスクバーのボタンを PowerShell と分けていたが、
#   実行時コンパイル（csc.exe）を無くすため廃止した（アイコン自体は Window.Icon で出るため残る）。
${themeFile} = "${sharedXamlDir}\theme.xaml"  # 画面の見た目（色・文字・コントロールの形）の共通定義
${theme} = $null                              # 読み込んだ theme.xaml（コードから色を引くときに使う）

function getGuiErrorLogFile {
    # 画面で起きた予期しないエラーの記録先（app_host.ps1 の writeErrorLog が使う）。今のワークスペースの中に置く
    return $workspace.GuiErrorLogFile
}

trap {
    # 記録できる状態（lib.ps1 の読み込み後）なら、内容をファイルにも残す
    if (Get-Command writeErrorLog -ErrorAction SilentlyContinue) {
        writeErrorLog "起動・実行中" $_
    }
    if ($null -ne $script:splash) {
        $script:splash.Close()
    }
    [System.Windows.MessageBox]::Show("予期しないエラーが発生しました。`n$($_.Exception.Message)", ${appTitle}, "OK", "Error") | Out-Null
    exit 1
}

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

# ---- 画面で使う型と、画面の土台 ----
# 型は継承元（shared）を先に読み込む。app_host は XAML の読み込みに使うため、ウィンドウを作る前に読み込む
. "$PSScriptRoot\..\shared\ui\types.ps1"
. "$PSScriptRoot\ui\types.ps1"
. "$PSScriptRoot\..\shared\ui\app_host.ps1"
stepSplash 50

# ---- ウィンドウと、画面の部品の対応 ----

$window = loadWindow "${xamlDir}\tebunko.xaml"
stepSplash 55

# タブの中身はタブごとのファイルに分けてある。読み込んでタブに入れ、x:Name の対応表（$ui）を作る。
# 別ファイルから読み込んだ中身は、そのファイルごとに名前を持つため、$window.FindName では見つからない。
# タブごとに FindName する（名前の一覧もタブごとに分けておく）
$tabs = @(
    @{ Tab = "IndexTab"; File = "tab_index.xaml"; Names = @(
        "IndexGrid", "IndexGridPlaceholder", "NewIndexButton", "EditIndexButton", "RemoveIndexButton",
        "IndexSummaryText", "IndexingStateText", "IndexingButton", "IndexingHint",
        "FailedPanel", "FailedHeading", "FailedGrid",
        "IndexingProgressPanel", "IndexingProgressText", "IndexingProgressEta", "IndexingProgress",
        "IndexingProgressDetail", "IndexingStopButton", "IndexingLogButton") }
    @{ Tab = "SearchTab"; File = "tab_search.xaml"; Names = @(
        "WordBox", "SearchButton", "RegexCheck", "CaseCheck", "ShapeCheck", "CommentCheck", "FileFilterBox", "FileFilterPlaceholder",
        "WordNotice", "SearchTargetText", "GoIndexTabButton", "FastSearchText",
        "IndexTree", "IndexTreePlaceholder", "CheckAllIndexButton", "UncheckAllIndexButton",
        "SummaryText", "SearchProgress", "FilterBox", "FilterPlaceholder", "ExpandAllButton", "CollapseAllButton", "ResultGrid", "IndexColumn",
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
foreach ($name in @("Tabs", "IndexTab", "SearchTab", "SettingsTab", "KillTab", "IndexTabHeader", "KillTabHeader", "StatusText")) {
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
. "$PSScriptRoot\..\shared\ui\shell.ps1"
# 画面から頼む短い仕事（startJob）のスレッド。長い仕事（件数の数え上げ等）の間もプレビューが待たないよう 2 つにする。
# 各スレッドは最初の仕事の前に lib.ps1 を 1 回だけ読み込む（docs/design/architecture/threads.md「スレッドの一覧」）
$script:backgroundQueue = [BackgroundQueue]::new(${backgroundWorkers}, ". '$(${libPath}.Replace("'", "''"))'", $Host)
. "$PSScriptRoot\..\shared\ui\folder_dialog.ps1"
. "$PSScriptRoot\ui\index_view.ps1"
. "$PSScriptRoot\ui\indexing_view.ps1"
. "$PSScriptRoot\ui\search_view.ps1"
. "$PSScriptRoot\ui\preview_view.ps1"
. "$PSScriptRoot\ui\settings_view.ps1"
stepSplash 80
. "$PSScriptRoot\ui\index_tab.ps1"
. "$PSScriptRoot\ui\indexing_tab.ps1"
. "$PSScriptRoot\ui\result_list.ps1"
. "$PSScriptRoot\ui\search_tab.ps1"
. "$PSScriptRoot\ui\preview.ps1"
. "$PSScriptRoot\ui\open_source.ps1"
. "$PSScriptRoot\ui\index_tree.ps1"
. "$PSScriptRoot\ui\process_tab.ps1"
. "$PSScriptRoot\ui\settings_tab.ps1"
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

# 閉じるときの順番（docs/design/architecture/threads.md「閉じるときの順番」）。インデックス作成は画面のプロセスのスレッドで動くため、
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
        $script:closeWaiting = $true
        $script:indexingTimer.Stop()
        $script:indexingSession.Stop()
        $ui.IndexingStopButton.IsEnabled = $false
        $ui.IndexingProgressText.Text = "インデックス作成を止めています…"
        $ui.IndexingProgressDetail.Text = "取り込み中のファイルが終わると、画面を閉じます。"
        setStatus "インデックス作成を止めてから閉じます…"
        $script:closeDeadline = (Get-Date).AddSeconds(${closeWaitSeconds})
        $script:closeTimer.Start()
        return
    }
    if ($script:search) {
        cancelSearch
    }
})

$window.Add_Loaded({
    safe {
        if ($ui.Tabs.SelectedItem -eq $ui.SearchTab) {
            $ui.WordBox.Focus() | Out-Null
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

function loadStartupData {
    try {
        loadWorkspaceViews
        updateKillBadge
    } finally {
        $script:startupLoaded = $true
    }
    setStatus ""
}

# ---- 起動 ----

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
    # （docs/design/architecture/threads.md「閉じるときの順番」）
    $script:closeTimer.Stop()
    $script:indexingTimer.Stop()
    if ($script:indexingSession) {
        $script:indexingSession.Close()
    }
    $script:searchService.Close()
    $script:jobTimer.Stop()
    $script:backgroundQueue.Close()
    $activateTimer.Stop()
    $activateEvent.Close()
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}