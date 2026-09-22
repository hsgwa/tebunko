# tebunko の画面（WPF）の起動口
#
# ［1 検索］（その下に［インデックス管理］［検索］）・［2 比較］・［9 プロセス停止］のタブを 1 つのウィンドウに組み立てる。
# 画面の枠は xaml\tebunko.xaml。タブの中身は、検索は ..\tebunko_grep\、比較は ..\tebunko_diff\、
# プロセス停止は ..\shared\ にある（下の読み込みの順に意味がある）。
#
# このフォルダ（scripts\tebunko\）は、ツールを組み立てるだけの場所。ツール同士は互いを読み込まない。
param (
    [string]$StartPage = ""   # 開くタブ（search = ［1 検索］、diff = ［2 比較］）。空なら起動時の状態で決める
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# zip 展開で付く Mark-of-the-Web（外部由来の印）を、scripts 配下から消す。印が残っていると
# RemoteSigned でスクリプトの読み込みがブロックされるため。通常は tebunko.bat が起動前に消すが、
# ショートカットから直接起動したときや、あとでファイルを差し替えたときのために、ここでも消しておく。
# （この gui.ps1 自身が印付きだと、この行に来る前にブロックされる。その場合は tebunko.bat から起動する）
try {
    Get-ChildItem -LiteralPath (Split-Path $PSScriptRoot -Parent) -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
} catch { }

# ツールの画面以外の部品（どちらも shared を読み込む。2 回読み込んでも同じ定義になる）
${grepDir} = "$PSScriptRoot\..\tebunko_grep"
${diffDir} = "$PSScriptRoot\..\tebunko_diff"
. "$PSScriptRoot\..\tebunko_grep\lib.ps1"
. "$PSScriptRoot\..\tebunko_diff\lib.ps1"

$ErrorActionPreference = "Stop"

${appTitle} = "tebunko"

# 画面定義（XAML）とアイコンの置き場所
${sharedXamlDir} = "$PSScriptRoot\..\shared\xaml"
${iconFile}      = "$PSScriptRoot\tebunko.ico"  # タイトルバーとタスクバーに出すアイコン（Window.Icon。loadWindow）
${themeFile}     = "${sharedXamlDir}\theme.xaml" # 画面の見た目（色・文字・コントロールの形）の共通定義
${theme}         = $null                         # 読み込んだ theme.xaml（コードから色を引くときに使う）

trap {
    # 記録できる状態（lib.ps1 の読み込み後）なら、内容をファイルにも残す
    if (Get-Command writeErrorLog -ErrorAction SilentlyContinue) {
        writeErrorLog "起動・実行中" $_
    }
    [System.Windows.MessageBox]::Show("予期しないエラーが発生しました。`n$($_.Exception.Message)", ${appTitle}, "OK", "Error") | Out-Null
    exit 1
}

# ---- 多重起動の防止（ツールの配置フォルダごと） ----
#
# すでに開いているときは、その画面のウィンドウを前面に出して終わる（もう一度起動するのは、
# たいてい「開いたつもりのウィンドウが他のウィンドウの裏にある」ときのため）。
# 知らせるのは名前付きイベントで行う。ここは型を読み込む前のため、.NET の機能だけを使う。

$md5 = New-Object System.Security.Cryptography.MD5CryptoServiceProvider
$instanceKey = [BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes(${rootDir}.ToLowerInvariant()))).Replace("-", "")
$mutexName = "Local\tebunko_gui_" + $instanceKey
$activateName = "Local\tebunko_gui_activate_" + $instanceKey
# 以前の版（tebunko_grep だけの画面）のミューテックス。同じフォルダで開いていれば、同じ設定・インデックスを書き換えるため開かない
$legacyMutexName = "Local\${appId}_gui_" + $instanceKey
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$createdNew)
if (!$createdNew) {
    $running = $null
    if ([System.Threading.EventWaitHandle]::TryOpenExisting($activateName, [ref]$running)) {
        [void]$running.Set()
        $running.Close()
        exit
    }
    [System.Windows.MessageBox]::Show("すでに開いています。", ${appTitle}, "OK", "Information") | Out-Null
    exit
}
$legacyMutex = $null
if ([System.Threading.Mutex]::TryOpenExisting($legacyMutexName, [ref]$legacyMutex)) {
    $legacyMutex.Dispose()
    $mutex.ReleaseMutex()
    $mutex.Dispose()
    [System.Windows.MessageBox]::Show("以前の版の画面（tebunko_grep）が開いています。閉じてから開き直してください。", ${appTitle}, "OK", "Information") | Out-Null
    exit
}
$activateEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, $activateName)

# ---- 画面で使う型と、画面の土台 ----
# 型は継承元（shared）を先に読み込む。app_host は XAML の読み込みに使うため、ウィンドウを作る前に読み込む
. "$PSScriptRoot\..\shared\ui\types.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\types_grep.ps1"
. "$PSScriptRoot\..\tebunko_diff\ui\types_diff.ps1"
. "$PSScriptRoot\..\shared\ui\app_host.ps1"

# ---- ウィンドウと、画面の部品の対応 ----

$window = loadWindow "$PSScriptRoot\xaml\tebunko.xaml"

# タブの中身はタブごとのファイルに分けてある。読み込んでタブに入れ、x:Name の対応表（$ui）を作る。
# 別ファイルから読み込んだ中身は、そのファイルごとに名前を持つため、$window.FindName では見つからない。
# タブごとに FindName する（名前の一覧もタブごとに分けておく）
$tabs = @(
    @{ Tab = "IndexTab"; File = "tab_index.xaml"; Dir = "${grepDir}\xaml"; Names = @(
        "IndexGrid", "IndexGridPlaceholder", "NewIndexButton", "EditIndexButton", "RemoveIndexButton",
        "IndexSummaryText", "IndexingStateText", "IndexingButton", "IndexingHint",
        "FailedPanel", "FailedHeading", "FailedGrid",
        "IndexingProgressPanel", "IndexingProgressText", "IndexingProgressEta", "IndexingProgress",
        "IndexingProgressDetail", "IndexingStopButton", "IndexingLogButton") }
    @{ Tab = "SearchTab"; File = "tab_search.xaml"; Dir = "${grepDir}\xaml"; Names = @(
        "WordBox", "SearchButton", "RegexCheck", "CaseCheck", "ShapeCheck", "CommentCheck", "FileFilterBox", "FileFilterPlaceholder",
        "WordNotice", "SearchTargetText", "GoIndexTabButton",
        "IndexTree", "IndexTreePlaceholder", "CheckAllIndexButton", "UncheckAllIndexButton",
        "SummaryText", "SearchProgress", "FilterBox", "FilterPlaceholder", "ExpandAllButton", "CollapseAllButton", "ResultGrid", "IndexColumn",
        "MenuOpen", "MenuOpenReadOnly", "MenuOpenNew", "MenuOpenFolder", "MenuCopy", "MenuCopyPath",
        "DetailPanel", "DetailTitle", "OpenButton", "OpenModeCombo", "OpenFolderButton",
        "PreviewScroll", "PreviewHeaderScroll", "PreviewHeader", "PreviewRows", "PreviewNote",
        "PreviewPlaceholder", "MenuPreviewCopy", "MenuPreviewCopyRow", "ExportButton") }
    @{ Tab = "DiffTab"; File = "tab_diff.xaml"; Dir = "${diffDir}\xaml"; Names = @("DiffRoot") }
    @{ Tab = "KillTab"; File = "tab_kill.xaml"; Dir = ${sharedXamlDir}; Names = @(
        "ProcessGrid", "ProcessSummaryText", "RefreshProcessButton",
        "KillAllButton", "KillSelectedButton", "KillBackgroundButton") }
)

$ui = @{}
foreach ($name in @("Tabs", "SearchPage", "SearchPageHeader", "SearchTabs", "IndexTab", "SearchTab", "DiffTab", "KillTab",
        "IndexTabHeader", "KillTabHeader", "StatusText")) {
    $ui[$name] = $window.FindName($name)
}
foreach ($tab in $tabs) {
    $content = loadXaml "$($tab.Dir)\$($tab.File)"
    $ui[$tab.Tab].Content = $content
    foreach ($name in $tab.Names) {
        $ui[$name] = $content.FindName($name)
    }
}
$taskbar = $window.TaskbarItemInfo

${okBrush}   = themeBrush "Ok"
${warnBrush} = themeBrush "Warn"
${ngBrush}   = themeBrush "Danger.Text"
${infoBrush} = themeBrush "Accent"
${grayBrush} = themeBrush "Ink.Muted"

# ---- 画面の中身（それぞれのファイルにイベントの登録まで入っている。$ui を作った後に読み込む） ----
. "$PSScriptRoot\..\shared\ui\shell.ps1"
. "$PSScriptRoot\..\shared\ui\folder_dialog.ps1"
. "$PSScriptRoot\..\shared\ui\open_file.ps1"
# 検索（［1 検索］）
. "$PSScriptRoot\..\tebunko_grep\ui\grep_page.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\index_view.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\indexing_view.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\search_view.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\preview_view.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\index_tab.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\indexing_tab.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\result_list.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\search_tab.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\preview.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\open_source.ps1"
. "$PSScriptRoot\..\tebunko_grep\ui\index_tree.ps1"
# 比較（［2 比較］）
. "$PSScriptRoot\..\tebunko_diff\ui\diff_view.ps1"
. "$PSScriptRoot\..\tebunko_diff\ui\diff_pane.ps1"
. "$PSScriptRoot\..\tebunko_diff\ui\diff_tree.ps1"
. "$PSScriptRoot\..\tebunko_diff\ui\diff_tab.ps1"
# プロセス停止（［9 プロセス停止］）。裏で Office を使っている作業を知らせる
. "$PSScriptRoot\..\shared\ui\process_tab.ps1"
registerBusyCheck { isIndexing } "インデックス作成"
registerBusyCheck { isDiffRunning } "比較"

# ============================================================================
# ウィンドウ全体
# ============================================================================

function selectPage {
    # 1 段目のタブを選ぶ（search / diff / kill）
    param (
        [string]$page
    )

    switch ($page) {
        "search" { $ui.Tabs.SelectedItem = $ui.SearchPage }
        "diff"   { $ui.Tabs.SelectedItem = $ui.DiffTab }
        "kill"   { $ui.Tabs.SelectedItem = $ui.KillTab }
    }
}

$ui.Tabs.Add_SelectionChanged({
    param ($sender, $e)
    # 中の表・一覧・2 段目のタブの選択変更も伝わってくるため、1 段目のタブの切り替えだけを扱う
    if ($e.OriginalSource -ne $ui.Tabs) {
        return
    }
    safe {
        if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
            refreshProcesses
            $script:processTimer.Start()
        } else {
            $script:processTimer.Stop()
        }
        if ($ui.Tabs.SelectedItem -eq $ui.SearchPage) {
            onGrepPageShown
        }
    }
})

$window.Add_Activated({
    safe {
        onGrepActivated
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
    if ($modifiers -eq "Control" -and ($e.Key -eq "D1" -or $e.Key -eq "NumPad1")) {
        selectPage "search"
        $e.Handled = $true
    } elseif ($modifiers -eq "Control" -and ($e.Key -eq "D2" -or $e.Key -eq "NumPad2")) {
        selectPage "diff"
        $e.Handled = $true
    } elseif ($modifiers -eq "Control" -and ($e.Key -eq "D9" -or $e.Key -eq "NumPad9")) {
        selectPage "kill"
        $e.Handled = $true
    } elseif ($e.Key -eq "F5" -and $ui.Tabs.SelectedItem -eq $ui.KillTab) {
        safe { refreshProcesses }
        $e.Handled = $true
    } elseif ($e.Key -eq "F" -and ($modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
        # Ctrl+F・Ctrl+Shift+F は、どのタブからでも［1 検索］の［検索］へ移る
        if (onGrepKeyDown $e $modifiers) {
            $e.Handled = $true
        }
    } elseif ($ui.Tabs.SelectedItem -eq $ui.DiffTab) {
        # ［2 比較］のキー操作（F5・F8・Esc など）
        if (onDiffKeyDown $e $modifiers) {
            $e.Handled = $true
        }
    } elseif (onGrepKeyDown $e $modifiers) {
        # ［1 検索］のキー操作（Ctrl+F はどのタブからでも［検索］へ移る）
        $e.Handled = $true
    }
})

$window.Add_Closing({
    param ($sender, $e)
    # 裏で動いている作業をどうするか聞く。どちらかで「閉じない」を選んだら閉じない
    if (!(confirmGrepClosing)) {
        $e.Cancel = $true
        return
    }
    if (!(confirmDiffClosing)) {
        $e.Cancel = $true
        return
    }
})

$window.Add_Loaded({
    safe {
        if ($ui.Tabs.SelectedItem -eq $ui.SearchPage -and $ui.SearchTabs.SelectedItem -eq $ui.SearchTab) {
            $ui.WordBox.Focus() | Out-Null
        }
    }
})

# ---- 起動 ----

startGrepPage
startDiffPage
updateKillBadge

# 起動時のタブ。指定が無ければ、比較を最後に使っていても［1 検索］で開く（検索の起動時の状態で 2 段目を決める）
if ($StartPage -eq "diff") {
    selectPage "diff"
} else {
    selectPage "search"
}
setStatus ""

# 多重起動したとき（2つ目のプロセスが $activateEvent を合図）に、この画面を前面へ出す。
# 画面のスレッドで一定間隔にイベントを確認する（P/Invoke を使わず、WPF の Activate で前面化する）
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
    closeGrepPage
    closeDiffPage
    $activateTimer.Stop()
    $activateEvent.Close()
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
