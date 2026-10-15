# 画面（WPF）の本体。gui.ps1（起動口）が、起動中の表示・Add-Type・Unblock-File・部品の読み込みを終えた後に呼ぶ。
#
# ui\ 配下のファイルは $PSScriptRoot を使わない（ここは tebunko\ui\ に置くため、$PSScriptRoot は ui\ を指してしまう）。
# tebunko\ のパスが要るときは、呼び出し元（gui.ps1）から渡された $TebunkoDir を使う。
#
# startGui は順に呼ぶだけにする。中身は ui\shell\ の下のファイルに分け、startGui の中で読み込む
# （$ui・$window・startGui のローカル変数・$script: の変数をそのまま見られるように、関数にせず dot-source する）。
# 画面を足すときは、領域の表（shell\regions.ps1）と読み込みの表（shell\screens.ps1）に行を足す。

function startGui {
    param (
        [string]$TebunkoDir
    )

    # 検索・画面の裏の仕事も同じプロセスのスレッドで動くため、画面を止める重い GC（全体の GC）をなるべく後回しにする
    # （docs/design/structure/closing.md「GC とメモリ」）
    [System.Runtime.GCSettings]::LatencyMode = [System.Runtime.GCLatencyMode]::SustainedLowLatency

    ${previewFirstLines}    = 50   # 行を選んだときに、プレビューへ最初に出す選択行の前後の行数（それぞれ）
    ${previewReadAheadLines} = 20  # 見えている行が、読んである範囲の端からこの行数以内に近づいたら、その向きに読み足す
    ${previewChunkLines}    = 200  # 1 回に読み足す行数（持つ行が上限を超えたとき、遠い側から捨てる行数も同じ）
    ${previewMaxRows}       = 3000 # プレビューに持つ行の上限（超えたら、見ている所から遠い側を捨てる）
    ${previewColumnMargin}  = 3    # 見えている列の左右に、先に作っておく列の数（横にスクロールするたびにセルを作り直さないため）
    ${previewCellRowMargin} = 10   # 見えている行の上下に、先にセルを作っておく行の数
    ${backgroundWorkers} = 2  # startJob のスレッドの数

    # 画面定義（XAML）とアイコンの置き場所
    ${xamlDir}       = "$TebunkoDir\xaml"
    ${sharedXamlDir} = "$TebunkoDir\..\shared\xaml"
    ${iconXamlFile}  = "${xamlDir}\app_icon.xaml"  # 窓・バージョン情報に出すアイコン（ベクターの絵）
    # アイコンは Window.Icon（loadWindow）でタイトルバー・タスクバーに出る。.ico（tebunko.ico）は画面では使わず、
    # インストーラー・ショートカットなど Windows が .ico しか受け付けない所だけで使う。
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

    # 壊れた設定ファイルの退避と、ワークスペースの決定
    . ([WildcardPattern]::Escape("$TebunkoDir\ui\shell\settings_recovery.ps1".Replace('`', '``')))

    # 検索結果の件数の上限（設定 searchLimit）。壊れた設定の退避のあとに、起動のときに 1 回読む
    ${searchLimit} = readSearchLimit

    $ErrorActionPreference = "Stop"

    # ウィンドウと、領域ごとの部品の対応表（$ui）
    $window = loadWindow "${xamlDir}\tebunko.xaml"
    stepSplash 55
    $ui = @{}
    . ([WildcardPattern]::Escape("$TebunkoDir\ui\shell\regions.ps1".Replace('`', '``')))

    ${okBrush}   = themeBrush "Ok"
    ${warnBrush} = themeBrush "Warn"
    ${ngBrush}   = themeBrush "Danger.Text"
    ${infoBrush} = themeBrush "Accent"
    ${grayBrush} = themeBrush "Ink.Muted"

    # 起動時の読み込み（loadStartupData）が済んだか。済むまでは、画面の切り替え・ウィンドウの前面化で読み直さない
    # （起動時の画面を選んだとき・ウィンドウを出したときにも呼ばれ、同じ読み込みが重なるため）。画面の読み込みの前に決める
    $script:startupLoaded = $false

    # 画面の中身（それぞれのファイルにイベントの登録まで入っている。$ui を作った後に読み込む）
    . ([WildcardPattern]::Escape("$TebunkoDir\ui\shell\services.ps1".Replace('`', '``')))
    . ([WildcardPattern]::Escape("$TebunkoDir\ui\shell\screens.ps1".Replace('`', '``')))

    # ウィンドウ全体: イベント・閉じるときの処理・起動時の読み込みとタブ選び・前面化の見張り
    . ([WildcardPattern]::Escape("$TebunkoDir\ui\shell\window_events.ps1".Replace('`', '``')))
    . ([WildcardPattern]::Escape("$TebunkoDir\ui\shell\closing.ps1".Replace('`', '``')))
    . ([WildcardPattern]::Escape("$TebunkoDir\ui\shell\startup.ps1".Replace('`', '``')))
    . ([WildcardPattern]::Escape("$TebunkoDir\ui\shell\activate_watch.ps1".Replace('`', '``')))

    try {
        [void]$window.ShowDialog()
    } finally {
        # インデックス作成のスレッド、検索の司令のスレッドと照合のプール、画面の裏の仕事のスレッドを片づける
        # （docs/design/structure/closing.md「閉じるときの順番」）。片づける順番はそのまま変えない。
        # 画面の裏の仕事（$script:backgroundQueue・$script:networkQueue）だけ、止まった仕事（届かない共有の
        # Test-Path など、OS の呼び出しで戻らないもの）を待たずに戻る Abandon（前は Close）を使う
        $script:closeTimer.Stop()
        $script:indexingTimer.Stop()
        try { closeIndexingWait } catch { }
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
    # 画面のスレッドの Dispatcher を止める。止めずに戻ると、PowerShell が終わるときの片づけで、まれに終了コード 5 で終わることがある
    # （docs/design/structure/closing.md「閉じるときの順番」）。例外で抜けるときは通らない
    # （起動の失敗を知らせる画面（reportStartupFailure）が、この Dispatcher を使うため）
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
}
