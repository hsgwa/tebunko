# startGui の一部: 共通の画面部品の読み込みと、画面の裏の仕事のスレッド。startGui の中で読み込む。
# $ui を作った後に読み込む。screens.ps1 より前。

    . ([WildcardPattern]::Escape("$TebunkoDir\..\shared\ui\message_view.ps1".Replace('`', '``')))
    . ([WildcardPattern]::Escape("$TebunkoDir\..\shared\ui\shell.ps1".Replace('`', '``')))
    # 標題は startup_error.ps1 の ${appTitle}（gui.ps1 が先に読み込む）。メッセージ・確認ダイアログを出す前に渡す
    setShellTitle ${appTitle}
    . ([WildcardPattern]::Escape("$TebunkoDir\..\shared\ui\data_grid.ps1".Replace('`', '``')))
    setDialogScrim $ui.ScrimOverlay
    # safe で包んでいない処理（PreviewKeyDown・Closing・活性化のタイマーなど）が
    # 投げた例外や、XAML の描画中に WPF が投げる例外を、画面のスレッドの Dispatcher で受ける
    [void](registerUnhandledErrorHandler $window.Dispatcher)
    # 画面から頼む短い仕事（startJob）のスレッド。長い仕事（件数の数え上げ等）の間もプレビューが待たないよう 2 つにする。
    # 各スレッドは最初の仕事の前に lib.ps1 を 1 回だけ読み込み、ワークスペースを決める
    # （docs/design/structure/threads.md「スレッドの一覧」）。
    # 列を作る式は 1 か所にまとめ、既定の列はここで作り、ネットワークの列は shell.ps1 に式（factory）だけ渡して
    # 初めて使うときに作らせる（届かない共有が無い利用者には、スレッドも lib.ps1 の読み込みも増えない）
    $backgroundLoad = getPartLoad lib -then "initWorkspace"
    $newBackgroundQueue = { [BackgroundQueue]::new(${backgroundWorkers}, $backgroundLoad, $Host) }
    setBackgroundQueue (& $newBackgroundQueue)
    setNetworkQueueFactory $newBackgroundQueue
