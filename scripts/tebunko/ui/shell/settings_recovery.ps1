# startGui の一部: 壊れた設定ファイルの退避と、ワークスペースの決定。startGui の中で読み込む。

    # ---- 壊れた設定ファイルの退避 ----
    # setting.config が JSON として読めないときは、別の名前に退避して既定の設定で起動する（画面が出てから知らせる。$script:settingsRecovery）。
    # 退避するのは、ここ（多重起動でない最初の起動）だけ。ほかの読み込みは、壊れたファイルを動かさずに例外にする（readSettings）
    $script:settingsRecovery = repairBrokenSettings
    # ワークスペース（setting.config の workspaceFolder）は、壊れた設定ファイルの退避の後に決める
    initWorkspace
    stepSplash 40
