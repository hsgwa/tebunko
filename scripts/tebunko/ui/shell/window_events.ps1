# startGui の一部: ウィンドウのイベント（Activated・PreviewKeyDown・Loaded・ContentRendered）。startGui の中で読み込む。

    function onWindowActivated {
        if (!$script:startupLoaded) {
            return
        }
        if (!(testIndexReadAllowed ([bool]$script:workspaceMoving))) {
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
        }
    }

    $window.Add_Activated({ onWindowActivated })

    function onWindowPreviewKeyDown {
        # 窓のキー（Ctrl+Tab・F5・Ctrl+F など）の入口。modifiers は押されている修飾キー（テストでは差し替える）
        param (
            $sender,
            $e,
            $modifiers = [System.Windows.Input.Keyboard]::Modifiers
        )

        $ctrl = ($modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
        $shift = ($modifiers -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0
        # Alt などが混ざったキーは、これまでどおり扱わない
        $others = $modifiers -band (-bnot ([System.Windows.Input.ModifierKeys]::Control -bor [System.Windows.Input.ModifierKeys]::Shift))
        if ($others -ne 0) { return }
        if (!(testMainContentOperable ([bool]$script:workspaceMoving))) {
            # ワークスペースの中身を移している間は、窓のキー（Ctrl+Tab・F5・Ctrl+F など）を受けない
            $e.Handled = $true
            return
        }
        $action = getShortcutAction ([string]$e.Key) $ctrl $shift (getCurrentScreen)
        if (invokeShortcutAction $action) {
            $e.Handled = $true
        }
    }

    $window.Add_PreviewKeyDown({
        param ($sender, $e)
        onWindowPreviewKeyDown $sender $e
    })

    $window.Add_Loaded({
        safe {
            if ((getCurrentScreen) -eq "SearchTab") {
                $ui.WordBox.Focus() | Out-Null
            }
            # 起動時のお知らせを開いている間は、前回残った Office の確認を出さず、閉じたあとに出す（leftover_dialog.ps1）
            $script:leftoverNoticesOpen = $true
            try {
                if ($script:settingsRecovery) {
                    closeSplash
                    showMessage (getSettingsRecoveryMessage $script:settingsRecovery) "OK" "Warning" | Out-Null
                }
                if ($script:workspaceBlock) {
                    # 知らせを読む間、起動中の表示が裏に残らないように先に閉じる
                    closeSplash
                    showMessage $script:workspaceBlock "OK" "Warning" | Out-Null
                }
            } finally {
                $script:leftoverNoticesOpen = $false
            }
            resumeLeftoverPrompt
        }
    })

    # 画面を描き終わったら、起動中の表示を閉じ、一覧の読み込みと別スレッドでの集計を始める。
    # ウィンドウを出す前に行うと、そのぶん画面が出るのが遅れるため（一覧は読み込むまで空で出る）
    $window.Add_ContentRendered({
        closeSplash
        # 描いた画面が映ってから読み込むよう、描画より優先度の低い Background で行う
        $window.Dispatcher.BeginInvoke([action]{ safe { loadStartupData } }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
    })
