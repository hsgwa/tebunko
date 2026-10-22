# startGui の一部: 窓の領域の表（$regions）と、x:Name の対応表（$ui）。startGui の中で読み込む（$window・${xamlDir} と、呼ぶ側が作る $ui を使う）。
# 画面（領域）を足すときは、ここに行を足す。

    # 窓の中身は領域ごとのファイルに分けてある（xaml\<画面>\<領域>.xaml）。読み込んで差す口に入れ、x:Name の対応表（$ui）を作る。
    # 別ファイルから読み込んだ中身は、そのファイルごとに名前を持つため、$window.FindName では見つからない。
    # 領域ごとに FindName する（名前の一覧も領域ごとに分けておく）。
    # 領域の表: File = xaml\ からの相対パス、Slot = 差す口（窓の x:Name）、Screen = 画面の名前（Slot が ContentHost のものだけ。
    # 画面の中身は $script:screenContents に持ち、selectScreen が選んだものを ContentHost に差す）、Names = 使う x:Name
    # 知らせの帯の部品（shell\notice_banner.xaml）の名前。3 か所に差す（下の表で Prefix を付けて持つ）
    $noticeBannerNames = @("NoticePanel", "NoticeIcon", "NoticeText", "NoticeButton", "NoticeClose")
    $regions = @(
        @{ File = "shell\nav.xaml"; Slot = "NavHost"; Names = @(
            "NavList", "SearchTab", "IndexTab", "SettingsTab", "IndexTabBadge", "NavPaneHost", "AboutLink") }
        @{ File = "shell\status_bar.xaml"; Slot = "StatusBarHost"; Names = @("StatusText", "IndexingStatusText", "IndexFooterFolders", "IndexFooterFiles") }
        @{ File = "index\index.xaml"; Slot = "ContentHost"; Screen = "IndexTab"; Names = @(
            "IndexListHost", "IndexDetailHost", "IndexDetailRow", "IndexSplitter",
            "IndexingProgressPanel", "IndexingBannerIcon", "IndexingProgressText", "IndexingProgressEta", "IndexingProgress",
            "IndexingProgressDetail", "IndexingStopButton", "IndexingResumeButton", "IndexingSearchButton", "IndexingBannerClose",
            "IndexNoticeHost") }
        @{ File = "index\index_list.xaml"; Slot = "IndexListHost"; Names = @(
            "IndexGrid", "IndexGridPlaceholder", "NewIndexButton", "ActionsButton", "ActionsMenu", "ActionUpdate", "ActionExport", "ActionImport", "ActionDelete", "SelectionCountText", "SelectAllCheckBox", "IndexingButton", "IndexingHint", "StatusColumnHeader",
            "IndexRowMenu", "RowMenuUpdate", "RowMenuOpenFolder", "RemoveIndexButton", "ExportIndexButton") }
        @{ File = "index\index_detail.xaml"; Slot = "IndexDetailHost"; Names = @(
            "IndexDetailTitle", "IndexSummaryText",
            "IndexDetailRows",
            "FailedPanel", "FailedHeading", "FailedGrid") }
        @{ File = "search\search.xaml"; Slot = "ContentHost"; Screen = "SearchTab"; Names = @(
            "SearchBarHost", "ResultListHost", "PreviewHost", "DetailRow", "PreviewSplitter", "ResultEmptyState", "GoIndexTabButton",
            "SearchNoticeHost") }
        @{ File = "shell\notice_banner.xaml"; Slot = "SearchNoticeHost"; Prefix = "Search"; Margin = @(-16, -12, -16, 8); Names = $noticeBannerNames }
        @{ File = "shell\notice_banner.xaml"; Slot = "IndexNoticeHost"; Prefix = "Index"; Names = $noticeBannerNames }
        @{ File = "shell\notice_banner.xaml"; Slot = "WorkspaceNoticeHost"; Prefix = "Workspace"; Names = $noticeBannerNames }
        @{ File = "search\search_bar.xaml"; Slot = "SearchBarHost"; Names = @(
            "WordBox", "SearchButton", "RegexCheck", "CaseCheck", "ShapeCheck", "CommentCheck", "ScopeButton",
            "ConditionsPanel", "ConditionsSpacer","KindChipExcel", "KindChipWord", "KindChipPowerPoint", "KindChipText", "KindChipPdf",
            "WordPlaceholder", "WordErrorRing",
            "BalloonLayer", "SearchKindBalloon", "SearchKindBalloonText", "RegexBalloon", "RegexBalloonText") }
        @{ File = "search\target_tree.xaml"; Slot = "NavPaneHost"; Names = @(
            "IndexTree", "CheckAllIndexButton", "UncheckAllIndexButton", "TargetCountText") }
        @{ File = "search\result_list.xaml"; Slot = "ResultListHost"; Names = @(
            "SummaryText", "SearchProgress", "FilterBox", "FilterPlaceholder", "ExpandAllButton", "CollapseAllButton", "ExportButton", "ExportButtonText", "ExportMenuButton", "MenuExportXlsx", "MenuExportCsv", "MenuExportTsv",
            "ResultGrid", "NoHitText",
            "ResultMenu", "MenuOpen", "MenuOpenReadOnly", "MenuOpenNew", "MenuOpenFolder", "MenuCopy", "MenuCopyPath", "MenuToggleGroup") }
        @{ File = "search\preview.xaml"; Slot = "PreviewHost"; Names = @(
            "DetailPanel", "DetailTitle", "OpenButton", "OpenMenuButton", "MenuOpenModeNormal", "MenuOpenModeNew", "MenuOpenModeReadOnly", "OpenFolderButton",
            "PreviewHeaderScroll", "PreviewHeaderPanel", "PreviewHeader", "PreviewRows",
            "PreviewPlaceholder", "PreviewMenu", "MenuPreviewOpen", "MenuPreviewCopy", "MenuPreviewCopyRow") }
        @{ File = "settings\settings.xaml"; Slot = "ContentHost"; Screen = "SettingsTab"; Names = @(
            "WorkspaceBox", "WorkspaceInputError", "WorkspaceNote", "WorkspaceCaution", "WorkspaceHowTo", "ChangeWorkspaceButton", "ResetWorkspaceButton", "SettingsFileText", "SettingsFileNote",
            "SearchLimitedRadio", "SearchNoneRadio", "SearchLimitBox", "SearchLimitInputError", "SearchLimitRange", "SearchLimitCaution") }
    )

    foreach ($name in @("NavHost", "ContentHost", "StatusBarHost", "ScrimOverlay",
            "WorkspaceNoticeHost")) {
        $ui[$name] = $window.FindName($name)
    }
    $script:screenContents = @{}
    foreach ($region in $regions) {
        $content = loadXaml "${xamlDir}\$($region.File)"
        if ($region.Slot -eq "ContentHost") {
            $script:screenContents[$region.Screen] = $content
        } else {
            $ui[$region.Slot].Content = $content
        }
        foreach ($name in $region.Names) {
            # 同じ部品（知らせの帯）を何か所かに差すものは、Prefix を付けた名前で持つ
            $key = "$($region.Prefix)$name"
            $ui[$key] = $content.FindName($name)
            if ($region.Prefix) {
                [System.Windows.Automation.AutomationProperties]::SetAutomationId($ui[$key], $key)
            }
        }
        if ($region.Margin) {
            $content.Margin = [System.Windows.Thickness]::new($region.Margin[0], $region.Margin[1], $region.Margin[2], $region.Margin[3])
        }
        stepSplash 70
    }
    # プレビューの行の入れ物（PreviewRows）は、行の仮想化のために、自分のテンプレートの中にスクロール（PreviewScroll）を持つ。
    # テンプレートの中の名前は上の FindName では取れないので、テンプレートを作ってから取る
    [void]$ui.PreviewRows.ApplyTemplate()
    $ui.PreviewScroll = $ui.PreviewRows.Template.FindName("PreviewScroll", $ui.PreviewRows)
    $taskbar = $window.TaskbarItemInfo
