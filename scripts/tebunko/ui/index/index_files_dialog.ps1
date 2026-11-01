# 取り込んだファイルの一覧のダイアログ（詳細の［ファイルの一覧］と、失敗の入口の［一覧で見る］から開く）。画面層。
# 文言と絞り込みの判断は index_detail_view.ps1（判断層）。取り込み一覧の読み方は getIndexFileEntries（状態層）。

function newIndexFilesDialogControls {
    # ダイアログを読み込み、部品の表（名前 → 部品）を返す
    param (
        $owner,
        [string]$indexName
    )

    $dialog = loadWindow "${xamlDir}\dialog_index_files.xaml"
    $dialog.Owner = $owner
    $dialog.Title = getIndexFilesDialogTitle $indexName
    $ctrl = @{ Dialog = $dialog }
    foreach ($name in @("IndexFilesHeading", "IndexFilesFilter", "IndexFilesAllRadio", "IndexFilesFailedRadio", "IndexFilesGrid", "IndexFilesEmptyText", "IndexFilesHint", "IndexFilesCloseButton")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    return $ctrl
}

function applyIndexFilesFilter {
    # 読み込めた一覧を「すべて」「失敗のみ」に絞って表に出す（読み直さない）。0 件の案内は表の中央、見出しは要約のまま
    param (
        $ctx
    )

    $ctrl = $ctx.Controls
    $ctrl.IndexFilesAllRadio.IsChecked = ($ctx.State.Filter -eq "All")
    $ctrl.IndexFilesFailedRadio.IsChecked = ($ctx.State.Filter -eq "Failed")
    if ($null -eq $ctx.State.Entries) {
        return
    }
    $shown = @(selectIndexFileRows $ctx.State.Entries.Rows $ctx.State.Filter)
    $ctrl.IndexFilesGrid.ItemsSource = $shown
    $notice = getIndexFilesEmptyNotice $ctx.State.Filter $shown.Count
    $ctrl.IndexFilesEmptyText.Text = $notice
    $ctrl.IndexFilesEmptyText.Visibility = $(if ($notice) { "Visible" } else { "Collapsed" })
}

function applyIndexFilesEntries {
    # 読み込めた一覧（getIndexFileEntries の結果）を出す。見出し・絞り込みの有効化・絞り込みどおりの行
    param (
        $ctx,
        $entries
    )

    $ctx.State.Entries = $entries
    $ctx.Controls.IndexFilesHeading.Text = getIndexFilesHeading ($entries.Rows.Count) ($entries.Failed) ($entries.Pending)
    $ctx.Controls.IndexFilesFilter.IsEnabled = $true
    # 失敗したファイルがあるときだけ、原因の列（4 列目）を出す
    $ctx.Controls.IndexFilesGrid.Columns[3].Visibility = $(if ([int]$entries.Failed -gt 0) { "Visible" } else { "Collapsed" })
    applyIndexFilesFilter $ctx
}

function showIndexFilesDialog {
    # 「<インデックス名> のファイル」を開く（閉じるまで戻らない。ほかのダイアログと同じく showOwnedDialog で、後ろの窓に暗幕を掛ける）。
    #   Filter  All（すべて）／Failed（失敗のみ）。開き直したときは前の絞り込みを持ち越さず、入口で決める
    # 取り込み一覧は数万行になることがあるため、読み込みと整形は裏のスレッドで行い、届いたら表に渡すだけにする
    param (
        $Owner,
        [string]$IndexName,
        [string]$Filter = "All"
    )

    $ctrl = newIndexFilesDialogControls $Owner $IndexName
    $dialog = $ctrl.Dialog
    $ctx = @{ Controls = $ctrl; State = @{ Entries = $null; Filter = $(if ($Filter -eq "Failed") { "Failed" } else { "All" }); Closed = $false } }
    $ctrl.IndexFilesHeading.Text = getIndexFilesLoadingText
    $ctrl.IndexFilesFilter.IsEnabled = $false
    $ctrl.IndexFilesAllRadio.IsChecked = ($ctx.State.Filter -eq "All")
    $ctrl.IndexFilesFailedRadio.IsChecked = ($ctx.State.Filter -eq "Failed")

    # ハンドラ・裏の仕事の結果を受けるところからは関数の名前が見つからないため、実体を変数に取って & で呼ぶ
    $applyFilter = ${function:applyIndexFilesFilter}
    $applyEntries = ${function:applyIndexFilesEntries}
    $openFolder = ${function:openListedFileFolder}
    $grid = $ctrl.IndexFilesGrid
    # クリックでなく Checked で受ける（支援技術の選択・キー操作でも絞り込みが動く）
    $ctrl.IndexFilesAllRadio.Add_Checked({ $ctx.State.Filter = "All"; & $applyFilter $ctx }.GetNewClosure())
    $ctrl.IndexFilesFailedRadio.Add_Checked({ $ctx.State.Filter = "Failed"; & $applyFilter $ctx }.GetNewClosure())
    # ［閉じる］は IsCancel（クリックでも Esc でも閉じる）
    $dialog.Add_Closed({
        $ctx.State.Closed = $true
        $grid.ItemsSource = $null
    }.GetNewClosure())
    $grid.Add_MouseDoubleClick({
        param ($sender, $e)
        # 行の上でのダブルクリックだけを対象にする（列見出し・スクロールバーは除く）
        $element = $e.OriginalSource
        while ($element -and !($element -is [System.Windows.Controls.DataGridRow])) {
            if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
                return
            }
            $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
        }
        if ($element) {
            & $openFolder $grid.SelectedItem
        }
    }.GetNewClosure())

    $labels = getIndexFileStateLabels
    $heading = $ctrl.IndexFilesHeading
    $failedText = ${function:getIndexFilesLoadFailedText}
    startJob {
        param ($path, $indexName, $labels)
        getIndexFileEntries -indexName $indexName -labels $labels -path $path
    } @($workspace.StatusFile, $IndexName, $labels) {
        param ($output, $errorText)
        if ($ctx.State.Closed) {
            # 待っている間に閉じた。結果は捨てる
            return
        }
        if ($errorText -or !$output -or $output.Count -eq 0) {
            $heading.Text = & $failedText
            return
        }
        & $applyEntries $ctx $output[0]
    }.GetNewClosure() (getWorkspaceJobQueue $workspace.StatusFile)

    [void](showOwnedDialog $dialog)
}
