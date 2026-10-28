# 更新の確認ダイアログの「サブフォルダを含める」（行ごとのチェック）のテスト。
# 実際のダイアログの xaml を同じプロセスの中で読み込み、チェックを 1 回切り替えて、ハンドラが何度も呼ばれ続けない（一覧全体を描き直すと、
# チェックが入れ直されて Checked が上がり続ける）こと、行・合計・説明・主ボタンが数え直しの見せ方に替わることを確かめる。
# WPF の要素は STA のスレッドでしか作れないため、STA でないときは飛ばす。
BeforeDiscovery {
    $sta = [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq "STA"
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\app_host.ps1"
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_types.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\indexing_view.ps1"
    # indexing_tab.ps1 は読み込むときに時計を作る（newTimer）。このテストでは動かさないので、偽物を渡す
    function newTimer { param ($ms, $action) return [pscustomobject]@{} }
    # 読み込むときにつなぐ画面の部品（本物の部品を作るだけで、画面は出さない）。STA のときだけ作れる
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq "STA") {
        $ui = @{
            FailedGrid = New-Object System.Windows.Controls.DataGrid
            IndexFilesGrid = New-Object System.Windows.Controls.DataGrid
            IndexFilesButton = New-Object System.Windows.Controls.Button
            ShowFailedFilesButton = New-Object System.Windows.Controls.Button
            IndexFilesAllRadio = New-Object System.Windows.Controls.RadioButton
            IndexFilesFailedRadio = New-Object System.Windows.Controls.RadioButton
            IndexFilesCloseButton = New-Object System.Windows.Controls.Button
            IndexingButton = New-Object System.Windows.Controls.Button
            IndexingStopButton = New-Object System.Windows.Controls.Button
            IndexingResumeButton = New-Object System.Windows.Controls.Button
            IndexingSearchButton = New-Object System.Windows.Controls.Button
        }
    } else {
        $ui = @{}
    }
    . "${scriptsDir}\tebunko\ui\indexing_tab.ps1"

    # 画面の要素の下にあるチェックボックスを、見える順に集める（ハンドラの中からも呼べるよう、関数でなくスクリプトブロックにする）
    $script:findCheckBoxes = {
        param ($element)
        $found = New-Object System.Collections.Generic.List[object]
        $stack = New-Object System.Collections.Generic.Stack[object]
        $stack.Push($element)
        while ($stack.Count -gt 0) {
            $node = $stack.Pop()
            if ($node -is [System.Windows.Controls.CheckBox]) { $found.Add($node) }
            $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($node)
            for ($k = $count - 1; $k -ge 0; $k--) { $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($node, $k)) }
        }
        return $found.ToArray()
    }
    function setRecountSession {
        # 数え直しの関数が読む状態（確認ダイアログの状態・動いているインデックス作成の受け渡しの口）は、その関数と同じスコープに置く
        param ($channel, $dialogState)
        $script:indexingSession = [pscustomobject]@{ Channel = $channel }
        $script:confirmDialog = $dialogState
        $script:indexingOnlyNames = @()
        $script:indexingConfirmedChoices = $null
    }

    function setConfirmedChoices { param ($choices) $script:indexingConfirmedChoices = $choices }

    function getConfirmedChoices { return $script:indexingConfirmedChoices }

    function getConfirmDialogState { return $script:confirmDialog }
    function setConfirmDialogState {
        # indexing_tab.ps1 の関数が読む状態は、その関数と同じスコープ（この BeforeAll）に置く
        param ($dialogState)
        $script:confirmDialog = $dialogState
        $script:indexingOnlyNames = @()
    }

    function newPlanRow {
        param ([string]$name, [bool]$choosable = $true)
        $row = [PlanRow]::new()
        $row.Name = $name; $row.Path = "C:\data\$name"; $row.TotalText = "3"; $row.StatusText = "要更新"; $row.Level = "Wait"
        $row.PlannedStatusText = "要更新"; $row.PlannedLevel = "Wait"; $row.Choosable = $choosable; $row.Editable = $choosable
        return $row
    }
}

Describe "確認ダイアログの「サブフォルダを含める」" -Tag Unit {
    It "チェックを 1 回外すと、ハンドラは繰り返さずに数回で止まり、行・合計・説明・主ボタンが数え直しの見せ方になる" -Skip:(!$sta) {
        $dialog = loadXaml "${scriptsDir}\tebunko\xaml\dialog_indexing_confirm.xaml"
        $ctrl = @{}
        foreach ($name in @("StartButton", "TotalText", "PlanGrid", "RecountHint", "RecountHintText")) { $ctrl[$name] = $dialog.FindName($name) }
        $rows = @((newPlanRow "営業"), (newPlanRow "顧客"))
        $ctrl.PlanGrid.ItemsSource = $rows
        setConfirmDialogState @{ Window = $dialog; Ctrl = $ctrl }
        # ハンドラの中は別のスコープで動くため、結果は共有の入れ物（ハッシュ）に書く
        $state = @{ Calls = 0; CallsBefore = 0; BoxCount = 0; Problem = $null; Dialog = $dialog }
        $find = $script:findCheckBoxes
        $update = ${function:updateIndexingRecountView}
        connectPlanGridChoice $ctrl.PlanGrid { $state.Calls++; & $update }.GetNewClosure()
        $state.Timer = New-Object System.Windows.Threading.DispatcherTimer
        $state.Timer.Interval = [TimeSpan]::FromMilliseconds(1500)
        $state.Timer.Add_Tick({ param ($sender, $e) $sender.Stop(); $state.Dialog.Close() }.GetNewClosure())
        $dialog.Add_ContentRendered({
            try {
                $boxes = @(& $find $ctrl.PlanGrid)
                $state.BoxCount = $boxes.Count
                $state.CallsBefore = $state.Calls
                $boxes[0].IsChecked = $false
            } catch {
                $state.Problem = $_.Exception.Message
            }
            # 描き直しの連鎖が続くなら、この間に呼び出しが増え続ける
            $state.Timer.Start()

        }.GetNewClosure())
        $null = $dialog.ShowDialog()

        $state.Problem | Should -BeNullOrEmpty
        $state.BoxCount | Should -Be 2
        # 切り替えによる呼び出しの数（読み込みのときの通知は除く）
        ($state.Calls - $state.CallsBefore) | Should -BeLessOrEqual 3
        $rows[0].Subfolders | Should -BeFalse
        $rows[0].Changed | Should -BeTrue
        $rows[0].TopOnlyText | Should -Be "直下のみ"
        $rows[0].StatusText | Should -Be "取り込み直し"
        $rows[1].Changed | Should -BeFalse
        $rows[1].StatusText | Should -Be "要更新"
        $ctrl.StartButton.Content | Should -Be "数え直して続ける"
        $ctrl.TotalText.Text | Should -Be "更新対象: 2 フォルダ（件数は数え直したあとに出ます）"
        $ctrl.RecountHintText.Text | Should -Be "選びを変えたので、件数を数え直します"
        $ctrl.RecountHint.Visibility | Should -Be "Visible"
    }

    It "削除予定の行（Choosable が偽）のチェックは使えない" -Skip:(!$sta) {
        $dialog = loadXaml "${scriptsDir}\tebunko\xaml\dialog_indexing_confirm.xaml"
        $grid = $dialog.FindName("PlanGrid")
        $grid.ItemsSource = @((newPlanRow "営業"), (newPlanRow "消える" $false))
        $state = @{ Enabled = @() }
        $find = $script:findCheckBoxes
        $dialog.Add_ContentRendered({
            $state.Enabled = @(& $find $grid | ForEach-Object { $_.IsEnabled })
            $dialog.Close()
        }.GetNewClosure())
        $null = $dialog.ShowDialog()
        $state.Enabled | Should -Be @($true, $false)
    }
}

Describe "数え直しの流れ（確認を開いたまま）" -Tag Unit {
    BeforeAll {
        # 呼び先は記録するだけの偽物にする（本物は設定を書く）
        function saveTargets { param ([switch]$Quiet) $script:recountLog.Add("save:" + (($script:targetItems | ForEach-Object { "$($_.Name)=$($_.Subfolders)" }) -join ",")) }
        function updateIndexDetailPanel { $script:recountLog.Add("detail") }
        function isIndexing { return $true }
    }

    It "終わり方が <label> のとき、選びは行に入れて設定へ書き、［要更新］の印は <marked>（applyConfirmedChoices）" -Skip:(!$sta) -TestCases @(
        @{ label = "0（その選びで取り込み終えた）"; exitCode = 0; marked = $false }
        @{ label = "2（取りやめ。取り込んでいない）"; exitCode = 2; marked = $true }
        @{ label = "1（エラー）"; exitCode = 1; marked = $true }
    ) {
        param ($label, $exitCode, $marked)
        $script:recountLog = New-Object System.Collections.Generic.List[string]
        $a = New-Object FolderItem; $a.Name = "営業"
        $b = New-Object FolderItem; $b.Name = "顧客"
        $script:targetItems = @($a, $b)

        applyConfirmedChoices @(@{ Name = "営業"; Subfolders = $false }) $exitCode

        @($script:recountLog) | Should -Be @("save:営業=False,顧客=True", "detail")
        $a.TopOnlyText | Should -Be "直下のみ"
        $a.SubfoldersChanged | Should -Be $marked
        $b.Subfolders | Should -BeTrue
    }

    It "終わり方が <label> のとき、その回の行（<round>）の［要更新］の印は、選びを変えていなくても <cleared>（writeConfirmedChoices）" -Skip:(!$sta) -TestCases @(
        @{ label = "0（取り込み終えた）・全部の回"; exitCode = 0; round = @(); cleared = @("営業", "顧客") }
        @{ label = "0（取り込み終えた）・選んだ回"; exitCode = 0; round = @("営業"); cleared = @("営業") }
        @{ label = "2（取りやめ）"; exitCode = 2; round = @(); cleared = @() }
        @{ label = "1（エラー）"; exitCode = 1; round = @(); cleared = @() }
    ) {
        param ($label, $exitCode, $round, $cleared)
        $script:recountLog = New-Object System.Collections.Generic.List[string]
        $items = @("営業", "顧客") | ForEach-Object { $item = New-Object FolderItem; $item.Name = $_; $item.Enabled = $true; $item.SubfoldersChanged = $true; $item }
        $script:targetItems = @($items)
        setConfirmedChoices $null   # 前の回の取りやめで印が立ったまま。今回は選びを変えていない

        writeConfirmedChoices $exitCode $round

        @($items | Where-Object { -not $_.SubfoldersChanged } | ForEach-Object { $_.Name }) | Should -Be $cleared
    }

    It "writeConfirmedChoices: 控えがあれば書く（取りやめで終わったときも）・控えが無ければ書かない。控えは 1 回で消える" -Skip:(!$sta) -TestCases @(
        @{ label = "取り込み終えた（0）"; exitCode = 0; choices = @(@{ Name = "営業"; Subfolders = $false }); saved = $true }
        @{ label = "取りやめ（2）"; exitCode = 2; choices = @(@{ Name = "営業"; Subfolders = $false }); saved = $true }
        @{ label = "控えが無い（確認で選びを変えずに終わった）"; exitCode = 2; choices = $null; saved = $false }
    ) {
        param ($label, $exitCode, $choices, $saved)
        $script:recountLog = New-Object System.Collections.Generic.List[string]
        $a = New-Object FolderItem; $a.Name = "営業"
        $script:targetItems = @($a)
        setConfirmedChoices $choices

        writeConfirmedChoices $exitCode

        (@($script:recountLog | Where-Object { $_ -like "save:*" }).Count -gt 0) | Should -Be $saved
        $a.Subfolders | Should -Be (-not $saved)
        (getConfirmedChoices) | Should -BeNullOrEmpty
    }

    It "［数え直して続ける］: 開いたまま全部の行が数え中になり、返事（Recount）をインデクサに返す。新しい Plan が届くと、一覧が差し替わって主ボタンが［更新を開始］に戻る" -Skip:(!$sta) {
        $dialog = loadXaml "${scriptsDir}\tebunko\xaml\dialog_indexing_confirm.xaml"
        $ctrl = @{}
        foreach ($name in @("StartButton", "TotalText", "PlanGrid", "RecountHint", "RecountHintText", "RetryCheck", "CloudWarnText", "CloudCheck", "CloudRememberCheck")) { $ctrl[$name] = $dialog.FindName($name) }
        $rows = @((newPlanRow "営業"), (newPlanRow "顧客"))
        $ctrl.PlanGrid.ItemsSource = $rows
        $channel = newIndexerChannel
        setRecountSession $channel @{ Window = $dialog; Ctrl = $ctrl; Plan = @() }

        startIndexingRecount (getConfirmDialogState) @(@{ Name = "営業"; Subfolders = $false })

        $channel.Answer.Recount | Should -BeTrue
        $channel.Answered.WaitOne(0) | Should -BeTrue
        @($rows | ForEach-Object { $_.StatusText }) | Should -Be @("数え中", "数え中")
        @($rows | ForEach-Object { $_.TotalText }) | Should -Be @("…", "…")
        @($rows | ForEach-Object { $_.Editable }) | Should -Be @($false, $false)
        $ctrl.StartButton.Content | Should -Be "数え直しています…"
        $ctrl.StartButton.IsEnabled | Should -BeFalse
        $ctrl.TotalText.Text | Should -Be "更新対象: 数え直しています…（［キャンセル］で取りやめ）"
        # 新しい Plan が届く（PlanVersion が上がる）まで、一覧は差し替えない
        checkIndexingRecountResult
        $ctrl.PlanGrid.ItemsSource | Should -Be $rows

        $channel.Plan = @((newIngestPlanRow "営業" "C:\data\営業" ${planKindIngest} 2 1 1 0 0 0 0), (newIngestPlanRow "顧客" "C:\data\顧客" ${planKindIngest} 5 0 0 0 0 0 0))
        $channel.PlanVersion = 1
        checkIndexingRecountResult

        $newRows = @($ctrl.PlanGrid.ItemsSource)
        @($newRows | ForEach-Object { $_.TotalText }) | Should -Be @("2", "5")
        @($newRows | ForEach-Object { $_.Name }) | Should -Be @("営業", "顧客")
        $ctrl.StartButton.IsEnabled | Should -BeTrue
        $ctrl.StartButton.Content | Should -Be "更新を開始"
        $script:indexingConfirmedChoices.Count | Should -Be 1
        $script:indexingConfirmedChoices[0].Name | Should -Be "営業"
    }
}