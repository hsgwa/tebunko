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
        $row.PlannedStatusText = "要更新"; $row.PlannedLevel = "Wait"; $row.Choosable = $choosable
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

Describe "数え直しの流れ（applyIndexingRecount）" -Tag Unit {
    BeforeAll {
        # 呼び先は記録するだけの偽物にする（本物は設定を書き、インデクサを始める）
        function saveTargets { param ([switch]$Quiet) $script:recountLog.Add("save:" + (($script:targetItems | ForEach-Object { "$($_.Name)=$($_.Subfolders)" }) -join ",")) }
        function updateIndexDetailPanel { $script:recountLog.Add("detail") }
        function updateIndexingButton { $script:recountLog.Add("button") }
        function startIndexing { param ([string[]]$onlyNames = @()) $script:recountLog.Add("start:" + ($onlyNames -join ",")) }
    }

    It "選びを行に入れて設定へ書いてから、同じ対象（選んだ名前）で始め直す。選ばなかった行は変えない" -Skip:(!$sta) {
        $script:recountLog = New-Object System.Collections.Generic.List[string]
        $a = New-Object FolderItem; $a.Name = "営業"
        $b = New-Object FolderItem; $b.Name = "顧客"
        $script:targetItems = @($a, $b)

        applyIndexingRecount @{ OnlyNames = @("営業"); Changes = @(@{ Name = "営業"; Subfolders = $false }) }

        @($script:recountLog) | Should -Be @("save:営業=False,顧客=True", "detail", "button", "start:営業")
        $a.Subfolders | Should -BeFalse
        $a.TopOnlyText | Should -Be "直下のみ"
        $a.SubfoldersChanged | Should -BeTrue
        $b.Subfolders | Should -BeTrue
    }

    It "［すべて更新］のとき（選んだ名前が空）は、空のまま始め直す" -Skip:(!$sta) {
        $script:recountLog = New-Object System.Collections.Generic.List[string]
        $a = New-Object FolderItem; $a.Name = "営業"
        $script:targetItems = @($a)

        applyIndexingRecount @{ OnlyNames = @(); Changes = @(@{ Name = "営業"; Subfolders = $false }) }

        @($script:recountLog)[-1] | Should -Be "start:"
    }
}