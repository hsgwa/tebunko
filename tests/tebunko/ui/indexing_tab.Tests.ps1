# インデックス作成の開始（tebunko\ui\indexing_tab.ps1 の startIndexing）のテスト。
# 画面と裏の列は偽物にする。ワークスペースが共有フォルダにある場面は、実在しない UNC のパスで表す。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    $script:statuses = New-Object System.Collections.Generic.List[string]
    $script:jobs = New-Object System.Collections.Generic.List[object]
    $script:continued = New-Object System.Collections.Generic.List[object]
    $script:buttonUpdates = 0

    function setStatus { param([string]$text) $script:statuses.Add($text) }
    function startJob {
        # 裏の列には出さず、依頼を取っておく（テストが onDone を呼んで、終わったことにする）
        param ([scriptblock]$scriptBlock, [object[]]$arguments, [scriptblock]$onDone, [string]$queue = "default")
        $script:jobs.Add(@{ ScriptBlock = $scriptBlock; Arguments = $arguments; OnDone = $onDone; Queue = $queue })
    }
    function updateIndexingButton { $script:buttonUpdates++ }
    function showMessage { param($text, $button, $kind) $script:statuses.Add("message:$text") }
    function selectScreen { param($name) }
    function showIndexingConfirmDialog { param ($plan) $script:dialogAnswer }
    function answerIndexingPlan { param ($channel, $answer) }
    function showIndexTextExtensionsDialog { param ($Owner, $IndexName, $Excluded, $Counts, $CountsSource, $CountsTime) $null }  # 拡張子を選ぶダイアログ（画面。Mock で差し替える）
    function continueStartIndexing {
        # 続きの中身（前の版のインデックスの確かめとインデクサの開始）は画面を使うため、受け取った引数だけを記録する
        param ($legacyState, [string[]]$onlyNames, $confirmText = $null)
        $script:continued.Add(@{ LegacyState = $legacyState; OnlyNames = $onlyNames; ConfirmText = $confirmText })
    }

    . "${scriptsDir}\tebunko\ui\indexing_view.ps1"

    # 画面の型を使うファイルは読み込まず、必要な関数だけを取り出す
    # （dot-source は、この場の scope で行う）
    foreach ($source in @(
            @{ File = "${scriptsDir}\tebunko\ui\indexing_tab.ps1"; Names = @("startIndexing", "finishWorkspaceCheck", "confirmIndexingTargets", "changeIndexingTextExtensions", "restartIndexingIfRequested", "startIndexingForRestart", "waitAndRestartIndexing", "getIndexingScanText", "finishIndexing", "finishImportConfirm", "applyImportConfirmText", "updateIndexingConfirmTotal") }
            @{ File = "${scriptsDir}\tebunko\ui\index\index_list.ps1"; Names = @("isIndexing", "isIndexingOrPreparing") }
        )) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($source.File, [ref]$null, [ref]$null)
        foreach ($functionName in $source.Names) {
            $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $functionName }, $true)
            . ([scriptblock]::Create($node.Extent.Text))
        }
    }

    function runJob {
        # 裏の仕事の中身を、この場の Mock が効くように作り直して（startJob と同じく文字列から）流す
        param ($job)
        $block = [scriptblock]::Create($job.ScriptBlock.ToString())
        $jobArguments = $job.Arguments
        & $block @jobArguments
    }
}

Describe "startIndexing" -Tag Unit {
    BeforeEach {
        $script:statuses.Clear()
        $script:jobs.Clear()
        $script:continued.Clear()
        $script:buttonUpdates = 0
        $script:indexingSession = $null
        $script:indexingPreparing = $false
        Mock getWorkspaceBlockMessage { "" }
    }

    Context "ネットワークのワークスペース" {
        BeforeEach {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws" }
        }

        It "場所に触れず、裏の列（network）に出して、確かめている間は他の操作を止める" {
            Mock getLegacyIndexState { throw "画面のスレッドで調べた" }

            startIndexing @("営業")

            Should -Invoke getLegacyIndexState -Times 0
            $script:jobs.Count | Should -Be 1
            $script:jobs[0].Queue | Should -Be "network"
            $script:indexingPreparing | Should -Be $true
            (isIndexingOrPreparing) | Should -Be $true
            $script:buttonUpdates | Should -BeGreaterThan 0
            $script:statuses[0] | Should -Be (getWorkspaceCheckingStatus)
            $script:continued.Count | Should -Be 0
        }

        It "確かめている間に、もう一度押しても、新しい依頼は出さない" {
            startIndexing
            startIndexing

            $script:jobs.Count | Should -Be 1
        }

        It "裏の仕事: 届かないときは、前の版のインデックスを調べずに知らせる" {
            Mock getPathState { @{ State = "Unreachable" } }
            Mock getLegacyIndexState { throw "届かない場所を調べた" }
            startIndexing

            $result = @(runJob $script:jobs[0])[0]

            $result.Unreachable | Should -Be $true
            Should -Invoke getLegacyIndexState -Times 0
        }

        It "裏の仕事: 届くときは、前の版のインデックスの状態を返す" {
            Mock getPathState { @{ State = "Found" } }
            Mock getLegacyIndexState { @{ HasLegacyIndex = $true; ContentEmpty = $false } }
            startIndexing

            $result = @(runJob $script:jobs[0])[0]

            $result.Unreachable | Should -Be $false
            $result.LegacyState.HasLegacyIndex | Should -Be $true
        }

        It "確かめが終わったら、止めていた操作を戻して、続きを始める（選んだ名前を引き継ぐ）" {
            startIndexing @("営業")

            & $script:jobs[0].OnDone @(@{ Unreachable = $false; LegacyState = @{ HasLegacyIndex = $false } }) ""

            $script:indexingPreparing | Should -Be $false
            $script:continued.Count | Should -Be 1
            $script:continued[0].OnlyNames | Should -Be @("営業")
            $script:statuses[$script:statuses.Count - 1] | Should -Be ""
        }

        It "届かなかったら、接続できないことを出して、始めない" {
            startIndexing

            & $script:jobs[0].OnDone @(@{ Unreachable = $true; LegacyState = $null }) ""

            $script:indexingPreparing | Should -Be $false
            $script:continued.Count | Should -Be 0
            $script:statuses[$script:statuses.Count - 1] | Should -Be "ワークスペースに接続できません：\\fileserver\共有\ws"
        }

        It "確かめに失敗したときも、始めない（止めていた操作は戻す）" {
            startIndexing

            & $script:jobs[0].OnDone @() "届きません"

            $script:indexingPreparing | Should -Be $false
            $script:continued.Count | Should -Be 0
            $script:statuses[$script:statuses.Count - 1] | Should -BeLike "ワークスペースに接続できません：*"
        }
    }

    Context "ローカルのワークスペース" {
        It "裏の仕事は出さず、その場で前の版のインデックスを調べて続きを始める" {
            $script:workspace = newTestWorkspace @{ Dir = "C:\Users\test\tebunko_ws" }
            Mock getLegacyIndexState { @{ HasLegacyIndex = $false; ContentEmpty = $true } }

            startIndexing @("営業")

            $script:jobs.Count | Should -Be 0
            $script:indexingPreparing | Should -Be $false
            $script:continued.Count | Should -Be 1
            $script:continued[0].OnlyNames | Should -Be @("営業")
        }
    }

    It "既定のワークスペースにほかのファイルがあるときは、確かめず始めない" {
        $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws" }
        Mock getWorkspaceBlockMessage { "ほかのファイルがあります" }

        startIndexing

        $script:jobs.Count | Should -Be 0
        $script:continued.Count | Should -Be 0
        $script:indexingPreparing | Should -Be $false
    }
}

Describe "startIndexing の confirmText" -Tag Unit {
    BeforeEach {
        $script:continued.Clear()
        $script:indexingSession = $null
        $script:indexingPreparing = $false
        $script:workspace = newTestWorkspace @{ Dir = "C:\ws" }
        Mock getWorkspaceBlockMessage { "" }
        Mock getWorkspaceJobQueue { "default" }
        Mock getLegacyIndexState { @{ HasLegacyIndex = $false; ContentEmpty = $true } }
    }

    It "渡した確認の文言を、続きにそのまま渡す（省けば渡らない）" {
        startIndexing -onlyNames @("営業") -confirmText @{ Title = "t" }
        startIndexing @("営業")

        $script:continued[0].ConfirmText.Title | Should -Be "t"
        $script:continued[1].ConfirmText | Should -BeNullOrEmpty
    }
}

Describe "updateIndexingConfirmTotal の主ボタン" -Tag Unit {
    BeforeEach {
        $names = @("StartButton", "TotalText", "SubfolderRemovalNote", "TextExcludeRemovalNote", "RetryCheck", "CloudWarnText", "CloudCheck", "CloudRememberCheck")
        $ctrl = @{}
        foreach ($name in $names) {
            $ctrl[$name] = [pscustomobject]@{ Content = ""; Text = ""; Visibility = ""; IsChecked = $false; IsEnabled = $false }
        }
        $script:confirmDialog = @{ Plan = @(); Ctrl = $ctrl; View = $null }
        $script:indexingOnlyNames = @("営業")
        Mock getIndexingConfirmView {
            @{ Confirm = @{ Text = "合計"; Button = "更新を開始" }; SubfolderRemovalNote = ""; TextExcludeRemovalNote = ""; RetryVisible = $false; RetryText = ""
                CloudVisible = $false; CloudWarn = ""; CloudCheckText = ""; Include = $false }
        }
    }

    It "<label>" -TestCases @(
        @{ label = "インポートの回は、何度書き直しても主ボタンは［作成する（更新を開始）］のまま"; text = @{ StartText = "作成する（更新を開始）" }; expected = "作成する（更新を開始）" }
        @{ label = "ふだんの更新は、判断層が決めた文言"; text = $null; expected = "更新を開始" }
    ) {
        param ($label, $text, $expected)
        $script:indexingConfirmText = $text

        updateIndexingConfirmTotal
        updateIndexingConfirmTotal

        $script:confirmDialog.Ctrl.StartButton.Content | Should -Be $expected
        $script:indexingConfirmText = $null
    }
}

Describe "applyImportConfirmText" -Tag Unit {
    It "窓の題・見出し・説明・［あとで］と、取り込んだ行のステータスだけを差し替える" {
        $heading = [pscustomobject]@{ Text = "インデックス更新の確認" }
        $hint = [pscustomobject]@{ Text = ""; Visibility = "Collapsed" }
        $dialog = [pscustomobject]@{ Title = "インデックス更新の確認"; Controls = @{ HeadingText = $heading; ImportIndexHint = $hint } } |
            Add-Member -MemberType ScriptMethod -Name FindName -Value { param ($name) $this.Controls[$name] } -PassThru
        $ctrl = @{ CancelButton = [pscustomobject]@{ Content = "キャンセル" } }
        $rows = @(
            [pscustomobject]@{ Name = "営業"; StatusText = "新規"; Level = "Ok" }
            [pscustomobject]@{ Name = "総務"; StatusText = "新規"; Level = "Ok" }
        )
        $text = @{ Name = "営業"; Title = "営業 をインポートしました"; Heading = "見出し"; Hint = "説明"; CancelText = "あとで"; RowStatus = "要更新" }

        applyImportConfirmText $dialog $ctrl $rows $text

        $dialog.Title | Should -Be "営業 をインポートしました"
        $heading.Text | Should -Be "見出し"
        $hint.Text | Should -Be "説明"
        $hint.Visibility | Should -Be "Visible"
        $ctrl.CancelButton.Content | Should -Be "あとで"
        $rows[0].StatusText | Should -Be "要更新"
        $rows[0].Level | Should -Be "Wait"
        $rows[1].StatusText | Should -Be "新規"
    }
}

Describe "finishIndexing（インポートの後の確認）" -Tag Unit {
    BeforeAll {
        function closeIndexingWait { }
        function rememberTextExtCounts { param ($plan) }
        function setIndexingBanner { param ($level, $kind) }
        function setIndexingBannerBehavior { }
        function getIndexingBannerLevel { param ($exitCode, $failed) "warn" }
        function getIndexingSkippedView { param ($skipped) $null }
        function refreshIndexingState { }
        function refreshIndexSummary { }
        function loadIndexTree { }
        function refreshLeftoverIndexes { }
        function restartIndexingIfRequested { param ($exitCode) }
        function closeIndexingBanner { $script:bannerClosed++ }
        function getIndexingProgress { $null }
    }

    BeforeEach {
        $script:statuses.Clear()
        $script:bannerClosed = 0
        $script:taskbar = [pscustomobject]@{ ProgressState = "" }
        $script:indexingTimer = [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Stop -Value { } -PassThru
        $script:ui = @{}
        foreach ($name in @("IndexingProgress", "IndexingProgressText", "IndexingProgressDetail", "IndexingProgressEta", "IndexingStopButton", "IndexingProgressPanel", "IndexingSearchButton")) {
            $script:ui[$name] = [pscustomobject]@{ IsIndeterminate = $false; Value = 0; Text = ""; Visibility = "" }
        }
        $script:indexingConfirmText = @{ SkippedStatus = "取りやめの文" }
        $script:indexingRestart = $null
    }

    It "拡張子を変えて取りやめた回（始め直しの印あり）は、取りやめの文を出さず、帯も閉じない" {
        $script:indexingCanceledAtConfirm = $true
        $script:indexingRestart = @{ OnlyNames = @("営業"); ConfirmText = $script:indexingConfirmText }
        $script:indexingSession = [pscustomobject]@{ Channel = @{ OnlySkipped = @(); Plan = @{} } } |
            Add-Member -MemberType ScriptMethod -Name GetExitCode -Value { 2 } -PassThru |
            Add-Member -MemberType ScriptMethod -Name GetError -Value { "" } -PassThru |
            Add-Member -MemberType ScriptMethod -Name Close -Value { } -PassThru |
            Add-Member -MemberType ScriptMethod -Name GetCloudSkipped -Value { @() } -PassThru

        finishIndexing

        $script:bannerClosed | Should -Be 0
        @($script:statuses).Count | Should -Be 0
    }

    It "<label>" -TestCases @(
        @{ label = "確認で取りやめたら、帯を閉じて取りやめの文をステータスバーに出し、文言を消す"; exit = 2; canceled = $true; closed = 1; statuses = @("取りやめの文") }
        @{ label = "更新を中止したときは、帯もステータスも触らず、文言を消す"; exit = 2; canceled = $false; closed = 0; statuses = @() }
        @{ label = "完了したときも、文言を消す"; exit = 0; canceled = $false; closed = 0; statuses = @() }
        @{ label = "続けられなかったときも、文言を消す（取りやめの文は出さず、エラーのメッセージだけ）"; exit = 1; canceled = $false; closed = 0; statuses = @("message:インデックスを更新できませんでした。`n`n詳しくはログを確認してください。") }
    ) {
        param ($label, $exit, $canceled, $closed, $statuses)
        $script:indexingCanceledAtConfirm = $canceled
        $code = $exit
        $script:indexingSession = [pscustomobject]@{ Channel = @{ OnlySkipped = @(); Plan = @{} } } |
            Add-Member -MemberType ScriptMethod -Name GetExitCode -Value { $code }.GetNewClosure() -PassThru |
            Add-Member -MemberType ScriptMethod -Name GetError -Value { "" } -PassThru |
            Add-Member -MemberType ScriptMethod -Name Close -Value { } -PassThru |
            Add-Member -MemberType ScriptMethod -Name GetCloudSkipped -Value { @() } -PassThru

        finishIndexing

        $script:indexingConfirmText | Should -BeNullOrEmpty
        $script:bannerClosed | Should -Be $closed
        @($script:statuses) | Should -Be @($statuses)
    }
}

Describe "confirmIndexingTargets" -Tag Unit {
    BeforeEach {
        $script:statuses.Clear()
        $script:indexingConfirmed = $false
        $script:indexingCanceledAtConfirm = $false
        $script:indexingRate = 1
        $script:indexingSession = @{ Channel = @{ Plan = @{} } }
        $script:ui = @{
            IndexingProgress     = [pscustomobject]@{ IsIndeterminate = $false }
            IndexingProgressText = [pscustomobject]@{ Text = "" }
            IndexingProgressDetail = [pscustomobject]@{ Text = "" }
            IndexingStopButton   = [pscustomobject]@{ IsEnabled = $true }
        }
    }

    It "更新を始める返事のときは、前の操作の結果の文を消す" {
        $script:dialogAnswer = @{ RetryFailed = $false; IncludeCloud = $false }

        confirmIndexingTargets

        $script:statuses.Count | Should -Be 1
        $script:statuses[0] | Should -Be ""
    }

    It "取りやめたときは、ステータスバーの文に触れない" {
        $script:dialogAnswer = $null

        confirmIndexingTargets

        $script:indexingCanceledAtConfirm | Should -Be $true
        $script:statuses.Count | Should -Be 0
    }
}

Describe "更新を始める道（continueStartIndexing から確認まで）のステータスバーの文" -Tag Unit {
    BeforeAll {
        # 続きの本物と、一覧の保存の本物を取り出し、画面と状態を触る口だけを偽物にして、道の全体を通す
        foreach ($source in @(
                @{ File = "${scriptsDir}\tebunko\ui\indexing_tab.ps1"; Names = @("continueStartIndexing") }
                @{ File = "${scriptsDir}\tebunko\ui\index\index_list.ps1"; Names = @("saveTargets") }
            )) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($source.File, [ref]$null, [ref]$null)
            foreach ($functionName in $source.Names) {
                $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $functionName }, $true)
                . ([scriptblock]::Create($node.Extent.Text))
            }
        }
        function writeTargetFolders { param ($folders) }
        function getTargetFolders { @() }
        function getTargetsKey { param ($folders) "" }
        function updateIndexDetailPanel { }
        function newIndexerChannel { param ($confirmTargets, $onlyNames) @{ Plan = @{} } }
        function newIndexingSession { param ($channel) @{ Channel = $channel } }
        function showIndexingPanel { }
    }

    BeforeEach {
        $script:statuses.Clear()
        $script:targetItems = @()
        $script:indexingTimer = [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Start -Value { } -PassThru
        $script:indexingConfirmed = $false
        $script:ui = @{
            IndexingProgress       = [pscustomobject]@{ IsIndeterminate = $false }
            IndexingProgressText   = [pscustomobject]@{ Text = "" }
            IndexingProgressDetail = [pscustomobject]@{ Text = "" }
            IndexingStopButton     = [pscustomobject]@{ IsEnabled = $true }
        }
        $legacy = @{ HasLegacyIndex = $false; ContentEmpty = $true }
        continueStartIndexing $legacy @()
    }

    It "確認で取りやめたら、ステータスバーに何も出さない（前の操作の結果を残す）" {
        $script:dialogAnswer = $null

        confirmIndexingTargets

        $script:statuses.Count | Should -Be 0
    }

    It "確認の文言（インポートの後）を渡した回はそれを持ち、渡さない回は持ち越さない" {
        $script:dialogAnswer = $null
        $text = @{ Title = "営業 をインポートしました" }
        continueStartIndexing @{ HasLegacyIndex = $false; ContentEmpty = $true } @() $text
        $script:indexingConfirmText.Title | Should -Be "営業 をインポートしました"

        continueStartIndexing @{ HasLegacyIndex = $false; ContentEmpty = $true } @()
        $script:indexingConfirmText | Should -BeNullOrEmpty
    }

    It "確認で進めたら、前の操作の結果の文を消す（空の 1 回だけ）" {
        $script:dialogAnswer = @{ RetryFailed = $false; IncludeCloud = $false }

        confirmIndexingTargets

        $script:statuses.Count | Should -Be 1
        $script:statuses[0] | Should -Be ""
    }
}

Describe "saveTargets" -Tag Unit {
    BeforeAll {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\index\index_list.ps1", [ref]$null, [ref]$null)
        $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "saveTargets" }, $true)
        . ([scriptblock]::Create($node.Extent.Text))
        function writeTargetFolders { param ($folders) }
        function getTargetFolders { @() }
        function getTargetsKey { param ($folders) "" }
    }

    BeforeEach {
        $script:statuses.Clear()
        $script:targetItems = @()
    }

    It "ふだんは、保存した旨をステータスバーに出す" {
        saveTargets

        $script:statuses.Count | Should -Be 1
        $script:statuses[0] | Should -BeLike "インデックス一覧を保存しました（*）"
    }

    It "-Quiet のときは、保存しても文を出さない" {
        saveTargets -Quiet

        $script:statuses.Count | Should -Be 0
    }
}

Describe "getWorkspaceCheckingStatus / getWorkspaceUnreachableStatus" -Tag Unit {
    It "確かめている間は、接続できないときに時間がかかることを知らせる" {
        getWorkspaceCheckingStatus | Should -Be "ワークスペースを確かめています…（共有フォルダに接続できないときは、しばらくかかります）"
    }

    It "届かないときは、場所を添えて知らせる" {
        getWorkspaceUnreachableStatus "\\fileserver\共有\ws" | Should -Be "ワークスペースに接続できません：\\fileserver\共有\ws"
    }
}

Describe "確認での取り込む拡張子の変更と、始め直し" -Tag Unit {
    BeforeEach {
        $script:indexingRestart = $null
        $script:indexingRecounting = $false
        $script:indexingOnlyNames = @("売上")
        $script:indexingSession = $null
        $script:indexingPreparing = $false
        $window = [pscustomobject]@{ DialogResult = $null }
        $planRow = [pscustomobject]@{ インデックス名 = "売上"; 区分 = ${planKindIngest}; TextExtensionCounts = @{ ".log" = 5 } }
        $script:confirmDialog = @{ Window = $window; Plan = @($planRow); TextExcludes = @{ "売上" = @(".log") } }
        Mock writeIndexTextExcludes { }
        Mock startIndexing { }
    }

    It "選びが変わった［OK］は、設定にすぐ書き、同じ対象での始め直しの印を立てて、確認を取りやめとして閉じる" {
        Mock showIndexTextExtensionsDialog { , @(".log", ".csv") }

        changeIndexingTextExtensions "売上"

        Should -Invoke writeIndexTextExcludes -Times 1 -ParameterFilter { $name -eq "売上" -and (@($excluded) -join ",") -eq ".log,.csv" }
        $script:indexingRestart.OnlyNames | Should -Be @("売上")
        $script:confirmDialog.Window.DialogResult | Should -Be $false
        # ダイアログには、今数えた件数と今の選びを渡す
        Should -Invoke showIndexTextExtensionsDialog -Times 1 -ParameterFilter { $IndexName -eq "売上" -and $CountsSource -eq "now" -and $Counts[".log"] -eq 5 -and (@($Excluded) -join ",") -eq ".log" }
    }

    It "インポートの後の確認で選びを変えたら、始め直しの印に確認の文言を持たせ、始め直しの startIndexing にそのまま渡す" {
        Mock showIndexTextExtensionsDialog { , @(".log", ".csv") }
        $script:indexingConfirmText = @{ Title = "売上 をインポートしました" }

        changeIndexingTextExtensions "売上"
        restartIndexingIfRequested 2

        Should -Invoke startIndexing -Times 1 -ParameterFilter { $confirmText.Title -eq "売上 をインポートしました" }
        $script:indexingConfirmText = $null
    }

    It "ふだんの更新の確認で選びを変えたら、始め直しに確認の文言を渡さない" {
        Mock showIndexTextExtensionsDialog { , @(".log", ".csv") }
        $script:indexingConfirmText = $null

        changeIndexingTextExtensions "売上"
        restartIndexingIfRequested 2

        Should -Invoke startIndexing -Times 1 -ParameterFilter { $null -eq $confirmText }
    }

    It "全部戻す（空の配列）［OK］も、選びが変わったものとして書く" {
        Mock showIndexTextExtensionsDialog { , @() }

        changeIndexingTextExtensions "売上"

        Should -Invoke writeIndexTextExcludes -Times 1
        $script:indexingRestart | Should -Not -BeNullOrEmpty
    }

    It "［キャンセル］と、選びが変わらない［OK］では、何もしない" -TestCases @(
        @{ result = "cancel" }
        @{ result = "same" }
    ) {
        param ($result)
        if ($result -eq "cancel") { Mock showIndexTextExtensionsDialog { $null } } else { Mock showIndexTextExtensionsDialog { , @(".LOG") } }

        changeIndexingTextExtensions "売上"

        Should -Invoke writeIndexTextExcludes -Times 0
        $script:indexingRestart | Should -BeNullOrEmpty
        $script:confirmDialog.Window.DialogResult | Should -BeNullOrEmpty
    }

    It "取りやめた回が終わったら、同じ onlyNames で startIndexing を 1 回だけ呼び、印は消える（2 回目は呼ばない）" {
        $script:indexingRestart = @{ OnlyNames = @("売上", "経理") }

        restartIndexingIfRequested 2
        restartIndexingIfRequested 2

        Should -Invoke startIndexing -Times 1 -ParameterFilter { (@($onlyNames) -join ",") -eq "売上,経理" }
        $script:indexingRestart | Should -BeNullOrEmpty
    }

    It "一覧の読み込みが続いている（targetsNaming。ネットワークのワークスペース）間は、すぐには呼ばず、待つ関数に任せる（更新の流れが消えない）" {
        Mock waitAndRestartIndexing {}
        $script:indexingRestart = @{ OnlyNames = @("売上") }
        $script:targetsNaming = $true
        try {
            restartIndexingIfRequested 2
        } finally {
            $script:targetsNaming = $false
        }

        Should -Invoke startIndexing -Times 0
        Should -Invoke waitAndRestartIndexing -Times 1 -ParameterFilter { (@($onlyNames) -join ",") -eq "売上" }
        $script:indexingRestart | Should -BeNullOrEmpty
    }

    It "エラーで終わった回では始め直さず、印を消す" {
        $script:indexingRestart = @{ OnlyNames = @("売上") }

        restartIndexingIfRequested 1

        Should -Invoke startIndexing -Times 0
        $script:indexingRestart | Should -BeNullOrEmpty
        $script:indexingRecounting | Should -Be $false
    }

    It "印が無ければ何も呼ばない" {
        restartIndexingIfRequested 0

        Should -Invoke startIndexing -Times 0
    }

    It "startIndexing が断られた（インデクサも準備中も無いまま）ときは、印も数えの印も消えて、もう呼ばない" {
        $script:indexingRestart = @{ OnlyNames = @() }

        restartIndexingIfRequested 2
        restartIndexingIfRequested 2

        Should -Invoke startIndexing -Times 1
        $script:indexingRestart | Should -BeNullOrEmpty
        $script:indexingRecounting | Should -Be $false
    }

    It "始め直しの数えの間だけ、帯に数えたファイル数を出す（ふだんは今までの文。口が無いときは読まない）" {
        $script:indexingSession = @{ Channel = @{ CountedFiles = 1240 } }
        getIndexingScanText | Should -Be "更新するファイルを調べています…"
        $script:indexingRecounting = $true
        getIndexingScanText | Should -Be "数えています…（1,240 ファイル）"
        $script:indexingSession = @{ Channel = @{ CountedFiles = 0 } }
        getIndexingScanText | Should -Be "数えています…"
        $script:indexingSession = $null
        getIndexingScanText | Should -Be "更新するファイルを調べています…"
    }
}
