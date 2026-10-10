# インデックス一覧の読み込み（tebunko\ui\index\index_list.ps1 の loadTargets）のテスト。
# 画面と裏の列は偽物にする。ワークスペースが共有フォルダにある場面は、実在しない UNC のパスで表す。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"

    $script:statuses = New-Object System.Collections.Generic.List[string]
    $script:jobs = New-Object System.Collections.Generic.List[object]
    $script:calls = New-Object System.Collections.Generic.List[string]
    $script:messages = New-Object System.Collections.Generic.List[string]

    function setStatus { param([string]$text) $script:statuses.Add($text) }
    function startJob {
        # 裏の列には出さず、依頼を取っておく（テストが onDone を呼んで、終わったことにする）
        param ([scriptblock]$scriptBlock, [object[]]$arguments, [scriptblock]$onDone, [string]$queue = "default")
        $script:jobs.Add(@{ ScriptBlock = $scriptBlock; Arguments = $arguments; OnDone = $onDone; Queue = $queue })
    }
    function updateIndexListView { $script:calls.Add("updateIndexListView") }
    function updateIndexingButton { $script:calls.Add("updateIndexingButton") }
    function isIndexing { return $false }
    function showMessage { param($text, $buttons, $icon) $script:messages.Add($text); return "OK" }
    function refreshFolderStatus { $script:calls.Add("refreshFolderStatus") }
    function refreshFastSearchStatus { $script:calls.Add("refreshFastSearchStatus") }
    function newFolderItem {
        param ([string]$path, [bool]$enabled, [string]$name)
        [pscustomobject]@{ Path = $path; Enabled = $enabled; Name = $name; Checked = $false }
    }
    function getTargetsKey { param($folders) "key" }
    function getIndexCheckedItems { param($items) @($items | Where-Object { $_.Checked }) }

    . "${scriptsDir}\tebunko\ui\index_view.ps1"

    # index_list.ps1 は画面の型を使うため、読み込まずに必要な関数だけを取り出す
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\index\index_list.ps1", [ref]$null, [ref]$null)
    foreach ($name in @("loadTargets", "finishTargetsNaming", "fillTargetItems", "getCurrentIndexJobBlocker", "isIndexingOrPreparing", "testIndexOperable")) {
        $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        . ([scriptblock]::Create($node.Extent.Text))
    }

    function runJob {
        # 裏の仕事の中身を、この場の Mock が効くように作り直して（startJob と同じく文字列から）流す
        param ($job)
        $block = [scriptblock]::Create($job.ScriptBlock.ToString())
        $jobArguments = $job.Arguments
        & $block @jobArguments
    }
}

Describe "loadTargets" -Tag Unit {
    BeforeEach {
        $script:statuses.Clear()
        $script:jobs.Clear()
        $script:calls.Clear()
        $script:targetsLoadRequest = @{ Value = 0 }
        $script:targetItems = New-Object "System.Collections.ObjectModel.ObservableCollection[object]"
        $script:loadingTargets = $false
    }

    Context "名前の決まっていないインデックスが無いとき" {
        It "裏の仕事は出さず、設定の一覧を並べる（チェックは名前で引き継ぐ）" {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws" }
            $script:targetItems.Add([pscustomobject]@{ Name = "営業"; Checked = $true })
            Mock getTargetFolders { @([pscustomobject]@{ Name = "営業"; Path = "\\fileserver\共有\営業"; Enabled = $true }, [pscustomobject]@{ Name = "技術"; Path = "\\fileserver\共有\技術"; Enabled = $true }) }

            loadTargets

            $script:jobs.Count | Should -Be 0
            $script:targetItems.Count | Should -Be 2
            $script:targetItems[0].Checked | Should -Be $true
            $script:targetItems[1].Checked | Should -Be $false
        }
    }

    Context "名前の決まっていないインデックスがあるとき" {
        BeforeEach {
            $script:unnamed = @([pscustomobject]@{ Name = ""; Path = "\\fileserver\共有\営業"; Enabled = $true })
            $script:named = @([pscustomobject]@{ Name = "営業"; Path = "\\fileserver\共有\営業"; Enabled = $true })
            $script:namedNow = $false
            Mock getTargetFolders { if ($script:namedNow) { $script:named } else { $script:unnamed } }
        }

        It "ネットワークのワークスペース: 場所に触れず、一覧を空にして、裏の列（network）で名前を決める" {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws"; StatusFile = "\\fileserver\共有\ws\work\ingest_status.tsv" }
            Mock readStatusFile { throw "画面のスレッドで取り込み一覧を読んだ" }
            Mock saveAssignedIndexNames { throw "画面のスレッドで保存した" }

            loadTargets

            Should -Invoke readStatusFile -Times 0
            Should -Invoke saveAssignedIndexNames -Times 0
            $script:jobs.Count | Should -Be 1
            $script:jobs[0].Queue | Should -Be "network"
            $script:targetItems.Count | Should -Be 0
            $script:statuses[0] | Should -Be (getIndexNamingStatus)
        }

        It "ネットワークのワークスペース: 裏の仕事は、取り込み一覧を読んで名前を割り当てて保存する" {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws"; StatusFile = "\\fileserver\共有\ws\work\ingest_status.tsv" }
            loadTargets
            Mock readStatusFile { @{ Folders = @(); Rows = @{} } }
            Mock assignIndexNames { @($folders) }
            Mock saveAssignedIndexNames { }

            runJob $script:jobs[0]

            Should -Invoke readStatusFile -Times 1 -Exactly -ParameterFilter { $path -eq "\\fileserver\共有\ws\work\ingest_status.tsv" }
            Should -Invoke assignIndexNames -Times 1 -Exactly
            Should -Invoke saveAssignedIndexNames -Times 1 -Exactly
        }

        It "ネットワークのワークスペース: 名前を決めている間は、追加・削除などの操作を止める（空の一覧を保存して、ほかのインデックスを消さないため）" {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws" }
            loadTargets

            $script:targetsNaming | Should -Be $true
            (getCurrentIndexJobBlocker) | Should -Be "インデックス名の決定中"
            testIndexOperable "追加" | Should -Be $false
            testIndexOperable "削除" | Should -Be $false
            $script:messages[0] | Should -Be (getIndexJobBlockedMessage "インデックス名の決定中" "追加")
        }

        It "ネットワークのワークスペース: 結果が届いたら、操作の停止を解く（失敗したときも）" {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws" }
            loadTargets
            $script:namedNow = $true

            & $script:jobs[0].OnDone @() ""

            $script:targetsNaming | Should -Be $false
            (getCurrentIndexJobBlocker) | Should -Be ""
            testIndexOperable "追加" | Should -Be $true

            $script:namedNow = $false
            loadTargets
            $script:targetsNaming | Should -Be $true
            & $script:jobs[1].OnDone @() "届きません"

            $script:targetsNaming | Should -Be $false
        }

        It "ネットワークのワークスペース: 結果が届いてから、名前の付いた一覧を並べる" {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws" }
            loadTargets
            $script:namedNow = $true

            & $script:jobs[0].OnDone @() ""

            $script:targetItems.Count | Should -Be 1
            $script:targetItems[0].Name | Should -Be "営業"
        }

        It "ネットワークのワークスペース: 失敗したら理由を出し、設定のままの一覧を並べる" {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws" }
            loadTargets

            & $script:jobs[0].OnDone @() "届きません"

            $script:statuses[$script:statuses.Count - 1] | Should -Be (getIndexNamingFailedStatus "届きません")
            $script:targetItems.Count | Should -Be 1
        }

        It "ネットワークのワークスペース: 待っている間にもう一度読み込んだら、前の結果は捨てる" {
            $script:workspace = newTestWorkspace @{ Dir = "\\fileserver\共有\ws" }
            loadTargets
            loadTargets
            $script:namedNow = $true

            & $script:jobs[0].OnDone @() ""
            $script:targetItems.Count | Should -Be 0
            & $script:jobs[1].OnDone @() ""

            $script:targetItems.Count | Should -Be 1
        }

        It "ローカルのワークスペース: その場で名前を決めて、並べる（裏の仕事は出さない）" {
            $script:workspace = newTestWorkspace @{ Dir = "C:\Users\test\tebunko_ws" }
            Mock readStatusFile { @{ Folders = @(); Rows = @{} } }
            Mock assignIndexNames { @($folders) }
            Mock saveAssignedIndexNames { $script:namedNow = $true }

            loadTargets

            $script:jobs.Count | Should -Be 0
            Should -Invoke saveAssignedIndexNames -Times 1 -Exactly
            $script:targetItems[0].Name | Should -Be "営業"
        }
    }
}
