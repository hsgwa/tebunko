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
    function continueStartIndexing {
        # 続きの中身（前の版のインデックスの確かめとインデクサの開始）は画面を使うため、受け取った引数だけを記録する
        param ($legacyState, [string[]]$onlyNames)
        $script:continued.Add(@{ LegacyState = $legacyState; OnlyNames = $onlyNames })
    }

    . "${scriptsDir}\tebunko\ui\indexing_view.ps1"

    # 画面の型を使うファイルは読み込まず、必要な関数だけを取り出す
    # （dot-source は、この場の scope で行う）
    foreach ($source in @(
            @{ File = "${scriptsDir}\tebunko\ui\indexing_tab.ps1"; Names = @("startIndexing", "finishWorkspaceCheck") }
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

Describe "getWorkspaceCheckingStatus / getWorkspaceUnreachableStatus" -Tag Unit {
    It "確かめている間は、接続できないときに時間がかかることを知らせる" {
        getWorkspaceCheckingStatus | Should -Be "ワークスペースを確かめています…（共有フォルダに接続できないときは、しばらくかかります）"
    }

    It "届かないときは、場所を添えて知らせる" {
        getWorkspaceUnreachableStatus "\\fileserver\共有\ws" | Should -Be "ワークスペースに接続できません：\\fileserver\共有\ws"
    }
}
