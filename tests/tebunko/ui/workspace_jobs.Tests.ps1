# 起動と切り替えで使うワークスペースの調べ（tebunko\ui\workspace_jobs.ps1）のテスト。
# 画面と裏の列は偽物にする。ネットワークの場所は、実在しない UNC のパスで表す（調べる関数は Mock にして、触れないことも確かめる）。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    $fake = @{ Statuses = (New-Object System.Collections.Generic.List[string]); Jobs = (New-Object System.Collections.Generic.List[object]) }
    function setStatus {
        param ([string]$text)
        $fake.Statuses.Add($text)
    }
    function startJob {
        # 裏の列には出さず、依頼を取っておく（テストが onDone を呼んで、終わったことにする）
        param ([scriptblock]$scriptBlock, [object[]]$arguments, [scriptblock]$onDone, [string]$queue = "default")
        $fake.Jobs.Add(@{ ScriptBlock = $scriptBlock; Arguments = $arguments; OnDone = $onDone; Queue = $queue })
    }

    . "${scriptsDir}\tebunko\ui\workspace_jobs.ps1"
}

Describe "testStartupIndexExists" -Tag Unit {
    It "ネットワークの場所は、場所に触れずに真を返す（［検索］を開く）" {
        Mock testIndexExists { throw "届かない場所を調べた" }

        testStartupIndexExists "\\fileserver\共有\tebunko_ws" | Should -Be $true
        Should -Invoke testIndexExists -Times 0
    }

    It "ローカルの場所は、testIndexExists の結果を返す" {
        Mock testIndexExists { $false }

        testStartupIndexExists "C:\Users\test\tebunko_ws" | Should -Be $false
        Should -Invoke testIndexExists -Times 1
    }
}

Describe "refreshLegacyIndexMessage" -Tag Unit {
    BeforeEach {
        $fake.Statuses.Clear()
        $fake.Jobs.Clear()
        $script:workspaceBlock = ""
        $script:legacyIndexMessage = ""
    }

    It "ローカルの場所は、その場で調べて知らせを覚える（ステータスには出さない）" {
        $script:workspace = [pscustomobject]@{ Dir = "C:\Users\test\tebunko_ws" }
        Mock getLegacyIndexState { @{ HasLegacyIndex = $true } }
        Mock getLegacyIndexMessage { "前の版のインデックスがあります" }

        refreshLegacyIndexMessage

        $script:legacyIndexMessage | Should -Be "前の版のインデックスがあります"
        $fake.Jobs.Count | Should -Be 0
        $fake.Statuses.Count | Should -Be 0
    }

    It "ネットワークの場所は、場所に触れず、裏の列（network）に出す" {
        $script:workspace = [pscustomobject]@{ Dir = "\\fileserver\共有\tebunko_ws" }
        Mock getLegacyIndexState { throw "届かない場所を調べた" }

        refreshLegacyIndexMessage

        Should -Invoke getLegacyIndexState -Times 0
        $fake.Jobs.Count | Should -Be 1
        $fake.Jobs[0].Queue | Should -Be "network"
        $script:legacyIndexMessage | Should -Be ""
    }

    It "裏で見つかったら、知らせを覚えてステータスに出す" {
        $script:workspace = [pscustomobject]@{ Dir = "\\fileserver\共有\tebunko_ws" }
        refreshLegacyIndexMessage

        & $fake.Jobs[0].OnDone @("前の版のインデックスがあります") ""

        $script:legacyIndexMessage | Should -Be "前の版のインデックスがあります"
        $fake.Statuses | Should -Be @("前の版のインデックスがあります")
    }

    It "裏で見つからなければ（空）、ステータスに出さない" {
        $script:workspace = [pscustomobject]@{ Dir = "\\fileserver\共有\tebunko_ws" }
        refreshLegacyIndexMessage

        & $fake.Jobs[0].OnDone @("") ""

        $fake.Statuses.Count | Should -Be 0
    }

    It "裏の調べが失敗したら、何も出さない" {
        $script:workspace = [pscustomobject]@{ Dir = "\\fileserver\共有\tebunko_ws" }
        refreshLegacyIndexMessage

        & $fake.Jobs[0].OnDone @() "届きません"

        $script:legacyIndexMessage | Should -Be ""
        $fake.Statuses.Count | Should -Be 0
    }

    It "起動時に別の知らせを出しているときは、ステータスを上書きしない" {
        $script:workspace = [pscustomobject]@{ Dir = "\\fileserver\共有\tebunko_ws" }
        $script:workspaceBlock = "既定のワークスペースが使えません"
        refreshLegacyIndexMessage

        & $fake.Jobs[0].OnDone @("前の版のインデックスがあります") ""

        $script:legacyIndexMessage | Should -Be "前の版のインデックスがあります"
        $fake.Statuses.Count | Should -Be 0
    }

    It "待っている間に調べ直したら、前の依頼の結果は捨てる" {
        $script:workspace = [pscustomobject]@{ Dir = "\\fileserver\共有\tebunko_ws" }
        refreshLegacyIndexMessage
        $script:workspace = [pscustomobject]@{ Dir = "\\fileserver\共有\ほか" }
        refreshLegacyIndexMessage

        & $fake.Jobs[0].OnDone @("古い知らせ") ""
        $fake.Statuses.Count | Should -Be 0
        & $fake.Jobs[1].OnDone @("新しい知らせ") ""

        $script:legacyIndexMessage | Should -Be "新しい知らせ"
        $fake.Statuses | Should -Be @("新しい知らせ")
    }
}
