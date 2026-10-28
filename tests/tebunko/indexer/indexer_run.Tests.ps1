# インデックス作成の司令（tebunko\indexer\indexer_run.ps1）のうち、確認の返事を待って、選びを変えた返事（Recount）なら数え直す所
# （applyRecountChoices・waitForIngestApproval）のテスト。司令全体を通す確かめは indexer.Tests.ps1 にある。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"
    . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"  # findDroppedIndexes（IngestPlanner）
    . "${scriptsDir}\tebunko\indexer\pending_publish.ps1"  # PendingPublish
    . "${scriptsDir}\tebunko\indexer\indexer_run.ps1"

    function newPrevious {
        $map = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        return , $map
    }

    function newCrawledPlanner {
        # 名前を割り当てて数えた IngestPlanner（司令が最初に作るものと同じ手順）
        param ($targetFolders, $status)
        $planner = [IngestPlanner]::new(@{
            Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed}
        })
        $planner.AssignNames($targetFolders, $status)
        $planner.Crawl($null, { }, { $false })
        return $planner
    }

    function newFakeReporter {
        # 画面の返事を順に返す偽物（呼ばれた回数と、渡された Plan を控える）
        param ([object[]]$answers)
        $reporter = [pscustomobject]@{ Answers = (New-Object System.Collections.Queue); Calls = 0; PlanSizes = (New-Object System.Collections.Generic.List[int]); OnCall = $null }
        foreach ($answer in $answers) { $reporter.Answers.Enqueue($answer) }
        Add-Member -InputObject $reporter -MemberType ScriptMethod -Name WaitForApproval -Value {
            param ($phase, $plan, $targetCount, $failedCount, $timeoutMinutes)
            $this.Calls++
            $this.PlanSizes.Add(@($plan | ForEach-Object { $_.ファイル数 } | Measure-Object -Sum).Sum)
            if ($null -ne $this.OnCall) { & $this.OnCall }
            return $this.Answers.Dequeue()
        }
        return $reporter
    }
}

Describe "applyRecountChoices" -Tag Unit {
    BeforeAll {
        $assigned = @([pscustomobject]@{ Name = "売上"; Path = "C:\data\売上" }, [pscustomobject]@{ Name = "経理"; Path = "C:\data\経理" })
        function newTargets {
            return @(
                [pscustomobject]@{ Name = ""; Path = "C:\data\売上"; Enabled = $true; Subfolders = $true },
                [pscustomobject]@{ Name = ""; Path = "C:\data\経理"; Enabled = $false; Subfolders = $true }
            )
        }
    }

    It "<label>" -TestCases @(
        @{ label = "名前に当たる項目の Subfolders だけを替える（ほかの値は引き継ぐ）"; changes = @(@{ Name = "売上"; Subfolders = $false }); expected = @($false, $true) }
        @{ label = "名前の無い Change は無視する"; changes = @(@{ Name = "存在しない"; Subfolders = $false }); expected = @($true, $true) }
        @{ label = "Change が無ければそのまま写す"; changes = @(); expected = @($true, $true) }
        @{ label = "2 つとも替える"; changes = @(@{ Name = "売上"; Subfolders = $false }, @{ Name = "経理"; Subfolders = $false }); expected = @($false, $false) }
    ) {
        param ($label, $changes, $expected)
        $result = @(applyRecountChoices (newTargets) $assigned $changes)
        @($result | ForEach-Object { $_.Subfolders }) | Should -Be $expected
        @($result | ForEach-Object { $_.Enabled }) | Should -Be @($true, $false)
    }

    It "名前がまだ無い新しいフォルダ（Name が空）は、1 回目に割り当てた名前を写しに入れる（数え直しのたびに割り当て直さず、同じ名前にする）" {
        $result = @(applyRecountChoices (newTargets) $assigned @())
        @($result | ForEach-Object { $_.Name }) | Should -Be @("売上", "経理")
    }

    It "元のクロール対象フォルダの一覧は書き換えない（写しに当てる）" {
        $targets = newTargets
        [void](applyRecountChoices $targets $assigned @(@{ Name = "売上"; Subfolders = $false }))
        $targets[0].Subfolders | Should -BeTrue
    }
}

Describe "waitForIngestApproval" -Tag Io {
    BeforeAll {
        $root = Join-Path $TestDrive "src"
        [System.IO.Directory]::CreateDirectory("$root\下") | Out-Null
        [System.IO.File]::WriteAllText("$root\直下.docx", "dummy")
        [System.IO.File]::WriteAllText("$root\下\奥.docx", "dummy")
        ${indexDir} = Join-Path $TestDrive "index_run"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $targetFolders = @([pscustomobject]@{ Name = ""; Path = $root; Enabled = $true; Subfolders = $true })
        function newStatus { return @{ Folders = @(); Rows = (newPrevious); LegacyFolders = @() } }
    }

    It "ふつうの返事（取り込む）は、数え直さずにそのまま返す" {
        $status = newStatus
        $planner = newCrawledPlanner $targetFolders $status
        $reporter = newFakeReporter @(@{ RetryFailed = $false })
        $channel = newIndexerChannel

        $result = waitForIngestApproval $planner $reporter $channel $targetFolders $status $null (New-Object PendingPublish) 60

        $result.Answer.RetryFailed | Should -BeFalse
        [object]::ReferenceEquals($result.Planner, $planner) | Should -BeTrue
        $reporter.Calls | Should -Be 1
    }

    It "取りやめの返事（`$null）は、そのまま `$null で返す" {
        $status = newStatus
        $planner = newCrawledPlanner $targetFolders $status
        $reporter = newFakeReporter @($null)

        $result = waitForIngestApproval $planner $reporter (newIndexerChannel) $targetFolders $status $null (New-Object PendingPublish) 60

        $result.Answer | Should -BeNullOrEmpty
        $reporter.Calls | Should -Be 1
    }

    It "選びを変えた返事（Recount）なら、その選びで数え直し、新しい Plan（直下だけの件数）でもう一度待つ" {
        $status = newStatus
        $planner = newCrawledPlanner $targetFolders $status
        $planner.Plan[0].ファイル数 | Should -Be 2
        $reporter = newFakeReporter @(@{ Recount = $true; Changes = @(@{ Name = "売上"; Subfolders = $false }) }, @{ RetryFailed = $false })
        # 数えるときに付いた名前（売上）に合わせて、Change の名前を引く
        $reporter.Answers.Clear()
        $name = [string]$planner.Folders[0].Name
        $reporter.Answers.Enqueue(@{ Recount = $true; Changes = @(@{ Name = $name; Subfolders = $false }) })
        $reporter.Answers.Enqueue(@{ RetryFailed = $false })

        $result = waitForIngestApproval $planner $reporter (newIndexerChannel) $targetFolders $status $null (New-Object PendingPublish) 60

        $reporter.Calls | Should -Be 2
        @($reporter.PlanSizes) | Should -Be @(2, 1)
        $result.Answer.RetryFailed | Should -BeFalse
        [object]::ReferenceEquals($result.Planner, $planner) | Should -BeFalse
        $result.Planner.Targets.Count | Should -Be 1
        # 元のクロール対象フォルダの一覧は書き換えない（設定に書くのは画面）
        $targetFolders[0].Subfolders | Should -BeTrue
    }

    It "数え直しの最中に中止が立つと、数え終えたところで待たずに取りやめの返事（`$null）で返す" {
        $status = newStatus
        $planner = newCrawledPlanner $targetFolders $status
        $name = [string]$planner.Folders[0].Name
        $channel = newIndexerChannel
        $reporter = newFakeReporter @(@{ Recount = $true; Changes = @(@{ Name = $name; Subfolders = $false }) })
        # 返事を返す直前に、画面が［キャンセル］を押した（数え直しの最中）ことにする
        $reporter.OnCall = { $channel.Stop = $true }.GetNewClosure()

        $result = waitForIngestApproval $planner $reporter $channel $targetFolders $status $null (New-Object PendingPublish) 60

        $result.Answer | Should -BeNullOrEmpty
        $reporter.Calls | Should -Be 1
    }

    It "数え直しの前に無くなっていたファイルの削除は、数え直しても 1 件だけ記録する（二重に入らない）" {
        $status = newStatus
        # 名前は、クロール対象フォルダの名前（フォルダ名の src）から割り当てられる
        $lostRow = newStatusRow "src\無い.docx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"
        $status.Rows["src\無い.docx"] = $lostRow
        $planner = newCrawledPlanner $targetFolders $status
        $name = [string]$planner.Folders[0].Name
        $pending = New-Object PendingPublish
        foreach ($removed in $planner.RemovedPaths) { $pending.Add($removed, $true) }
        $reporter = newFakeReporter @(@{ Recount = $true; Changes = @(@{ Name = $name; Subfolders = $false }) }, @{ RetryFailed = $false })

        [void](waitForIngestApproval $planner $reporter (newIndexerChannel) $targetFolders $status $null $pending 60)

        $names = @($pending.Unpublished.Values | ForEach-Object { $_ } | ForEach-Object { $_ })
        $names.Count | Should -Be 1
        $names[0] | Should -Be "無い.docx"
    }

    It "外して数え直し→付け直して数え直し→開始、と進めても、前の数えにしか無い削除は残らない（サブフォルダのファイルが本文インデックスから外れて検索に出なくなるのを防ぐ）" {
        # 取り込み済みのサブフォルダのファイル（前回の行と更新日時・サイズが同じ＝済）
        $info = Get-Item -LiteralPath "$root\下\奥.docx"
        $updated = formatFileTime $info.LastWriteTime
        $status = newStatus
        $status.Rows["src\下\奥.docx"] = (newStatusRow "src\下\奥.docx" $updated ([string]$info.Length) ${stateDone} 1 $updated)
        $planner = newCrawledPlanner $targetFolders $status
        $name = [string]$planner.Folders[0].Name
        $pending = New-Object PendingPublish
        foreach ($removed in $planner.RemovedPaths) { $pending.Add($removed, $true) }
        @($pending.Unpublished.Values).Count | Should -Be 0 -Because "最初の数え（サブフォルダも含める）では、奥.docx は無くなっていない"
        $reporter = newFakeReporter @(
            @{ Recount = $true; Changes = @(@{ Name = $name; Subfolders = $false }) },
            @{ Recount = $true; Changes = @(@{ Name = $name; Subfolders = $true }) },
            @{ RetryFailed = $false })

        $result = waitForIngestApproval $planner $reporter (newIndexerChannel) $targetFolders $status $null $pending 60

        $reporter.Calls | Should -Be 3
        $result.Planner.RemovedPaths.Count | Should -Be 0
        # 1 回目の数え直し（直下だけ）で「無くなった」ことにした奥.docx は、付け直したあとの数えでは済のまま。外す記録に残してはならない
        @($pending.Unpublished.Values | ForEach-Object { $_ } | ForEach-Object { $_ }).Count | Should -Be 0
        @($pending.Folders.Values | ForEach-Object { $_ } | ForEach-Object { $_ }).Count | Should -Be 0
    }
}
