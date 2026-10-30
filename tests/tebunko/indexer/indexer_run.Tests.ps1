# インデックス作成の司令（tebunko\indexer\indexer_run.ps1）のうち、数えたあと・取り込み終えたあとの処理の関数
# （applyRemovedPaths・recordIngestedSettings）のテスト。司令全体を通す確かめは indexer.Tests.ps1 にある。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"
    . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"  # findDroppedIndexes（IngestPlanner）
    . "${scriptsDir}\tebunko\indexer\pending_publish.ps1"  # PendingPublish
    . "${scriptsDir}\tebunko\indexer\indexer_run.ps1"
}

Describe "applyRemovedPaths" -Tag Io {
    It "無くなったファイルのインデックスのフォルダを消し、外す予定を `$pending に足す" {
        ${indexDir} = Join-Path $TestDrive "index_apply"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $bookDir = getBookDir "src\無い.docx"
        [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
        $pending = New-Object PendingPublish

        applyRemovedPaths @("src\無い.docx") $pending

        Test-Path -LiteralPath $bookDir | Should -BeFalse
        $names = @($pending.Unpublished.Values | ForEach-Object { $_ } | ForEach-Object { $_ })
        $names | Should -Be @("無い.docx")
    }

    It "空でも何もしない" {
        $pending = New-Object PendingPublish
        { applyRemovedPaths @() $pending } | Should -Not -Throw
        $pending.Unpublished.Count | Should -Be 0
    }
}

Describe "recordIngestedSettings" -Tag Unit {
    BeforeAll {
        function newRecordPlanner {
            # 名前を割り当てた Folders と、取り込む行（区分 = 取り込む）を持つ IngestPlanner（Crawl は通さない）
            param ($folders, $plan)
            $planner = [IngestPlanner]::new(@{
                Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed}
            })
            $planner.Folders = @($folders)
            foreach ($row in $plan) { $planner.Plan.Add($row) }
            return $planner
        }

        function newFakeCatalog {
            $catalog = [pscustomobject]@{ Written = $null; Fail = $false }
            Add-Member -InputObject $catalog -MemberType ScriptMethod -Name StoreIngestedSubfolders -Value {
                param ($values)
                if ($this.Fail) { throw "書けません" }
                $this.Written = $values
            }
            return $catalog
        }
    }

    It "取り込んだフォルダ（区分が取り込む）が使った subfolders を、パスを鍵にして書く。チェックなし・フォルダなしの行は書かない" {
        $planner = newRecordPlanner @(
            [pscustomobject]@{ Name = "売上"; Path = "C:\data\売上"; Enabled = $true; Subfolders = $false },
            [pscustomobject]@{ Name = "経理"; Path = "C:\data\経理"; Enabled = $true; Subfolders = $true },
            [pscustomobject]@{ Name = "総務"; Path = "C:\data\総務"; Enabled = $false; Subfolders = $false },
            [pscustomobject]@{ Name = "企画"; Path = "C:\data\企画"; Enabled = $true; Subfolders = $false }
        ) @(
            (newIngestPlanRow "売上" "C:\data\売上" ${planKindIngest}),
            (newIngestPlanRow "経理" "C:\data\経理" ${planKindIngest}),
            (newIngestPlanRow "総務" "C:\data\総務" ${planKindUnchecked}),
            (newIngestPlanRow "企画" "C:\data\企画" ${planKindMissing})
        )
        $catalog = newFakeCatalog

        recordIngestedSettings $planner $catalog

        $catalog.Written.Count | Should -Be 2
        $catalog.Written["C:\data\売上"] | Should -BeFalse
        $catalog.Written["C:\data\経理"] | Should -BeTrue
    }

    It "選んだものだけの回は、選ばなかったフォルダ（Plan に取り込む行が無い）を書かない" {
        $planner = newRecordPlanner @(
            [pscustomobject]@{ Name = "売上"; Path = "C:\data\売上"; Enabled = $true; Subfolders = $false },
            [pscustomobject]@{ Name = "経理"; Path = "C:\data\経理"; Enabled = $true; Subfolders = $false }
        ) @((newIngestPlanRow "売上" "C:\data\売上" ${planKindIngest}))
        $catalog = newFakeCatalog

        recordIngestedSettings $planner $catalog

        @($catalog.Written.Keys) | Should -Be @("C:\data\売上")
    }

    It "書く対象が無ければ、設定を書かない" {
        $planner = newRecordPlanner @([pscustomobject]@{ Name = "売上"; Path = "C:\data\売上"; Enabled = $false; Subfolders = $true }) @((newIngestPlanRow "売上" "C:\data\売上" ${planKindUnchecked}))
        $catalog = newFakeCatalog

        recordIngestedSettings $planner $catalog

        $catalog.Written | Should -BeNullOrEmpty
    }

    It "書けなくても例外にせず、ログに黄色で残す" {
        Mock writeIndexerLog {}
        $planner = newRecordPlanner @([pscustomobject]@{ Name = "売上"; Path = "C:\data\売上"; Enabled = $true; Subfolders = $false }) @((newIngestPlanRow "売上" "C:\data\売上" ${planKindIngest}))
        $catalog = newFakeCatalog
        $catalog.Fail = $true

        { recordIngestedSettings $planner $catalog } | Should -Not -Throw

        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*前回取り込んだときのサブフォルダの設定を保存できませんでした*" -and $color -eq "Yellow" }
    }
}
