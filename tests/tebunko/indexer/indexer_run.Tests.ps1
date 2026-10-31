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
            Add-Member -InputObject $catalog -MemberType ScriptMethod -Name StoreIngestedTextExcludes -Value { param ($values) }
            Add-Member -InputObject $catalog -MemberType ScriptMethod -Name StoreTextExtCounts -Value { param ($name, $counts, $time) }
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

Describe "recordIngestedTextExtensions" -Tag Unit {
    BeforeAll {
        function newExtPlanner {
            param ($folders, $plan, $excludes, $counts)
            $planner = [IngestPlanner]::new(@{
                Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed}
            })
            $planner.Folders = @($folders)
            foreach ($row in $plan) { $planner.Plan.Add($row) }
            $planner.TextExcludes = $excludes
            foreach ($name in @($counts.Keys)) { $planner.TextCountsByIndex[$name] = $counts[$name] }
            return $planner
        }

        function newExtCatalog {
            $catalog = [pscustomobject]@{ WrittenExcludes = $null; WrittenCounts = @(); Fail = $false }
            Add-Member -InputObject $catalog -MemberType ScriptMethod -Name StoreIngestedTextExcludes -Value { param ($values) if ($this.Fail) { throw "書けません" }; $this.WrittenExcludes = $values }
            Add-Member -InputObject $catalog -MemberType ScriptMethod -Name StoreTextExtCounts -Value { param ($name, $counts, $time) $this.WrittenCounts += @(@{ Name = $name; Counts = $counts; Time = $time }) }
            return $catalog
        }
    }

    It "取り込んだインデックスの件数の控えを書き、名前の有効なフォルダをすべて回した名前だけ、使った取り込まない拡張子を書く" {
        $planner = newExtPlanner @(
            [pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $true },
            [pscustomobject]@{ Name = "営業"; Path = "C:\b"; Enabled = $true },
            [pscustomobject]@{ Name = "経理"; Path = "C:\c"; Enabled = $true }
        ) @(
            (newIngestPlanRow "営業" "C:\a" ${planKindIngest}),
            (newIngestPlanRow "経理" "C:\c" ${planKindIngest})
        ) @{ "営業" = @(".log"); "経理" = @() } @{ "営業" = @{ ".txt" = 3 }; "経理" = @{ ".md" = 1 } }
        $catalog = newExtCatalog

        recordIngestedTextExtensions $planner $catalog

        @($catalog.WrittenCounts | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @("営業", "経理")
        $catalog.WrittenCounts[0].Time | Should -Match '^\d{4}/\d{2}/\d{2} '
        # 営業は C:\b を回していないので書かない。経理は全部回したので（空の一覧として）書く
        @($catalog.WrittenExcludes.Keys) | Should -Be @("経理")
        @($catalog.WrittenExcludes["経理"]).Count | Should -Be 0
    }

    It "全部回した名前の取り込まない拡張子を、そのまま書く" {
        $planner = newExtPlanner @([pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $true }) @((newIngestPlanRow "営業" "C:\a" ${planKindIngest})) @{ "営業" = @(".log", ".md") } @{}
        $catalog = newExtCatalog

        recordIngestedTextExtensions $planner $catalog

        ($catalog.WrittenExcludes["営業"] -join ",") | Should -Be ".log,.md"
    }

    It "取り込んでいない行（チェックなし・フォルダなし）は書かない" {
        $planner = newExtPlanner @([pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $false }) @((newIngestPlanRow "営業" "C:\a" ${planKindUnchecked})) @{} @{}
        $catalog = newExtCatalog

        recordIngestedTextExtensions $planner $catalog

        $catalog.WrittenCounts.Count | Should -Be 0
        $catalog.WrittenExcludes | Should -BeNullOrEmpty
    }

    It "書けなくても例外にせず、ログに黄色で残す" {
        Mock writeIndexerLog {}
        $planner = newExtPlanner @([pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $true }) @((newIngestPlanRow "営業" "C:\a" ${planKindIngest})) @{ "営業" = @(".log") } @{}
        $catalog = newExtCatalog
        $catalog.Fail = $true

        { recordIngestedTextExtensions $planner $catalog } | Should -Not -Throw

        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*前回取り込んだときの拡張子の設定・拡張子ごとの件数を保存できませんでした*" -and $color -eq "Yellow" }
    }
}
