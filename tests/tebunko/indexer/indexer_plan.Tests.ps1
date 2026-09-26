# クロール（tebunko\indexer\indexer_plan.ps1 の createTargetList）のテスト。
BeforeDiscovery {
    # -TestCases の表が使う一覧の状態（$stateDone など）
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"

    function newPrevious {
        param ($rows = @())
        $map = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $rows) { $map[$row.相対パス] = $row }
        return , $map
    }

    function newCounts {
        param ([hashtable]$pairs = @{})
        $counts = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($key in $pairs.Keys) { $counts[$key] = $pairs[$key] }
        return , $counts
    }
}

Describe "createTargetList" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "src"
        [System.IO.Directory]::CreateDirectory($source) | Out-Null
        $file = Join-Path $source "a.xlsx"
        [System.IO.File]::WriteAllText($file, "dummy")
        $info = Get-Item -LiteralPath $file
        $updated = formatFileTime $info.LastWriteTime
        $size = [string]$info.Length
        $folder = @{ Path = $source; Name = "売上" }
        # removeBookDir が実際のインデックスを見ないよう、テスト用のフォルダに向ける
        ${indexDir} = Join-Path $TestDrive "index"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
    }

    # 一覧の行（state が $null なら一覧に無い。modified は元のファイルが更新されたか、version は抽出版）と
    # TSV の数え上げ（tsv が $null なら数えない・0 なら数え上げに無い）→ 取り込み対象・失敗の数と、取り込み予定のどの件数に数えるか
    It "<name>" -TestCases @(
        @{ name = "一覧に無いファイルは取り込み対象になる（新規）"; state = $null; modified = $false; version = "2"; tsv = $null; targets = 1; failed = 0; field = "新規" }
        @{ name = "取り込み済みで更新が無ければ取り込まない"; state = $stateDone; modified = $false; version = "2"; tsv = 1; targets = 0; failed = 0; field = "" }
        @{ name = "前の抽出版で取り込んだファイルは、更新が無くても取り込み直す（更新ありに数える）"; state = $stateDone; modified = $false; version = ""; tsv = 1; targets = 1; failed = 0; field = "更新あり" }
        @{ name = "取り込み済みでも TSV が無ければ取り込み直す（インデックスなし）"; state = $stateDone; modified = $false; version = "2"; tsv = 0; targets = 1; failed = 0; field = "インデックスなし" }
        @{ name = "更新されていれば取り込み対象になる（更新あり）"; state = $stateDone; modified = $true; version = "2"; tsv = 1; targets = 1; failed = 0; field = "更新あり" }
        @{ name = "前回失敗して更新が無ければ、取り込み対象ではなく失敗として返す"; state = $stateFailed; modified = $false; version = "2"; tsv = 1; targets = 0; failed = 1; field = "前回失敗" }
        @{ name = "前回「未取り込み」で終わっていれば取り込み対象になる（前回未完了）"; state = $stateNew; modified = $false; version = "2"; tsv = 1; targets = 1; failed = 0; field = "前回未完了" }
    ) {
        param ($name, $state, $modified, $version, $tsv, $targets, $failed, $field)
        $rows = @()
        if ($null -ne $state) {
            $rowUpdated = if ($modified) { "2000/01/01 00:00:00" } else { $updated }
            $rows = @((newStatusRow "売上\a.xlsx" $rowUpdated $size $state 1 $updated "" $version))
        }
        $counts = $null
        if ($tsv -eq 0) {
            $counts = newCounts
        } elseif ($tsv) {
            $counts = newCounts @{ "売上\a.xlsx" = $tsv }
        }
        $result = createTargetList $folder (newPrevious $rows) $counts
        $result.Rows.Count | Should -Be 1
        $result.Targets.Count | Should -Be $targets
        $result.Failed.Count | Should -Be $failed
        if ($field) {
            $result.Plan.$field | Should -Be 1
        }
        $result.Plan.ファイル数 | Should -Be 1
        $result.Plan.取り込み対象 | Should -Be $targets
    }
}

Describe "createTargetList（サブフォルダ・無くなったファイル）" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "src2"
        [System.IO.Directory]::CreateDirectory("$source\2024") | Out-Null
        $file = Join-Path $source "2024\b.docx"
        [System.IO.File]::WriteAllText($file, "dummy")
        (Get-Item -LiteralPath $file).LastWriteTime = [datetime]"2024/04/01 09:00:00"
        $folder = @{ Path = $source; Name = "経理" }
        ${indexDir} = Join-Path $TestDrive "index2"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $bookDir = Join-Path ${indexDir} "経理\2024\b.docx"

        function newIndexTsv([string]$name, [datetime]$time) {
            [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
            $path = Join-Path $bookDir $name
            [System.IO.File]::WriteAllText($path, "1`t見積")
            (Get-Item -LiteralPath $path).LastWriteTime = $time
        }
    }

    AfterEach {
        if (Test-Path -LiteralPath ${indexDir}) { Remove-Item -LiteralPath ${indexDir} -Recurse -Force }
    }

    It "サブフォルダのファイルは、インデックス名とフォルダからの相対パスでつなぐ" {
        $result = createTargetList $folder (newPrevious) $null
        $result.Rows[0].相対パス | Should -Be "経理\2024\b.docx"
    }

    It "一覧に無いファイルは、元のファイルより新しいインデックスのフォルダが残っていても取り込み対象にする" {
        newIndexTsv "1ページ.tsv" ([datetime]"2024/04/02 09:00:00")
        $result = createTargetList $folder (newPrevious) $null
        $result.Targets.Count | Should -Be 1
        $result.Plan.新規 | Should -Be 1
    }

    It "一覧にあって元のファイルが無くなったものは、インデックスを消して一覧から除く" {
        $gone = Join-Path ${indexDir} "経理\消えた.xlsx"
        [System.IO.Directory]::CreateDirectory($gone) | Out-Null
        [System.IO.File]::WriteAllText("$gone\Sheet1.tsv", "1`t見積")
        $previous = newPrevious @((newStatusRow "経理\消えた.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理\消えた.xlsx" }).Count | Should -Be 0
        Test-Path -LiteralPath $gone | Should -Be $false
    }

    It "ほかのインデックスの行には触らない（名前が前方一致するインデックスも別のものとして扱う）" {
        $other = Join-Path ${indexDir} "経理2\c.xlsx"
        [System.IO.Directory]::CreateDirectory($other) | Out-Null
        $previous = newPrevious @((newStatusRow "経理2\c.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理2\c.xlsx" }).Count | Should -Be 0
        Test-Path -LiteralPath $other | Should -Be $true
    }

    It "アクセスできないフォルダがあったときは、見つからなかったファイルの行とインデックスを残す" {
        Mock findOfficeFiles { @{ Root = $source; Files = @(Get-Item -LiteralPath $file); HasError = $true } }
        $gone = Join-Path ${indexDir} "経理\読めない\d.xlsx"
        [System.IO.Directory]::CreateDirectory($gone) | Out-Null
        $previous = newPrevious @((newStatusRow "経理\読めない\d.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理\読めない\d.xlsx" }).Count | Should -Be 1
        Test-Path -LiteralPath $gone | Should -Be $true
        # \\?\ の付かないパスで列挙されたファイルも、フォルダからの相対パスにする
        $result.Rows[0].相対パス | Should -Be "経理\2024\b.docx"
    }
}

Describe "findOfficeFiles" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "scan"
        [System.IO.Directory]::CreateDirectory("$source\下\さらに下") | Out-Null
        foreach ($name in @("a.xlsx", "下\b.DOCX", "下\さらに下\c.pptm", "d.txt", "e.pdf", ('~$' + "a.xlsx"), "f.xls", "g.ppt")) {
            [System.IO.File]::WriteAllText((Join-Path $source $name), "dummy")
        }
    }

    It "サブフォルダも含めて Office の拡張子のファイルだけを返す（大文字の拡張子も含め、~$ で始まるロックファイルは除く）" {
        $scan = findOfficeFiles $source
        @($scan.Files | ForEach-Object { $_.Name } | Sort-Object) -join "," | Should -Be "a.xlsx,b.DOCX,c.pptm,f.xls,g.ppt"
        $scan.HasError | Should -Be $false
    }

    It "末尾に \ を付けたフォルダでも同じフォルダを列挙する" {
        $scan = findOfficeFiles "$source\"
        $scan.Root.TrimEnd("\") | Should -Be $source
        $scan.Files.Count | Should -Be 5
    }
}

Describe "waitForIndexingApproval" -Tag Io {
    BeforeAll {
        $plan = @((newIngestPlanRow "営業" "C:\共有\営業部" ${planKindIngest} 3 2 1 1 0 0 1))

        function startAnswer($channel, $answer, [bool]$cancel = $false) {
            # 画面の代わりに、別のスレッドで待たれている間の中身を控えてから返事をする
            $ps = [powershell]::Create()
            [void]$ps.AddScript({
                param ($channel, $answer, $cancel)
                while ($null -eq $channel.Plan) { Start-Sleep -Milliseconds 20 }
                $seen = @{ Plan = @($channel.Plan); Progress = $channel.Progress }
                if ($cancel) {
                    $channel.Stop = $true
                } else {
                    $channel.Answer = $answer
                }
                [void]$channel.Answered.Set()
                $seen
            }).AddArgument($channel).AddArgument($answer).AddArgument($cancel)
            return @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
        }

        function endAnswer($job) {
            try { return $job.PowerShell.EndInvoke($job.Handle)[0] } finally { $job.PowerShell.Dispose() }
        }
    }

    It "取り込み予定と確認待ちの進み具合を入れてから待ち、画面が開始を選んだら返事を返す" {
        $channel = newIndexerChannel
        $job = startAnswer $channel @{ RetryFailed = $false }
        $answer = waitForIndexingApproval $channel $plan 2 1
        $seen = endAnswer $job
        $answer.RetryFailed | Should -Be $false
        $seen.Plan.Count | Should -Be 1
        $seen.Plan[0].インデックス名 | Should -Be "営業"
        $seen.Progress.Phase | Should -Be ${indexingPhaseConfirm}
        $seen.Progress.Remaining | Should -Be 2
        $seen.Progress.Failed | Should -Be 1
        # 返事を受けたら、取り込み予定は外す
        $channel.Plan | Should -BeNullOrEmpty
    }

    It "画面が失敗分の再取り込みを選んだら RetryFailed を返す" {
        $channel = newIndexerChannel
        $job = startAnswer $channel @{ RetryFailed = $true }
        (waitForIndexingApproval $channel $plan 2 1).RetryFailed | Should -Be $true
        [void](endAnswer $job)
    }

    It "中止を求められたら `$null を返す" {
        $channel = newIndexerChannel
        $job = startAnswer $channel $null $true
        waitForIndexingApproval $channel $plan 2 1 | Should -BeNullOrEmpty
        [void](endAnswer $job)
    }

    It "前に残った返事は使わない（待ち始める前に消す）" {
        $channel = newIndexerChannel
        answerIndexingPlan $channel @{ RetryFailed = $true }
        $channel.Stop = $false
        waitForIndexingApproval $channel $plan 2 1 0 | Should -BeNullOrEmpty
    }

    It "制限時間を過ぎても返事が無ければ `$null を返し、取り込み予定を外す" {
        $channel = newIndexerChannel
        waitForIndexingApproval $channel $plan 2 1 0 | Should -BeNullOrEmpty
        $channel.Plan | Should -BeNullOrEmpty
    }
}
