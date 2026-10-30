# インデックス作成の司令の進み具合・確認待ち（tebunko\indexer\indexing_reporter.ps1 の IndexingReporter）のテスト。
# waitForIndexingApproval（関数）だったテストは、期待値を変えずにこのクラスのテストへ移した。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexing_reporter.ps1"
}

Describe "IndexingReporter.Progress" -Tag Unit {
    It "受け渡しの口に進み具合を書く" {
        $channel = newIndexerChannel
        $reporter = [IndexingReporter]::new($channel)
        $reporter.Progress(${indexingPhaseIngest}, 12, 34, 5, "営業\見積.xlsx")
        $progress = readIndexingProgress $channel
        $progress.Phase | Should -Be ${indexingPhaseIngest}
        $progress.Processed | Should -Be 12
        $progress.Remaining | Should -Be 34
        $progress.Failed | Should -Be 5
        $progress.Detail | Should -Be "営業\見積.xlsx"
    }
}

Describe "IndexingReporter.WaitForApproval" -Tag Io {
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
        $reporter = [IndexingReporter]::new($channel)
        $job = startAnswer $channel @{ RetryFailed = $false }
        $answer = $reporter.WaitForApproval(${indexingPhaseConfirm}, $plan, 2, 1, 60)
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
        $reporter = [IndexingReporter]::new($channel)
        $job = startAnswer $channel @{ RetryFailed = $true }
        ($reporter.WaitForApproval(${indexingPhaseConfirm}, $plan, 2, 1, 60)).RetryFailed | Should -Be $true
        [void](endAnswer $job)
    }

    It "中止を求められたら `$null を返す" {
        $channel = newIndexerChannel
        $reporter = [IndexingReporter]::new($channel)
        $job = startAnswer $channel $null $true
        $reporter.WaitForApproval(${indexingPhaseConfirm}, $plan, 2, 1, 60) | Should -BeNullOrEmpty
        [void](endAnswer $job)
    }

    It "前に残った返事は使わない（待ち始める前に消す）" {
        $channel = newIndexerChannel
        $reporter = [IndexingReporter]::new($channel)
        answerIndexingPlan $channel @{ RetryFailed = $true }
        $channel.Stop = $false
        $reporter.WaitForApproval(${indexingPhaseConfirm}, $plan, 2, 1, 0) | Should -BeNullOrEmpty
    }

    It "制限時間を過ぎても返事が無ければ `$null を返し、取り込み予定を外す" {
        $channel = newIndexerChannel
        $reporter = [IndexingReporter]::new($channel)
        $reporter.WaitForApproval(${indexingPhaseConfirm}, $plan, 2, 1, 0) | Should -BeNullOrEmpty
        $channel.Plan | Should -BeNullOrEmpty
    }
}
