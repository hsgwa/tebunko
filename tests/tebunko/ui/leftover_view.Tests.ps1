# 起動時の「前回残った Office の確認」の文言と判断（tebunko\ui\leftover_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\leftover_view.ps1"

    # getOfficeProcesses の要素に見立てた値。既定は「止めてよい残り物」
    # spec: @{ Id; Name（既定 EXCEL）; Owned（既定 $true）; Owner（既定 Gone）; Window（既定 $false）}（-TestCases の段階では関数が使えないため、値で渡す）
    function newOfficeProcess {
        param ([hashtable]$spec)
        $name = $(if ($spec.Name) { $spec.Name } else { "EXCEL" })
        return [pscustomobject]@{
            Id = $spec.Id; ProcessName = $name; AppName = ${officeProcessNames}[$name]
            StartTime = [datetime]"2030-01-01T09:00:00"
            HasWindow = [bool]$spec.Window
            Owned = $(if ($spec.ContainsKey("Owned")) { $spec.Owned } else { $true })
            Owner = $(if ($spec.ContainsKey("Owner")) { $spec.Owner } else { "Gone" })
        }
    }
}

Describe "getLeftoverPrompt" -Tag Unit {
    It "<name>は `$null（確認しない）" -TestCases @(
        @{ name = "プロセスが無い"; list = @() }
        @{ name = "記録が無い（利用者・ほかのアプリの Office）"; list = @(@{ Id = 1; Name = "EXCEL"; Owned = $false; Owner = "" }) }
        @{ name = "持ち主が今の画面"; list = @(@{ Id = 1; Name = "EXCEL"; Owned = $true; Owner = "Self" }) }
        @{ name = "持ち主がほかの画面・インデックス作成（生きている）"; list = @(@{ Id = 1; Name = "EXCEL"; Owned = $true; Owner = "Other" }) }
        @{ name = "窓がある（利用者が見ている可能性）"; list = @(@{ Id = 1; Name = "EXCEL"; Owned = $true; Owner = "Gone"; Window = $true }) }
    ) {
        getLeftoverPrompt @($list | ForEach-Object { newOfficeProcess $_ }) | Should -Be $null
    }

    It "残り物だけを対象にし、件数・内訳（Excel・Word・PowerPoint の順、0 件は出さない）を数える" {
        $list = @(
            (newOfficeProcess @{ Id = 1; Name = "WINWORD" })
            (newOfficeProcess @{ Id = 2; Name = "EXCEL" })
            (newOfficeProcess @{ Id = 3; Name = "EXCEL" })
            (newOfficeProcess @{ Id = 4; Name = "EXCEL"; Owned = $true; Owner = "Other" })
            (newOfficeProcess @{ Id = 5; Name = "POWERPNT"; Owned = $false; Owner = "" })
        )
        $prompt = getLeftoverPrompt $list
        @($prompt.Targets).Id | Should -Be @(1, 2, 3)
        $prompt.Count | Should -Be 3
        $prompt.Summary | Should -Be "Excel 2 件・Word 1 件"
        $prompt.Heading | Should -Be "前回の更新で起動した Office が 3 件、残ったまま動いています。終了しますか？"
    }

    It "文言・選択肢が決まっている" {
        $prompt = getLeftoverPrompt @((newOfficeProcess @{ Id = 1; Name = "POWERPNT" }))
        $prompt.Title | Should -Be "Office の終了"
        $prompt.Summary | Should -Be "PowerPoint 1 件"
        $prompt.Facts.Count | Should -Be 2
        @($prompt.Facts | ForEach-Object { $_.Kind }) | Should -Be @("Kept", "Kept")
        $prompt.Facts[0].Title | Should -Be "編集中のファイルは閉じません"
        $prompt.Facts[0].Detail | Should -Be "終了するのは、tebunko がバックグラウンドで起動した PowerPoint 1 件だけです"
        $prompt.Facts[1].Title | Should -Be "保存していない作業が失われることはありません"
        $prompt.Facts[1].Detail | Should -Be "残っているのは、更新のために起動したものです"
        $prompt.Hint | Should -Be "［今回は終了しない］を選んだときは、次に起動したときにもう一度お聞きします。"
        $prompt.CancelText | Should -Be "今回は終了しない"
        $prompt.Choices.Count | Should -Be 1
        $prompt.Choices[0].Text | Should -Be "終了する"
        $prompt.Choices[0].Value | Should -Be "stop"
        $prompt.Choices[0].ContainsKey("Danger") | Should -Be $false
    }

    It "一覧の行は、アプリ名・PID・起動時刻（時刻は生の値。書式は画面側）" {
        $prompt = getLeftoverPrompt @((newOfficeProcess @{ Id = 7; Name = "WINWORD" }))
        $prompt.Details.Count | Should -Be 1
        $prompt.Details[0].AppName | Should -Be "Word"
        $prompt.Details[0].Id | Should -Be 7
        $prompt.Details[0].StartTime | Should -Be ([datetime]"2030-01-01T09:00:00")
    }

    It "起動時刻が読めなくても残り物として出す（時刻は `$null）" {
        $p = newOfficeProcess @{ Id = 7 }
        $p.StartTime = $null
        $prompt = getLeftoverPrompt @($p)
        $prompt.Count | Should -Be 1
        $prompt.Details[0].StartTime | Should -Be $null
    }
}

Describe "getLeftoverTargets" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "出した PID が今も残り物なら、そのまま止める対象にする"; shown = @(1, 2); fresh = @(@{ Id = 1 }, @{ Id = 2 }); expected = @(1, 2) }
        @{ name = "確認の後に増えた残り物は、出していないので止めない"; shown = @(1); fresh = @(@{ Id = 1 }, @{ Id = 2 }); expected = @(1) }
        @{ name = "もう無い PID は外す"; shown = @(1, 2); fresh = @(@{ Id = 2 }); expected = @(2) }
        @{ name = "窓を持つようになった（利用者が開いた）ものは外す"; shown = @(1, 2); fresh = @(@{ Id = 1; Name = "EXCEL"; Owned = $true; Owner = "Gone"; Window = $true }, @{ Id = 2 }); expected = @(2) }
        @{ name = "記録が外れた（PID が別のプロセスになった）ものは外す"; shown = @(1); fresh = @(@{ Id = 1; Name = "EXCEL"; Owned = $false; Owner = "" }); expected = @() }
        @{ name = "持ち主が別の生きているプロセスになったものは外す"; shown = @(1); fresh = @(@{ Id = 1; Name = "EXCEL"; Owned = $true; Owner = "Other" }); expected = @() }
        @{ name = "読み直した一覧が空なら何も止めない"; shown = @(1); fresh = @(); expected = @() }
    ) {
        $current = @($fresh | ForEach-Object { newOfficeProcess $_ })
        @(getLeftoverTargets $shown $current | ForEach-Object { $_.Id }) | Should -Be $expected
    }
}

Describe "getLeftoverResultText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "全部止めた"; results = @(@{ Id = 1; Status = "Stopped" }, @{ Id = 2; Status = "Stopped" }, @{ Id = 3; Status = "Stopped" }); expected = "Office を 3 件終了しました。" }
        @{ name = "一部が別のプロセスに変わっていた"; results = @(@{ Id = 1; Status = "Stopped" }, @{ Id = 2; Status = "Stopped" }, @{ Id = 9316; Status = "Changed" }); expected = "Office を 2 件終了しました。PID 9316 は確認の後に別のプロセスに変わったため、終了しませんでした。" }
        @{ name = "変わった PID が複数なら、1 つの文に「・」でつなぐ"; results = @(@{ Id = 1; Status = "Stopped" }, @{ Id = 9316; Status = "Changed" }, @{ Id = 9320; Status = "Changed" }); expected = "Office を 1 件終了しました。PID 9316・9320 は確認の後に別のプロセスに変わったため、終了しませんでした。" }
        @{ name = "1 件も止めておらず、変わったものがあれば、件数の文は出さない"; results = @(@{ Id = 9316; Status = "Changed" }); expected = "PID 9316 は確認の後に別のプロセスに変わったため、終了しませんでした。" }
        @{ name = "止める前に無くなったものは、数にも理由にも入れない"; results = @(@{ Id = 1; Status = "Stopped" }, @{ Id = 2; Status = "Gone" }); expected = "Office を 1 件終了しました。" }
        @{ name = "止められなかったものは、PID と理由を添える"; results = @(@{ Id = 1; Status = "Stopped" }, @{ Id = 2; Status = "Failed"; Reason = "アクセスが拒否されました" }); expected = "Office を 1 件終了しました。PID 2 を終了できませんでした：アクセスが拒否されました" }
        @{ name = "変わった・失敗が両方ある"; results = @(@{ Id = 1; Status = "Stopped" }, @{ Id = 2; Status = "Changed" }, @{ Id = 3; Status = "Failed"; Reason = "拒否" }); expected = "Office を 1 件終了しました。PID 2 は確認の後に別のプロセスに変わったため、終了しませんでした。PID 3 を終了できませんでした：拒否" }
        @{ name = "全部もう無かった"; results = @(@{ Id = 1; Status = "Gone" }); expected = $null }
        @{ name = "結果が空"; results = @(); expected = $null }
    ) {
        getLeftoverResultText $results | Should -Be $expected
    }
}

Describe "getLeftoverPromptTiming" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "残り物が無ければ見送る"; count = 0; closing = $false; indexing = $false; notices = $false; dialogs = 0; expected = "Skip" }
        @{ name = "ウィンドウを閉じている途中なら見送る"; count = 2; closing = $true; indexing = $false; notices = $false; dialogs = 0; expected = "Skip" }
        @{ name = "インデックス作成中なら見送る"; count = 2; closing = $false; indexing = $true; notices = $false; dialogs = 0; expected = "Skip" }
        @{ name = "起動時のお知らせが閉じていなければ、閉じたあとにもう一度判断する"; count = 2; closing = $false; indexing = $false; notices = $true; dialogs = 1; expected = "Defer" }
        @{ name = "ほかのダイアログが開いていれば見送る"; count = 2; closing = $false; indexing = $false; notices = $false; dialogs = 1; expected = "Skip" }
        @{ name = "条件がそろえば出す"; count = 2; closing = $false; indexing = $false; notices = $false; dialogs = 0; expected = "Show" }
        @{ name = "残り物が無ければ、お知らせが開いていても出し直さない"; count = 0; closing = $false; indexing = $false; notices = $true; dialogs = 0; expected = "Skip" }
        @{ name = "閉じている途中は、お知らせが開いていても出し直さない"; count = 2; closing = $true; indexing = $false; notices = $true; dialogs = 0; expected = "Skip" }
        @{ name = "インデックス作成中は、お知らせが開いていても出し直さない"; count = 2; closing = $false; indexing = $true; notices = $true; dialogs = 0; expected = "Skip" }
    ) {
        getLeftoverPromptTiming $count $closing $indexing $notices $dialogs | Should -Be $expected
    }
}

Describe "getLeftoverFailureText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "確認の失敗（理由なし）"; phase = "Check"; err = ""; expected = "前回残った Office を確認できませんでした。" }
        @{ name = "確認の失敗（理由は 1 行目だけ）"; phase = "Check"; err = "アクセスが拒否されました`r`n発生場所 …"; expected = "前回残った Office を確認できませんでした（アクセスが拒否されました）。" }
        @{ name = "確認の失敗（LF だけの理由も 1 行目だけ）"; phase = "Check"; err = "アクセスが拒否されました`n発生場所 …"; expected = "前回残った Office を確認できませんでした（アクセスが拒否されました）。" }
        @{ name = "終了の失敗"; phase = "Stop"; err = "タイムアウト"; expected = "Office を終了できませんでした（タイムアウト）。残った Office は、次に tebunko を起動したときにもう一度確認できます。" }
        @{ name = "終了の失敗（理由なし）"; phase = "Stop"; err = $null; expected = "Office を終了できませんでした。残った Office は、次に tebunko を起動したときにもう一度確認できます。" }
    ) {
        getLeftoverFailureText $phase $err | Should -Be $expected
    }
}

Describe "getLeftoverTimeText・getLeftoverDetailToggleText" -Tag Unit {
    It "起動した時刻を「月/日 時:分」にし、読めなければ「不明」" -TestCases @(
        @{ value = [datetime]"2030-10-04T18:32:59"; expected = "10/04 18:32" }
        @{ value = [datetime]"2030-01-02T03:04:00"; expected = "01/02 03:04" }
        @{ value = $null; expected = "不明" }
        @{ value = "読めない"; expected = "不明" }
    ) {
        getLeftoverTimeText $value | Should -Be $expected
    }

    It "詳細の開け閉めの文字" -TestCases @(
        @{ open = $false; expected = "詳細を表示" }
        @{ open = $true; expected = "詳細を隠す" }
    ) {
        getLeftoverDetailToggleText $open | Should -Be $expected
    }

    It "確認の一覧の行に、アプリ・PID・時刻の文字が入る" {
        $prompt = getLeftoverPrompt @(
            (newOfficeProcess @{ Id = 12840; Name = "EXCEL" })
            (newOfficeProcess @{ Id = 9316; Name = "WINWORD" })
        )
        @($prompt.Details | ForEach-Object { $_.AppName }) | Should -Be @("Excel", "Word")
        @($prompt.Details | ForEach-Object { $_.Id }) | Should -Be @(12840, 9316)
        $prompt.Details[0].TimeText | Should -Be "01/01 09:00"
    }
}

Describe "偽の行（写真・画面のテスト用）" -Tag Unit {
    It "ファイルが指定されていなければ本物（Real）、指定されていれば偽（Fake）" -TestCases @(
        @{ file = $null; expected = "Real" }
        @{ file = ""; expected = "Real" }
        @{ file = "C:\temp\rows.json"; expected = "Fake" }
    ) {
        getLeftoverSourceMode $file | Should -Be $expected
    }

    It "偽の行は、残り物として確認に出る（記録あり・持ち主なし・窓なし）" {
        $rows = @(
            [pscustomobject]@{ Id = 12840; ProcessName = "EXCEL"; StartTime = "2030-10-04T18:32:00" }
            [pscustomobject]@{ Id = 9316; ProcessName = "WINWORD"; StartTime = "2030-10-04T18:41:00"; StopStatus = "Changed" }
        )
        $processes = @(convertLeftoverFakeRows $rows)
        $processes.Count | Should -Be 2
        @($processes | ForEach-Object { $_.AppName }) | Should -Be @("Excel", "Word")
        @(selectLeftoverProcesses $processes).Count | Should -Be 2
        (getLeftoverPrompt $processes).Details[1].TimeText | Should -Be "10/04 18:41"
    }

    It "偽の行の終了の結果は、StopStatus どおりに作る（プロセスには触れない）" {
        $targets = @(convertLeftoverFakeRows @(
            [pscustomobject]@{ Id = 1; ProcessName = "EXCEL" }
            [pscustomobject]@{ Id = 2; ProcessName = "WINWORD"; StopStatus = "Changed" }
            [pscustomobject]@{ Id = 3; ProcessName = "EXCEL"; StopStatus = "Failed" }
        ))
        $results = @(getLeftoverFakeStopResults $targets)
        @($results | ForEach-Object { $_.Status }) | Should -Be @("Stopped", "Changed", "Failed")
        $results[0].Stopped | Should -BeTrue
        $results[1].Reason | Should -Be "PID 2 は確認の後に別のプロセスに変わったため、終了しませんでした。"
        getLeftoverResultText $results | Should -BeLike "Office を 1 件終了しました。PID 2 は*PID 3 を終了できませんでした：*"
    }
}
