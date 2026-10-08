# 画面のスモークテスト: 起動時の「前回残った Office の確認」の、本物の道筋（偽の行を使わない）。
# 記録を読む → 確認 → 取り直して照らし合わせる → 止める → 記録が消える、までを通しで確かめる。
# 本物の Office には触らない。止める相手は、このテストが起動して PID を控えた偽のプロセス
# （%windir%\System32\PING.EXE を EXCEL.EXE の名前で写したもの。窓が無い）だけ。
# 記録の置き場と作業フォルダは $TestDrive の下（画面のコピーの work）で、実機の記録・設定には触らない。
# 記録の無い Office は、本物でも偽でも対象にならない（記録を書くのはこのテストだけ）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"

    function newFakeExcel {
        # 偽の Excel を起動して返す（Handle を取って、終わった後も起動時刻などが読めるようにする）
        param ([string]$Dir)
        [void][IO.Directory]::CreateDirectory($Dir)
        Copy-Item -LiteralPath (Join-Path $env:windir "System32\PING.EXE") -Destination "$Dir\EXCEL.EXE"
        $p = Start-Process "$Dir\EXCEL.EXE" -ArgumentList "127.0.0.1", "-n", "600" -WindowStyle Hidden -PassThru
        $null = $p.Handle
        return $p
    }
}

Describe "起動時の前回残った Office の確認（本物の道筋）" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        # 偽の行の差し込み口が残っていると本物の枝を通らないので、控えて消す（AfterAll で戻す）
        $script:leftoverEnvBefore = $env:TEBUNKO_GUI_LEFTOVER_FILE
        Remove-Item Env:\TEBUNKO_GUI_LEFTOVER_FILE -ErrorAction SilentlyContinue
        $script:tool = newGuiTool $TestDrive
        $script:recordDir = Join-Path $script:tool.Work "office_pids\$(getMachineKey)"
        [void][IO.Directory]::CreateDirectory($script:recordDir)

        # 記録のある偽の Excel（持ち主は、もう終わったプロセス）
        $script:recorded = newFakeExcel (Join-Path $TestDrive "fake_a")
        # 記録の無い偽の Excel（残り物にならない）
        $script:unrecorded = newFakeExcel (Join-Path $TestDrive "fake_b")
        $owner = Start-Process cmd.exe -ArgumentList "/c", "exit" -WindowStyle Hidden -PassThru
        $null = $owner.Handle
        $ownerTicks = getOfficeStartTicks $owner
        $owner.WaitForExit()
        $ownerTicks | Should -BeGreaterThan 0
        $script:deadId = $owner.Id
        $script:deadTicks = $ownerTicks
        $script:recordFile = Join-Path $script:recordDir "$($script:recorded.Id).txt"
        $ok = addOfficeRecord $script:recordDir $script:recorded.Id "EXCEL" (getOfficeStartTicks $script:recorded) $owner.Id $ownerTicks
        $ok | Should -BeTrue
    }

    AfterAll {
        foreach ($p in @($script:recorded, $script:unrecorded)) {
            if ($p -and !$p.HasExited) {
                Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
                [void]$p.WaitForExit(5000)
            }
        }
        if ($null -ne $script:leftoverEnvBefore) { $env:TEBUNKO_GUI_LEFTOVER_FILE = $script:leftoverEnvBefore }
    }

    It "記録のある残り物だけが確認に出る。［今回は終了しない］では止まらず、［終了する］で止まって記録が消える" {
        $S = startGui $script:tool "R1"
        invokeGuiScene $S {
            setGuiStep $S "確認ダイアログ（記録のある 1 件だけ）"
            $dialog = waitGuiWindow $S "Office の終了" -Id "HeadingText" -Text "残ったまま動いています"
            (getGuiText (findGui $dialog -Id "HeadingText")) | Should -BeLike "*Office が 1 件*"
            setGuiStep $S "［今回は終了しない］"
            clickGui $S $dialog "LeftoverCancelButton" "［今回は終了しない］"
            waitGuiWindowClosed $S $dialog "Office の終了"
            $script:recorded.Refresh()
            $script:recorded.HasExited | Should -BeFalse -Because "止めない選択をしたので動いたまま"
            Test-Path -LiteralPath $script:recordFile | Should -BeTrue
            closeGui $S
        }

        $S = startGui $script:tool "R2"
        invokeGuiScene $S {
            setGuiStep $S "起動し直すとまた確認が出る → ［終了する］"
            answerGuiConfirm $S "Office の終了" "残ったまま動いています" "終了する"
            waitGui $S "記録のある偽のプロセスが終わる" ${guiDefaultTimeout} { $script:recorded.Refresh(); $script:recorded.HasExited } | Out-Null
            waitGui $S "結果がステータスに出る" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "StatusText")) -like "Office を 1 件終了しました。*"
            } | Out-Null
            waitGui $S "記録が消える" ${guiDefaultTimeout} { !(Test-Path -LiteralPath $script:recordFile) } | Out-Null
            $script:unrecorded.Refresh()
            $script:unrecorded.HasExited | Should -BeFalse -Because "記録の無いプロセスは止めない"
            closeGui $S
        }
    }

    It "記録の無い偽のプロセスだけが動いているとき、確認は出ない" {
        Test-Path -LiteralPath $script:recordFile | Should -BeFalse -Because "前の It で記録は消えている"
        # 起動時の調べが済んだ印にする記録: もう無いプロセスの記録は、調べる（記録を読む）ときに消える。
        # 消えたのを待てば、調べが済んだ後に確認が出ないことを見られる
        $marker = Join-Path $script:recordDir "$($script:deadId).txt"
        addOfficeRecord $script:recordDir $script:deadId "EXCEL" $script:deadTicks $script:deadId $script:deadTicks | Should -BeTrue
        $S = startGui $script:tool "R3"
        invokeGuiScene $S {
            setGuiStep $S "起動時の調べが済むのを待つ"
            waitGui $S "調べが済んで印の記録が消える" ${guiDefaultTimeout} { !(Test-Path -LiteralPath $marker) } | Out-Null
            setGuiStep $S "確認が出ないこと"
            # 調べの結果は画面のスレッドで受けて、すぐに確認を出す。受けるまでの短い間だけ余分に見る
            $deadline = [datetime]::UtcNow.AddSeconds(3)
            $shown = $false
            while ([datetime]::UtcNow -lt $deadline) {
                if (@(getGuiOtherWindows $S).Count -gt 0) { $shown = $true; break }
                Start-Sleep -Milliseconds 250
            }
            $shown | Should -BeFalse
            $script:unrecorded.Refresh()
            $script:unrecorded.HasExited | Should -BeFalse
            closeGui $S
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        foreach ($p in @($script:recorded, $script:unrecorded)) {
            if (!$p.HasExited) {
                Stop-Process -Id $p.Id -Force
                [void]$p.WaitForExit(5000)
            }
        }
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
