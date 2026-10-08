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
        $script:recordFile = Join-Path $script:recordDir "$($script:recorded.Id).txt"
        $ok = addOfficeRecord $script:recordDir $script:recorded.Id "EXCEL" (getOfficeStartTicks $script:recorded) $owner.Id $ownerTicks
        $ok | Should -BeTrue
    }

    AfterAll {
        foreach ($p in @($script:recorded, $script:unrecorded)) {
            if ($p -and !$p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
        }
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
        $S = startGui $script:tool "R3"
        invokeGuiScene $S {
            setGuiStep $S "確認が出ないこと"
            $deadline = [datetime]::UtcNow.AddSeconds(8)
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
            if (!$p.HasExited) { Stop-Process -Id $p.Id -Force }
        }
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
