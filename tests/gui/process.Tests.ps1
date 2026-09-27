# 画面のスモークテスト S7: プロセス停止（共通の関数は gui_helpers.ps1）。
# 画面遷移の一覧（docs\design\testing\index.md「画面のスモークテスト」）の #36・#37・#38 を確かめる。
#
# 本物の Office を止めないよう、Office の代わりに偽のプロセス（%windir%\System32\PING.EXE を EXCEL.EXE の名前で写したもの。
# 窓が無いのでバックグラウンドのプロセスになる）を使う。［すべて終了］［バックグラウンドのみ終了］は、確認で［キャンセル］だけを押し、
# 実際に止めるのは偽のプロセスを選んだ［選択したプロセスを終了］だけにする。
# ただし確認を出す作りが壊れていると、押した時点で本物の Office が保存されずに止まる。そのため、手元（GITHUB_ACTIONS が無いとき）で
# 偽のプロセスのほかに Excel・Word・PowerPoint が動いていたら、#37 の段階だけを飛ばし、理由をログに出す。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S7 プロセス停止" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        $fakeDir = Join-Path $TestDrive "fake"
        [void][IO.Directory]::CreateDirectory($fakeDir)
        Copy-Item -LiteralPath (Join-Path $env:windir "System32\PING.EXE") -Destination "$fakeDir\EXCEL.EXE"
        $script:fake = Start-Process "$fakeDir\EXCEL.EXE" -ArgumentList "127.0.0.1", "-n", "600" -WindowStyle Hidden -PassThru
        $null = $script:fake.Handle
    }

    AfterAll {
        if ($script:fake -and !$script:fake.HasExited) { Stop-Process -Id $script:fake.Id -Force -ErrorAction SilentlyContinue }
    }

    It "一覧に偽のプロセスが出て、確認のキャンセルと、選んで終了が動く" {
        $S = startGui $script:tool "S7"
        invokeGuiScene $S {
            $pidText = [string]$script:fake.Id
            $findFakeRow = {
                @(getGuiGridRows (findGui $S.Window -Id "ProcessGrid")) | Where-Object { (getGuiRowTexts $_) -contains $pidText } | Select-Object -First 1
            }

            # ［9 プロセス停止］に切り替えると一覧に偽のプロセスが出て、［更新］でも出る（#36）
            setGuiStep $S "［9 プロセス停止］の一覧に偽のプロセスが出る"
            selectGuiTab $S "KillTab" "ProcessGrid"
            waitGui $S "一覧に偽のプロセス（PID $pidText）" ${guiDefaultTimeout} $findFakeRow | Out-Null
            setGuiStep $S "［更新］"
            clickGui $S $S.Window "RefreshProcessButton" "［更新］"
            waitGui $S "更新の後も一覧に偽のプロセス" ${guiDefaultTimeout} $findFakeRow | Out-Null

            # ［すべて終了］［バックグラウンドのみ終了］は確認で［キャンセル］だけを押す（#37）。
            # 手元で偽のほかに Office が動いていれば、この段階だけを飛ばす
            $others = @(Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.Id -ne $script:fake.Id })
            if ($env:GITHUB_ACTIONS -ne "true" -and $others.Count -gt 0) {
                Write-Host "Office が動いているため、すべて終了の確かめを飛ばした（#37。動いている Office: $(($others | ForEach-Object { "$($_.ProcessName) $($_.Id)" }) -join '・')）"
            } else {
                foreach ($id in "KillAllButton", "KillBackgroundButton") {
                    setGuiStep $S "［$id］→ 確認で［キャンセル］"
                    clickGui $S $S.Window $id "［$id］"
                    $confirm = waitGuiWindow $S "終了の確認（$id）" -Id "HeadingText" -Text "終了しますか"
                    clickGuiByName $S $confirm "キャンセル"
                    waitGuiWindowClosed $S $confirm "終了の確認"
                    $script:fake.Refresh()
                    $script:fake.HasExited | Should -BeFalse -Because "キャンセルしたので偽のプロセスは動いたまま"
                }
            }

            # 偽のプロセスを選んで［選択したプロセスを終了］。確認でキャンセルすると残り、［終了する］で一覧から消える（#38）
            setGuiStep $S "偽のプロセスを選ぶ"
            $row = waitGui $S "一覧に偽のプロセス" ${guiDefaultTimeout} $findFakeRow
            selectGui $row
            setGuiStep $S "［選択したプロセスを終了］→［キャンセル］"
            clickGui $S $S.Window "KillSelectedButton" "［選択したプロセスを終了］"
            $confirm = waitGuiWindow $S "終了の確認" -Id "HeadingText" -Text "終了しますか"
            clickGuiByName $S $confirm "キャンセル"
            waitGuiWindowClosed $S $confirm "終了の確認"
            $script:fake.Refresh()
            $script:fake.HasExited | Should -BeFalse
            setGuiStep $S "［選択したプロセスを終了］→［終了する］"
            clickGui $S $S.Window "KillSelectedButton" "［選択したプロセスを終了］"
            $confirm = waitGuiWindow $S "終了の確認" -Id "HeadingText" -Text "終了しますか"
            clickGuiByName $S $confirm "終了する"
            waitGuiWindowClosed $S $confirm "終了の確認"
            waitGui $S "偽のプロセスが終了する" ${guiDefaultTimeout} { $script:fake.Refresh(); $script:fake.HasExited } | Out-Null
            waitGui $S "一覧から偽のプロセスが消える" ${guiDefaultTimeout} { !(& $findFakeRow) } | Out-Null

            closeGui $S
        }
    }

    It "偽のプロセスが残っていない。利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        if (!$script:fake.HasExited) { Stop-Process -Id $script:fake.Id -Force }
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
