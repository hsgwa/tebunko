# 画面のスモークテスト: 起動時の「前回残った Office の確認」。
# 本物の Office には触らない。偽の行（環境変数 TEBUNKO_GUI_LEFTOVER_FILE が指す JSON）を場面から差し込み、
# 確認ダイアログ・詳細の開け閉め・［今回は終了しない］・［終了する］の結果（ステータス）を確かめる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"

    function newLeftoverFile {
        param ([string]$Path, [string]$ThirdStatus)
        $rows = @(
            @{ Id = 12840; ProcessName = "EXCEL"; StartTime = "2030-10-04T18:32:00" }
            @{ Id = 15012; ProcessName = "EXCEL"; StartTime = "2030-10-04T18:32:00" }
            @{ Id = 9316; ProcessName = "WINWORD"; StartTime = "2030-10-04T18:41:00"; StopStatus = $ThirdStatus }
        )
        $json = ConvertTo-Json -InputObject $rows
        [IO.File]::WriteAllText($Path, $json, (New-Object Text.UTF8Encoding($false)))
    }
}

Describe "起動時の前回残った Office の確認" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
    }
    AfterEach {
        Remove-Item Env:\TEBUNKO_GUI_LEFTOVER_FILE -ErrorAction SilentlyContinue
    }

    It "［今回は終了しない］を選ぶと、何も変えずに閉じる。詳細を開くと PID が並ぶ" {
        $tool = newGuiTool $TestDrive
        $file = Join-Path $TestDrive "rows_cancel.json"
        newLeftoverFile $file "Stopped"
        $env:TEBUNKO_GUI_LEFTOVER_FILE = $file
        $S = startGui $tool "L1"
        invokeGuiScene $S {
            setGuiStep $S "確認ダイアログ"
            $dialog = waitGuiWindow $S "Office の終了" -Id "HeadingText" -Text "残ったまま動いています"
            (getGuiText (findGui $dialog -Id "HeadingText")) | Should -BeLike "*Office が 3 件*"
            (getGuiTexts $dialog) -join " " | Should -BeLike "*編集中のファイルは閉じません*"

            setGuiStep $S "詳細を開く"
            toggleGui (findGui $dialog -Id "LeftoverDetailToggle")
            waitGui $S "詳細に PID が出る" ${guiDefaultTimeout} { (@(getGuiTexts $dialog) -join " ") -like "*12840*9316*" } | Out-Null
            getGuiToggleState (findGui $dialog -Id "LeftoverDetailToggle") | Should -Be "On"

            setGuiStep $S "［今回は終了しない］"
            clickGui $S $dialog "LeftoverCancelButton" "［今回は終了しない］"
            waitGuiWindowClosed $S $dialog "Office の終了"
            (getGuiText (findGui $S.Window -Id "StatusText")) | Should -Not -BeLike "*Office を*終了*"
            closeGui $S
        }
    }

    It "［終了する］を選ぶと、変わっていた行の PID がステータスに出る（偽の行だけを扱い、本物は止めない）" {
        $tool = newGuiTool (Join-Path $TestDrive "stop")
        $file = Join-Path $TestDrive "rows_stop.json"
        newLeftoverFile $file "Changed"
        $env:TEBUNKO_GUI_LEFTOVER_FILE = $file
        $S = startGui $tool "L2"
        invokeGuiScene $S {
            answerGuiConfirm $S "Office の終了" "残ったまま動いています" "終了する"
            waitGui $S "結果がステータスに出る" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "StatusText")) -like "*PID 9316 は確認の後に別のプロセスに変わったため*"
            } | Out-Null
            (getGuiText (findGui $S.Window -Id "StatusText")) | Should -BeLike "Office を 2 件終了しました。*"
            closeGui $S
        }
    }

    AfterAll {
        $after = getGuiEnvSnapshot
        compareGuiEnvSnapshot $script:envBefore $after | Should -BeNullOrEmpty
    }
}
