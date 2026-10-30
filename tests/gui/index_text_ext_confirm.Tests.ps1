# 画面のスモークテスト: 更新の確認ダイアログの「取り込み対象の拡張子」の［変更…］。
# ［OK］で選びを書いて数え直し、確認を出し直す。［キャンセル］では確認に戻る（設定は書かない）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "確認ダイアログの［変更…］" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        $script:source = Join-Path $TestDrive "元のフォルダ\営業"
        newGuiSourceFolder $script:source
        foreach ($name in @("議事録.txt", "動作記録.log")) {
            [IO.File]::WriteAllText((Join-Path $script:source $name), "サンプル`r`n", (New-Object System.Text.UTF8Encoding($true)))
        }
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "営業"; path = $script:source; enabled = $true }) -Force
        writeGuiConfig $script:tool.Dir $config
    }

    It "［キャンセル］で確認に戻り、［OK］で設定に書いて確認が出し直される" {
        $S = startGui $script:tool "TE1"
        invokeGuiScene $S {
            selectGuiTab $S "IndexTab" "IndexingButton"
            setGuiStep $S "［すべて更新］→ 確認の［変更…］"
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $confirm "ChangeIndexTextExtensionsButton" "確認の拡張子の［変更…］"
            $dialog = waitGuiWindow $S "取り込み対象の拡張子のダイアログ" -Id "TextExtOkButton"

            setGuiStep $S "［キャンセル］→ 確認のまま（設定は書かない）"
            clickGui $S $dialog "TextExtCancelButton" "ダイアログの［キャンセル］"
            waitGui $S "ダイアログが閉じる" ${guiDefaultTimeout} { @(getGuiOtherWindows $S | Where-Object { findGui $_ -Id "TextExtOkButton" }).Count -eq 0 } | Out-Null
            (findGui $confirm -Id "StartButton") | Should -Not -BeNullOrEmpty
            (readTextExtensionSettings $script:tool.Config).IndexExcludes.Count | Should -Be 0

            setGuiStep $S "［変更…］で .log を外して［OK］→ 確認が出し直される"
            clickGui $S $confirm "ChangeIndexTextExtensionsButton" "確認の拡張子の［変更…］"
            $dialog = waitGuiWindow $S "取り込み対象の拡張子のダイアログ" -Id "TextExtOkButton"
            toggleGui (findGui $dialog -Name ".log" -Type "CheckBox")
            clickGui $S $dialog "TextExtOkButton" "ダイアログの［OK］"
            waitGui $S "ダイアログと最初の確認が閉じる" ${guiDefaultTimeout} { @(getGuiOtherWindows $S | Where-Object { (findGui $_ -Id "TextExtOkButton") -or (findGui $_ -Id "StartButton") }).Count -eq 0 } | Out-Null
            (@((readTextExtensionSettings $script:tool.Config).IndexExcludes["営業"]) -join ",") | Should -Be ".log"
            $again = waitGuiWindow $S "出し直された確認" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $again "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $again "出し直された確認"
            closeGui $S
        }
    }

    It "取り込み済みのインデックスで、外す拡張子を足して［OK］しても、確認が出し直される（更新の流れが消えない）" {
        $tool2 = newGuiTool (Join-Path $TestDrive "ws2")
        $config = readGuiConfig $tool2
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "営業"; path = $script:source; enabled = $true }) -Force
        writeGuiConfig $tool2.Dir $config
        $S = startGui $tool2 "TE2"
        invokeGuiScene $S {
            selectGuiTab $S "IndexTab" "IndexingButton"
            setGuiStep $S "取り込む"
            startGuiIndexing $S
            waitGui $S "取り込みの完了" ${guiIndexTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*" -and (findGui $S.Window -Id "IndexingButton").Current.IsEnabled
            } | Out-Null

            setGuiStep $S ".log を対象外にして［すべて更新］→ 確認の［変更…］で .txt も外して［OK］"
            writeIndexTextExcludes "営業" @(".log") $tool2.Config
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $confirm "ChangeIndexTextExtensionsButton" "確認の拡張子の［変更…］"
            $dialog = waitGuiWindow $S "取り込み対象の拡張子のダイアログ" -Id "TextExtOkButton"
            toggleGui (findGui $dialog -Name ".txt" -Type "CheckBox")
            clickGui $S $dialog "TextExtOkButton" "ダイアログの［OK］"
            waitGui $S "ダイアログと最初の確認が閉じる" ${guiDefaultTimeout} { @(getGuiOtherWindows $S | Where-Object { (findGui $_ -Id "TextExtOkButton") -or (findGui $_ -Id "StartButton") }).Count -eq 0 } | Out-Null
            $again = waitGuiWindow $S "出し直された確認" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $again "CancelButton" "［キャンセル］"
            waitGui $S "取りやめが終わる" ${guiIndexTimeout} { (findGui $S.Window -Id "IndexingButton").Current.IsEnabled } | Out-Null
            closeGui $S
        }
    }

    AfterAll {
        $after = getGuiEnvSnapshot
        compareGuiEnvSnapshot $script:envBefore $after | Should -BeNullOrEmpty
    }
}
