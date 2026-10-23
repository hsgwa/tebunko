# 画面のスモークテスト: テキストファイルの拡張子の扱いを選ぶダイアログ（［設定］の［変更…］と［アクション ▾］の［テキストの拡張子…］）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "テキストファイルの拡張子のダイアログ" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
    }

    It "［設定］の要約が出て、［変更…］と［アクション ▾］のどちらからも同じダイアログが開き、［OK］で設定を保つ" {
        $tool = newGuiTool $TestDrive @{ excludedTextExtensions = @(".log"); notepadTextExtensions = @(".html") }
        $S = startGui $tool "TX1"
        invokeGuiScene $S {
            setGuiStep $S "［設定］の要約"
            selectGuiTab $S "SettingsTab" "ChangeTextExtensionsButton"
            (getGuiText (findGui $S.Window -Id "TextExtensionSummary")) | Should -Be "取り込む 75 個のうち 74 個 ・ メモ帳で開く 16 個"

            setGuiStep $S "［変更…］でダイアログを開く"
            clickGui $S $S.Window "ChangeTextExtensionsButton" "［変更…］"
            $dialog = waitGuiWindow $S "テキストファイルの拡張子のダイアログ" -Id "TextExtOkButton"
            @(findAllGui $dialog -Type "ComboBox").Count | Should -Be 75
            (getGuiTexts $dialog) -join " " | Should -BeLike "*次にインデックス作成で件数を数えたときにインデックスから外します*"

            setGuiStep $S "［キャンセル］"
            clickGui $S $dialog "TextExtCancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "テキストファイルの拡張子"

            setGuiStep $S "［アクション ▾］の［テキストの拡張子…］から開いて［OK］"
            selectGuiTab $S "IndexTab" "ActionsButton"
            clickGuiAction $S "ActionTextExtensions" "［テキストの拡張子…］"
            $dialog = waitGuiWindow $S "テキストファイルの拡張子のダイアログ" -Id "TextExtOkButton"
            clickGui $S $dialog "TextExtOkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "テキストファイルの拡張子"

            $config = readGuiConfig $tool
            (@($config.excludedTextExtensions) -join ",") | Should -Be ".log"
            (@($config.notepadTextExtensions) -join ",") | Should -Be ".html"
            closeGui $S
        }
    }

    AfterAll {
        $after = getGuiEnvSnapshot
        compareGuiEnvSnapshot $script:envBefore $after | Should -BeNullOrEmpty
    }
}
