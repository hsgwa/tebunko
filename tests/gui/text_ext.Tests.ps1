# 画面のスモークテスト: テキストファイルの拡張子の扱いを選ぶダイアログ（［設定］の［変更…］と［アクション ▾］の［テキストの拡張子…］）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"

    function selectGuiComboItem {
        # ダイアログの ComboBox（Name で探す）を開き、文言の項目を選ぶ
        param ($S, $Dialog, [string]$ComboName, [string]$ItemText)
        $combo = waitGuiByName $S $Dialog $ComboName "ComboBox"
        $combo.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand()
        $item = waitGuiByName $S $combo $ItemText "ListItem"
        selectGui $item
        $combo.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Collapse()
    }

    function getGuiComboSelected {
        # ダイアログの ComboBox（Name で探す）で、今選ばれている項目の文言
        param ($Dialog, [string]$ComboName)
        $combo = findGui $Dialog -Name $ComboName -Type "ComboBox"
        $selected = @($combo.GetCurrentPattern([Windows.Automation.SelectionPattern]::Pattern).Current.GetSelection())
        return $selected[0].Current.Name
    }
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

            setGuiStep $S "選びを変えて［OK］→ 保存され、開き直すと残っている"
            clickGuiAction $S "ActionTextExtensions" "［テキストの拡張子…］"
            $dialog = waitGuiWindow $S "テキストファイルの拡張子のダイアログ" -Id "TextExtOkButton"
            selectGuiComboItem $S $dialog ".txt の扱い" "取り込まない"
            clickGui $S $dialog "TextExtOkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "テキストファイルの拡張子"
            $config = readGuiConfig $tool
            (@($config.excludedTextExtensions) -join ",") | Should -Be ".txt,.log"  # 固定の一覧の順
            (@($config.notepadTextExtensions) -join ",") | Should -Be ".html"

            clickGuiAction $S "ActionTextExtensions" "［テキストの拡張子…］"
            $dialog = waitGuiWindow $S "テキストファイルの拡張子のダイアログ" -Id "TextExtOkButton"
            getGuiComboSelected $dialog ".txt の扱い" | Should -Be "取り込まない"
            getGuiComboSelected $dialog ".html の扱い" | Should -Be "メモ帳で開く"
            getGuiComboSelected $dialog ".csv の扱い" | Should -Be "既定のアプリで開く"
            clickGui $S $dialog "TextExtCancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "テキストファイルの拡張子"
            closeGui $S
        }
    }

    AfterAll {
        $after = getGuiEnvSnapshot
        compareGuiEnvSnapshot $script:envBefore $after | Should -BeNullOrEmpty
    }
}
