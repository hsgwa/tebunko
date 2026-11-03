# 画面のスモークテスト: テキストの開き方（拡張子ごとに「既定のアプリで開く」か「エディタで開く」か）を選ぶダイアログ（［設定］の「テキストファイル」の［変更…］）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"

    function selectGuiWay {
        # ダイアログの行（拡張子）の開き方のボタン（RadioButton。Name は「<拡張子> ： <文言>」）を選ぶ
        param ($S, $Dialog, [string]$Extension, [string]$WayText)
        $radio = waitGuiByName $S $Dialog "$Extension ： $WayText" "RadioButton"
        selectGui $radio
    }

    function getGuiSelectedWay {
        # ダイアログの行（拡張子）で、選ばれている開き方の文言
        param ($Dialog, [string]$Extension)
        foreach ($wayText in @("既定のアプリで開く", "エディタで開く")) {
            $radio = findGui $Dialog -Name "$Extension ： $wayText" -Type "RadioButton"
            if ($radio -and (isGuiSelected $radio)) { return $wayText }
        }
        return ""
    }
}

Describe "テキストの開き方のダイアログ" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
    }

    It "［設定］の要約とエディタの行が出て、［変更…］でダイアログが開き、選びが保存され、開き直すと残っている" {
        $tool = newGuiTool $TestDrive @{ notepadTextExtensions = @(".css") }
        $S = startGui $tool "TX1"
        invokeGuiScene $S {
            setGuiStep $S "［設定］の要約とエディタの行"
            selectGuiTab $S "SettingsTab" "ChangeTextExtensionsButton"
            (getGuiText (findGui $S.Window -Id "TextExtensionSummary")) | Should -Be "変更したもの 1 個（エディタで開く）"
            (getGuiText (findGui $S.Window -Id "EditorHint")) | Should -Be "「エディタで開く」に使うアプリです。今はメモ帳（既定）です。"
            (findGui $S.Window -Id "ChangeEditorButton").Current.IsEnabled | Should -BeTrue
            (findGui $S.Window -Id "ResetEditorButton").Current.IsEnabled | Should -BeFalse -Because "メモ帳（既定）のときは戻すものが無い"
            testGuiVisible (findGui $S.Window -Id "EditorPathText") | Should -BeFalse
            testGuiVisible (findGui $S.Window -Id "EditorErrorText") | Should -BeFalse

            setGuiStep $S "［変更…］でダイアログを開く"
            clickGui $S $S.Window "ChangeTextExtensionsButton" "［変更…］"
            $dialog = waitGuiWindow $S "テキストの開き方のダイアログ" -Id "TextExtOkButton"
            @(findAllGui $dialog -Type "RadioButton").Count | Should -Be 150
            (getGuiText (findGui $dialog -Id "TextExtSummary")) | Should -Be "変更したもの 1 個（エディタで開く）"
            (getGuiText (findGui $dialog -Id "TextExtEditorName")) | Should -Be "メモ帳（既定）"
            (findGui $dialog -Id "TextExtEditorChangeButton").Current.IsEnabled | Should -BeTrue
            (findGui $dialog -Id "TextExtEditorResetButton").Current.IsEnabled | Should -BeFalse
            testGuiVisible (findGui $dialog -Id "TextExtEditorPath") | Should -BeFalse
            testGuiVisible (findGui $dialog -Id "TextExtEditorError") | Should -BeFalse
            (getGuiText (findGui $dialog -Id "NoteText")) | Should -BeLike "*「エディタで開く」は、下の「エディタ」のアプリ（既定はメモ帳）で開きます。*"
            (findGui $dialog -Name "将来の対応") | Should -BeNullOrEmpty
            # 実行・登録になる形式は「既定のアプリで開く」を選べない
            (findGui $dialog -Name ".bat ： 既定のアプリで開く" -Type "RadioButton").Current.IsEnabled | Should -BeFalse
            getGuiSelectedWay $dialog ".bat" | Should -Be "エディタで開く"
            getGuiSelectedWay $dialog ".css" | Should -Be "エディタで開く"
            getGuiSelectedWay $dialog ".txt" | Should -Be "既定のアプリで開く"

            setGuiStep $S "［変更したものだけ］"
            toggleGui (findGui $dialog -Id "TextExtChangedOnly")
            waitGui $S "要約が変わる" ${guiDefaultTimeout} { (getGuiText (findGui $dialog -Id "TextExtSummary")) -eq "変更したもの 1 個を表示" } | Out-Null
            toggleGui (findGui $dialog -Id "TextExtChangedOnly")

            setGuiStep $S "［キャンセル］"
            clickGui $S $dialog "TextExtCancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "テキストの開き方"

            setGuiStep $S "選びを変えて［OK］"
            clickGui $S $S.Window "ChangeTextExtensionsButton" "［変更…］"
            $dialog = waitGuiWindow $S "テキストの開き方のダイアログ" -Id "TextExtOkButton"
            selectGuiWay $S $dialog ".html" "エディタで開く"
            waitGui $S "要約が変わる" ${guiDefaultTimeout} { (getGuiText (findGui $dialog -Id "TextExtSummary")) -eq "変更したもの 2 個（エディタで開く）" } | Out-Null
            clickGui $S $dialog "TextExtOkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "テキストの開き方"
            $config = readGuiConfig $tool
            (@($config.notepadTextExtensions) -join ",") | Should -Be ".html,.css"  # 固定の一覧の順
            (getGuiText (findGui $S.Window -Id "TextExtensionSummary")) | Should -Be "変更したもの 2 個（エディタで開く）"

            setGuiStep $S "開き直すと選びが残っている"
            clickGui $S $S.Window "ChangeTextExtensionsButton" "［変更…］"
            $dialog = waitGuiWindow $S "テキストの開き方のダイアログ" -Id "TextExtOkButton"
            getGuiSelectedWay $dialog ".html" | Should -Be "エディタで開く"
            getGuiSelectedWay $dialog ".css" | Should -Be "エディタで開く"
            getGuiSelectedWay $dialog ".csv" | Should -Be "既定のアプリで開く"
            clickGui $S $dialog "TextExtCancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "テキストの開き方"
            closeGui $S
        }
    }

    It "エディタが選んである（textEditorPath）と、名前・パスの段が出て、［メモ帳に戻す］が押せ、押すと既定に戻る（設定の行とダイアログの行）" {
        $tool = newGuiTool $TestDrive @{ textEditorPath = "C:\Tools\sakura.exe" }
        $S = startGui $tool "TX2"
        invokeGuiScene $S {
            setGuiStep $S "［設定］の行"
            selectGuiTab $S "SettingsTab" "ChangeTextExtensionsButton"
            (getGuiText (findGui $S.Window -Id "EditorHint")) | Should -Be "「エディタで開く」に使うアプリです。今はsakuraです。"
            testGuiVisible (findGui $S.Window -Id "EditorPathText") | Should -BeTrue
            (getGuiText (findGui $S.Window -Id "EditorPathText")) | Should -Be "C:\Tools\sakura.exe"
            (findGui $S.Window -Id "ResetEditorButton").Current.IsEnabled | Should -BeTrue

            setGuiStep $S "ダイアログの行"
            clickGui $S $S.Window "ChangeTextExtensionsButton" "［変更…］"
            $dialog = waitGuiWindow $S "テキストの開き方のダイアログ" -Id "TextExtOkButton"
            (getGuiText (findGui $dialog -Id "TextExtEditorName")) | Should -Be "sakura"
            (getGuiText (findGui $dialog -Id "TextExtEditorPath")) | Should -Be "C:\Tools\sakura.exe"
            (findGui $dialog -Id "TextExtEditorResetButton").Current.IsEnabled | Should -BeTrue

            setGuiStep $S "ダイアログの［メモ帳に戻す］"
            clickGui $S $dialog "TextExtEditorResetButton" "［メモ帳に戻す］"
            waitGui $S "名前が戻る" ${guiDefaultTimeout} { (getGuiText (findGui $dialog -Id "TextExtEditorName")) -eq "メモ帳（既定）" } | Out-Null
            testGuiVisible (findGui $dialog -Id "TextExtEditorPath") | Should -BeFalse
            (findGui $dialog -Id "TextExtEditorResetButton").Current.IsEnabled | Should -BeFalse
            clickGui $S $dialog "TextExtCancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "テキストの開き方"
            (getGuiText (findGui $S.Window -Id "EditorHint")) | Should -Be "「エディタで開く」に使うアプリです。今はメモ帳（既定）です。"
            testGuiVisible (findGui $S.Window -Id "EditorPathText") | Should -BeFalse
            (readGuiConfig $tool).textEditorPath | Should -Be ""
            closeGui $S
        }
    }

    AfterAll {
        $after = getGuiEnvSnapshot
        compareGuiEnvSnapshot $script:envBefore $after | Should -BeNullOrEmpty
    }
}
