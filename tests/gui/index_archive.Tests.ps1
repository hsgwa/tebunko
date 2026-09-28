# 画面のスモークテスト: ［1 インデックス管理］のエクスポート・インポート（gui_helpers.ps1 の useGuiFileOpenPicker）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S2b エクスポート・インポート" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        $script:source = Join-Path $TestDrive "元のフォルダ\営業"
        newGuiSourceFolder $script:source
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "営業"; path = $script:source; enabled = $true }) -Force
        writeGuiConfig $script:tool.Dir $config
        $script:exportDest = Join-Path $TestDrive "エクスポート先"
        [void][IO.Directory]::CreateDirectory($script:exportDest)
    }

    It "エクスポート → 削除 → インポートで、同じインデックスが一覧に戻る" {
        $S = startGui $script:tool "S2b"
        invokeGuiScene $S {
            setGuiStep $S "起動時のタブ・取り込み"
            selectGuiTab $S "IndexTab" "IndexingButton"
            startGuiIndexing $S
            waitGui $S "取り込みの完了" ${guiIndexTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*"
            } | Out-Null

            setGuiStep $S "［エクスポート…］"
            $row = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            selectGui $row
            clickGui $S $S.Window "ExportIndexButton" "［エクスポート…］"
            useGuiFolderPicker $S $script:exportDest
            waitGui $S "書き出しの完了" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "StatusText")) -like "*書き出しました*"
            } | Out-Null
            $zip = @(Get-ChildItem -LiteralPath $script:exportDest -Filter "*.zip")
            $zip.Count | Should -Be 1

            setGuiStep $S "一覧から削除（インポートし直すため）"
            clickGui $S $S.Window "RemoveIndexButton" "［削除］"
            answerGuiConfirm $S "削除の確認" "一覧から削除しますか" "削除する"
            waitGui $S "一覧から消える" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count -eq 0 } | Out-Null

            setGuiStep $S "［インポート…］"
            clickGui $S $S.Window "ImportIndexButton" "［インポート…］"
            useGuiFileOpenPicker $S $zip[0].FullName
            $dialog = waitGuiWindow $S "インポートのダイアログ" -Id "NameBox"
            getGuiValue (findGui $dialog -Id "NameBox") | Should -Be "営業"
            getGuiValue (findGui $dialog -Id "FolderBox") | Should -Be $script:source
            clickGui $S $dialog "OkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "インポートのダイアログ"

            setGuiStep $S "インポートの完了"
            waitGui $S "インポートの完了" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "StatusText")) -like "*インポートしました*"
            } | Out-Null
            $row = waitGui $S "一覧に戻る" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            (getGuiRowTexts $row) | Should -Contain "営業"
            Test-Path -LiteralPath "$($script:tool.Work)\content_index\営業" | Should -BeTrue

            closeGui $S
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
