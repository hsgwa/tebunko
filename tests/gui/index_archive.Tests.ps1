# 画面のスモークテスト: ［インデックス管理］のエクスポート・インポート（gui_helpers.ps1 の useGuiFileOpenPicker）。
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
            checkGuiRow $S $row
            clickGuiAction $S "ActionExport" "［エクスポート…］"
            $dialog = waitGuiWindow $S "エクスポートのダイアログ" -Id "ExportPathBox"
            getGuiText (findGui $dialog -Id "ExportTitleText") | Should -Be "営業" -Because "見出しの代わりに対象の名前を枠に出す"
            getGuiText (findGui $dialog -Id "ExportFileText") | Should -BeLike "営業_インデックス_*.zip として保存します"
            getGuiText (findGui $dialog -Id "CautionText") | Should -BeLike "このファイルには、元のファイルの本文と元のフォルダの場所が含まれます。*"
            setGuiText $S (findGui $dialog -Id "ExportPathBox") $script:exportDest
            clickGui $S $dialog "ExportButton" "［エクスポート］"
            waitGuiWindowClosed $S $dialog "エクスポートのダイアログ"
            waitGui $S "書き出しの完了" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "StatusText")) -like "*書き出しました*"
            } | Out-Null
            (getGuiText (findGui $S.Window -Id "StatusText")).Contains("インデックス [営業] を") | Should -BeTrue -Because "書き出したインデックスの名前がステータスに出る"
            $zip = @(Get-ChildItem -LiteralPath $script:exportDest -Filter "*.zip")
            $zip.Count | Should -Be 1
            $script:zipPath = $zip[0].FullName

            setGuiStep $S "一覧から削除（インポートし直すため）"
            clickGuiAction $S "ActionDelete" "［削除…］"
            answerGuiConfirm $S "削除の確認" "インデックスを削除しますか" "削除する"
            waitGui $S "一覧から消える" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count -eq 0 } | Out-Null

            setGuiStep $S "［インポート…］"
            clickGuiAction $S "ActionImport" "［インポート…］"
            useGuiFileOpenPicker $S $zip[0].FullName
            $dialog = waitGuiWindow $S "インポートのダイアログ" -Id "NameBox"
            getGuiValue (findGui $dialog -Id "NameBox") | Should -Be "営業"
            getGuiValue (findGui $dialog -Id "FolderBox") | Should -Be $script:source
            getGuiText (findGui $dialog -Id "ImportTitleText") | Should -Be "営業" -Because "見出しの代わりに zip の中身を枠に出す"
            getGuiText (findGui $dialog -Id "ImportSizeText") | Should -Match "^\d+ ファイル・\d+\.\d MB"
            waitGui $S "元のフォルダが見つかった表示" ${guiDefaultTimeout} { (getGuiText (findGui $dialog -Id "SourceCheckText")) -eq "見つかりました" } | Out-Null
            clickGui $S $dialog "ImportButton" "［インポート］"
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

    It "2 件にチェックを付けると、まとめてエクスポートし、まとめて削除できる" {
        $tool3 = newGuiTool (Join-Path $TestDrive "まとめて")
        $sourceA = Join-Path $TestDrive "まとめて元\営業"
        $sourceB = Join-Path $TestDrive "まとめて元\総務"
        newGuiSourceFolder $sourceA
        newGuiSourceFolder $sourceB
        $config = readGuiConfig $tool3
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(
            @{ name = "営業"; path = $sourceA; enabled = $true }, @{ name = "総務"; path = $sourceB; enabled = $true }) -Force
        writeGuiConfig $tool3.Dir $config
        $dest = Join-Path $TestDrive "まとめて先"
        [void][IO.Directory]::CreateDirectory($dest)

        $S = startGui $tool3 "S2d"
        invokeGuiScene $S {
            setGuiStep $S "取り込み"
            selectGuiTab $S "IndexTab" "IndexingButton"
            startGuiIndexing $S
            waitGui $S "取り込みの完了" ${guiIndexTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*"
            } | Out-Null

            setGuiStep $S "2 件にチェックを付けてまとめて［エクスポート…］"
            foreach ($row in @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))) { checkGuiRow $S $row }
            clickGuiAction $S "ActionExport" "［エクスポート…］"
            $dialog = waitGuiWindow $S "エクスポートのダイアログ" -Id "ExportPathBox"
            setGuiText $S (findGui $dialog -Id "ExportPathBox") $dest
            clickGui $S $dialog "ExportButton" "［エクスポート］"
            waitGuiWindowClosed $S $dialog "エクスポートのダイアログ"
            waitGui $S "書き出しの完了" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "StatusText")) -like "*2 件のインデックスをエクスポートしました*"
            } | Out-Null
            @(Get-ChildItem -LiteralPath $dest -Filter "*.zip").Count | Should -Be 2

            setGuiStep $S "まとめて［削除…］"
            clickGuiAction $S "ActionDelete" "［削除…］"
            answerGuiConfirm $S "まとめての削除の確認" "2 件のインデックスを削除しますか" "削除する"
            waitGui $S "一覧から消える" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count -eq 0 } | Out-Null
            (getGuiText (findGui $S.Window -Id "StatusText")) | Should -BeLike "*2 件のインデックスを削除しました*"

            closeGui $S
        }
    }

    It "別のワークスペース: ［設定］で空のフォルダに切り替えてから、書き出した zip をインポートすると、その中にインデックスができる" {
        $tool2 = newGuiTool (Join-Path $TestDrive "別のPC")
        $workspaceB = Join-Path $TestDrive "別のワークスペース"
        [void][IO.Directory]::CreateDirectory($workspaceB)

        $S = startGui $tool2 "S2c"
        invokeGuiScene $S {
            setGuiStep $S "［設定］で空のフォルダに切り替える"
            selectGuiTab $S "SettingsTab" "ChangeWorkspaceButton"
            clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
            useGuiFolderPicker $S $workspaceB
            answerGuiConfirm $S "ワークスペースを変える確認" "へ移動します" "移動する"
            waitGui $S "ワークスペースが切り替わる" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "WorkspaceText")) -eq $workspaceB } | Out-Null

            setGuiStep $S "［インポート…］"
            selectGuiTab $S "IndexTab" "ActionsButton"
            clickGuiAction $S "ActionImport" "［インポート…］"
            useGuiFileOpenPicker $S $script:zipPath
            $dialog = waitGuiWindow $S "インポートのダイアログ" -Id "NameBox"
            getGuiValue (findGui $dialog -Id "NameBox") | Should -Be "営業"
            clickGui $S $dialog "ImportButton" "［インポート］"
            waitGuiWindowClosed $S $dialog "インポートのダイアログ"

            waitGui $S "インポートの完了" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "StatusText")).Contains("インデックス [営業] をインポートしました")
            } | Out-Null
            $row = waitGui $S "一覧に出る" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            (getGuiRowTexts $row) | Should -Contain "営業"
            Test-Path -LiteralPath "$workspaceB\content_index\営業" | Should -BeTrue -Because "切り替えた先のワークスペースに入る"
            Test-Path -LiteralPath "$($tool2.Work)\content_index\営業" | Should -BeFalse -Because "切り替える前のワークスペースには入らない"

            closeGui $S
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
