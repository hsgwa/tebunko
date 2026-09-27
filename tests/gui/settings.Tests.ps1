# 画面のスモークテスト S5: ワークスペースの変更、S6: 既定のワークスペース（CI だけ）（共通の関数は gui_helpers.ps1）。
# 画面遷移の一覧（docs\design\testing\index.md「画面のスモークテスト」）の #30〜#34 を S5 で、#5・#22・#35 を S6 で確かめる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S5 ワークスペースの変更" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        newGuiSampleIndex $script:tool $TestDrive
        # 選ぶフォルダ: 空のフォルダ・空でないフォルダ・インデックスのあるフォルダ（ほかの人が共有したワークスペースに見立てる）
        $script:emptyDir = Join-Path $TestDrive "空のフォルダ"
        $script:nonEmptyDir = Join-Path $TestDrive "ほかのファイルがあるフォルダ"
        $script:sharedDir = Join-Path $TestDrive "共有のワークスペース"
        [void][IO.Directory]::CreateDirectory($script:emptyDir)
        [void][IO.Directory]::CreateDirectory($script:nonEmptyDir)
        Set-Content -LiteralPath "$script:nonEmptyDir\メモ.txt" -Value "メモ" -Encoding UTF8
        newGuiSampleIndex @{ Work = $script:sharedDir } (Join-Path $TestDrive "共有の元") "共有"
    }

    It "［変更…］のキャンセル・使えないフォルダ・空のフォルダ・空でないフォルダ・インデックスのあるフォルダが動く" {
        $S = startGui $script:tool "S5"
        invokeGuiScene $S {
            $workspaceText = { getGuiText (findGui $S.Window -Id "WorkspaceText") }
            $changeWorkspace = {
                param ($path)
                clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
                useGuiFolderPicker $S $path
            }

            setGuiStep $S "［8 設定］を開く"
            selectGuiTab $S "SettingsTab" "ChangeWorkspaceButton"
            & $workspaceText | Should -Be $script:tool.Work

            # ［変更…］→ OS のフォルダ選択で［キャンセル］すると変わらない（#30）
            setGuiStep $S "［変更…］→ フォルダ選択で［キャンセル］"
            & $changeWorkspace $null
            Start-Sleep -Milliseconds 500
            & $workspaceText | Should -Be $script:tool.Work

            # 今のインデックスの中のフォルダは使えない（メッセージボックス）（#34）
            setGuiStep $S "［変更…］→ 使えないフォルダ（今のインデックスの中）"
            & $changeWorkspace "$($script:tool.Work)\index\営業"
            closeGuiMessage $S "インデックスのフォルダの中です" "使えないフォルダの警告" | Out-Null
            & $workspaceText | Should -Be $script:tool.Work

            # 空のフォルダ: 確認でキャンセルすると変わらず、実行すると中身が移って切り替わる（#31）
            setGuiStep $S "［変更…］→ 空のフォルダ → 確認で［キャンセル］"
            & $changeWorkspace $script:emptyDir
            $confirm = waitGuiWindow $S "ワークスペースを変える確認" -Id "HeadingText" -Text "ワークスペースを変えますか"
            clickGuiByName $S $confirm "キャンセル"
            waitGuiWindowClosed $S $confirm "ワークスペースを変える確認"
            & $workspaceText | Should -Be $script:tool.Work
            setGuiStep $S "［変更…］→ 空のフォルダ → 確認で［ワークスペースを変える］"
            & $changeWorkspace $script:emptyDir
            $confirm = waitGuiWindow $S "ワークスペースを変える確認" -Id "HeadingText" -Text "ワークスペースを変えますか"
            clickGuiByNameLike $S $confirm "ワークスペースを変える*"
            waitGuiWindowClosed $S $confirm "ワークスペースを変える確認"
            waitGui $S "ワークスペースが空のフォルダに変わる" ${guiDefaultTimeout} { (& $workspaceText) -eq $script:emptyDir } | Out-Null
            Test-Path -LiteralPath "$($script:emptyDir)\index\営業" | Should -BeTrue -Because "今のワークスペースの中身が移る"
            (readGuiConfig $script:tool).workspaceFolder | Should -Be $script:emptyDir

            # 空でないフォルダ: 確認でキャンセル、中にフォルダを作って切り替わる（#32）
            setGuiStep $S "［変更…］→ 空でないフォルダ → 確認で［キャンセル］"
            & $changeWorkspace $script:nonEmptyDir
            $confirm = waitGuiWindow $S "空でないフォルダの確認" -Id "HeadingText" -Text "空ではありません"
            clickGuiByName $S $confirm "キャンセル"
            waitGuiWindowClosed $S $confirm "空でないフォルダの確認"
            & $workspaceText | Should -Be $script:emptyDir
            setGuiStep $S "［変更…］→ 空でないフォルダ → 中に「workspace」フォルダを作る"
            & $changeWorkspace $script:nonEmptyDir
            $confirm = waitGuiWindow $S "空でないフォルダの確認" -Id "HeadingText" -Text "空ではありません"
            clickGuiByNameLike $S $confirm "中に*"
            waitGuiWindowClosed $S $confirm "空でないフォルダの確認"
            waitGui $S "ワークスペースが中の workspace に変わる" ${guiDefaultTimeout} { (& $workspaceText) -eq "$($script:nonEmptyDir)\workspace" } | Out-Null
            Test-Path -LiteralPath "$($script:nonEmptyDir)\workspace\index\営業" | Should -BeTrue

            # インデックスのあるフォルダ: 確認でキャンセル、あるインデックスを使って切り替わる（#33）
            setGuiStep $S "［変更…］→ インデックスのあるフォルダ → 確認で［キャンセル］"
            & $changeWorkspace $script:sharedDir
            $confirm = waitGuiWindow $S "インデックスのあるフォルダの確認" -Id "HeadingText" -Text "すでにインデックスがあります"
            clickGuiByName $S $confirm "キャンセル"
            waitGuiWindowClosed $S $confirm "インデックスのあるフォルダの確認"
            & $workspaceText | Should -Be "$($script:nonEmptyDir)\workspace"
            setGuiStep $S "［変更…］→ インデックスのあるフォルダ → 確認で［あるインデックスを使う］"
            & $changeWorkspace $script:sharedDir
            $confirm = waitGuiWindow $S "インデックスのあるフォルダの確認" -Id "HeadingText" -Text "すでにインデックスがあります"
            clickGuiByNameLike $S $confirm "あるインデックスを使う*"
            waitGuiWindowClosed $S $confirm "インデックスのあるフォルダの確認"
            waitGui $S "ワークスペースがインデックスのあるフォルダに変わる" ${guiDefaultTimeout} { (& $workspaceText) -eq $script:sharedDir } | Out-Null
            (readGuiConfig $script:tool).workspaceFolder | Should -Be $script:sharedDir

            closeGui $S
        }
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

# S6 は、既定のワークスペース（%USERPROFILE%\Documents\tebunko_ws。利用者の本物のワークスペース）を使う。手元で流すと利用者のワークスペースに
# 触るため、CI（GITHUB_ACTIONS）だけで流す。手元では、理由を出して飛ばす
Describe "S6 既定のワークスペース（CI だけ）" -Tag Gui {
    BeforeAll {
        $script:runS6 = ($env:GITHUB_ACTIONS -eq "true")
        if ($script:runS6) {
            $script:defaultWork = Join-Path ([Environment]::GetFolderPath("UserProfile")) "Documents\tebunko_ws"
            $script:defaultWorkExisted = Test-Path -LiteralPath $script:defaultWork
            $script:tool = newGuiTool $TestDrive @{ workspaceFolder = "" }
            $script:source = Join-Path $TestDrive "元のフォルダ\営業"
            newGuiSourceFolder $script:source
            $config = readGuiConfig $script:tool
            $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "営業"; path = $script:source; enabled = $true }) -Force
            writeGuiConfig $script:tool.Dir $config
            # 既定のワークスペースに、tebunko のものではないファイルを置く
            [void][IO.Directory]::CreateDirectory($script:defaultWork)
            Set-Content -LiteralPath "$script:defaultWork\ほかのファイル.txt" -Value "ほかのファイル" -Encoding UTF8
            $script:emptyDir = Join-Path $TestDrive "空のフォルダ"
            [void][IO.Directory]::CreateDirectory($script:emptyDir)
        }
    }

    AfterAll {
        # 流したのは CI のランナーだけ。作ったものを消す
        if ($script:runS6 -and !$script:defaultWorkExisted -and (Test-Path -LiteralPath $script:defaultWork)) {
            Remove-Item -LiteralPath $script:defaultWork -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It "起動時の警告・作成の開始の警告・［既定に戻す］が動く" {
        if (!$script:runS6) {
            Set-ItResult -Skipped -Because "既定のワークスペース（利用者の本物のワークスペース）を使うため、CI（GITHUB_ACTIONS）だけで流す"
            return
        }
        $S = startGui $script:tool "S6"
        invokeGuiScene $S {
            # 既定のワークスペースにほかのファイルがあると、起動時に警告が出て、［8 設定］が選ばれる（#5）
            setGuiStep $S "起動時の警告"
            closeGuiMessage $S "空のフォルダではありません" "起動時の警告" | Out-Null
            getGuiSelectedTab $S | Should -Be "SettingsTab"

            # ［インデックス作成を開始］も警告が出て、［8 設定］が選ばれる（#22）
            setGuiStep $S "［インデックス作成を開始］の警告"
            selectGuiTab $S "IndexTab" "IndexingButton"
            clickGui $S $S.Window "IndexingButton" "［インデックス作成を開始］"
            closeGuiMessage $S "空のフォルダではありません" "作成の開始の警告" | Out-Null
            waitGui $S "［8 設定］が選ばれる" ${guiDefaultTimeout} { (getGuiSelectedTab $S) -eq "SettingsTab" } | Out-Null

            # ほかのフォルダに変えてから、ほかのファイルを消して［既定に戻す］（#35）
            setGuiStep $S "空のフォルダに変える"
            clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
            useGuiFolderPicker $S $script:emptyDir
            $confirm = waitGuiWindow $S "ワークスペースを変える確認" -Id "HeadingText" -Text "ワークスペースを変えますか"
            clickGuiByNameLike $S $confirm "ワークスペースを変える*"
            waitGuiWindowClosed $S $confirm "ワークスペースを変える確認"
            waitGui $S "［既定に戻す］が出る" ${guiDefaultTimeout} { $b = findGui $S.Window -Id "ResetWorkspaceButton"; $b -and !$b.Current.IsOffscreen } | Out-Null

            setGuiStep $S "［既定に戻す］（ほかのファイルがあるので警告）"
            clickGui $S $S.Window "ResetWorkspaceButton" "［既定に戻す］"
            closeGuiMessage $S "空のフォルダではありません" "既定に戻す警告" | Out-Null

            setGuiStep $S "ほかのファイルを消して［既定に戻す］→ 確認"
            Remove-Item -LiteralPath "$($script:defaultWork)\ほかのファイル.txt" -Force
            clickGui $S $S.Window "ResetWorkspaceButton" "［既定に戻す］"
            $confirm = waitGuiWindow $S "既定に戻す確認" -Id "HeadingText" -Text "ワークスペースを変えますか"
            clickGuiByName $S $confirm "キャンセル"
            waitGuiWindowClosed $S $confirm "既定に戻す確認"
            clickGui $S $S.Window "ResetWorkspaceButton" "［既定に戻す］"
            $confirm = waitGuiWindow $S "既定に戻す確認" -Id "HeadingText" -Text "ワークスペースを変えますか"
            clickGuiByNameLike $S $confirm "ワークスペースを変える*"
            waitGuiWindowClosed $S $confirm "既定に戻す確認"
            waitGui $S "ワークスペースが既定に戻る" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "WorkspaceText")) -eq $script:defaultWork } | Out-Null

            closeGui $S
        }
    }
}
