# 画面のスモークテスト S5: ワークスペースの変更、S6: 既定のワークスペース（共通の関数は gui_helpers.ps1）。
# 画面遷移の一覧（docs\design\testing\gui-smoke.md「画面のスモークテスト」）の #30〜#34 を S5 で、#5・#22・#35 を S6 で確かめる。
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
        $script:typedDir = Join-Path $TestDrive "入力する空のフォルダ"
        [void][IO.Directory]::CreateDirectory($script:typedDir)
        [void][IO.Directory]::CreateDirectory($script:emptyDir)
        [void][IO.Directory]::CreateDirectory($script:nonEmptyDir)
        Set-Content -LiteralPath "$script:nonEmptyDir\メモ.txt" -Value "メモ" -Encoding UTF8
        newGuiSampleIndex @{ Work = $script:sharedDir } (Join-Path $TestDrive "共有の元") "共有"
    }

    It "［変更…］のキャンセル・使えないフォルダ・空のフォルダ・空でないフォルダ・インデックスのあるフォルダが動く" {
        $S = startGui $script:tool "S5"
        invokeGuiScene $S {
            $workspaceText = { getGuiValue (findGui $S.Window -Id "WorkspaceBox") }
            $changeWorkspace = {
                param ($path)
                clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
                useGuiFolderPicker $S $path
            }

            setGuiStep $S "［設定］を開く"
            selectGuiTab $S "SettingsTab" "ChangeWorkspaceButton"
            & $workspaceText | Should -Be $script:tool.Work

            # 題と説明は置かず、カード「インデックス設定」の見出しの帯から始まる。行は「ワークスペース」「設定ファイル」だけ（表示設定・最大並列の言葉と、消した説明は出ない）
            $texts = @(getGuiTexts $S.Window)
            $texts | Should -Contain "インデックス設定"
            $texts | Should -Contain "ワークスペース"
            $texts | Should -Contain "設定ファイル"
            $texts | Should -Not -Contain "インデックスの保存先を設定します。"
            # 設定ファイルの場所は、押す欄に見えない読み取り専用の文字（選んでコピーできる TextBox）
            $fileText = findGui $S.Window -Id "SettingsFileText" -Type "Edit"
            $fileText | Should -Not -BeNullOrEmpty
            $fileText.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).Current.IsReadOnly | Should -BeTrue
            @($texts | Where-Object { $_ -eq "設定" }).Count | Should -BeLessOrEqual 1 -Because "題「設定」を置かない（ナビの項目の 1 つだけ）"
            @($texts | Where-Object { $_ -like "*表示設定*" -or $_ -like "*最大並列*" -or $_ -like "*インデックスの保存先*" }) | Should -BeNullOrEmpty

            # ［変更…］→ OS のフォルダ選択で［キャンセル］すると変わらない（#30）
            setGuiStep $S "［変更…］→ フォルダ選択で［キャンセル］"
            & $changeWorkspace $null
            Start-Sleep -Milliseconds 500
            & $workspaceText | Should -Be $script:tool.Work

            # 今のインデックスの中のフォルダは使えない（メッセージボックス）（#34）
            setGuiStep $S "［変更…］→ 使えないフォルダ（今のインデックスの中）"
            & $changeWorkspace "$($script:tool.Work)\content_index\営業"
            closeGuiMessage $S "インデックスのフォルダの中です" "使えないフォルダの警告" | Out-Null
            & $workspaceText | Should -Be $script:tool.Work

            # 空のフォルダ: 確認でキャンセルすると変わらず、実行すると中身が移って切り替わる（#31）
            setGuiStep $S "［変更…］→ 空のフォルダ → 確認で［キャンセル］"
            & $changeWorkspace $script:emptyDir
            answerGuiConfirm $S "ワークスペースを変える確認" "へ移動します" "キャンセル"
            & $workspaceText | Should -Be $script:tool.Work
            setGuiStep $S "［変更…］→ 空のフォルダ → 確認で［ワークスペースを変える］"
            & $changeWorkspace $script:emptyDir
            answerGuiConfirm $S "ワークスペースを変える確認" "へ移動します" "移動する"
            waitGui $S "ワークスペースが空のフォルダに変わる" ${guiDefaultTimeout} { (& $workspaceText) -eq $script:emptyDir } | Out-Null
            Test-Path -LiteralPath "$($script:emptyDir)\content_index\営業" | Should -BeTrue -Because "今のワークスペースの中身が移る"
            (readGuiConfig $script:tool).workspaceFolder | Should -Be $script:emptyDir
            # 終わりの処理に、移した先と元の場所が届いている（空だと、ステータスの場所が抜ける）
            $statusText = { getGuiText (findGui $S.Window -Id "StatusText") }
            waitGui $S "移したステータスに場所が入る" ${guiDefaultTimeout} { (& $statusText) -like "*「$($script:emptyDir)」に変え*移す前の場所：$($script:tool.Work)*" } | Out-Null

            # 空でないフォルダ: 確認でキャンセル、中にフォルダを作って切り替わる（#32）
            setGuiStep $S "［変更…］→ 空でないフォルダ → 確認で［キャンセル］"
            & $changeWorkspace $script:nonEmptyDir
            answerGuiConfirm $S "空でないフォルダの確認" "空ではありません" "キャンセル"
            & $workspaceText | Should -Be $script:emptyDir
            setGuiStep $S "［変更…］→ 空でないフォルダ → 中に「workspace」フォルダを作る"
            & $changeWorkspace $script:nonEmptyDir
            answerGuiConfirm $S "空でないフォルダの確認" "空ではありません" "中に*" -Like
            waitGui $S "ワークスペースが中の workspace に変わる" ${guiDefaultTimeout} { (& $workspaceText) -eq "$($script:nonEmptyDir)\workspace" } | Out-Null
            Test-Path -LiteralPath "$($script:nonEmptyDir)\workspace\content_index\営業" | Should -BeTrue

            # インデックスのあるフォルダ: 確認でキャンセル、あるインデックスを使って切り替わる（#33）
            setGuiStep $S "［変更…］→ インデックスのあるフォルダ → 確認で［キャンセル］"
            & $changeWorkspace $script:sharedDir
            answerGuiConfirm $S "インデックスのあるフォルダの確認" "すでにインデックスがあります" "キャンセル"
            & $workspaceText | Should -Be "$($script:nonEmptyDir)\workspace"
            setGuiStep $S "［変更…］→ インデックスのあるフォルダ → 確認で［あるインデックスを使う］"
            & $changeWorkspace $script:sharedDir
            answerGuiConfirm $S "インデックスのあるフォルダの確認" "すでにインデックスがあります" "そのフォルダのインデックスを使う"
            waitGui $S "ワークスペースがインデックスのあるフォルダに変わる" ${guiDefaultTimeout} { (& $workspaceText) -eq $script:sharedDir } | Out-Null
            (readGuiConfig $script:tool).workspaceFolder | Should -Be $script:sharedDir
            waitGui $S "あるインデックスを使うステータスに場所が入る" ${guiDefaultTimeout} { (& $statusText) -like "*「$($script:sharedDir)」に変え、そこにあるインデックスを使います*" } | Out-Null

            # 欄に直接入力して変える（欄から出たときに確定する。確かめ・確認は［変更…］と同じ）
            $typeWorkspace = {
                param ($text)
                $box = findGui $S.Window -Id "WorkspaceBox"
                $box.SetFocus()
                setGuiText $S $box $text
                (findGui $S.Window -Id "ChangeWorkspaceButton").SetFocus()
            }
            $inputError = { getGuiText (findGui $S.Window -Id "WorkspaceInputError") }
            setGuiStep $S "欄に入力: 無いフォルダ → 警告が出て変わらず、欄は今の場所に戻る"
            & $typeWorkspace (Join-Path $TestDrive "無いフォルダ")
            closeGuiMessage $S "ファイルを作れません" "無いフォルダの警告" | Out-Null
            & $workspaceText | Should -Be $script:sharedDir
            setGuiStep $S "欄に入力: 相対パス → 欄の下に理由が出る"
            & $typeWorkspace "相対\パス"
            waitGui $S "理由が出る（相対）" ${guiDefaultTimeout} { (& $inputError) -like "*ドライブ文字*" } | Out-Null
            & $workspaceText | Should -Be $script:sharedDir
            (readGuiConfig $script:tool).workspaceFolder | Should -Be $script:sharedDir
            setGuiStep $S "欄に入力: ネットワークのパス → 入力では受けず、［変更…］へ案内する（共有には触れない）"
            & $typeWorkspace "\\server\share\ws"
            waitGui $S "理由が出る（ネットワーク）" ${guiDefaultTimeout} { (& $inputError) -like "*［変更…］*" } | Out-Null
            & $workspaceText | Should -Be $script:sharedDir
            setGuiStep $S "欄に入力して Enter: 不備のとき、欄の文字は残り、理由が出る"
            $box = findGui $S.Window -Id "WorkspaceBox"
            $box.SetFocus()
            setGuiText $S $box "相対\パス"
            setGuiStep $S "キーを送る"
            pressGuiKey $S.Window 0x0D
            waitGui $S "理由が出る（Enter）" ${guiDefaultTimeout} { (& $inputError) -like "*ドライブ文字*" } | Out-Null
            & $workspaceText | Should -Be "相対\パス"
            setGuiStep $S "Esc: 欄が今の場所に戻り、理由が消える"
            pressGuiKey $S.Window 0x1B
            waitGui $S "欄が戻り、理由が消える（Esc）" ${guiDefaultTimeout} { (& $workspaceText) -eq $script:sharedDir -and (& $inputError) -eq "" } | Out-Null
            (readGuiConfig $script:tool).workspaceFolder | Should -Be $script:sharedDir
            setGuiStep $S "入力の途中で［変更…］ → フォルダ選択で［キャンセル］: 入力した文字と理由が消え、欄は今の場所に戻る"
            $box.SetFocus()
            setGuiText $S $box "相対\パス"
            pressGuiKey $S.Window 0x0D
            waitGui $S "理由が出る（［変更…］の前）" ${guiDefaultTimeout} { (& $inputError) -like "*ドライブ文字*" } | Out-Null
            & $changeWorkspace $null
            waitGui $S "欄が戻り、理由が消える（［変更…］のキャンセル）" ${guiDefaultTimeout} { (& $workspaceText) -eq $script:sharedDir -and (& $inputError) -eq "" } | Out-Null
            (readGuiConfig $script:tool).workspaceFolder | Should -Be $script:sharedDir
            setGuiStep $S "欄に入力: 今のインデックスの中 → 警告が出て変わらない"
            & $typeWorkspace "$($script:sharedDir)\content_index\共有"
            closeGuiMessage $S "インデックスのフォルダの中です" "使えないフォルダの警告" | Out-Null
            & $workspaceText | Should -Be $script:sharedDir
            setGuiStep $S "欄に入力: 空のフォルダ → 確認で［移動する］ → 切り替わる"
            & $typeWorkspace $script:typedDir
            answerGuiConfirm $S "ワークスペースを変える確認" "へ移動します" "移動する"
            waitGui $S "ワークスペースが入力したフォルダに変わる" ${guiDefaultTimeout} { (& $workspaceText) -eq $script:typedDir } | Out-Null
            (readGuiConfig $script:tool).workspaceFolder | Should -Be $script:typedDir
            (& $inputError) | Should -Be ""

            closeGui $S
        }
    }

    It "移動が失敗したら、警告にフォルダ名が入り、元のワークスペースのまま使える" {
        # 移す先の content_index の中のファイルをこちらで開いたままにして、消す処理を失敗させる（TestDrive の中だけ）
        $failTool = newGuiTool (Join-Path $TestDrive "失敗")
        newGuiSampleIndex $failTool $TestDrive "失敗用"
        $failDir = Join-Path $TestDrive "消せないインデックスのある先"
        [void][IO.Directory]::CreateDirectory("$failDir\content_index\あ")
        $held = [IO.File]::Open("$failDir\content_index\あ\a.tsv", [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        try {
            $S = startGui $failTool "S5f"
            invokeGuiScene $S {
                selectGuiTab $S "SettingsTab" "ChangeWorkspaceButton"
                clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
                useGuiFolderPicker $S $failDir
                answerGuiConfirm $S "インデックスのあるフォルダの確認" "すでにインデックスがあります" "今のインデックスを移動する*" -Like
                $warning = waitGuiWindow $S "移動の失敗の警告" -Text "削除できませんでした"
                $text = @(getGuiTexts $warning) -join " "
                $text | Should -BeLike "*「$failDir」のインデックスを削除できませんでした*"
                closeGuiMessage $S "削除できませんでした" "移動の失敗の警告" | Out-Null
                (getGuiValue (findGui $S.Window -Id "WorkspaceBox")) | Should -Be $failTool.Work
                (readGuiConfig $failTool).workspaceFolder | Should -Be $failTool.Work
                Test-Path -LiteralPath "$($failTool.Work)\content_index\失敗用" | Should -BeTrue
                closeGui $S
            }
        } finally {
            $held.Dispose()
        }
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

# S6 は、既定のワークスペースを使う場面。既定は環境変数 TEBUNKO_DEFAULT_WORKSPACE で $TestDrive の中に差し替えて起動する
# （利用者の既定のワークスペースには触れない。手元でも CI でも流す）
Describe "S6 既定のワークスペース" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive @{ workspaceFolder = "" }
        $script:defaultWork = $script:tool.DefaultWorkspace
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

    It "起動時の警告・作成の開始の警告・［既定に戻す］が動く" {
        $S = startGui $script:tool "S6"
        invokeGuiScene $S {
            # 既定のワークスペースにほかのファイルがあると、起動時に警告が出て、［設定］が選ばれる（#5）
            setGuiStep $S "起動時の警告"
            closeGuiMessage $S "空のフォルダではありません" "起動時の警告" | Out-Null
            getGuiSelectedTab $S | Should -Be "SettingsTab"

            # ［すべて更新］も警告が出て、［設定］が選ばれる（#22）
            setGuiStep $S "［すべて更新］の警告"
            selectGuiTab $S "IndexTab" "IndexingButton"
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            closeGuiMessage $S "空のフォルダではありません" "作成の開始の警告" | Out-Null
            waitGui $S "［設定］が選ばれる" ${guiDefaultTimeout} { (getGuiSelectedTab $S) -eq "SettingsTab" } | Out-Null

            # ほかのフォルダに変えてから、ほかのファイルを消して［既定に戻す］（#35）
            setGuiStep $S "空のフォルダに変える"
            clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
            useGuiFolderPicker $S $script:emptyDir
            answerGuiConfirm $S "ワークスペースを変える確認" "へ移動します" "移動する"
            waitGui $S "［既定に戻す］が出る" ${guiDefaultTimeout} { $b = findGui $S.Window -Id "ResetWorkspaceButton"; $b -and !$b.Current.IsOffscreen } | Out-Null

            setGuiStep $S "入力の途中で［既定に戻す］（ほかのファイルがあるので警告）: 入力した文字と理由が消え、欄は今の場所に戻る"
            $box = findGui $S.Window -Id "WorkspaceBox"
            $box.SetFocus()
            setGuiText $S $box "相対\パス"
            pressGuiKey $S.Window 0x0D
            waitGui $S "理由が出る（［既定に戻す］の前）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "WorkspaceInputError")) -like "*ドライブ文字*" } | Out-Null
            clickGui $S $S.Window "ResetWorkspaceButton" "［既定に戻す］"
            closeGuiMessage $S "空のフォルダではありません" "既定に戻す警告" | Out-Null
            waitGui $S "欄が戻り、理由が消える（［既定に戻す］の警告）" ${guiDefaultTimeout} { (getGuiValue (findGui $S.Window -Id "WorkspaceBox")) -eq $script:emptyDir -and (getGuiText (findGui $S.Window -Id "WorkspaceInputError")) -eq "" } | Out-Null

            setGuiStep $S "ほかのファイルを消して［既定に戻す］→ 確認"
            Remove-Item -LiteralPath "$($script:defaultWork)\ほかのファイル.txt" -Force
            clickGui $S $S.Window "ResetWorkspaceButton" "［既定に戻す］"
            answerGuiConfirm $S "既定に戻す確認" "へ移動します" "キャンセル"
            clickGui $S $S.Window "ResetWorkspaceButton" "［既定に戻す］"
            answerGuiConfirm $S "既定に戻す確認" "へ移動します" "戻す"
            waitGui $S "ワークスペースが既定に戻る" ${guiDefaultTimeout} { (getGuiValue (findGui $S.Window -Id "WorkspaceBox")) -eq $script:defaultWork } | Out-Null

            closeGui $S
        }
    }

    It "利用者の環境（利用者の既定のワークスペースを含む）に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
