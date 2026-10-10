# 画面のテスト S10: 設定から外れたインデックスの削除予定を、確認のダイアログで見せてから消す（共通の関数は gui_helpers.ps1）。
# 更新するファイルが 0 件で削除予定だけのときも、［更新を開始］と［キャンセル］が残り、キャンセルすれば消えないことを確かめる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S10 設定から外れたインデックスの削除予定" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:sourceA = Join-Path $TestDrive "元のフォルダ\営業"
        $script:sourceB = Join-Path $TestDrive "元のフォルダ\資料"
        newGuiSourceFolder $script:sourceA
        newGuiSourceFolder $script:sourceB
        $script:tool = newGuiTool $TestDrive @{ targetFolders = @(
                @{ name = "営業"; path = $script:sourceA; enabled = $true },
                @{ name = "資料"; path = $script:sourceB; enabled = $true }) }
    }

    It "2 つ作ってから 1 つを設定から外すと、確認に削除予定が出て、キャンセルでは消えず、［更新を開始］で消える" {
        $S = startGui $script:tool "S10a"
        invokeGuiScene $S {
            setGuiStep $S "2 つのインデックスを作る"
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $confirm "StartButton" "確認の［更新を開始］"
            waitGuiWindowClosed $S $confirm "取り込みの確認"
            waitGui $S "取り込みの完了" ${guiIndexTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*" -and
                    (findGui $S.Window -Id "IndexingButton").Current.IsEnabled
            } | Out-Null
            closeGui $S
        }
        Test-Path -LiteralPath "$($script:tool.Work)\content_index\営業" | Should -BeTrue
        Test-Path -LiteralPath "$($script:tool.Work)\content_index\資料" | Should -BeTrue

        # 資料だけを設定に残す（営業は設定から外れる）
        $config = [ordered]@{ workspaceFolder = $script:tool.Work; targetFolders = @(@{ name = "資料"; path = $script:sourceB; enabled = $true }) }
        writeGuiConfig $script:tool.Dir $config

        $S = startGui $script:tool "S10b"
        invokeGuiScene $S {
            selectGuiTab $S "IndexTab" "IndexingButton"
            setGuiStep $S "［すべて更新］→ 削除予定だけの確認（更新 0 件）"
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            $texts = (@(getGuiTexts $confirm)) -join "`n"
            $texts | Should -BeLike "*削除予定*"
            $texts | Should -BeLike "*設定に無いインデックス 1 件を削除します。*"
            (findGui $confirm -Id "StartButton").Current.Name | Should -Be "更新を開始"
            (findGui $confirm -Id "CancelButton").Current.IsOffscreen | Should -BeFalse

            setGuiStep $S "［キャンセル］では消えない"
            clickGui $S $confirm "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $confirm "取り込みの確認"
            waitGui $S "取りやめて［すべて更新］に戻る" ${guiDefaultTimeout} { (findGui $S.Window -Id "IndexingButton").Current.IsEnabled } | Out-Null
            Test-Path -LiteralPath "$($script:tool.Work)\content_index\営業" | Should -BeTrue

            setGuiStep $S "［更新を開始］で消える"
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $confirm "StartButton" "確認の［更新を開始］"
            waitGuiWindowClosed $S $confirm "取り込みの確認"
            waitGui $S "営業のインデックスが消える" ${guiIndexTimeout} { !(Test-Path -LiteralPath "$($script:tool.Work)\content_index\営業") } | Out-Null
            Test-Path -LiteralPath "$($script:tool.Work)\content_index\資料" | Should -BeTrue
            waitGui $S "更新が終わる" ${guiIndexTimeout} { (findGui $S.Window -Id "IndexingButton").Current.IsEnabled } | Out-Null
            closeGui $S
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
