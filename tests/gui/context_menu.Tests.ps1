# 画面のスモークテスト: 検索結果の右クリックメニュー（共通の関数は gui_helpers.ps1）。
# マウスの右クリックは合成できないため、行を選んでフォーカスを移し、メニューキー（VK_APPS）のメッセージで開く。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "検索結果の右クリックメニュー" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        newGuiSampleIndex $script:tool $TestDrive
        # 既定の開き方を「読み取り専用」にして起動する
        $config = [ordered]@{ workspaceFolder = $script:tool.Work; openMode = "readOnly" }
        writeGuiConfig $script:tool.Dir $config
    }

    It "行のメニューは先頭に既定の開き方を出し、見出しのメニューは読み取り専用で開くから始まる" {
        $S = startGui $script:tool "ContextMenu1"
        $script:rowMenu = $null
        $script:groupMenu = $null
        invokeGuiScene $S {
            $menuNames = {
                # メニューキーで開いて、出ている項目の名前を取る
                param ($item, [string]$focusId = "ResultGrid")
                if ($item) { selectGui $item }
                (findGui $S.Window -Id $focusId).SetFocus()
                pressGuiKey $S.Window 0x5D
                $names = waitGui $S "メニューが開く" ${guiDefaultTimeout} {
                    $found = @(getGuiOtherWindows $S | ForEach-Object { findAllGui $_ -Type MenuItem } | ForEach-Object { $_.Current.Name })
                    if ($found.Count -gt 0) { , $found }
                }
                # メニューを閉じる（Esc）。閉じないと、画面を閉じる手順が進まない
                foreach ($window in @(getGuiOtherWindows $S)) {
                    pressGuiKey $window 0x1B
                }
                waitGui $S "メニューが閉じる" ${guiDefaultTimeout} { @(getGuiOtherWindows $S).Count -eq 0 } | Out-Null
                return $names
            }
            setGuiText $S (findGui $S.Window -Id "WordBox") "単価"
            waitGuiEnabled $S (findGui $S.Window -Id "SearchButton") "［検索］"
            clickGui $S $S.Window "SearchButton" "［検索］"
            waitGui $S "該当 2 件" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "2 件（*" } | Out-Null
            clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
            waitGui $S "結果の行が出る" ${guiDefaultTimeout} { @(getGuiHitRows (findGui $S.Window -Id "ResultGrid")).Count -eq 2 } | Out-Null

            $grid = findGui $S.Window -Id "ResultGrid"
            $script:rowMenu = & $menuNames (@(getGuiHitRows $grid))[0]
            # 見出しの行（ヒットの行ではない行）
            $header = @(getGuiGridRows $grid | Where-Object { @(findAllGui $_ -Type Text).Count -eq 0 })[0]
            $script:groupMenu = & $menuNames $header
            # 行を選び直すとプレビューが出る。プレビューの表の上でもメニューが開く
            selectGui (@(getGuiHitRows $grid))[0]
            $script:previewMenu = & $menuNames $null "PreviewScroll"
            closeGui $S
        }
        ($script:rowMenu -join "/") | Should -Be "読み取り専用で開く/開く/新規で開く/フォルダを開く/選んだ行をコピー/ファイルのパスをコピー"
        ($script:groupMenu -join "/") | Should -Be "読み取り専用で開く/フォルダを開く/ファイルのパスをコピー/この結果を折りたたむ"
        ($script:previewMenu -join "/") | Should -Be "元のファイルのこの場所を開く/選んだセルをコピー/この行をコピー"
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
