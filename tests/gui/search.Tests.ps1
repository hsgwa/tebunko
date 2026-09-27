# 画面のスモークテスト S4: 検索の遷移（共通の関数は gui_helpers.ps1）。
# 画面遷移の一覧（docs\design\testing\index.md「画面のスモークテスト」）の #24・#25・#27・#29 を確かめる。
# #28（右クリックのメニュー・プレビューのメニュー）は対象外（UI オートメーションではメニューを開けない）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S4 検索の遷移" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        # 元のファイル（見積.xlsx・議事録.docx）は実在しない
        newGuiSampleIndex $script:tool $TestDrive
    }

    It "不正な正規表現・検索対象の選択・結果の開閉と絞り込み・見つからないファイルの確認が動く" {
        $S = startGui $script:tool "S4"
        invokeGuiScene $S {
            $summary = { getGuiText (findGui $S.Window -Id "SummaryText") }
            $hitRows = { getGuiHitRows (findGui $S.Window -Id "ResultGrid") }
            $search = {
                param ($word)
                setGuiText $S (findGui $S.Window -Id "WordBox") $word
                waitGuiEnabled $S (findGui $S.Window -Id "SearchButton") "［検索］"
                clickGui $S $S.Window "SearchButton" "［検索］"
            }

            getGuiSelectedTab $S | Should -Be "SearchTab"

            # 正規表現で不正な式を入れると、注意が出る（文字どおり検索にする案内。［検索］は押せたまま）。式を直すと消える（#24）
            setGuiStep $S "正規表現で不正な式"
            $regex = waitGuiById $S $S.Window "RegexCheck"
            toggleGui $regex
            waitGui $S "「正規表現を使う」が付く" ${guiDefaultTimeout} { (getGuiToggleState (findGui $S.Window -Id "RegexCheck")) -eq "On" } | Out-Null
            setGuiText $S (findGui $S.Window -Id "WordBox") "("
            waitGui $S "注意（WordNotice）が出る" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "WordNotice")) -like "*文字どおり検索*" } | Out-Null
            setGuiStep $S "式を直す"
            setGuiText $S (findGui $S.Window -Id "WordBox") "単価"
            waitGui $S "注意が消える" ${guiDefaultTimeout} { (findGui $S.Window -Id "WordNotice").Current.IsOffscreen } | Out-Null
            toggleGui (findGui $S.Window -Id "RegexCheck")

            # 検索対象のツリーで［すべて解除］すると検索できず、［すべて選択］で戻る（#25）
            setGuiStep $S "検索対象の［すべて解除］"
            clickGui $S $S.Window "UncheckAllIndexButton" "［すべて解除］"
            waitGui $S "検索対象が「なし」になり［検索］が押せない" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "SearchTargetText")) -like "検索対象：なし*" -and !(findGui $S.Window -Id "SearchButton").Current.IsEnabled
            } | Out-Null
            setGuiStep $S "検索対象の［すべて選択］"
            clickGui $S $S.Window "CheckAllIndexButton" "［すべて選択］"
            waitGui $S "検索対象が戻り［検索］が押せる" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "SearchTargetText")) -like "検索対象：すべて*" -and (findGui $S.Window -Id "SearchButton").Current.IsEnabled
            } | Out-Null

            # 検索して、［すべて展開］［すべて折りたたむ］・絞り込み（#27）
            setGuiStep $S "検索"
            & $search "単価"
            waitGui $S "該当 2 件" ${guiDefaultTimeout} { (& $summary) -like "該当 2 件*" } | Out-Null
            setGuiStep $S "［すべて展開］"
            clickGui $S $S.Window "ExpandAllButton" "［すべて展開］"
            waitGui $S "結果の行が 2 件出る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null
            setGuiStep $S "［すべて折りたたむ］"
            clickGui $S $S.Window "CollapseAllButton" "［すべて折りたたむ］"
            waitGui $S "結果の行が隠れる" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 0 } | Out-Null
            setGuiStep $S "結果の絞り込み"
            clickGui $S $S.Window "ExpandAllButton" "［すべて展開］"
            waitGui $S "結果の行が 2 件出る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null
            setGuiText $S (findGui $S.Window -Id "FilterBox") "議事録"
            waitGui $S "絞り込んだ件数（結果の行が 1 件）" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 1 } | Out-Null
            setGuiText $S (findGui $S.Window -Id "FilterBox") ""
            waitGui $S "絞り込みを消すと 2 件に戻る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null

            # 元のファイルが無い行で［… で開く］を押すと、確認が出る。［キャンセル］／［フォルダを選ぶ］→ フォルダ選択［キャンセル］（#29）
            setGuiStep $S "結果の行を選ぶ"
            selectGui (@(& $hitRows) | Select-Object -Last 1)
            waitGuiEnabled $S (findGui $S.Window -Id "OpenButton") "［開く］"
            setGuiStep $S "元のファイルが無い行の［開く］→［キャンセル］"
            clickGui $S $S.Window "OpenButton" "［開く］"
            answerGuiConfirm $S "見つからない確認" "が見つかりません" "キャンセル"
            setGuiStep $S "元のファイルが無い行の［開く］→［フォルダを選ぶ］→ フォルダ選択［キャンセル］"
            clickGui $S $S.Window "OpenButton" "［開く］"
            $confirm = waitGuiWindow $S "見つからない確認" -Id "HeadingText" -Text "が見つかりません"
            clickGuiByName $S $confirm "フォルダを選ぶ"
            useGuiFolderPicker $S
            waitGuiWindowClosed $S $confirm "見つからない確認"
            # フォルダ選択のキャンセルの直後は、後片付けが少し遅れて窓の一覧に残ることがあるため、時間で待つ
            waitGui $S "確認に戻らず、ほかの窓が残らない" ${guiDefaultTimeout} { @(getGuiOtherWindows $S).Count -eq 0 } | Out-Null

            closeGui $S
        }
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
