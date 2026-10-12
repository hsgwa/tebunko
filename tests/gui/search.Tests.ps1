# 画面のスモークテスト S4: 検索の遷移（共通の関数は gui_helpers.ps1）。
# 画面遷移の一覧（docs\design\testing\gui-smoke.md「画面のスモークテスト」）の #24・#25・#27・#29 を確かめる。
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

            # 正規表現で不正な式を入れると、［正規表現］の下に吹き出しが出て、入力欄の枠が赤くなり、［検索］は押せなくなる（検索バーの高さは変わらない）。式を直すと消える（#24）
            setGuiStep $S "正規表現で不正な式"
            $regex = waitGuiById $S $S.Window "RegexCheck"
            toggleGui $regex
            waitGui $S "「正規表現を使う」が付く" ${guiDefaultTimeout} { (getGuiToggleState (findGui $S.Window -Id "RegexCheck")) -eq "On" } | Out-Null
            setGuiText $S (findGui $S.Window -Id "WordBox") "("
            waitGui $S "吹き出し（RegexBalloon）が出る" ${guiDefaultTimeout} {
                $text = findGui $S.Window -Id "RegexBalloonText"
                !$text.Current.IsOffscreen -and (getGuiText $text) -eq "正規表現が正しくありません"
            } | Out-Null
            $regexRect = (findGui $S.Window -Id "RegexBalloonText").Current.BoundingRectangle
            $regexCheckRect = (findGui $S.Window -Id "RegexCheck").Current.BoundingRectangle
            ($regexRect.Top - $regexCheckRect.Bottom) | Should -BeGreaterThan 0 -Because "吹き出しは［正規表現］の下に出す"
            ($regexRect.Top - $regexCheckRect.Bottom) | Should -BeLessThan 15 -Because "吹き出しは［正規表現］のすぐ下に出す"
            $regexRect.Right | Should -BeLessOrEqual $S.Window.Current.BoundingRectangle.Right -Because "吹き出しは窓の右端からはみ出さない"
            (findGui $S.Window -Id "SearchButton").Current.IsEnabled | Should -BeFalse -Because "正規表現が正しくないと検索を始めない"
            $resultTop = (findGui $S.Window -Id "FilterBox").Current.BoundingRectangle.Top
            $resultTop | Should -BeLessThan 100000 -Because "絞り込みの欄が見えている（高さの比べの基準）"
            setGuiStep $S "式を直す"
            setGuiText $S (findGui $S.Window -Id "WordBox") "単価"
            waitGui $S "吹き出しが消え［検索］が押せる" ${guiDefaultTimeout} {
                (findGui $S.Window -Id "RegexBalloonText").Current.IsOffscreen -and (findGui $S.Window -Id "SearchButton").Current.IsEnabled
            } | Out-Null
            (findGui $S.Window -Id "FilterBox").Current.BoundingRectangle.Top | Should -Be $resultTop -Because "吹き出しを出しても検索バーの高さは変わらない"
            toggleGui (findGui $S.Window -Id "RegexCheck")

            # 検索対象のツリーで［すべて解除］すると検索できず、［すべて選択］で戻る（#25）
            setGuiStep $S "検索対象の［すべて解除］"
            clickGui $S $S.Window "UncheckAllIndexButton" "［解除］"
            waitGui $S "検索対象が「なし」になり［検索］が押せない" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "TargetCountText")) -like "検索対象 0 / *" -and !(findGui $S.Window -Id "SearchButton").Current.IsEnabled
            } | Out-Null
            setGuiStep $S "検索対象の［すべて選択］"
            clickGui $S $S.Window "CheckAllIndexButton" "［すべて］"
            waitGui $S "検索対象が戻り［検索］が押せる" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "TargetCountText")) -match "^検索対象 (\d+) / \1$" -and (findGui $S.Window -Id "SearchButton").Current.IsEnabled
            } | Out-Null

            # 検索して、［すべて開く］［すべて折りたたむ］・絞り込み（#27）
            setGuiStep $S "検索"
            & $search "単価"
            waitGui $S "該当 2 件" ${guiDefaultTimeout} { (& $summary) -like "一致 2 件（*" } | Out-Null
            setGuiStep $S "［すべて開く］"
            clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
            waitGui $S "結果の行が 2 件出る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null
            setGuiStep $S "［すべて折りたたむ］"
            clickGui $S $S.Window "CollapseAllButton" "［すべて折りたたむ］"
            waitGui $S "結果の行が隠れる" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 0 } | Out-Null
            setGuiStep $S "結果の絞り込み"
            clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
            waitGui $S "結果の行が 2 件出る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null
            # 全部開いたら、表の行が全部見える（最後の行の下端が表の下端の内側）。下の余白は行 1 つぶん未満
            $gridRect = { (findGui $S.Window -Id "ResultGrid").Current.BoundingRectangle }
            $lastRowBottom = { @(& $hitRows)[-1].Current.BoundingRectangle.Bottom }
            waitGui $S "開いた 2 件が表に全部入る" ${guiDefaultTimeout} { (& $lastRowBottom) -le (& $gridRect).Bottom } | Out-Null
            ((& $gridRect).Bottom - (& $lastRowBottom)) | Should -BeLessThan 40 -Because "表の下に余白が長く残らない"
            $previewTwo = (findGui $S.Window -Id "PreviewSplitter").Current.BoundingRectangle.Top
            setGuiText $S (findGui $S.Window -Id "FilterBox") "議事録"
            waitGui $S "絞り込んだ件数（結果の行が 1 件）" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 1 } | Out-Null
            # 結果が少ないときは、結果の一覧が内容の高さまでで、余白が長く残らない（2 件のときより高くならない）
            waitGui $S "結果の一覧が内容の高さに縮み、最後の行が表に入ったまま、下の余白が行 1 つぶん未満になる" ${guiDefaultTimeout} {
                $gap = (& $gridRect).Bottom - (& $lastRowBottom)
                $gap -ge 0 -and $gap -lt 40
            } | Out-Null
            (findGui $S.Window -Id "PreviewSplitter").Current.BoundingRectangle.Top | Should -BeLessThan $previewTwo -Because "表が縮んだぶん、プレビューが上に広がる"
            setGuiText $S (findGui $S.Window -Id "FilterBox") ""
            waitGui $S "絞り込みを消すと 2 件に戻る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null

            # 種類のチップを外すと、その種類のファイルが検索から外れる。戻すと元に戻る（全部外すこともできるが、そのままでは検索を始めない）
            setGuiStep $S "種類のチップ［Word］を外して検索"
            toggleGui (findGui $S.Window -Id "KindChipWord")
            waitGui $S "チップ［Word］が外れる" ${guiDefaultTimeout} { (getGuiToggleState (findGui $S.Window -Id "KindChipWord")) -eq "Off" } | Out-Null
            & $search "単価"
            waitGui $S "Word を除いた該当 1 件" ${guiDefaultTimeout} { (& $summary) -like "一致 1 件（*" } | Out-Null
            setGuiStep $S "チップ［Word］を戻して検索"
            toggleGui (findGui $S.Window -Id "KindChipWord")
            waitGui $S "チップ［Word］が付く" ${guiDefaultTimeout} { (getGuiToggleState (findGui $S.Window -Id "KindChipWord")) -eq "On" } | Out-Null
            & $search "単価"
            waitGui $S "該当 2 件に戻る" ${guiDefaultTimeout} { (& $summary) -like "一致 2 件（*" } | Out-Null

            # 種類のチップを全部外せる。1 つも選んでいないと検索を始めず、種類のチップの下の吹き出しで知らせる。1 つ選べば検索できる
            $chipIds = @("KindChipExcel", "KindChipWord", "KindChipPowerPoint", "KindChipText", "KindChipPdf")
            setGuiStep $S "種類のチップを全部外す"
            foreach ($chipId in $chipIds) { toggleGui (findGui $S.Window -Id $chipId) }
            waitGui $S "チップが全部外れる" ${guiDefaultTimeout} {
                @($chipIds | Where-Object { (getGuiToggleState (findGui $S.Window -Id $_)) -eq "On" }).Count -eq 0
            } | Out-Null
            & $search "単価"
            waitGui $S "種類を選ぶ吹き出し" ${guiDefaultTimeout} {
                $text = findGui $S.Window -Id "SearchKindBalloonText"
                !$text.Current.IsOffscreen -and (getGuiText $text) -eq "種類を 1 つ以上選んでください"
            } | Out-Null
            $balloonRect = (findGui $S.Window -Id "SearchKindBalloonText").Current.BoundingRectangle
            $chipRect = (findGui $S.Window -Id "KindChipExcel").Current.BoundingRectangle
            # 吹き出しの文字は、チップの左端から枠・余白・アイコンの分（20〜40）だけ右、チップの下端から 0〜15 下
            ($balloonRect.Left - $chipRect.Left) | Should -BeGreaterThan 20 -Because "左端はチップの左端にそろう"
            ($balloonRect.Left - $chipRect.Left) | Should -BeLessThan 40 -Because "左端はチップの左端にそろう"
            ($balloonRect.Top - $chipRect.Bottom) | Should -BeGreaterThan 0 -Because "吹き出しは種類のチップの下に出す"
            ($balloonRect.Top - $chipRect.Bottom) | Should -BeLessThan 15 -Because "吹き出しはチップのすぐ下に出す"
            (findGui $S.Window -Id "FilterBox").Current.BoundingRectangle.Top | Should -Be $resultTop -Because "吹き出しを出しても検索バーの高さは変わらない"
            (findGui $S.Window -Id "SearchButton").Current.IsEnabled | Should -BeTrue -Because "種類が無くても［検索］は押せる"
            (& $summary) | Should -BeLike "一致 2 件（*" -Because "検索を始めないため、前の結果のまま"
            setGuiStep $S "チップ［Excel］だけ選んで検索"
            toggleGui (findGui $S.Window -Id "KindChipExcel")
            & $search "単価"
            waitGui $S "種類を選んだので吹き出しが消える" ${guiDefaultTimeout} { (findGui $S.Window -Id "SearchKindBalloonText").Current.IsOffscreen } | Out-Null
            waitGui $S "Excel だけの該当 1 件" ${guiDefaultTimeout} { (& $summary) -like "一致 1 件（*" } | Out-Null
            setGuiStep $S "チップを全部戻して検索"
            foreach ($chipId in @("KindChipWord", "KindChipPowerPoint", "KindChipText", "KindChipPdf")) { toggleGui (findGui $S.Window -Id $chipId) }
            & $search "単価"
            waitGui $S "該当 2 件に戻る" ${guiDefaultTimeout} { (& $summary) -like "一致 2 件（*" } | Out-Null
            (findGui $S.Window -Id "NoHitText").Current.IsOffscreen | Should -BeTrue -Because "結果があるときは「一致する結果はありません」を出さない"
            setGuiStep $S "0 件の検索"
            & $search "存在しない語zzzqqq"
            waitGui $S "0 件の文が出る" ${guiDefaultTimeout} {
                $noHit = findGui $S.Window -Id "NoHitText"
                !$noHit.Current.IsOffscreen -and (getGuiText $noHit) -eq "一致する結果はありません"
            } | Out-Null
            & $search "単価"
            waitGui $S "0 件の文が消え、該当 2 件に戻る" ${guiDefaultTimeout} { (& $summary) -like "一致 2 件（*" -and (findGui $S.Window -Id "NoHitText").Current.IsOffscreen } | Out-Null
            clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
            waitGui $S "結果の行が 2 件出る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null

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
            # 確認に戻って「見つからない確認」が再び出ないことを確かめる（結果の見出し行のツールヒントが一瞬出ることがあるため、
            # 窓の数そのものではなく、確認ダイアログが無いことを見る）
            Start-Sleep -Milliseconds 500
            @(getGuiOtherWindows $S | Where-Object {
                $heading = findGui $_ -Id "HeadingText"
                $heading -and $heading.Current.Name -like "*が見つかりません*"
            }).Count | Should -Be 0 -Because "フォルダ選択をキャンセルすると、確認に戻らずに終わる"

            closeGui $S
        }
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
