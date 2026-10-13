# 画面でのプレビュー（全部の列・横の仮想化・縦の読み足し）の確かめ（共通の関数は gui_helpers.ps1）。
# 2,000 列の Excel の行を選び、一致したセルと見出しが見えること、右端までのスクロール、右へスクロールしたまま下・上へ
# スクロールしても読み足しが効くこと、1 列だけの行（Word）が幅いっぱいに出ることを、UI オートメーションで実機の画面を使って確かめる。
# 列の見出しのドラッグ・範囲を選んだコピーは、マウスの合成をしないため画面では確かめない（tests\tebunko\ui\ の単体テストが確かめる）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "プレビューの全部の列と読み足し" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        $script:columns = 2000
        $script:rows = 300
        # 2,000 列 × 300 行のシート。語は 150 行目の 1,500 列目にだけ置く（最初に見える範囲の外）。セルは「r<行>c<列>」
        $tsvRoot = Join-Path $TestDrive "tsv_営業"
        $lines = New-Object System.Collections.Generic.List[string]
        for ($r = 1; $r -le $script:rows; $r++) {
            $cells = New-Object System.Collections.Generic.List[string]
            for ($c = 1; $c -le $script:columns; $c++) { $cells.Add($(if ($r -eq 150 -and $c -eq 1500) { "横の語" } else { "r${r}c${c}" })) }
            $lines.Add(($cells -join "`t"))
        }
        newTsv "$tsvRoot\営業\広い.xlsx\$(toIndexFileName "表")" $lines.ToArray()
        # 1 列だけの Word のページ
        newTsv "$tsvRoot\営業\議事録.docx\$(toIndexFileName "ページ001")" @("前の行", "狭い語の行", "後ろの行")
        [void](newContentIndexFiles $tsvRoot "$($script:tool.Work)\content_index")
        writeGuiConfig $script:tool.Dir ([ordered]@{ workspaceFolder = $script:tool.Work })
    }

    It "全部の列を持ち、見えている範囲だけを作り、横・縦のスクロールと読み足しが効く" {
        $S = startGui $script:tool "PreviewWide"
        invokeGuiScene $S {
            $cellNames = { @(getGuiPreviewTexts $S.Window | ForEach-Object { $_.Text.Current.Name }) }
            $hasCell = { param ($pattern) @(& $cellNames | Where-Object { $_ -match $pattern }).Count -gt 0 }
            $scrollTo = {
                # 横・縦の位置を割合（0〜100。-1 は動かさない）で指定してスクロールする
                param ($horizontal, $vertical)
                $pattern = (waitGuiPreviewScroll $S).GetCurrentPattern([Windows.Automation.ScrollPattern]::Pattern)
                $pattern.SetScrollPercent($horizontal, $vertical)
            }
            $horizontalPercent = { (waitGuiPreviewScroll $S).GetCurrentPattern([Windows.Automation.ScrollPattern]::Pattern).Current.HorizontalScrollPercent }
            $cellCount = {
                # 作ってあるセルの数と、見えている範囲から決まる上限。上限は（見えている列 + 左右に 3 列ずつ）×（見えている行 + 上下に 10 行ずつ）
                # （前の 3 と 10 は、作る範囲に足す余白の列・行の数）。見えているセルは、スクロールの枠と重なるもの
                $scrollRect = (waitGuiPreviewScroll $S).Current.BoundingRectangle
                $texts = @(getGuiPreviewTexts $S.Window)
                $visible = @($texts | ForEach-Object { $_.Text.Current.BoundingRectangle } | Where-Object {
                        $_.Right -gt $scrollRect.Left -and $_.Left -lt $scrollRect.Right -and $_.Bottom -gt $scrollRect.Top -and $_.Top -lt $scrollRect.Bottom })
                $visibleColumns = @($visible | ForEach-Object { [Math]::Round($_.Left) } | Sort-Object -Unique).Count
                $visibleRows = @($visible | ForEach-Object { [Math]::Round($_.Top) } | Sort-Object -Unique).Count
                @{ Count = $texts.Count; Limit = ($visibleColumns + 6) * ($visibleRows + 20); Columns = $visibleColumns; Rows = $visibleRows }
            }
            $search = {
                param ($word)
                setGuiText $S (findGui $S.Window -Id "WordBox") $word
                waitGuiEnabled $S (findGui $S.Window -Id "SearchButton") "［検索］"
                clickGui $S $S.Window "SearchButton" "［検索］"
                waitGui $S "該当 1 件" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "一致 1 件（*" } | Out-Null
                clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
                $row = waitGui $S "結果の行が出る" ${guiDefaultTimeout} { @(getGuiHitRows (findGui $S.Window -Id "ResultGrid")) | Select-Object -First 1 }
                selectGui $row
            }

            setGuiStep $S "2,000 列の行を選ぶ"
            & $search "横の語"
            $hit = waitGui $S "一致したセルが出る" ${guiDefaultTimeout} { @(getGuiPreviewTexts $S.Window | Where-Object { $_.Text.Current.Name -eq "横の語" }) | Select-Object -First 1 }
            $made = & $cellCount
            $made.Count | Should -BeGreaterThan 0
            $made.Count | Should -BeLessOrEqual $made.Limit -Because "見えている範囲（$($made.Columns) 列 × $($made.Rows) 行）と余白の分だけを作る（全部で $($script:columns * $script:rows) セル）"

            setGuiStep $S "一致したセルが見える範囲にあり、その列の見出しも見える"
            $scrollRect = (waitGuiPreviewScroll $S).Current.BoundingRectangle
            $hitRect = $hit.Text.Current.BoundingRectangle
            $hitRect.Left | Should -BeGreaterOrEqual ($scrollRect.Left - 1)
            $hitRect.Right | Should -BeLessOrEqual ($scrollRect.Right + 1)
            $headerTexts = @(findAllGui (findGui $S.Window -Id "PreviewHeaderScroll") -Type Text)
            $headerTexts.Count | Should -BeGreaterThan 0
            $headerTexts.Count | Should -BeLessThan 100 -Because "見えている列の見出しだけを作る"
            @($headerTexts | Where-Object { $_.Current.BoundingRectangle.Left -ge ($hitRect.Left - 2) -and $_.Current.BoundingRectangle.Left -lt $hitRect.Right }).Count | Should -BeGreaterThan 0 -Because "一致したセルの上に見出しがある"

            setGuiStep $S "右端までスクロールする"
            & $scrollTo 100 -1
            waitGui $S "右端の列（2000 列目）のセルが出る" ${guiDefaultTimeout} { & $hasCell 'c2000$' } | Out-Null
            (& $hasCell '^r\d+c1$') | Should -Be $false -Because "左端の列は作り直されて消える"
            $made = & $cellCount
            $made.Count | Should -BeLessOrEqual $made.Limit -Because "右端へ動かしても、見えている範囲と余白の分だけを作る"

            setGuiStep $S "右端のまま下へスクロールして、最後の行まで読み足す"
            & $scrollTo -1 100
            waitGui $S "最後の行（300 行目）の右端のセル" ${guiDefaultTimeout} {
                if (-not (& $hasCell '^r300c2000$')) { & $scrollTo -1 100 }
                & $hasCell '^r300c2000$'
            } | Out-Null
            (& $horizontalPercent) | Should -BeGreaterThan 95 -Because "下へ読み足しても横の位置は保たれる"

            setGuiStep $S "右端のまま上へスクロールして、先頭の行まで読み足す"
            & $scrollTo -1 0
            waitGui $S "先頭の行（1 行目）の右端のセル" ${guiDefaultTimeout} {
                if (-not (& $hasCell '^r1c2000$')) { & $scrollTo -1 0 }
                & $hasCell '^r1c2000$'
            } | Out-Null
            (& $horizontalPercent) | Should -BeGreaterThan 95 -Because "上へ読み足しても横の位置は保たれる"

            setGuiStep $S "左端へ戻る"
            & $scrollTo 0 -1
            waitGui $S "左端の列（1 列目）のセルが出る" ${guiDefaultTimeout} { & $hasCell '^r1c1$' } | Out-Null

            setGuiStep $S "1 列だけの行（Word）は幅いっぱいに出る"
            & $search "狭い語の行"
            $wide = waitGui $S "Word の行のセルが出る" ${guiDefaultTimeout} { @(getGuiPreviewTexts $S.Window | Where-Object { $_.Text.Current.Name -eq "狭い語の行" }) | Select-Object -First 1 }
            $viewport = (waitGuiPreviewScroll $S).Current.BoundingRectangle
            $wide.Text.Current.BoundingRectangle.Width | Should -BeGreaterThan ($viewport.Width * 0.6) -Because "1 列だけのときは幅いっぱいに広げる"

            closeGui $S
        }
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
