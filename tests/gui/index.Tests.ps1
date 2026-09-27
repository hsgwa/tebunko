# 画面のスモークテスト S2: インデックスの管理と作成、S3: 作成中の操作（共通の関数は gui_helpers.ps1）。
# 画面遷移の一覧（docs\design\testing\index.md「画面のスモークテスト」）の #3・#11〜#18・#26 を S2 で確かめる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S2 インデックスの管理と作成" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        $script:source = Join-Path $TestDrive "元のフォルダ\営業"
        newGuiSourceFolder $script:source -Broken
    }

    It "追加・編集・［作成］の切り替え・作成・削除が動く" {
        $S = startGui $script:tool "S2"
        invokeGuiScene $S {
            # インデックスが無いので［1 インデックス管理］が選ばれる（#3）
            setGuiStep $S "起動時のタブ（インデックスが無い）"
            getGuiSelectedTab $S | Should -Be "IndexTab"

            # ［2 検索］の［インデックスを作成する］で［1］へ（#26）
            setGuiStep $S "［2 検索］の［インデックスを作成する］"
            selectGuiTab $S "SearchTab" "GoIndexTabButton"
            clickGui $S $S.Window "GoIndexTabButton" "［インデックスを作成する］"
            waitGui $S "［1 インデックス管理］が選ばれる" ${guiDefaultTimeout} { (getGuiSelectedTab $S) -eq "IndexTab" } | Out-Null

            # 追加: キャンセル（#11）
            setGuiStep $S "［追加…］→［キャンセル］"
            clickGui $S $S.Window "NewIndexButton" "［追加…］"
            $dialog = waitGuiWindow $S "インデックスの追加のダイアログ" -Id "FolderBox"
            clickGui $S $dialog "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "追加のダイアログ"
            @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count | Should -Be 0

            # 追加: 入力が足りないまま［OK］は、ダイアログの中に注意が出る（#12）
            setGuiStep $S "［追加…］→ 入力が足りないまま［OK］"
            clickGui $S $S.Window "NewIndexButton" "［追加…］"
            $dialog = waitGuiWindow $S "インデックスの追加のダイアログ" -Id "FolderBox"
            clickGui $S $dialog "OkButton" "［OK］"
            waitGui $S "注意（ErrorText）" ${guiDefaultTimeout} { (getGuiText (findGui $dialog -Id "ErrorText")) -ne "" } | Out-Null

            # ［参照…］: キャンセルすると変わらない、フォルダを選ぶと欄に入る（#13）
            setGuiStep $S "［参照…］→ OS のフォルダ選択（キャンセル）"
            clickGui $S $dialog "BrowseButton" "［参照…］"
            useGuiFolderPicker $S
            getGuiValue (findGui $dialog -Id "FolderBox") | Should -Be ""
            setGuiStep $S "［参照…］→ OS のフォルダ選択（フォルダを選ぶ）"
            clickGui $S $dialog "BrowseButton" "［参照…］"
            useGuiFolderPicker $S $script:source
            waitGui $S "フォルダの欄に入る" ${guiDefaultTimeout} { (getGuiValue (findGui $dialog -Id "FolderBox")) -eq $script:source } | Out-Null
            getGuiValue (findGui $dialog -Id "NameBox") | Should -Be "営業"

            # ［OK］で一覧に加わる（#11）
            setGuiStep $S "［OK］で追加"
            clickGui $S $dialog "OkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "追加のダイアログ"
            $row = waitGui $S "一覧に加わる" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            (getGuiRowTexts $row) | Should -Contain "営業"

            # 編集: キャンセル・名前の変更（#14）
            setGuiStep $S "［編集…］→［キャンセル］"
            selectGui $row
            clickGui $S $S.Window "EditIndexButton" "［編集…］"
            $dialog = waitGuiWindow $S "インデックスの編集のダイアログ" -Id "NameBox"
            clickGui $S $dialog "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "編集のダイアログ"
            setGuiStep $S "［編集…］→ 名前を変えて［OK］"
            clickGui $S $S.Window "EditIndexButton" "［編集…］"
            $dialog = waitGuiWindow $S "インデックスの編集のダイアログ" -Id "NameBox"
            setGuiText $S (findGui $dialog -Id "NameBox") "資料"
            clickGui $S $dialog "OkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "編集のダイアログ"
            $row = waitGui $S "一覧の名前が変わる" ${guiDefaultTimeout} {
                $r = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1
                if ($r -and ((getGuiRowTexts $r) -contains "資料")) { $r }
            }

            # ［作成］のチェックを切り替える（#16）。
            # UI オートメーションの Toggle は、チェックの状態を変えるだけで Click イベントを起こさないため、Click で行う設定への保存と
            # ［インデックス作成を開始］の可否の更新は確かめられない（マウスの操作でだけ動く。docs\design\testing\index.md「画面のスモークテスト」の対象外）
            setGuiStep $S "［作成］のチェックの切り替え"
            $check = findGui $row -Type CheckBox
            getGuiToggleState $check | Should -Be "On"
            toggleGui $check
            waitGui $S "チェックが外れる" ${guiDefaultTimeout} { (getGuiToggleState (findGui $row -Type CheckBox)) -eq "Off" } | Out-Null
            toggleGui (findGui $row -Type CheckBox)
            waitGui $S "チェックが付く" ${guiDefaultTimeout} { (getGuiToggleState (findGui $row -Type CheckBox)) -eq "On" } | Out-Null

            # 作成: 確認でキャンセルすると取りやめ、もう一度で取り込む（#17）
            setGuiStep $S "［インデックス作成を開始］→ 確認で［キャンセル］"
            clickGui $S $S.Window "IndexingButton" "［インデックス作成を開始］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $confirm "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $confirm "取り込みの確認"
            waitGui $S "取りやめて［インデックス作成を開始］に戻る" ${guiDefaultTimeout} {
                $b = findGui $S.Window -Id "IndexingButton"
                $b.Current.IsEnabled -and $b.Current.Name -eq "インデックス作成を開始"
            } | Out-Null
            @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count | Should -Be 1
            (getGuiText (findGui $S.Window -Id "IndexSummaryText")) | Should -BeLike "まだインデックスがありません*"

            setGuiStep $S "［インデックス作成を開始］→ 確認で［インデックス作成を開始］"
            $sw = [Diagnostics.Stopwatch]::StartNew()
            clickGui $S $S.Window "IndexingButton" "［インデックス作成を開始］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $confirm "StartButton" "確認の［インデックス作成を開始］"
            waitGuiWindowClosed $S $confirm "取り込みの確認"

            # 取り込みが終わると、完了の表示と失敗したファイルの一覧が出る（#18）
            setGuiStep $S "取り込みの完了・失敗したファイルの一覧"
            waitGui $S "取り込みの完了（集約ファイルができ、失敗 1 件が出る）" ${guiIndexTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*" -and
                    (getGuiText (findGui $S.Window -Id "FailedHeading")) -like "*失敗したファイル 1 件*" -and
                    (findGui $S.Window -Id "IndexingButton").Current.IsEnabled
            } | Out-Null
            $S.Timing["取り込み"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
            @(getGuiGridRows (findGui $S.Window -Id "FailedGrid")).Count | Should -Be 1
            (getGuiRowTexts @(getGuiGridRows (findGui $S.Window -Id "FailedGrid"))[0]) -join " " | Should -BeLike "*壊れた文書.docx*"
            Test-Path -LiteralPath "$($script:tool.Work)\index\資料" | Should -BeTrue

            # 削除: キャンセルすると残り、［削除する］で消える（#15）
            setGuiStep $S "［削除］→［キャンセル］"
            $row = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            selectGui $row
            clickGui $S $S.Window "RemoveIndexButton" "［削除］"
            $confirm = waitGuiWindow $S "削除の確認" -Id "HeadingText" -Text "一覧から削除しますか"
            clickGuiByName $S $confirm "キャンセル"
            waitGuiWindowClosed $S $confirm "削除の確認"
            @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count | Should -Be 1
            setGuiStep $S "［削除］→［削除する］"
            clickGui $S $S.Window "RemoveIndexButton" "［削除］"
            $confirm = waitGuiWindow $S "削除の確認" -Id "HeadingText" -Text "一覧から削除しますか"
            clickGuiByName $S $confirm "削除する"
            waitGuiWindowClosed $S $confirm "削除の確認"
            waitGui $S "一覧から消える" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count -eq 0 } | Out-Null
            waitGui $S "インデックスのフォルダが消える" ${guiDefaultTimeout} { !(Test-Path -LiteralPath "$($script:tool.Work)\index\資料") } | Out-Null

            closeGui $S
            Write-Host ("S2 の秒数: " + (($S.Timing.GetEnumerator() | ForEach-Object { "$($_.Key) $($_.Value)" }) -join "・"))
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
