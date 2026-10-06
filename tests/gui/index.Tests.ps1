# 画面のスモークテスト S2: インデックスの管理と作成、S3: 作成中の操作（共通の関数は gui_helpers.ps1）。
# 画面遷移の一覧（docs\design\testing\gui-smoke.md「画面のスモークテスト」）の #3・#11〜#18・#26 を S2 で確かめる。
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

            # ［2 検索］の［インデックス管理へ］で［1］へ（#26）
            setGuiStep $S "［2 検索］の［インデックス管理へ］"
            selectGuiTab $S "SearchTab" "GoIndexTabButton"
            clickGui $S $S.Window "GoIndexTabButton" "［インデックス管理へ］"
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
            $detailTitle = { getGuiText (findGui $S.Window -Id "IndexDetailTitle") }
            $detailTexts = { @(getGuiTexts $S.Window) }
            & $detailTitle | Should -Be "インデックスの状態"
            clickGui $S $dialog "OkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "追加のダイアログ"
            $row = waitGui $S "一覧に加わる" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            (getGuiRowTexts $row) | Should -Contain "営業"

            # 編集: キャンセル・名前の変更（#14）
            setGuiStep $S "［編集…］→［キャンセル］"
            selectGui $row
            waitGui $S "詳細の見出し（営業）" ${guiDefaultTimeout} { (& $detailTitle) -eq "営業 - 詳細" } | Out-Null
            clickGuiRowMenu $S $row "EditIndexButton" "［編集…］"
            $dialog = waitGuiWindow $S "インデックスの編集のダイアログ" -Id "NameBox"
            clickGui $S $dialog "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "編集のダイアログ"
            setGuiStep $S "［編集…］→ 名前を変えて［OK］"
            clickGuiRowMenu $S $row "EditIndexButton" "［編集…］"
            $dialog = waitGuiWindow $S "インデックスの編集のダイアログ" -Id "NameBox"
            setGuiText $S (findGui $dialog -Id "NameBox") "資料"
            clickGui $S $dialog "OkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "編集のダイアログ"
            $row = waitGui $S "一覧の名前が変わる" ${guiDefaultTimeout} {
                $r = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1
                if ($r -and ((getGuiRowTexts $r) -contains "資料")) { $r }
            }

            waitGui $S "詳細の見出し（資料）" ${guiDefaultTimeout} { (& $detailTitle) -eq "資料 - 詳細" } | Out-Null

            # ［作成］のチェックを切り替えると、設定への保存と［インデックス作成を開始］の可否に反映される（#16）
            setGuiStep $S "［作成］のチェックの切り替え"
            $check = findGui $row -Type CheckBox
            getGuiToggleState $check | Should -Be "On"
            toggleGui $check
            waitGui $S "チェックが外れる" ${guiDefaultTimeout} { (getGuiToggleState (findGui $row -Type CheckBox)) -eq "Off" } | Out-Null
            waitGui $S "詳細の「インデックス作成」が変わる" ${guiDefaultTimeout} { @(& $detailTexts | Where-Object { $_ -like "対象にしない*" }).Count -eq 1 } | Out-Null
            waitGui $S "設定の enabled が false になる" ${guiDefaultTimeout} { (readGuiConfig $script:tool).targetFolders[0].enabled -eq $false } | Out-Null
            waitGui $S "［インデックス作成を開始］が押せなくなる" ${guiDefaultTimeout} { !(findGui $S.Window -Id "IndexingButton").Current.IsEnabled } | Out-Null
            toggleGui (findGui $row -Type CheckBox)
            waitGui $S "チェックが付く" ${guiDefaultTimeout} { (getGuiToggleState (findGui $row -Type CheckBox)) -eq "On" } | Out-Null
            waitGui $S "詳細の「インデックス作成」が戻る" ${guiDefaultTimeout} { @(& $detailTexts | Where-Object { $_ -eq "対象にする" }).Count -eq 1 } | Out-Null
            waitGui $S "設定の enabled が true に戻る" ${guiDefaultTimeout} { (readGuiConfig $script:tool).targetFolders[0].enabled -eq $true } | Out-Null
            waitGui $S "［インデックス作成を開始］が押せるようになる" ${guiDefaultTimeout} { (findGui $S.Window -Id "IndexingButton").Current.IsEnabled } | Out-Null

            # 起動時に、チェックの付いた行が表示されるだけでは、保存が重ねて走らないことを確かめる（#16）。
            # loadTargets 自体は保存を呼ばないため、起動し直した後の書き込みは 1 件でもあれば不具合（deliberate な保存と混じらず区別できる）
            setGuiStep $S "閉じて起動し直す"
            closeGui $S
            $reloadWriteTime = (Get-Item -LiteralPath $script:tool.Config).LastWriteTimeUtc
            $S.Window = $null
            $S.Process = startGuiProcess $script:tool
            waitGuiStarted $S
            setGuiStep $S "起動時に、チェックの付いた行の表示で余計な保存が走らないこと"
            waitGui $S "一覧に表示される" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 } | Out-Null
            Start-Sleep -Milliseconds 1000
            (Get-Item -LiteralPath $script:tool.Config).LastWriteTimeUtc | Should -Be $reloadWriteTime -Because "起動時に、チェックの付いた行が表示されるだけでは設定ファイルを書き直さない（loadTargets は保存を呼ばない）"

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
            waitGui $S "「まだインデックスがありません」" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "まだインデックスがありません*" } | Out-Null

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
            Test-Path -LiteralPath "$($script:tool.Work)\content_index\資料" | Should -BeTrue

            # 削除: キャンセルすると残り、［削除する］で消える（#15）
            setGuiStep $S "［削除］→［キャンセル］"
            $row = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            selectGui $row
            clickGuiRowMenu $S $row "RemoveIndexButton" "［削除］"
            answerGuiConfirm $S "削除の確認" "インデックスを削除しますか" "キャンセル"
            @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count | Should -Be 1
            setGuiStep $S "［削除］→［削除する］"
            clickGuiRowMenu $S $row "RemoveIndexButton" "［削除］"
            answerGuiConfirm $S "削除の確認" "インデックスを削除しますか" "削除する"
            waitGui $S "一覧から消える" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count -eq 0 } | Out-Null
            waitGui $S "インデックスのフォルダが消える" ${guiDefaultTimeout} { !(Test-Path -LiteralPath "$($script:tool.Work)\content_index\資料") } | Out-Null

            closeGui $S
            Write-Host ("S2 の秒数: " + (($S.Timing.GetEnumerator() | ForEach-Object { "$($_.Key) $($_.Value)" }) -join "・"))
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

# S3 は、取り込みの最中に操作する。読み取りのスレッドを 1 つ（ingestThreads）にして、ファイルの数を多めにし、
# 「中止」してもすぐには終わらないようにする。取り込みが終わってしまったときは「間に合わなかった」と分かる文言で失敗にする
# （$script:s3Copies を増やす）
Describe "S3 作成中の操作" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:s3Copies = 300
        $script:tool = newGuiTool $TestDrive @{ ingestThreads = 1 }
        $script:source = Join-Path $TestDrive "元のフォルダ\大量"
        newGuiSourceFolder $script:source -Copies $script:s3Copies
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "大量"; path = $script:source; enabled = $true }) -Force
        writeGuiConfig $script:tool.Dir $config
    }

    It "作成中は追加・編集・削除が押せず、中止の確認が動く" {
        $tooFast = "取り込みが終わってしまい、取り込み中の操作が間に合わなかった。tests\gui\index.Tests.ps1 の s3Copies（ファイルの数）を増やす"

        # 1 回目の起動: 取り込みを始め、取り込み中の操作を確かめて、中止する
        $S = startGui $script:tool "S3"
        invokeGuiScene $S {
            getGuiSelectedTab $S | Should -Be "IndexTab"
            setGuiStep $S "取り込みを始める"
            $sw = [Diagnostics.Stopwatch]::StartNew()
            startGuiIndexing $S
            waitGui $S "取り込み中（［インデックス作成中…］）" ${guiDefaultTimeout} { testGuiIndexing $S } | Out-Null

            # 取り込み中は［追加…］［編集…］［削除］が押せない（#20）
            setGuiStep $S "取り込み中の［追加…］［編集…］［削除］"
            if (!(testGuiIndexing $S)) { throw $tooFast }
            (findGui $S.Window -Id "NewIndexButton").Current.IsEnabled | Should -BeFalse -Because "取り込み中は NewIndexButton が押せない"
            # ［編集…］［削除］は行のメニューの中。［⋯］で開いて、押せないことを確かめてから Esc で閉じる
            $row = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            invokeGui $S (waitGuiById $S $row "IndexRowMenuButton") "行の［⋯］" -NoWait
            $menu = waitGuiWindow $S "行のメニュー" -Id "EditIndexButton"
            foreach ($id in "EditIndexButton", "RemoveIndexButton") {
                (findGui $menu -Id $id).Current.IsEnabled | Should -BeFalse -Because "取り込み中は $id が押せない"
            }
            pressGuiKey $menu 0x1B
            waitGuiWindowClosed $S $menu "行のメニュー"

            # ［8 設定］の［変更…］はメッセージボックスで断られる（#21）
            setGuiStep $S "取り込み中の［8 設定］の［変更…］"
            selectGuiTab $S "SettingsTab" "ChangeWorkspaceButton"
            clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
            closeGuiMessage $S "作成中はワークスペースを変えられません" "作成中の警告" | Out-Null
            selectGuiTab $S "IndexTab" "IndexingStopButton"

            # ［中止］→ 確認で［キャンセル］なら続き、［中止する］なら止まる（#19）
            setGuiStep $S "［中止］→ 確認で［キャンセル］"
            if (!(testGuiIndexing $S)) { throw $tooFast }
            clickGui $S $S.Window "IndexingStopButton" "［中止］"
            answerGuiConfirm $S "中止の確認" "中止しますか" "キャンセル"
            if (!(testGuiIndexing $S)) { throw $tooFast }
            setGuiStep $S "［中止］→ 確認で［中止する］"
            clickGui $S $S.Window "IndexingStopButton" "［中止］"
            answerGuiConfirm $S "中止の確認" "中止しますか" "中止する"
            waitGui $S "取り込みが止まる（［続きから再開］）" ${guiIndexTimeout} {
                $b = findGui $S.Window -Id "IndexingButton"
                $b.Current.IsEnabled -and $b.Current.Name -like "続きから再開*"
            } | Out-Null
            $S.Timing["中止まで"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)

            # 取り込みが止まると、また押せる（#20）
            setGuiStep $S "止まった後の［追加…］［編集…］［削除］"
            (findGui $S.Window -Id "NewIndexButton").Current.IsEnabled | Should -BeTrue
            closeGui $S
        }
    }

    It "中断した取り込みから再開でき、取り込み中に閉じる確認が動く" {
        # #4（取り込みが中断していると起動時に［1 インデックス管理］が選ばれる）は、gui.ps1 の起動時の判定
        # （$script:indexingState が非同期に読み込まれる前に決めているため、Pending の判定が効かない）に見つかった
        # 不具合により、1 件でも取り込み済みだと ［2 検索］が選ばれる。別の fix（起票済み。Backlog）で直すまで、ここではタブを
        # 明示的に選んで続きの確かめ（#10）を行う。IndexingStateText の中断の文言は、選び直した後に出ることを確かめる
        $tooFast = "取り込みが終わってしまい、取り込み中の操作が間に合わなかった。tests\gui\index.Tests.ps1 の s3Copies（ファイルの数）を増やす"

        # 2 回目の起動: 続きから再開し、閉じる操作を確かめる
        $S = startGui $script:tool "S3"
        invokeGuiScene $S {
            setGuiStep $S "起動時のタブを［1 インデックス管理］にする（#4 は別の fix で直すまでの回避）"
            selectGuiTab $S "IndexTab" "IndexingStateText"
            waitGui $S "「まだ取り込んでいないファイルがあります」" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexingStateText")) -like "*まだ取り込んでいないファイルがあります*" } | Out-Null

            setGuiStep $S "続きから再開"
            startGuiIndexing $S
            waitGui $S "取り込み中（［インデックス作成中…］）" ${guiDefaultTimeout} { testGuiIndexing $S } | Out-Null

            # 取り込み中に閉じる。確認で［閉じない］なら続き、［インデックス作成を止めて閉じる］なら止めてから終了する（#10）
            $tooFastGuard = { if (!(testGuiIndexing $S)) { throw $tooFast } }
            setGuiStep $S "取り込み中に閉じる → 確認で［閉じない］"
            if (!(testGuiIndexing $S)) { throw $tooFast }
            closeGuiWindowAsync $S $S.Window
            answerGuiConfirm $S "閉じる確認" "中止して閉じますか" "閉じない" -Guard $tooFastGuard
            $S.Process.HasExited | Should -BeFalse
            if (!(testGuiIndexing $S)) { throw $tooFast }
            setGuiStep $S "取り込み中に閉じる → 確認で［インデックス作成を止めて閉じる］"
            closeGuiWindowAsync $S $S.Window
            $confirm = waitGuiWindow $S "閉じる確認" -Id "HeadingText" -Text "中止して閉じますか" -Guard $tooFastGuard
            clickGuiByName $S $confirm "中止して閉じる"
            waitGui $S "取り込みを止めて画面が終了する" ${guiIndexTimeout} -AllowExited { $S.Process.HasExited } | Out-Null
            $S.Process.ExitCode | Should -Be 0
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
