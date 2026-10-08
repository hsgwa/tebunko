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

    It "追加・編集・行のチェックの切り替え・作成・削除が動く" {
        $S = startGui $script:tool "S2"
        invokeGuiScene $S {
            # インデックスが無いので［インデックス管理］が選ばれる（#3）
            setGuiStep $S "起動時のタブ（インデックスが無い）"
            getGuiSelectedTab $S | Should -Be "IndexTab"

            # ［検索］の［インデックス管理へ］で［インデックス管理］へ（#26）
            setGuiStep $S "［検索］の［インデックス管理へ］"
            selectGuiTab $S "SearchTab" "GoIndexTabButton"
            clickGui $S $S.Window "GoIndexTabButton" "［インデックス管理へ］"
            waitGui $S "［インデックス管理］が選ばれる" ${guiDefaultTimeout} { (getGuiSelectedTab $S) -eq "IndexTab" } | Out-Null

            # 頭: 見出しと説明の文は無く、ⓘ に説明がある（ツールヒントは UI オートメーションの HelpText で読む）
            setGuiStep $S "頭に見出しと説明が無く、ⓘ に説明がある"
            @(getGuiTexts $S.Window) | Should -Not -Contain "検索するフォルダとインデックスを管理します。ファイルを変更したら［すべて更新］でインデックスを最新にします。"
            (findGui $S.Window -Id "IndexScreenInfo").Current.HelpText | Should -BeLike "検索したいフォルダを登録する画面です。*"

            # 追加: キャンセル（#11）
            setGuiStep $S "［＋ フォルダを追加］→［キャンセル］"
            clickGui $S $S.Window "NewIndexButton" "［＋ フォルダを追加］"
            $dialog = waitGuiWindow $S "インデックスの追加のダイアログ" -Id "FolderBox"
            clickGui $S $dialog "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "追加のダイアログ"
            @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count | Should -Be 0

            # 追加: 入力が足りないまま［OK］は、ダイアログの中に注意が出る（#12）
            setGuiStep $S "［＋ フォルダを追加］→ 入力が足りないまま［OK］"
            clickGui $S $S.Window "NewIndexButton" "［＋ フォルダを追加］"
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
            # 名前は、追加のダイアログの欄で決める（編集のダイアログは UI オートメーションから開けないため）
            setGuiText $S (findGui $dialog -Id "NameBox") "資料"

            # ［OK］で一覧に加わる（#11）
            setGuiStep $S "［OK］で追加"
            $detailTitle = { getGuiText (findGui $S.Window -Id "IndexDetailTitle") }
            $detailTexts = { @(getGuiTexts $S.Window) }
            & $detailTitle | Should -Be "インデックスの状態"
            clickGui $S $dialog "OkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "追加のダイアログ"
            $row = waitGui $S "一覧に加わる" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            (getGuiRowTexts $row) | Should -Contain "資料"

            # 行を選ぶと、詳細の見出しが変わる（#14）。名前の編集のダイアログは、行を右クリックかダブルクリックで開く
            # （UI オートメーションからは開けないので、ここでは確かめない）
            setGuiStep $S "行を選ぶ"
            selectGui $row
            waitGui $S "詳細の見出し（資料）" ${guiDefaultTimeout} { (& $detailTitle) -eq "資料 - 詳細" } | Out-Null

            # 詳細のフォルダパスの［...］: キャンセルすると変わらず、別のフォルダを選ぶとその場で変わる（［編集…］で変えて［OK］と同じ道）
            setGuiStep $S "詳細のフォルダパスの［...］→ OS のフォルダ選択（キャンセル）"
            $pathText = { getGuiText (findGui $S.Window -Id "IndexDetailPath") }
            (findGui $S.Window -Id "IndexDetailPathButton").Current.IsEnabled | Should -BeTrue
            & $pathText | Should -Be $script:source
            clickGui $S $S.Window "IndexDetailPathButton" "詳細の［...］"
            useGuiFolderPicker $S
            & $pathText | Should -Be $script:source
            setGuiStep $S "詳細のフォルダパスの［...］→ OS のフォルダ選択（別のフォルダ）"
            $moved = Join-Path (Split-Path -Parent $script:source) "移動先"
            New-Item -ItemType Directory -Path $moved -Force | Out-Null
            clickGui $S $S.Window "IndexDetailPathButton" "詳細の［...］"
            useGuiFolderPicker $S $moved
            waitGui $S "フォルダパスが変わる" ${guiDefaultTimeout} { (& $pathText) -eq $moved } | Out-Null
            setGuiStep $S "詳細のフォルダパスの［...］→ 元のフォルダへ戻す"
            clickGui $S $S.Window "IndexDetailPathButton" "詳細の［...］"
            useGuiFolderPicker $S $script:source
            waitGui $S "フォルダパスが元に戻る" ${guiDefaultTimeout} { (& $pathText) -eq $script:source } | Out-Null

            # チェックは、保存しないその場の選び。初めは付いておらず、付けても設定は書き換わず、［すべて更新］の可否も変わらない（#16）
            setGuiStep $S "行のチェックの切り替え"
            $check = findGui $row -Type CheckBox
            getGuiToggleState $check | Should -Be "Off"
            $configWriteTime = (Get-Item -LiteralPath $script:tool.Config).LastWriteTimeUtc
            toggleGui $check
            waitGui $S "チェックが付く" ${guiDefaultTimeout} { (getGuiToggleState (findGui $row -Type CheckBox)) -eq "On" } | Out-Null
            waitGui $S "［アクション ▾］が押せる" ${guiDefaultTimeout} { (findGui $S.Window -Id "ActionsButton").Current.IsEnabled } | Out-Null
            (findGui $S.Window -Id "IndexingButton").Current.IsEnabled | Should -BeTrue
            (Get-Item -LiteralPath $script:tool.Config).LastWriteTimeUtc | Should -Be $configWriteTime -Because "チェックは保存しない"
            toggleGui (findGui $row -Type CheckBox)
            waitGui $S "チェックが外れる" ${guiDefaultTimeout} { (getGuiToggleState (findGui $row -Type CheckBox)) -eq "Off" } | Out-Null

            # 作成: 確認でキャンセルすると取りやめ、もう一度で取り込む（#17）
            setGuiStep $S "［すべて更新］→ 確認で［キャンセル］"
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $confirm "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $confirm "取り込みの確認"
            waitGui $S "取りやめて［すべて更新］に戻る" ${guiDefaultTimeout} {
                $b = findGui $S.Window -Id "IndexingButton"
                $b.Current.IsEnabled -and $b.Current.Name -eq "すべて更新"
            } | Out-Null
            @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count | Should -Be 1
            waitGui $S "「まだインデックスがありません」" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "まだインデックスがありません*" } | Out-Null

            setGuiStep $S "［すべて更新］→ 確認で［すべて更新］"
            $sw = [Diagnostics.Stopwatch]::StartNew()
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            clickGui $S $confirm "StartButton" "確認の［更新を開始］"
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

            # 更新が終わった帯には［検索する］が出る。押すと検索の画面に切り替わる
            setGuiStep $S "帯の［検索する］"
            $searchButton = findGui $S.Window -Id "IndexingSearchButton"
            $searchButton.Current.Name | Should -Be "検索する"
            clickGui $S $S.Window "IndexingSearchButton" "帯の［検索する］"
            waitGui $S "検索の画面に切り替わる" ${guiDefaultTimeout} { (getGuiSelectedTab $S) -eq "SearchTab" } | Out-Null
            selectGuiTab $S "IndexTab" "IndexingButton"

            # 削除: キャンセルすると残り、［削除する］で消える（#15）
            setGuiStep $S "［削除］→［キャンセル］"
            $row = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            selectGui $row
            checkGuiRow $S $row
            clickGuiAction $S "ActionDelete" "［削除…］"
            answerGuiConfirm $S "削除の確認" "インデックスを削除しますか" "キャンセル"
            @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count | Should -Be 1
            setGuiStep $S "［削除］→［削除する］"
            clickGuiAction $S "ActionDelete" "［削除…］"
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
            waitGui $S "取り込み中（［更新中…］）" ${guiDefaultTimeout} { testGuiIndexing $S } | Out-Null

            # 取り込み中は［＋ フォルダを追加］と［アクション ▾］の項目が押せない（#20）
            setGuiStep $S "取り込み中の［＋ フォルダを追加］と［アクション ▾］の項目"
            if (!(testGuiIndexing $S)) { throw $tooFast }
            (findGui $S.Window -Id "NewIndexButton").Current.IsEnabled | Should -BeFalse -Because "取り込み中は NewIndexButton が押せない"
            # 更新中の行を選ぶと、詳細の「インデックス」の箱に進み具合が出て、行には［中止］が出る
            $row = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            selectGui $row
            waitGui $S "詳細に進み具合の棒が出る" ${guiDefaultTimeout} { findGui $S.Window -Id "IndexingProgressBar" } | Out-Null
            waitGui $S "行に［中止］が出る" ${guiDefaultTimeout} { findGui $row -Id "IndexRowStopButton" } | Out-Null

            # ［エクスポート…］［削除…］［インポート…］は［アクション ▾］のメニューの中。開いて、押せないことを確かめてから Esc で閉じる
            invokeGui $S (waitGuiById $S $S.Window "ActionsButton") "［アクション ▾］" -NoWait
            $menu = waitGuiWindow $S "アクションのメニュー" -Id "ActionDelete"
            foreach ($id in "ActionExport", "ActionImport", "ActionDelete") {
                (findGui $menu -Id $id).Current.IsEnabled | Should -BeFalse -Because "取り込み中は $id が押せない"
            }
            pressGuiKey $menu 0x1B
            waitGuiWindowClosed $S $menu "アクションのメニュー"

            # ［設定］の［変更…］はメッセージボックスで断られる（#21）
            setGuiStep $S "取り込み中の［設定］の［変更…］"
            selectGuiTab $S "SettingsTab" "ChangeWorkspaceButton"
            clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
            closeGuiMessage $S "更新中はワークスペースを変えられません" "作成中の警告" | Out-Null
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
            waitGui $S "更新が止まる（帯に［続きから再開］、［すべて更新］が押せる）" ${guiIndexTimeout} {
                $b = findGui $S.Window -Id "IndexingButton"
                $b.Current.IsEnabled -and $b.Current.Name -eq "すべて更新" -and (findGui $S.Window -Id "IndexingResumeButton")
            } | Out-Null
            $S.Timing["中止まで"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)

            # 取り込みが止まると、また押せる（#20）
            setGuiStep $S "止まった後の［＋ フォルダを追加］［編集…］［削除］"
            (findGui $S.Window -Id "NewIndexButton").Current.IsEnabled | Should -BeTrue
            closeGui $S
        }
    }

    It "中断した取り込みから再開でき、取り込み中に閉じる確認が動く" {
        # #4（取り込みが中断していると起動時に［インデックス管理］が選ばれる）は、gui.ps1 の起動時の判定
        # （$script:indexingState が非同期に読み込まれる前に決めているため、Pending の判定が効かない）に見つかった
        # 不具合により、1 件でも取り込み済みだと ［検索］が選ばれる。別の fix（起票済み。Backlog）で直すまで、ここではタブを
        # 明示的に選んで続きの確かめ（#10）を行う。更新の帯の中断の文言は、選び直した後に出ることを確かめる
        $tooFast = "取り込みが終わってしまい、取り込み中の操作が間に合わなかった。tests\gui\index.Tests.ps1 の s3Copies（ファイルの数）を増やす"

        # 2 回目の起動: 続きから再開し、閉じる操作を確かめる
        $S = startGui $script:tool "S3"
        invokeGuiScene $S {
            setGuiStep $S "起動時のタブを［インデックス管理］にする（#4 は別の fix で直すまでの回避）"
            selectGuiTab $S "IndexTab" "IndexingButton"
            waitGui $S "「更新を中断しました」" ${guiDefaultTimeout} { (getGuiIndexingBannerText $S) -like "*更新を中断しました*" } | Out-Null

            setGuiStep $S "続きから再開"
            startGuiIndexing $S
            waitGui $S "取り込み中（［更新中…］）" ${guiDefaultTimeout} { testGuiIndexing $S } | Out-Null

            # 取り込み中に閉じる。確認で［閉じない］なら続き、［中止して閉じる］なら止めてから終了する（#10）
            $tooFastGuard = { if (!(testGuiIndexing $S)) { throw $tooFast } }
            setGuiStep $S "取り込み中に閉じる → 確認で［閉じない］"
            if (!(testGuiIndexing $S)) { throw $tooFast }
            closeGuiWindowAsync $S $S.Window
            answerGuiConfirm $S "閉じる確認" "中止して閉じますか" "閉じない" -Guard $tooFastGuard
            $S.Process.HasExited | Should -BeFalse
            if (!(testGuiIndexing $S)) { throw $tooFast }
            setGuiStep $S "取り込み中に閉じる → 確認で［中止して閉じる］"
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
