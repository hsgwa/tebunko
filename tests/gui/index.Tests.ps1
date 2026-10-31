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

            # 頭: 見出しと説明の文は無く、ⓘ も無い
            setGuiStep $S "頭に見出し・説明・ⓘ が無い"
            @(getGuiTexts $S.Window) | Should -Not -Contain "検索するフォルダとインデックスを管理します。ファイルを変更したら［すべて更新］でインデックスを最新にします。"
            findGui $S.Window -Id "IndexScreenInfo" | Should -BeNullOrEmpty

            # 追加: フォルダ選びをキャンセルすると何も加わらない（#11）
            setGuiStep $S "［＋ フォルダを追加］→ OS のフォルダ選択（キャンセル）"
            clickGui $S $S.Window "NewIndexButton" "［＋ フォルダを追加］"
            useGuiFolderPicker $S
            @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")).Count | Should -Be 0

            # 一覧の表は右の領域の横いっぱいに広がる（仕切りと同じ左右の端）。
            # 窓を最小の幅（1024）にしても、頭の 3 つのボタンと説明は右で切れない。説明が出るのは、まだフォルダが無い今だけ（追加すると説明は消える）なので、ここで測る（ナビを最大に広げた幅は、ドラッグでしか作れないので、単体では見られない）
            setGuiStep $S "一覧の横幅と頭のボタン（最小の幅）"
            resizeGuiWindow $S 1024 600
            [math]::Round($S.Window.Current.BoundingRectangle.Width) | Should -BeGreaterThan 1000
            [math]::Round($S.Window.Current.BoundingRectangle.Width) | Should -BeLessThan 1100
            $edge = {
                param ($id)
                return (findGui $S.Window -Id $id).Current.BoundingRectangle
            }
            # 表の右端は窓の右端まで届く（窓の枠・影の分だけ内側。以前は右に 16 の余白があり、枠の分と合わせて 20 を超えていた）
            $win = $S.Window.Current.BoundingRectangle
            $gridRect = & $edge "IndexGrid"
            ($win.Right - $gridRect.Right) | Should -BeLessOrEqual 12 -Because "表が右の領域の端まで広がる"
            $area = $gridRect
            $buttons = New-Object System.Collections.Generic.List[object]
            foreach ($id in @("IndexingButton", "ActionsButton", "NewIndexButton")) { $buttons.Add((& $edge $id)) }
            foreach ($b in $buttons) { $b.Right | Should -BeLessOrEqual $area.Right -Because "頭のボタンが右で切れない" }
            $buttons[1].Left | Should -BeGreaterThan $buttons[0].Right -Because "頭のボタンが重ならない"
            $buttons[2].Left | Should -BeGreaterThan $buttons[1].Right -Because "頭のボタンが重ならない"
            $hintRect = & $edge "IndexingHint"
            $hintRect.Width | Should -BeGreaterThan 0 -Because "説明が出ている状態で測る（空だと矩形が空になり、比べても何も確かめられない）"
            $hintRect.Right | Should -BeLessOrEqual $area.Right -Because "説明が右で切れない"
            resizeGuiWindow $S 1280 800

            # 追加: フォルダを選ぶと、フォルダ名のインデックスが一覧に加わる（追加のダイアログは無い）
            setGuiStep $S "［＋ フォルダを追加］→ OS のフォルダ選択（フォルダを選ぶ）"
            clickGui $S $S.Window "NewIndexButton" "［＋ フォルダを追加］"
            useGuiFolderPicker $S $script:source
            $detailTitle = { getGuiText (findGui $S.Window -Id "IndexDetailTitle") }
            $row = waitGui $S "一覧に加わる" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            (getGuiRowTexts $row) | Should -Contain "営業"

            # 行を選ぶと、詳細の見出しが変わる（#14）
            setGuiStep $S "行を選ぶ"
            selectGui $row
            waitGui $S "詳細の見出し（営業）" ${guiDefaultTimeout} { (& $detailTitle) -eq "営業 - 詳細" } | Out-Null

            # 窓を低くしても、詳細を縦にスクロールして、一番下の「すべてのインデックスの合計」まで届く（使えない理由と手順の出し分けは、判断層の単体テストで確かめる）
            setGuiStep $S "詳細のスクロール"
            resizeGuiWindow $S 1024 600
            [math]::Round($S.Window.Current.BoundingRectangle.Width) | Should -BeGreaterThan 1000
            [math]::Round($S.Window.Current.BoundingRectangle.Width) | Should -BeLessThan 1100
            $summaryText = waitGuiById $S $S.Window "IndexSummaryText"
            $scrollItem = $null
            if ($summaryText.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern, [ref]$scrollItem)) { $scrollItem.ScrollIntoView() }
            waitGui $S "合計までスクロールで届く" ${guiDefaultTimeout} {
                $rect = (findGui $S.Window -Id "IndexSummaryText").Current.BoundingRectangle
                $rect.Bottom -le $S.Window.Current.BoundingRectangle.Bottom
            } | Out-Null
            resizeGuiWindow $S 1280 800

            # 詳細の名前の欄で直接直す。使えない名前は欄の下に理由が出て名前は変わらず、正しい名前は確定で一覧に反映される（Enter か欄から出る）
            setGuiStep $S "詳細の名前の欄で名前を直す"
            $nameBox = findGui $S.Window -Id "IndexDetailName"
            $nameBox.Current.IsEnabled | Should -BeTrue
            $nameBox.SetFocus()
            setGuiText $S $nameBox "a\b"
            (findGui $S.Window -Id "IndexGrid").SetFocus()
            waitGui $S "名前の注意が出る" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexDetailEditError")) -ne "" } | Out-Null
            getGuiValue (findGui $S.Window -Id "IndexDetailName") | Should -Be "営業" -Because "使えない名前は欄から出ると元に戻る"
            $nameBox = findGui $S.Window -Id "IndexDetailName"
            $nameBox.SetFocus()
            setGuiText $S $nameBox "資料"
            (findGui $S.Window -Id "IndexGrid").SetFocus()
            waitGui $S "一覧の名前が変わる" ${guiDefaultTimeout} { (getGuiRowTexts @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]) -contains "資料" } | Out-Null
            waitGui $S "詳細の見出し（資料）" ${guiDefaultTimeout} { (& $detailTitle) -eq "資料 - 詳細" } | Out-Null
            $row = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            selectGui $row

            # 詳細のフォルダパスの［変更…］: キャンセルすると変わらず、別のフォルダを選ぶとその場で変わる（欄に入力して確定したときと同じ道）
            setGuiStep $S "詳細のフォルダパスの［変更…］→ OS のフォルダ選択（キャンセル）"
            $pathText = { getGuiValue (findGui $S.Window -Id "IndexDetailPath") }
            (findGui $S.Window -Id "IndexDetailPathButton").Current.IsEnabled | Should -BeTrue
            & $pathText | Should -Be $script:source
            clickGui $S $S.Window "IndexDetailPathButton" "詳細の［変更…］"
            useGuiFolderPicker $S
            & $pathText | Should -Be $script:source
            setGuiStep $S "詳細のフォルダパスの［変更…］→ OS のフォルダ選択（別のフォルダ）"
            $moved = Join-Path (Split-Path -Parent $script:source) "移動先"
            New-Item -ItemType Directory -Path $moved -Force | Out-Null
            clickGui $S $S.Window "IndexDetailPathButton" "詳細の［変更…］"
            useGuiFolderPicker $S $moved
            waitGui $S "フォルダパスが変わる" ${guiDefaultTimeout} { (& $pathText) -eq $moved } | Out-Null
            setGuiStep $S "詳細のフォルダパスの［変更…］→ 元のフォルダへ戻す"
            clickGui $S $S.Window "IndexDetailPathButton" "詳細の［変更…］"
            useGuiFolderPicker $S $script:source
            waitGui $S "フォルダパスが元に戻る" ${guiDefaultTimeout} { (& $pathText) -eq $script:source } | Out-Null

            (findGui $S.Window -Id "IndexDetailPathButton").Current.Name | Should -Be "フォルダの場所を変える"
            # チェックは、保存しないその場の選び。初めは付いておらず、付けても設定は書き換わず、［すべて更新］の可否も変わらない（#16）
            setGuiStep $S "行のチェックの切り替え"
            $check = findGui $row -Type CheckBox
            getGuiToggleState $check | Should -Be "Off"
            $configWriteTime = (Get-Item -LiteralPath $script:tool.Config).LastWriteTimeUtc
            toggleGui $check
            waitGui $S "チェックが付く" ${guiDefaultTimeout} { (getGuiToggleState (findGui $row -Type CheckBox)) -eq "On" } | Out-Null
            # チェックの左端は、頭の「N / M 件を選択中」の左端にそろう（表の左端から 16。セルの内側の余白の分は、チェックの左の余白で引いてある）
            waitGui $S "件数の文が出る" ${guiDefaultTimeout} { (findGui $S.Window -Id "SelectionCountText").Current.BoundingRectangle.Width -gt 0 } | Out-Null
            $countLeft = (findGui $S.Window -Id "SelectionCountText").Current.BoundingRectangle.Left
            $checkLeft = (findGui $row -Type CheckBox).Current.BoundingRectangle.Left
            [Math]::Abs($checkLeft - $countLeft) | Should -BeLessOrEqual 2 -Because "チェックの左端と件数の文の左端がそろう"
            waitGui $S "［アクション ▾］が押せる" ${guiDefaultTimeout} { (findGui $S.Window -Id "ActionsButton").Current.IsEnabled } | Out-Null
            (findGui $S.Window -Id "IndexingButton").Current.IsEnabled | Should -BeTrue
            (Get-Item -LiteralPath $script:tool.Config).LastWriteTimeUtc | Should -Be $configWriteTime -Because "チェックは保存しない"
            toggleGui (findGui $row -Type CheckBox)
            waitGui $S "チェックが外れる" ${guiDefaultTimeout} { (getGuiToggleState (findGui $row -Type CheckBox)) -eq "Off" } | Out-Null

            # 作成: 確認でキャンセルすると取りやめ、もう一度で取り込む（#17）
            setGuiStep $S "［すべて更新］→ 確認で［キャンセル］"
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $confirm = waitGuiWindow $S "取り込みの確認のダイアログ" -Id "StartButton" -Timeout ${guiIndexTimeout}
            # 確認のダイアログと［インデックス管理］に、インデックスの保存先（ワークスペース）が出る。ローカルなので共有の警告は出ない（［フォルダを開く］は押さない）
            $confirmTexts = @(getGuiTexts $confirm) -join "`n"
            $confirmTexts | Should -BeLike "*インデックスの保存先*$($script:tool.Work)*"
            $confirmTexts | Should -Not -BeLike "*共有フォルダを選択しています*"
            testGuiVisible (findGui $S.Window -Id "IndexLocationWarning") | Should -BeFalse
            testGuiVisible (findGui $confirm -Id "OpenIndexLocationButton") | Should -BeTrue
            (getGuiText (findGui $S.Window -Id "IndexLocationText")) | Should -BeLike "$($script:tool.Work)*"
            testGuiVisible (findGui $S.Window -Id "OpenIndexLocationButton") | Should -BeTrue
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
                    (getGuiText (findGui $S.Window -Id "FailedFilesHeading")) -like "*失敗したファイル 1 件*" -and
                    (findGui $S.Window -Id "IndexingButton").Current.IsEnabled
            } | Out-Null
            $S.Timing["取り込み"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
            # 失敗は 1 行の入口になり、［一覧で見る］で［ファイルの一覧］が「失敗のみ」で開く
            testGuiVisible (findGui $S.Window -Id "ShowFailedFilesButton") | Should -BeTrue
            clickGui $S $S.Window "ShowFailedFilesButton" "失敗の入口の［一覧で見る］"
            waitGui $S "失敗のみの一覧が出る" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexFilesHeading")) -like "ファイル * 件*" -and
                    @(getGuiGridRows (findGui $S.Window -Id "IndexFilesGrid")).Count -ge 1
            } | Out-Null
            $failedOnly = @(getGuiGridRows (findGui $S.Window -Id "IndexFilesGrid"))
            $failedOnly.Count | Should -Be 1 -Because "失敗のみなので、失敗の 1 件だけ並ぶ"
            (getGuiRowTexts $failedOnly[0]) -join " " | Should -BeLike "*壊れた文書.docx*失敗*"
            clickGui $S $S.Window "IndexFilesCloseButton" "［閉じる］"
            waitGui $S "一覧が閉じる" ${guiDefaultTimeout} { !(testGuiVisible (findGui $S.Window -Id "IndexFilesGrid")) } | Out-Null
            # 失敗のあるだけの行は、エラーでなく警告の表示で、行の［更新］が使える
            $failRow = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            $failRowTexts = (getGuiRowTexts $failRow) -join " "
            $failRowTexts | Should -BeLike "*警告*"
            $failRowTexts | Should -Not -BeLike "*エラー*"
            $rowUpdate = findGui $failRow -Id "IndexRowUpdateButton"
            testGuiVisible $rowUpdate | Should -BeTrue
            $rowUpdate.Current.IsEnabled | Should -BeTrue
            Test-Path -LiteralPath "$($script:tool.Work)\store\content\資料" | Should -BeTrue

            # ［ファイルの一覧］: 取り込んだファイルが並び、失敗したものが分かる。［閉じる］で閉じる
            setGuiStep $S "［ファイルの一覧］"
            # 入口は右の列「インデックス情報」の見出しの右端（見本どおり。窓の右半分にあり、情報の欄に「うち失敗」「うち未取り込み」が出る）
            $winRect = $S.Window.Current.BoundingRectangle
            (findGui $S.Window -Id "IndexFilesButton").Current.BoundingRectangle.Left | Should -BeGreaterThan ($winRect.Left + $winRect.Width / 2)
            (getGuiTexts $S.Window) -join "`n" | Should -BeLike "*うち失敗*うち未取り込み*"
            clickGui $S $S.Window "IndexFilesButton" "詳細の［ファイルの一覧］"
            waitGui $S "取り込んだファイルが並ぶ" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexFilesHeading")) -like "ファイル * 件*"
            } | Out-Null
            # 押したら、一覧の見出しが窓の中に見える（詳細欄をスクロールしないと見えない所に開いて、何も出ないように見えるのを防ぐ）
            $headingRect = (findGui $S.Window -Id "IndexFilesHeading").Current.BoundingRectangle
            $paneRect = (findGui $S.Window -Id "IndexDetailTitle").Current.BoundingRectangle
            $winRect = $S.Window.Current.BoundingRectangle
            $headingRect.Top | Should -BeGreaterThan $paneRect.Bottom -Because "一覧の見出しが詳細の見出しの下にある"
            $headingRect.Bottom | Should -BeLessThan $winRect.Bottom -Because "一覧の見出しが窓の中に見える"
            # 表の最初の行まで見える（見出しだけが端に見えている状態にしない）
            waitGui $S "一覧の最初の行が見える所まで送られる" ${guiDefaultTimeout} {
                $rows = @(getGuiGridRows (findGui $S.Window -Id "IndexFilesGrid"))
                $rows.Count -gt 0 -and !$rows[0].Current.IsOffscreen
            } | Out-Null
            $fileRows = @(getGuiGridRows (findGui $S.Window -Id "IndexFilesGrid"))
            $fileRows.Count | Should -BeGreaterThan 1
            ($fileRows | ForEach-Object { (getGuiRowTexts $_) -join " " }) -join "`n" | Should -BeLike "*壊れた文書.docx*失敗*"
            clickGui $S $S.Window "IndexFilesCloseButton" "［閉じる］"
            waitGui $S "一覧が閉じる" ${guiDefaultTimeout} { !(testGuiVisible (findGui $S.Window -Id "IndexFilesGrid")) } | Out-Null

            # 更新が終わった帯には［検索する］が出る。押すと検索の画面に切り替わる
            setGuiStep $S "帯の［検索する］"
            $searchButton = findGui $S.Window -Id "IndexingSearchButton"
            $searchButton.Current.Name | Should -Be "検索する"
            clickGui $S $S.Window "IndexingSearchButton" "帯の［検索する］"
            waitGui $S "検索の画面に切り替わる" ${guiDefaultTimeout} { (getGuiSelectedTab $S) -eq "SearchTab" } | Out-Null
            selectGuiTab $S "IndexTab" "IndexingButton"

            # 残りが無い終わりの帯（注意）は閉じるボタン（×）で閉じられる
            setGuiStep $S "帯の×で閉じる"
            clickGui $S $S.Window "IndexingBannerClose" "帯の×"
            waitGui $S "帯が消える" ${guiDefaultTimeout} { !(testGuiVisible (findGui $S.Window -Id "IndexingProgressText")) } | Out-Null

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
            waitGui $S "インデックスのフォルダが消える" ${guiDefaultTimeout} { !(Test-Path -LiteralPath "$($script:tool.Work)\store\content\資料") } | Out-Null

            closeGui $S
            Write-Host ("S2 の秒数: " + (($S.Timing.GetEnumerator() | ForEach-Object { "$($_.Key) $($_.Value)" }) -join "・"))
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

Describe "S2 詳細の『サブフォルダを含める』" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        $script:source = Join-Path $TestDrive "元のフォルダ\営業"
        newGuiSourceFolder $script:source
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "営業"; path = $script:source; enabled = $true }) -Force
        writeGuiConfig $script:tool.Dir $config
    }

    It "外すと設定にすぐ書かれ、札『直下のみ』と［要更新］が出る。付け直すと外れ、一覧を読み直しても残る。行を選び直しただけでは書かない" {
        $S = startGui $script:tool "S2sub"
        invokeGuiScene $S {
            getGuiSelectedTab $S | Should -Be "IndexTab"
            setGuiStep $S "取り込む（［すべて更新］→［更新を開始］）"
            startGuiIndexing $S
            waitGui $S "取り込みの完了" ${guiIndexTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*" -and (findGui $S.Window -Id "IndexingButton").Current.IsEnabled
            } | Out-Null
            $rowTexts = { (getGuiRowTexts @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]) -join " " }
            (& $rowTexts) | Should -Not -BeLike "*直下のみ*"
            (& $rowTexts) | Should -Not -BeLike "*要更新*"

            setGuiStep $S "行を選ぶ（付いている。選び直しただけでは設定を書き換えない）"
            selectGui @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            waitGui $S "詳細の見出し（営業）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexDetailTitle")) -eq "営業 - 詳細" } | Out-Null
            $box = findGui $S.Window -Id "IndexDetailSubfolders"
            $box.Current.IsEnabled | Should -BeTrue
            getGuiToggleState $box | Should -Be "On"
            $configWriteTime = (Get-Item -LiteralPath $script:tool.Config).LastWriteTimeUtc
            selectGui @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            Start-Sleep -Milliseconds 500
            (Get-Item -LiteralPath $script:tool.Config).LastWriteTimeUtc | Should -Be $configWriteTime

            setGuiStep $S "外す → 設定に subfolders: false・札と［要更新］"
            toggleGui (findGui $S.Window -Id "IndexDetailSubfolders")
            waitGui $S "設定に false が書かれる" ${guiDefaultTimeout} { (readGuiConfig $script:tool).targetFolders[0].subfolders -eq $false } | Out-Null
            waitGui $S "札と［要更新］が出る" ${guiDefaultTimeout} { (& $rowTexts) -like "*直下のみ*" -and (& $rowTexts) -like "*要更新*" } | Out-Null

            setGuiStep $S "付け直す → 設定は true・［要更新］が外れる"
            toggleGui (findGui $S.Window -Id "IndexDetailSubfolders")
            waitGui $S "設定が true に戻る" ${guiDefaultTimeout} { (readGuiConfig $script:tool).targetFolders[0].subfolders -eq $true } | Out-Null
            waitGui $S "札と［要更新］が消える" ${guiDefaultTimeout} { (& $rowTexts) -notlike "*直下のみ*" -and (& $rowTexts) -notlike "*要更新*" } | Out-Null

            setGuiStep $S "外したまま、タブを移って戻っても［要更新］が残る"
            toggleGui (findGui $S.Window -Id "IndexDetailSubfolders")
            waitGui $S "［要更新］が出る" ${guiDefaultTimeout} { (& $rowTexts) -like "*要更新*" } | Out-Null
            selectGuiTab $S "SearchTab"
            selectGuiTab $S "IndexTab" "IndexGrid"
            waitGui $S "戻っても札と［要更新］が残る" ${guiDefaultTimeout} { (& $rowTexts) -like "*直下のみ*" -and (& $rowTexts) -like "*要更新*" } | Out-Null
            closeGui $S
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

Describe "S2 詳細の『取り込み対象の拡張子』" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        $script:source = Join-Path $TestDrive "元のフォルダ\営業"
        newGuiSourceFolder $script:source
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "営業"; path = $script:source; enabled = $true }) -Force
        writeGuiConfig $script:tool.Dir $config
    }

    It "［変更…］で .log を外すと設定にすぐ書かれ、一覧の行に［要更新］が出る。戻すと消える。見出しの『Excel』を押すと Excel の拡張子がまとめて外れ、もう一度押すと戻る" {
        $S = startGui $script:tool "S2ext"
        invokeGuiScene $S {
            getGuiSelectedTab $S | Should -Be "IndexTab"
            setGuiStep $S "取り込む（［すべて更新］→［更新を開始］）"
            startGuiIndexing $S
            waitGui $S "取り込みの完了" ${guiIndexTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*" -and (findGui $S.Window -Id "IndexingButton").Current.IsEnabled
            } | Out-Null
            $rowTexts = { (getGuiRowTexts @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]) -join " " }
            (& $rowTexts) | Should -Not -BeLike "*要更新*"

            setGuiStep $S "行を選ぶ（詳細の『取り込み対象の拡張子』は全部取り込む）"
            selectGui @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            waitGui $S "詳細の見出し（営業）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexDetailTitle")) -eq "営業 - 詳細" } | Out-Null
            (getGuiText (findGui $S.Window -Id "IndexDetailTextExtSummary")) | Should -Be "すべて取り込み"
            $change = findGui $S.Window -Id "ChangeIndexTextExtensionsButton"
            $change.Current.IsEnabled | Should -BeTrue

            setGuiStep $S "［変更…］で .log を外す → 設定に indexExtensionExcludes・［要更新］"
            invokeGui $S $change "［変更…］" -NoWait
            $dialog = waitGuiWindow $S "取り込み対象の拡張子のダイアログ" -Id "TextExtOkButton"
            toggleGui (findGui $dialog -Name ".log" -Type CheckBox)
            invokeGui $S (findGui $dialog -Id "TextExtOkButton") "［OK］"
            waitGui $S "設定に .log が書かれる" ${guiDefaultTimeout} { @((readGuiConfig $script:tool).indexExtensionExcludes."営業") -contains ".log" } | Out-Null
            waitGui $S "一覧の行に［要更新］が出る" ${guiDefaultTimeout} { (& $rowTexts) -like "*要更新*" } | Out-Null
            (getGuiText (findGui $S.Window -Id "IndexDetailTextExtSummary")) | Should -Be "対象外あり"

            setGuiStep $S "もう一度開いて .log を戻す → 設定から消え、［要更新］が外れる"
            invokeGui $S (findGui $S.Window -Id "ChangeIndexTextExtensionsButton") "［変更…］" -NoWait
            $dialog = waitGuiWindow $S "取り込み対象の拡張子のダイアログ" -Id "TextExtOkButton"
            toggleGui (findGui $dialog -Name ".log" -Type CheckBox)
            invokeGui $S (findGui $dialog -Id "TextExtOkButton") "［OK］"
            waitGui $S "設定から .log が消える" ${guiDefaultTimeout} { @((readGuiConfig $script:tool).indexExtensionExcludes."営業") -notcontains ".log" } | Out-Null
            waitGui $S "［要更新］が消える" ${guiDefaultTimeout} { (& $rowTexts) -notlike "*要更新*" } | Out-Null
            setGuiStep $S "［変更…］で見出し『Excel』を押す → .xlsx などが全部外れる"
            invokeGui $S (findGui $S.Window -Id "ChangeIndexTextExtensionsButton") "［変更…］" -NoWait
            $dialog = waitGuiWindow $S "取り込み対象の拡張子のダイアログ" -Id "TextExtOkButton"
            $heading = findGui $dialog -Id "TextExtGroup_Excel" -Type CheckBox
            getGuiToggleState $heading | Should -Be "On"
            toggleGui $heading
            waitGui $S "Excel の行が全部外れ、見出しが空になる" ${guiDefaultTimeout} {
                (getGuiToggleState (findGui $dialog -Name ".xlsx" -Type CheckBox)) -eq "Off" -and (getGuiToggleState (findGui $dialog -Id "TextExtGroup_Excel" -Type CheckBox)) -eq "Off"
            } | Out-Null
            (getGuiToggleState (findGui $dialog -Name ".docx" -Type CheckBox)) | Should -Be "On"
            invokeGui $S (findGui $dialog -Id "TextExtOkButton") "［OK］"
            waitGui $S "設定に Excel の 4 つが書かれる" ${guiDefaultTimeout} { try { (@((readGuiConfig $script:tool).indexExtensionExcludes."営業") | Sort-Object) -join "," -eq ".xls,.xlsb,.xlsm,.xlsx" } catch { $false } } | Out-Null

            setGuiStep $S "もう一度開いて .xlsx だけ戻す → 見出しは一部。見出しを押すと全部戻り、設定から消える"
            invokeGui $S (findGui $S.Window -Id "ChangeIndexTextExtensionsButton") "［変更…］" -NoWait
            $dialog = waitGuiWindow $S "取り込み対象の拡張子のダイアログ" -Id "TextExtOkButton"
            toggleGui (findGui $dialog -Name ".xlsx" -Type CheckBox)
            waitGui $S "見出しが一部になる" ${guiDefaultTimeout} { (getGuiToggleState (findGui $dialog -Id "TextExtGroup_Excel" -Type CheckBox)) -eq "Indeterminate" } | Out-Null
            toggleGui (findGui $dialog -Id "TextExtGroup_Excel" -Type CheckBox)
            waitGui $S "Excel の行が全部戻る" ${guiDefaultTimeout} { (getGuiToggleState (findGui $dialog -Name ".xls" -Type CheckBox)) -eq "On" -and (getGuiToggleState (findGui $dialog -Id "TextExtGroup_Excel" -Type CheckBox)) -eq "On" } | Out-Null
            invokeGui $S (findGui $dialog -Id "TextExtOkButton") "［OK］"
            waitGui $S "設定から Excel の分が消える" ${guiDefaultTimeout} { try { @((readGuiConfig $script:tool).indexExtensionExcludes."営業" | Where-Object { $_ }).Count -eq 0 } catch { $false } } | Out-Null
            closeGui $S
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
            (findGui $S.Window -Id "IndexDetailSubfolders").Current.IsEnabled | Should -BeFalse -Because "更新中は『サブフォルダを含める』が押せない"
            (findGui $S.Window -Id "ChangeIndexTextExtensionsButton").Current.IsEnabled | Should -BeFalse -Because "更新中は『取り込み対象の拡張子』の［変更…］が押せない"
            waitGui $S "行に［中止］が出る" ${guiDefaultTimeout} { findGui $row -Id "IndexRowStopButton" } | Out-Null

            # ［エクスポート…］［削除…］［インポート…］は［アクション ▾］のメニューの中。開いて、押せないことを確かめてから Esc で閉じる
            invokeGui $S (waitGuiById $S $S.Window "ActionsButton") "［アクション ▾］" -NoWait
            $menu = waitGuiWindow $S "アクションのメニュー" -Id "ActionDelete"
            foreach ($id in "ActionUpdate", "ActionExport", "ActionImport", "ActionDelete") {
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
                $b.Current.IsEnabled -and $b.Current.Name -eq "すべて更新" -and (testGuiVisible (findGui $S.Window -Id "IndexingResumeButton"))
            } | Out-Null
            $S.Timing["中止まで"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)

            # 中止のあとは残りがあるので、帯に［続きから再開］が出る。閉じるボタンもあり、閉じても再開の入口（［すべて更新］）は残る
            testGuiVisible (findGui $S.Window -Id "IndexingBannerClose") | Should -BeTrue -Because "中止の帯も閉じられる"

            # 帯のボタンが 2 つ並ぶ（残りがあれば［検索する］と［続きから再開］）ときは、見える枠の間が 10
            $search = findGui $S.Window -Id "IndexingSearchButton"
            $resume = findGui $S.Window -Id "IndexingResumeButton"
            $search | Should -Not -BeNullOrEmpty
            $search.Current.IsOffscreen | Should -BeFalse
            $gap = $resume.Current.BoundingRectangle.Left - $search.Current.BoundingRectangle.Right
            $gap | Should -BeGreaterThan 9.5
            $gap | Should -BeLessThan 10.5

            setGuiStep $S "中止の帯を×で閉じる（再開の入口は残る）"
            clickGui $S $S.Window "IndexingBannerClose" "帯の×"
            waitGui $S "帯が消える" ${guiDefaultTimeout} { !(testGuiVisible (findGui $S.Window -Id "IndexingProgressText")) } | Out-Null
            $b = findGui $S.Window -Id "IndexingButton"
            $b.Current.IsEnabled | Should -BeTrue
            $b.Current.Name | Should -Be "すべて更新"

            # 取り込みが止まると、また押せる（#20）
            setGuiStep $S "止まった後の［＋ フォルダを追加］・詳細の名前とフォルダパスの欄・［削除］"
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
            testGuiVisible (findGui $S.Window -Id "IndexingBannerClose") | Should -BeTrue -Because "中断の帯も閉じられる"

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
            # 止まるのを待つ間の待つ間のダイアログ（秒読み・押せない［閉じない］）は、止まるのが速い機械では見える前に消えるため、
            # ここでは確かめない（実機で見る。文言は wait_view.Tests.ps1）。止まったら画面が終了することを確かめる
            waitGui $S "取り込みを止めて画面が終了する" ${guiIndexTimeout} -AllowExited { $S.Process.HasExited } | Out-Null
            $S.Process.ExitCode | Should -Be 0
        }
    }

    It "中断した取り込みを再開して最後まで終えると、帯に閉じるボタンが出る" {
        # 中断の帯（残りあり）では閉じられないが、残りが無くなった完了の帯は閉じられる（8 秒は待たない）。
        # 最後まで終えられるよう、ファイルの数を減らした別の作業場所で行う
        $tooFast = "取り込みが終わってしまい、中止が間に合わなかった。tests\gui\index.Tests.ps1 の少量の場面のファイルの数（Copies）を増やす"
        $smallDir = Join-Path $TestDrive "少量"
        $tool = newGuiTool $smallDir @{ ingestThreads = 1 }
        $source = Join-Path $smallDir "元のフォルダ\少量"
        newGuiSourceFolder $source -Copies 60
        $config = readGuiConfig $tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "少量"; path = $source; enabled = $true }) -Force
        writeGuiConfig $tool.Dir $config

        $S = startGui $tool "S3"
        invokeGuiScene $S {
            setGuiStep $S "取り込みを始めて中止する"
            startGuiIndexing $S
            waitGui $S "取り込み中（［更新中…］）" ${guiDefaultTimeout} { testGuiIndexing $S } | Out-Null
            if (!(testGuiIndexing $S)) { throw $tooFast }
            clickGui $S $S.Window "IndexingStopButton" "［中止］"
            answerGuiConfirm $S "中止の確認" "中止しますか" "中止する"
            waitGui $S "更新が止まる（帯に［続きから再開］）" ${guiIndexTimeout} {
                $b = findGui $S.Window -Id "IndexingButton"
                $b.Current.IsEnabled -and $b.Current.Name -eq "すべて更新" -and (testGuiVisible (findGui $S.Window -Id "IndexingResumeButton"))
            } | Out-Null
            setGuiStep $S "続きから再開して最後まで終える"
            startGuiIndexing $S
            waitGui $S "完了の帯に閉じるボタンが出る" ${guiIndexTimeout} {
                !(testGuiIndexing $S) -and (getGuiIndexingBannerText $S) -like "*更新が終わりました*" -and
                    (testGuiVisible (findGui $S.Window -Id "IndexingBannerClose"))
            } | Out-Null
            closeGui $S
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
