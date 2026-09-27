# 検索結果から元のファイルを開く・パスをコピーする・結果を書き出す。


# ---- 元のファイルを開く・コピー・出力 ----


function getSourcePath {
    # 元のファイルのパス（ファイルがあるかは確かめない）。元の場所が分からなければ $null
    param (
        $row
    )

    return resolveSourcePath $row $script:sourceFolderMaps
}

$script:openSourceRequest = [ref]0  # 元のファイルを確かめる依頼の番号。増やすたびに前の依頼の結果を捨てる。
# [ref] のまま閉じ込める（スクリプトブロックの中で $script:openSourceRequest を読むと、遠く離れたスレッド・
# タイマーから呼ばれたときに増やす前の値のまま固まって読めることがあるため。tests\tebunko\ui\open_source.Tests.ps1 で確かめている）
$script:openSourcePendingRow = [ref]$null  # 裏のスレッドに依頼を出したままの行（.Value が $null なら待っている依頼は無い）。
# $script:openSourceRequest と同じ理由で [ref] の箱にする。GetNewClosure() の中で
# 「$script:openSourcePendingRow = ...」と書き換えても、閉じ込めた側だけの変数になり元の箱には届かない
# （.Value を書き換えれば、同じ箱を見ている側すべてに届く）。
# 同じ行をもう一度開いても新しい依頼を出さない（届かない共有では、列の 2 つのスレッドが同じ行の依頼で塞がるため）
$script:openSourcePendingPath = [ref]""  # 待っている行を、確かめている間の文言に出すためのパス（もう一度開いたときに出し直す）

function cancelPendingSourceLookup {
    # 待っている元のファイルの確認を打ち切る（新しく検索を始めたとき・ワークスペースを変えたときに呼ぶ）。
    # 依頼の番号を進めて前の依頼の結果を捨て、待っている行の記録・カーソルを戻す
    $script:openSourceRequest.Value++
    $script:openSourcePendingRow.Value = $null
    $window.Cursor = $null
}

function findSourceFile {
    # 元のファイルの場所を確かめ、見つかれば onFound { param($path) } を呼ぶ（見つからない・選ばなかったときは呼ばない）。
    # ネットワークにあるときは裏のスレッドで確かめ、結果が届くまで画面のスレッドは待たない（届かない共有で「応答なし」に
    # しないため）。ローカルならその場で確かめる。待っている間に同じ行をもう一度開いても、新しい依頼は出さない
    # （確かめている間の文言は出し直す。開き方（通常・読み取り専用・フォルダを開く 等）が違っても、同じ行の待ちは 1 つにまとめる）
    param (
        $row,
        [scriptblock]$onFound
    )

    $pendingRowBox = $script:openSourcePendingRow
    if ($row -eq $pendingRowBox.Value) {
        # この行はもう裏のスレッドに依頼済みで、まだ結果が届いていない。確かめている最中であることを出し直す
        setStatus (getSourceCheckingStatus $script:openSourcePendingPath.Value)
        return
    }

    $requestBox = $script:openSourceRequest
    $requestBox.Value++
    $requestId = $requestBox.Value
    $location = getSourceLocation $row $script:sourceFolderMaps
    $book = $row.Book

    # 呼ぶ関数は変数で捕まえてから閉じ込める（スクリプトブロックの中で名前のまま呼ぶと、遠く離れたスレッド・
    # タイマーから呼ばれたときに見つからないことがあるため）
    $applyState = ${function:applySourceFileState}
    $apply = {
        param ($state)
        if ($row -eq $pendingRowBox.Value) {
            $pendingRowBox.Value = $null
        }
        if ($requestId -ne $requestBox.Value) {
            # 待っている間に別の行を開いた・新しく検索した・ワークスペースを変えた。前の依頼は捨てる
            return
        }
        & $applyState $state $location $book $onFound
    }.GetNewClosure()

    if (!$location.Known) {
        # 元のフォルダの記録が無い（Folder が空）。row.Root・row.RelDir（インデックスの中の位置）から知らせる
        $relPath = if ($row.RelDir) { "$($row.RelDir)\$book" } else { $book }
        promptSourceMissing $location $book (joinSourcePath $row.Root $row.RelDir $book) $relPath "" $onFound
        return
    }
    if (!(testNetworkPath $location.Folder)) {
        & $apply (findSourceFileState $location $book)
        return
    }
    $pendingRowBox.Value = $row
    $checkingPath = joinSourcePath $location.Folder $location.Rest $book
    $script:openSourcePendingPath.Value = $checkingPath
    setStatus (getSourceCheckingStatus $checkingPath)
    $window.Cursor = [System.Windows.Input.Cursors]::AppStarting
    # ${pathStateOther} も、外の変数を直に書かず、いったんローカル変数に受けてから閉じ込める（$apply と同じ理由）
    $otherState = ${pathStateOther}
    startJob {
        param ($location, $book)
        findSourceFileState $location $book
    } @($location, $book) {
        param ($output, $errorText)
        if ($errorText) {
            & $apply @{ State = $otherState; Message = $errorText }
        } else {
            & $apply $output[0]
        }
    }.GetNewClosure() "network"
}

function applySourceFileState {
    # findSourceFileState（または裏の仕事）の結果を画面に反映する。見つかれば onFound を呼ぶ。
    # 見つからない・接続できない・その他のときは、見つかるまで（またはあきらめるまで）確認・フォルダ選択で進める
    param (
        $state,
        $location,
        [string]$book,
        [scriptblock]$onFound
    )

    $window.Cursor = $null
    if ($state.State -eq ${pathStateFound}) {
        if ($state.Alias -and $location.Name) {
            setIndexSourceFolder $location.Name $state.Alias
            $script:sourceFolderMaps = @{}
            setStatus "インデックス [$($location.Name)] の元のフォルダを $($state.Alias) に変えました"
        }
        & $onFound $state.Path
        return
    }
    if ($state.State -eq ${pathStateUnreachable} -or $state.State -eq ${pathStateOther}) {
        promptSourceConnectFailure $state $location $book $onFound
        return
    }
    # State=Missing はここまで来た時点で必ず Known（findSourceFileState は Known のときだけ呼ぶ）
    $path = joinSourcePath $location.Folder $location.Rest $book
    $relPath = if ($location.Rest) { "$($location.Rest)\$book" } else { $book }
    promptSourceMissing $location $book $path $relPath ([string]$state.Initial) $onFound
}

function promptSourceConnectFailure {
    # 接続できない・その他のとき。知らせて、［フォルダを選ぶ］を選べば見つからないときと同じ流れに進む
    param (
        $state,
        $location,
        [string]$book,
        [scriptblock]$onFound
    )

    $dialog = getSourceConnectFailureDialog $book $state.State $location.Folder $state.Message
    setStatus (getSourceConnectFailureStatus $state.State $location.Folder $state.Message)
    $answer = showConfirm -heading $dialog.Heading `
        -facts @((factGone $dialog.Title $dialog.Detail), (factNext $dialog.Hint)) `
        -choices @(@{ Text = "フォルダを選ぶ"; Value = "pick" })
    if ($answer -eq "pick") {
        $path = joinSourcePath $location.Folder $location.Rest $book
        $relPath = if ($location.Rest) { "$($location.Rest)\$book" } else { $book }
        promptSourceMissing $location $book $path $relPath "" $onFound
    }
}

function promptSourceMissing {
    # 見つからない・記録が無いとき、フォルダを選んでもらって探す（見つかるまで繰り返す）。
    # findMovedSource は、選んだ直後のフォルダ（届いている）を調べるため画面のスレッドのままにする。
    #   path: 知らせに出す（記録されていた・見つからなかった）パス
    #   relPath: インデックスの中の相対パス（記録が無いときの知らせに使う）
    #   initial: フォルダ選択の開始フォルダ（findSourceFileState が返した Initial。無ければ空）
    param (
        $location,
        [string]$book,
        [string]$path,
        [string]$relPath,
        [string]$initial,
        [scriptblock]$onFound
    )

    if ($location.Known) {
        $missing = factGone "記録されていた場所にありません" $path
        $description = "「$($location.Folder)」に当たるフォルダ（または $book のあるフォルダ）を選んでください"
    } else {
        $missing = factGone "このファイルが今どこにあるか、記録がありません" $relPath
        $description = "$book のあるフォルダ（またはインデックス [$($location.Name)] の元のフォルダ）を選んでください"
    }
    $facts = @(
        $missing,
        (factNext "今ある場所のフォルダを選べば開けます" "選んだ場所はインデックス「$($location.Name)」に覚えさせるので、同じインデックスのほかのファイルも次から開けます")
    )

    while ($true) {
        if ((showConfirm -heading "$book が見つかりません" -facts $facts `
                -choices @(@{ Text = "フォルダを選ぶ"; Value = "pick" })) -ne "pick") {
            setStatus (getSourceNotFoundStatus $path)
            return
        }
        $picked = selectFolder $description $initial
        if (!$picked) {
            setStatus (getSourceNotFoundStatus $path)
            return
        }

        $found = findMovedSource $picked $location.Rest $book
        if ($found) {
            if ($found.Root -and $location.Name) {
                setIndexSourceFolder $location.Name $found.Root
                $script:sourceFolderMaps = @{}
                setStatus "インデックス [$($location.Name)] の元のフォルダを $($found.Root) に変えました"
            } else {
                setStatus "開きました：$($found.Path)"
            }
            & $onFound $found.Path
            return
        }
        $facts = @(
            (factGone "選んだフォルダの中にありませんでした" "選んだフォルダ：${picked}`n探したファイル：${relPath}"),
            (factNext "別のフォルダを選んで、もう一度探せます")
        )
        $initial = $picked
    }
}

function openWithShell {
    # ファイルを既定のアプリで開く。開き方（mode）は、エクスプローラーの右クリックメニューと同じ動詞で行う。
    #   読み取り専用 → OpenAsReadOnly、新規 → New（元のファイルを基にした無題の文書。元のファイルを占有しない）
    # その動詞が登録されていない種類のファイルは、そのまま開いて $false を返す
    param (
        [string]$path,
        [string]$mode
    )

    $verb = switch ($mode) {
        ${openModeReadOnly} { "OpenAsReadOnly" }
        ${openModeNew}      { "New" }
        default             { $null }
    }
    if ($verb) {
        # Start-Process は [ ] をワイルドカードとして扱うため使わない。
        # ProcessStartInfo.Verbs は New を一覧に含めないため、実行してみて、無ければ例外で分かる
        $info = New-Object System.Diagnostics.ProcessStartInfo($path)
        $info.UseShellExecute = $true
        $info.Verb = $verb
        try {
            [void][System.Diagnostics.Process]::Start($info)
            return $true
        } catch [System.ComponentModel.Win32Exception] {
            # その動詞が登録されていない
        }
    }
    Invoke-Item -LiteralPath $path
    return (-not $verb)
}

function openInExcel {
    # 表示中の Excel（無ければ新しく起動）でブックを開き、該当シートの該当セルを選択する。
    # 開き方（mode）: 通常 = そのまま開く / 読み取り専用 = ReadOnly で開く /
    #                 新規 = 元のファイルを基にした新しいブック（無題）として開く（元のファイルを占有しない）
    param (
        [string]$path,
        [string]$location,
        [string]$cell,
        [string]$mode = ${openModeNormal}
    )

    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        # インデクサがバックグラウンドで使っている Excel は使わない
        if (!$excel.Visible) {
            $excel = $null
        }
    } catch {
        $excel = $null
    }
    if ($null -eq $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.UserControl = $true
    }

    $book = $null
    if ($mode -eq ${openModeNew}) {
        # 元のファイルをテンプレートとして新しいブックを作る（読み込んだ後は元のファイルを開いたままにしない）
        $book = $excel.Workbooks.Add($path)
    } else {
        # 既に開いているブックがあれば、そのまま使う（同じブックを二重に開けないため。開き方も既に開いたときのまま）
        foreach ($openBook in $excel.Workbooks) {
            if ($openBook.FullName -eq $path) {
                $book = $openBook
                break
            }
        }
        if ($null -eq $book) {
            #   引数: Filename, UpdateLinks, ReadOnly
            $book = $excel.Workbooks.Open($path, [Type]::Missing, ($mode -eq ${openModeReadOnly}))
        }
    }
    $book.Activate()

    # 場所はシート名。名前が同じシートを選ぶ。
    # 以前の版のインデックスは、ファイル名に使えない文字を全角に置き換えてあるため、同じ名前のシートが無ければ
    # 全角に置き換えて一致するシートを選ぶ（`衝突"` と `衝突”` のように、置き換えると重なるシートがあるため、同じ名前を優先する）
    $target = $null
    $sameSafeName = $null
    foreach ($sheet in $book.Worksheets) {
        if ($sheet.Name -eq $location) {
            $target = $sheet
            break
        }
        if ($null -eq $sameSafeName -and (toSafeFileName $sheet.Name) -eq $location) {
            $sameSafeName = $sheet
        }
    }
    if ($null -eq $target) {
        $target = $sameSafeName
    }
    if ($null -ne $target) {
        $target.Activate()
        if ($cell) {
            $target.Range($cell).Select()
        }
    }
    if ($excel.WindowState -eq -4140) {
        # 最小化されていれば元に戻す（xlMinimized → xlNormal）
        $excel.WindowState = -4143
    }
    # Excel は Visible にして前面に出す（P/Invoke の SetForegroundWindow は実行時コンパイル（csc.exe）を無くすため廃止）
    $excel.Visible = $true
    try { $excel.ActiveWindow.Activate() } catch { }
}

function getOpenMode {
    # ダブルクリック・Enter・［開く］での開き方（［開き方］の選択。${openModes} のいずれか）
    $item = $ui.OpenModeCombo.SelectedItem
    if ($null -ne $item -and (${openModes} -contains $item.Tag)) {
        return [string]$item.Tag
    }
    return ${openModeNormal}
}

function setOpenMode {
    # ［開き方］の選択を設定の値に合わせる（起動時。選んだことにはしないため、設定は保存しない）
    param (
        [string]$mode
    )

    $script:loadingOpenMode = $true
    try {
        foreach ($item in $ui.OpenModeCombo.Items) {
            if ($item.Tag -eq $mode) {
                $ui.OpenModeCombo.SelectedItem = $item
                return
            }
        }
        $ui.OpenModeCombo.SelectedIndex = 0
    } finally {
        $script:loadingOpenMode = $false
    }
}

function updateOpenMenu {
    # 右クリックメニューは3つの開き方をすべて出し、既定の開き方（ダブルクリック・Enter と同じ）に Enter を表示する
    $mode = getOpenMode
    $ui.MenuOpen.InputGestureText         = $(if ($mode -eq ${openModeNormal})   { "Enter" } else { "" })
    $ui.MenuOpenReadOnly.InputGestureText = $(if ($mode -eq ${openModeReadOnly}) { "Enter" } else { "" })
    $ui.MenuOpenNew.InputGestureText      = $(if ($mode -eq ${openModeNew})      { "Enter" } else { "" })
}

function openSource {
    # 選択行の元のファイルを開く。mode で開き方（通常・読み取り専用・新規）を指定する。
    # 元のファイルがネットワークにあれば、見つかるまで（または見つからない・接続できないと分かるまで）画面のスレッドは待たない
    param (
        [string]$mode = (getOpenMode)
    )

    $row = getCurrentHitRow
    if ($null -eq $row) {
        return
    }
    # 呼ぶ関数は変数で捕まえてから閉じ込める（findSourceFile の $apply と同じ理由）
    $openFound = ${function:openFoundSource}
    findSourceFile $row {
        param ($path)
        & $openFound $path $row $mode
    }.GetNewClosure()
}

function openFoundSource {
    # findSourceFile が見つけたファイルを開く（openSource の続き）
    param (
        [string]$path,
        $row,
        [string]$mode
    )

    $how = switch ($mode) {
        ${openModeReadOnly} { "読み取り専用で開きました" }
        ${openModeNew}      { "新規で開きました" }
        default             { "開きました" }
    }
    $fallback = switch ($mode) {
        ${openModeReadOnly} { "読み取り専用で開けなかったため、元のファイルを開きました" }
        default             { "新規で開けなかったため、元のファイルを開きました" }
    }

    if ($row.IsExcel) {
        setStatus "Excel で開いています：${path}"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            # 図形・コメントの場所（"<シート名>[図形]" 等）は、そのシートの図形の左上・コメントのセルを選ぶ
            openInExcel $path (splitObjectPlace $row.Location).Base $row.MatchCell $mode
            setStatus "${how}：${path}"
        } catch {
            # Excel を操作できない場合（ダイアログを表示中など）は、ファイルを開くだけにする
            if (openWithShell $path $mode) {
                setStatus "${how}（該当セルへの移動はできませんでした）：${path}"
            } else {
                setStatus "${fallback}（該当セルへの移動はできませんでした）：${path}"
            }
        } finally {
            $window.Cursor = $null
        }
    } elseif (openWithShell $path $mode) {
        setStatus "${how}：${path}"
    } else {
        setStatus "${fallback}：${path}"
    }
}

function openSourceFolder {
    $row = getCurrentHitRow
    if ($null -eq $row) {
        return
    }
    findSourceFile $row {
        param ($path)
        Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"${path}`""
    }
}

function copySelectedRows {
    $rows = getSelectedRows
    if ($rows.Count -eq 0) {
        return
    }
    $result = toSearchResultLines $rows
    $text = ((@($result.Header) + $result.Lines.ToArray()) -join "`r`n") + "`r`n"
    [System.Windows.Clipboard]::SetText($text)
    setStatus "$($rows.Count) 行をコピーしました（Excel に貼り付けると、元の列の位置に並びます）"
}

function copySourcePath {
    $row = getCurrentHitRow
    if ($null -eq $row) {
        return
    }
    $path = getSourcePath $row
    if (!$path) {
        # 元のファイルの場所が分からない（インデックスだけを別の PC にコピーした等）ときは、インデックスの中の位置（フォルダ\元のファイル名）を写す
        $path = if ($row.RelDir) { "$($row.RelDir)\$($row.Book)" } else { $row.Book }
    }
    [System.Windows.Clipboard]::SetText($path)
    setStatus "パスをコピーしました：${path}"
}

function exportResults {
    if ($null -eq $script:lastSearch) {
        setStatus "先に検索してください。"
        return
    }
    $rows = getViewRows
    try {
        [System.IO.Directory]::CreateDirectory($workspace.Dir) | Out-Null
        $writer = New-Object System.IO.StreamWriter($workspace.ResultFile, $false, ${utf8Bom})
        try {
            writeSearchResult $writer $script:lastSearch.Word $rows
        } finally {
            $writer.Close()
        }
    } catch [System.IO.IOException] {
        setStatus "検索結果.txt に書き込めません。開いているアプリを閉じてから、もう一度出力してください。"
        return
    } catch [System.UnauthorizedAccessException] {
        # 読み取り専用・書き込み権限が無いときは、アプリを閉じても直らないため別の文言にする
        setStatus "検索結果.txt に書き込む権限がありません（読み取り専用など）。$($workspace.ResultFile) を確かめてから、もう一度出力してください。"
        return
    }
    Invoke-Item -LiteralPath $workspace.ResultFile
    if ($rows.Count -lt $script:hitCount) {
        setStatus "絞り込み後の $($rows.Count.ToString('N0')) 件を検索結果.txt に出力しました"
    } else {
        setStatus "検索結果.txt に出力しました（$($rows.Count.ToString('N0')) 件）"
    }
}

$ui.ResultGrid.Add_MouseDoubleClick({
    param ($sender, $e)
    # 行の上でのダブルクリックだけを対象にする（列見出し・スクロールバーは除く）
    $element = $e.OriginalSource
    while ($element -and !($element -is [System.Windows.Controls.DataGridRow])) {
        if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
            return
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    # 見出しの行は、クリックで閉じる・開く（result_list.ps1）ので、ダブルクリックでは開かない
    if ($element -and !($element.Item -is [FileGroup])) {
        safe { openSource }
    }
})
$ui.ResultGrid.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        # 見出しの行では閉じる・開く。行では元のファイルを開く
        safe {
            if ($ui.ResultGrid.SelectedItem -is [FileGroup]) {
                toggleFileGroup $ui.ResultGrid.SelectedItem
            } else {
                openSource
            }
        }
        $e.Handled = $true
    } elseif ($e.Key -eq "C" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") {
        safe { copySelectedRows }
        $e.Handled = $true
    }
})
$ui.MenuOpen.Add_Click({ safe { openSource ${openModeNormal} } })
$ui.MenuOpenReadOnly.Add_Click({ safe { openSource ${openModeReadOnly} } })
$ui.MenuOpenNew.Add_Click({ safe { openSource ${openModeNew} } })
$ui.OpenModeCombo.Add_SelectionChanged({
    safe {
        # 起動時の読み込みでは保存しない（設定していない利用者の setting.config を作らないため）
        if (-not $script:loadingOpenMode) {
            writeOpenMode (getOpenMode)
        }
        updateOpenMenu
    }
})
$ui.MenuOpenFolder.Add_Click({ safe { openSourceFolder } })
$ui.OpenButton.Add_Click({ safe { openSource } })
$ui.OpenFolderButton.Add_Click({ safe { openSourceFolder } })

$ui.MenuCopy.Add_Click({ safe { copySelectedRows } })
$ui.MenuCopyPath.Add_Click({ safe { copySourcePath } })
$ui.ExportButton.Add_Click({ safe { exportResults } })
