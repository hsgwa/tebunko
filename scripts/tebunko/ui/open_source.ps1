# 検索結果から元のファイルを開く・パスをコピーする・結果を書き出す。


# ---- 元のファイルを開く・コピー・出力 ----


$script:openSourceRequest = [ref]0  # 元のファイルを確かめる依頼の番号。増やすたびに前の依頼の結果を捨てる。
# [ref] のまま閉じ込める（スクリプトブロックの中で $script:openSourceRequest を読むと、遠く離れたスレッド・
# タイマーから呼ばれたときに増やす前の値のまま固まって読めることがあるため。tests\tebunko\ui\open_source.Tests.ps1 で確かめている）
$script:openSourcePendingRow = [ref]$null  # 裏のスレッドに依頼を出したままの行（.Value が $null なら待っている依頼は無い）。
# $script:openSourceRequest と同じ理由で [ref] の箱にする。GetNewClosure() の中で
# 「$script:openSourcePendingRow = ...」と書き換えても、閉じ込めた側だけの変数になり元の箱には届かない
# （.Value を書き換えれば、同じ箱を見ている側すべてに届く）。
# 同じ行をもう一度開いても新しい依頼を出さない（届かない共有では、列の 2 つのスレッドが同じ行の依頼で塞がるため）
$script:openSourcePendingPath = [ref]""  # 待っている行を、確かめている間の文言に出すためのパス（もう一度開いたときに出し直す）

# パスのコピーの依頼は、[開く] とは別の番号・箱にする（同じ番号を共有すると、先に出した方の結果が黙って捨てられるため）
$script:copySourceRequest = [ref]0
$script:copySourcePendingRow = [ref]$null

function cancelPendingSourceLookup {
    # 待っている元のファイルの確認を打ち切る（新しく検索を始めたとき・ワークスペースを変えたときに呼ぶ）。
    # 依頼の番号を進めて前の依頼の結果を捨て、待っている行の記録・カーソルを戻す
    $script:openSourceRequest.Value++
    $script:openSourcePendingRow.Value = $null
    $script:copySourceRequest.Value++
    $script:copySourcePendingRow.Value = $null
    $window.Cursor = $null
}

function findSourceLocationAsync {
    # 検索結果の行の元のファイルの場所（getSourceLocation の結果）を求め、onLocation { param($location) } に渡す。
    # キャッシュ（$script:sourceFolderMaps）にあればその場で渡す。無ければ対応を読む必要があり、
    # ネットワークの場所なら裏の列（network）で読み、結果が届くまで画面のスレッドは待たない。ローカルならその場で読む
    #   requestBox・requestId: 呼ぶ側が進めた依頼の番号（待っている間に別の依頼が出たら、この結果は捨てる）
    #   pendingRowBox: 裏に依頼を出している間の行を入れる箱
    param (
        $row,
        [scriptblock]$onLocation,
        $requestBox,
        [int]$requestId,
        $pendingRowBox
    )

    $found = findSourceLocationInMaps $row $script:sourceFolderMaps
    if ($found.Resolved) {
        & $onLocation $found.Location
        return
    }
    if (!(testNetworkPath $found.Pending)) {
        & $onLocation (getSourceLocation $row $script:sourceFolderMaps)
        return
    }

    $pendingRowBox.Value = $row
    $script:openSourcePendingPath.Value = ""
    setStatus (getSourceLookingStatus)
    $window.Cursor = [System.Windows.Input.Cursors]::AppStarting
    # 呼ぶ関数は変数で捕まえてから閉じ込める（クロージャの中では、名前のままでは関数を引けないため）
    $applyLocation = ${function:applySourceLocation}
    $maps = $script:sourceFolderMaps
    $sendMaps = $maps.Clone()
    $statusPath = [string]$workspace.StatusFile
    $settingsPath = [string]${settingsFile}
    $indexDir = [string]$workspace.IndexDir
    startJob {
        param ($row, $maps, $statusPath, $settingsPath, $indexDir)
        readSourceLocation $row $maps $statusPath $settingsPath $indexDir
    } @($row, $sendMaps, $statusPath, $settingsPath, $indexDir) {
        param ($output, $errorText)
        if ($row -eq $pendingRowBox.Value) {
            $pendingRowBox.Value = $null
        }
        if ($requestId -ne $requestBox.Value) {
            # 待っている間に別の行を開いた・新しく検索した・ワークスペースを変えた。前の依頼は捨てる
            return
        }
        & $applyLocation $output $errorText $maps $onLocation
    }.GetNewClosure() (getWorkspaceJobQueue @($found.Pending, $indexDir))
}

function applySourceLocation {
    # findSourceLocationAsync の裏の仕事の結果を反映する。読めたら、読んだ対応をキャッシュ（maps）に足して onLocation を呼ぶ
    param (
        $output,
        $errorText,
        [hashtable]$maps,
        [scriptblock]$onLocation
    )

    $window.Cursor = $null
    if ($errorText -or !$output -or @($output).Count -eq 0) {
        setStatus (getSourceLookupFailedStatus ([string]$errorText))
        return
    }
    $result = @($output)[0]
    foreach ($key in @($result.Maps.Keys)) {
        $maps[$key] = $result.Maps[$key]
    }
    & $onLocation $result.Location
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
        $pendingPath = $script:openSourcePendingPath.Value
        setStatus $(if ($pendingPath) { getSourceCheckingStatus $pendingPath } else { getSourceLookingStatus })
        return
    }

    $requestBox = $script:openSourceRequest
    $requestBox.Value++
    $requestId = $requestBox.Value
    $continue = ${function:continueFindSourceFile}
    findSourceLocationAsync $row {
        param ($location)
        & $continue $row $location $requestBox $requestId $pendingRowBox $onFound
    }.GetNewClosure() $requestBox $requestId $pendingRowBox
}

function continueFindSourceFile {
    # findSourceFile の続き。元のファイルの場所（location）が分かったので、ファイルがあるかを確かめる
    param (
        $row,
        $location,
        $requestBox,
        [int]$requestId,
        $pendingRowBox,
        [scriptblock]$onFound
    )

    $book = $row.Book

    # 見つけたときの続きには、パスに加えてインデックスの名前も渡す（もらったインデックスかをキャッシュから引き直さないため）
    $foundInner = $onFound
    $locationName = [string]$location.Name
    $onFound = { param ($path) & $foundInner $path $locationName }.GetNewClosure()

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
    if ($requestId -ne $requestBox.Value) {
        # 場所を調べている間に別の行を開いた・新しく検索した・ワークスペースを変えた。前の依頼は捨てる（確認も出さない）
        return
    }
    if (testSourceNeedsConfirm $location.Name $location.Known (getConfirmedSourceNames).Confirmed) {
        # 設定に無い名前（content_index\ に手でコピーしたインデックス）の元のフォルダは、持ち主が書いたもの。
        # 元のフォルダに触れる前（有無の確認・ネットワークへの接続の前）に利用者に確かめる
        $answer = promptSourceConfirm $location $book
        $window.Cursor = $null
        if ($answer -eq "pick") {
            $path = joinSourcePath $location.Folder $location.Rest $book
            $relPath = if ($location.Rest) { "$($location.Rest)\$book" } else { $book }
            promptSourceMissing $location $book $path $relPath "" $onFound -pickFirst
            return
        }
        if ($answer -ne "use") {
            setStatus (getSourceConfirmCanceledStatus)
            return
        }
        setIndexSourceFolder $location.Name $location.Folder
        $script:sourceFolderMaps = @{}
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
    }.GetNewClosure() (getWorkspaceJobQueue $location.Folder)
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
    $answer = showConfirm -title "元のファイルが見つかりません" -heading $dialog.Heading `
        -facts @((factGone $dialog.Title $dialog.Detail), (factNext $dialog.Hint)) `
        -choices @(@{ Text = "フォルダを選ぶ"; Value = "pick" })
    if ($answer -eq "pick") {
        $path = joinSourcePath $location.Folder $location.Rest $book
        $relPath = if ($location.Rest) { "$($location.Rest)\$book" } else { $book }
        promptSourceMissing $location $book $path $relPath "" $onFound
    }
}

function promptSourceConfirm {
    # もらったインデックスの元のフォルダを、使う前に利用者に確かめる。"use"（このフォルダを使う）・"pick"（フォルダを選ぶ）、
    # キャンセルは $null を返す
    param (
        $location,
        [string]$book
    )

    $dialog = getSourceConfirmDialog $book $location.Name $location.Folder (testNetworkPath (normalizeFolderPath ([string]$location.Folder)))
    return (showConfirm -title "元のフォルダを確かめてください" -heading $dialog.Heading -hint $dialog.Hint `
        -facts @((factWarn $dialog.Title $dialog.Detail)) `
        -choices @(@{ Text = $dialog.PickText; Value = "pick" }, @{ Text = $dialog.UseText; Value = "use" }))
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
        [scriptblock]$onFound,
        [switch]$pickFirst   # 最初の「見つかりません」の確認を飛ばして、すぐフォルダ選択から始める
    )

    if ($location.Known) {
        $heading = "元のファイルが見つかりません。フォルダを移動した場合は、移動先のフォルダを選んでください。"
        $hint = "元の場所: $path"
        $description = "「$($location.Folder)」に当たるフォルダ（または $book のあるフォルダ）を選んでください"
    } else {
        $heading = "元のファイルが見つかりません。場所の記録もないため、$book のあるフォルダを選んでください。"
        $hint = "ファイル: $relPath"
        $description = "$book のあるフォルダ（またはインデックス [$($location.Name)] の元のフォルダ）を選んでください"
    }
    $facts = @()
    $skipConfirm = [bool]$pickFirst
    while ($true) {
        if (!$skipConfirm -and (showConfirm -title "元のファイルが見つかりません" -heading $heading -hint $hint -facts $facts `
                -choices @(@{ Text = "フォルダを選ぶ"; Value = "pick" })) -ne "pick") {
            setStatus (getSourceNotFoundStatus $path)
            return
        }
        $skipConfirm = $false
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
    # （noFallback のときは、そのまま開かずに $false を返す。もらったインデックスのマクロを持てる形式を、通常で開かないため）
    param (
        [string]$path,
        [string]$mode,
        [switch]$noFallback
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
    if ($noFallback) {
        return $false
    }
    Invoke-Item -LiteralPath $path
    return (-not $verb)
}

function openWithNotepad {
    # ファイルをメモ帳（固定のパス）で開く。.bat・.ps1・.js など、既定のアプリで開くと実行・登録になる拡張子はこちらを使う
    # （openWithShell・Invoke-Item は使わない。既定のアプリの登録がどうなっていても実行されない）。
    # メモ帳が無い（Windows 11 の「オプション機能」で外された等）・起動できないときは、既定のアプリには戻さずに $false を返す
    param (
        [string]$path
    )

    try {
        Start-Process -FilePath "$env:SystemRoot\System32\notepad.exe" -ArgumentList "`"${path}`""
        return $true
    } catch {
        return $false
    }
}

function setExcelAutomationSecurity {
    # Excel のマクロの扱い（AutomationSecurity）を設定し、前の値を返す。設定できなければ例外のまま外へ出す
    # （マクロが動く設定のまま開かないため。呼ぶ側が、ファイルを開くだけの道に切り替える）
    param (
        $excel,
        [int]$value
    )

    $previous = $excel.AutomationSecurity
    $excel.AutomationSecurity = $value
    return $previous
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

    # プログラムから開いたときの Excel の既定はマクロ有効のため、開く間だけ Excel の設定に従う（msoAutomationSecurityByUI = 2。
    # 既定ではマクロを止めて警告する）にし、開き終えたら元の値に戻す（使っていた Excel の設定を変えたままにしない）
    # 設定できなければ例外のまま外へ出し、開かない（マクロが動く設定のまま開かないため。呼ぶ側が、ファイルを開くだけの道に切り替える）
    $previousSecurity = setExcelAutomationSecurity $excel 2
    $book = $null
    try {
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
    } finally {
        try { $excel.AutomationSecurity = $previousSecurity } catch { }
    }
    $book.Activate()

    # 場所はシート名（ワークシート・グラフシートの両方）。名前が同じシートを選ぶ。
    # 以前の版のインデックスは、ファイル名に使えない文字を全角に置き換えてあるため、同じ名前のシートが無ければ
    # 全角に置き換えて一致するシートを選ぶ（`衝突"` と `衝突”` のように、置き換えると重なるシートがあるため、同じ名前を優先する）。
    # `Sheets` はワークシート・グラフシートを表示順のまま 1 つの列で回せるため、同じ名前の一致が
    # 全角に置き換えた一致より先に見つかるよう、両方の種類を 1 回のループで調べる（グラフシートはセルを選べない）。
    # グラフシートかどうかは `Charts` コレクションの名前で判定する（`Sheets` の要素の `Type` は、グラフシートでは
    # 期待どおりの値（xlChart）にならない。実機で確かめると、既定のグラフの種類の値が返ることがある）
    $chartNames = New-Object System.Collections.Generic.HashSet[string]
    foreach ($chart in $book.Charts) { [void]$chartNames.Add($chart.Name) }

    $target = $null
    $sameSafeName = $null
    $isChartSheet = $false
    $sameSafeIsChartSheet = $false
    foreach ($sheet in $book.Sheets) {
        $isChart = $chartNames.Contains($sheet.Name)
        if ($sheet.Name -eq $location) {
            $target = $sheet
            $isChartSheet = $isChart
            break
        }
        if ($null -eq $sameSafeName -and (toSafeFileName $sheet.Name) -eq $location) {
            $sameSafeName = $sheet
            $sameSafeIsChartSheet = $isChart
        }
    }
    if ($null -eq $target -and $null -ne $sameSafeName) {
        $target = $sameSafeName
        $isChartSheet = $sameSafeIsChartSheet
    }
    if ($null -ne $target) {
        $target.Activate()
        if ($cell -and -not $isChartSheet) {
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
    # ダブルクリック・Enter・［開く］での開き方（${openModes} のいずれか。選んでいなければ通常）
    if ($null -ne $script:openMode -and (${openModes} -contains $script:openMode)) {
        return [string]$script:openMode
    }
    return ${openModeNormal}
}

function setOpenMode {
    # 既定の開き方を設定の値に合わせる（起動時。選んだことにはしないため、設定は保存しない）
    param (
        [string]$mode
    )

    $script:openMode = if (${openModes} -contains $mode) { $mode } else { ${openModeNormal} }
}

function setContextMenuItems {
    # 右クリックメニューを、判断層が決めた並び（@{ Id; Header; Bold } の並び）で組み直す。
    # parts は Id → 項目（MenuItem）の表。区切りは Id が "separator"。ここでは並べるだけで、並び・文言・太字は決めない
    param (
        $menu,
        $items,
        $parts
    )

    $menu.Items.Clear()
    foreach ($item in $items) {
        if ($item.Id -eq "separator") {
            [void]$menu.Items.Add((New-Object System.Windows.Controls.Separator))
            continue
        }
        $menuItem = $parts[$item.Id]
        $menuItem.Header = $item.Header
        $menuItem.FontWeight = $(if ($item.Bold) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal })
        [void]$menu.Items.Add($menuItem)
    }
}

function selectResultRowForMenu {
    # 右クリックした行を選んでからメニューを出す（選んでいる別の行にメニューが効かないようにする）。
    # 行（ヒットした行）を複数選んでいて、その中の行を右クリックしたときだけ、選びを変えない（［選んだ行をコピー］のため）。
    # 見出しや、見出しを含む選びのときは、右クリックした行だけを選ぶ。メニューの種類も押したときの対象も、この選びで決まる
    param (
        $row
    )

    if (!(testResultMenuKeepSelection $row.IsSelected $row.Item @($ui.ResultGrid.SelectedItems))) {
        $ui.ResultGrid.SelectedItem = $row.Item
    }
}

function showResultMenu {
    # 結果の右クリックメニューを組み直す。出してよければ $true（行の無い所では $false）。
    # cursorLeft が負のときはキーボードで開いたとき（選んでいる行に出す）
    param (
        $source,
        [double]$cursorLeft
    )

    if ($cursorLeft -ge 0 -and $null -eq (getDataGridRowAt $source)) {
        return $false
    }
    $item = $ui.ResultGrid.SelectedItem
    if ($null -eq $item) {
        return $false
    }

    $context = getResultMenuContext $item
    $parts = @{
        openNormal = $ui.MenuOpen; openNew = $ui.MenuOpenNew; openReadOnly = $ui.MenuOpenReadOnly; openFolder = $ui.MenuOpenFolder
        copyRows = $ui.MenuCopy; copyPath = $ui.MenuCopyPath; toggleGroup = $ui.MenuToggleGroup
    }
    setContextMenuItems $ui.ResultMenu (getResultMenuItems $context.Target (getOpenMode) $context.Expanded) $parts
    return $true
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
        param ($path, $name)
        & $openFound $path $row $mode $name
    }.GetNewClosure()
}

function openFoundSource {
    # findSourceFile が見つけたファイルを開く（openSource の続き）
    param (
        [string]$path,
        $row,
        [string]$mode,
        [string]$indexName = ""   # 元のファイルを見つけたインデックスの名前（findSourceFile が渡す）
    )

    # もらったインデックス（このワークスペースで自分が作ったのではないもの）のマクロを持てる形式は、読み取り専用で開く
    $received = testSourceReceived $indexName @((getConfirmedSourceNames).Crawled)
    $openMode = getSourceOpenMode $path $mode $received
    $mode = $openMode.Mode
    $strict = $openMode.Strict
    $notice = if ($openMode.Notice) { "（$($openMode.Notice)）" } else { "" }

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
            setStatus "${how}：${path}${notice}"
        } catch {
            # Excel を操作できない場合（ダイアログを表示中など）は、ファイルを開くだけにする
            if (openWithShell $path $mode -noFallback:$strict) {
                setStatus "${how}（該当セルへの移動はできませんでした）：${path}${notice}"
            } elseif ($strict) {
                setStatus (getSourceReadOnlyFailedStatus $path)
            } else {
                setStatus "${fallback}（該当セルへの移動はできませんでした）：${path}"
            }
        } finally {
            $window.Cursor = $null
        }
    } elseif ($row.IsText) {
        # テキストは既定のアプリに開き方（読み取り専用・新規）の動詞が無いことが多く、毎回「開けなかったため…」と出るのを避けるため、
        # 動詞を試さずにそのまま開く（openWithShell に通常の開き方を渡すと動詞を試さない）。行への移動はしない。
        # .bat・.ps1・.js など、既定のアプリで開くと実行・登録になる拡張子はメモ帳で開く
        if (testTextOpenWithNotepad $path) {
            if (openWithNotepad $path) {
                setStatus "メモ帳で開きました：${path}"
            } else {
                setStatus "メモ帳で開けませんでした：${path}"
            }
        } else {
            [void](openWithShell $path ${openModeNormal})
            setStatus "開きました：${path}"
        }
    } elseif (openWithShell $path $mode -noFallback:$strict) {
        setStatus "${how}：${path}${notice}"
    } elseif ($strict) {
        setStatus (getSourceReadOnlyFailedStatus $path)
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
    # 元のファイルのパスをクリップボードに写す。元の場所の対応を読む必要があるときは、ネットワークの場所なら裏で読む
    $row = getCurrentHitRow
    if ($null -eq $row) {
        return
    }
    $requestBox = $script:copySourceRequest
    $requestBox.Value++
    $requestId = $requestBox.Value
    $pendingRowBox = $script:copySourcePendingRow
    $complete = ${function:completeCopySourcePath}
    findSourceLocationAsync $row {
        param ($location)
        & $complete $row $location
    }.GetNewClosure() $requestBox $requestId $pendingRowBox
}

function setClipboardText {
    param (
        [string]$text
    )

    [System.Windows.Clipboard]::SetText($text)
}

function completeCopySourcePath {
    # copySourcePath の続き。場所が分かったので、パスを写す
    param (
        $row,
        $location
    )

    $path = $null
    if ($location.Known) {
        $path = joinSourcePath $location.Folder $location.Rest $row.Book
    }
    if (!$path) {
        # 元のファイルの場所が分からない（インデックスだけを別の PC にコピーした等）ときは、インデックスの中の位置（フォルダ\元のファイル名）を写す
        $path = if ($row.RelDir) { "$($row.RelDir)\$($row.Book)" } else { $row.Book }
    }
    setClipboardText $path
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
        setStatus "search_results.txt に書き込めません。開いているアプリを閉じてから、もう一度出力してください。"
        return
    } catch [System.UnauthorizedAccessException] {
        # 読み取り専用・書き込み権限が無いときは、アプリを閉じても直らないため別の文言にする
        setStatus "search_results.txt に書き込む権限がありません（読み取り専用など）。$($workspace.ResultFile) を確かめてから、もう一度出力してください。"
        return
    }
    Invoke-Item -LiteralPath $workspace.ResultFile
    if ($rows.Count -lt $script:hitCount) {
        setStatus "絞り込み後の $($rows.Count.ToString('N0')) 件を search_results.txt に出力しました"
    } else {
        setStatus "search_results.txt に出力しました（$($rows.Count.ToString('N0')) 件）"
    }
}

$ui.ResultGrid.Add_MouseDoubleClick({
    param ($sender, $e)
    # 行の上でのダブルクリックだけを対象にする（列見出し・スクロールバーは除く）。
    # 見出しの行は、クリックで閉じる・開く（result_list.ps1）ので、ダブルクリックでは開かない
    $row = getDataGridRowAt $e.OriginalSource
    if ($row -and !($row.Item -is [FileGroup])) {
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
# 右クリックした行を選び、行の無い所ではメニューを出さない（インデックスの一覧の右クリックと同じ）
$ui.ResultGrid.Add_PreviewMouseRightButtonDown({
    param ($sender, $e)
    safe {
        $row = getDataGridRowAt $e.OriginalSource
        if ($row) {
            selectResultRowForMenu $row
        }
    }
})
$ui.ResultGrid.Add_ContextMenuOpening({
    param ($sender, $e)
    safe {
        if (!(showResultMenu $e.OriginalSource $e.CursorLeft)) {
            $e.Handled = $true
        }
    }
})
$ui.MenuOpen.Add_Click({ safe { openSource ${openModeNormal} } })
$ui.MenuOpenReadOnly.Add_Click({ safe { openSource ${openModeReadOnly} } })
$ui.MenuOpenNew.Add_Click({ safe { openSource ${openModeNew} } })
$ui.MenuOpenFolder.Add_Click({ safe { openSourceFolder } })
$ui.OpenButton.Add_Click({ safe { openSource } })
# ［開く ▾］のメニュー。選んだ開き方は次からの既定（ダブルクリック・Enter・［開く］）にもなり、そのまま開く
function selectOpenMode {
    param (
        [string]$mode
    )

    if (${openModes} -contains $mode) {
        $script:openMode = $mode
        writeOpenMode $mode
    }
    openSource $mode
}
$ui.OpenMenuButton.Add_Click({
    $menu = $ui.OpenMenuButton.ContextMenu
    $menu.PlacementTarget = $ui.OpenMenuButton
    $menu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
    $menu.IsOpen = $true
})
$ui.MenuOpenModeNormal.Add_Click({ safe { selectOpenMode ${openModeNormal} } })
$ui.MenuOpenModeNew.Add_Click({ safe { selectOpenMode ${openModeNew} } })
$ui.MenuOpenModeReadOnly.Add_Click({ safe { selectOpenMode ${openModeReadOnly} } })
$ui.OpenFolderButton.Add_Click({ safe { openSourceFolder } })

$ui.MenuToggleGroup.Add_Click({
    safe {
        if ($ui.ResultGrid.SelectedItem -is [FileGroup]) {
            toggleFileGroup $ui.ResultGrid.SelectedItem
        }
    }
})
$ui.MenuCopy.Add_Click({ safe { copySelectedRows } })
$ui.MenuCopyPath.Add_Click({ safe { copySourcePath } })
$ui.ExportButton.Add_Click({ safe { exportResults } })
