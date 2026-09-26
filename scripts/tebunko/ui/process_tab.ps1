# ［9 プロセス停止］タブ（実行中の Office の一覧と強制終了）。

# ============================================================================
# ［9 プロセス停止］
# ============================================================================

$script:processes = @()

function refreshProcesses {
    # 実行中の Office の一覧を取り直して表示する（すぐに一覧が要るとき。ボタンの操作など）
    applyProcesses @(getOfficeProcesses)
}

function refreshProcessesInBackground {
    # 一覧の取り直しを別スレッドで行う（タブを開いている間の定期的な更新。画面のスレッドでプロセスを調べない）
    if ($script:processRefreshing) {
        return
    }
    $script:processRefreshing = $true
    startJob {
        , @(getOfficeProcesses)
    } @() {
        param ($output, $errorText)
        $script:processRefreshing = $false
        if (!$errorText -and $output -and $output.Count -gt 0) {
            applyProcesses @($output[0])
        }
    }
}

function applyProcesses {
    param (
        [object[]]$list
    )

    $selectedIds = @($ui.ProcessGrid.SelectedItems | ForEach-Object { $_.Id })
    $script:processes = @($list | Sort-Object @{ Expression = { !$_.Background } }, @{ Expression = { $_.StartTime } })

    $rows = New-Object System.Collections.ArrayList
    foreach ($process in $script:processes) {
        $row = New-Object ProcRow
        $row.Id = $process.Id
        $row.AppName = $process.AppName
        $row.Background = $process.Background
        $row.StartText = formatTime $process.StartTime
        $row.MemoryText = "$($process.MemoryMB.ToString('N0')) MB"
        $row.TitleText = if ($process.Title) { $process.Title } else { "（なし）" }
        [void]$rows.Add($row)
    }
    $ui.ProcessGrid.ItemsSource = $rows
    foreach ($row in $rows) {
        if ($selectedIds -contains $row.Id) {
            [void]$ui.ProcessGrid.SelectedItems.Add($row)
        }
    }

    $background = @($script:processes | Where-Object { $_.Background }).Count
    $visible = $script:processes.Count - $background
    if ($script:processes.Count -eq 0) {
        $ui.ProcessSummaryText.Text = "実行中の Excel・Word・PowerPoint はありません。（$(Get-Date -Format 'H:mm:ss') 時点）"
    } else {
        $ui.ProcessSummaryText.Text = "$($script:processes.Count) 件（バックグラウンド $background 件・画面に表示中 $visible 件） ・ $(Get-Date -Format 'H:mm:ss') 時点"
    }
    $ui.KillBackgroundButton.Content = "バックグラウンドのみ終了（$background 件）"
    $ui.KillBackgroundButton.IsEnabled = $background -gt 0
    $ui.KillAllButton.IsEnabled = $script:processes.Count -gt 0
    $ui.KillSelectedButton.IsEnabled = $ui.ProcessGrid.SelectedItems.Count -gt 0
    updateKillBadge $background
}

function updateKillBadge {
    param (
        $background = $null
    )

    if ($null -eq $background) {
        # 数えるのは別スレッドで行い、数が分かったら印を付け直す（ウィンドウを前面にするたびに呼ばれるため、画面のスレッドで調べない）
        startJob {
            @(getOfficeProcesses | Where-Object { $_.Background }).Count
        } @() {
            param ($output, $errorText)
            if (!$errorText -and $output -and $output.Count -gt 0) {
                updateKillBadge ([int]$output[0])
            }
        }
        return
    }
    # インデックス作成中はバックグラウンドの Excel 等があって当然なので、印を付けない
    $ui.KillTabHeader.Text = if ($background -gt 0 -and !(isIndexing)) { "⚠ 9 プロセス停止" } else { "9 プロセス停止" }
}

function killProcesses {
    param (
        [object[]]$targets
    )

    if ($targets.Count -eq 0) {
        return
    }

    $describe = {
        param ([object[]]$list)
        @(${officeProcessNames}.Values | ForEach-Object {
            $name = $_
            $count = @($list | Where-Object { $_.AppName -eq $name }).Count
            if ($count -gt 0) { "$name $count 件" }
        }) -join "・"
    }
    $visibleTargets = @($targets | Where-Object { !$_.Background })
    if ($visibleTargets.Count -gt 0) {
        $heading = "開いたままの Office を $($targets.Count) 件、強制的に終了しますか？"
        $facts = @(
            (factGone "保存していない内容は失われます" "画面に出ているもの：$(& $describe $visibleTargets)"),
            (factKept "ファイル自体は消えません" "保存し忘れがないか、先に画面で確かめてください")
        )
    } else {
        $heading = "バックグラウンドの Office を $($targets.Count) 件終了しますか？"
        $facts = @(
            (factKept "画面に出ているファイルはありません" (& $describe $targets)),
            (factNext "残ったまま動いていたものを片付けます" "次のインデックス作成で作り直されます")
        )
    }
    if (isIndexing) {
        $facts += factGone "いま取り込み中のファイルは失敗あつかいになります" "インデックス作成が終わってから終了するのが安全です"
    }
    $answer = showConfirm -heading $heading -facts $facts `
        -choices @(@{ Text = "終了する"; Value = "stop"; Danger = ($visibleTargets.Count -gt 0 -or (isIndexing)) })
    if ($answer -ne "stop") {
        return
    }

    $results = @(stopOfficeProcesses @($targets | ForEach-Object { $_.Id }))
    $stopped = @($results | Where-Object { $_.Stopped }).Count
    $failures = @($results | Where-Object { !$_.Stopped } | ForEach-Object { "PID $($_.Id) を終了できませんでした：$($_.Message)" })
    setStatus (@("$stopped 個のプロセスを終了しました。") + $failures -join "　")
    Start-Sleep -Milliseconds 300
    refreshProcesses
}

$script:processRefreshing = $false
$script:processTimer = newTimer 5000 { safe { refreshProcessesInBackground } }

$ui.RefreshProcessButton.Add_Click({ safe { refreshProcesses } })
$ui.ProcessGrid.Add_SelectionChanged({ $ui.KillSelectedButton.IsEnabled = $ui.ProcessGrid.SelectedItems.Count -gt 0 })
$ui.KillBackgroundButton.Add_Click({ safe { refreshProcesses; killProcesses @($script:processes | Where-Object { $_.Background }) } })
$ui.KillSelectedButton.Add_Click({
    safe {
        $ids = @($ui.ProcessGrid.SelectedItems | ForEach-Object { $_.Id })
        killProcesses @($script:processes | Where-Object { $ids -contains $_.Id })
    }
})
$ui.KillAllButton.Add_Click({ safe { refreshProcesses; killProcesses $script:processes } })
