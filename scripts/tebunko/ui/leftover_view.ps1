# 起動時の「前回残った Office の確認」に出す文言と、出すかどうかの判断（判断層）。
# 材料は getOfficeProcesses（shared/office/office_process.ps1）の結果。画面に触らないため、そのままテストできる
# （tests/tebunko/ui/leftover_view.Tests.ps1）。止める処理は stopOfficeProcesses だけが行う。
# 止めてよいのは、このツールが起動した記録があり（Owned）、起動した画面・インデックス作成がもう無く（Owner が Gone）、
# 窓を持たない（HasWindow でない）Office だけ。利用者が開いている Office は、記録が無いか窓があるため、対象にならない。

function selectLeftoverProcesses {
    # getOfficeProcesses の結果から、残り物（確認して止めてよいもの）だけを返す
    param (
        [object[]]$processes
    )

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($process in @($processes)) {
        if ($process.Owned -and $process.Owner -eq "Gone" -and !$process.HasWindow) {
            $result.Add($process)
        }
    }
    return $result.ToArray()
}

function getLeftoverSummary {
    # 「Excel 2 件・Word 1 件」。Excel・Word・PowerPoint の順で、0 件のアプリは出さない
    param (
        [object[]]$targets
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($name in ${officeProcessNames}.Keys) {
        $count = @($targets | Where-Object { $_.ProcessName -eq $name }).Count
        if ($count -gt 0) {
            $parts.Add("$(${officeProcessNames}[$name]) $count 件")
        }
    }
    return ($parts -join "・")
}

function getLeftoverPrompt {
    # 確認の内容を返す。残り物が無ければ $null。
    #   Targets: 残り物（getOfficeProcesses の要素）  Count: 件数  Summary: 「Excel 2 件・Word 1 件」
    #   Title・Heading・Facts（@{ Kind = "Kept"; Title; Detail } の 2 行）・Hint・Choices（@{ Text; Value }）・CancelText
    #   Details: 一覧の行 @{ AppName; Id; StartTime; TimeText }（StartTime は読めなければ $null。TimeText は「10/04 18:32」。読めなければ「不明」）
    param (
        [object[]]$processes
    )

    $targets = @(selectLeftoverProcesses $processes)
    if ($targets.Count -eq 0) {
        return $null
    }
    $summary = getLeftoverSummary $targets
    $details = New-Object System.Collections.Generic.List[object]
    foreach ($target in $targets) {
        $details.Add(@{ AppName = $target.AppName; Id = $target.Id; StartTime = $target.StartTime; TimeText = (getLeftoverTimeText $target.StartTime) })
    }
    return @{
        Targets    = $targets
        Count      = $targets.Count
        Summary    = $summary
        Title      = "Office の終了"
        Heading    = "前回のインデックス作成で起動した Office が $($targets.Count) 件、残ったまま動いています。終了しますか？"
        Facts      = @(
            @{ Kind = "Kept"; Title = "編集中のファイルは閉じません"; Detail = "終了するのは、tebunko がバックグラウンドで起動した ${summary}だけです" }
            @{ Kind = "Kept"; Title = "保存していない作業が失われることはありません"; Detail = "残っているのは、インデックス作成のために起動したものです" }
        )
        Hint       = "［今回は終了しない］を選んだときは、次に起動したときにもう一度お聞きします。"
        Choices    = @(@{ Text = "終了する"; Value = "stop" })
        CancelText = "今回は終了しない"
        Details    = $details.ToArray()
    }
}

function getLeftoverTimeText {
    # 一覧の「起動した時刻」。「10/04 18:32」。読めなければ「不明」
    param (
        $startTime
    )

    if ($null -eq $startTime) {
        return "不明"
    }
    try {
        return ([datetime]$startTime).ToString("MM/dd HH:mm")
    } catch {
        return "不明"
    }
}

function getLeftoverDetailToggleText {
    # 一覧の開け閉めのボタンの文字
    param (
        [bool]$open
    )

    return $(if ($open) { "詳細を隠す" } else { "詳細を表示" })
}

function getLeftoverTargets {
    # 確認を出してから止めるまでの間に状況が変わっていないかを照らし直す。
    # shownIds: 確認に出した PID。fresh: 読み直した getOfficeProcesses の結果。
    # 今も残り物であるもののうち、確認に出した PID だけを返す（確認の後に増えたものは、確認していないので止めない）
    param (
        [int[]]$shownIds,
        [object[]]$fresh
    )

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($target in @(selectLeftoverProcesses $fresh)) {
        if (@($shownIds) -contains [int]$target.Id) {
            $result.Add($target)
        }
    }
    return $result.ToArray()
}

function getLeftoverResultText {
    # stopOfficeProcesses の結果（@{ Id; Stopped; Status; Reason }）から、ステータスバーに出す文を作る。
    # 何も止めておらず言うことも無ければ $null（ステータスは変えない）。もう無かったもの（Gone）は数にも理由にも入れない
    param (
        [object[]]$results
    )

    $stopped = 0
    $changed = New-Object System.Collections.Generic.List[string]
    $failed = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($results)) {
        switch ($item.Status) {
            "Stopped" { $stopped++ }
            "Changed" { $changed.Add([string]$item.Id) }
            "Failed" { $failed.Add("PID $($item.Id) を終了できませんでした：$($item.Reason)") }
        }
    }
    $text = ""
    if ($stopped -gt 0) {
        $text += "Office を $stopped 件終了しました。"
    }
    if ($changed.Count -gt 0) {
        $text += "PID $($changed -join "・") は確認の後に別のプロセスに変わったため、終了しませんでした。"
    }
    if ($failed.Count -gt 0) {
        $text += ($failed -join "")
    }
    if ($text -eq "") {
        return $null
    }
    return $text
}

function getLeftoverPromptTiming {
    # 起動時の確認を、今出すか・見送るか・あとで出し直すかを決める。上から見て最初に当てはまったもの。
    #   残り物が 0 件 → Skip
    #   ウィンドウを閉じている途中 → Skip
    #   インデックス作成中（そのインデックス作成の Office を止めないため）→ Skip
    #   起動時のお知らせがまだ閉じていない → Defer（閉じたあとに、もう一度判断する）
    #   ほかのダイアログが開いている → Skip
    #   それ以外 → Show
    param (
        [int]$leftoverCount,
        [bool]$closing,
        [bool]$indexing,
        [bool]$noticesOpen,
        [int]$otherDialogCount
    )

    if ($leftoverCount -le 0 -or $closing -or $indexing) {
        return "Skip"
    }
    if ($noticesOpen) {
        return "Defer"
    }
    if ($otherDialogCount -ge 1) {
        return "Skip"
    }
    return "Show"
}

function getOfficePidQueue {
    # 記録の置き場所を読む startJob の列。共有に届かないと待たされるため、ネットワークの場所なら専用の列（"network"）
    param (
        [string]$dir
    )

    return $(if (testNetworkPath $dir) { "network" } else { "default" })
}

# ---- 写真・画面のテスト用の偽の行 ----
# 写真と画面のテストは、本物の Office を残さずに確認の画面を出すため、場面（環境変数 TEBUNKO_GUI_LEFTOVER_FILE が指す JSON）から
# 偽の行を差し込む。環境変数が無い（ふだんの起動）ときは、必ず本物（getOfficeProcesses・stopOfficeProcesses）を使う。
# 偽の行のときは、終了の処理も偽（結果を作るだけで、プロセスには触れない）。tests/meta/leftover_seam.Tests.ps1 が、この決まりを確かめる。

function getLeftoverSourceMode {
    # 偽の行のファイルが指定されていれば "Fake"、なければ "Real"
    param (
        [string]$fakeFile
    )

    return $(if ([string]::IsNullOrEmpty($fakeFile)) { "Real" } else { "Fake" })
}

function convertLeftoverFakeRows {
    # 偽の行（@{ Id; ProcessName; StartTime; StopStatus }）を getOfficeProcesses の要素の形にする。残り物として扱う（記録あり・持ち主なし・窓なし）。
    # StopStatus は、終了の結果に使う（Stopped・Changed・Gone・Failed。省けば Stopped）
    param (
        [object[]]$rows
    )

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($rows)) {
        $name = $(if ($row.ProcessName) { [string]$row.ProcessName } else { "EXCEL" })
        $start = $null
        if ($row.StartTime) {
            try { $start = [datetime]$row.StartTime } catch { $start = $null }
        }
        $result.Add([pscustomobject]@{
            Id          = [int]$row.Id
            ProcessName = $name
            AppName     = ${officeProcessNames}[$name]
            StartTime   = $start
            HasWindow   = $false
            Owned       = $true
            Owner       = "Gone"
            StopStatus  = $(if ($row.StopStatus) { [string]$row.StopStatus } else { "Stopped" })
        })
    }
    return $result.ToArray()
}

function getLeftoverFakeStopResults {
    # 偽の行の「終了」の結果。stopOfficeProcesses と同じ形（@{ Id; Stopped; Status; Reason }）。プロセスには触れない
    param (
        [object[]]$targets
    )

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($target in @($targets)) {
        $status = $(if ($target.StopStatus) { [string]$target.StopStatus } else { "Stopped" })
        $reason = ""
        if ($status -eq "Changed") {
            $reason = "PID $($target.Id) は確認の後に別のプロセスに変わったため、終了しませんでした。"
        } elseif ($status -eq "Failed") {
            $reason = "アクセスが拒否されました"
        }
        $results.Add(@{ Id = [int]$target.Id; Stopped = ($status -eq "Stopped"); Status = $status; Reason = $reason })
    }
    return $results.ToArray()
}
