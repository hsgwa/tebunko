# 起動の時間の計測（tools\perf\measure_startup.ps1）の集計。起動とは分けてあり、単体テストで確かめる（tests\tools\perf_startup.Tests.ps1）。
#
# 入力は、画面の起動口が書く起動の記録（ワークスペース直下のstartup_trace.tsv。scripts\tebunko\ui\startup_trace.ps1）の行。
# 1 行が「段階の名前<TAB>PowerShell のプロセスが始まってからの経過ミリ秒」。段階の時間は、隣り合う行の差で求める。
# 画面が出る（A）・一覧が出る（B）は、記録の中の決まった名前の行（プロセスが始まってからの経過ミリ秒のまま使う）。

. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))

# 終点の段階の名前（scripts\tebunko\ui\shell\window_events.ps1・startup.ps1 の markStartup と同じ名前）
${startupMarkWindow} = "画面が出る"
${startupMarkList} = "一覧が出る"

function readStartupTrace {
    # 記録の行を読む: @{ Rows（Name・Ms の並び）; Problems（飛ばした行の理由の並び） }
    #   壊れた行（2 つに分かれない・名前が空・数でない・負の数）と、時刻が前の行より戻っている行は飛ばして、理由を Problems に入れる
    param (
        [string[]]$lines
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $problems = New-Object System.Collections.Generic.List[string]
    $previous = -1.0
    $number = 0
    foreach ($line in @($lines)) {
        $number++
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $parts = $line.Split("`t")
        $ms = 0.0
        if ($parts.Count -ne 2 -or $parts[0].Trim() -eq "" -or
            !([double]::TryParse($parts[1], [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ms)) -or $ms -lt 0) {
            $problems.Add("${number} 行目を飛ばした（壊れた行）: $line")
            continue
        }
        if ($ms -lt $previous) {
            $problems.Add("${number} 行目を飛ばした（時刻が前の行より戻っている）: $line")
            continue
        }
        $previous = $ms
        $rows.Add([ordered]@{ Name = $parts[0].Trim(); Ms = $ms })
    }
    return @{ Rows = $rows.ToArray(); Problems = $problems.ToArray() }
}

function getStartupRunTimes {
    # 1 回の起動の記録（readStartupTrace の Rows）から、段階ごとの時間と終点の時刻を求める:
    #   @{ Stages（Name・Ms の並び。最初の段階はプロセスの始まりから、あとは前の行との差）; WindowMs（画面が出る・累計）; ListMs（一覧が出る・累計）}
    #   終点の行が無いときは、その値は $null
    param (
        [object[]]$rows
    )

    $stages = New-Object System.Collections.Generic.List[object]
    $previous = 0.0
    $windowMs = $null
    $listMs = $null
    foreach ($row in @($rows)) {
        $stages.Add([ordered]@{ Name = $row.Name; Ms = $row.Ms - $previous })
        $previous = $row.Ms
        if ($row.Name -eq ${startupMarkWindow}) { $windowMs = $row.Ms }
        if ($row.Name -eq ${startupMarkList}) { $listMs = $row.Ms }
    }
    return @{ Stages = $stages.ToArray(); WindowMs = $windowMs; ListMs = $listMs }
}

function getStartupSummary {
    # 何回かの起動（getStartupRunTimes の結果に、外から測った ExternalMs を足したもの）を、段階ごとに集計する:
    #   @{ Stages（Name と Stats の並び。段階は最初に現れた順）; Window・List・External（getStats の結果。無ければ $null）; Runs（回数） }
    param (
        [object[]]$runs
    )

    $order = New-Object System.Collections.Generic.List[string]
    $values = @{}
    $windowValues = New-Object System.Collections.Generic.List[double]
    $listValues = New-Object System.Collections.Generic.List[double]
    $externalValues = New-Object System.Collections.Generic.List[double]
    foreach ($run in @($runs)) {
        foreach ($stage in @($run.Stages)) {
            if (!$values.ContainsKey($stage.Name)) {
                $values[$stage.Name] = New-Object System.Collections.Generic.List[double]
                $order.Add($stage.Name)
            }
            $values[$stage.Name].Add([double]$stage.Ms)
        }
        if ($null -ne $run.WindowMs) { $windowValues.Add([double]$run.WindowMs) }
        if ($null -ne $run.ListMs) { $listValues.Add([double]$run.ListMs) }
        if ($null -ne $run.ExternalMs) { $externalValues.Add([double]$run.ExternalMs) }
    }
    $stages = @($order | ForEach-Object { [ordered]@{ Name = $_; Stats = (getStats $values[$_].ToArray()) } })
    return @{
        Stages = $stages
        Window = (getStats $windowValues.ToArray())
        List = (getStats $listValues.ToArray())
        External = (getStats $externalValues.ToArray())
        Runs = @($runs).Count
    }
}

function formatStartupMs {
    # ミリ秒を整数で書く（カルチャに左右されない）。値が無ければ -
    param (
        $value
    )

    if ($null -eq $value) {
        return "-"
    }
    return ([long][Math]::Round([double]$value)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
}

function formatStartupStats {
    # 最小 / 中央値 / 最大 を 1 つの欄に書く
    param (
        $stats
    )

    if ($null -eq $stats) {
        return "-"
    }
    return "{0} / {1} / {2}" -f (formatStartupMs $stats.Min), (formatStartupMs $stats.Median), (formatStartupMs $stats.Max)
}

function formatStartupDiff {
    # 後の中央値 - 前の中央値（ミリ秒）と割合。どちらかが無ければ -
    param (
        $before,
        $after
    )

    if ($null -eq $before -or $null -eq $after) {
        return "-"
    }
    $diff = [double]$after.Median - [double]$before.Median
    $sign = if ($diff -gt 0) { "+" } else { "" }
    $text = "${sign}$(formatStartupMs $diff)"
    if ([double]$before.Median -ne 0) {
        $percent = [Math]::Round($diff * 100 / [double]$before.Median, 1)
        $text += " ({0}{1}%)" -f $sign, $percent.ToString("0.#", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $text
}

function getStartupTableRows {
    # 表にする行（Label・Before・After の Stats）を、段階 → 終点の順に並べる。前（before）が無いときは Before は $null
    param (
        $after,
        $before = $null
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($stage in @($after.Stages)) { if (!$names.Contains($stage.Name)) { $names.Add($stage.Name) } }
    if ($null -ne $before) {
        foreach ($stage in @($before.Stages)) { if (!$names.Contains($stage.Name)) { $names.Add($stage.Name) } }
    }
    foreach ($name in $names) {
        $afterStats = @($after.Stages | Where-Object { $_.Name -eq $name } | ForEach-Object { $_.Stats }) | Select-Object -First 1
        $beforeStats = $null
        if ($null -ne $before) {
            $beforeStats = @($before.Stages | Where-Object { $_.Name -eq $name } | ForEach-Object { $_.Stats }) | Select-Object -First 1
        }
        $rows.Add([ordered]@{ Label = $name; Before = $beforeStats; After = $afterStats })
    }
    $rows.Add([ordered]@{ Label = "**(A) 画面が出る（プロセスの始まりから）**"; Before = $(if ($null -ne $before) { $before.Window } else { $null }); After = $after.Window })
    $rows.Add([ordered]@{ Label = "**(B) 一覧が出る（プロセスの始まりから）**"; Before = $(if ($null -ne $before) { $before.List } else { $null }); After = $after.List })
    $rows.Add([ordered]@{ Label = "外から測った、tebunko.bat を起動してから本体の窓が出るまで"; Before = $(if ($null -ne $before) { $before.External } else { $null }); After = $after.External })
    return $rows.ToArray()
}

function formatStartupTable {
    # 集計（getStartupSummary）を Markdown の表の行にする。値は ミリ秒の「最小 / 中央値 / 最大」。
    # before を渡すと、前と後を並べ、中央値の差（と割合）の欄を付ける
    param (
        $after,
        $before = $null
    )

    $lines = New-Object System.Collections.Generic.List[string]
    if ($null -eq $before) {
        $lines.Add("| 段階 | 最小 / 中央値 / 最大（ms） |")
        $lines.Add("|---|---|")
        foreach ($row in @(getStartupTableRows $after)) {
            $lines.Add(("| {0} | {1} |" -f $row.Label, (formatStartupStats $row.After)))
        }
        return $lines.ToArray()
    }
    $lines.Add("| 段階 | 前: 最小 / 中央値 / 最大（ms） | 後: 最小 / 中央値 / 最大（ms） | 中央値の差 |")
    $lines.Add("|---|---|---|---|")
    foreach ($row in @(getStartupTableRows $after $before)) {
        $lines.Add(("| {0} | {1} | {2} | {3} |" -f $row.Label, (formatStartupStats $row.Before), (formatStartupStats $row.After), (formatStartupDiff $row.Before $row.After)))
    }
    return $lines.ToArray()
}
