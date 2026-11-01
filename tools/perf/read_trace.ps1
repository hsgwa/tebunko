# 処理の記録（shared\core\trace.ps1 が書いた trace\*.tsv）を読んで、区間ごとの時間を集計する。
#
# どの区分の記録にも使える。区切りの名前を定義（Defs）で渡せば、その順・その前後関係で表にする。定義が無ければ、記録に出てきた名前から作る。
#   .\tools\perf\read_trace.ps1 -Dir <ワークスペース>\trace -Category preview
# 他のスクリプト（measure_startup.ps1）は、このファイルを dot-source して関数を使う。

param (
    # 記録のフォルダ（ワークスペースの trace\）。省略すると関数だけを読み込む
    [string]$Dir,
    [string]$Category = "startup",
    # 指定した PID のプロセスの記録だけ（省略すると全部）
    [int]$ProcessId = 0
)

. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))

function readTraceLines {
    # TSV の行を読む。書いている最中の画面が開いたままでも読めるよう、共有を許して開く。開けなければ空
    param (
        [string]$Path
    )

    try {
        $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            return @($reader.ReadToEnd() -split "\r?\n")
        } finally {
            $stream.Dispose()
        }
    } catch {
        return @()
    }
}

function readTraceRows {
    # フォルダの TSV から、指定の区分の行を読む。ファイル名は <起点の日時>_<PID>_<スレッド>.tsv。
    # 戻り値は @{ Run = "<起点の日時>_<PID>"; Process = <PID>; Thread; ElapsedMs; Category; Name } の並び
    param (
        [string]$Dir,
        [string]$Category,
        [int]$ProcessId = 0
    )

    $rows = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Dir)) {
        return $rows.ToArray()
    }
    foreach ($file in @(Get-ChildItem -LiteralPath $Dir -Filter "*.tsv" | Sort-Object Name)) {
        if ($file.BaseName -notmatch '^(?<start>[0-9]{8}-[0-9]{6})_(?<pid>[0-9]+)_(?<thread>[0-9]+)$') {
            continue
        }
        if ($ProcessId -gt 0 -and [int]$Matches.pid -ne $ProcessId) {
            continue
        }
        $run = "$($Matches.start)_$($Matches.pid)"
        $processId = [int]$Matches.pid
        foreach ($line in @(readTraceLines $file.FullName)) {
            $cells = $line -split "`t"
            $elapsed = 0.0
            if ($cells.Count -lt 5 -or -not [double]::TryParse($cells[1], [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$elapsed)) {
                continue
            }
            if ($Category -and $cells[2] -ne $Category) {
                continue
            }
            $rows.Add([pscustomobject]@{ Run = $run; Process = $processId; Thread = [int]$cells[4]; ElapsedMs = $elapsed; Category = $cells[2]; Name = $cells[3] })
        }
    }
    return $rows.ToArray()
}

function getTraceRuns {
    # 行を実行（プロセス）ごとにまとめる。Markers は 名前 → 最初の経過ミリ秒（同じ名前が何度か出ても、最初の 1 回）。
    # 定義（Defs）に名前が無い行は、そのまま名前で入れる
    param (
        [object[]]$Rows
    )

    $runs = New-Object System.Collections.Generic.List[object]
    foreach ($group in @($Rows | Group-Object Run | Sort-Object Name)) {
        $markers = @{}
        foreach ($row in @($group.Group | Sort-Object ElapsedMs)) {
            if (-not $markers.ContainsKey($row.Name)) {
                $markers[$row.Name] = $row.ElapsedMs
            }
        }
        $runs.Add(@{ Markers = $markers; Failed = $false })
    }
    return $runs.ToArray()
}

function newTraceDefsFromRuns {
    # 定義が無いとき、記録に出てきた名前を、経過の中央値の順に並べて定義にする（前の区切りが区間の始まり）
    param (
        [object[]]$Runs
    )

    $values = @{}
    foreach ($run in $Runs) {
        if ($run.Failed) { continue }
        foreach ($key in $run.Markers.Keys) {
            if (-not $values.ContainsKey($key)) { $values[$key] = New-Object System.Collections.Generic.List[double] }
            $values[$key].Add([double]$run.Markers[$key])
        }
    }
    $ordered = @($values.Keys | Sort-Object { (getStats $values[$_].ToArray()).Median })
    $defs = New-Object System.Collections.Generic.List[object]
    $previous = $null
    foreach ($key in $ordered) {
        $defs.Add(@{ Key = $key; Label = $key; Previous = $previous; Kind = "mark" })
        $previous = $key
    }
    return $defs.ToArray()
}

function formatTraceMs {
    param (
        $Value
    )

    if ($null -eq $Value) { return "–" }
    return ([Math]::Round([double]$Value)).ToString("0", [System.Globalization.CultureInfo]::InvariantCulture)
}

function getTraceTable {
    # 実行ごとの区切り（Runs）から、区間ごとの時間の表（Markdown）を返す。
    #   Runs: @{ Markers = @{ <Key> = 経過ミリ秒 }; Failed = <失敗した回か> } の並び。失敗の回は数えず、件数だけ添える
    #   Defs: @{ Key; Label; Previous = <区間の始まりの Key。無ければ始まりは 0>; Kind = "mark"（経過の時刻）または "duration"（Markers の値がそのまま長さ） } の並び。省くと記録の名前から作る
    # 区間の長さは、終点の経過 - 始まりの経過（同じ回の中で引く）。欠けた区切りは「–」。長さが負なら、始まりより先に起きた（重なっている）
    param (
        [object[]]$Runs,
        [object[]]$Defs = $null
    )

    if ($null -eq $Defs -or $Defs.Count -eq 0) {
        $Defs = @(newTraceDefsFromRuns $Runs)
    }
    $okRuns = @($Runs | Where-Object { -not $_.Failed })
    $failed = @($Runs | Where-Object { $_.Failed }).Count
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("| 区切り | 区間の終点 | 長さ 中央値 (ms) | 長さ 最大 (ms) | 経過 中央値 (ms) | 経過 最大 (ms) | 回数 |")
    $lines.Add("|---:|---|---:|---:|---:|---:|---:|")
    $number = 0
    foreach ($def in $Defs) {
        $number++
        $lengths = New-Object System.Collections.Generic.List[double]
        $elapsedList = New-Object System.Collections.Generic.List[double]
        foreach ($run in $okRuns) {
            if (-not $run.Markers.ContainsKey($def.Key)) { continue }
            $value = [double]$run.Markers[$def.Key]
            if ($def.Kind -eq "duration") {
                $lengths.Add($value)
                continue
            }
            $elapsedList.Add($value)
            if ($def.Previous) {
                if ($run.Markers.ContainsKey($def.Previous)) { $lengths.Add($value - [double]$run.Markers[$def.Previous]) }
            } else {
                $lengths.Add($value)
            }
        }
        $lengthStats = getStats $lengths.ToArray()
        $elapsedStats = getStats $elapsedList.ToArray()
        $count = [Math]::Max($lengths.Count, $elapsedList.Count)
        $cells = @(
            $number, $def.Label,
            $(if ($lengthStats) { formatTraceMs $lengthStats.Median } else { "–" }), $(if ($lengthStats) { formatTraceMs $lengthStats.Max } else { "–" }),
            $(if ($elapsedStats) { formatTraceMs $elapsedStats.Median } else { "–" }), $(if ($elapsedStats) { formatTraceMs $elapsedStats.Max } else { "–" }),
            $count)
        $lines.Add("| " + ($cells -join " | ") + " |")
    }
    $lines.Add("")
    $lines.Add("回数 $($okRuns.Count) 回（失敗 $failed 回は除く）")
    return ($lines -join "`r`n")
}

if ($Dir) {
    $rows = @(readTraceRows $Dir $Category $ProcessId)
    getTraceTable @(getTraceRuns $rows)
}
