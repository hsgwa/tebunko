# Office からの取り込みの計測（measure_ingest.ps1・measure_perf.ps1）で使う部品。
# 取り込みの計測が tebunko の受け渡しの口・取り込み一覧から読んだ値を、確かめて数える処理をまとめる（単体でテストする）。

# 受け渡しの口（Progress.Phase）の値（計測の口。tebunko の core\paths.ps1 の ${indexingPhase*} と同じ）
$ingestPhaseNames = @("クロール", "確認", "取り込み", "仕上げ")
$ingestPhaseIngest = "取り込み"
# 取り込み一覧の状態の値とファイル名（計測の口）
$ingestStateDone = "済"
$ingestStateFailed = "失敗"
$ingestStatusFileName = "取り込み一覧.tsv"

function getIngestKindCounts {
    # データのフォルダにあるファイルを、種類（拡張子）ごとに数える。@{ xlsx; docx; pptx; doc; ppt; Total }
    param (
        [string]$dataDir
    )

    $counts = [ordered]@{ xlsx = 0; docx = 0; pptx = 0; doc = 0; ppt = 0 }
    $total = 0
    foreach ($f in [System.IO.Directory]::GetFiles($dataDir, "*", "AllDirectories")) {
        $ext = [System.IO.Path]::GetExtension($f).TrimStart(".").ToLowerInvariant()
        if ($counts.Contains($ext)) { $counts[$ext]++; $total++ }
    }
    $counts.Total = $total
    return $counts
}

function getIngestLanes {
    # データにある種類から、使う Office のレーンを決める（Excel・Word・PowerPoint は種類ごとに 1 つ。.docx・.pptx は Office を使わない）
    param (
        $counts
    )

    $lanes = New-Object System.Collections.Generic.List[string]
    if ($counts.xlsx -gt 0) { $lanes.Add("Excel") }
    if ($counts.doc -gt 0) { $lanes.Add("Word") }
    if ($counts.ppt -gt 0) { $lanes.Add("PowerPoint") }
    if ($lanes.Count -eq 0) { return "なし" }
    return ($lanes -join "・")
}

function readIngestStatusCounts {
    # 取り込み一覧（TSV）から、成功（済）と失敗の数を数える。@{ Done; Failed; Other }。
    # 見出しの行（クロール対象フォルダの行の後にある。「相対パス」で始まる最初の行）から「相対パス」「状態」の列を名前で探し、列の数が見出しと同じ行だけを見る
    # （クロール対象フォルダの行は列の数が違うので外れる）。同じ相対パスが何行あっても、最後の行の状態を使う
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) { throw "取り込み一覧がありません（取り込みが動いていません）。" }
    $lines = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) -split "\r?\n"
    $headerIndex = [Array]::FindIndex([string[]]$lines, [Predicate[string]]{ param($l) $l.StartsWith("相対パス`t") })
    if ($headerIndex -lt 0) { throw "取り込み一覧に見出しの行がありません。" }
    $header = $lines[$headerIndex].Split("`t")
    $pathColumn = [Array]::IndexOf($header, "相対パス")
    $stateColumn = [Array]::IndexOf($header, "状態")
    if ($pathColumn -lt 0 -or $stateColumn -lt 0) { throw "取り込み一覧の見出しに「相対パス」「状態」の列がありません。" }
    $states = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($line in @($lines | Select-Object -Skip ($headerIndex + 1))) {
        $f = $line.Split("`t")
        if ($f.Count -ne $header.Count -or $f[$pathColumn] -eq "") { continue }
        $states[$f[$pathColumn]] = $f[$stateColumn]
    }
    $done = 0; $failed = 0; $other = 0
    foreach ($s in $states.Values) {
        if ($s -eq $ingestStateDone) { $done++ } elseif ($s -eq $ingestStateFailed) { $failed++ } else { $other++ }
    }
    return @{ Done = $done; Failed = $failed; Other = $other }
}

function assertIngestCounts {
    # 成功 + 失敗 = データのファイル数でなければ失敗にする（取り込まれなかったファイルを黙って見逃さない）
    param (
        $status,
        [int]$total
    )

    if (($status.Done + $status.Failed) -ne $total) {
        throw ("取り込み一覧の成功 {0} + 失敗 {1} が、データのファイル数 {2} と合いません（未取り込み・その他 {3}）。" -f $status.Done, $status.Failed, $total, $status.Other)
    }
}

function getIngestPhaseSeconds {
    # 記録した段階（@{ Name; StartMs; EndMs } の配列）から、段階ごとの秒を返す（同じ段階が何度かあれば足す）。
    # 取り込みの段階が一度も読めなかったとき、知らない段階の名前があったときは失敗にする
    # （1 ファイルあたりの ms は取り込みの段階の秒から出すので、黙って 0 秒や空の行を出さない）
    param (
        [object[]]$phases
    )

    $seconds = [ordered]@{}
    foreach ($p in @($phases)) {
        if ($ingestPhaseNames -notcontains $p.Name) { throw "知らない段階の名前が受け渡しの口に来ました: $($p.Name)" }
        if ($null -eq $p.EndMs) { throw "段階「$($p.Name)」が閉じていません。" }
        if (!$seconds.Contains($p.Name)) { $seconds[$p.Name] = 0.0 }
        $seconds[$p.Name] += ($p.EndMs - $p.StartMs) / 1000
    }
    if (!$seconds.Contains($ingestPhaseIngest)) { throw "取り込みの段階が読めませんでした（受け渡しの口の Progress.Phase が変わった疑いがあります）。" }
    return $seconds
}

function getOfficeVersions {
    # 入っている Office（Excel・Word・PowerPoint）のビルド番号を返す。取れないときは「不明」
    $versions = [ordered]@{}
    foreach ($app in @(@("Excel", "EXCEL.EXE"), @("Word", "WINWORD.EXE"), @("PowerPoint", "POWERPNT.EXE"))) {
        $version = "不明"
        try {
            $exe = ([string](Get-Item -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$($app[1])" -ErrorAction Stop).GetValue("")).Trim('"')
            $found = (Get-Item -LiteralPath $exe -ErrorAction Stop).VersionInfo.FileVersion
            if ($found) { $version = [string]$found }
        } catch { }
        $versions[$app[0]] = $version
    }
    return $versions
}

function newIngestResult {
    # 取り込みを何回か測った結果（measure_ingest.ps1 の JSON を読んだもの）から、result.json の Ingest を作る。
    # 全体・段階ごと・1 ファイルあたりは最小・中央値・平均・最大（getStats）。リソースは、全体の時間が中央値の回のもの
    param (
        [object[]]$runs
    )

    $first = $runs[0]
    $byTime = @($runs | Sort-Object { [double]$_.Seconds })
    $median = $byTime[[int][Math]::Floor(($byTime.Count - 1) / 2)]
    $phases = New-Object System.Collections.Generic.List[object]
    foreach ($name in $ingestPhaseNames) {
        $values = [double[]]@($runs | ForEach-Object { @($_.Phases) | Where-Object { $_.Phase -eq $name } | ForEach-Object { $_.Seconds } })
        if ($values.Count) { $phases.Add([ordered]@{ Phase = $name; Stats = (getStats $values) }) }
    }
    return [ordered]@{
        Files = $first.Files; Total = $first.Total
        Done = ($runs | ForEach-Object { $_.Done } | Measure-Object -Minimum).Minimum
        Failed = ($runs | ForEach-Object { $_.Failed } | Measure-Object -Maximum).Maximum
        Threads = $first.Threads; Lanes = $first.Lanes; Repeat = $runs.Count
        Seconds = (getStats ([double[]]@($runs | ForEach-Object { $_.Seconds })))
        PhaseSeconds = $phases.ToArray()
        PerFileMs = (getStats ([double[]]@($runs | Where-Object { $null -ne $_.PerFileMs } | ForEach-Object { $_.PerFileMs })))
        Resources = $median.Resources
        ResourceRun = ([Array]::IndexOf($runs, $median) + 1)
        PeakWorkingSetMB = ($runs | ForEach-Object { $_.PeakWorkingSetMB } | Measure-Object -Maximum).Maximum
    }
}