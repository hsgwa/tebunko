# tebunko を起動してから、最初の検索の結果が返るまでの時間を、区間ごとに測る。
#
# ツール（scripts\）を一時フォルダに写し、-Workspace（インデックス入りのワークスペース）を指す設定ファイルを置いて、
# 環境変数 TEBUNKO_TRACE=startup,search で画面を -Count 回起動する（1 回目は温めとして別に出す）。
# 窓が出たら UI オートメーションで検索の欄に語を入れて検索を押し、記録（ワークスペースの trace\）に startup の行が書かれたら閉じる。
# 区間は shared\core\trace.ps1 の記録の区切り（処理の記録。docs\design\testing\perf.md）。区切り 12 は外から測る（窓が入力を受け付けなかった最長の時間）。
# 起動は tebunko.bat と同じ呼び出しの形（powershell.exe -Command "& gui.ps1"）で、起こしたプロセスの PID だけを閉じる・止める（名前で探して止めない）。
# 集計は read_trace.ps1 の関数。結果の表（Markdown）を出力し、-Out があればファイルにも書く。
#   .\tools\perf\measure_startup.ps1 -Workspace <ワークスペース> -Words "まれな語","ない語" -Count 5

param (
    # 測る tebunko のフォルダ（scripts\ を持つ）。省略すると、このリポジトリ
    [string]$Tool,
    # インデックスのあるワークスペース
    [string]$Workspace,
    # 検索する語（語ごとに -Count 回ずつ測る）
    [string[]]$Words = @("まれ", "存在しない語zzqx"),
    [int]$Count = 5,
    # 1 回の起動で、記録が書かれるまで待つ秒数
    [int]$TimeoutSeconds = 120,
    [string]$Out,
    # 記録を入れずに（TEBUNKO_TRACE なし）、外から「起動 → 検索の結果の件数が出る」だけを測る。記録を入れたことで遅くなっていないかを比べるため
    [switch]$NoTrace
)

. ([WildcardPattern]::Escape("$PSScriptRoot\read_trace.ps1".Replace('`', '``')))

function getStartupTraceDefs {
    # startup の区切り 1〜12（shared\core\trace.ps1 の startup。名前は記録の名前と同じ）
    return @(
        @{ Key = "01 PowerShell の起動"; Label = "PowerShell の起動（プロセスの開始 → gui.ps1）"; Previous = $null; Kind = "mark" }
        @{ Key = "stepSplash 5"; Label = "WPF の読み込みと起動中の表示"; Previous = "01 PowerShell の起動"; Kind = "mark" }
        @{ Key = "stepSplash 50"; Label = "shared・settings・lib の読み込み"; Previous = "stepSplash 5"; Kind = "mark" }
        @{ Key = "04 画面の組み立て（Loaded）"; Label = "多重起動の判定・ワークスペース・XAML と部品の読み込み（→ Loaded）"; Previous = "stepSplash 50"; Kind = "mark" }
        @{ Key = "05 窓が出る（ContentRendered）"; Label = "窓が出る（Loaded → ContentRendered）"; Previous = "04 画面の組み立て（Loaded）"; Kind = "mark" }
        @{ Key = "06 起動時の読み込み（画面のスレッドの分）"; Label = "起動時の読み込み（画面のスレッドの分）"; Previous = "05 窓が出る（ContentRendered）"; Kind = "mark" }
        @{ Key = "07 インデックスのツリーが画面に入る"; Label = "ツリーが画面に入る（裏の列）"; Previous = "06 起動時の読み込み（画面のスレッドの分）"; Kind = "mark" }
        @{ Key = "08 検索の司令の用意が済む"; Label = "検索の司令のスレッドの用意が済む（共通部品の読み込み後から。画面と並んで進む）"; Previous = "stepSplash 50"; Kind = "mark" }
        @{ Key = "09 検索の送り"; Label = "検索の送り（ツリーが入ってから。道具が語を入れて押す）"; Previous = "07 インデックスのツリーが画面に入る"; Kind = "mark" }
        @{ Key = "10 表に最初の行が出る"; Label = "表に最初の行が出る"; Previous = "09 検索の送り"; Kind = "mark" }
        @{ Key = "11 総数が決まる"; Label = "総数が決まる"; Previous = "10 表に最初の行が出る"; Kind = "mark" }
        @{ Key = "窓の最長の止まり"; Label = "送りまでに、窓が入力を受け付けなかった最長の時間（外から測る）"; Previous = $null; Kind = "duration" }
    )
}

function getSearchTraceDefs {
    # search の区切り（最初の検索の内訳。名前は記録の名前と同じ。画面のスレッドと司令のスレッドの記録が、同じ起点で並ぶ）
    return @(
        @{ Key = "画面: 送り"; Label = "画面: 送り（起点から）"; Previous = $null; Kind = "mark" }
        @{ Key = "司令: 受け取り"; Label = "司令: 受け取り"; Previous = "画面: 送り"; Kind = "mark" }
        @{ Key = "司令: gram 索引で絞り終え"; Label = "司令: gram 索引で絞り終え"; Previous = "司令: 受け取り"; Kind = "mark" }
        @{ Key = "司令: ファイルの列挙の終わり"; Label = "司令: ファイルの列挙の終わり"; Previous = "司令: gram 索引で絞り終え"; Kind = "mark" }
        @{ Key = "司令: searchContentIndex の終わり"; Label = "司令: 読み込みと照合（searchContentIndex）の終わり"; Previous = "司令: ファイルの列挙の終わり"; Kind = "mark" }
        @{ Key = "画面: 最初の行"; Label = "画面: 表に最初の行（送りから）"; Previous = "画面: 送り"; Kind = "mark" }
        @{ Key = "画面: 総数"; Label = "画面: 総数が決まる（司令の終わりから）"; Previous = "司令: searchContentIndex の終わり"; Kind = "mark" }
    )
}

function getStartupRunFromRows {
    # 1 回の起動の記録（行）から、実行 1 回分（Markers）を作る。外から測った値（Outside・BlockMs）を足す。
    # SearchRows（search の行）があれば、最初の検索の内訳（SearchMarkers。名前に "画面: 送り" などが入る）も持つ
    param (
        [object[]]$Rows,
        [double]$OutsideMs,
        [double]$BlockMs,
        [object[]]$SearchRows = @()
    )

    $runs = @(getTraceRuns $Rows)
    $run = if ($runs.Count -gt 0 -and $runs[0] -is [hashtable]) { $runs[0] } else { @{ Markers = @{}; Failed = $true } }
    if ($run.Markers.ContainsKey("11 総数が決まる") -eq $false) {
        $run.Failed = $true
    }
    $run.Markers["窓の最長の止まり"] = $BlockMs
    $searchRuns = @(getTraceRuns $SearchRows)
    $run.SearchMarkers = $(if ($searchRuns.Count -gt 0 -and $searchRuns[0] -is [hashtable]) { $searchRuns[0].Markers } else { @{} })
    $run.OutsideMs = $OutsideMs
    return $run
}

function getStartupSummaryTable {
    # 区間の表に、外から測った「起動 → 記録が書かれる」と、記録の区切り 11 の経過を並べる（記録の口のずれの確かめ）
    param (
        [object[]]$Runs
    )

    $table = getTraceTable $Runs (getStartupTraceDefs)
    $ok = @($Runs | Where-Object { -not $_.Failed -and $null -ne $_.OutsideMs })
    if ($ok.Count -eq 0) {
        return $table
    }
    $outside = getStats ([double[]]@($ok | ForEach-Object { $_.OutsideMs }))
    $inside = getStats ([double[]]@($ok | ForEach-Object { $_.Markers["11 総数が決まる"] }))
    $diff = if ($outside.Median -gt 0) { [Math]::Round(100 * ($outside.Median - $inside.Median) / $outside.Median, 1) } else { 0 }
    return $table + "`r`n`r`n外から測った 起動 → 記録が書かれる: 中央値 $(formatTraceMs $outside.Median) ms・最大 $(formatTraceMs $outside.Max) ms／記録の区切り 11 の経過: 中央値 $(formatTraceMs $inside.Median) ms（差 $diff %）"
}

function getStartupSearchTable {
    # 最初の検索の内訳（search の区切り）の表。司令の分は、司令のスレッドが同じ起点から測った経過
    param (
        [object[]]$Runs
    )

    $searchRuns = @($Runs | ForEach-Object { @{ Markers = $_.SearchMarkers; Failed = ($_.Failed -or $_.SearchMarkers.Count -eq 0) } })
    return getTraceTable $searchRuns (getSearchTraceDefs)
}

function invokeStartupMeasure {
    # 画面を 1 回起動して、最初の検索までを測る。戻り値は実行 1 回分（getStartupRunFromRows）
    param (
        [string]$ToolDir,
        [string]$Word,
        [string]$TraceDir,
        [string]$DefaultWorkspace,
        [int]$TimeoutSeconds,
        [switch]$NoTrace
    )

    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $ae = [Windows.Automation.AutomationElement]
    $gui = Join-Path $ToolDir "scripts\tebunko\gui.ps1"
    $command = "& '$($gui.Replace("'", "''"))'"
    $savedTrace = $env:TEBUNKO_TRACE
    $savedDefault = $env:TEBUNKO_DEFAULT_WORKSPACE
    $env:TEBUNKO_TRACE = $(if ($NoTrace) { $null } else { "startup,search" })
    $env:TEBUNKO_DEFAULT_WORKSPACE = $DefaultWorkspace
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $process = Start-Process powershell.exe -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "RemoteSigned", "-Command", $command) -PassThru -WindowStyle Hidden
    } finally {
        $env:TEBUNKO_TRACE = $savedTrace
        $env:TEBUNKO_DEFAULT_WORKSPACE = $savedDefault
    }
    $null = $process.Handle
    $blockMs = 0.0
    $outsideMs = 0.0
    $rows = @()
    $searchRows = @()
    try {
        $deadline = $watch.Elapsed.TotalSeconds + $TimeoutSeconds
        $window = $null
        $wordBox = $null
        $button = $null
        $sent = $false
        while ($watch.Elapsed.TotalSeconds -lt $deadline -and -not $process.HasExited) {
            $probe = [System.Diagnostics.Stopwatch]::StartNew()
            if ($null -eq $window) {
                $condition = New-Object Windows.Automation.PropertyCondition($ae::ProcessIdProperty, $process.Id)
                foreach ($top in @($ae::RootElement.FindAll("Children", $condition) | ForEach-Object { $_ })) {
                    $found = $top.FindFirst("Descendants", (New-Object Windows.Automation.PropertyCondition($ae::AutomationIdProperty, "NavList")))
                    if ($found) { $window = $top }
                }
            } elseif (-not $sent) {
                if ($null -eq $wordBox) {
                    $wordBox = $window.FindFirst("Descendants", (New-Object Windows.Automation.PropertyCondition($ae::AutomationIdProperty, "WordBox")))
                    $button = $window.FindFirst("Descendants", (New-Object Windows.Automation.PropertyCondition($ae::AutomationIdProperty, "SearchButton")))
                }
                if ($wordBox -and $button) {
                    if ($wordBox.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).Current.Value -ne $Word) {
                        $wordBox.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Word)
                    }
                    # ［検索］が押せるのは、ツリーが入ってから（読み込み中は押せない）
                    if ($button.Current.IsEnabled) {
                        $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
                        $sent = $true
                    }
                }
            } elseif ($NoTrace) {
                # 記録が無いので、検索の結果の要約（SummaryText）が「検索中…」でなくなった時を終わりとする
                $summary = $window.FindFirst("Descendants", (New-Object Windows.Automation.PropertyCondition($ae::AutomationIdProperty, "SummaryText")))
                if ($summary -and $summary.Current.Name -like "一致*") {
                    $outsideMs = $watch.Elapsed.TotalMilliseconds
                    break
                }
            } else {
                $rows = @(readTraceRows $TraceDir "startup" $process.Id)
                if (@($rows | Where-Object { $_.Name -eq "11 総数が決まる" }).Count -gt 0) {
                    $outsideMs = $watch.Elapsed.TotalMilliseconds
                    # 司令のスレッドの記録は、画面が終わりに気づくのと前後して書かれる。少し待って読む
                    Start-Sleep -Milliseconds 400
                    $searchRows = @(readTraceRows $TraceDir "search" $process.Id)
                    break
                }
            }
            # UI オートメーションの呼び出しは、窓が応答しない間は返らない。1 回の呼び出しの長さの最大を、止まりの長さとする（窓が出てから送りまで）
            if ($window -and -not $sent) {
                $blockMs = [Math]::Max($blockMs, $probe.Elapsed.TotalMilliseconds)
            }
            Start-Sleep -Milliseconds 20
        }
    } finally {
        if (-not $process.HasExited) {
            [void]$process.CloseMainWindow()
            if (-not $process.WaitForExit(20000)) {
                $process.Kill()
                [void]$process.WaitForExit(10000)
            }
        }
        $pidText = "_$($process.Id)_"
        if ([System.IO.Directory]::Exists($TraceDir)) {
            foreach ($path in @([System.IO.Directory]::GetFiles($TraceDir, "*.tsv") | Where-Object { [System.IO.Path]::GetFileName($_).Contains($pidText) })) {
                try { [System.IO.File]::Delete($path) } catch { $null = $_ }
            }
        }
    }
    $run = getStartupRunFromRows $rows $outsideMs $blockMs $searchRows
    if ($NoTrace) {
        $run.Failed = ($outsideMs -le 0)
    }
    return $run
}

function invokeStartupMeasureAll {
    param (
        [string]$Tool,
        [string]$Workspace,
        [string[]]$Words,
        [int]$Count,
        [int]$TimeoutSeconds,
        [string]$Out,
        [switch]$NoTrace
    )

    if (-not $Tool) { $Tool = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath }
    # powershell.exe -File で呼ぶと、-Words "a","b" は "a,b" という 1 つの文字列で渡る。カンマで分ける
    $Words = @($Words | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $Workspace -or -not (Test-Path -LiteralPath $Workspace)) {
        throw "-Workspace に、インデックスのあるワークスペースのフォルダを渡してください。"
    }
    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ("tebunko-startup-" + [guid]::NewGuid().ToString("N"))
    $toolCopy = Join-Path $temp "tool"
    [void][System.IO.Directory]::CreateDirectory($toolCopy)
    $defaultWorkspace = Join-Path $temp "default_workspace"
    [void][System.IO.Directory]::CreateDirectory($defaultWorkspace)
    Copy-Item -LiteralPath (Join-Path $Tool "scripts") -Destination (Join-Path $toolCopy "scripts") -Recurse
    $config = [ordered]@{ workspaceFolder = (Resolve-Path -LiteralPath $Workspace).ProviderPath }
    [System.IO.File]::WriteAllText((Join-Path $toolCopy "setting.config"), (ConvertTo-Json -InputObject $config), (New-Object System.Text.UTF8Encoding($false)))
    $traceDir = Join-Path $config.workspaceFolder "trace"
    $report = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($word in $Words) {
            $warm = invokeStartupMeasure $toolCopy $word $traceDir $defaultWorkspace $TimeoutSeconds -NoTrace:$NoTrace
            $runs = New-Object System.Collections.Generic.List[object]
            for ($i = 1; $i -le $Count; $i++) {
                $runs.Add((invokeStartupMeasure $toolCopy $word $traceDir $defaultWorkspace $TimeoutSeconds -NoTrace:$NoTrace))
            }
            if ($NoTrace) {
                $outside = getStats ([double[]]@($runs | Where-Object { -not $_.Failed } | ForEach-Object { $_.OutsideMs }))
                $report.Add("### 語「$word」記録なし（温め 1 回を除く $Count 回）")
                $report.Add("")
                $report.Add($(if ($outside) { "外から測った 起動 → 検索の結果の件数が出る: 中央値 $(formatTraceMs $outside.Median) ms・最大 $(formatTraceMs $outside.Max) ms（$($outside.Count) 回）" } else { "測れなかった" }))
                $report.Add("")
                continue
            }
            $report.Add("### 語「$word」（温め 1 回を除く $Count 回）")
            $report.Add("")
            $report.Add((getStartupSummaryTable $runs.ToArray()))
            $report.Add("")
            $report.Add("最初の検索の内訳（search。経過は同じ起点から。長さは表の前の区切りとの差）")
            $report.Add("")
            $report.Add((getStartupSearchTable $runs.ToArray()))
            $report.Add("")
            $warmText = if ($warm.Failed) { "失敗" } else { "総数まで $(formatTraceMs $warm.Markers['11 総数が決まる']) ms" }
            $report.Add("温め 1 回目: $warmText")
            $report.Add("")
        }
    } finally {
        try { [System.IO.Directory]::Delete($temp, $true) } catch { $null = $_ }
    }
    $text = $report -join "`r`n"
    if ($Out) {
        [System.IO.File]::WriteAllText($Out, $text, (New-Object System.Text.UTF8Encoding($true)))
    }
    return $text
}

if ($MyInvocation.InvocationName -ne ".") {
    invokeStartupMeasureAll $Tool $Workspace $Words $Count $TimeoutSeconds $Out -NoTrace:$NoTrace
}
