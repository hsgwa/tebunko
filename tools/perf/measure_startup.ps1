<#
.SYNOPSIS
tebunko.bat を何回か起動して、画面が出るまでの時間を段階ごとに測り、Markdown の表で出す。

.DESCRIPTION
ツール（tebunko.bat と scripts\）を一時フォルダに写し、設定ファイルでワークスペースを指定して、tebunko.bat を起動する。
最初の 1 回は温めとして捨てる。1 回ごとに、起動の前に起動の記録（ワークスペース直下の startup_trace.tsv）を消し、環境変数
TEBUNKO_STARTUP_TRACE=1 を付けて起動する。記録が書かれたら（画面の一覧が出たら）本体の窓を閉じて次へ進む。
道具の側でも「tebunko.bat を起動してから本体の窓（NavList を持つ窓）が出るまで」を外から測る（記録の口が無い版でも測れる値）。
記録の口が無い版（-Before に古い版を渡したとき）は、本体の窓が出てから -Settle 秒待って閉じる。
-Before を渡すと、前と後を交互に（前・後・前・後…）起動して同じ表に並べる（機械の温まり・ほかの負荷の偏りを打ち消すため）。
-Timeout 秒までに本体の窓か記録が出ない回（知らせのダイアログ・多重起動・起動の失敗）は失敗として数え、
道具が起動したプロセス（tebunko.bat の子の conhost と PowerShell）だけを PID で止めて次へ進む。名前でプロセスを探して止めない。

利用者の既定のワークスペースには触れない（TEBUNKO_DEFAULT_WORKSPACE を一時フォルダに差し替えて起動する）。
ワークスペース（-Workspace）は、そのまま使う（起動の記録を書く。インデックスは読むだけ）。測るときはローカルのフォルダを渡す
（ネットワークのパスには記録を書かない）。

.PARAMETER Workspace
ワークスペースのフォルダ。省略すると、空の一時フォルダ（インデックス無し）。

.PARAMETER Before
前の版のツールのフォルダ（tebunko.bat と scripts\ があるフォルダ）。渡すと前と後を交互に測って並べる。

.PARAMETER Runs
測る回数（温めの 1 回は別）。

.PARAMETER Timeout
1 回の起動を待つ上限（秒）。

.PARAMETER Settle
記録の口が無い版で、本体の窓が出てから閉じるまで待つ秒数。

.PARAMETER Out
表を書くファイル（省略すると画面に出すだけ）。
#>
param (
    [string]$Workspace = "",
    [string]$Before = "",
    [int]$Runs = 10,
    [int]$Timeout = 120,
    [int]$Settle = 5,
    [string]$Out = ""
)

$ErrorActionPreference = "Stop"
. ([WildcardPattern]::Escape("$PSScriptRoot\startup_common.ps1".Replace('`', '``')))
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$traceFileName = "startup_trace.tsv"

function copyStartupTool {
    # tebunko.bat と scripts\ を dir の下に写し、ワークスペースを指す設定ファイルを置く。ツールのフォルダを返す
    param (
        [string]$source,
        [string]$dir,
        [string]$workspace
    )

    [void][System.IO.Directory]::CreateDirectory($dir)
    Copy-Item -LiteralPath (Join-Path $source "tebunko.bat") -Destination $dir
    Copy-Item -LiteralPath (Join-Path $source "scripts") -Destination (Join-Path $dir "scripts") -Recurse
    $config = [ordered]@{ workspaceFolder = $workspace }
    [System.IO.File]::WriteAllText((Join-Path $dir "setting.config"), (ConvertTo-Json -InputObject $config), (New-Object System.Text.UTF8Encoding($false)))
    return $dir
}

function findStartedProcesses {
    # tebunko.bat（cmd.exe の PID）が起動した conhost.exe と、その子の powershell.exe を探す: @{ Conhost; PowerShell }（まだ無ければ $null）。
    # 親の PID をたどるので、同じ機械でほかの tebunko が動いていても、自分が起動したものだけが見つかる
    param (
        [int]$cmdId
    )

    $conhost = Get-CimInstance Win32_Process -Filter "Name = 'conhost.exe' AND ParentProcessId = $cmdId" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $conhost) {
        return $null
    }
    $shell = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' AND ParentProcessId = $($conhost.ProcessId)" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $shell) {
        return $null
    }
    return @{ Conhost = [int]$conhost.ProcessId; PowerShell = [int]$shell.ProcessId }
}

function testMainWindow {
    # そのプロセスに、本体の窓（NavList を持つ窓）があるか
    param (
        [int]$processId
    )

    try {
        $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $processId)
        $idCondition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, "NavList")
        foreach ($window in [System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $condition)) {
            if ($null -ne $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $idCondition)) {
                return $true
            }
        }
    } catch {
        # 窓を調べられなかった回は、まだ出ていないものとして次の周で調べ直す
        $null = $_.Exception
    }
    return $false
}

function closeStartedGui {
    # 本体を閉じる（CloseMainWindow。効かなければ UI オートメーションの Close）。終わらなければ、控えた PID だけを止める
    param (
        $started
    )

    $process = Get-Process -Id $started.PowerShell -ErrorAction SilentlyContinue
    if ($null -ne $process) {
        [void]$process.CloseMainWindow()
        if (!$process.WaitForExit(10000)) {
            try {
                $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $started.PowerShell)
                foreach ($window in [System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $condition)) {
                    $pattern = $null
                    if ($window.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern, [ref]$pattern)) { $pattern.Close() }
                }
            } catch {
                # 閉じられなければ、待ったあとに PID で止める
                $null = $_.Exception
            }
            [void]$process.WaitForExit(20000)
        }
        $process.Dispose()
    }
    stopStartedPids $started
}

function stopStartedPids {
    # 控えた PID（conhost と PowerShell）が残っていれば止める
    param (
        $started
    )

    foreach ($id in @($started.PowerShell, $started.Conhost)) {
        if ($null -ne $id) {
            Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
        }
    }
}

function invokeStartupRun {
    # 1 回起動して、記録と外から測った時間を返す: @{ Ok; Reason; Run（getStartupRunTimes の結果 + ExternalMs）; Problems }
    param (
        [string]$tool,
        [string]$workspace,
        [string]$defaultWorkspace,
        [int]$timeout,
        [int]$settle
    )

    $traceFile = Join-Path $workspace $traceFileName
    if (Test-Path -LiteralPath $traceFile) {
        Remove-Item -LiteralPath $traceFile -Force
    }
    $supportsTrace = (Get-Content -LiteralPath (Join-Path $tool "scripts\tebunko\gui.ps1") -Raw -Encoding UTF8) -match "TEBUNKO_STARTUP_TRACE"

    $previousTrace = $env:TEBUNKO_STARTUP_TRACE
    $previousDefault = $env:TEBUNKO_DEFAULT_WORKSPACE
    $env:TEBUNKO_STARTUP_TRACE = "1"
    $env:TEBUNKO_DEFAULT_WORKSPACE = $defaultWorkspace
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $cmd = Start-Process -FilePath "cmd.exe" -ArgumentList @("/c", ('"{0}"' -f (Join-Path $tool "tebunko.bat"))) -PassThru -WindowStyle Hidden
    } finally {
        $env:TEBUNKO_STARTUP_TRACE = $previousTrace
        $env:TEBUNKO_DEFAULT_WORKSPACE = $previousDefault
    }

    $started = $null
    $externalMs = $null
    $windowSeen = $null
    $result = @{ Ok = $false; Reason = ""; Run = $null; Problems = @() }
    try {
        while ($true) {
            if ($null -eq $started) {
                $started = findStartedProcesses $cmd.Id
            }
            if ($null -ne $started -and $null -eq $externalMs -and (testMainWindow $started.PowerShell)) {
                $externalMs = $watch.Elapsed.TotalMilliseconds
                $windowSeen = [System.Diagnostics.Stopwatch]::StartNew()
            }
            $traceDone = $false
            if ($supportsTrace -and (Test-Path -LiteralPath $traceFile)) {
                try {
                    $read = readStartupTrace ([System.IO.File]::ReadAllLines($traceFile, [System.Text.Encoding]::UTF8))
                    $traceDone = @($read.Rows | Where-Object { $_.Name -eq ${startupMarkList} }).Count -gt 0
                } catch {
                    $traceDone = $false
                }
            }
            if ($traceDone -and $null -ne $externalMs) {
                $run = getStartupRunTimes $read.Rows
                $run.ExternalMs = $externalMs
                $result.Run = $run
                $result.Problems = $read.Problems
                $result.Ok = $true
                break
            }
            if (!$supportsTrace -and $null -ne $windowSeen -and $windowSeen.Elapsed.TotalSeconds -ge $settle) {
                $result.Run = @{ Stages = @(); WindowMs = $null; ListMs = $null; ExternalMs = $externalMs }
                $result.Ok = $true
                break
            }
            if ($watch.Elapsed.TotalSeconds -gt $timeout) {
                $result.Reason = "$timeout 秒以内に、本体の窓か記録が出なかった（知らせのダイアログ・多重起動・起動の失敗）"
                break
            }
            Start-Sleep -Milliseconds 50
        }
    } finally {
        if ($null -ne $started) {
            closeStartedGui $started
        } elseif ($null -ne $cmd -and !$cmd.HasExited) {
            Stop-Process -Id $cmd.Id -Force -ErrorAction SilentlyContinue
        }
        $cmd.Dispose()
    }
    return $result
}

# ---- 準備 ----
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("tebunko_startup_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
[void][System.IO.Directory]::CreateDirectory($work)
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
if ($Workspace -eq "") {
    $Workspace = Join-Path $work "workspace"
}
[void][System.IO.Directory]::CreateDirectory($Workspace)
$defaultWorkspace = Join-Path $work "default_workspace"
[void][System.IO.Directory]::CreateDirectory($defaultWorkspace)

$tools = [ordered]@{}
$tools["after"] = copyStartupTool $repoRoot (Join-Path $work "after") $Workspace
if ($Before -ne "") {
    $tools["before"] = copyStartupTool (Resolve-Path $Before).Path (Join-Path $work "before") $Workspace
}

$runsByKind = @{ after = New-Object System.Collections.Generic.List[object]; before = New-Object System.Collections.Generic.List[object] }
$failures = New-Object System.Collections.Generic.List[string]
$problems = New-Object System.Collections.Generic.List[string]
try {
    # 温め（捨てる）。前の版も 1 回ずつ起動して、ファイルをキャッシュに載せる
    foreach ($kind in $tools.Keys) {
        Write-Output "温め（$kind）…"
        [void](invokeStartupRun $tools[$kind] $Workspace $defaultWorkspace $Timeout $Settle)
    }
    for ($i = 1; $i -le $Runs; $i++) {
        foreach ($kind in @($tools.Keys)) {
            Write-Output "$i / $Runs 回目（$kind）…"
            $r = invokeStartupRun $tools[$kind] $Workspace $defaultWorkspace $Timeout $Settle
            if ($r.Ok) {
                $runsByKind[$kind].Add($r.Run)
                foreach ($p in $r.Problems) { $problems.Add("${kind} ${i} 回目: $p") }
            } else {
                $failures.Add("${kind} ${i} 回目: $($r.Reason)")
            }
        }
    }
} finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

$after = getStartupSummary $runsByKind["after"].ToArray()
$beforeSummary = $null
if ($tools.Contains("before")) {
    $beforeSummary = getStartupSummary $runsByKind["before"].ToArray()
}
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("起動の時間（ms）。成功した回数: 後 $($after.Runs) / $Runs" + $(if ($null -ne $beforeSummary) { "・前 $($beforeSummary.Runs) / $Runs" } else { "" }) + "。")
$lines.Add("")
foreach ($line in @(formatStartupTable $after $beforeSummary)) { $lines.Add($line) }
if ($failures.Count -gt 0) {
    $lines.Add("")
    $lines.Add("失敗した回:")
    foreach ($f in $failures) { $lines.Add("- $f") }
}
if ($problems.Count -gt 0) {
    $lines.Add("")
    $lines.Add("記録の行を飛ばした回:")
    foreach ($p in $problems) { $lines.Add("- $p") }
}
$text = $lines -join "`r`n"
Write-Output $text
if ($Out -ne "") {
    [System.IO.File]::WriteAllText($Out, $text + "`r`n", (New-Object System.Text.UTF8Encoding($true)))
}
