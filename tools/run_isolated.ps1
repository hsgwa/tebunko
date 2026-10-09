# 実機の確かめを、利用者の既定のワークスペースに触れない形で行う。
#
#   .\tools\run_isolated.ps1                       使い捨てのフォルダにツールの写しを作り、画面を開く（閉じるまで待つ）
#   .\tools\run_isolated.ps1 -Single               展開せずに動く単一 .ps1 版で開く
#   .\tools\run_isolated.ps1 -WaitSeconds 30     画面を 30 秒だけ開き、過ぎたら起動した PID とその子だけを止める
#   .\tools\run_isolated.ps1 -Settings @{ ... }    写しの setting.config に書く項目（targetFolders など）
#   .\tools\run_isolated.ps1 -Command { param ($ToolDir, $Workspace) ... }   画面を開かず、写しの場所と使い捨てのワークスペースを受け取って任意の確かめを流す
#
# 流れ:
#   1. %TEMP% の下の新しいフォルダに、ツールの写し（scripts\ と tebunko.bat。-Single なら単一 .ps1）と、使い捨てのワークスペースを作る。
#   2. 子プロセスの環境変数 TEBUNKO_DEFAULT_WORKSPACE に、使い捨てのワークスペースを入れる（既定のワークスペースがそこになる）。
#   3. 起動の前に、その環境で設定ファイルの場所と work の場所を求め、どちらも使い捨てのフォルダの中でなければ、起動せずに終了コード 1 で止まる。
#   4. 起動の前後で、既定のワークスペースと、写し元のリポジトリの setting.config・work\ の名前・大きさ・更新時刻を比べる
#      （中身は読まない）。違いがあれば一覧を出して終了コード 1。
#   5. 使い捨てのフォルダは終わったら消す。起動したプロセスは PID を出す（止めるときは PID で止める）。
param (
    [switch]$Single,
    [hashtable]$Settings = @{},
    [scriptblock]$Command,
    [int]$WaitSeconds = 0,     # 画面を待つ秒数（0 は閉じるまで待つ）。過ぎたら、起動した PID とその子だけを止める
    [string]$RealWorkspace     # 既定のワークスペースの代わりに比べる場所（テスト用。既定は利用者の既定のワークスペース）
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path "$PSScriptRoot\..").Path
. "$PSScriptRoot\isolation\isolation_common.ps1"

function getIsolatedProbeResult {
    # 使い捨ての環境で、写しの設定ファイルの場所と work の場所を求める（子プロセス。書き込まない）
    param ([string]$ScriptsDir, [string]$Base, [string]$WorkspaceDir)

    $probe = Join-Path $Base "probe.ps1"
    $text = @'
param([string]$Lib)
. $Lib
"settingsFile=" + ${settingsFile}
"workDir=" + (getWorkDir)
'@
    [System.IO.File]::WriteAllText($probe, $text, (New-Object System.Text.UTF8Encoding($true)))
    $previous = $env:TEBUNKO_DEFAULT_WORKSPACE
    $env:TEBUNKO_DEFAULT_WORKSPACE = $WorkspaceDir
    try {
        $lines = @(& powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $probe -Lib (Join-Path $ScriptsDir "tebunko\lib.ps1"))
    } finally {
        $env:TEBUNKO_DEFAULT_WORKSPACE = $previous
    }
    $result = @{}
    foreach ($line in $lines) {
        if ($line -match '^(settingsFile|workDir)=(.*)$') { $result[$Matches[1]] = $Matches[2] }
    }
    return $result
}

function invokeIsolated {
    param ([bool]$SingleScript, [hashtable]$Config, [scriptblock]$Action, [int]$Wait, [string]$RealDir)

    if (!$RealDir) { $RealDir = getRealDefaultWorkspace }
    $base = newIsolatedFolder
    $tool = Join-Path $base "tool"
    $WorkspaceDir = Join-Path $base "workspace"
    $exitCode = 0
    try {
        $watch = newIsolationWatch ([ordered]@{
            "既定のワークスペース" = $RealDir
            "リポジトリの setting.config" = "$repoRoot\setting.config"
            "リポジトリの work" = "$repoRoot\work"
        })
        Write-Host "使い捨てのフォルダ: $base"
        [void][System.IO.Directory]::CreateDirectory($tool)
        Copy-Item -LiteralPath "$repoRoot\scripts" -Destination "$tool\scripts" -Recurse
        Copy-Item -LiteralPath "$repoRoot\tebunko.bat" -Destination "$tool\tebunko.bat"
        if ($Config.Count -gt 0) {
            [System.IO.File]::WriteAllText("$tool\setting.config", (ConvertTo-Json -InputObject $Config -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
        }

        # 書く前に、書き先が使い捨ての中であることを確かめる
        $probe = getIsolatedProbeResult "$tool\scripts" $base $WorkspaceDir
        Write-Host "設定ファイルの場所: $($probe.settingsFile)"
        Write-Host "work の場所: $($probe.workDir)"
        $outside = @(@($probe.settingsFile, $probe.workDir) | Where-Object { !(testIsolationPathInside $_ $base) })
        if ($probe.Count -ne 2 -or $outside.Count -gt 0) {
            Write-Host "書き込み先が使い捨てのフォルダの外を指しているため、起動しません: $($outside -join ', ')" -ForegroundColor Red
            $exitCode = 1
        } else {
            [void][System.IO.Directory]::CreateDirectory($WorkspaceDir)
            $previous = $env:TEBUNKO_DEFAULT_WORKSPACE
            $env:TEBUNKO_DEFAULT_WORKSPACE = $WorkspaceDir
            try {
                if ($Action) {
                    & $Action $tool $WorkspaceDir | Out-Host
                } else {
                    $gui = "$tool\scripts\tebunko\gui.ps1"
                    if ($SingleScript) {
                        $gui = Join-Path $tool "tebunko-isolated.ps1"
                        & "$repoRoot\tools\new_single_script.ps1" -Version "v0.0.0-isolated" -OutFile $gui | Out-Null
                    }
                    $p = Start-Process powershell.exe -PassThru -WindowStyle Hidden -ArgumentList @(
                        "-NoProfile", "-STA", "-ExecutionPolicy", "RemoteSigned", "-Command", "& '$($gui.Replace("'", "''"))'")
                    $null = $p.Handle   # ExitCode を取るため、起動の直後にハンドルを持つ
                    Write-Host "起動したプロセスの PID: $($p.Id)$(if ($Wait -gt 0) { "（$Wait 秒たったら止めます。それまでに画面を閉じれば、そこで終わります）" } else { "（画面を閉じるまで待ちます）" })"
                    $waited = waitIsolationProcess $p $Wait
                    $verdict = getIsolationLaunchVerdict $waited $Wait $p.Id
                    if ($verdict.Line) { Write-Host $verdict.Line -ForegroundColor $(if ($verdict.ExitCode -ne 0) { "Red" } else { "Gray" }) }
                    if ($verdict.ExitCode -ne 0) { $exitCode = 1 }
                }
            } finally {
                $env:TEBUNKO_DEFAULT_WORKSPACE = $previous
            }
        }
    } finally {
        # 写しの作成・事前の確かめ・確かめの途中の例外のどの道でも、既定のワークスペースを比べてから消す
        try {
            if (!$watch) {
                Write-Host "前後の比べを用意できなかったため、比べていません" -ForegroundColor Red
                $exitCode = 1
            } else {
                $report = getIsolationReport @(getIsolationWatchDiffs $watch) "起動の前後で、既定のワークスペースなどに違いがありました"
                if ($report.ExitCode -ne 0) {
                    $report.Lines | ForEach-Object { Write-Host $_ -ForegroundColor Red }
                    $exitCode = 1
                } else {
                    Write-Host "既定のワークスペースなどに違いなし"
                }
            }
        } finally {
            Remove-Item -LiteralPath $base -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    return $exitCode
}

exit (invokeIsolated $Single.IsPresent $Settings $Command $WaitSeconds $RealWorkspace)
