# 実機の確かめを、利用者の本物のワークスペースに触れない形で行う。
#
#   .\tools\run_isolated.ps1                       使い捨てのフォルダにツールの写しを作り、画面を開く（閉じるまで待つ）
#   .\tools\run_isolated.ps1 -Single               展開せずに動く単一 .ps1 版で開く
#   .\tools\run_isolated.ps1 -Settings @{ ... }    写しの setting.config に書く項目（targetFolders など）
#   .\tools\run_isolated.ps1 -Command { param ($ToolDir, $Workspace) ... }   画面を開かず、写しの場所と使い捨てのワークスペースを受け取って任意の確かめを流す
#
# 流れ:
#   1. %TEMP% の下の新しいフォルダに、ツールの写し（scripts\ と tebunko.bat。-Single なら単一 .ps1）と、使い捨てのワークスペースを作る。
#   2. 子プロセスの環境変数 TEBUNKO_DEFAULT_WORKSPACE に、使い捨てのワークスペースを入れる（既定のワークスペースがそこになる）。
#   3. 起動の前に、その環境で設定ファイルの場所と work の場所を求め、どちらも使い捨てのフォルダの中でなければ、起動せずに終了コード 1 で止まる。
#   4. 起動の前後で、本物の既定のワークスペースと、写し元のリポジトリの setting.config・work\ の名前・大きさ・更新時刻を比べる
#      （中身は読まない）。違いがあれば一覧を出して終了コード 1。
#   5. 使い捨てのフォルダは終わったら消す。起動したプロセスは PID を出す（止めるときは PID で止める）。
param (
    [switch]$Single,
    [hashtable]$Settings = @{},
    [scriptblock]$Command,
    [string]$Workspace,        # 使い捨てのワークスペースの場所（既定は使い捨てのフォルダの中）。外を指すと起動せずに止まる（テスト用）
    [string]$RealWorkspace     # 本物の代わりに比べる場所（テスト用。既定は利用者の本物の既定のワークスペース）
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
    param ([bool]$SingleScript, [hashtable]$Config, [scriptblock]$Action, [string]$WorkspaceDir, [string]$RealDir)

    if (!$RealDir) { $RealDir = getRealDefaultWorkspace }
    $base = newIsolatedFolder
    $tool = Join-Path $base "tool"
    if (!$WorkspaceDir) { $WorkspaceDir = Join-Path $base "workspace" }
    $realBefore = getFolderSnapshot $RealDir
    $configBefore = getFolderSnapshot "$repoRoot\setting.config"
    $workBefore = getFolderSnapshot "$repoRoot\work"
    try {
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
        $outside = @(@($probe.settingsFile, $probe.workDir) | Where-Object { !(testPathInside $_ $base) })
        if ($probe.Count -ne 2 -or $outside.Count -gt 0) {
            Write-Host "書き込み先が使い捨てのフォルダの外を指しているため、起動しません: $($outside -join ', ')" -ForegroundColor Red
            return 1
        }

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
                Write-Host "起動したプロセスの PID: $($p.Id)（画面を閉じるまで待ちます）"
                $p.WaitForExit()
            }
        } finally {
            $env:TEBUNKO_DEFAULT_WORKSPACE = $previous
        }

        $diffs = @(compareFolderSnapshot $realBefore (getFolderSnapshot $RealDir) "本物の既定のワークスペース") +
            @(compareFolderSnapshot $configBefore (getFolderSnapshot "$repoRoot\setting.config") "リポジトリの setting.config") +
            @(compareFolderSnapshot $workBefore (getFolderSnapshot "$repoRoot\work") "リポジトリの work")
        if ($diffs.Count -gt 0) {
            Write-Host "本物のフォルダに違いがあります:" -ForegroundColor Red
            $diffs | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
            return 1
        }
        Write-Host "本物のフォルダに違いなし"
        return 0
    } finally {
        Remove-Item -LiteralPath $base -Recurse -Force -ErrorAction SilentlyContinue
    }
}

exit (invokeIsolated $Single.IsPresent $Settings $Command $Workspace $RealWorkspace)
