# 実機の確かめを、利用者の既定のワークスペースに触れない形で行う。
#
#   .\tools\run_isolated.ps1                       使い捨てのフォルダにツールの写しを作り、画面を開く（閉じるまで待つ）
#   .\tools\run_isolated.ps1 -Single               展開せずに動く単一 .ps1 版で開く
#   .\tools\run_isolated.ps1 -WaitSeconds 30     画面を 30 秒だけ開き、過ぎたら起動した PID とその子だけを止める
#   .\tools\run_isolated.ps1 -Settings @{ ... }    写しの setting.config に書く項目（targetFolders など）
#   .\tools\run_isolated.ps1 -Command { param ($ToolDir, $Workspace) ... }   画面を開かず、写しの場所と使い捨てのワークスペースを受け取って任意の確かめを流す
#   .\tools\run_isolated.ps1 -ToolDir <展開した配布物の tebunko フォルダ> -Entry indexer -Settings @{ ... }   古い版（フォルダ）を隔てて動かす
#   .\tools\run_isolated.ps1 -ToolScript <単一 .ps1> -Entry indexer -Settings @{ ... }                       古い版（単一 .ps1）を隔てて動かす
#   -Entry gui|indexer で起動するものを選ぶ（既定 gui）。-ToolArgs は起動するものの後ろに足す（空白を含む値は引用符でくくる。今の版の gui には渡せない）。
#   古い版・配布物は、必ずこの道具を通して起動する（TEBUNKO_DEFAULT_WORKSPACE を読まない版もあるため）。
#
# 流れ:
#   1. %TEMP% の下の新しいフォルダに、ツールの写し（scripts\ と tebunko.bat。-Single なら単一 .ps1）と、使い捨てのワークスペースを作る。
#   2. 子プロセスの環境変数 TEBUNKO_DEFAULT_WORKSPACE に、使い捨てのワークスペースを入れる（既定のワークスペースがそこになる）。
#   3. 起動の前に、その環境で設定ファイルの場所と work の場所を求め、どちらも使い捨てのフォルダの中でなければ、起動せずに終了コード 1 で止まる。
#   4. 起動の前後で、既定のワークスペースと、写し元のリポジトリの setting.config・work\ の名前・大きさ・更新時刻を比べる
#      （中身は読まない）。違いがあれば一覧を出して終了コード 1。
#   5. 使い捨てのフォルダは終わったら消す。起動したプロセスは PID を出す（止めるときは PID で止める）。
#
# -ToolDir・-ToolScript のとき（古い版）は、上に加えて次を行う。
#   - 写しの横の setting.config の workspaceFolder に使い捨てのワークスペースを必ず入れる（環境変数を読まない版でも書き先を固定する）。
#     -Settings に workspaceFolder があれば起動しない。
#   - 起動する子プロセスだけ、USERPROFILE・APPDATA・LOCALAPPDATA・TEMP・TMP・TEBUNKO_DEFAULT_WORKSPACE を使い捨ての下に向ける
#     （環境変数を読む版にだけ効く。GetFolderPath を使う版を守るのは、上の workspaceFolder の固定と起動後の前後の比べ）。
#   - -ToolDir は起動の前に、写しの lib.ps1 で設定ファイルと work の場所を求めて確かめる。-ToolScript は読み込んで確かめる口が無いので、
#     起動後の判定で守る。
#   - 起動後、使い捨てのワークスペースにできたファイルの名前を出す。1 つもなければ終了コード 1（よそに書いた合図）。
param (
    [switch]$Single,
    [hashtable]$Settings = @{},
    [scriptblock]$Command,
    [int]$WaitSeconds = 0,     # 画面を待つ秒数（0 は閉じるまで待つ）。過ぎたら、起動した PID とその子だけを止める
    [string]$RealWorkspace,    # 既定のワークスペースの代わりに比べる場所（テスト用。既定は利用者の既定のワークスペース）
    [string]$ToolDir,          # 古い版（展開した配布物の tebunko フォルダ。scripts\ と tebunko.bat がある）。元のフォルダには書かない
    [string]$ToolScript,       # 古い版（単一 .ps1）。-ToolDir とは同時に指定できない
    [ValidateSet("gui", "indexer")]
    [string]$Entry = "gui",    # 起動するもの
    [string[]]$ToolArgs = @()  # 起動するものの後ろに足す引数
)

$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath("$PSScriptRoot\..")
. ([WildcardPattern]::Escape("$PSScriptRoot\isolation\isolation_common.ps1".Replace('`', '``')))

function getIsolatedProbeResult {
    # 使い捨ての環境で、写しの設定ファイルの場所と work の場所を求める（子プロセス。書き込まない）。lib.ps1 が無ければ空
    param ([string]$ScriptsDir, [string]$Base, [string]$WorkspaceDir)

    $lib = Join-Path $ScriptsDir "tebunko\lib.ps1"
    if (!(Test-Path -LiteralPath $lib -PathType Leaf)) { return @{} }
    $probe = Join-Path $Base "probe.ps1"
    $text = @'
param([string]$Lib)
. ([WildcardPattern]::Escape($Lib.Replace('`', '``')))
"settingsFile=" + ${settingsFile}
"workDir=" + (getWorkDir)
'@
    [System.IO.File]::WriteAllText($probe, $text, (New-Object System.Text.UTF8Encoding($true)))
    $previous = $env:TEBUNKO_DEFAULT_WORKSPACE
    $env:TEBUNKO_DEFAULT_WORKSPACE = $WorkspaceDir
    try {
        $lines = @(& powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $probe -Lib $lib)
    } finally {
        $env:TEBUNKO_DEFAULT_WORKSPACE = $previous
    }
    $result = @{}
    foreach ($line in $lines) {
        if ($line -match '^(settingsFile|workDir)=(.*)$') { $result[$Matches[1]] = $Matches[2] }
    }
    return $result
}

function getIsolatedArgumentError {
    # 古い版の指定（-ToolDir・-ToolScript）の誤りを 1 つ返す（無ければ空）。写しを作る前に調べる
    param ([string]$Dir, [string]$Script, [hashtable]$Config, [string]$EntryName = "gui", [string[]]$ArgList = @(), [bool]$HasCommand = $false)

    if (!$Dir -and !$Script) {
        if (!$HasCommand -and $EntryName -eq "gui" -and $ArgList.Count -gt 0) { return "今の版の gui には -ToolArgs を渡せません（-Entry indexer か -ToolDir・-ToolScript で使います）" }
        return ""
    }
    if ($Dir -and $Script) { return "-ToolDir と -ToolScript は同時に指定できません" }
    if ($Dir -and !(Test-Path -LiteralPath (Join-Path $Dir "scripts") -PathType Container)) { return "-ToolDir に scripts\ が見つかりません: $Dir" }
    if ($Script -and !(Test-Path -LiteralPath $Script -PathType Leaf)) { return "-ToolScript のファイルが見つかりません: $Script" }
    if ($Config.ContainsKey("workspaceFolder")) { return "古い版は書き先を使い捨てのワークスペースに固定するため、-Settings に workspaceFolder は指定できません" }
    return ""
}

function useIsolatedChildEnv {
    # 子プロセスが見る環境変数を、使い捨ての下に向ける（元の値を Saved に入れる。restoreIsolatedChildEnv で戻す）
    param ([string]$Base, [string]$ProfileDir, [string]$WorkspaceDir, [System.Collections.IDictionary]$Saved)

    $values = [ordered]@{
        USERPROFILE               = $ProfileDir
        APPDATA                   = "$Base\appdata\Roaming"
        LOCALAPPDATA              = "$Base\appdata\Local"
        TEMP                      = "$Base\temp"
        TMP                       = "$Base\temp"
        TEBUNKO_DEFAULT_WORKSPACE = $WorkspaceDir
    }
    # 先に元の値を全部控える（途中で例外になっても、呼んだ側が Saved で戻せる）
    foreach ($name in $values.Keys) {
        $Saved[$name] = [System.Environment]::GetEnvironmentVariable($name)
    }
    foreach ($name in $values.Keys) {
        [void][System.IO.Directory]::CreateDirectory($values[$name])
        [System.Environment]::SetEnvironmentVariable($name, $values[$name])
    }
}

function restoreIsolatedChildEnv {
    param ($Saved)

    foreach ($name in $Saved.Keys) {
        [System.Environment]::SetEnvironmentVariable($name, $Saved[$name])
    }
}

function invokeIsolated {
    param ([bool]$SingleScript, [hashtable]$Config, [scriptblock]$Action, [int]$Wait, [string]$RealDir, [hashtable]$Legacy = @{})

    if (!$RealDir) { $RealDir = getRealDefaultWorkspace }
    $base = newIsolatedFolder
    $tool = Join-Path $base "tool"
    $isLegacy = [bool]($Legacy.Dir -or $Legacy.Script)
    $profileDir = Join-Path $base "userprofile"
    # 古い版は、使い捨てのプロファイルの下の既定の場所をワークスペースにする。子の USERPROFILE を向けても
    # [Environment]::GetFolderPath("UserProfile") は変わらないので（v0.4.0 はこれで既定の場所を求める）、書き先を守るのは
    # setting.config の workspaceFolder の固定と、起動後の前後の比べ。USERPROFILE を環境変数で読む版だけ、ここに落ちる
    $WorkspaceDir = if ($isLegacy) { Join-Path $profileDir "Documents\tebunko_ws" } else { Join-Path $base "workspace" }
    $exitCode = 0
    $watch = $null
    $savedEnv = [ordered]@{}
    try {
        $watch = newIsolationWatch ([ordered]@{
            "既定のワークスペース" = $RealDir
            "リポジトリの setting.config" = "$repoRoot\setting.config"
            "リポジトリの work" = "$repoRoot\work"
        })
        Write-Host "使い捨てのフォルダ: $base"
        [void][System.IO.Directory]::CreateDirectory($tool)
        $legacyScript = $null
        if ($Legacy.Script) {
            $legacyScript = Join-Path $tool ([System.IO.Path]::GetFileName($Legacy.Script))
            Copy-Item -LiteralPath $Legacy.Script -Destination $legacyScript
        } elseif ($Legacy.Dir) {
            Copy-Item -LiteralPath (Join-Path $Legacy.Dir "scripts") -Destination "$tool\scripts" -Recurse
            $bat = Join-Path $Legacy.Dir "tebunko.bat"
            if (Test-Path -LiteralPath $bat) { Copy-Item -LiteralPath $bat -Destination "$tool\tebunko.bat" }
        } else {
            Copy-Item -LiteralPath "$repoRoot\scripts" -Destination "$tool\scripts" -Recurse
            Copy-Item -LiteralPath "$repoRoot\tebunko.bat" -Destination "$tool\tebunko.bat"
        }
        if ($isLegacy) {
            $Config = $Config.Clone()
            $Config["workspaceFolder"] = $WorkspaceDir
            useIsolatedChildEnv $base $profileDir $WorkspaceDir $savedEnv
        }
        if ($Config.Count -gt 0) {
            [System.IO.File]::WriteAllText("$tool\setting.config", (ConvertTo-Json -InputObject $Config -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
        }

        # 書く前に、書き先が使い捨ての中であることを確かめる（単一 .ps1 は読み込んで確かめる口が無いので、設定で固定した場所を出すだけ）
        if ($Legacy.Script) {
            $probe = @{ settingsFile = "$tool\setting.config"; workDir = $WorkspaceDir }
            Write-Information -InformationAction Continue "設定で固定した work の場所: $WorkspaceDir"
        } else {
            $probe = getIsolatedProbeResult "$tool\scripts" $base $WorkspaceDir
            Write-Information -InformationAction Continue "設定ファイルの場所: $($probe.settingsFile)"
            Write-Information -InformationAction Continue "work の場所: $($probe.workDir)"
        }
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
                    $file = "$tool\scripts\tebunko\$Entry.ps1"
                    $isSingle = $SingleScript -or [bool]$Legacy.Script
                    if ($Legacy.Script) {
                        $file = $legacyScript
                    } elseif ($SingleScript) {
                        $file = Join-Path $tool "tebunko-isolated.ps1"
                        & ([WildcardPattern]::Escape("$repoRoot\tools\new_single_script.ps1".Replace('`', '``'))) -Version "v0.0.0-isolated" -OutFile $file | Out-Null
                    }
                    $launchArgs = getIsolatedLaunchArgs $file $Entry (!$isLegacy -and $Entry -eq "gui") $isSingle @($Legacy.Args)
                    $p = Start-Process powershell.exe -PassThru -WindowStyle Hidden -ArgumentList $launchArgs
                    $null = $p.Handle   # ExitCode を取るため、起動の直後にハンドルを持つ
                    Write-Host "起動したプロセスの PID: $($p.Id)$(if ($Wait -gt 0) { "（$Wait 秒たったら止めます。それまでに画面を閉じれば、そこで終わります）" } else { "（画面を閉じるまで待ちます）" })"
                    $waited = waitIsolationProcess $p $Wait
                    $verdict = getIsolationLaunchVerdict $waited $Wait $p.Id
                    if ($verdict.Line) { Write-Host $verdict.Line -ForegroundColor $(if ($verdict.ExitCode -ne 0) { "Red" } else { "Gray" }) }
                    if ($verdict.ExitCode -ne 0) { $exitCode = 1 }
                    if ($isLegacy) {
                        $made = @(getWorkspaceFileNames $WorkspaceDir)
                        if ($made.Count -eq 0) {
                            Write-Host "使い捨てのワークスペースに、ファイルが 1 つもできていません（よそに書いた可能性があります）: $WorkspaceDir" -ForegroundColor Red
                            $exitCode = 1
                        } else {
                            Write-Information -InformationAction Continue "使い捨てのワークスペースにできたファイル: $($made -join ', ')"
                        }
                    }
                }
            } finally {
                $env:TEBUNKO_DEFAULT_WORKSPACE = $previous
            }
        }
    } finally {
        if ($savedEnv.Count -gt 0) { restoreIsolatedChildEnv $savedEnv }
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

$argumentError = getIsolatedArgumentError $ToolDir $ToolScript $Settings $Entry $ToolArgs ($null -ne $Command)
if ($argumentError) {
    Write-Host $argumentError -ForegroundColor Red
    exit 1
}
exit (invokeIsolated $Single.IsPresent $Settings $Command $WaitSeconds $RealWorkspace @{ Dir = $ToolDir; Script = $ToolScript; Args = $ToolArgs })
