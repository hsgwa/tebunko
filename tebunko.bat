@echo off
rem tebunko launcher.
rem Start PowerShell once. In that process:
rem 1) Clear Mark-of-the-Web (added when a downloaded zip is extracted) so that
rem    RemoteSigned does not block the scripts. Inline -Command is not affected by
rem    the execution policy, so this runs even when the mark is present.
rem    Only the scripts folder needs it. The work folder holds the index (tens of
rem    thousands of TSV files), and unblocking it would slow every start.
rem 2) Run the GUI script. RemoteSigned applies to it, and it is no longer marked.
rem    Doing both in one process saves one PowerShell start (1-2 seconds).
rem 3) If the GUI script fails before its own window opens (restricted execution
rem    policy, Constrained Language Mode, missing files, ...), catch it here and
rem    show the reason with Notepad (a message box may not be available in those
rem    cases, and this inline command is not affected by the execution policy or
rem    the language mode). The reason texts are under scripts\tebunko\startup\
rem    (Japanese text does not belong in this ASCII batch file). The record is
rem    saved to a fixed folder (the workspace is not known before startup).
rem    Written only with cmdlets, so it still works in Constrained Language Mode.
rem Run PowerShell through conhost.exe. When Windows Terminal is the default
rem terminal app, it receives the console window and ignores -WindowStyle Hidden,
rem so the window stays open while the GUI runs. conhost.exe keeps the classic
rem console, which -WindowStyle Hidden can hide.
rem PowerShell is not looked up on PATH; it points at the copy that ships with
rem Windows, same as installer\tebunko.cs. If it is missing, Notepad shows the
rem reason and this batch file exits (no record: PowerShell itself is unusable).
set "PS1=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS1%" (
    start "" notepad.exe "%~dp0scripts\tebunko\startup\no_powershell.txt"
    exit /b 1
)
start "" conhost.exe "%PS1%" -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -Command "$opener='notepad.exe';$root='%~dp0';$startup=Join-Path $root 'scripts\tebunko\startup';$gui=Join-Path $root 'scripts\tebunko\gui.ps1';Get-ChildItem -LiteralPath (Join-Path $root 'scripts') -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue; try { & $gui } catch { $err = $_; if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') { $reason = Join-Path $startup 'constrained_language.txt' } elseif (-not (Test-Path -LiteralPath $gui)) { $reason = Join-Path $startup 'missing_files.txt' } elseif ($err.FullyQualifiedErrorId -like 'UnauthorizedAccess*') { $reason = Join-Path $startup 'execution_policy.txt' } else { $reason = Join-Path $startup 'failed.txt' }; $reasonLines = @(if (Test-Path -LiteralPath $reason) { Get-Content -LiteralPath $reason -Encoding UTF8 } else { @('tebunko could not start.') }); $detailLines = @(('==== {0} startup ====' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')), ('Tool: {0}' -f $root), ('LanguageMode: {0}' -f $ExecutionContext.SessionState.LanguageMode), ('PSVersion: {0}' -f $PSVersionTable.PSVersion), (Get-ExecutionPolicy -List | Out-String), ('{0}' -f $err.Exception.Message)); $allLines = $reasonLines + '' + $detailLines; $openTarget = $reason; foreach ($candidate in @((Join-Path $env:LOCALAPPDATA 'tebunko\startup_error.txt'), (Join-Path $env:TEMP 'tebunko_startup_error.txt'))) { try { $dir = Split-Path -Parent $candidate; if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir -ErrorAction Stop | Out-Null }; Set-Content -LiteralPath $candidate -Value $allLines -Encoding UTF8 -ErrorAction Stop; $openTarget = $candidate; break } catch { } }; & $opener $openTarget }"
