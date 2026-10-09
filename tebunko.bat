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
rem    saved in the tool folder (the workspace is not known before startup).
rem    Written only with cmdlets, so it still works in Constrained Language Mode.
rem    The -Command text is built below in PSCMD, piece by piece (one set per
rem    step), so no single line is too long to read. No %% in any piece other
rem    than %%~dp0 (expanded once, here, into the PSCMD value below). No delayed
rem    expansion; each set line expands %%PSCMD%% to its value so far, in order.
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
rem Step 1: the notepad opener (kept as a single assignment so a test can
rem replace it) and the tool's folder paths, used by every later step.
set "PSCMD=$opener='notepad.exe';$root='%~dp0';$startup=Join-Path $root 'scripts\tebunko\startup';$gui=Join-Path $root 'scripts\tebunko\gui.ps1';"
rem Step 2: clear Mark-of-the-Web from the scripts folder (see the comment above).
set "PSCMD=%PSCMD%Get-ChildItem -LiteralPath (Join-Path $root 'scripts') -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue; "
rem Step 3: run the GUI script, catching a failure from before its own window opens.
set "PSCMD=%PSCMD%try { & $gui } catch { $err = $_; "
rem Step 4: in the catch, pick the reason (missing gui.ps1, Constrained Language
rem Mode, execution policy, or something else) and its text file.
set "PSCMD=%PSCMD%if (-not (Test-Path -LiteralPath $gui)) { $reason = Join-Path $startup 'missing_files.txt' } elseif ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') { $reason = Join-Path $startup 'constrained_language.txt' } elseif ($err.FullyQualifiedErrorId -like 'UnauthorizedAccess*') { $reason = Join-Path $startup 'execution_policy.txt' } else { $reason = Join-Path $startup 'failed.txt' }; $reasonLines = @(if (Test-Path -LiteralPath $reason) { Get-Content -LiteralPath $reason -Encoding UTF8 } else { @('tebunko could not start.') }); "
rem Step 5: build the record's detail lines (date, tool path, language mode,
rem PowerShell version, execution policy, and the error message).
set "PSCMD=%PSCMD%$detailLines = @(('==== {0} startup ====' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')), ('Tool: {0}' -f $root), ('LanguageMode: {0}' -f $ExecutionContext.SessionState.LanguageMode), ('PSVersion: {0}' -f $PSVersionTable.PSVersion), (Get-ExecutionPolicy -List | Out-String), ('{0}' -f $err.Exception.Message)); $allLines = $reasonLines + '' + $detailLines; "
rem Step 6: write the record in the tool folder (no record if it cannot be written; then the reason is shown), then show it.
set "PSCMD=%PSCMD%$openTarget = $reason; $candidate = Join-Path $root 'startup_error.txt'; try { Set-Content -LiteralPath $candidate -Value $allLines -Encoding UTF8 -ErrorAction Stop; $openTarget = $candidate } catch { }; & $opener $openTarget }"
start "" conhost.exe "%PS1%" -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -Command "%PSCMD%"
