@echo off
rem tebunko_grep launcher.
rem 1) Clear Mark-of-the-Web (added when a downloaded zip is extracted) so that
rem    RemoteSigned does not block the scripts. Inline -Command is not affected by
rem    the execution policy, so this runs even when the mark is present.
rem    Only the scripts folder needs it. The work folder holds the index (tens of
rem    thousands of TSV files), and unblocking it would slow every start.
rem 2) Check in this console whether the GUI can run (start.ps1). start.ps1 also
rem    runs in Constrained Language Mode (AppLocker/WDAC), where gui.ps1 stops at
rem    its first line and shows nothing. There start.ps1 shows the reason here
rem    and starts the restricted mode in this console.
rem    Exit code 10: the GUI can run. 0: the restricted mode ended.
rem    12: start.ps1 failed and already showed the error.
rem    Any other code: start.ps1 itself could not run (for example, a group policy
rem    sets AllSigned). Show start_failed.txt with an inline -Command, which the
rem    execution policy does not block. The bat stays ASCII (cmd reads it in the
rem    ANSI code page), so the Japanese text is in that file.
rem 3) Start the GUI under RemoteSigned. Run PowerShell through conhost.exe.
rem    When Windows Terminal is the default terminal app, it receives the
rem    console window and ignores -WindowStyle Hidden, so the window stays
rem    open while the GUI runs. conhost.exe keeps the classic console, which
rem    -WindowStyle Hidden can hide.
powershell -NoProfile -Command "Get-ChildItem -LiteralPath '%~dp0scripts' -Recurse -File | Unblock-File" 1>nul 2>nul
powershell -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0scripts\tebunko_grep\start.ps1"
if %errorlevel% equ 10 goto gui
if %errorlevel% equ 0 goto end
if %errorlevel% equ 12 goto wait
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0scripts\tebunko_grep\start_failed.txt' -Encoding UTF8; Get-ExecutionPolicy -List | Format-Table -AutoSize | Out-String"
:wait
pause
goto end
:gui
start "" conhost.exe powershell -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "%~dp0scripts\tebunko_grep\gui.ps1"
:end
