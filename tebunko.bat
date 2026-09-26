@echo off
rem tebunko launcher.
rem 1) Clear Mark-of-the-Web (added when a downloaded zip is extracted) so that
rem    RemoteSigned does not block the scripts. Inline -Command is not affected by
rem    the execution policy, so this runs even when the mark is present.
rem    Only the scripts folder needs it. The work folder holds the index (tens of
rem    thousands of TSV files), and unblocking it would slow every start.
rem 2) Start the GUI under RemoteSigned. Run PowerShell through conhost.exe.
rem    When Windows Terminal is the default terminal app, it receives the
rem    console window and ignores -WindowStyle Hidden, so the window stays
rem    open while the GUI runs. conhost.exe keeps the classic console, which
rem    -WindowStyle Hidden can hide.
powershell -NoProfile -Command "Get-ChildItem -LiteralPath '%~dp0scripts' -Recurse -File | Unblock-File" 1>nul 2>nul
start "" conhost.exe powershell -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "%~dp0scripts\tebunko\gui.ps1"
