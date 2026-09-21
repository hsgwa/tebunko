@echo off
rem tebunko_grep launcher.
rem 1) Clear Mark-of-the-Web (added when a downloaded zip is extracted) so that
rem    RemoteSigned does not block the scripts. Inline -Command is not affected by
rem    the execution policy, so this runs even when the mark is present.
rem    Only the scripts folder needs it. The work folder holds the index (tens of
rem    thousands of TSV files), and unblocking it would slow every start.
rem 2) Start the GUI under RemoteSigned.
powershell -NoProfile -Command "Get-ChildItem -LiteralPath '%~dp0scripts' -Recurse -File | Unblock-File" 1>nul 2>nul
start "" powershell -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "%~dp0scripts\tebunko_grep\gui.ps1"
