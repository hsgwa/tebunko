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
rem Run PowerShell through conhost.exe. When Windows Terminal is the default
rem terminal app, it receives the console window and ignores -WindowStyle Hidden,
rem so the window stays open while the GUI runs. conhost.exe keeps the classic
rem console, which -WindowStyle Hidden can hide.
start "" conhost.exe powershell -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -Command "Get-ChildItem -LiteralPath '%~dp0scripts' -Recurse -File | Unblock-File -ErrorAction SilentlyContinue; & '%~dp0scripts\tebunko\gui.ps1'"
