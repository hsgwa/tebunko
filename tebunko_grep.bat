@echo off
rem Kept for compatibility with shortcuts made for the previous version (tebunko_grep).
rem It opens the same window as tebunko.bat, starting on the search tab.
rem 1) Clear Mark-of-the-Web so that RemoteSigned does not block the scripts (same as tebunko.bat).
rem 2) Start the GUI under RemoteSigned.
powershell -NoProfile -Command "Get-ChildItem -LiteralPath '%~dp0scripts' -Recurse -File | Unblock-File" 1>nul 2>nul
start "" powershell -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "%~dp0scripts\tebunko\gui.ps1" -StartPage search
