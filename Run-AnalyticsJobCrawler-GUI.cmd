@echo off
setlocal
cd /d "%~dp0"

start "" powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Launch-AnalyticsJobCrawlerGui.ps1"
exit /b 0
