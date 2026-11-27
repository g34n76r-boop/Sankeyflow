@echo off
rem Double-click to remove SankeyFlow from desktop Excel (this Windows user only).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sankeyflow-setup.ps1" -Uninstall
echo.
pause
