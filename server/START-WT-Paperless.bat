@echo off
chcp 65001 >nul
title WT Paperless - Warehouse and Transportation
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 goto NONODE

node "WT-Paperless-Server.js"
echo.
echo   Server stopped. Press any key to close this window.
pause >nul
exit /b

:NONODE
echo.
echo   ============================================================
echo    Node.js is required but was not found on this computer.
echo   ============================================================
echo.
echo    1. Install Node.js LTS from  https://nodejs.org
echo    2. Then double-click this file again.
echo.
echo    Without Node.js you can still open WT-Paperless-WebApp.html
echo    directly - data will be stored in the browser instead of Excel.
echo.
pause
