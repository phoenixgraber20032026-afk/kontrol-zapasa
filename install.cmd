@echo off
cd /d "%~dp0"
where node >/dev/null 2>nul
if errorlevel 1 goto nonode
node install.mjs
pause
exit /b
:nonode
echo.
echo   Node.js is not installed. Download LTS from https://nodejs.org , install, then run install.cmd again.
echo   (Node.js ne ustanovlen: skachaite LTS s https://nodejs.org i zapustite install.cmd snova)
echo.
pause
