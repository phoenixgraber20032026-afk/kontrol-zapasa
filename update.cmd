@echo off
cd /d "%~dp0"
where node >/dev/null 2>nul
if errorlevel 1 goto nonode
node install.mjs --update
pause
exit /b
:nonode
echo   Node.js is not installed. Download LTS from https://nodejs.org
pause
