@echo off
chcp 65001 >nul
cd /d "%~dp0"
where node >nul 2>nul || (echo Нужен Node.js 18+ — установите с https://nodejs.org и запустите снова & pause & exit /b 1)
node install.mjs
pause
