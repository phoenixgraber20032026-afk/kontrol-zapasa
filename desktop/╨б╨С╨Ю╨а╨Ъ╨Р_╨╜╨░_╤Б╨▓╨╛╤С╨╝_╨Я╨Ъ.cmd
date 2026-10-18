@echo off
chcp 65001 >nul
echo Сборка установщика «Контроль запаса» на этом компьютере (нужен Node.js 18+ с nodejs.org)
cd /d "%~dp0"
where node >nul 2>nul || (echo Не найден Node.js — установите с https://nodejs.org и запустите снова & pause & exit /b 1)
call npm install --no-audit --no-fund || (pause & exit /b 1)
node ..\tools\build.mjs
call npm run dist || (pause & exit /b 1)
echo.
echo Готово: установщик в папке desktop\dist\KontrolZapasa-Setup-*.exe
pause
