# Контроль неснижаемого запаса

Программа контроля складского запаса по выгрузке 1С «Остатки по складам»: светофор статусов, критический запас,
группы замен, аналоги, заказ, предупреждения. **Данные хранятся в Google-таблице** («Контроль_неснижаемого_запаса»,
скрипт KontrolZapasa v2.0 + KZ_Web), а работать можно тремя способами — везде одна и та же программа:

| Способ | Как открыть | Для кого |
|---|---|---|
| Внутри Google | ссылка веб-приложения Apps Script `…/exec` | как раньше |
| **Сайт** | `https://<логин>.github.io/kontrol-zapasa/` (папка `app/`, GitHub Pages) | все сотрудники, телефон; в браузере можно «Установить как приложение» (PWA) |
| **Настольная программа** | установщик `KontrolZapasa-Setup-x.y.z.exe` из Releases | Windows; обновляется сама |

Уровни доступа те же: **ПРОСМОТР** без пароля, **WIZARD** по паролю. Сайт и программа ничего не хранят на своих
серверах: страница только отправляет запросы в ваш Apps Script (`POST …/exec`, `{"api":"kzwXxx","args":[…]}`).

## Структура

```
gas/        исходники для Apps Script: KZ_Web.gs (сервер), KZ_WebApp.html (интерфейс — единый исходник),
            appsscript.json; сюда же clasp pull кладёт основной скрипт KontrolZapasa
app/        сайт (GitHub Pages): index.html = копия gas/KZ_WebApp.html (собирает tools/build.mjs),
            chart.umd.min.js, manifest.json, sw.js, иконки
desktop/    настольная версия (Electron + electron-builder): main.js, package.json, build/icon.ico
tools/      build.mjs — сборка app/ из gas/ и синхронизация версий
.github/    workflows: pages.yml (сайт), desktop.yml (установщик Windows по метке v*)
```

Подробная пошаговая инструкция — `УСТАНОВКА_GitHub_и_exe.md`.

## Выпуск новой версии (коротко)
1. Правим `gas/KZ_WebApp.html` / `gas/KZ_Web.gs` (номер в `KZW.VERSION`), `clasp push` в Apps Script → новая версия развёртывания.
2. `node tools/build.mjs` → commit → push в `main` — сайт обновится сам (Actions → Pages).
3. `git tag v1.7.0 && git push --tags` — Actions соберёт `KontrolZapasa-Setup-1.7.0.exe` и опубликует в Releases;
   установленные программы предложат обновление.
