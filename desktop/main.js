/**
 * Контроль неснижаемого запаса — настольная версия (Electron).
 * Окно показывает ту же страницу app/index.html, что и сайт; данные ходят в Google-таблицу через
 * веб-приложение Apps Script (адрес /exec задаётся один раз на экране входа и хранится в профиле программы).
 * Обновления: при наличии нового выпуска на GitHub (Releases) программа скачивает его сама и предлагает перезапуск.
 */
const { app, BrowserWindow, Menu, shell, dialog, session } = require('electron');
const path = require('path');

let win = null;
const APP_TITLE = 'Контроль неснижаемого запаса';
const SITE_URL = 'https://phoenixgraber20032026-afk.github.io/kontrol-zapasa/';   // адрес сайта на GitHub Pages (например https://<логин>.github.io/kontrol-zapasa/app/) — появится пункт меню «Открыть в браузере»

function createWindow() {
  win = new BrowserWindow({
    width: 1500, height: 940, minWidth: 900, minHeight: 600, show: false,
    title: APP_TITLE, backgroundColor: '#f4f6f9', autoHideMenuBar: false,
    icon: path.join(__dirname, 'build', 'icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: false }
  });
  win.loadFile(path.join(__dirname, 'app', 'index.html'));
  win.once('ready-to-show', () => { win.show(); });
  win.on('page-title-updated', (e, t) => { if (!/Контроль/.test(t)) e.preventDefault(); });
  // внешние ссылки (таблица Google, сайты) — в обычном браузере
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('file:')) { e.preventDefault(); shell.openExternal(url); } });
  win.on('closed', () => { win = null; });
}

function js(code) { if (win) win.webContents.executeJavaScript(code).catch(() => {}); }

function buildMenu() {
  const tpl = [
    { label: 'Программа', submenu: [
      { label: 'Обновить данные  (F5)', accelerator: 'F5', click: () => js("document.getElementById('refreshBtn') && document.getElementById('app').style.display !== 'none' ? document.getElementById('refreshBtn').click() : location.reload()") },
      { label: 'Перезагрузить страницу', accelerator: 'CmdOrCtrl+R', click: () => win && win.reload() },
      { type: 'separator' },
      { label: 'Адрес программы Google…', click: () => js("(function(){try{localStorage.removeItem('kzw_api');localStorage.removeItem('kzw_token');}catch(e){}location.reload();})()") },
      ...(SITE_URL ? [{ label: 'Открыть в браузере (сайт)', click: () => shell.openExternal(SITE_URL) }] : []),
      { type: 'separator' },
      { label: 'Выход', role: 'quit' }
    ] },
    { label: 'Вид', submenu: [
      { label: 'Полный экран', role: 'togglefullscreen' },
      { label: 'Крупнее', role: 'zoomIn' }, { label: 'Мельче', role: 'zoomOut' }, { label: 'Обычный масштаб', role: 'resetZoom' },
      { type: 'separator' },
      { label: 'Инструменты разработчика', role: 'toggleDevTools' }
    ] },
    { label: 'Справка', submenu: [
      { label: 'Справка по программе', click: () => js("(function(){var t=document.querySelector('.tab[data-tab=\"help\"]'); if(t) t.click();})()") },
      { label: 'Проверить обновления', click: () => checkUpdates(true) },
      { label: 'О программе', click: () => dialog.showMessageBox(win, { type: 'info', title: 'О программе', message: APP_TITLE, detail: 'Версия ' + app.getVersion() + '\nДанные хранятся в Google-таблице «Контроль_неснижаемого_запаса».\nООО «Инструмент-Ресурс»' }) }
    ] }
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(tpl));
}

function checkUpdates(manual) {
  if (!app.isPackaged) { if (manual) dialog.showMessageBox(win, { message: 'Проверка обновлений работает только в установленной версии.' }); return; }
  let updater;
  try { updater = require('electron-updater').autoUpdater; } catch (e) { return; }
  updater.autoDownload = true;
  updater.on('update-downloaded', (info) => {
    dialog.showMessageBox(win, { type: 'question', buttons: ['Перезапустить сейчас', 'Позже'], defaultId: 0, title: 'Обновление', message: 'Загружена новая версия ' + info.version, detail: 'Перезапустить программу, чтобы установить её?' })
      .then(r => { if (r.response === 0) updater.quitAndInstall(); });
  });
  updater.on('error', (e) => { if (manual) dialog.showMessageBox(win, { type: 'warning', message: 'Не удалось проверить обновления', detail: String(e && e.message || e) }); });
  updater.on('update-not-available', () => { if (manual) dialog.showMessageBox(win, { message: 'У вас последняя версия ' + app.getVersion() }); });
  updater.checkForUpdates().catch(() => {});
}

app.setAppUserModelId('by.air.kontrol-zapasa');
app.whenReady().then(() => {
  // разрешаем звук предупреждения без клика
  session.defaultSession.setPermissionRequestHandler((wc, permission, cb) => cb(permission === 'notifications' || permission === 'media'));
  createWindow(); buildMenu();
  setTimeout(() => checkUpdates(false), 8000);
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('window-all-closed', () => { app.quit(); });
app.commandLine.appendSwitch('autoplay-policy', 'no-user-gesture-required');
