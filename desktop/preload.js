// Мост между окном и программой: страница узнаёт, что работает в настольной версии (подпись на экране входа).
const { contextBridge } = require('electron');
contextBridge.exposeInMainWorld('KZW_DESKTOP', true);
