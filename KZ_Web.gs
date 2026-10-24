/**
 * ============================================================================
 *  КОНТРОЛЬ НЕСНИЖАЕМОГО ЗАПАСА — ВЕБ-ПРИЛОЖЕНИЕ (серверная часть)  v2.0
 * ============================================================================
 *  Файл добавляется в проект Apps Script таблицы «Контроль_неснижаемого_запаса»
 *  РЯДОМ с основным скриптом (Code.gs, KontrolZapasa v2.0) и файлом KZ_WebApp.html.
 *  Ничего из основного скрипта не переопределяет: использует его константы
 *  (SH_*, C_*, WH_NAMES, TYPES, KZ) и функции (kzRun_, kzLoadState_, kzClearState_,
 *  convertToGoogleSheet_, normName_, toDate_, keyStr_, syncXref_, ensureGroupPriority_,
 *  ensureXrefSheet_, kzCtx_, detectSeparator_, kzFlushSig_, ensureRows_, lastRowByCol_, fmtDate_).
 *
 *  Уровни доступа:  ПРОСМОТР — без пароля (всё видно, ничего не меняется);
 *                   WIZARD   — пароль (по умолчанию 2801; меняется в приложении),
 *                              загрузка 1С, пересборка, минимумы/группы, правка листов, конструктор.
 *
 *  Развёртывание: редактор скриптов → «Начать развёртывание» → «Новое развёртывание» →
 *  тип «Веб-приложение», «Выполнять от моего имени», доступ «Все». После каждого
 *  изменения кода — «Управление развёртываниями» → ✎ → Версия: «Новая версия» → «Развернуть».
 *  Для загрузки Excel из приложения в проекте должна быть подключена служба «Drive API»
 *  (Службы → + → Drive API) — она же нужна основному скрипту для меню «Загрузить остатки из 1С».
 *
 *  Необязательно: в onOpen основного скрипта можно добавить пункт меню
 *     .addItem('Веб-приложение: ссылка', 'kzwShowWebAppUrl')
 */

var KZW = {
  VERSION: '2.0',
  HTML_FILE: 'KZ_WebApp',
  TITLE: 'Контроль неснижаемого запаса',
  TOKEN_TTL_MS: 12 * 3600 * 1000,
  DEFAULT_PASS: { wizard: '2801' },
  MAX_ROWS: 3000, MAX_COLS: 60,
  SHEET_PAGES: 'Веб-страницы',
  PAGES_HEADER: ['ID', 'Название', 'Иконка', 'Порядок', 'Доступ (viewer,wizard)', 'Вкладка в меню (Да/Нет)', 'Содержимое (JSON блоков)'],
  HIDDEN_SHEETS: ['Данные 1С', 'Номенклатура', 'Омонимы', 'Снимки остатков'],       // служебные (большие) листы — по умолчанию не показываются; WIZARD может включить в конструкторе
  PROP_PENDING: 'KZW_PENDING',                                     // параметры отложенной пересборки (запуск по триггеру)
  TRIGGER_FN: 'kzwRunRebuildTrigger',
  HISTORY_MAX: 400
};
var KZW_RANK = { viewer: 1, wizard: 2 };

// ─────────────────────────────────────────────────────────────── вход в приложение
function doGet(e) {
  if (e && e.parameter && e.parameter.ping) return kzwJson_({ ok: true, app: 'kz', ver: KZW.VERSION, title: KZW.TITLE, time: kzwDateTimeStr_(new Date()) });   // v1.7: проверка адреса из режима «сайт»
  var t = HtmlService.createTemplateFromFile(KZW.HTML_FILE);
  return t.evaluate().setTitle(KZW.TITLE).addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * v1.7 — режим «сайт» / настольная программа: страница размещена вне Google (GitHub Pages, .exe),
 * а данные ходят POST-запросами на /exec: тело {"api":"kzwXxx","args":[…]} (text/plain — без preflight),
 * ответ {"ok":true,"result":…} или {"ok":false,"error":"…"}. Разрешены только публичные функции kzw[A-Z]…
 * (без «_»), права внутри них проверяет kzwAuth_ по токену — как и при вызове через google.script.run.
 */
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return kzwJson_({ ok: false, error: 'Неверный формат запроса' }); }
  if (!body || typeof body.api !== 'string') return kzwJson_({ ok: false, error: 'Нет имени функции (api)' });
  return kzwJson_(kzwApi_(body.api, body.args || []));
}
function kzwApi_(name, args) {
  if (!/^kzw[A-Z][A-Za-z0-9]*$/.test(name) || name === 'kzwRunRebuildTrigger') return { ok: false, error: 'Функция недоступна: ' + name };
  var G = (typeof globalThis !== 'undefined') ? globalThis : this, fn = G[name];
  if (typeof fn !== 'function') return { ok: false, error: 'Функция не найдена: ' + name };
  try { var r = fn.apply(null, Array.isArray(args) ? args : []); return { ok: true, result: r === undefined ? null : r }; }
  catch (err) { return { ok: false, error: String(err && err.message || err) }; }
}
function kzwJson_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/** Меню / редактор: показать ссылку на веб-приложение. */
function kzwShowWebAppUrl() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  var text = url ? ('Ссылка на веб-приложение: ' + url) : 'Веб-приложение ещё не развёрнуто: редактор скриптов → «Начать развёртывание» → «Новое развёртывание» → «Веб-приложение» (выполнять от моего имени, доступ — все).';
  try {
    var ui = SpreadsheetApp.getUi();
    if (!url) { ui.alert(text); return; }
    var html = HtmlService.createHtmlOutput('<div style="font:13px Arial;padding:10px">Ссылка для пользователей:<br><br><input style="width:100%;font:13px monospace" value="' + url + '" onclick="this.select()" readonly>' +
      '<br><br><a href="' + url + '" target="_blank">Открыть</a> &nbsp;·&nbsp; ПРОСМОТР — без пароля, WIZARD — по паролю (меняется в приложении).</div>').setWidth(560).setHeight(160);
    ui.showModalDialog(html, 'Веб-приложение «' + KZW.TITLE + '»');
  } catch (err) { console.log(text); }
}

// ─────────────────────────────────────────────────────────────── авторизация
function kzwPass_() {
  var p = PropertiesService.getScriptProperties().getProperty('KZW_PASS_WIZARD');
  return p === null ? KZW.DEFAULT_PASS.wizard : p;
}
function kzwLogin(password) {
  password = String(password || '').trim();
  var role = 'viewer';
  if (password) { if (password === kzwPass_()) role = 'wizard'; else return { error: 'Неверный пароль' }; }
  return { token: kzwMintToken_(role), role: role };
}
/** Новый токен сессии для роли (вход по паролю или через Битрикс24); попутно чистит просроченные. */
function kzwMintToken_(role, extra) {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var props = PropertiesService.getScriptProperties(), all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('KZW_TOK_') === 0) { try { if (JSON.parse(all[k]).exp < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } }
  });
  var s = { role: role, exp: now + KZW.TOKEN_TTL_MS }; if (extra) Object.keys(extra).forEach(function (k) { s[k] = extra[k]; });
  props.setProperty('KZW_TOK_' + token, JSON.stringify(s));
  return token;
}
function kzwLogout(token) { try { PropertiesService.getScriptProperties().deleteProperty('KZW_TOK_' + token); } catch (e) {} return true; }
function kzwAuth_(token, minRole) {
  if (token === '__internal__') return 'wizard';   // v2.0: вызовы внутри скрипта (советник, сводка, резерв)
  var raw = token ? PropertiesService.getScriptProperties().getProperty('KZW_TOK_' + token) : null;
  if (!raw) throw new Error('AUTH: сессия истекла — войдите заново');
  var s = JSON.parse(raw);
  if (s.exp < Date.now()) throw new Error('AUTH: сессия истекла — войдите заново');
  if (minRole && KZW_RANK[s.role] < KZW_RANK[minRole]) throw new Error('Недостаточно прав: нужен уровень WIZARD');
  if (minRole === 'wizard') kzwEdgeBump_();   // v1.8: любая правка WIZARD → ускоритель сбрасывает кэш
  return s.role;
}
function kzwSetPassword(token, wizardPass) {
  kzwAuth_(token, 'wizard');
  if (!wizardPass || !String(wizardPass).trim()) throw new Error('Пустой пароль');
  PropertiesService.getScriptProperties().setProperty('KZW_PASS_WIZARD', String(wizardPass).trim());
  return 'Пароль WIZARD сохранён';
}

// ─────────────────────────────────────────────────────────────── служебное
function kzwDateStr_(d) { return (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''; }
function kzwDateTimeStr_(d) { return (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(d == null ? '' : d); }
function kzwNum_(v) { if (typeof v === 'number') return v; var n = parseFloat(String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.')); return isNaN(n) ? 0 : n; }
function kzwSs_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function kzwSheet_(ss, name) { var sh = ss.getSheetByName(name); if (!sh) throw new Error('Лист не найден: ' + name); return sh; }
/** Ключ как строка для сопоставления «Остатки» ↔ «Справочник» (даты/числа — как в keyStr_ основного скрипта). */
function kzwKey_(v) { return (typeof keyStr_ === 'function') ? keyStr_(v) : String(v == null ? '' : v).trim(); }

// ─────────────────────────────────────────────────────────────── первичная загрузка
function kzwBootstrap(token, password) {
  if (!token) { var lg = kzwLogin(password); if (lg.error) return lg; token = lg.token; }
  var role = kzwAuth_(token);
  var ss = kzwSs_();
  var res = {
    token: token, role: role, title: KZW.TITLE, today: kzwDateStr_(new Date()),
    sheets: kzwVisibleSheets_(ss, role), allSheets: ss.getSheets().map(function (s) { return s.getName(); }),
    pages: kzwPages_(ss, role), webSettings: kzwSettings_(ss),
    whNames: WH_NAMES.slice(), whOut: WH_OUT, types: TYPES.map(function (T) { return T.t; }),
    sheetNames: { ost: SH_OST, spr: SH_SPR, grp: SH_GRP, set: SH_SET, hist: SH_HI, crit: SH_CRIT, sum: SH_SUM, xref: SH_XREF, stat: SH_ST, raw: SH_RAW },
    scriptVersion: (typeof KZ !== 'undefined' && KZ.VERSION) || ''
  };
  res.data = kzwGetData(token);
  return res;
}
/** v1.8: то же без пакета данных — для ускорителя (мета по роли отдельно от data). */
function kzwBootMeta(token) {
  var role = kzwAuth_(token), ss = kzwSs_();
  return {
    role: role, title: KZW.TITLE, today: kzwDateStr_(new Date()),
    sheets: kzwVisibleSheets_(ss, role), allSheets: ss.getSheets().map(function (s) { return s.getName(); }),
    pages: kzwPages_(ss, role), webSettings: kzwSettings_(ss),
    whNames: WH_NAMES.slice(), whOut: WH_OUT, types: TYPES.map(function (T) { return T.t; }),
    sheetNames: { ost: SH_OST, spr: SH_SPR, grp: SH_GRP, set: SH_SET, hist: SH_HI, crit: SH_CRIT, sum: SH_SUM, xref: SH_XREF, stat: SH_ST, raw: SH_RAW },
    scriptVersion: (typeof KZ !== 'undefined' && KZ.VERSION) || ''
  };
}

/** Все данные для приложения одним пакетом: позиции «Остатков» + комментарии «Справочника», группы, история, настройки. */
function kzwGetData(token) {
  kzwAuth_(token);
  var ss = kzwSs_();
  var out = { items: [], groups: [], history: [], settings: kzwReadSettings_(ss), loadedAt: kzwDateTimeStr_(new Date()) };
  try { kzwApplyCritFormulas_(ss, false); kzwFreezeHeaders_(ss, false); } catch (eC) {}   // v1.2: формулы с критическим запасом и закреплённые шапки — восстанавливаются после пересборки
  // Справочник: комментарий и статус по ключу
  var spr = {}, wsS = ss.getSheetByName(SH_SPR);
  if (wsS && wsS.getLastRow() >= 2) {
    var sprCols = Math.min(wsS.getMaxColumns(), KZW_EXCL.COL), sv = wsS.getRange(2, 1, wsS.getLastRow() - 1, sprCols).getValues();
    for (var i = 0; i < sv.length; i++) { var k = kzwKey_(sv[i][0]); if (k && !spr[k]) spr[k] = { c: String(sv[i][9] == null ? '' : sv[i][9]), s: String(sv[i][11] == null ? '' : sv[i][11]), r: i + 2, crit: (sv[i][KZW_CRIT.COL - 1] === '' || sv[i][KZW_CRIT.COL - 1] === null || sv[i][KZW_CRIT.COL - 1] === undefined) ? null : kzwNum_(sv[i][KZW_CRIT.COL - 1]), excl: String(sv[i][KZW_EXCL.COL - 1] == null ? '' : sv[i][KZW_EXCL.COL - 1]).trim() }; }
  }
  // Остатки
  var wsO = kzwSheet_(ss, SH_OST), lr = wsO.getLastRow();
  if (lr >= 2) {
    var v = wsO.getRange(2, 1, lr - 1, OST_COLS).getValues();
    for (var r = 0; r < v.length; r++) {
      var row = v[r]; if (row[0] === '' || row[0] === null) continue;
      var key = kzwKey_(row[0]), sp = spr[key] || { c: '', s: '', r: 0, crit: null, excl: '' };
      var wh = []; for (var j = 0; j < N_WH; j++) wh.push(kzwNum_(row[C_WH0 - 1 + j]));
      out.items.push([
        key, kzwKey_(row[C_ART - 1]), String(row[C_NAME - 1] == null ? '' : row[C_NAME - 1]), String(row[C_BRAND - 1] || '') || 'Не определено',
        String(row[C_CAT - 1] || ''), String(row[C_UNIT - 1] || ''), kzwNum_(row[C_TOTAL - 1]), wh, kzwNum_(row[C_OUT - 1]),
        kzwDateStr_(row[C_OLD - 1]), kzwDateStr_(row[C_NEW - 1]), Math.round(kzwNum_(row[C_PRICE - 1]) * 100) / 100, Math.round(kzwNum_(row[C_VALUE - 1]) * 100) / 100,
        kzwNum_(row[C_ROWS - 1]), (row[C_MIN - 1] === '' || row[C_MIN - 1] === null) ? null : kzwNum_(row[C_MIN - 1]),
        (row[C_TGT - 1] === '' || row[C_TGT - 1] === null) ? null : kzwNum_(row[C_TGT - 1]), kzwNum_(row[C_DEF - 1]),
        String(row[C_GRP - 1] || '').trim(), String(row[C_STAT - 1] || ''), kzwNum_(row[C_PRIO - 1]), String(row[C_TYPE - 1] || ''),
        sp.c, sp.s, r + 2, sp.r, sp.crit === undefined ? null : sp.crit, sp.excl || ''
      ]);
    }
  }
  out.itemCols = ['key', 'art', 'name', 'brand', 'cat', 'unit', 'tot', 'wh', 'out', 'old', 'nw', 'price', 'value', 'rows', 'min', 'tgt', 'def', 'grp', 'stat', 'prio', 'type', 'comment', 'sprStat', 'ostRow', 'sprRow', 'crit', 'excl'];
  // Группы замен
  var wsG = ss.getSheetByName(SH_GRP);
  if (wsG && wsG.getLastRow() >= 2) {
    var gv = wsG.getRange(2, 1, wsG.getLastRow() - 1, 8).getValues();
    for (var g = 0; g < gv.length; g++) {
      var id = String(gv[g][0] || '').trim(); if (!id) continue;
      out.groups.push({ id: id, purpose: String(gv[g][1] || ''), min: (gv[g][2] === '' || gv[g][2] === null) ? null : kzwNum_(gv[g][2]), comment: String(gv[g][3] || ''),
        stock: kzwNum_(gv[g][4]), count: kzwNum_(gv[g][5]), stat: String(gv[g][6] || ''), prio: kzwNum_(gv[g][7]), row: g + 2 });
    }
  }
  // История
  var wsH = ss.getSheetByName(SH_HI);
  if (wsH && wsH.getLastRow() >= 5) {
    var hv = wsH.getRange(5, 1, wsH.getLastRow() - 4, 12).getValues();
    for (var h = 0; h < hv.length; h++) {
      if (!hv[h][2]) continue;
      out.history.push([kzwDateTimeStr_(hv[h][0]), String(hv[h][1] || ''), String(hv[h][2] || '')].concat(hv[h].slice(3, 12).map(kzwNum_)));
    }
    if (out.history.length > KZW.HISTORY_MAX) out.history = out.history.slice(out.history.length - KZW.HISTORY_MAX);
  }
  out.state = kzwStateInfo_(ss);
  try { if (!out.state.busy && kzwSnapshot_(ss, out.items.map(function (a) { return { key: a[0], tot: a[6], value: a[12], stat: a[18] }; })) > 0) { kzwB24Notify_(ss, out); try { kzwOrdersAutoClose_(ss, out.items); } catch (e1) {} try { kzwReserveRefresh_(ss); } catch (e2) {} } } catch (eS) {}   // v1.3: снимок остатков на дату загрузки; v1.9: уведомление в Битрикс24; v2.0: автозакрытие заказов, резерв
  return out;
}

/** Настройки: пороги, последняя загрузка, таблица складов. */
function kzwReadSettings_(ss) {
  var ws = ss.getSheetByName(SH_SET), s = { yellow: 0.2, defMin: null, lastLoad: '', lastFile: '', state: '', warehouses: [], whHeaderRow: 9 };
  if (!ws) return s;
  var v = ws.getRange(1, 1, 6, 3).getValues();
  s.yellow = kzwNum_(v[1][1]); s.defMin = (v[2][1] === '' || v[2][1] === null) ? null : kzwNum_(v[2][1]);
  s.lastLoad = kzwDateTimeStr_(v[3][1]); s.lastFile = String(v[4][1] || ''); s.state = String(v[5][1] || '');
  var h = (typeof settingsHeaderRow_ === 'function') ? settingsHeaderRow_(ws) : 9, last = ws.getLastRow();
  s.whHeaderRow = h;
  if (last > h) {
    var wv = ws.getRange(h + 1, 1, last - h, 4).getValues();
    for (var i = 0; i < wv.length; i++) {
      var nm = String(wv[i][0] || '').trim(); if (!nm) continue;
      s.warehouses.push({ row: h + 1 + i, name: nm, group: String(wv[i][1] || '').trim() || 'Прочие', inc: (typeof whIsYes_ === 'function') ? whIsYes_(wv[i][2]) : String(wv[i][2]).toLowerCase() !== 'нет', count: kzwNum_(wv[i][3]) });
    }
  }
  return s;
}
/** Состояние пересборки: есть ли незавершённая, текст из Настройки!B6, ждёт ли отложенный запуск. */
function kzwStateInfo_(ss) {
  var st = null; try { st = kzLoadState_(); } catch (e) {}
  var text = ''; try { text = String(ss.getSheetByName(SH_SET).getRange('B6').getValue() || ''); } catch (e2) {}
  var pending = null; try { pending = PropertiesService.getScriptProperties().getProperty(KZW.PROP_PENDING); } catch (e3) {}
  // реально ли что-то выполняется: kzRun_ держит блокировку документа; продолжение после паузы — триггер kzContinue
  var running = false; try { var lock = LockService.getDocumentLock(); if (lock.tryLock(0)) lock.releaseLock(); else running = true; } catch (e4) {}
  var hasTrigger = false; try { hasTrigger = ScriptApp.getProjectTriggers().some(function (t) { var h = t.getHandlerFunction(); return h === KZ.CONTINUE_FN || h === KZW.TRIGGER_FN; }); } catch (e5) {}
  // защита от «вечной пересборки»: отложенный запуск старше 15 мин или триггер продолжения при состоянии старше 40 мин — считаем зависшими
  var now = Date.now(), pendAge = 0, stAge = 0;
  try { pendAge = pending ? (now - (JSON.parse(pending).at || 0)) / 60000 : 0; } catch (e6) { pendAge = 999; }
  try { stAge = st && st.started ? (now - new Date(st.started).getTime()) / 60000 : 0; } catch (e7) {}
  var pendOk = !!pending && pendAge < 15, trigOk = hasTrigger && !!st && stAge < 40;
  var busy = running || pendOk || trigOk;
  var stale = !busy && (!!st || !!pending || hasTrigger);
  return { busy: busy, stale: stale, running: running, hasTrigger: hasTrigger, stage: st ? st.stage : null, done: st ? st.done : [], started: st ? st.started : '',
    ageMin: Math.round(stAge), pendAgeMin: Math.round(pendAge), text: text, pending: !!pending, paused: !!st && !pending && /Пауза/.test(text) };
}
function kzwStatus(token) {
  kzwAuth_(token);
  var ss = kzwSs_(), s = kzwStateInfo_(ss);
  try { var ws = ss.getSheetByName(SH_SET); s.lastLoad = kzwDateTimeStr_(ws.getRange('B4').getValue()); s.lastFile = String(ws.getRange('B5').getValue() || ''); } catch (e) {}
  return s;
}

// ─────────────────────────────────────────────────────────────── правки WIZARD: минимумы, группы, настройки
/** Позиция: мин. запас (G), целевой (H), группа (I), комментарий (J) в «Справочнике»; возвращает пересчитанные значения строки «Остатков». */
function kzwSetItem(token, key, fields) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), ws = kzwSheet_(ss, SH_SPR), last = ws.getLastRow();
  key = String(key || '').trim(); if (!key) throw new Error('Пустой ключ');
  var keys = ws.getRange(2, 1, Math.max(last - 1, 1), 1).getValues(), row = 0;
  for (var i = 0; i < keys.length; i++) if (kzwKey_(keys[i][0]) === key) { row = i + 2; break; }
  if (!row) throw new Error('Позиция «' + key + '» не найдена в «Справочнике». Выполните «Обновить отчёты».');
  var upd = {};
  var put = function (col, val, name) {
    if (val === undefined) return;
    var cell = ws.getRange(row, col);
    if (val === null || val === '') cell.clearContent(); else cell.setValue(val);
    upd[name] = val;
  };
  var numOrNull = function (v) { if (v === undefined) return undefined; if (v === null || v === '') return null; var n = kzwNum_(v); if (n < 0) throw new Error('Отрицательное значение'); return n; };
  put(7, numOrNull(fields.min), 'min'); put(8, numOrNull(fields.tgt), 'tgt');
  if (fields.crit !== undefined) { kzwEnsureCritCol_(ss); put(KZW_CRIT.COL, numOrNull(fields.crit), 'crit'); }
  if (fields.grp !== undefined) { put(9, String(fields.grp || '').trim(), 'grp'); try { kzwXrefSetKeyGroup_(ss, keys[row - 2][0], String(fields.grp || '').trim()); } catch (eX) {} }   // и в «Аналоги», иначе syncXref_ вернёт старую группу
  if (fields.comment !== undefined) put(10, String(fields.comment || ''), 'comment');
  if (fields.excl !== undefined) { kzwEnsureExclCol_(ss); put(KZW_EXCL.COL, String(fields.excl || '').trim(), 'excl'); }   // v1.6: исключение из контроля (снято с производства / нерелевантно)
  SpreadsheetApp.flush();
  // пересчитанные формулы «Остатков» (U:AB) — по ключу
  var wsO = kzwSheet_(ss, SH_OST), lo = wsO.getLastRow(), ok = wsO.getRange(2, 1, Math.max(lo - 1, 1), 1).getValues(), orow = 0;
  for (var j = 0; j < ok.length; j++) if (kzwKey_(ok[j][0]) === key) { orow = j + 2; break; }
  var res = { key: key, upd: upd };
  if (orow) {
    var v = wsO.getRange(orow, C_MIN, 1, 8).getValues()[0];
    res.calc = { min: (v[0] === '' ? null : kzwNum_(v[0])), tgt: (v[1] === '' ? null : kzwNum_(v[1])), def: kzwNum_(v[2]), grp: String(v[3] || ''), stat: String(v[4] || ''), prio: kzwNum_(v[5]) };
  }
  return res;
}
/** Группа замен: назначение (B), мин. запас группы (C), комментарий (D). */
function kzwSetGroup(token, id, fields) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), ws = kzwSheet_(ss, SH_GRP), last = Math.max(ws.getLastRow(), 1);
  id = String(id || '').trim();
  var ids = ws.getRange(2, 1, Math.max(last - 1, 1), 1).getValues(), row = 0;
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) { row = i + 2; break; }
  if (!row) throw new Error('Группа ' + id + ' не найдена');
  if (fields.purpose !== undefined) ws.getRange(row, 2).setValue(String(fields.purpose || ''));
  if (fields.min !== undefined) { if (fields.min === null || fields.min === '') ws.getRange(row, 3).clearContent(); else ws.getRange(row, 3).setValue(kzwNum_(fields.min)); }
  if (fields.comment !== undefined) ws.getRange(row, 4).setValue(String(fields.comment || ''));
  SpreadsheetApp.flush();
  var v = ws.getRange(row, 1, 1, 8).getValues()[0];
  return { id: id, purpose: String(v[1] || ''), min: (v[2] === '' ? null : kzwNum_(v[2])), comment: String(v[3] || ''), stock: kzwNum_(v[4]), count: kzwNum_(v[5]), stat: String(v[6] || ''), prio: kzwNum_(v[7]), row: row };
}
/** Новая группа замен: следующий свободный ID G###; формулы E:H ставит ensureGroupPriority_ основного скрипта. */
function kzwAddGroup(token, fields) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), ws = kzwSheet_(ss, SH_GRP), last = Math.max((typeof lastRowByCol_ === 'function') ? lastRowByCol_(ws, 1) : ws.getLastRow(), 1);
  var ids = last >= 2 ? ws.getRange(2, 1, last - 1, 1).getValues() : [], maxN = 0;
  ids.forEach(function (r) { var m = String(r[0]).match(/^GR?(\d+)$/i); if (m) maxN = Math.max(maxN, +m[1]); });   // в книге встречаются G001 и GR42
  var id = 'G' + ('000' + (maxN + 1)).slice(-3), row = last + 1;
  if (typeof ensureRows_ === 'function') ensureRows_(ws, row + 5);
  ws.getRange(row, 1, 1, 4).setValues([[id, String(fields.purpose || ''), (fields.min === null || fields.min === undefined || fields.min === '') ? '' : kzwNum_(fields.min), String(fields.comment || 'создано в веб-приложении')]]);
  try { var ctx = kzCtx_(); detectSeparator_(ss, ctx); ensureGroupPriority_(ctx, ss); kzFlushSig_(ctx); } catch (e) {}
  SpreadsheetApp.flush();
  var v = ws.getRange(row, 1, 1, 8).getValues()[0];
  return { id: id, purpose: String(v[1] || ''), min: (v[2] === '' ? null : kzwNum_(v[2])), comment: String(v[3] || ''), stock: kzwNum_(v[4]), count: kzwNum_(v[5]), stat: String(v[6] || ''), prio: kzwNum_(v[7]), row: row };
}
/** Настройки контроля: порог жёлтой зоны (B2), минимум по умолчанию (B3). */
function kzwSetSettings(token, s) {
  kzwAuth_(token, 'wizard');
  var ws = kzwSheet_(kzwSs_(), SH_SET);
  if (s.yellow !== undefined) { var y = kzwNum_(s.yellow); if (y > 1) y = y / 100; ws.getRange('B2').setValue(y); }
  if (s.defMin !== undefined) { if (s.defMin === null || s.defMin === '') ws.getRange('B3').clearContent(); else ws.getRange('B3').setValue(kzwNum_(s.defMin)); }
  SpreadsheetApp.flush();
  return kzwReadSettings_(kzwSs_());
}
/** Склад 1С: колонка контроля (B) и признак «Учитывать в контроле» (C). Вступает в силу после пересборки. */
function kzwSetWarehouse(token, row, fields) {
  kzwAuth_(token, 'wizard');
  var ws = kzwSheet_(kzwSs_(), SH_SET), r = +row;
  var h = (typeof settingsHeaderRow_ === 'function') ? settingsHeaderRow_(ws) : 9;
  if (!(r > h)) throw new Error('Неверная строка склада');
  if (fields.group !== undefined) ws.getRange(r, 2).setValue(String(fields.group || 'Прочие'));
  if (fields.inc !== undefined) ws.getRange(r, 3).setValue(fields.inc ? 'Да' : 'Нет');
  return kzwReadSettings_(kzwSs_());
}
/** Применить группы из листа «Аналоги» к «Справочнику» (как пункт меню). */
function kzwSyncXref(token) {
  kzwAuth_(token, 'wizard');
  var ctx = kzCtx_(), ss = ctx.ss;
  detectSeparator_(ss, ctx);
  ensureXrefSheet_(ctx, ss);
  var r = syncXref_(ss, null, null);
  ensureGroupPriority_(ctx, ss);
  kzFlushSig_(ctx);
  var msg = 'Связей применено: ' + r.applied + '. Новых групп: ' + r.newGroups + '.';
  if (r.unresolved.length) msg += ' Не найдено позиций: ' + r.unresolved.length + ' (' + r.unresolved.slice(0, 8).join(', ') + (r.unresolved.length > 8 ? '…' : '') + ').';
  if (r.conflicts.length) msg += ' Ключ в нескольких группах (взята первая): ' + r.conflicts.slice(0, 8).join(', ') + '.';
  return msg + ' Чтобы группы появились в отчётах — запустите «Обновить отчёты».';
}

// ─────────────────────────────────────────────────────────────── загрузка 1С из приложения и пересборка
/**
 * Файл выгрузки «Остатки по складам» (xlsx/xls) из браузера: base64 → временный файл на Диске →
 * конвертация в Google-таблицу (Drive API) → лист «Данные 1С» → отложенный запуск kzRun_ по триггеру
 * (чтобы запрос из браузера не упирался в лимит времени). Ход — kzwStatus().
 */
function kzwUpload1C(token, fileName, base64, mime) {
  kzwAuth_(token, 'wizard');
  if (kzLoadState_()) throw new Error('Есть незавершённая пересборка — сначала «Продолжить» или «Сбросить» на вкладке «Загрузка 1С».');
  if (PropertiesService.getScriptProperties().getProperty(KZW.PROP_PENDING)) throw new Error('Пересборка уже запущена и ещё не завершилась. Подождите.');
  if (!base64) throw new Error('Пустой файл');
  var ss = kzwSs_(), tmpFile = null, tmpId = null, out;
  try {
    var bytes = Utilities.base64Decode(base64);
    var isGsFile = false;
    tmpFile = DriveApp.createFile(Utilities.newBlob(bytes, mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'web_1C_' + (fileName || 'upload.xlsx')));
    tmpId = convertToGoogleSheet_(tmpFile);
    var srcSs = SpreadsheetApp.openById(tmpId);
    var wsS = srcSs.getSheetByName('TDSheet') || srcSs.getSheets()[0];
    var v = wsS.getDataRange().getValues();
    if (v.length < 2 || String(v[0][2] || '').toLowerCase().indexOf('тмц') < 0 || String(v[0][6] || '').toLowerCase().indexOf('кол-во') < 0)
      throw new Error('Файл не похож на выгрузку «Остатки по складам»: в 1-й строке ожидаются заголовки «ТМЦ» (колонка C) и «Кол-во остатков» (колонка G).');
    out = [];
    for (var r = 1; r < v.length; r++) {
      var row = v[r], nm = normName_(row[2]);
      if (!nm || String(row[0] || '').trim().toLowerCase() === 'итог') continue;
      var o = new Array(13);
      for (var c = 0; c < 13; c++) {
        var val = row[c] === undefined ? '' : row[c];
        if (c === 0) o[c] = String(val).trim();
        else if (c >= 1 && c <= 5) o[c] = normName_(val);
        else if (c === 9) o[c] = toDate_(val) || val;
        else o[c] = val;
      }
      out.push(o);
    }
    if (!out.length) throw new Error('В файле нет строк с остатками');
    var ws = kzwSheet_(ss, SH_RAW), maxR = ws.getMaxRows();
    if (maxR > 1) ws.getRange(2, 1, maxR - 1, 13).clearContent();
    ensureRows_(ws, out.length + 1);
    ws.getRange(2, 1, out.length, 13).setValues(out);
    ws.getRange(2, 10, out.length, 1).setNumberFormat(FMT_DATE);
    var wsSet = kzwSheet_(ss, SH_SET);
    wsSet.getRange('B4').setValue(new Date()).setNumberFormat(FMT_DATE);
    wsSet.getRange('B5').setValue(String(fileName || tmpFile.getName()));
    wsSet.getRange('A6').setValue('Состояние пересборки'); wsSet.getRange('B6').setValue('Загружено строк 1С: ' + out.length + '. Пересборка запускается…');
    SpreadsheetApp.flush();
  } catch (e) {
    var m = String(e && e.message || e);
    if (/Drive/i.test(m) && /не определ|is not defined|undefined/i.test(m)) m = 'Для чтения Excel нужен сервис «Drive API»: редактор скриптов → Службы (+) → Drive API → Добавить. ' + m;
    throw new Error('Ошибка загрузки: ' + m);
  } finally {
    try { if (tmpId) DriveApp.getFileById(tmpId).setTrashed(true); } catch (e1) {}
    try { if (tmpFile) tmpFile.setTrashed(true); } catch (e2) {}
  }
  kzwScheduleRebuild_({ fromLoad: true, loadedRows: out.length, loadedFile: String(fileName || '') });
  return { rows: out.length, started: true };
}

/** Пересборка отчётов из приложения: mode = 'new' | 'resume' | 'reset'. Выполняется по триггеру, ход — kzwStatus(). */
function kzwStartRebuild(token, mode) {
  kzwAuth_(token, 'wizard');
  var props = PropertiesService.getScriptProperties();
  if (mode === 'reset') { kzClearState_(); props.deleteProperty(KZW.PROP_PENDING); kzwDeleteOwnTriggers_(); try { var lk = LockService.getDocumentLock(); if (lk.tryLock(0)) lk.releaseLock(); } catch (eL) {} try { kzwSheet_(kzwSs_(), SH_SET).getRange('B6').setValue('Состояние сброшено ' + kzwDateTimeStr_(new Date())); } catch (e) {} return { reset: true }; }
  if (props.getProperty(KZW.PROP_PENDING)) throw new Error('Пересборка уже запущена и ещё не началась / идёт. Подождите.');
  var st = kzLoadState_();
  if (mode === 'resume') { if (!st) throw new Error('Незавершённой пересборки нет'); kzwScheduleRebuild_({ resume: true }); return { started: true, resume: true }; }
  if (st) { var info = kzwStateInfo_(kzwSs_()); if (info.busy) throw new Error('Есть незавершённая пересборка — выберите «Продолжить» или «Сбросить».'); kzClearState_(); props.deleteProperty(KZW.PROP_PENDING); kzwDeleteOwnTriggers_(); }
  kzwScheduleRebuild_({ fromLoad: false });
  return { started: true };
}
function kzwScheduleRebuild_(opts) {
  opts = opts || {}; opts.byTrigger = true; opts.at = Date.now();
  PropertiesService.getScriptProperties().setProperty(KZW.PROP_PENDING, JSON.stringify(opts));
  kzwDeleteOwnTriggers_();
  ScriptApp.newTrigger(KZW.TRIGGER_FN).timeBased().after(1000).create();
}
function kzwDeleteOwnTriggers_() {
  try { ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === KZW.TRIGGER_FN) ScriptApp.deleteTrigger(t); }); } catch (e) {}
}
/** Обработчик триггера: запускает конвейер kzRun_ основного скрипта (сам продолжит себя при нехватке времени). */
function kzwRunRebuildTrigger() {
  var props = PropertiesService.getScriptProperties(), raw = props.getProperty(KZW.PROP_PENDING);
  kzwDeleteOwnTriggers_();
  if (!raw) return;
  var opts = {}; try { opts = JSON.parse(raw) || {}; } catch (e) {}
  opts.byTrigger = true;
  var res = null;
  try { res = kzRun_(opts); }
  catch (err) { res = { error: String(err && err.message || err) }; }
  finally { props.deleteProperty(KZW.PROP_PENDING); }
  if (res && res.error) console.error('kzwRunRebuildTrigger: ' + res.error);
  if (res && res.stat) { try { kzwApplyCritFormulas_(kzwSs_(), true); } catch (e2) { console.error('crit formulas: ' + e2.message); } }
}

// ─────────────────────────────────────────────────────────────── любой лист книги (просмотр / правка WIZARD)
function kzwGetSheet(token, name) {
  var role = kzwAuth_(token);
  var ss = kzwSs_(), sh = kzwSheet_(ss, name);
  var lr = Math.min(Math.max(sh.getLastRow(), 1), KZW.MAX_ROWS), lc = Math.min(Math.max(sh.getLastColumn(), 1), KZW.MAX_COLS);
  var rng = sh.getRange(1, 1, lr, lc);
  var widths = []; for (var c = 0; c < lc; c++) widths.push(sh.getColumnWidth(c + 1));
  var hiddenCols = []; for (var hc = 1; hc <= lc; hc++) { try { if (sh.isColumnHiddenByUser(hc)) hiddenCols.push(hc); } catch (e0) {} }
  var merges = [];
  try { rng.getMergedRanges().forEach(function (m) { merges.push([m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()]); }); } catch (e) {}
  var bgs = rng.getBackgrounds(), fcs = rng.getFontColors(), fws = rng.getFontWeights(), fms = rng.getFormulas();
  var styles = {}, formulas = [];
  for (var r = 0; r < lr; r++) for (var cc = 0; cc < lc; cc++) {
    var bg = bgs[r][cc], fc = fcs[r][cc], fw = fws[r][cc];
    var st = [(bg && bg !== '#ffffff') ? bg : '', (fc && fc !== '#000000') ? fc : '', fw === 'bold' ? 1 : 0];
    if (st[0] || st[1] || st[2]) styles[(r + 1) + ':' + (cc + 1)] = st;
    if (fms[r][cc]) formulas.push((r + 1) + ':' + (cc + 1));
  }
  return { name: name, rows: lr, cols: lc, widths: widths, hiddenCols: hiddenCols, frozenRows: sh.getFrozenRows(), frozenCols: sh.getFrozenColumns(),
    charts: sh.getCharts().length, values: rng.getDisplayValues(), styles: styles, formulas: formulas, merges: merges,
    editable: role === 'wizard', truncated: sh.getLastRow() > KZW.MAX_ROWS || sh.getLastColumn() > KZW.MAX_COLS };
}
function kzwGetCharts(token, name) {
  kzwAuth_(token);
  var sh = kzwSs_().getSheetByName(name); if (!sh) return [];
  var out = [], cache = CacheService.getScriptCache();
  sh.getCharts().forEach(function (ch, i) {
    try {
      var opts = ch.getOptions(), key = 'kzw_chart_' + Utilities.base64EncodeWebSafe(name).slice(0, 40) + '_' + i, png = null;
      try { png = cache.get(key); } catch (e) {}
      if (!png) { png = Utilities.base64Encode(ch.getAs('image/png').getBytes()); try { if (png.length < 95000) cache.put(key, png, 600); } catch (e2) {} }
      var title = ''; try { title = String(opts.get('title') || ''); } catch (e3) {}
      out.push({ title: title, png: png });
    } catch (err) { out.push({ error: String(err && err.message || err) }); }
  });
  return out;
}
function kzwSetCell(token, name, row, col, value, force) {
  kzwAuth_(token, 'wizard');
  var sh = kzwSheet_(kzwSs_(), name), cell = sh.getRange(+row, +col);
  if (cell.getFormula() && !force) throw new Error('Ячейка содержит формулу');
  var v = value;
  if (typeof v === 'string') {
    var s = v.trim();
    if (s.charAt(0) === '=') { cell.setFormula(s); return { display: cell.getDisplayValue(), formula: true }; }
    if (/^-?\d+([.,]\d+)?$/.test(s)) v = parseFloat(s.replace(',', '.'));
    else if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(s)) { var m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); v = new Date(+m[3], +m[2] - 1, +m[1]); }
    else v = s;
  }
  cell.setValue(v);
  return { display: cell.getDisplayValue() };
}
function kzwGetCellFormula(token, name, row, col) {
  kzwAuth_(token, 'wizard');
  var c = kzwSheet_(kzwSs_(), name).getRange(+row, +col);
  return { formula: c.getFormula(), value: c.getDisplayValue() };
}
function kzwSetCellStyle(token, name, row, col, st) {
  kzwAuth_(token, 'wizard');
  var rng = kzwSheet_(kzwSs_(), name).getRange(+row, +col, +(st.rows || 1), +(st.cols || 1));
  if (st.bold !== undefined && st.bold !== null) rng.setFontWeight(st.bold ? 'bold' : 'normal');
  if (st.fc) rng.setFontColor(st.fc === 'auto' ? null : st.fc);
  if (st.bg) rng.setBackground(st.bg === 'none' ? null : st.bg);
  if (st.size) rng.setFontSize(+st.size);
  if (st.italic !== undefined && st.italic !== null) rng.setFontStyle(st.italic ? 'italic' : 'normal');
  if (st.align) rng.setHorizontalAlignment(st.align);
  return true;
}
/** Переместить столбец (WIZARD). Листы, которые строит скрипт («Остатки», «Справочник», вкладки брендов…), трогать нельзя — он опирается на номера колонок. */
function kzwMoveColumn(token, name, col, dir) {
  kzwAuth_(token, 'wizard');
  if (kzwIsScriptSheet_(name)) throw new Error('Лист «' + name + '» формирует скрипт — порядок его столбцов менять нельзя');
  var sh = kzwSheet_(kzwSs_(), name);
  col = +col; dir = +dir;
  var dest = dir < 0 ? col - 1 : col + 2;
  if (dest < 1 || col + dir < 1 || col + dir > sh.getMaxColumns()) throw new Error('Дальше двигать некуда');
  var fc = 0; try { fc = sh.getFrozenColumns(); } catch (e) {}
  if (fc) { try { sh.setFrozenColumns(0); } catch (e1) {} }
  var merges = [];
  try { sh.getDataRange().getMergedRanges().forEach(function (m) { if (m.getNumColumns() > 1) { merges.push({ row: m.getRow(), rows: m.getNumRows(), col: m.getColumn(), cols: m.getNumColumns() }); m.breakApart(); } }); } catch (e2) {}
  try { sh.moveColumns(sh.getRange(1, col, sh.getMaxRows(), 1), dest); }
  finally { merges.forEach(function (m) { try { sh.getRange(m.row, m.col, m.rows, m.cols).merge(); } catch (e3) {} }); if (fc) { try { sh.setFrozenColumns(fc); } catch (e4) {} } }
  return { col: col + dir };
}
function kzwIsScriptSheet_(name) {
  var fixed = [SH_RAW, SH_NOM, SH_SPR, SH_OST, SH_SET, SH_GRP, SH_SOST, SH_HOM, SH_CRIT, SH_SUM, SH_RA, SH_RB, SH_RT, SH_XREF, SH_ST, SH_HI, SH_CROSS, SH_SUGG, KZW.SHEET_PAGES];
  if (fixed.indexOf(name) !== -1) return true;
  var sh = kzwSs_().getSheetByName(name);
  try { return !!(sh && typeof looksLikeBrandTab_ === 'function' && looksLikeBrandTab_(sh)); } catch (e) { return false; }
}
function kzwAddSheet(token, name) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(); name = String(name || '').trim();
  if (!name) throw new Error('Пустое имя');
  if (ss.getSheetByName(name)) throw new Error('Лист с таким именем уже есть');
  var sh = ss.insertSheet(name); sh.getRange('A1').setValue(name).setFontWeight('bold').setFontSize(13);
  return { sheets: kzwVisibleSheets_(ss, 'wizard'), all: ss.getSheets().map(function (x) { return x.getName(); }) };
}
function kzwRenameSheet(token, oldName, newName) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), sh = kzwSheet_(ss, oldName);
  if (kzwIsScriptSheet_(oldName)) throw new Error('Лист «' + oldName + '» используется скриптом — переименование запрещено (своё название вкладки можно задать в конструкторе)');
  newName = String(newName || '').trim(); if (!newName) throw new Error('Пустое имя');
  if (ss.getSheetByName(newName)) throw new Error('Лист с таким именем уже есть');
  sh.setName(newName);
  return kzwVisibleSheets_(ss, 'wizard');
}
function kzwGetRange(token, sheetName, rangeA1) {
  kzwAuth_(token);
  var sh = kzwSheet_(kzwSs_(), sheetName), rng;
  try { rng = rangeA1 ? sh.getRange(rangeA1) : sh.getDataRange(); } catch (e) { throw new Error('Неверный диапазон: ' + rangeA1); }
  if (rng.getNumRows() > 500) rng = sh.getRange(rng.getRow(), rng.getColumn(), 500, rng.getNumColumns());
  var vals = rng.getDisplayValues(), bgs = rng.getBackgrounds(), fcs = rng.getFontColors(), fws = rng.getFontWeights(), styles = {};
  for (var r = 0; r < vals.length; r++) for (var c = 0; c < vals[r].length; c++) {
    var st = [(bgs[r][c] && bgs[r][c] !== '#ffffff') ? bgs[r][c] : '', (fcs[r][c] && fcs[r][c] !== '#000000') ? fcs[r][c] : '', fws[r][c] === 'bold' ? 1 : 0];
    if (st[0] || st[1] || st[2]) styles[(r + 1) + ':' + (c + 1)] = st;
  }
  return { values: vals, styles: styles };
}

// ─────────────────────────────────────────────────────────────── конструктор страниц и настройки интерфейса (лист «Веб-страницы»)
function kzwPagesSheet_(ss, create) {
  var sh = ss.getSheetByName(KZW.SHEET_PAGES);
  if (!sh && create) {
    sh = ss.insertSheet(KZW.SHEET_PAGES);
    sh.getRange(1, 1, 1, KZW.PAGES_HEADER.length).setValues([KZW.PAGES_HEADER]).setBackground('#1f3864').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.getRange('A2').setValue('_settings'); sh.getRange('G2').setValue('{}');
    sh.setColumnWidth(2, 220); sh.setColumnWidth(7, 600); sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}
function kzwReadPages_(ss) {
  var sh = kzwPagesSheet_(ss, false);
  if (!sh || sh.getLastRow() < 2) return { pages: [], settings: {} };
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues(), pages = [], settings = {};
  vals.forEach(function (r, i) {
    var id = String(r[0]).trim(); if (!id) return;
    var json = {}; try { json = r[6] ? JSON.parse(r[6]) : {}; } catch (e) { json = { error: 'JSON: ' + e.message }; }
    if (id === '_settings') { settings = json || {}; return; }
    pages.push({ id: id, title: String(r[1] || ''), icon: String(r[2] || '📄'), order: +r[3] || 0,
      roles: String(r[4] || 'viewer,wizard').split(',').map(function (x) { return x.trim(); }).filter(Boolean),
      tab: String(r[5] || '').trim().toLowerCase() === 'да', blocks: Array.isArray(json) ? json : (json.blocks || []), rowIndex: i + 2 });
  });
  pages.sort(function (a, b) { return a.order - b.order || a.title.localeCompare(b.title, 'ru'); });
  return { pages: pages, settings: settings };
}
function kzwPages_(ss, role) { return kzwReadPages_(ss).pages.filter(function (p) { return role === 'wizard' || p.roles.indexOf(role) !== -1; }).map(function (p) { delete p.rowIndex; return p; }); }
function kzwSettings_(ss) { return kzwReadPages_(ss).settings || {}; }
function kzwVisibleSheets_(ss, role) {
  var st = kzwSettings_(ss), hidden = (st.hiddenSheets || (st.hiddenSheetsSet ? [] : KZW.HIDDEN_SHEETS)).concat([KZW.SHEET_PAGES]);
  return ss.getSheets().filter(function (sh) { return hidden.indexOf(sh.getName()) === -1 && !sh.isSheetHidden(); }).map(function (sh) { return sh.getName(); });
}
function kzwWriteSettingsRow_(ss, settings) {
  var sh = kzwPagesSheet_(ss, true), vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues(), r = 0;
  for (var i = 1; i < vals.length; i++) if (String(vals[i][0]) === '_settings') { r = i + 1; break; }
  if (!r) { r = sh.getLastRow() + 1; sh.getRange(r, 1).setValue('_settings'); }
  sh.getRange(r, 7).setValue(JSON.stringify(settings || {}));
}
function kzwSavePage(token, page) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), sh = kzwPagesSheet_(ss, true), all = kzwReadPages_(ss), id = String(page.id || '').trim() || ('p' + Date.now().toString(36));
  var row = [id, String(page.title || 'Без названия'), String(page.icon || '📄'), +page.order || 0, (page.roles || ['viewer', 'wizard']).join(','), page.tab ? 'Да' : 'Нет', JSON.stringify(page.blocks || [])];
  var existing = all.pages.filter(function (p) { return p.id === id; })[0];
  sh.getRange(existing ? existing.rowIndex : sh.getLastRow() + 1, 1, 1, 7).setValues([row]);
  return { id: id, pages: kzwPages_(ss, 'wizard') };
}
function kzwDeletePage(token, id) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), sh = kzwPagesSheet_(ss, false);
  if (!sh) return { pages: [] };
  var p = kzwReadPages_(ss).pages.filter(function (x) { return x.id === id; })[0];
  if (p) sh.deleteRow(p.rowIndex);
  return { pages: kzwPages_(ss, 'wizard') };
}
function kzwSaveSettings(token, settings) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(); settings = settings || {}; settings.hiddenSheetsSet = true;
  kzwWriteSettingsRow_(ss, settings);
  return { settings: settings, sheets: kzwVisibleSheets_(ss, 'wizard') };
}
function kzwSavePinned(token, list) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), st = kzwSettings_(ss);
  st.ui = st.ui || {}; st.ui.pinnedSheets = (list || []).map(String).slice(0, 12);
  kzwWriteSettingsRow_(ss, st);
  return st;
}

// ═══════════════════════════════════════════════════════════════ v1.1: СВЕРКА ДУБЛЕЙ, АНАЛОГИ (Cross), БАЗА АНАЛОГОВ
var KZW_REC = { SHEET: 'Сверка', H: 3, HDR: ['ID', 'Признак совпадения', 'Ключи позиций (через ;)', 'Решение', 'Группа замен', 'Дата решения', 'Примечание', 'Позиций', 'Наименования'], MAX_CLUSTER: 12 };

/** Строка «Аналоги» ↔ ключ: поставить/снять группу (чтобы syncXref_ не вернул старую связь). */
function kzwXrefSetKeyGroup_(ss, key, grp) {
  var ws = ss.getSheetByName(SH_XREF); if (!ws) return;
  var last = xrefLastUsedRow_(ws); if (last <= XREF_H) { if (grp) kzwXrefAppend_(ws, last, [[grp, key]]); return; }
  var v = ws.getRange(XREF_H + 1, 1, last - XREF_H, 3).getValues(), ks = kzwKey_(key), found = false;
  for (var i = 0; i < v.length; i++) {
    if (kzwKey_(v[i][1]) === ks || kzwKey_(v[i][2]) === ks) { found = true; if (String(v[i][0]).trim() !== String(grp)) ws.getRange(XREF_H + 1 + i, 1).setValue(grp || ''); }
  }
  if (!found && grp) kzwXrefAppend_(ws, last, [[grp, key]]);
}
function kzwXrefAppend_(ws, last, rows) {
  var r1 = Math.max(last, XREF_H) + 1;
  ensureRows_(ws, r1 + rows.length + 5);
  var lastF = ws.getRange(XREF_H + 1, 3, Math.max(ws.getMaxRows() - XREF_H, 1), 1).getFormulas().reduce(function (mx, row, i) { return row[0] ? XREF_H + 1 + i : mx; }, XREF_H);
  if (r1 + rows.length > lastF) writeXrefFormulas_(ws, lastF + 1, r1 + rows.length + 100);
  ws.getRange(r1, 1, rows.length, 2).setValues(rows);
  ws.getRange(r1, 12, rows.length, 1).setValue('веб-приложение ' + fmtDate_(new Date()));
}
/** Следующий свободный ID группы (учитывает G001 и GR42). */
function kzwNextGroupId_(ss) {
  var ws = ss.getSheetByName(SH_GRP), last = Math.max(lastRowByCol_(ws, 1), 1), maxN = 0;
  if (last >= 2) ws.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { var m = String(r[0]).match(/^GR?(\d+)$/i); if (m) maxN = Math.max(maxN, +m[1]); });
  return 'G' + ('000' + (maxN + 1)).slice(-3);
}
/** Позиции «Остатков» для сверки/аналогов (как в основном скрипте). */
function kzwStock_(ss) {
  return readOst_({ ss: ss }).map(function (r) { return { key: r.key, ks: r.ks, art: r.art, name: r.name, brand: r.brand, qty: r.tot, grp: r.grp, type: r.type }; });
}

// ── лист «Сверка»
function kzwRecSheet_(ss, create) {
  var sh = ss.getSheetByName(KZW_REC.SHEET);
  if (!sh && create) {
    var after = ss.getSheetByName(SH_XREF) || ss.getSheetByName(SH_GRP);
    sh = ss.insertSheet(KZW_REC.SHEET, after ? after.getIndex() : ss.getNumSheets());
    sh.getRange('A1').setValue('Сверка одинаковых артикулов и наименований — кандидаты в дубли / один товар под разными ключами').setFontWeight('bold').setFontSize(13).setFontColor(COLOR_HEAD);
    sh.getRange('A2').setValue('Заполняется веб-приложением (вкладка «Сверка»). Решение «идентичный» объединяет позиции в группу замен (общий остаток), «разные» — оставляет раздельно и больше не предлагает. Можно править и здесь: столбец «Решение» (идентичный / разные / пусто).').setFontColor(COLOR_SUB).setWrap(true);
    sh.setRowHeight(2, 40);
    hdr_(sh, KZW_REC.H, 1, KZW_REC.HDR); sh.setFrozenRows(KZW_REC.H);
    [80, 200, 320, 100, 90, 110, 260, 60, 520].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange('C:C').setNumberFormat('@');
  }
  return sh;
}
function kzwRecRead_(ss) {
  var sh = kzwRecSheet_(ss, false), out = {};
  if (!sh || sh.getLastRow() <= KZW_REC.H) return out;
  sh.getRange(KZW_REC.H + 1, 1, sh.getLastRow() - KZW_REC.H, 7).getValues().forEach(function (r, i) {
    var keys = String(r[2] || '').split(';').map(function (x) { return x.trim(); }).filter(Boolean); if (!keys.length) return;
    out[keys.slice().sort().join(';')] = { id: String(r[0]), reason: String(r[1]), keys: keys, decision: kzwRecDec_(r[3]), grp: String(r[4] || '').trim(), date: kzwDateStr_(r[5]) || String(r[5] || ''), note: String(r[6] || ''), row: KZW_REC.H + 1 + i };
  });
  return out;
}
function kzwRecDec_(v) { v = String(v || '').trim().toLowerCase(); if (!v) return ''; if (v.indexOf('идент') === 0 || v === 'same' || v === 'да') return 'same'; if (v.indexOf('разн') === 0 || v === 'diff' || v === 'нет') return 'diff'; return ''; }
function kzwRecHash_(s) { var h = 5381; for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return 'R' + (h >>> 0).toString(36).toUpperCase(); }
function kzwNormName_(s) { return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]+/g, ''); }
function kzwNameTokens_(s) { var d = {}; String(s || '').toUpperCase().replace(/[^A-ZА-ЯЁ0-9]+/g, ' ').split(' ').forEach(function (p) { if (p.length > 1) d[p] = 1; }); return d; }

/**
 * Поиск кандидатов: одинаковый артикул у разных ключей, одинаковое (нормализованное) наименование,
 * артикул одной позиции внутри наименования другой, очень похожие наименования одного бренда.
 * Возвращает кластеры (2..MAX_CLUSTER позиций) с причинами. Решения берутся из листа «Сверка».
 */
function kzwReconcileCompute_(ss) {
  var stock = kzwStock_(ss), items = [];
  stock.forEach(function (it) {
    var art = anlNorm_(it.art); if (art.length < 3) art = '';
    var nn = kzwNormName_(it.name); if (nn.length < 6) nn = '';
    items.push({ it: it, art: art, nn: nn, tok: kzwNameTokens_(it.name) });
  });
  var par = items.map(function (_, i) { return i; }), why = {};
  function f(x) { while (par[x] !== x) { par[x] = par[par[x]]; x = par[x]; } return x; }
  function u(a, b, t) { var ra = f(a), rb = f(b); if (ra !== rb) par[rb] = ra; var k = a < b ? a + '|' + b : b + '|' + a; (why[k] = why[k] || []).push(t); }
  var byArt = {}, byName = {}, byTok = {};
  items.forEach(function (o, i) {
    if (o.art) (byArt[o.art] = byArt[o.art] || []).push(i);
    if (o.nn) (byName[o.nn] = byName[o.nn] || []).push(i);
    Object.keys(o.tok).forEach(function (t) { var k = anlNorm_(t); if (k.length >= 5 && /\d/.test(k)) (byTok[k] = byTok[k] || []).push(i); });
  });
  Object.keys(byArt).forEach(function (a) { var l = byArt[a]; if (l.length > 1 && l.length <= KZW_REC.MAX_CLUSTER) for (var i = 1; i < l.length; i++) u(l[0], l[i], 'одинаковый артикул ' + items[l[0]].it.art); });
  Object.keys(byName).forEach(function (n) { var l = byName[n]; if (l.length > 1 && l.length <= KZW_REC.MAX_CLUSTER) for (var i = 1; i < l.length; i++) u(l[0], l[i], 'одинаковое наименование'); });
  items.forEach(function (o, i) {
    if (!o.art || o.art.length < 5 || !/\d/.test(o.art)) return;
    var l = byTok[o.art]; if (!l || l.length > 6) return;
    l.forEach(function (j) { if (j !== i && items[j].art !== o.art) u(i, j, 'артикул ' + o.it.art + ' в наименовании'); });
  });
  // похожие наименования одного бренда (Жаккар ≥ 0.75) — только внутри бренда, чтобы не перебирать всё
  var byBrand = {}; items.forEach(function (o, i) { (byBrand[o.it.brand] = byBrand[o.it.brand] || []).push(i); });
  Object.keys(byBrand).forEach(function (b) {
    var l = byBrand[b]; if (l.length > 400) return;
    for (var x = 0; x < l.length; x++) for (var y = x + 1; y < l.length; y++) {
      var A = items[l[x]], B = items[l[y]]; if (f(l[x]) === f(l[y])) continue;
      var na = Object.keys(A.tok).length, nb = Object.keys(B.tok).length; if (na < 3 || nb < 3) continue;
      var s = jaccard_(A.tok, B.tok); if (s >= 0.75) u(l[x], l[y], 'похожее наименование (' + Math.round(s * 100) + '%)');
    }
  });
  var comps = {}; items.forEach(function (o, i) { var r = f(i); (comps[r] = comps[r] || []).push(i); });
  var dec = kzwRecRead_(ss), out = [], skipped = 0;
  Object.keys(comps).forEach(function (r) {
    var mem = comps[r]; if (mem.length < 2) return;
    if (mem.length > KZW_REC.MAX_CLUSTER) { skipped++; return; }
    var keys = mem.map(function (i) { return items[i].it.ks; }).sort(), kk = keys.join(';');
    var reasons = [], memSet = {}; mem.forEach(function (i) { memSet[i] = 1; });
    Object.keys(why).forEach(function (k) { var p = k.split('|'); if (memSet[p[0]] && memSet[p[1]]) why[k].forEach(function (t) { if (reasons.indexOf(t) < 0) reasons.push(t); }); });
    var grps = {}; mem.forEach(function (i) { if (items[i].it.grp) grps[items[i].it.grp] = 1; });
    var gl = Object.keys(grps), allSame = gl.length === 1 && mem.every(function (i) { return items[i].it.grp === gl[0]; });
    var d = dec[kk] || null, decision = d ? d.decision : '', grp = d ? d.grp : '';
    if (!decision && allSame) { decision = 'same'; grp = gl[0]; }
    out.push({ id: d ? d.id : kzwRecHash_(kk), reason: reasons.slice(0, 4).join('; '), keys: keys, decision: decision, grp: grp || (gl.length === 1 ? gl[0] : ''), groups: gl, auto: !d && allSame, date: d ? d.date : '', note: d ? d.note : '',
      names: mem.map(function (i) { return items[i].it.name; }) });
  });
  out.sort(function (a, b) { return (a.decision ? 1 : 0) - (b.decision ? 1 : 0) || a.reason.localeCompare(b.reason, 'ru'); });
  return { clusters: out, skipped: skipped, total: stock.length };
}
/** Записать текущие кластеры в лист «Сверка» (решения сохраняются). */
function kzwRecWrite_(ss, clusters) {
  var sh = kzwRecSheet_(ss, true), old = kzwRecRead_(ss), rows = [];
  var seen = {};
  clusters.forEach(function (c) {
    var kk = c.keys.slice().sort().join(';'), o = old[kk]; seen[kk] = 1;
    rows.push([c.id, c.reason, c.keys.join('; '), c.decision === 'same' ? 'идентичный' : (c.decision === 'diff' ? 'разные' : ''), c.grp || '', (o && o.date) || (c.decision ? fmtDate_(new Date()) : ''), (o && o.note) || (c.auto ? 'уже в одной группе' : ''), c.keys.length, c.names.join(' | ')]);
  });
  Object.keys(old).forEach(function (kk) { if (!seen[kk] && old[kk].decision) { var o = old[kk]; rows.push([o.id, o.reason + ' (кластер больше не находится)', o.keys.join('; '), o.decision === 'same' ? 'идентичный' : 'разные', o.grp, o.date, o.note, o.keys.length, '']); } });
  var maxR = sh.getMaxRows();
  if (maxR > KZW_REC.H) sh.getRange(KZW_REC.H + 1, 1, maxR - KZW_REC.H, 9).clearContent();
  if (rows.length) { ensureRows_(sh, KZW_REC.H + rows.length); sh.getRange(KZW_REC.H + 1, 1, rows.length, 9).setValues(rows).setFontFamily('Arial').setFontSize(10); }
  return rows.length;
}
function kzwReconcile(token, recompute) {
  kzwAuth_(token);
  var ss = kzwSs_(), sh = kzwRecSheet_(ss, false);
  if (recompute || !sh || sh.getLastRow() <= KZW_REC.H) {
    var r = kzwReconcileCompute_(ss);
    kzwRecWrite_(ss, r.clusters);
    return { clusters: r.clusters, skipped: r.skipped, total: r.total, computed: kzwDateTimeStr_(new Date()) };
  }
  var dec = kzwRecRead_(ss), out = [];
  Object.keys(dec).forEach(function (k) { var d = dec[k]; out.push({ id: d.id, reason: d.reason, keys: d.keys, decision: d.decision, grp: d.grp, groups: d.grp ? [d.grp] : [], date: d.date, note: d.note, names: [] }); });
  out.sort(function (a, b) { return (a.decision ? 1 : 0) - (b.decision ? 1 : 0) || a.reason.localeCompare(b.reason, 'ru'); });
  return { clusters: out, skipped: 0, total: 0, computed: '' };
}
/**
 * Решение по кластеру. decision: 'same' — объединить keys (подмножество кластера) в группу замен (общий остаток);
 * 'diff' — разные товары (разъединить, если ранее были объединены этой сверкой); '' — сбросить решение.
 */
function kzwReconcileDecide(token, id, decision, keys, note) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), sh = kzwRecSheet_(ss, true), all = kzwRecRead_(ss), rec = null, kk = null;
  Object.keys(all).forEach(function (k) { if (all[k].id === id) { rec = all[k]; kk = k; } });
  if (!rec) throw new Error('Строка сверки ' + id + ' не найдена — обновите сверку');
  keys = (keys && keys.length) ? keys.map(String) : rec.keys;
  var stock = kzwStock_(ss), byKs = {}; stock.forEach(function (s) { byKs[s.ks] = s; });
  var members = keys.map(function (k) { return byKs[k]; }).filter(Boolean);
  if (members.length < 2 && decision === 'same') throw new Error('Для объединения нужно минимум две позиции');
  var res = { id: id, decision: decision, grp: '', msg: '' };
  detectSeparator_(ss, kzCtx_());
  if (decision === 'same') {
    var grp = '';
    members.forEach(function (m) { if (!grp && m.grp) grp = m.grp; });
    if (!grp) grp = kzwNextGroupId_(ss);
    var purpose = members.map(function (m) { return m.name; }).sort(function (a, b) { return a.length - b.length; })[0];
    ensureXrefSheet_(kzCtx_(), ss);
    var ent = members.map(function (m) { return { grp: grp, name: purpose, key: m.key, comment: 'сверка: один товар · ' + fmtDate_(new Date()) }; });
    var r = anlApply_(ss, ent, 'сверка (веб)');
    res.grp = grp; res.msg = 'Объединено в группу ' + grp + (r.created.length ? ' (создана)' : '') + ': ' + members.length + ' поз.';
  } else if (decision === 'diff') {
    if (rec.grp && rec.decision === 'same') {        // ранее объединены этой сверкой — разъединить
      var wsP = ss.getSheetByName(SH_SPR), lastP = wsP.getLastRow(), col = wsP.getRange(2, 1, Math.max(lastP - 1, 1), 9).getValues(), n = 0;
      members.forEach(function (m) { for (var i = 0; i < col.length; i++) if (kzwKey_(col[i][0]) === m.ks && String(col[i][8]).trim() === rec.grp) { wsP.getRange(i + 2, 9).clearContent(); n++; } kzwXrefSetKeyGroup_(ss, m.key, ''); });
      res.msg = 'Разъединено: снята группа ' + rec.grp + ' у ' + n + ' поз.';
    } else res.msg = 'Отмечено как разные товары — больше не предлагается.';
  } else res.msg = 'Решение сброшено';
  sh.getRange(rec.row, 3, 1, 5).setValues([[keys.join('; '), decision === 'same' ? 'идентичный' : (decision === 'diff' ? 'разные' : ''), decision === 'same' ? res.grp : (decision === 'diff' ? '' : rec.grp), decision ? new Date() : '', note || rec.note || '']]);
  SpreadsheetApp.flush();
  return res;
}

// ── аналоги по базе Cross (движок основного скрипта anlCompute_ / anlApply_)
function kzwRejected_(ss) { var st = kzwSettings_(ss); return st.rejectedAnalogs || {}; }
function kzwCrossStats_(ss) {
  var rows = anlReadCross_(ss), base = 0, manual = 0, src = {};
  rows.forEach(function (r) { if (anlIsBaseSrc_(r.src)) base++; else manual++; var s = r.src.split('·')[0].trim() || 'вручную'; src[s] = (src[s] || 0) + 1; });
  return { rows: rows.length, base: base, manual: manual, sources: src };
}
/** Предложения аналогов: write=true — также переписать листы «Предложения аналогов» и колонку H «Cross» (как пункт меню). */
function kzwAnalogs(token, write) {
  var role = kzwAuth_(token);
  var ctx = kzCtx_(), ss = ctx.ss; detectSeparator_(ss, ctx);
  anlEnsureCross_(ss);
  var res;
  if (write && role === 'wizard') { res = anlSuggest_(ctx, ss).res; kzFlushSig_(ctx); }
  else res = anlCompute_(kzwStock_(ss), anlReadCross_(ss), anlReadGroups_(ss));
  var rej = kzwRejected_(ss), blocks = [];
  res.blocks.forEach(function (b) {
    var rows = b.rows.map(function (x) { return { cand: x.cand && !rej[x.it.ks + '|' + x.target], rejected: !!(x.cand && rej[x.it.ks + '|' + x.target]), kind: x.kind, label: x.label, target: x.target, tname: x.tname, key: x.it.key, ks: x.it.ks, art: x.it.art, name: x.it.name, brand: x.it.brand, qty: x.it.qty, grp: x.it.grp, total: x.total, why: x.why }; });
    if (rows.some(function (r) { return r.cand; })) blocks.push({ kind: b.kind, rows: rows });
  });
  var ref = res.ref.slice(0, 1500).map(function (x) { return { key: x.it.key, ks: x.it.ks, art: x.it.art, name: x.it.name, brand: x.it.brand, qty: x.it.qty, grp: x.it.grp, n: x.n, an: x.an, others: x.others, ids: x.ids }; });
  return { blocks: blocks, ref: ref, cnt: res.cnt, items: res.items, wide: res.wide, cross: kzwCrossStats_(ss), computed: kzwDateTimeStr_(new Date()) };
}
/** Принять предложения: ent = [{grp, name, key}] → «Группы замен» + «Аналоги» + Справочник. */
function kzwApplyAnalogs(token, ent) {
  kzwAuth_(token, 'wizard');
  if (!ent || !ent.length) throw new Error('Нет выбранных предложений');
  var ctx = kzCtx_(), ss = ctx.ss; detectSeparator_(ss, ctx);
  ensureXrefSheet_(ctx, ss);
  var r = anlApply_(ss, ent.map(function (e) { return { grp: String(e.grp).trim(), name: String(e.name || e.grp), key: e.key, comment: 'аналог по Cross · ' + fmtDate_(new Date()) }; }), 'авто: Cross (веб)');
  try { anlSuggest_(ctx, ss); kzFlushSig_(ctx); } catch (e) {}
  return 'Принято: ' + ent.length + ' поз. Новых групп: ' + r.created.length + (r.created.length ? ' (' + r.created.join(', ') + ')' : '') + '. Строк в «Аналоги»: +' + r.added + ', изменено ' + r.moved + '. Связей в Справочник: ' + r.sync.applied + '.';
}
/** Отклонить предложение (не аналог) — больше не показывать для этой позиции и группы. */
function kzwRejectAnalog(token, ks, target, undo) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), st = kzwSettings_(ss); st.rejectedAnalogs = st.rejectedAnalogs || {};
  if (undo) delete st.rejectedAnalogs[ks + '|' + target]; else st.rejectedAnalogs[ks + '|' + target] = fmtDate_(new Date());
  kzwWriteSettingsRow_(ss, st);
  return Object.keys(st.rejectedAnalogs).length;
}
/**
 * Загрузка базы аналогов фильтров в «Cross замен фильтров». Форматы:
 *  1) CSV в формате листа (первая колонка заголовка «ID строки»);
 *  2) любой CSV/XLSX с колонками «Артикул» и «Аналог…» (+ по желанию «Тип», «Бренд», «Комментарий»);
 *  3) CSV «Артикул;Аналоги» без заголовка.
 * Режимы: append — добавить новые строки к базе (дубли пропускаются); replace — заменить прежнюю базу, ручные строки сохранить (anlWriteCross_); all — заменить всё.
 */
function kzwUploadCross(token, fileName, base64, mime, mode) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), name = String(fileName || 'база'), data = null, tmp = null, tmpId = null;
  try {
    if (/\.(xlsx|xls|xlsm)$/i.test(name)) {
      tmp = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(base64), mime || 'application/octet-stream', 'web_cross_' + name));
      tmpId = convertToGoogleSheet_(tmp);
      data = SpreadsheetApp.openById(tmpId).getSheets()[0].getDataRange().getValues().map(function (r) { return r.map(function (v) { return v instanceof Date ? fmtDate_(v) : String(v == null ? '' : v); }); });
    } else {
      var text = Utilities.newBlob(Utilities.base64Decode(base64)).getDataAsString('UTF-8').replace(/^﻿/, '');
      if (text.indexOf('�') !== -1) { try { text = Utilities.newBlob(Utilities.base64Decode(base64)).getDataAsString('windows-1251'); } catch (e) {} }
      var head = text.split(/\r?\n/)[0], sep = (head.split(';').length >= head.split(',').length && head.split(';').length >= head.split('\t').length) ? ';' : (head.split('\t').length > head.split(',').length ? '\t' : ',');
      data = Utilities.parseCsv(text, sep);
    }
  } finally {
    try { if (tmpId) DriveApp.getFileById(tmpId).setTrashed(true); } catch (e1) {}
    try { if (tmp) tmp.setTrashed(true); } catch (e2) {}
  }
  if (!data || !data.length) throw new Error('Файл пуст');
  var hdr = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); }), src = 'база v2 (' + fmtDate_(new Date()) + ' · ' + name + ')', base = [];
  var find = function (re) { for (var i = 0; i < hdr.length; i++) if (re.test(hdr[i])) return i; return -1; };
  if (hdr[0] === 'id строки') {
    base = data.slice(1).filter(function (r) { return r.length >= 5 && (String(r[2]).trim() || String(r[4]).trim()); }).map(function (r) { return [r[0], r[1], r[2], r[3], r[4], anlIsBaseSrc_(r[5]) ? r[5] : src, r[6] || '', '']; });
  } else {
    var iArt = find(/^(артикул|art|код|code|part)/), iAn = -1, anCols = [], iType = find(/^(тип|type|вид)/), iBrand = find(/^(бренд|brand|производ)/), iCom = find(/коммент|примеч|comment|модел|оборудован/);
    hdr.forEach(function (h, i) { if (/аналог|analog|cross|замен|equivalent/.test(h)) anCols.push(i); });
    var hasHeader = iArt >= 0 || anCols.length;
    if (!hasHeader) { iArt = 0; anCols = [1]; if (data[0].length > 2) for (var c = 2; c < data[0].length; c++) anCols.push(c); }
    if (iArt < 0) iArt = 0;
    var rows = hasHeader ? data.slice(1) : data, n = 0;
    rows.forEach(function (r) {
      var art = String(r[iArt] || '').trim(); if (!art) return;
      var an = anCols.map(function (c) { return String(r[c] || '').trim(); }).filter(Boolean).join('; ').replace(/\s*[,\n]\s*/g, '; ');
      if (!an) return;
      n++;
      base.push(['W' + ('0000' + n).slice(-5), iType >= 0 ? String(r[iType] || '').trim() : '', art, iBrand >= 0 ? String(r[iBrand] || '').trim() : '', an, src, iCom >= 0 ? String(r[iCom] || '').trim() : '', '']);
    });
  }
  if (!base.length) throw new Error('В файле не найдено строк вида «Артикул → аналоги». Нужны колонки «Артикул» и «Аналоги» (или формат листа «Cross замен фильтров»).');
  var mode = String(mode || 'append');   // 'append' — добавить к базе (по умолчанию), 'replace' — заменить прежнюю базу (ручные строки остаются), 'all' — заменить всё
  var sh = anlEnsureCross_(ss), hr = anlCrossHeaderRow_(sh), lr = sh.getLastRow(), skipped = 0;
  if (mode === 'append') {
    var have = {};
    anlReadCross_(ss).forEach(function (r) { have[anlNorm_(r.main) + '|' + anlNorm_(r.an)] = 1; });
    var fresh = base.filter(function (r) { var k = anlNorm_(r[2]) + '|' + anlNorm_(r[4]); if (have[k]) { skipped++; return false; } have[k] = 1; return true; });
    if (fresh.length) {
      ensureRows_(sh, lr + fresh.length);
      for (var i = 0; i < fresh.length; i += 2000) { var part = fresh.slice(i, i + 2000); sh.getRange(lr + 1 + i, 1, part.length, 8).setValues(part); }
    }
    return { loaded: fresh.length, skipped: skipped, manualKept: -1, mode: mode, stats: kzwCrossStats_(ss) };
  }
  if (mode === 'all' && lr > hr) sh.getRange(hr + 1, 1, lr - hr, 8).clearContent();
  var w = anlWriteCross_(ss, base);
  return { loaded: base.length, skipped: 0, manualKept: w.manual, mode: mode, stats: kzwCrossStats_(ss) };
}

// ═══════════════════════════════════════════════════════════════ v1.2: КРИТИЧЕСКИЙ ЗАПАС, МАССОВОЕ ЗАДАНИЕ, ЗАКРЕПЛЕНИЕ ШАПОК
var KZW_CRIT = { COL: 13, HDR: 'Критический запас', FLAG: 'KZW_FROZEN_V1' };   // «Справочник»!M

/** Колонка «Критический запас» в «Справочнике» (M): создать при первом обращении. */
function kzwEnsureCritCol_(ss) {
  var ws = kzwSheet_(ss, SH_SPR);
  if (ws.getMaxColumns() < KZW_CRIT.COL) ws.insertColumnsAfter(ws.getMaxColumns(), KZW_CRIT.COL - ws.getMaxColumns());
  var h = ws.getRange(1, KZW_CRIT.COL);
  if (String(h.getValue()).trim() !== KZW_CRIT.HDR) {
    var src = ws.getRange(1, 7);   // оформление как у «Мин. запас»
    h.setValue(KZW_CRIT.HDR); try { src.copyTo(h, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) {}
    ws.setColumnWidth(KZW_CRIT.COL, 110);
    var last = Math.max(ws.getLastRow(), 2); ws.getRange(2, KZW_CRIT.COL, Math.max(ws.getMaxRows() - 1, 1), 1).setBackground(COLOR_MANUAL);
    try { ws.getRange(1, KZW_CRIT.COL).setNote('Критический запас (веб-приложение): если задан, остаток ≤ критического — КРАСНЫЙ, ≤ минимального — ЖЁЛТЫЙ. Если пусто — красная зона ниже минимума, жёлтая по порогу «Настройки»!B2.'); } catch (e2) {}
  }
  return ws;
}
/**
 * Формулы статуса в «Остатках» (Y) с учётом критического запаса:
 *   крит. задан: остаток ≤ крит → КРАСНЫЙ: ниже критического; ≤ мин → ЖЁЛТЫЙ; иначе ОК
 *   крит. пуст:  как в основном скрипте (ниже минимума → КРАСНЫЙ, ≤ мин×(1+порог) → ЖЁЛТЫЙ).
 * Основной скрипт при пересборке переписывает Y своей формулой — поэтому вызывается после каждой пересборки из приложения и при загрузке данных.
 */
function kzwApplyCritFormulas_(ss, force) {
  var ws = kzwSheet_(ss, SH_OST), n = ws.getLastRow() - 1; if (n < 1) return 0;
  var probe = ws.getRange(2, C_STAT).getFormula();
  if (!force && probe.indexOf('исключено') !== -1) return 0;         // v1.6: формула уже с учётом исключений (и групп замен)
  if (!probe) return 0;                                   // пересборка не завершена / лист пуст
  if (kzLoadState_()) return 0;                           // идёт пересборка — не трогаем
  kzwEnsureCritCol_(ss); kzwEnsureExclCol_(ss);
  var fo = [];
  for (var i = 0; i < n; i++) {
    var r = i + 2, a = '$A' + r, g = 'G' + r, U = 'U' + r, X = '$X' + r;
    var crit = 'IFERROR(INDEX(\'Справочник\'!$M:$M,MATCH(' + a + ',\'Справочник\'!$A:$A,0)),"")';
    // v1.4: позиция в группе замен, общий остаток которой выше порога группы, в красную/жёлтую зону не попадает.
    //   порог группы = «Мин. запас группы» (Группы замен!C), если задан, иначе наибольший минимум среди позиций группы
    var gMin = 'IFERROR(INDEX(\'Группы замен\'!$C:$C,MATCH(' + X + ',\'Группы замен\'!$A:$A,0)),"")';
    var gThr = 'IF(' + gMin + '<>"",' + gMin + ',MAXIFS($U:$U,$X:$X,' + X + '))';
    var covered = 'AND(' + X + '<>"",SUMIF($X:$X,' + X + ',$G:$G)>' + gThr + ')';
    var exc = 'IFERROR(INDEX(\'Справочник\'!$N:$N,MATCH(' + a + ',\'Справочник\'!$A:$A,0)),"")';   // v1.6: исключённые (снятые / нерелевантные) — вне светофора
    fo.push(['=IF(' + exc + '<>"","— исключено",IF(' + U + '="","— мин. не задан",IF(' + covered + ',"ОК",IF(' + g + '<=0,"КРАСНЫЙ: нет в наличии",IF(' + crit + '<>"",IF(' + g + '<=' + crit + ',"КРАСНЫЙ: ниже критического",IF(' + g + '<=' + U + ',"ЖЁЛТЫЙ: на грани","ОК")),IF(' + g + '<' + U + ',"КРАСНЫЙ: ниже минимума",IF(' + g + '<=' + U + '*(1+' + REF_YELLOW + '),"ЖЁЛТЫЙ: на грани","ОК")))))))']);
  }
  ws.getRange(2, C_STAT, n, 1).setFormulas(locAll_(fo));
  SpreadsheetApp.flush();
  return n;
}
function kzwApplyCritFormulas(token) { kzwAuth_(token, 'wizard'); var n = kzwApplyCritFormulas_(kzwSs_(), true); SpreadsheetApp.flush(); return 'Формулы статуса (критический запас + покрытие группой замен) применены: строк ' + n; }

/**
 * Массовое задание параметров: keys — список ключей (или null = все позиции «Справочника»), fields {min,tgt,crit,grp,comment}:
 * undefined — не менять, null/'' — очистить, число — записать. Одна запись на столбец.
 */
function kzwBulkSet(token, keys, fields) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), ws = kzwEnsureCritCol_(ss), last = ws.getLastRow(); if (last < 2) return { changed: 0 };
  if (fields.excl !== undefined) kzwEnsureExclCol_(ss);
  var want = keys && keys.length ? {} : null; (keys || []).forEach(function (k) { want[String(k)] = 1; });
  var cols = { min: 7, tgt: 8, grp: 9, comment: 10, crit: KZW_CRIT.COL, excl: KZW_EXCL.COL }, changed = 0, touched = 0;
  var colA = ws.getRange(2, 1, last - 1, 1).getValues(), rowsSel = [];
  for (var i = 0; i < colA.length; i++) { var k = kzwKey_(colA[i][0]); if (k && (!want || want[k])) rowsSel.push(i); }
  if (!rowsSel.length) return { changed: 0 };
  Object.keys(cols).forEach(function (f) {
    if (fields[f] === undefined) return;
    var v = fields[f]; if (f === 'min' || f === 'tgt' || f === 'crit') { v = (v === null || v === '') ? '' : kzwNum_(v); if (v !== '' && v < 0) throw new Error('Отрицательное значение'); } else v = String(v || '');
    var rng = ws.getRange(2, cols[f], last - 1, 1), vals = rng.getValues(), dirty = false;
    rowsSel.forEach(function (i) { if (String(vals[i][0]) !== String(v)) { vals[i][0] = v; dirty = true; changed++; } });
    if (dirty) rng.setValues(vals);
    touched++;
  });
  if (fields.grp !== undefined) rowsSel.forEach(function (i) { try { kzwXrefSetKeyGroup_(ss, colA[i][0], String(fields.grp || '').trim()); } catch (e) {} });
  SpreadsheetApp.flush();
  return { changed: changed, rows: rowsSel.length, fields: touched };
}

/** Закрепить строку заголовка на листах, где она не закреплена (Справочник, Группы замен, Номенклатура, Данные 1С, Настройки, правила, прочие табличные). */
function kzwFreezeHeaders_(ss, force) {
  var props = PropertiesService.getScriptProperties();
  if (!force && props.getProperty(KZW_CRIT.FLAG)) return [];
  var known = {}; known[SH_SPR] = 1; known[SH_GRP] = 1; known[SH_NOM] = 1; known[SH_RAW] = 1; known[SH_RA] = 1; known[SH_RB] = 1; known[SH_RT] = 1; known[SH_SOST] = 1; known[KZW_REC.SHEET] = KZW_REC.H; known[SH_HOM] = 3;
  var done = [];
  ss.getSheets().forEach(function (sh) {
    try {
      if (sh.getFrozenRows() > 0 || sh.isSheetHidden()) return;
      var name = sh.getName(), row = known[name] || 0;
      if (!row) {   // эвристика: первая из 6 строк с ≥3 непустыми ячейками и данными ниже
        var lr = Math.min(sh.getLastRow(), 7), lc = Math.min(sh.getLastColumn(), 30); if (lr < 2 || lc < 3) return;
        var v = sh.getRange(1, 1, lr, lc).getValues();
        for (var i = 0; i < lr - 1; i++) { var n1 = v[i].filter(function (x) { return x !== ''; }).length, n2 = v[i + 1].filter(function (x) { return x !== ''; }).length; if (n1 >= 3 && n2 >= 2) { row = i + 1; break; } }
        if (!row) return;
      }
      sh.setFrozenRows(row); done.push(name + ' (' + row + ')');
    } catch (e) {}
  });
  props.setProperty(KZW_CRIT.FLAG, '1');
  return done;
}
function kzwFreezeHeaders(token) { kzwAuth_(token, 'wizard'); var d = kzwFreezeHeaders_(kzwSs_(), true); return d.length ? 'Закреплены шапки: ' + d.join(', ') : 'На всех листах шапка уже закреплена'; }

// ═══════════════════════════════════════════════════════════════ v1.3: ПОСТУПЛЕНИЯ (партии 1С), СНИМКИ ОСТАТКОВ ПО ПОЗИЦИЯМ, СТАТИСТИКА ПО ТОВАРУ
var KZW_SNAP = { SHEET: 'Снимки остатков', KEEP: 36, H: 1 };

/**
 * Партии из «Данных 1С» (строки выгрузки «Остатки по складам»), привязанные к ключам «Остатков».
 * Поступление = дата партии (ТТН); в выгрузке есть только партии, которые ещё лежат на складе.
 * Возвращает [{d, key, wh, inc, qty, price, sum, sup, party}].
 */
function kzwArrivals(token) {
  kzwAuth_(token);
  var ss = kzwSs_(), wsRaw = ss.getSheetByName(SH_RAW), out = [];
  if (!wsRaw || wsRaw.getLastRow() < 2) return { rows: out, loaded: '' };
  var raw = wsRaw.getRange(2, 1, wsRaw.getLastRow() - 1, 10).getValues();
  var nom = loadNomen_(ss), whMap = loadWarehouses_(ss);
  var ksOf = {}; kzwStock_(ss).forEach(function (it) { ksOf[canonKey_(it.key)] = it.ks; ksOf[it.ks] = it.ks; });
  for (var r = 0; r < raw.length; r++) {
    var nm = normName_(raw[r][2]); if (!nm) continue;
    var rec = nom[nm], ky = rec ? canonKey_(rec.key) : canonKey_(nm), ks = ksOf[ky] || ksOf[keyStr_(ky)] || ky;
    var w = whMap[whNorm_(normName_(raw[r][1]))], qty = toDbl_(raw[r][6]), pr = kzwNum_(raw[r][7]), dt = toDate_(raw[r][9]);
    out.push({ d: dt ? kzwDateStr_(dt) : '', key: ks, wh: w ? w.group : 'Прочие', inc: w ? !!w.inc : true, qty: qty, price: Math.round(pr * 100) / 100, sum: Math.round(kzwNum_(raw[r][8]) * 100) / 100,
      sup: String(raw[r][5] || '').trim(), party: String(raw[r][4] || '').trim() });
  }
  var lastLoad = ''; try { lastLoad = kzwDateTimeStr_(ss.getSheetByName(SH_SET).getRange('B4').getValue()); } catch (e) {}
  return { rows: out, loaded: lastLoad };
}

/** Снимок остатков по позициям на дату загрузки 1С — для динамики по товару. Пишется один раз на каждую новую загрузку (при чтении данных). */
function kzwSnapshot_(ss, items) {
  var set = ss.getSheetByName(SH_SET); if (!set) return 0;
  var load = set.getRange('B4').getValue(); if (!(load instanceof Date)) return 0;
  var stamp = kzwDateStr_(load), sh = ss.getSheetByName(KZW_SNAP.SHEET);
  if (!sh) {
    sh = ss.insertSheet(KZW_SNAP.SHEET, ss.getNumSheets());
    sh.getRange(1, 1, 1, 5).setValues([['Дата загрузки', 'Ключ', 'Остаток в контроле', 'Стоимость, BYN', 'Статус']]).setBackground(COLOR_HEAD).setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1); sh.getRange('B:B').setNumberFormat('@'); try { sh.hideSheet(); } catch (e) {}
  }
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('KZW_SNAP_LAST') === stamp) return 0;
  var last = sh.getLastRow(), dates = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { var d = r[0] instanceof Date ? kzwDateStr_(r[0]) : String(r[0]); if (d) dates[d] = 1; });
  if (dates[stamp]) { props.setProperty('KZW_SNAP_LAST', stamp); return 0; }
  if (sh.getMaxColumns() < 5) sh.insertColumnsAfter(sh.getMaxColumns(), 5 - sh.getMaxColumns());
  if (String(sh.getRange(1, 5).getValue()) !== 'Статус') sh.getRange(1, 5).setValue('Статус').setBackground(COLOR_HEAD).setFontColor('#FFFFFF').setFontWeight('bold');
  var rows = items.map(function (it) { return [load, it.key, it.tot, it.value, String(it.stat || '')]; });
  if (!rows.length) return 0;
  // ограничение глубины: оставить последние KEEP загрузок
  var dl = Object.keys(dates).sort();
  if (dl.length >= KZW_SNAP.KEEP) {
    var drop = {}; dl.slice(0, dl.length - KZW_SNAP.KEEP + 1).forEach(function (d) { drop[d] = 1; });
    var all = sh.getRange(2, 1, last - 1, 5).getValues().filter(function (r) { var d = r[0] instanceof Date ? kzwDateStr_(r[0]) : String(r[0]); return !drop[d]; });
    sh.getRange(2, 1, last - 1, 5).clearContent();
    if (all.length) sh.getRange(2, 1, all.length, 5).setValues(all);
    last = all.length + 1;
  }
  ensureRows_(sh, last + rows.length + 5);
  for (var i = 0; i < rows.length; i += 2000) { var part = rows.slice(i, i + 2000); sh.getRange(last + 1 + i, 1, part.length, 5).setValues(part); }
  sh.getRange(last + 1, 1, rows.length, 1).setNumberFormat('dd.MM.yyyy');
  props.setProperty('KZW_SNAP_LAST', stamp);
  return rows.length;
}
/** Динамика остатков по ключам: {dates:[...], series:{key:[tot...]}, value:{key:[...]}} (пропуски = null). keys — до 400 ключей. */
function kzwSnapshots(token, keys) {
  kzwAuth_(token);
  var ss = kzwSs_(), sh = ss.getSheetByName(KZW_SNAP.SHEET), res = { dates: [], series: {}, value: {} };
  if (!sh || sh.getLastRow() < 2) return res;
  var want = {}; (keys || []).slice(0, 400).forEach(function (k) { want[String(k)] = 1; });
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(), dset = {};
  v.forEach(function (r) { var d = r[0] instanceof Date ? kzwDateStr_(r[0]) : String(r[0]); if (d) dset[d] = 1; });
  res.dates = Object.keys(dset).sort(); var di = {}; res.dates.forEach(function (d, i) { di[d] = i; });
  v.forEach(function (r) {
    var k = kzwKey_(r[1]); if (!want[k]) return;
    var d = r[0] instanceof Date ? kzwDateStr_(r[0]) : String(r[0]); if (di[d] === undefined) return;
    if (!res.series[k]) { res.series[k] = res.dates.map(function () { return null; }); res.value[k] = res.dates.map(function () { return null; }); }
    res.series[k][di[d]] = kzwNum_(r[2]); res.value[k][di[d]] = kzwNum_(r[3]);
  });
  return res;
}

// ═══════════════════════════════════════════════════════════════ v1.6: ИСКЛЮЧЕНИЯ, ЗАКАЗЫ
var KZW_EXCL = { COL: 14, HDR: 'Исключено' };   // «Справочник»!N — причина исключения (снято с производства / нерелевантно); пусто — позиция в контроле

/** Колонка «Исключено» в «Справочнике» (N): создать при первом обращении. */
function kzwEnsureExclCol_(ss) {
  var ws = kzwEnsureCritCol_(ss);
  if (ws.getMaxColumns() < KZW_EXCL.COL) ws.insertColumnsAfter(ws.getMaxColumns(), KZW_EXCL.COL - ws.getMaxColumns());
  var h = ws.getRange(1, KZW_EXCL.COL);
  if (String(h.getValue()).trim() !== KZW_EXCL.HDR) {
    h.setValue(KZW_EXCL.HDR); try { ws.getRange(1, 7).copyTo(h, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) {}
    ws.setColumnWidth(KZW_EXCL.COL, 160);
    ws.getRange(2, KZW_EXCL.COL, Math.max(ws.getMaxRows() - 1, 1), 1).setBackground('#e7e6e6');
    try { h.setNote('Исключено (веб-приложение): любой текст — причина (снято с производства, нерелевантно…). Исключённая позиция не попадает в красную/жёлтую зону, «Критический запас» и список заказа; её остаток по-прежнему учитывается в стоимости склада и в остатке группы замен. Пусто — позиция контролируется.'); } catch (e2) {}
  }
  return ws;
}

/** Исключить / вернуть в контроль список ключей. reason — причина ('' = вернуть). */
function kzwSetExcluded(token, keys, reason) {
  kzwAuth_(token, 'wizard');
  var r = kzwBulkSet(token, keys, { excl: String(reason || '').trim() });
  try { kzwApplyCritFormulas_(kzwSs_(), false); } catch (e) {}
  return r;
}

// ─── лист «Заказы»: сохранённые списки заказа (песочница → таблица)
var KZW_ORD = { SHEET: 'Заказы', HDR: ['ID заказа', 'Название', 'Статус', 'Создан', 'Изменён', 'Ключ', 'Артикул', 'Наименование', 'Бренд', 'Тип', 'Ед.', 'Кол-во к заказу', 'Цена, BYN', 'Сумма, BYN', 'Группа замен', 'Остаток на момент', 'Мин. запас', 'Склады (фильтр)', 'Поставщик', 'Примечание'], STATUSES: ['черновик', 'согласован', 'отправлен поставщику', 'получен', 'отменён'] };
function kzwOrdSheet_(ss, create) {
  var ws = ss.getSheetByName(KZW_ORD.SHEET);
  if (!ws && create) {
    ws = ss.insertSheet(KZW_ORD.SHEET);
    ws.getRange(1, 1, 1, KZW_ORD.HDR.length).setValues([KZW_ORD.HDR]).setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
    ws.setFrozenRows(1); ws.setColumnWidth(8, 320);
  }
  return ws;
}
/** Все сохранённые заказы: [{id,name,status,created,updated,lines:[{key,art,name,brand,type,unit,qty,price,sum,grp,stock,min,wh,sup,note}]}]. */
function kzwOrders(token, ssOpt) {
  kzwAuth_(token);
  var ws = kzwOrdSheet_(ssOpt || kzwSs_(), false); if (!ws || ws.getLastRow() < 2) return { orders: [], statuses: KZW_ORD.STATUSES };
  var v = ws.getRange(2, 1, ws.getLastRow() - 1, KZW_ORD.HDR.length).getValues(), by = {}, list = [];
  v.forEach(function (r) {
    var id = String(r[0] || '').trim(); if (!id) return;
    var o = by[id]; if (!o) { o = by[id] = { id: id, name: String(r[1] || ''), status: String(r[2] || 'черновик'), created: kzwDateTimeStr_(r[3]), updated: kzwDateTimeStr_(r[4]), lines: [] }; list.push(o); }
    o.lines.push({ key: kzwKey_(r[5]), art: String(r[6] || ''), name: String(r[7] || ''), brand: String(r[8] || ''), type: String(r[9] || ''), unit: String(r[10] || ''), qty: kzwNum_(r[11]), price: kzwNum_(r[12]), sum: kzwNum_(r[13]), grp: String(r[14] || ''), stock: kzwNum_(r[15]), min: (r[16] === '' ? null : kzwNum_(r[16])), wh: String(r[17] || ''), sup: String(r[18] || ''), note: String(r[19] || '') });
  });
  list.forEach(function (o) { o.n = o.lines.length; o.sum = o.lines.reduce(function (a, l) { return a + l.sum; }, 0); });
  return { orders: list.reverse(), statuses: KZW_ORD.STATUSES };
}
/** Сохранить заказ (новый или заменить строки существующего id). order = {id?, name, status?, lines:[...]}. */
function kzwSaveOrder(token, order) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), ws = kzwOrdSheet_(ss, true), now = new Date();
  var id = String(order.id || '').trim(), created = now, col = ws.getLastRow() >= 2 ? ws.getRange(2, 1, ws.getLastRow() - 1, 4).getValues() : [];
  if (!id) {   // новый номер: Zггммдд-ччмм, при совпадении — суффикс -2, -3…
    var base = 'Z' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMdd-HHmm'), used = {}; col.forEach(function (r) { used[String(r[0]).trim()] = 1; });
    id = base; for (var k = 2; used[id]; k++) id = base + '-' + k;
  } else {     // заменить строки существующего заказа (удаление снизу вверх), сохранить дату создания
    for (var i = col.length - 1; i >= 0; i--) if (String(col[i][0]).trim() === id) { if (col[i][3] instanceof Date) created = col[i][3]; ws.deleteRow(i + 2); }
  }
  var lines = (order.lines || []).filter(function (l) { return l && l.key; });
  if (!lines.length) return kzwOrders(token);
  var status = KZW_ORD.STATUSES.indexOf(order.status) !== -1 ? order.status : 'черновик';
  var rows = lines.map(function (l) { var q = kzwNum_(l.qty), p = kzwNum_(l.price); return [id, String(order.name || ''), status, created, now, String(l.key), String(l.art || ''), String(l.name || ''), String(l.brand || ''), String(l.type || ''), String(l.unit || ''), q, p, Math.round(q * p * 100) / 100, String(l.grp || ''), kzwNum_(l.stock), (l.min === null || l.min === undefined) ? '' : kzwNum_(l.min), String(l.wh || ''), String(l.sup || ''), String(l.note || '')]; });
  ws.getRange(ws.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();
  var res = kzwOrders(token); res.savedId = id; return res;
}
/** Сменить статус заказа (черновик → согласован → отправлен поставщику → получен / отменён). */
function kzwSetOrderStatus(token, id, status) {
  kzwAuth_(token, 'wizard');
  if (KZW_ORD.STATUSES.indexOf(status) === -1) throw new Error('Неизвестный статус: ' + status);
  var ws = kzwOrdSheet_(kzwSs_(), false); if (!ws || ws.getLastRow() < 2) throw new Error('Заказов нет');
  var rng = ws.getRange(2, 1, ws.getLastRow() - 1, 5), v = rng.getValues(), n = 0, now = new Date();
  v.forEach(function (r) { if (String(r[0]).trim() === id) { r[2] = status; r[4] = now; n++; } });
  if (!n) throw new Error('Заказ ' + id + ' не найден');
  rng.setValues(v); SpreadsheetApp.flush();
  return kzwOrders(token);
}
/** Удалить заказ целиком. */
function kzwDeleteOrder(token, id) {
  kzwAuth_(token, 'wizard');
  var ws = kzwOrdSheet_(kzwSs_(), false); if (!ws || ws.getLastRow() < 2) return kzwOrders(token);
  var col = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) if (String(col[i][0]).trim() === id) ws.deleteRow(i + 2);
  SpreadsheetApp.flush();
  return kzwOrders(token);
}


// ═══════════════════════════════════════════════════════════════ v1.8: УСКОРИТЕЛЬ CLOUDFLARE (kz-edge)
/*
 *  Worker kz-edge держит готовый пакет данных (kzwGetData, партии, заказы, листы) в D1 и отдаёт его за
 *  доли секунды; все записи он передаёт сюда без изменений. Чтобы кэш не отставал:
 *   1) kzwAuth_(…, 'wizard') перед каждой правкой дёргает POST <EDGE_URL>/edge/bump (ключ EDGE_KEY);
 *   2) Worker раз в 5 минут вызывает kzwEdgeSig и сравнивает подпись (ловит правки прямо в листах);
 *   3) после пересборки (kzwStatus: busy → false) Worker перечитывает данные сам.
 *  Подключение: WIZARD → «Пароль» → «⚡ Ускоритель» (kzwEdgeSetup) — адрес Worker'а и ключ из CONNECT.txt.
 */
var KZW_EDGE = { URL: 'EDGE_URL', KEY: 'EDGE_KEY' };
function kzwEdgeProps_() { var p = PropertiesService.getScriptProperties(); return { url: String(p.getProperty(KZW_EDGE.URL) || '').replace(/\/+$/, ''), key: String(p.getProperty(KZW_EDGE.KEY) || '') }; }
function kzwEdgeBump_() {
  try {
    var e = kzwEdgeProps_(); if (!e.url || !e.key) return;
    var c = CacheService.getScriptCache(); if (c.get('kzw_bump')) return;   // не чаще раза в 5 с
    c.put('kzw_bump', '1', 5);
    UrlFetchApp.fetch(e.url + '/edge/bump', { method: 'post', contentType: 'application/json', payload: '{}', headers: { 'X-Edge-Key': e.key }, muteHttpExceptions: true });
  } catch (err) {}
}
/** Подпись данных, которые влияют на пакет приложения: настройки, параметры контроля, группы, веб-страницы, заказы, список листов, состояние пересборки. ~1 с. */
function kzwEdgeSig(token) {
  kzwAuth_(token);
  var ss = kzwSs_(), parts = [], add = function (x) { parts.push(typeof x === 'string' ? x : JSON.stringify(x)); };
  var st = ss.getSheetByName(SH_SET); if (st) add(st.getRange('B2:B6').getValues());
  var spr = ss.getSheetByName(SH_SPR); if (spr && spr.getLastRow() > 1) add(spr.getRange(2, 7, spr.getLastRow() - 1, Math.max(1, Math.min(spr.getMaxColumns(), KZW_EXCL.COL) - 6)).getValues());
  var grp = ss.getSheetByName(SH_GRP); if (grp && grp.getLastRow() > 1) add(grp.getRange(2, 1, grp.getLastRow() - 1, 4).getValues());
  var pg = ss.getSheetByName(KZW.SHEET_PAGES); if (pg && pg.getLastRow() > 1) add(pg.getRange(2, 1, pg.getLastRow() - 1, 7).getValues());
  var ord = ss.getSheetByName(KZW_ORD.SHEET); if (ord && ord.getLastRow() > 1) add(ord.getRange(2, 1, ord.getLastRow() - 1, 5).getValues());
  var ost = ss.getSheetByName(SH_OST); if (ost) add([ost.getLastRow(), String(ost.getRange(2, C_STAT).getFormula()).length]);
  add(ss.getSheets().map(function (x) { return x.getName(); }));
  var state = null; try { state = kzLoadState_(); } catch (e) {}
  var str = parts.join('\u0001'), h = 5381; for (var i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
  return { sig: (h >>> 0).toString(36) + '-' + str.length, busy: !!state, lastLoad: st ? kzwDateTimeStr_(st.getRange('B4').getValue()) : '', ver: KZW.VERSION, at: kzwDateTimeStr_(new Date()) };
}
/** Подключить ускоритель: адрес Worker'а + ключ; проверяется /edge/ping. */
function kzwEdgeSetup(token, url, key) {
  kzwAuth_(token, 'wizard');
  url = String(url || '').trim().replace(/\/+$/, ''); key = String(key || '').trim();
  if (!/^https:\/\/[a-z0-9.-]+\.workers\.dev$|^https:\/\/[^\/\s]+$/.test(url)) throw new Error('Адрес ускорителя должен быть вида https://kz-edge.<имя>.workers.dev');
  if (key.length < 16) throw new Error('Ключ слишком короткий — скопируйте его из CONNECT.txt');
  var r = UrlFetchApp.fetch(url + '/edge/ping', { headers: { 'X-Edge-Key': key }, muteHttpExceptions: true });
  var code = r.getResponseCode(), body = {}; try { body = JSON.parse(r.getContentText()); } catch (e) {}
  if (code !== 200 || !body.ok) throw new Error('Ускоритель не ответил (' + code + '): ' + (body.error || r.getContentText().slice(0, 120)));
  var p = PropertiesService.getScriptProperties(); p.setProperty(KZW_EDGE.URL, url); p.setProperty(KZW_EDGE.KEY, key);
  try { UrlFetchApp.fetch(url + '/edge/bump', { method: 'post', contentType: 'application/json', payload: '{}', headers: { 'X-Edge-Key': key }, muteHttpExceptions: true }); } catch (e2) {}
  return kzwEdgeStatus(token);
}
function kzwEdgeOff(token) { kzwAuth_(token, 'wizard'); var p = PropertiesService.getScriptProperties(); p.deleteProperty(KZW_EDGE.URL); p.deleteProperty(KZW_EDGE.KEY); return { on: false }; }
function kzwEdgeStatus(token) {
  kzwAuth_(token);
  var e = kzwEdgeProps_(); if (!e.url) return { on: false };
  var res = { on: true, url: e.url };
  try { var r = UrlFetchApp.fetch(e.url + '/edge/status', { headers: { 'X-Edge-Key': e.key }, muteHttpExceptions: true }); res.code = r.getResponseCode(); try { res.edge = JSON.parse(r.getContentText()); } catch (e1) { res.edge = null; } res.okPing = res.code === 200 && !!(res.edge && res.edge.ok); }
  catch (err) { res.okPing = false; res.error = String(err && err.message || err); }
  return res;
}


// ═══════════════════════════════════════════════════════════════ v1.9: БИТРИКС24
/*
 *  Локальное приложение Битрикс24 работает через ускоритель (kz-edge): Битрикс открывает <EDGE_URL>/b24/app,
 *  Worker проверяет пользователя (user.current на портале) и просит здесь токен — kzwB24Login (ключ EDGE_KEY).
 *  Роль: WIZARD — если e-mail или ID пользователя Битрикс24 перечислен в B24_WIZARDS, иначе ПРОСМОТР.
 *  Вкладка в сделке: Worker берёт товары сделки и спрашивает kzwB24Lookup — остатки / статусы / группы замен.
 *  Уведомления: после новой выгрузки 1С (новый снимок) kzwB24Notify_ шлёт через входящий вебхук B24_HOOK
 *  сообщение получателям B24_NOTIFY (ID пользователей — im.notify.system; chatNN — im.message.add).
 */
var KZW_B24 = { DOMAIN: 'B24_DOMAIN', WIZARDS: 'B24_WIZARDS', HOOK: 'B24_HOOK', NOTIFY: 'B24_NOTIFY', ON: 'B24_NOTIFY_ON' };
function kzwB24Props_() {
  var p = PropertiesService.getScriptProperties();
  return { domain: String(p.getProperty(KZW_B24.DOMAIN) || '').toLowerCase().replace(/^https?:\/\//, '').replace(/\/.*$/, ''), wizards: String(p.getProperty(KZW_B24.WIZARDS) || ''), hook: String(p.getProperty(KZW_B24.HOOK) || '').replace(/\/+$/, ''), notify: String(p.getProperty(KZW_B24.NOTIFY) || ''), notifyOn: p.getProperty(KZW_B24.ON) !== '0' };
}
function kzwB24Setup(token, cfg) {
  kzwAuth_(token, 'wizard');
  var p = PropertiesService.getScriptProperties(), c = cfg || {};
  var dom = String(c.domain || '').trim().toLowerCase().replace(/^https?:\/\//, '').replace(/\/.*$/, '');
  if (dom && !/^[a-z0-9.-]+\.bitrix24\.[a-z]+$|^[a-z0-9.-]+\.[a-z]{2,}$/.test(dom)) throw new Error('Домен портала должен быть вида company.bitrix24.by');
  var hook = String(c.hook || '').trim().replace(/\/+$/, '');
  if (hook === '__keep__') hook = kzwB24Props_().hook;   // поле пустое — оставить прежний вебхук
  if (hook && !/^https:\/\/[^\/\s]+\/rest\/\d+\/[A-Za-z0-9]+$/.test(hook)) throw new Error('Входящий вебхук должен быть вида https://portal.bitrix24.by/rest/1/xxxxxxxx');
  p.setProperty(KZW_B24.DOMAIN, dom); p.setProperty(KZW_B24.WIZARDS, String(c.wizards || '').trim()); p.setProperty(KZW_B24.HOOK, hook); p.setProperty(KZW_B24.NOTIFY, String(c.notify || '').trim()); p.setProperty(KZW_B24.ON, c.notifyOn === false ? '0' : '1');
  return kzwB24Status(token);
}
function kzwB24Status(token) {
  kzwAuth_(token);
  var b = kzwB24Props_(), e = kzwEdgeProps_();
  var res = { domain: b.domain, wizards: b.wizards, hookSet: !!b.hook, hookMasked: b.hook ? b.hook.replace(/\/rest\/(\d+)\/[A-Za-z0-9]+$/, '/rest/$1/••••••') : '', notify: b.notify, notifyOn: b.notifyOn, edge: e.url, app: e.url ? e.url + '/b24/app' : '', install: e.url ? e.url + '/b24/install' : '' };
  if (b.hook) { try { var r = kzwB24Rest_('profile', {}); res.hookOk = !!(r && r.result); res.hookUser = r && r.result ? (r.result.NAME + ' ' + r.result.LAST_NAME).trim() : ''; res.hookErr = r && r.error_description || ''; } catch (err) { res.hookOk = false; res.hookErr = String(err.message || err); } }
  return res;
}
function kzwB24Rest_(method, params) {
  var b = kzwB24Props_(); if (!b.hook) throw new Error('Входящий вебхук Битрикс24 не задан');
  var r = UrlFetchApp.fetch(b.hook + '/' + method + '.json', { method: 'post', contentType: 'application/json', payload: JSON.stringify(params || {}), muteHttpExceptions: true });
  try { return JSON.parse(r.getContentText()); } catch (e) { return { error: 'bad_response', error_description: r.getContentText().slice(0, 200) }; }
}
/** Пробное уведомление получателям (кнопка «Проверить» в приложении). */
function kzwB24Test(token) {
  kzwAuth_(token, 'wizard');
  var n = kzwB24Send_('Проверка связи: уведомления «Контроль неснижаемого запаса» подключены. ' + (kzwEdgeProps_().url || ''));
  return 'Отправлено сообщений: ' + n.sent + (n.errors.length ? ' · ошибки: ' + n.errors.join('; ') : '');
}
function kzwB24Send_(text) {
  var b = kzwB24Props_(), out = { sent: 0, errors: [] }; if (!b.hook || !b.notify) return out;
  b.notify.split(/[,;\s]+/).filter(Boolean).forEach(function (to) {
    try {
      var r = /^chat/i.test(to) ? kzwB24Rest_('im.message.add', { DIALOG_ID: to, MESSAGE: text }) : kzwB24Rest_('im.notify.system', { USER_ID: to.replace(/^user/i, ''), MESSAGE: text });
      if (r && r.result) out.sent++; else out.errors.push(to + ': ' + (r && r.error_description || 'нет ответа'));
    } catch (e) { out.errors.push(to + ': ' + e.message); }
  });
  return out;
}
/** После новой выгрузки 1С: какие позиции стали КРАСНЫМИ по сравнению с предыдущим снимком → сообщение в Битрикс24. */
function kzwB24Notify_(ss, out) {
  var b = kzwB24Props_(); if (!b.hook || !b.notify || !b.notifyOn) return;
  var items = out.items, reds = items.filter(function (a) { return String(a[18] || '').indexOf('КРАСНЫЙ') === 0; });
  var prevRed = null, sh = ss.getSheetByName(KZW_SNAP.SHEET);
  if (sh && sh.getLastRow() > 1 && sh.getMaxColumns() >= 5) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues(), byDate = {};
    v.forEach(function (r) { var d = r[0] instanceof Date ? kzwDateStr_(r[0]) : String(r[0]); if (!d) return; (byDate[d] = byDate[d] || []).push(r); });
    var dates = Object.keys(byDate).sort(), prev = dates.length >= 2 ? byDate[dates[dates.length - 2]] : null;
    if (prev && prev.some(function (r) { return r[4]; })) { prevRed = {}; prev.forEach(function (r) { if (String(r[4] || '').indexOf('КРАСНЫЙ') === 0) prevRed[String(r[1])] = 1; }); }
  }
  var fresh = prevRed ? reds.filter(function (a) { return !prevRed[a[0]]; }) : null;
  var url = kzwEdgeProps_().url || '';
  var lines = ['[B]Контроль неснижаемого запаса[/B] — загружена выгрузка 1С от ' + (out.settings.lastLoad || '') + '.', 'КРАСНЫХ позиций: ' + reds.length + (fresh ? ', из них новых: ' + fresh.length : '') + '.'];
  (fresh && fresh.length ? fresh : []).slice(0, 15).forEach(function (a) { lines.push('• ' + a[2] + (a[3] ? ' · ' + a[3] : '') + ' — остаток ' + a[6] + (a[14] !== null && a[14] !== undefined ? ' / мин. ' + a[14] : '')); });
  if (fresh && fresh.length > 15) lines.push('… и ещё ' + (fresh.length - 15));
  if (url) lines.push('Открыть: ' + url);
  if (fresh && !fresh.length && reds.length) lines.push('Новых критических позиций нет.');
  kzwB24Send_(lines.join('\n'));
}
/** Вход пользователя Битрикс24 (вызывает Worker ключом EDGE_KEY): проверка портала, роль по списку WIZARD, токен. */
function kzwB24Login(key, info) {
  var e = kzwEdgeProps_(); if (!e.key || String(key || '') !== e.key) throw new Error('Неверный ключ ускорителя');
  var b = kzwB24Props_(), dom = String(info && info.domain || '').toLowerCase().replace(/^https?:\/\//, '').replace(/\/.*$/, '');
  if (!b.domain) throw new Error('В программе не указан домен портала Битрикс24 (WIZARD → Пароль → Битрикс24)');
  if (dom !== b.domain) throw new Error('Портал ' + dom + ' не совпадает с настроенным ' + b.domain);
  var email = String(info.email || '').toLowerCase(), id = String(info.userId || '');
  var wiz = b.wizards.toLowerCase().split(/[,;\s]+/).filter(Boolean), role = (wiz.indexOf(email) !== -1 || (id && wiz.indexOf(id) !== -1) || wiz.indexOf('id' + id) !== -1) ? 'wizard' : 'viewer';
  var token = kzwMintToken_(role, { b24: id, name: String(info.name || '') });
  return { token: token, role: role, name: String(info.name || ''), email: email };
}
/** Товары сделки → позиции склада: по артикулу/ключу (точно или как слово в названии), иначе по названию. */
function kzwB24Lookup(token, rows, ssOpt) {
  kzwAuth_(token);
  var ss = ssOpt || kzwSs_(), wsO = kzwSheet_(ss, SH_OST), lr = wsO.getLastRow(); if (lr < 2) return [];
  var v = wsO.getRange(2, 1, lr - 1, OST_COLS).getValues(), items = [];
  var norm = function (s) { return String(s == null ? '' : s).toUpperCase().replace(/[\s\-_\/.,]/g, ''); };
  for (var i = 0; i < v.length; i++) { var r = v[i]; if (!r[0]) continue; items.push({ key: String(r[0]), art: String(r[C_ART - 1] || ''), name: String(r[C_NAME - 1] || ''), brand: String(r[C_BRAND - 1] || ''), tot: kzwNum_(r[C_TOTAL - 1]), min: r[C_MIN - 1] === '' ? null : kzwNum_(r[C_MIN - 1]), stat: String(r[C_STAT - 1] || ''), grp: String(r[C_GRP - 1] || ''), nk: norm(r[0]), na: norm(r[C_ART - 1]), nn: norm(r[C_NAME - 1]) }); }
  var grpTot = {}; items.forEach(function (it) { if (it.grp) grpTot[it.grp] = (grpTot[it.grp] || 0) + it.tot; });
  var pick = function (it) { return { key: it.key, art: it.art, name: it.name, brand: it.brand, tot: it.tot, min: it.min, stat: it.stat, grp: it.grp, grpTot: it.grp ? grpTot[it.grp] : null }; };
  return (rows || []).slice(0, ssOpt ? 5000 : 60).map(function (q) {
    var name = String(q.name || ''), code = norm(q.code || ''), found = [], seen = {};
    var add = function (it) { if (!seen[it.key] && found.length < 6) { seen[it.key] = 1; found.push(pick(it)); } };
    if (code && code.length >= 3) items.forEach(function (it) { if (it.na === code || it.nk === code) add(it); });
    var toks = name.toUpperCase().split(/[\s,;()]+/).map(function (t) { return norm(t); }).filter(function (t) { return t.length >= 4 && /\d/.test(t); });
    toks.forEach(function (t) { items.forEach(function (it) { if (it.na === t || it.nk === t) add(it); }); });
    if (!found.length) toks.forEach(function (t) { items.forEach(function (it) { if (it.nn.indexOf(t) !== -1 || (it.na && t.indexOf(it.na) !== -1 && it.na.length >= 5)) add(it); }); });
    if (!found.length && name.length >= 6) { var nn = norm(name); items.forEach(function (it) { if (it.nn === nn) add(it); }); }
    // аналоги: другие позиции тех же групп замен
    var alt = []; found.forEach(function (f) { if (f.grp) items.forEach(function (it) { if (it.grp === f.grp && !seen[it.key] && alt.length < 8) { seen[it.key] = 1; alt.push(pick(it)); } }); });
    return { name: name, code: String(q.code || ''), qty: kzwNum_(q.qty), found: found, alt: alt };
  });
}


// ═══════════════════════════════════════════════════════════════ v2.0: СОВЕТНИК (расход, срок до исчерпания, ABC/XYZ, точка заказа, резерв, наборы, неликвид)
var KZW_ADV = { RESERVE: 'Резерв Б24', SETS: 'Наборы', DIGEST_FN: 'kzwAdvisorDigest', DIGEST_FLAG: 'KZW_DIGEST_ON' };
var KZW_ADV_DEF = { leadDefault: 30, lead: {}, safetyPct: 30, cyclDays: 30, abcA: 80, abcB: 95, xyzX: 0.5, xyzY: 1.0, deadDays: 365, season: {}, minDays: 14, reserveOn: true, alertClasses: 'AB' };
function kzwAdvSettings_(ss) { var st = kzwSettings_(ss), a = st.advisor || {}, out = {}; Object.keys(KZW_ADV_DEF).forEach(function (k) { out[k] = a[k] === undefined || a[k] === null || a[k] === '' ? KZW_ADV_DEF[k] : a[k]; }); return out; }

/** Расход по снимкам остатков: по ключу — списано (c), пришло (a), дней (d), загрузок (n), интервалы [дней, списано] для оценки колебаний. */
function kzwConsumption_(ss) {
  var sh = ss.getSheetByName(KZW_SNAP.SHEET), res = { keys: {}, loads: 0, days: 0, from: '', to: '' };
  if (!sh || sh.getLastRow() < 2) return res;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues(), byKey = {}, dates = {};
  for (var i = 0; i < v.length; i++) { var d = v[i][0] instanceof Date ? v[i][0].getTime() : Date.parse(v[i][0]); if (!d) continue; var k = String(v[i][1]); dates[d] = 1; (byKey[k] = byKey[k] || []).push([d, kzwNum_(v[i][2])]); }
  var dl = Object.keys(dates).map(Number).sort(function (a, b) { return a - b; }); res.loads = dl.length;
  if (dl.length >= 2) { res.days = Math.round((dl[dl.length - 1] - dl[0]) / 86400000); res.from = kzwDateStr_(new Date(dl[0])); res.to = kzwDateStr_(new Date(dl[dl.length - 1])); }
  Object.keys(byKey).forEach(function (k) {
    var seq = byKey[k].sort(function (a, b) { return a[0] - b[0]; }), c = 0, a = 0, iv = [];
    for (var j = 1; j < seq.length; j++) { var diff = seq[j][1] - seq[j - 1][1], days = (seq[j][0] - seq[j - 1][0]) / 86400000; if (diff < 0) { c += -diff; iv.push([days, -diff]); } else { a += diff; iv.push([days, 0]); } }
    res.keys[k] = { c: c, a: a, d: seq.length >= 2 ? Math.round((seq[seq.length - 1][0] - seq[0][0]) / 86400000) : 0, n: seq.length, iv: iv, last: seq[seq.length - 1][1] };
  });
  return res;
}
/** Резерв под открытые сделки Битрикс24 (входящий вебхук): лист «Резерв Б24» [ключ, кол-во, сделок, товар]. */
function kzwReserveRefresh_(ss) {
  var b = kzwB24Props_(); if (!b.hook) return { n: 0, off: true };
  var deals = [], start = 0;
  for (var guard = 0; guard < 40; guard++) { var r = kzwB24Rest_('crm.deal.list', { filter: { CLOSED: 'N' }, select: ['ID', 'TITLE'], start: start }); if (!r || !r.result) break; deals = deals.concat(r.result); if (r.next === undefined || r.next === null) break; start = r.next; }
  var rows = [];
  for (var i = 0; i < deals.length; i += 50) {
    var cmd = {}; deals.slice(i, i + 50).forEach(function (d) { cmd['d' + d.ID] = 'crm.deal.productrows.get?id=' + d.ID; });
    var br = kzwB24Rest_('batch', { halt: 0, cmd: cmd }); var res = br && br.result && br.result.result || {};
    Object.keys(res).forEach(function (k) { (res[k] || []).forEach(function (p) { rows.push({ deal: k.slice(1), name: String(p.PRODUCT_NAME || ''), code: String(p.PRODUCT_CODE || ''), qty: kzwNum_(p.QUANTITY) }); }); });
  }
  var by = {};
  if (rows.length) {
    var look = kzwB24Match_(ss, rows);
    look.forEach(function (q, idx) { if (q.found.length) { var f = q.found[0], x = by[f.key] || (by[f.key] = { qty: 0, deals: {}, name: f.name }); x.qty += rows[idx].qty; x.deals[rows[idx].deal] = 1; } });
  }
  var sh = ss.getSheetByName(KZW_ADV.RESERVE); if (!sh) { sh = ss.insertSheet(KZW_ADV.RESERVE, ss.getNumSheets()); sh.getRange(1, 1, 1, 5).setValues([['Ключ', 'Резерв, ед.', 'Сделок', 'Наименование', 'Обновлено']]).setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff'); sh.setFrozenRows(1); try { sh.hideSheet(); } catch (e) {} }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  var out = Object.keys(by).map(function (k) { return [k, by[k].qty, Object.keys(by[k].deals).length, by[k].name, new Date()]; });
  if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
  PropertiesService.getScriptProperties().setProperty('KZW_RESERVE_AT', kzwDateTimeStr_(new Date()));
  return { n: out.length, deals: deals.length, rows: rows.length };
}
function kzwReserveRefresh(token) { kzwAuth_(token, 'wizard'); return kzwReserveRefresh_(kzwSs_()); }
function kzwReserve_(ss) { var sh = ss.getSheetByName(KZW_ADV.RESERVE), m = {}; if (!sh || sh.getLastRow() < 2) return m; sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) { if (r[0]) m[String(r[0])] = { qty: kzwNum_(r[1]), deals: kzwNum_(r[2]) }; }); return m; }
/** Сопоставление названий/кодов с позициями (общее для вкладки сделки и резерва). */
function kzwB24Match_(ss, rows) { return kzwB24Lookup('__internal__', rows, ss); }
/** Наборы ТО: лист «Наборы» [Набор, Ключ/артикул, Кол-во]. Загрузка CSV (разделитель ; , или таб). */
function kzwUploadSets(token, text, mode) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), lines = String(text || '').replace(/^﻿/, '').split(/\r?\n/).filter(function (l) { return l.trim(); });
  if (!lines.length) throw new Error('Пустой файл');
  var sep = (typeof detectSeparator_ === 'function') ? null : null; sep = lines[0].indexOf('\t') !== -1 ? '\t' : (lines[0].split(';').length >= lines[0].split(',').length ? ';' : ',');
  var rows = [];
  lines.forEach(function (l, i) { var c = l.split(sep).map(function (x) { return x.trim().replace(/^"|"$/g, ''); }); if (i === 0 && /набор|set|kit/i.test(c[0]) && !/\d/.test(c[2] || '')) return; if (c[0] && c[1]) rows.push([c[0], c[1], kzwNum_(c[2]) || 1]); });
  if (!rows.length) throw new Error('Не найдено строк вида «Набор;Ключ или артикул;Кол-во»');
  var sh = ss.getSheetByName(KZW_ADV.SETS); if (!sh) { sh = ss.insertSheet(KZW_ADV.SETS, ss.getNumSheets()); sh.getRange(1, 1, 1, 3).setValues([['Набор', 'Ключ / артикул', 'Кол-во']]).setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff'); sh.setFrozenRows(1); }
  if (mode !== 'append' && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  var last = Math.max(sh.getLastRow(), 1); sh.getRange(last + 1, 1, rows.length, 3).setValues(rows);
  return { rows: rows.length, sets: Object.keys(rows.reduce(function (m, r) { m[r[0]] = 1; return m; }, {})).length };
}
function kzwSets_(ss) { var sh = ss.getSheetByName(KZW_ADV.SETS), m = {}; if (!sh || sh.getLastRow() < 2) return m; sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) { if (r[0] && r[1]) (m[String(r[0])] = m[String(r[0])] || []).push({ ref: String(r[1]), qty: kzwNum_(r[2]) || 1 }); }); return m; }
/** Автозакрытие заказов: «отправлен поставщику» → «получен», если по каждой строке остаток вырос не меньше чем на половину заказанного с момента заказа. */
function kzwOrdersAutoClose_(ss, items) {
  var ws = kzwOrdSheet_(ss, false); if (!ws || ws.getLastRow() < 2) return 0;
  var sn = ss.getSheetByName(KZW_SNAP.SHEET); if (!sn || sn.getLastRow() < 2) return 0;
  var v = sn.getRange(2, 1, sn.getLastRow() - 1, 3).getValues(), byDate = {};
  v.forEach(function (r) { var d = r[0] instanceof Date ? kzwDateStr_(r[0]) : String(r[0]); (byDate[d] = byDate[d] || {})[String(r[1])] = kzwNum_(r[2]); });
  var dates = Object.keys(byDate).sort(); if (dates.length < 2) return 0;
  var prev = byDate[dates[dates.length - 2]], now = {}; items.forEach(function (a) { now[a[0]] = a[6]; });
  var rng = ws.getRange(2, 1, ws.getLastRow() - 1, 20), rows = rng.getValues(), byId = {};
  rows.forEach(function (r, i) { var id = String(r[0]).trim(); if (!id || r[2] !== 'отправлен поставщику') return; (byId[id] = byId[id] || []).push(i); });
  var closed = 0;
  Object.keys(byId).forEach(function (id) {
    var ok = byId[id].every(function (i) { var k = String(rows[i][5]), q = kzwNum_(rows[i][11]); return (now[k] || 0) - (prev[k] || 0) >= Math.max(0.5 * q, 0.5); });
    if (ok) { byId[id].forEach(function (i) { rows[i][2] = 'получен'; rows[i][4] = new Date(); rows[i][19] = (rows[i][19] ? rows[i][19] + ' · ' : '') + 'закрыт автоматически по приходу'; }); closed++; }
  });
  if (closed) rng.setValues(rows);
  return closed;
}
function kzwItemsLite_(ss) {
  var spr = {}, wsS = ss.getSheetByName(SH_SPR);
  if (wsS && wsS.getLastRow() >= 2) { var cols = Math.min(wsS.getMaxColumns(), KZW_EXCL.COL), sv = wsS.getRange(2, 1, wsS.getLastRow() - 1, cols).getValues(); sv.forEach(function (r) { var k = kzwKey_(r[0]); if (k && !spr[k]) spr[k] = { crit: (r[KZW_CRIT.COL - 1] === '' || r[KZW_CRIT.COL - 1] === undefined || r[KZW_CRIT.COL - 1] === null) ? null : kzwNum_(r[KZW_CRIT.COL - 1]), excl: String(r[KZW_EXCL.COL - 1] == null ? '' : r[KZW_EXCL.COL - 1]).trim() }; }); }
  var wsO = kzwSheet_(ss, SH_OST), lr = wsO.getLastRow(), items = []; if (lr < 2) return items;
  var v = wsO.getRange(2, 1, lr - 1, OST_COLS).getValues();
  for (var i = 0; i < v.length; i++) { var r = v[i]; if (r[0] === '' || r[0] === null) continue; var k = kzwKey_(r[0]), sp = spr[k] || { crit: null, excl: '' };
    items.push({ key: k, art: kzwKey_(r[C_ART - 1]), name: String(r[C_NAME - 1] || ''), brand: String(r[C_BRAND - 1] || '') || 'Не определено', type: String(r[C_TYPE - 1] || ''), unit: String(r[C_UNIT - 1] || ''), tot: kzwNum_(r[C_TOTAL - 1]), min: (r[C_MIN - 1] === '' || r[C_MIN - 1] === null) ? null : kzwNum_(r[C_MIN - 1]), tgt: (r[C_TGT - 1] === '' || r[C_TGT - 1] === null) ? null : kzwNum_(r[C_TGT - 1]), crit: sp.crit, excl: sp.excl, price: Math.round(kzwNum_(r[C_PRICE - 1]) * 100) / 100, value: Math.round(kzwNum_(r[C_VALUE - 1]) * 100) / 100, old: kzwDateStr_(r[C_OLD - 1]), nw: kzwDateStr_(r[C_NEW - 1]), grp: String(r[C_GRP - 1] || '').trim(), stat: String(r[C_STAT - 1] || '') }); }
  return items;
}
/**
 * Советник: по каждой позиции — расход в день (по снимкам), класс ABC (по стоимости расхода, иначе по стоимости запаса) и XYZ
 * (колебания расхода), срок поставки по бренду, «хватит дней», точка заказа, расчётный минимум, рекомендуемое количество
 * с учётом «в пути» и резерва; плюс список рекомендаций (alerts) и неполные наборы.
 */
function kzwAdvisor(token) {
  kzwAuth_(token);
  var ss = kzwSs_(), A = kzwAdvSettings_(ss), items = kzwItemsLite_(ss), cons = kzwConsumption_(ss), reserve = A.reserveOn ? kzwReserve_(ss) : {}, sets = kzwSets_(ss);
  var ord = kzwOrders('__internal__', ss), transit = {}, lateOrders = [], now = new Date(), today = kzwDateStr_(now), month = now.getMonth() + 1;
  var leadOf = function (brand) { var l = A.lead && A.lead[brand]; return l === undefined || l === null || l === '' ? kzwNum_(A.leadDefault) || 30 : kzwNum_(l); };
  ord.orders.forEach(function (o) { if (o.status === 'согласован' || o.status === 'отправлен поставщику') { o.lines.forEach(function (l) { transit[l.key] = (transit[l.key] || 0) + l.qty; }); var age = (now - new Date(o.updated || o.created)) / 86400000, maxLead = Math.max.apply(null, o.lines.map(function (l) { return leadOf(l.brand); }).concat([0])); if (o.status === 'отправлен поставщику' && age > maxLead + 14) lateOrders.push({ id: o.id, name: o.name, days: Math.round(age), n: o.n }); } });
  var byKey = {}, byArt = {}; items.forEach(function (it) { byKey[it.key] = it; if (it.art) byArt[it.art.toUpperCase()] = it; });
  // группы замен: суммарный остаток
  var grpTot = {}; items.forEach(function (it) { if (it.grp) grpTot[it.grp] = (grpTot[it.grp] || 0) + it.tot; });
  var basisOk = cons.loads >= 2 && cons.days >= (kzwNum_(A.minDays) || 14);
  items.forEach(function (it) {
    var c = cons.keys[it.key], rate = null, cv = null;
    if (basisOk && c && c.n >= 2 && c.d >= 7) { rate = c.c / c.d; if (c.iv.length >= 4) { var rates = c.iv.map(function (x) { return x[0] > 0 ? x[1] / x[0] : 0; }), mean = rates.reduce(function (a, b) { return a + b; }, 0) / rates.length; if (mean > 0) { var sd = Math.sqrt(rates.reduce(function (a, b) { return a + (b - mean) * (b - mean); }, 0) / rates.length); cv = Math.round(sd / mean * 100) / 100; } } }
    it.rate = rate; it.consumed = c ? c.c : null; it.cv = cv; it.lead = leadOf(it.brand);
    var seas = A.season && A.season[it.type] && A.season[it.type][month - 1]; it.season = seas ? kzwNum_(seas) || 1 : 1;
    it.transit = transit[it.key] || 0; it.reserve = reserve[it.key] ? reserve[it.key].qty : 0;
    it.avail = Math.round((it.tot + it.transit - it.reserve) * 1000) / 1000;
    it.days = rate ? (it.avail <= 0 ? 0 : Math.round(it.avail / rate)) : null;
    it.rop = rate ? Math.ceil(rate * it.lead * (1 + kzwNum_(A.safetyPct) / 100) * it.season) : null;   // точка заказа = расчётный минимум
    it.autoMin = it.rop;
    var need = rate ? Math.ceil(rate * (it.lead + kzwNum_(A.cyclDays)) * it.season) - it.avail : ((it.tgt !== null ? it.tgt : (it.min !== null ? it.min : 0)) - it.avail);
    it.recQty = Math.max(0, Math.ceil(need * 1000) / 1000);
    it.deadDays = it.nw ? Math.round((now - new Date(it.nw)) / 86400000) : null;
    it.dead = !it.excl && it.tot > 0 && ((rate !== null && it.consumed === 0 && (c ? c.d : 0) >= Math.min(180, kzwNum_(A.deadDays))) || (rate === null && it.deadDays !== null && it.deadDays >= kzwNum_(A.deadDays)));
    it.cls = String(it.stat || '').indexOf('КРАСНЫЙ') === 0 ? 'red' : (String(it.stat || '').indexOf('ЖЁЛТЫЙ') === 0 ? 'yel' : (it.excl ? 'exc' : (it.min === null ? 'none' : 'grn')));
  });
  // ABC
  var useCons = basisOk && items.some(function (it) { return it.consumed > 0; });
  var ranked = items.filter(function (it) { return !it.excl; }).map(function (it) { return { it: it, w: useCons ? (it.consumed || 0) * it.price : it.value }; }).sort(function (a, b) { return b.w - a.w; });
  var totalW = ranked.reduce(function (a, x) { return a + x.w; }, 0), acc = 0;
  ranked.forEach(function (x) { acc += x.w; var sh = totalW ? acc / totalW * 100 : 100; x.it.abc = x.w <= 0 ? 'C' : (sh <= kzwNum_(A.abcA) ? 'A' : (sh <= kzwNum_(A.abcB) ? 'B' : 'C')); });
  items.forEach(function (it) { if (!it.abc) it.abc = it.excl ? '—' : 'C'; it.xyz = it.cv === null ? '—' : (it.cv < kzwNum_(A.xyzX) ? 'X' : (it.cv < kzwNum_(A.xyzY) ? 'Y' : 'Z')); it.grpTot = it.grp ? grpTot[it.grp] : null; });
  // рекомендации
  var alerts = [], K = String(A.alertClasses || 'AB'), isTop = function (it) { return K.indexOf(it.abc) !== -1; }, live = items.filter(function (it) { return !it.excl; });
  var sumQ = function (arr) { return Math.round(arr.reduce(function (a, it) { return a + it.recQty * it.price; }, 0)); };
  var orderNow = live.filter(function (it) { return it.rate && it.days !== null && it.days <= it.lead && isTop(it) && it.recQty > 0; });
  var orderSoon = live.filter(function (it) { return it.rate && it.days !== null && it.days > it.lead && it.days <= it.lead + 14 && isTop(it) && it.recQty > 0; });
  var redTop = live.filter(function (it) { return it.cls === 'red' && isTop(it) && !(it.grp && it.grpTot > (it.min || 0)) && it.transit <= 0; });
  var noMin = live.filter(function (it) { return it.min === null && it.rate > 0; });
  var minLow = live.filter(function (it) { return it.min !== null && it.autoMin !== null && isTop(it) && it.autoMin >= 2 * Math.max(it.min, 1) && it.autoMin - it.min >= 1; });
  var minHigh = live.filter(function (it) { return it.min !== null && it.autoMin !== null && it.min >= 3 && it.autoMin <= it.min * 0.4; });
  var reserveShort = live.filter(function (it) { return it.reserve > 0 && it.tot + it.transit < it.reserve; });
  var dead = live.filter(function (it) { return it.dead; });
  var push = function (id, level, title, text, keys, action) { alerts.push({ id: id, level: level, title: title, text: text, keys: (keys || []).slice(0, 300), n: keys ? keys.length : 0, action: action || '' }); };
  if (!basisOk) push('basis', 'info', 'Расход пока не оценён', 'Снимков остатков: ' + cons.loads + (cons.days ? ' за ' + cons.days + ' дн.' : '') + '. Расход, «хватит дней» и расчётные минимумы появятся после нескольких загрузок 1С (не меньше двух с разницей ≥ ' + (kzwNum_(A.minDays) || 14) + ' дней). Пока советы строятся по минимумам и стоимости запаса.');
  else if (cons.days < 60) push('basis', 'info', 'Оценка расхода ещё грубая', 'История: ' + cons.loads + ' загрузок за ' + cons.days + ' дн. (' + cons.from + ' – ' + cons.to + '). Точность вырастет после 2–3 месяцев регулярных загрузок.');
  if (orderNow.length) push('order_now', 'red', 'Заказать на этой неделе: ' + orderNow.length + ' поз. (' + K.split('').join('/') + '-класс)', 'Остатка с учётом «в пути» и резерва хватит меньше, чем на срок поставки. Сумма по рекомендуемому количеству ≈ ' + sumQ(orderNow).toLocaleString('ru-RU') + ' BYN.', orderNow.map(function (it) { return it.key; }), 'order');
  if (orderSoon.length) push('order_soon', 'yel', 'Точка заказа через 1–2 недели: ' + orderSoon.length + ' поз.', 'Запас закончится через срок поставки + до 14 дней. Удобно объединить с текущим заказом поставщику.', orderSoon.map(function (it) { return it.key; }), 'order');
  if (redTop.length) push('red_top', 'red', 'КРАСНЫЕ позиции ' + K.split('').join('/') + '-класса без заказа: ' + redTop.length, 'Важные позиции (по расходу или стоимости) уже в красной зоне, группа замен не покрывает, в пути ничего нет.', redTop.map(function (it) { return it.key; }), 'order');
  if (reserveShort.length) push('reserve', 'red', 'Не хватает под открытые сделки Битрикс24: ' + reserveShort.length + ' поз.', 'Товары в открытых сделках превышают остаток + в пути. Проверьте сроки сделок и закажите.', reserveShort.map(function (it) { return it.key; }), 'order');
  if (noMin.length) push('nomin', 'yel', 'Есть расход, но нет минимума: ' + noMin.length + ' поз.', 'Позиции расходуются, а светофор их не контролирует. Можно принять расчётные минимумы одной кнопкой.', noMin.map(function (it) { return it.key; }), 'automin');
  if (minLow.length) push('min_low', 'yel', 'Минимум занижен: ' + minLow.length + ' поз.', 'Расчётный минимум (расход × срок поставки × страховой запас) вдвое и более выше заданного — позиции попадут в красную зону слишком поздно.', minLow.map(function (it) { return it.key; }), 'automin');
  if (minHigh.length) push('min_high', 'info', 'Минимум, вероятно, завышен: ' + minHigh.length + ' поз.', 'Заданный минимум больше расчётного в 2,5 раза и более — лишние деньги в запасе. Проверьте перед снижением.', minHigh.map(function (it) { return it.key; }), 'automin');
  if (lateOrders.length) push('late', 'yel', 'Заказы дольше срока поставки: ' + lateOrders.length, lateOrders.slice(0, 5).map(function (o) { return o.id + (o.name ? ' «' + o.name + '»' : '') + ' — ' + o.days + ' дн., позиций ' + o.n; }).join('; ') + '. Уточните у поставщика или закройте заказ.', [], 'orders');
  if (dead.length) push('dead', 'info', 'Залежалый запас: ' + dead.length + ' поз. на ' + Math.round(dead.reduce(function (a, it) { return a + it.value; }, 0)).toLocaleString('ru-RU') + ' BYN', 'Без движения дольше ' + kzwNum_(A.deadDays) + ' дней. Кандидаты в исключения, распродажу или возврат поставщику.', dead.map(function (it) { return it.key; }), 'dead');
  // сезон: следующий месяц с коэффициентом > 1.2
  Object.keys(A.season || {}).forEach(function (t) { var arr = A.season[t] || [], nxt = kzwNum_(arr[month % 12]); if (nxt >= 1.2) push('season_' + t, 'info', 'Сезон: ' + t + ' — в следующем месяце расход ×' + nxt, 'По сезонным коэффициентам спрос вырастет. Расчётные минимумы и количества уже учитывают текущий месяц; проверьте запас до пика.', live.filter(function (it) { return it.type === t && isTop(it); }).map(function (it) { return it.key; }), 'items'); });
  // наборы
  var setsOut = [];
  Object.keys(sets).forEach(function (name) {
    var parts = sets[name].map(function (p) { var it = byKey[p.ref] || byArt[p.ref.toUpperCase()]; return { ref: p.ref, qty: p.qty, key: it ? it.key : '', name: it ? it.name : '', tot: it ? it.tot : null, avail: it ? it.avail : null, ok: it ? it.avail >= p.qty : false }; });
    var kits = parts.length ? Math.floor(Math.min.apply(null, parts.map(function (p) { return p.avail === null ? 0 : Math.floor(p.avail / p.qty); }))) : 0;
    setsOut.push({ name: name, parts: parts, kits: kits, missing: parts.filter(function (p) { return !p.ok; }) });
  });
  var incomplete = setsOut.filter(function (s) { return s.missing.length; });
  if (incomplete.length) push('sets', 'yel', 'Неполные наборы ТО: ' + incomplete.length + ' из ' + setsOut.length, incomplete.slice(0, 6).map(function (s) { return s.name + ' (нет: ' + s.missing.map(function (p) { return p.ref; }).join(', ') + ')'; }).join('; '), incomplete.reduce(function (a, s) { return a.concat(s.missing.filter(function (p) { return p.key; }).map(function (p) { return p.key; })); }, []), 'order');
  var excludedWithStock = items.filter(function (it) { return it.excl && it.tot > 0; });
  if (excludedWithStock.length) push('excl_stock', 'info', 'Исключённые позиции с остатком: ' + excludedWithStock.length + ' на ' + Math.round(excludedWithStock.reduce(function (a, it) { return a + it.value; }, 0)).toLocaleString('ru-RU') + ' BYN', 'Товар снят с производства / нерелевантен, но лежит на складе — распродажа или возврат.', excludedWithStock.map(function (it) { return it.key; }), 'excl');
  var order = { red: 0, yel: 1, info: 2 }; alerts.sort(function (a, b) { return order[a.level] - order[b.level]; });
  var cols = ['key', 'art', 'name', 'brand', 'type', 'unit', 'tot', 'min', 'tgt', 'crit', 'price', 'value', 'nw', 'grp', 'grpTot', 'stat', 'cls', 'excl', 'rate', 'consumed', 'cv', 'abc', 'xyz', 'lead', 'season', 'transit', 'reserve', 'avail', 'days', 'rop', 'autoMin', 'recQty', 'deadDays', 'dead'];
  return { at: kzwDateTimeStr_(now), today: today, cols: cols, items: items.map(function (it) { return cols.map(function (c) { return it[c] === undefined ? null : it[c]; }); }), alerts: alerts, sets: setsOut, basis: { ok: basisOk, loads: cons.loads, days: cons.days, from: cons.from, to: cons.to, useCons: useCons }, settings: A, reserveAt: PropertiesService.getScriptProperties().getProperty('KZW_RESERVE_AT') || '', reserveOn: !!kzwB24Props_().hook, digestOn: PropertiesService.getScriptProperties().getProperty(KZW_ADV.DIGEST_FLAG) === '1', transitKeys: Object.keys(transit).length };
}
/** Еженедельная сводка Советника в Битрикс24 (триггер по понедельникам, 9:00) и ручной запуск. */
function kzwAdvisorDigest() {
  var b = kzwB24Props_(); if (!b.hook || !b.notify) return 'Битрикс24 не настроен';
  var adv = kzwAdvisor('__internal__'), lines = ['[B]Советник — Контроль неснижаемого запаса[/B] (' + adv.at + ')'];
  adv.alerts.filter(function (a) { return a.level !== 'info' || a.id === 'dead'; }).slice(0, 8).forEach(function (a) { lines.push((a.level === 'red' ? '🔴 ' : (a.level === 'yel' ? '🟡 ' : '• ')) + a.title); });
  if (lines.length === 1) lines.push('Срочных рекомендаций нет.');
  var url = kzwEdgeProps_().url; if (url) lines.push('Открыть Советник: ' + url + '/?tab=advisor');
  var r = kzwB24Send_(lines.join('\n')); return 'Отправлено: ' + r.sent + (r.errors.length ? ' · ошибки: ' + r.errors.join('; ') : '');
}
function kzwAdvisorDigestNow(token) { kzwAuth_(token, 'wizard'); return kzwAdvisorDigest(); }
function kzwAdvisorSchedule(token, on) {
  kzwAuth_(token, 'wizard');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === KZW_ADV.DIGEST_FN) ScriptApp.deleteTrigger(t); });
  if (on) ScriptApp.newTrigger(KZW_ADV.DIGEST_FN).timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  PropertiesService.getScriptProperties().setProperty(KZW_ADV.DIGEST_FLAG, on ? '1' : '0');
  return { on: !!on };
}
/** v2.0: принять расчётные минимумы — pairs [[ключ, мин.], …]; пишет столбец G «Справочника» по каждому ключу своё значение. */
function kzwSetMinMany(token, pairs) {
  kzwAuth_(token, 'wizard');
  var ss = kzwSs_(), ws = kzwSheet_(ss, SH_SPR), last = ws.getLastRow(); if (last < 2 || !pairs || !pairs.length) return { changed: 0 };
  var want = {}; pairs.forEach(function (p) { var v = kzwNum_(p[1]); if (p[0] && v >= 0) want[String(p[0])] = Math.round(v * 1000) / 1000; });
  var colA = ws.getRange(2, 1, last - 1, 1).getValues(), rng = ws.getRange(2, 7, last - 1, 1), vals = rng.getValues(), changed = 0;
  for (var i = 0; i < colA.length; i++) { var k = kzwKey_(colA[i][0]); if (k && want[k] !== undefined && String(vals[i][0]) !== String(want[k])) { vals[i][0] = want[k]; changed++; } }
  if (changed) { rng.setValues(vals); SpreadsheetApp.flush(); }
  return { changed: changed };
}
