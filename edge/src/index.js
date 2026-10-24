/**
 * kz-edge — ускоритель программы «Контроль неснижаемого запаса» (Cloudflare Worker + D1).
 *
 * Что делает:
 *  - раздаёт страницу приложения (public/ = app/ сайта) и отвечает на /api так же, как Apps Script /exec
 *    ({"api":"kzwXxx","args":[…]} → {"ok":true,"result":…});
 *  - тяжёлые чтения (пакет данных, партии 1С, заказы, листы, сверка, аналоги, снимки) отдаёт из кэша D1
 *    за десятки миллисекунд; всё остальное (вход, правки, загрузка 1С) передаёт в Google без изменений;
 *  - кэш сбрасывается: после любой правки через Worker; по сигналу Apps Script (/edge/bump перед правкой WIZARD);
 *    по завершении пересборки (kzwStatus busy → false); по расписанию каждые 5 минут — сравнение подписи kzwEdgeSig.
 *  Данные остаются в Google-таблице; на Cloudflare — только их копия для быстрого чтения и список сессий (токен → роль).
 *
 *  Ради лимита процессорного времени бесплатного тарифа большие ответы Google НЕ разбираются и не сжимаются:
 *  текст JSON хранится в D1 как есть и склеивается в ответ строками.
 */

const VER = '2.0';
const TTL = { sheet: 15 * 60 * 1000, charts: 60 * 60 * 1000 };
const H = { 'content-type': 'application/json;charset=utf-8', 'access-control-allow-origin': '*', 'access-control-allow-headers': 'content-type,x-edge-key', 'cache-control': 'no-store' };

const mem = { ready: false, inflight: new Map(), dirtyAt: 0, lastBusy: null, svcTok: null };

export default {
  async fetch(req, env, ctx) {
    const url = new URL(req.url);
    try {
      if (req.method === 'OPTIONS') return new Response(null, { status: 204, headers: H });
      if (url.pathname === '/api' || url.pathname === '/exec') return api(req, url, env, ctx);
      if (url.pathname.startsWith('/edge/')) return edge(req, url, env, ctx);
      if (url.pathname.startsWith('/b24/')) return b24(req, url, env, ctx);
      if (url.pathname === '/ping') return json({ ok: true, app: 'kz', ver: VER, edge: true });
    } catch (e) {
      return json({ ok: false, error: 'edge: ' + (e && e.message || e) }, 500);
    }
    return env.ASSETS.fetch(req);
  },
  async scheduled(ev, env, ctx) { ctx.waitUntil(syncCheck(env)); }
};

const json = (obj, status) => new Response(JSON.stringify(obj), { status: status || 200, headers: H });
const okText = (resultText) => new Response('{"ok":true,"result":' + resultText + '}', { headers: H });
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// ─── D1
async function ready(env) {
  if (mem.ready) return;
  await env.DB.batch([
    env.DB.prepare('CREATE TABLE IF NOT EXISTS kv (k TEXT PRIMARY KEY, v TEXT, at INTEGER)'),
    env.DB.prepare('CREATE TABLE IF NOT EXISTS sess (tok TEXT PRIMARY KEY, role TEXT, at INTEGER)'),
    env.DB.prepare('CREATE TABLE IF NOT EXISTS st (k TEXT PRIMARY KEY, v TEXT)')
  ]);
  mem.ready = true;
}
async function kvGet(env, k, ttl) { const r = await env.DB.prepare('SELECT v, at FROM kv WHERE k = ?').bind(k).first(); if (!r || (ttl && Date.now() - r.at > ttl)) return null; return r; }
async function kvPut(env, k, text) { await env.DB.prepare('INSERT OR REPLACE INTO kv (k, v, at) VALUES (?, ?, ?)').bind(k, text, Date.now()).run(); return text; }
async function stGet(env, k) { const r = await env.DB.prepare('SELECT v FROM st WHERE k = ?').bind(k).first(); return r ? r.v : null; }
async function stSet(env, k, v) { await env.DB.prepare('INSERT OR REPLACE INTO st (k, v) VALUES (?, ?)').bind(k, String(v)).run(); }
async function sessRole(env, tok) { if (!tok) return null; const r = await env.DB.prepare('SELECT role FROM sess WHERE tok = ?').bind(tok).first(); return r ? r.role : null; }
async function sessPut(env, tok, role) { if (tok) await env.DB.prepare('INSERT OR REPLACE INTO sess (tok, role, at) VALUES (?, ?, ?)').bind(tok, role || 'viewer', Date.now()).run(); }
async function sessDel(env, tok) { if (tok) await env.DB.prepare('DELETE FROM sess WHERE tok = ?').bind(tok).run(); }
async function invalidate(env) { mem.dirtyAt = Date.now(); await env.DB.prepare("DELETE FROM kv").run(); }

// ─── Google Apps Script: ответ {"ok":true,"result":…} — результат вырезается как текст, без разбора
async function gasRaw(env, api, args) {
  if (!env.GAS_URL) throw new Error('В настройках Worker не задан GAS_URL (адрес /exec)');
  const r = await fetch(env.GAS_URL, { method: 'POST', headers: { 'content-type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ api, args: args || [] }), redirect: 'follow' });
  const txt = await r.text();
  if (txt.startsWith('{"ok":true,"result":') && txt.endsWith('}')) return txt.slice(20, -1);
  let j; try { j = JSON.parse(txt); } catch (e) { throw new Error(/<html/i.test(txt) ? 'Google ответил страницей вместо данных — проверьте GAS_URL (/exec, доступ «Все», KZ_Web ≥ 1.8)' : 'Непонятный ответ Google'); }
  if (j && j.ok) return JSON.stringify(j.result === undefined ? null : j.result);
  const err = new Error((j && j.error) || 'Ошибка Google'); err.gas = true; throw err;
}
async function gas(env, api, args) { return JSON.parse(await gasRaw(env, api, args)); }   // для небольших ответов
async function svcToken(env, renew) {
  if (!renew) { if (mem.svcTok) return mem.svcTok; const t = await stGet(env, 'svc_tok'); if (t) { mem.svcTok = t; return t; } }
  const r = await gas(env, 'kzwLogin', ['']); if (!r || r.error || !r.token) throw new Error('Служебный вход не удался: ' + (r && r.error || ''));
  mem.svcTok = r.token; await stSet(env, 'svc_tok', r.token); await sessPut(env, r.token, 'viewer'); return r.token;
}
async function withSvc(env, fn) { try { return await fn(await svcToken(env)); } catch (e) { if (/AUTH:/.test(e.message)) return fn(await svcToken(env, true)); throw e; } }

// ─── кэш текста с защитой от одновременных запросов
async function cachedText(env, key, ttl, loader) {
  const hit = await kvGet(env, key, ttl); if (hit) return hit.v;
  if (mem.inflight.has(key)) return mem.inflight.get(key);
  const p = (async () => { try { return await kvPut(env, key, await loader()); } finally { mem.inflight.delete(key); } })();
  mem.inflight.set(key, p); return p;
}
/** Мета по роли (kzwBootMeta без token/role — они подставляются при ответе). */
async function metaText(env, role, token) {
  return cachedText(env, 'meta:' + role, 0, async () => { const t = await gasRaw(env, 'kzwBootMeta', [token]); return t; });
}
async function dataText(env, token) { return cachedText(env, 'data', 0, () => gasRaw(env, 'kzwGetData', [token])); }
/** Ответ kzwBootstrap из кусков: {"token":…,"role":…,"edge":{…},"data":<data>, <мета без первой скобки> */
function composeBoot(token, role, meta, data, cached) {
  return '{"token":' + JSON.stringify(token) + ',"role":' + JSON.stringify(role) + ',"edge":' + JSON.stringify({ ver: VER, cached: !!cached }) + ',"data":' + data + ',' + meta.replace(/^\{/, '');
}
/** Фоновое обновление после правки: подождать, перечитать токеном автора (мета его роли, данные, заказы, партии). */
async function prefetch(env, token, role, delayMs) {
  const my = Date.now(); mem.dirtyAt = my;
  if (delayMs) await sleep(delayMs);
  if (mem.dirtyAt !== my) return;
  try {
    await kvPut(env, 'data', await gasRaw(env, 'kzwGetData', [token]));
    await kvPut(env, 'meta:' + (role || 'viewer'), await gasRaw(env, 'kzwBootMeta', [token]));
    await kvPut(env, 'orders', await gasRaw(env, 'kzwOrders', [token]));
    await kvPut(env, 'arr', await gasRaw(env, 'kzwArrivals', [token]));
  } catch (e) {}
}
async function refreshAll(env) { await invalidate(env); await withSvc(env, (t) => prefetch(env, t, 'viewer', 0)); }

// ─── /api
const READS = { kzwGetRange: 1, kzwGetCellFormula: 1, kzwEdgeStatus: 1, kzwEdgeSig: 1, kzwBootMeta: 1, kzwB24Status: 1, kzwB24Lookup: 1, kzwAdvisorDigestNow: 1 };
function hashKeys(keys) { let h = 5381; const s = (keys || []).slice().sort().join('|'); for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return (h >>> 0).toString(36); }

async function api(req, url, env, ctx) {
  if (req.method === 'GET') { if (url.searchParams.get('ping')) return json({ ok: true, app: 'kz', ver: VER, edge: true, title: 'Контроль неснижаемого запаса' }); return json({ ok: false, error: 'POST {api,args}' }, 405); }
  let body; try { body = await req.json(); } catch (e) { return json({ ok: false, error: 'Неверный формат запроса' }); }
  const name = String(body.api || ''), args = Array.isArray(body.args) ? body.args : [];
  if (!/^kzw[A-Z][A-Za-z0-9]*$/.test(name) || name === 'kzwB24Login') return json({ ok: false, error: 'Функция недоступна: ' + name });
  await ready(env);
  try { return okText(await handle(name, args, env, ctx)); }
  catch (e) { return json({ ok: false, error: String(e && e.message || e) }); }
}

/** Возвращает ТЕКСТ результата. */
async function handle(name, args, env, ctx) {
  const token = args[0];
  switch (name) {
    case 'kzwLogin': { const t = await gasRaw(env, name, args); const r = JSON.parse(t); if (r && r.token) await sessPut(env, r.token, r.role); return t; }
    case 'kzwLogout': await sessDel(env, token); return gasRaw(env, name, args).catch(() => 'true');
    case 'kzwBootstrap': return bootstrap(args, env);
    case 'kzwGetData': return viaCache(env, token, 'data', 0, name, args, (tok) => dataText(env, tok));
    case 'kzwArrivals': return viaCache(env, token, 'arr', 0, name, args);
    case 'kzwOrders': return viaCache(env, token, 'orders', 0, name, args);
    case 'kzwGetSheet': return viaCache(env, token, 'sheet:' + args[1], TTL.sheet, name, args);
    case 'kzwGetCharts': return viaCache(env, token, 'charts:' + args[1], TTL.charts, name, args);
    case 'kzwReconcile': if (args[1]) return kvPut(env, 'recon', await gasRaw(env, name, args)); return viaCache(env, token, 'recon', 0, name, args);
    case 'kzwAnalogs': if (args[1]) return kvPut(env, 'anl', await gasRaw(env, name, args)); return viaCache(env, token, 'anl', 0, name, args);
    case 'kzwSnapshots': return viaCache(env, token, 'snap:' + hashKeys(args[1]), 0, name, args);
    case 'kzwAdvisor': return viaCache(env, token, 'adv', 0, name, args);
    case 'kzwStatus': {
      const t = await gasRaw(env, name, args), st = JSON.parse(t);
      if (mem.lastBusy === true && st.busy === false) ctx.waitUntil(afterRebuild(env, token));
      mem.lastBusy = !!st.busy; return t;
    }
    default: {
      if (READS[name]) return gasRaw(env, name, args);
      const t = await gasRaw(env, name, args);               // запись — в Google
      await invalidate(env);
      if (name === 'kzwUpload1C' || name === 'kzwStartRebuild') mem.lastBusy = true;
      else ctx.waitUntil(prefetch(env, token, await sessRole(env, token), 1500));
      return t;
    }
  }
}
async function afterRebuild(env, token) { await invalidate(env); await sleep(2000); await prefetch(env, token, await sessRole(env, token), 0); }
/** Чтение через кэш: незнакомый токен — первый раз через Google (он проверит сессию). */
async function viaCache(env, token, key, ttl, name, args, loader) {
  const role = await sessRole(env, token);
  if (!role) { const t = await gasRaw(env, name, args); await sessPut(env, token, 'viewer'); return kvPut(env, key, t); }
  return loader ? loader(token) : cachedText(env, key, ttl, () => gasRaw(env, name, args));
}
async function bootstrap(args, env) {
  let token = args[0], password = args[1], role = null, fresh = false;
  if (!token) { const lg = await gas(env, 'kzwLogin', [password || '']); if (!lg || lg.error) return JSON.stringify(lg); token = lg.token; role = lg.role; await sessPut(env, token, role); }
  else role = await sessRole(env, token);
  if (!role) {   // незнакомый токен: мета из Google проверит сессию и скажет роль
    const mt = await gasRaw(env, 'kzwBootMeta', [token]), m = JSON.parse(mt); role = m.role; await sessPut(env, token, role); await kvPut(env, 'meta:' + role, mt); fresh = true;
  }
  const hadData = !!(await kvGet(env, 'data', 0));
  const [meta, data] = await Promise.all([metaText(env, role, token), dataText(env, token)]);
  return composeBoot(token, role, meta, data, hadData && !fresh);
}

// ─── /edge/* — служебные вызовы Apps Script и установщика (ключ EDGE_KEY)
async function edge(req, url, env, ctx) {
  const key = req.headers.get('x-edge-key') || url.searchParams.get('key') || '';
  if (!env.EDGE_KEY || key !== env.EDGE_KEY) return json({ ok: false, error: 'Неверный ключ ускорителя' }, 401);
  await ready(env);
  const p = url.pathname.slice(6);
  if (p === 'ping') return json({ ok: true, ver: VER, gas: !!env.GAS_URL });
  if (p === 'bump') { await invalidate(env); ctx.waitUntil((async () => { await sleep(7000); if (Date.now() - mem.dirtyAt >= 6500) await refreshAll(env).catch(() => {}); })()); return json({ ok: true }); }
  if (p === 'refresh') { await refreshAll(env); return json({ ok: true }); }
  if (p === 'status') {
    const rows = (await env.DB.prepare('SELECT k, at, length(v) AS n FROM kv ORDER BY k').all()).results || [];
    const sess = await env.DB.prepare('SELECT COUNT(*) AS n FROM sess').first();
    return json({ ok: true, ver: VER, gas: !!env.GAS_URL, cache: rows.map(r => ({ k: r.k, ageSec: Math.round((Date.now() - r.at) / 1000), kb: Math.round(r.n / 1024) })), sessions: sess ? sess.n : 0, sig: await stGet(env, 'sig'), sigAt: await stGet(env, 'sig_at'), sigErr: await stGet(env, 'sig_err') });
  }
  return json({ ok: false, error: 'нет такого пути' }, 404);
}

// ─── расписание: сравнить подпись данных, при расхождении перечитать
async function syncCheck(env) {
  await ready(env);
  try {
    const s = await withSvc(env, (t) => gas(env, 'kzwEdgeSig', [t]));
    const prev = await stGet(env, 'sig');
    await stSet(env, 'sig_at', new Date().toISOString()); await stSet(env, 'sig_err', '');
    if (s.busy) { mem.lastBusy = true; return; }
    if (mem.lastBusy === true || prev !== s.sig) { mem.lastBusy = false; await refreshAll(env); await stSet(env, 'sig', s.sig); }
    await env.DB.prepare('DELETE FROM sess WHERE at < ?').bind(Date.now() - 24 * 3600 * 1000).run();
  } catch (e) { await stSet(env, 'sig_err', String(e && e.message || e)); }
}

// ─── /b24/* — локальное приложение Битрикс24 (серверное): Битрикс открывает /b24/app POST-запросом
//     с DOMAIN, AUTH_ID, PLACEMENT, PLACEMENT_OPTIONS; пользователь проверяется user.current на портале,
//     роль и токен выдаёт Apps Script (kzwB24Login, ключ EDGE_KEY). /b24/install — установка (вкладка в сделке).
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const page = (body, status) => new Response('<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Контроль запаса</title><style>body{font:14px/1.4 Arial,sans-serif;margin:0;padding:14px;color:#222;background:#fff}h3{margin:0 0 10px;color:#1f3864}table{border-collapse:collapse;width:100%;font-size:13px}th,td{border:1px solid #d0d4d9;padding:5px 8px;vertical-align:top}th{background:#1f3864;color:#fff;text-align:left}.st{display:inline-block;padding:2px 8px;border-radius:5px;font-weight:bold;font-size:12px}.red{background:#f8cbad;color:#7a1c00}.yel{background:#ffe699;color:#7a5a00}.grn{background:#c6e0b4;color:#1e5e1e}.none{background:#eee;color:#666}.exc{background:#e7e6e6;color:#555;text-decoration:line-through}.hint{color:#666;font-size:12px}a{color:#1a4d8f}.num{text-align:right;white-space:nowrap}.alt{color:#555;font-size:12px;margin-top:3px}button{background:#1a73e8;color:#fff;border:0;border-radius:6px;padding:8px 14px;cursor:pointer;font:inherit}</style></head><body>' + body + '</body></html>', { status: status || 200, headers: { 'content-type': 'text/html;charset=utf-8', 'cache-control': 'no-store' } });
async function b24rest(env, domain, auth, method, params) {
  const base = env.B24_TEST_BASE ? env.B24_TEST_BASE : 'https://' + domain + '/rest';
  const r = await fetch(base + '/' + method + '.json', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(Object.assign({}, params || {}, { auth })) });
  try { return await r.json(); } catch (e) { return { error: 'bad_response' }; }
}
function stCls(st) { st = String(st || ''); if (st.indexOf('— исключено') === 0) return 'exc'; if (st.indexOf('КРАСНЫЙ') === 0) return 'red'; if (st.indexOf('ЖЁЛТЫЙ') === 0) return 'yel'; if (st.indexOf('ОК') === 0) return 'grn'; return 'none'; }
async function b24(req, url, env, ctx) {
  await ready(env);
  const p = url.pathname.slice(5);
  let form = null; if (req.method === 'POST') { try { form = await req.formData(); } catch (e) { form = null; } }
  const g = (k) => String((form && form.get(k)) || url.searchParams.get(k) || '');
  const domain = g('DOMAIN').toLowerCase(), auth = g('AUTH_ID'), placement = g('PLACEMENT') || 'DEFAULT';
  let opts = {}; try { opts = JSON.parse(g('PLACEMENT_OPTIONS') || '{}') || {}; } catch (e) {}
  if (!domain || !auth) return page('<h3>Контроль неснижаемого запаса</h3><p>Эта страница открывается из Битрикс24 (приложение в левом меню или вкладка «Остатки на складе» в сделке).</p>', 400);
  if (!/^[a-z0-9.-]+$/.test(domain)) return page('<p>Неверный адрес портала.</p>', 400);
  const me = await b24rest(env, domain, auth, 'user.current', {});
  if (!me || !me.result) return page('<p>Не удалось проверить пользователя Битрикс24: ' + esc(me && (me.error_description || me.error) || 'нет ответа') + '</p>', 403);
  const u = me.result, name = ((u.NAME || '') + ' ' + (u.LAST_NAME || '')).trim();
  if (p === 'install') {
    const handler = url.origin + '/b24/app', out = [];
    for (const pl of ['CRM_DEAL_DETAIL_TAB']) {
      await b24rest(env, domain, auth, 'placement.unbind', { PLACEMENT: pl, HANDLER: handler });
      const r = await b24rest(env, domain, auth, 'placement.bind', { PLACEMENT: pl, HANDLER: handler, TITLE: 'Остатки на складе', DESCRIPTION: 'Контроль неснижаемого запаса: остатки, статусы и аналоги по товарам сделки' });
      out.push(pl + ': ' + (r && r.result ? 'подключено' : 'ошибка — ' + esc(r && (r.error_description || r.error) || '')));
    }
    return page('<h3>Установка «Контроль неснижаемого запаса»</h3><p>Портал: <b>' + esc(domain) + '</b>, установил: ' + esc(name) + '.</p><ul>' + out.map(x => '<li>' + x + '</li>').join('') + '</ul><p class="hint">Пункт левого меню включается в настройках приложения («Показывать в левом меню»). В программе (WIZARD → Пароль → Битрикс24) должен быть указан домен портала <b>' + esc(domain) + '</b> и список сотрудников с уровнем WIZARD.</p><p><button onclick="finish()">Завершить установку</button></p><script src="//api.bitrix24.com/api/v1/"></script><script>function finish(){try{BX24.init(function(){BX24.installFinish();});}catch(e){alert("Готово — закройте окно");}}</script>');
  }
  let lg; try { lg = await gas(env, 'kzwB24Login', [env.EDGE_KEY, { domain, userId: String(u.ID), email: u.EMAIL || '', name, admin: !!u.ADMIN }]); }
  catch (e) { return page('<h3>Вход через Битрикс24 не выполнен</h3><p>' + esc(e.message) + '</p><p class="hint">WIZARD: откройте программу → «Пароль» → блок «Битрикс24» и укажите домен портала ' + esc(domain) + '.</p>', 403); }
  await sessPut(env, lg.token, lg.role);
  if (/^CRM_(DEAL|LEAD|SMART_INVOICE|QUOTE)_DETAIL_TAB$/.test(placement)) {
    const id = opts.ID || opts.id || '';
    const method = placement === 'CRM_LEAD_DETAIL_TAB' ? 'crm.lead.productrows.get' : (placement === 'CRM_QUOTE_DETAIL_TAB' ? 'crm.quote.productrows.get' : 'crm.deal.productrows.get');
    const pr = await b24rest(env, domain, auth, method, { id });
    const rows = (pr && pr.result) || [];
    if (!rows.length) return page('<h3>Остатки на складе</h3><p class="hint">В ' + (placement === 'CRM_LEAD_DETAIL_TAB' ? 'лиде' : 'сделке') + ' нет товаров' + (pr && pr.error_description ? ' (' + esc(pr.error_description) + ')' : '') + '.</p>');
    let look = []; try { look = await gas(env, 'kzwB24Lookup', [lg.token, rows.map(r => ({ name: r.PRODUCT_NAME, code: r.PRODUCT_CODE || '', qty: r.QUANTITY }))]); } catch (e) { return page('<p>Ошибка поиска: ' + esc(e.message) + '</p>', 500); }
    const link = (k) => url.origin + '/?b24tok=' + encodeURIComponent(lg.token) + '&b24=1&key=' + encodeURIComponent(k);
    const line = (f) => '<span class="st ' + stCls(f.stat) + '">' + esc(f.stat) + '</span> <a href="' + link(f.key) + '" target="_blank">' + esc(f.key) + '</a> ' + esc(f.name) + (f.brand ? ' · ' + esc(f.brand) : '') + ' — остаток <b>' + esc(f.tot) + '</b>' + (f.min !== null && f.min !== undefined ? ' / мин. ' + esc(f.min) : '') + (f.grp ? ' · группа ' + esc(f.grp) + ' (' + esc(f.grpTot) + ')' : '');
    const trs = look.map(q => '<tr><td>' + esc(q.name) + (q.code ? '<div class="hint">' + esc(q.code) + '</div>' : '') + '</td><td class="num">' + esc(q.qty) + '</td><td>' + (q.found.length ? q.found.map(f => '<div>' + line(f) + (q.qty > f.tot ? ' <span class="st red">не хватает ' + esc(Math.round((q.qty - f.tot) * 1000) / 1000) + '</span>' : '') + '</div>').join('') : '<span class="hint">не найдено в контроле запаса</span>') + (q.alt.length ? '<div class="alt">Аналоги (группа замен): ' + q.alt.map(a => esc(a.key) + ' — ' + esc(a.tot)).join('; ') + '</div>' : '') + '</td></tr>').join('');
    return page('<h3>Остатки на складе по товарам</h3><table><tr><th>Товар в сделке</th><th>Кол-во</th><th>Позиция склада · статус · остаток в контроле</th></tr>' + trs + '</table><p class="hint">Данные 1С из программы «Контроль неснижаемого запаса». <a href="' + url.origin + '/?b24tok=' + encodeURIComponent(lg.token) + '&b24=1" target="_blank">Открыть программу</a> · ' + esc(name) + ' (' + (lg.role === 'wizard' ? 'WIZARD' : 'ПРОСМОТР') + ')</p>');
  }
  // приложение целиком (левое меню): передаём токен странице
  return new Response(null, { status: 303, headers: { location: url.origin + '/?b24tok=' + encodeURIComponent(lg.token) + '&b24=1' } });
}
