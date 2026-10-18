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

const VER = '1.8';
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
const READS = { kzwGetRange: 1, kzwGetCellFormula: 1, kzwEdgeStatus: 1, kzwEdgeSig: 1, kzwBootMeta: 1 };
function hashKeys(keys) { let h = 5381; const s = (keys || []).slice().sort().join('|'); for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return (h >>> 0).toString(36); }

async function api(req, url, env, ctx) {
  if (req.method === 'GET') { if (url.searchParams.get('ping')) return json({ ok: true, app: 'kz', ver: VER, edge: true, title: 'Контроль неснижаемого запаса' }); return json({ ok: false, error: 'POST {api,args}' }, 405); }
  let body; try { body = await req.json(); } catch (e) { return json({ ok: false, error: 'Неверный формат запроса' }); }
  const name = String(body.api || ''), args = Array.isArray(body.args) ? body.args : [];
  if (!/^kzw[A-Z][A-Za-z0-9]*$/.test(name)) return json({ ok: false, error: 'Функция недоступна: ' + name });
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
