// Установка / обновление ускорителя kz-edge (Cloudflare Workers + D1). Запуск: install.cmd или update.cmd
// Требуется Node.js 18+. Аккаунт Cloudflare — бесплатный.
import { execSync, spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { createInterface } from 'node:readline/promises';
import { randomBytes } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url)), root = join(here, '..');
const update = process.argv.includes('--update');
const W = ['npx', '-y', 'wrangler@4.40.0'];
const log = (s) => console.log('\n' + s);
function run(args, opts = {}) {
  const r = spawnSync(W[0], W.slice(1).concat(args), { cwd: here, stdio: opts.input ? ['pipe', 'pipe', 'inherit'] : 'inherit', shell: true, encoding: 'utf8', input: opts.input });
  if (r.status !== 0 && !opts.soft) throw new Error('wrangler ' + args[0] + ' завершился с ошибкой');
  return r.stdout || '';
}
function runOut(args) { const r = spawnSync(W[0], W.slice(1).concat(args), { cwd: here, shell: true, encoding: 'utf8' }); process.stdout.write(r.stdout || ''); process.stderr.write(r.stderr || ''); return (r.stdout || '') + (r.stderr || ''); }

(async () => {
  const nodeMajor = +process.versions.node.split('.')[0]; if (nodeMajor < 18) { console.error('Нужен Node.js 18 или новее: https://nodejs.org'); process.exit(1); }
  let toml = readFileSync(join(here, 'wrangler.toml'), 'utf8');
  log('1/5  Сборка страницы (edge/public из app/)…');
  execSync('node tools/build.mjs', { cwd: root, stdio: 'inherit' });

  // Восстановление настроек, если wrangler.toml перезаписан новым архивом (REPLACE_WITH_…): берём из CONNECT.txt или из Cloudflare
  const keyFile = join(here, 'CONNECT.txt'), saved = existsSync(keyFile) ? readFileSync(keyFile, 'utf8') : '';
  if (update && /REPLACE_WITH/.test(toml)) {
    log('Настройки в wrangler.toml сброшены (новый архив распакован поверх старого) — восстанавливаю…');
    run(['whoami']);
    let gas = (saved.match(/GAS_URL:\s*(\S+)/) || [])[1];
    if (!gas) { const rl = createInterface({ input: process.stdin, output: process.stdout }); gas = (await rl.question('\nАдрес веб-приложения Apps Script (…/exec): ')).trim().replace(/\/dev$/, '/exec').replace(/\?.*$/, ''); rl.close(); }
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(gas)) { console.error('Это не ссылка /exec'); process.exit(1); }
    toml = toml.replace(/GAS_URL = "[^"]*"/, `GAS_URL = "${gas}"`);
    let id = (saved.match(/D1_ID:\s*(\S+)/) || [])[1];
    if (!id) { const out = runOut(['d1', 'list', '--json']); try { id = (JSON.parse(out.slice(out.indexOf('['))).find(d => d.name === 'kz-db') || {}).uuid; } catch (e) {} }
    if (!id) { console.error('Не удалось определить ID базы kz-db — откройте панель Cloudflare → Storage & Databases → D1 → kz-db и впишите database_id в wrangler.toml'); process.exit(1); }
    toml = toml.replace('REPLACE_WITH_D1_ID', id);
    writeFileSync(join(here, 'wrangler.toml'), toml);
  }
  if (!update) {
    log('2/5  Вход в Cloudflare (откроется браузер; если аккаунта нет — зарегистрируйте бесплатный)…');
    run(['login'], { soft: true });
    run(['whoami']);

    const rl = createInterface({ input: process.stdin, output: process.stdout });
    let gas = (toml.match(/GAS_URL = "([^"]*)"/) || [])[1];
    if (!gas || gas.startsWith('REPLACE')) {
      gas = (await rl.question('\nАдрес веб-приложения Apps Script (…/exec, версия KZ_Web 1.8, доступ «Все»): ')).trim().replace(/\/dev$/, '/exec').replace(/\?.*$/, '');
      if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(gas)) { console.error('Это не ссылка /exec'); process.exit(1); }
      toml = toml.replace(/GAS_URL = "[^"]*"/, `GAS_URL = "${gas}"`);
    }
    rl.close();

    if (/REPLACE_WITH_D1_ID/.test(toml)) {
      log('3/5  База D1 (kz-db)…');
      let out = runOut(['d1', 'create', 'kz-db']);
      let id = (out.match(/database_id\s*=\s*"([0-9a-f-]{36})"/) || out.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/) || [])[1];
      if (!id) { out = runOut(['d1', 'list', '--json']); try { id = (JSON.parse(out.slice(out.indexOf('['))).find(d => d.name === 'kz-db') || {}).uuid; } catch (e) {} }
      if (!id) { console.error('Не удалось определить ID базы — создайте kz-db в панели Cloudflare и впишите database_id в wrangler.toml'); process.exit(1); }
      toml = toml.replace('REPLACE_WITH_D1_ID', id);
    }
    writeFileSync(join(here, 'wrangler.toml'), toml);
  } else log('2/5–3/5  Обновление: вход и база уже настроены.');

  log('4/5  Публикация Worker…');
  let deployOut = '', ok = false;
  for (let i = 1; i <= 6 && !ok; i++) { deployOut = runOut(['deploy']); ok = /workers\.dev|Deployed|Uploaded/.test(deployOut) && !/ERROR|Error:/.test(deployOut); if (!ok) { console.log(`Повтор ${i}/6 через 10 с…`); Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 10000); } }
  if (!ok) { console.error('Публикация не удалась — см. вывод выше'); process.exit(1); }
  const url = (deployOut.match(/https:\/\/kz-edge[\w.-]*\.workers\.dev/) || [])[0] || '(адрес см. в панели Cloudflare → Workers → kz-edge)';

  let key = (saved.match(/EDGE_KEY:\s*(\S+)/) || [])[1] || '';
  if (!key || !update) {
    log('5/5  Ключ EDGE_KEY…');
    key = key || randomBytes(24).toString('base64url');
    run(['secret', 'put', 'EDGE_KEY'], { input: key + '\n' });
  }
  writeFileSync(keyFile, `Ускоритель «Контроль неснижаемого запаса» — подключение\n\nАдрес ускорителя: ${url}\nEDGE_KEY: ${key}\nGAS_URL: ${(toml.match(/GAS_URL = "([^"]*)"/) || [])[1] || ''}\nD1_ID: ${(toml.match(/database_id = "([^"]*)"/) || [])[1] || ''}\n\n1. Откройте программу как WIZARD → вкладка «Пароль» → блок «⚡ Ускоритель Cloudflare»: вставьте адрес и ключ → «Подключить».\n2. Сотрудникам дайте адрес ускорителя как ссылку на программу (ПРОСМОТР без пароля). В настольной версии и на сайте GitHub его можно указать как «Адрес программы Google».\n3. Обновление после замены файлов: update.cmd. Этот файл хранит ключ и адреса — не удаляйте его и не выкладывайте в репозиторий.\n\nВыпущено: ${new Date().toLocaleString('ru-RU')}\n`, 'utf8');
  log(`Готово. Адрес ускорителя: ${url}\nКлюч и инструкция — в файле CONNECT.txt (не выкладывайте его в репозиторий: он в .gitignore).`);
})().catch(e => { console.error('\nОшибка: ' + e.message); process.exit(1); });
