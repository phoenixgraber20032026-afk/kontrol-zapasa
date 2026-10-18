// Сборка папки сайта app/ из исходников gas/: index.html = KZ_WebApp.html (единый исходник для Google и сайта).
import { copyFileSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
const html = readFileSync('gas/KZ_WebApp.html', 'utf8');
writeFileSync('app/index.html', html);
if (!existsSync('app/chart.umd.min.js')) console.warn('Нет app/chart.umd.min.js — скачайте https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js в app/');
// версия в sw.js = версия из KZ_Web.gs, чтобы кэш обновлялся при каждом выпуске
const ver = (readFileSync('gas/KZ_Web.gs', 'utf8').match(/VERSION:\s*'([^']+)'/) || [])[1] || Date.now();
writeFileSync('app/sw.js', readFileSync('app/sw.js', 'utf8').replace(/var CACHE = '[^']+'/, `var CACHE = 'kz-app-v${ver}'`));
// версия настольной программы
const pkgPath = 'desktop/package.json';
if (existsSync(pkgPath)) { const pkg = JSON.parse(readFileSync(pkgPath, 'utf8')); if (/^\d+\.\d+$/.test(String(ver))) pkg.version = ver + '.0'; writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + '\n'); }
console.log('app/ собрана, версия', ver);
