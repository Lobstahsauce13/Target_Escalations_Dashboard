// Drives the Target Escalation Dashboard in headless Chromium. Serves the repo itself, reads one command per line on stdin.
// Run from the repo root:  node .claude/skills/run-target-escalations-dashboard/driver.mjs <<'CMDS' ... CMDS
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import readline from 'node:readline';
import { execSync } from 'node:child_process';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const { chromium } = require(path.join(execSync('npm root -g').toString().trim(), 'playwright'));

const ROOT = process.cwd();
if (!fs.existsSync(path.join(ROOT, 'index.html'))) throw new Error('Run from the repo root (no index.html in ' + ROOT + ')');
const OUT = process.env.SHOTS || '/tmp/target-shots';
fs.mkdirSync(OUT, { recursive: true });
const SLOT = { export: '#file', data: '#vfile', presence: '#pfile', tracker: '#tfile', parts: '#xfile' };
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml' };

const server = http.createServer((req, res) => {
  const p = path.join(ROOT, decodeURIComponent(new URL(req.url, 'http://x').pathname));
  if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': TYPES[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
});
await new Promise(r => server.listen(0, '127.0.0.1', r));
const URL0 = `http://127.0.0.1:${server.address().port}/index.html`;

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, timezoneId: 'America/New_York', acceptDownloads: true });
const page = await ctx.newPage();
const errs = [];
page.on('pageerror', e => errs.push('pageerror: ' + e.message));
page.on('console', m => m.type() === 'error' && errs.push('console: ' + m.text()));
// index.html pulls SheetJS from cdnjs; serve the repo's copy so runs don't depend on the network.
await page.route('**/cdnjs.cloudflare.com/**/xlsx.full.min.js', r => r.fulfill({ path: path.join(ROOT, 'xlsx.full.min.js'), contentType: 'text/javascript' }));
const vis = () => page.frame({ url: /visits\.html/ });
const toast = async () => {
  const t = await page.waitForSelector('#toast.on', { timeout: 5000 }).then(() => page.textContent('#toast')).catch(() => '');
  await page.evaluate(() => document.getElementById('toast')?.classList.remove('on'));
  return t.trim();
};

const CMD = {
  async open() {
    await page.goto(URL0);
    await page.waitForFunction(() => window.XLSX && document.getElementById('visView')?.contentWindow.loadVisits);
    return 'opened ' + URL0;
  },
  async theme(t) { await page.emulateMedia({ colorScheme: t }); return 'theme ' + t; },
  async load(slot, file) {
    if (!SLOT[slot]) throw new Error('slot must be one of ' + Object.keys(SLOT).join(', '));
    await page.setInputFiles(SLOT[slot], path.resolve(file));
    return 'toast: ' + (await toast() || '(none)') + ' | ' + (await page.textContent('#fpill')).trim();
  },
  async loaded(n) { await page.waitForFunction(n => document.getElementById('fpill').textContent.includes(n + ' of 5'), n, { timeout: 120000 }); return (await page.textContent('#fpill')).trim(); },
  async closeimports() { if (await page.isVisible('#imclose')) await page.click('#imclose'); return 'closed'; },
  async tab(app) { await page.click(`#apptabs [data-app="${app}"]`); await page.waitForTimeout(300); return 'tab ' + app; },
  async search(q) { await page.fill('#s', q); await page.dispatchEvent('#s', 'input'); await page.waitForTimeout(400); return (await page.textContent('#cnt')).trim(); },
  async click(sel) { await page.click(sel); await page.waitForTimeout(300); return 'clicked ' + sel; },
  async fill(sel, ...v) { await page.fill(sel, v.join(' ')); return 'filled ' + sel; },
  async text(sel) { return (await page.innerText(sel)).replace(/\s+/g, ' ').trim().slice(0, 2000); },
  async eval(...js) { return JSON.stringify(await page.evaluate(js.join(' '))); },
  async veval(...js) { return JSON.stringify(await vis().evaluate(js.join(' '))); },
  async toast() { return await toast() || '(no toast)'; },
  async shot(name = 'shot') { const f = path.join(OUT, name + '.png'); await page.screenshot({ path: f }); return f; },
  async wait(ms) { await page.waitForTimeout(+ms); return 'waited ' + ms; },
  async errors() { return errs.length ? errs.join('\n') : 'no errors'; },
};

let failed = 0;
for await (const line of readline.createInterface({ input: process.stdin })) {
  const l = line.trim();
  if (!l || l.startsWith('#')) continue;
  const [c, ...a] = l.split(/\s+/);
  try {
    if (!CMD[c]) throw new Error('unknown command; known: ' + Object.keys(CMD).join(' '));
    console.log(`> ${l}\n${await CMD[c](...a)}`);
  } catch (e) { failed++; console.log(`> ${l}\nERROR ${e.message.split('\n')[0]}`); }
}
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
