import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';

/**
 * 用 Chrome + CDP 把 HTML 打成 A4 PDF（带页码页脚）。
 *
 * 用法：
 *   node print_pdf.mjs <srcDir> <out.pdf> [footerLeft] [--port 9411]
 *
 * srcDir 约定：
 *   - 若存在 style.css，则把该目录下所有 *.html 按文件名排序拼接，
 *     并内联 style.css 生成 _book.html（分章片段工作流）；
 *   - 否则优先用 index.html / book.html 单文件。
 * 图片用相对路径（如 img/xxx.jpg），最终以 file:// 方式加载。
 */

const args = process.argv.slice(2);
const positional = args.filter(a => !a.startsWith('--'));
const flag = k => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : undefined; };

const SRCDIR = path.resolve(positional[0] || '.');
const PDFOUT = path.resolve(positional[1] || 'out.pdf');
const FOOTER_LEFT = positional[2] || '复习讲义';
const PORT = Number(flag('--port') || 9411);

const CHROME_CANDIDATES = [
  process.env.CHROME_PATH,
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  '/usr/bin/google-chrome',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
].filter(Boolean);
const CHROME = CHROME_CANDIDATES.find(p => { try { return fs.existsSync(p); } catch { return false; } });
if (!CHROME) { console.error('找不到 Chrome/Edge，请设置 CHROME_PATH'); process.exit(1); }

/* ---------- 1. 组装 HTML ---------- */
const cssPath = path.join(SRCDIR, 'style.css');
let html;
const OUTHTML = path.join(SRCDIR, '_book.html');

if (fs.existsSync(cssPath)) {
  const css = fs.readFileSync(cssPath, 'utf8');
  const files = fs.readdirSync(SRCDIR)
    .filter(f => /^\d\d_.*\.html$/.test(f))
    .sort();
  let body = '';
  for (const f of files) body += `\n<!-- ===== ${f} ===== -->\n` + fs.readFileSync(path.join(SRCDIR, f), 'utf8');
  html = `<!DOCTYPE html>\n<html lang="zh-CN">\n<head><meta charset="utf-8">\n<style>${css}</style>\n</head>\n<body>${body}\n</body></html>`;
  console.log(`assembled ${files.length} fragments`);
} else {
  const single = ['index.html', 'book.html'].map(f => path.join(SRCDIR, f)).find(p => fs.existsSync(p));
  if (!single) { console.error('srcDir 内既无 style.css 也无 index.html'); process.exit(1); }
  html = fs.readFileSync(single, 'utf8');
}
fs.writeFileSync(OUTHTML, html, 'utf8');
console.log(`html -> ${OUTHTML} (${(html.length / 1024).toFixed(0)} KB)`);

/* ---------- 2. 启动 Chrome ---------- */
const udd = path.join(SRCDIR, '.chromeprofile');
fs.rmSync(udd, { recursive: true, force: true });
const chrome = spawn(CHROME, [
  '--headless=new', '--disable-gpu', '--no-sandbox', '--no-first-run',
  '--disable-extensions', '--hide-scrollbars', '--force-device-scale-factor=1',
  '--font-render-hinting=none', '--allow-file-access-from-files',
  `--user-data-dir=${udd}`, `--remote-debugging-port=${PORT}`, 'about:blank',
], { stdio: 'ignore' });

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function getWs() {
  for (let i = 0; i < 60; i++) {
    try {
      const r = await fetch(`http://127.0.0.1:${PORT}/json/version`);
      const j = await r.json();
      if (j.webSocketDebuggerUrl) return j.webSocketDebuggerUrl;
    } catch { }
    await sleep(400);
  }
  throw new Error('chrome devtools 未就绪');
}

/* ---------- 3. 最小 CDP 客户端 ---------- */
const ws = new WebSocket(await getWs());
await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });

let msgId = 0;
const pending = new Map();
const listeners = [];
ws.onmessage = ev => {
  const m = JSON.parse(ev.data);
  if (m.id && pending.has(m.id)) {
    const { res, rej } = pending.get(m.id);
    pending.delete(m.id);
    m.error ? rej(new Error(JSON.stringify(m.error))) : res(m.result);
  } else if (m.method) for (const l of listeners) l(m);
};
function send(method, params = {}, sessionId) {
  const id = ++msgId;
  const payload = { id, method, params };
  if (sessionId) payload.sessionId = sessionId;
  ws.send(JSON.stringify(payload));
  return new Promise((res, rej) => pending.set(id, { res, rej }));
}

const { targetId } = await send('Target.createTarget', { url: 'about:blank' });
const { sessionId } = await send('Target.attachToTarget', { targetId, flatten: true });
await send('Page.enable', {}, sessionId);
await send('Emulation.setDefaultBackgroundColorOverride',
  { color: { r: 255, g: 255, b: 255, a: 1 } }, sessionId);

const loaded = new Promise(res => {
  const h = m => { if (m.method === 'Page.loadEventFired') { listeners.splice(listeners.indexOf(h), 1); res(); } };
  listeners.push(h);
});

const fileUrl = 'file:///' + OUTHTML.replace(/\\/g, '/').replace(/ /g, '%20').replace(/#/g, '%23');
await send('Page.navigate', { url: fileUrl }, sessionId);
await loaded;
await sleep(2500);   // 等图片解码与字体

/* ---------- 4. 打印（页脚页码只能靠 footerTemplate） ---------- */
const footer = `<div style="width:100%;margin:0 14mm;font-family:'Microsoft YaHei',sans-serif;font-size:7.6pt;color:#8b98a8;border-top:.4pt solid #dfe6ee;padding-top:2mm;">
  <table style="width:100%;border:0;border-collapse:collapse;"><tr>
    <td style="text-align:left;border:0;padding:0;">${FOOTER_LEFT.replace(/</g, '&lt;')}</td>
    <td style="text-align:right;border:0;padding:0;">第 <span class="pageNumber"></span> 页 / 共 <span class="totalPages"></span> 页</td>
  </tr></table></div>`;

const pdf = await send('Page.printToPDF', {
  printBackground: true,
  paperWidth: 8.27, paperHeight: 11.69,     // A4
  marginTop: 0.52, marginBottom: 0.58, marginLeft: 0.55, marginRight: 0.55,
  displayHeaderFooter: true,
  headerTemplate: '<div></div>',
  footerTemplate: footer,
  preferCSSPageSize: false,
  scale: 1,
}, sessionId);

fs.writeFileSync(PDFOUT, Buffer.from(pdf.data, 'base64'));
console.log(`PDF -> ${PDFOUT}  ${(fs.statSync(PDFOUT).size / 1024 / 1024).toFixed(2)} MB`);

try { await send('Browser.close'); } catch { }
chrome.kill();
await sleep(500);
process.exit(0);
