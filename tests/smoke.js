// Preizkus strani: prilepi besedilo AJPES -> polja podjetja -> naloži izmišljeno bilanco -> Excel ima D3 do D5.
// Uporaba: strežnik z ALLOW_NO_AUTH=1 na BASE (privzeto http://localhost:8781), nato: node tests/smoke.js
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path'), os = require('os');
const BASE = process.env.BASE || 'http://localhost:8781';
const AJPES = fs.readFileSync(path.join(__dirname, 'fixture_ajpes.txt'), 'utf8');
const BIL = path.join(__dirname, 'fixture_cbk_lp2025.xlsx');
const fails = [];
const ok = (c, m) => { console.log((c ? '  ✓ ' : '  ✗ ') + m); if (!c) fails.push(m); };
const paste = (page, text) => page.evaluate(t => { const dt = new DataTransfer(); dt.setData('text/plain', t); document.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); }, text);

(async () => {
  const browser = await chromium.launch();
  const page = await (await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true })).newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errs.push('console: ' + m.text()); });

  console.log('1. Začetni zaslon');
  await page.goto(BASE); await page.waitForTimeout(800);
  ok(await page.locator('#coDz').isVisible(), 'polje za vnos podjetja je vidno');
  ok(await page.locator('#cbkDz').isVisible(), 'polje za bilanco je vidno');
  ok(!(await page.locator('.tb-tabs').isVisible()), 'ni zavihkov z obrazci');

  console.log('2. Podjetje iz AJPES');
  await paste(page, AJPES); await page.waitForTimeout(800);
  ok((await page.inputValue('#coNaziv')).includes('TESTNO PODJETJE'), 'naziv podjetja');
  ok((await page.inputValue('#coMat')) === '1234567000', 'matična številka');
  ok((await page.textContent('#coMsg')).includes('št. partnerja'), 'opozorilo, da št. partnerja manjka');
  await page.fill('#coPart', '123456');

  console.log('3. Pretvorba bilance');
  await page.setInputFiles('#cbkFile', BIL);
  await page.selectOption('#cbkTip', '8');
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 60000 }), page.click('#cbkGo')]);
  const out = path.join(os.tmpdir(), 'omni-cbk-test.xlsx'); await dl.saveAs(out);
  await page.waitForTimeout(500);
  ok(await page.locator('#cbkRes').isVisible(), 'prikazan je povzetek pretvorbe');
  const py = require('child_process').spawnSync('python3', ['-c', `import openpyxl,sys;w=openpyxl.load_workbook(sys.argv[1])['BS'];print(w['D3'].value+'|'+str(w['D4'].value)+'|'+str(w['D5'].value))`, out]).stdout.toString().trim();
  ok(py === 'TESTNO PODJETJE, trgovina in storitve, d.o.o.|1234567000|123456', 'Excel: D3, D4, D5 (' + py + ')');

  console.log('4. Napake v brskalniku');
  ok(errs.length === 0, 'brez napak v konzoli' + (errs.length ? ': ' + errs.join(' | ') : ''));
  await browser.close();
  if (fails.length) { console.log('\nNEUSPEH: ' + fails.length); process.exit(1); }
  console.log('\nVSE V REDU');
})();
