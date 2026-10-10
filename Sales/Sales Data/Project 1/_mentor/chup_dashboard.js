// Chụp ảnh dashboard cho slide: cd /opt/node-tools (nơi có playwright) && node <đường dẫn>/chup_dashboard.js <thư mục assets> <file html>
const { chromium } = require('playwright');
(async () => {
  const OUT = process.argv[2], HTML = process.argv[3];
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1400, height: 900 }, deviceScaleFactor: 2 });
  await p.goto('file://' + HTML);
  await p.waitForTimeout(1500);
  const info = await p.evaluate(() => ({month: document.querySelector('#aq-month').value, date: document.querySelector('#aq-date').textContent,
     opts: [...document.querySelectorAll('#aq-month option')].slice(-3).map(o=>o.value+':'+o.textContent)}));
  console.log(JSON.stringify(info));
  for (const s of ['s1','s2','s3','s4']) {
    await p.click(`button[data-p=${s}]`);
    await p.waitForTimeout(800);
    const el = await p.$('#aq-canvas');
    await el.screenshot({ path: `${OUT}/dash_${s}.png` });
    const figs = await p.$$(`#${s} figure.c`);
    for (let i = 0; i < figs.length; i++) await figs[i].screenshot({ path: `${OUT}/dash_${s}_fig${i+1}.png` });
    const ks = await p.$$(`#${s} .k`);
    for (let i = 0; i < ks.length; i++) await ks[i].screenshot({ path: `${OUT}/dash_${s}_kpi${i+1}.png` });
    const txt = await p.evaluate((s) => [...document.querySelectorAll(`#${s} .k, #${s} figcaption`)].map(e => e.innerText.replace(/\n/g,' | ')), s);
    console.log(s, JSON.stringify(txt, null, 0));
  }
  await b.close();
})();
