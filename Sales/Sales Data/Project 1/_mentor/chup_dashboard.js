// Chụp ảnh dashboard cho slide: cd /opt/node-tools (nơi có playwright) && node <đường dẫn>/chup_dashboard.js <thư mục assets> <file html>
// Mỗi trang: dash_sN.png (toàn trang), dash_sN_figK.png (từng biểu đồ/bảng), dash_sN_kpiK.png (từng thẻ KPI).
const { chromium } = require('playwright');
(async () => {
  const OUT = process.argv[2], HTML = process.argv[3];
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1400, height: 900 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  await p.goto('file://' + HTML);
  await p.waitForTimeout(1200);
  for (const s of ['s1','s2','s3','s4']) {
    await p.click(`button[data-p=${s}]`);
    await p.waitForTimeout(600);
    await (await p.$('.wrap')).screenshot({ path: `${OUT}/dash_${s}.png` });
    const figs = await p.$$(`#${s} .card`);
    for (let i = 0; i < figs.length; i++) await figs[i].screenshot({ path: `${OUT}/dash_${s}_fig${i+1}.png` });
    const ks = await p.$$(`#${s} .k`);
    for (let i = 0; i < ks.length; i++) await ks[i].screenshot({ path: `${OUT}/dash_${s}_kpi${i+1}.png` });
  }
  console.log('ERRORS', JSON.stringify(errs));
  await b.close();
})();
