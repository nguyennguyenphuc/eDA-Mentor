// Chụp ảnh dashboard cho slide: cd /opt/node-tools (nơi có playwright) && node <đường dẫn>/chup_dashboard.js <thư mục assets> <file html>
// Mỗi trang (tab) một ảnh: dash_sN.png — dùng ở slide Phần 7 và để so với dashboard Excel (xuat_pdf_kiem_tra.py).
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
  }
  console.log('ERRORS', JSON.stringify(errs));
  await b.close();
})();
