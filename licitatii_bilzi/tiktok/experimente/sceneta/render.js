// Randează cadrele scenetei (JPEG, opace). Utilizare: node render.js frames.json out_dir [start] [pas]
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
(async () => {
  const [framesFile, outDir, start = '0', step = '1'] = process.argv.slice(2);
  const frames = JSON.parse(fs.readFileSync(framesFile, 'utf8'));
  fs.mkdirSync(outDir, { recursive: true });
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1080, height: 1920 } });
  await page.goto('file://' + path.join(__dirname, 'scene.html'));
  await page.evaluate(() => document.fonts.ready);
  for (let i = +start; i < frames.length; i += +step) {
    await page.evaluate((s) => window.renderFrame(s), frames[i]);
    await page.screenshot({ path: path.join(outDir, `f_${String(i).padStart(5, '0')}.jpg`), type: 'jpeg', quality: 92 });
  }
  await browser.close();
})();
