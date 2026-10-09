// Randează cadrele stratului cu personajul (PNG transparent 1080x820).
// Utilizare: node render.js frames.json out_dir
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const [framesFile, outDir] = process.argv.slice(2);
  const frames = JSON.parse(fs.readFileSync(framesFile, 'utf8'));
  fs.mkdirSync(outDir, { recursive: true });
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1080, height: 1020 } });
  await page.goto('file://' + path.join(__dirname, 'avatar.html'));
  await page.evaluate(() => document.fonts.ready);
  for (let i = 0; i < frames.length; i++) {
    await page.evaluate((s) => window.renderFrame(s), frames[i]);
    await page.screenshot({ path: path.join(outDir, `f_${String(i).padStart(5, '0')}.png`), omitBackground: true });
  }
  await browser.close();
})();
