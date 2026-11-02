// Convertit en PDF les cartes et la synthèse produites par bilan_mensuel.py
// Usage : node outils/html_vers_pdf.js <dossier_sortie>
const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

(async () => {
  const dir = path.resolve(process.argv[2] || "bilan");
  const files = fs.readdirSync(dir).filter(f => /^(cartes|synthese)_.*\.html$/.test(f));
  const browser = await chromium.launch();
  const page = await browser.newPage();
  for (const f of files) {
    await page.goto("file://" + path.join(dir, f), { waitUntil: "networkidle" });
    await page.evaluate(() => document.fonts.ready);
    const out = path.join(dir, f.replace(/\.html$/, ".pdf"));
    await page.pdf({ path: out, preferCSSPageSize: true, printBackground: true });
    console.log("PDF :", out);
  }
  await browser.close();
})();
