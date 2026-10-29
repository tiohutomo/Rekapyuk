// Uji asap: membuka www/ di Chromium tanpa internet-CDN, lalu memeriksa pustaka lokal.
const { chromium } = require("playwright");

const URL = process.env.SMOKE_URL || "http://localhost:8080/index.html";
function fail(msg) {
  console.error("::error::" + msg);
  console.error(msg);
  process.exit(1);
}

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errs = [];
  const external = [];
  page.on("pageerror", (e) => errs.push(String(e)));
  page.on("request", (r) => {
    const u = r.url();
    if (!u.startsWith("http://localhost") && !u.startsWith("data:") && !u.startsWith("blob:")) external.push(u);
  });
  await page.goto(URL);
  await page.waitForTimeout(1500);

  const xl = await page.evaluate(() => typeof XLSX);
  if (xl !== "object") fail("Pustaka Excel lokal tidak termuat (typeof XLSX = " + xl + ")");

  const pdf = await page.evaluate(async () => {
    try { const l = await loadPdf(); return typeof l.getDocument + "|" + l.GlobalWorkerOptions.workerSrc; } catch (e) { return "ERR " + JSON.stringify(e); }
  });
  if (!String(pdf).startsWith("function|") || String(pdf).includes("http")) fail("Pembaca PDF lokal bermasalah: " + pdf);

  const ocr = await page.evaluate(async () => {
    try {
      const c = document.createElement("canvas");
      c.width = 900; c.height = 260;
      const g = c.getContext("2d");
      g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height);
      g.fillStyle = "#111"; g.font = "bold 40px Arial";
      g.fillText("25 Okt 2026", 30, 60);
      g.fillText("Transfer Rupiah  - Rp 5.500.000", 30, 130);
      g.fillText("Ke MUAMALAT", 30, 190);
      const blob = await new Promise((r) => c.toBlob(r, "image/png"));
      const s = await claude.use("sample");
      return await s.json("x", { images: [blob] });
    } catch (e) { return { error: String((e && (e.message || e.code)) || e) }; }
  });
  if (!Array.isArray(ocr) || !ocr.length) fail("OCR lokal gagal: " + JSON.stringify(ocr));
  const r = ocr[0];
  if (r.n !== 5500000 || r.t !== "2026-10-25") fail("Hasil OCR tidak sesuai: " + JSON.stringify(ocr));

  if (external.length) fail("Masih memanggil alamat luar: " + external.slice(0, 3).join(", "));
  if (errs.length) fail("Error halaman: " + errs.join(" | "));
  console.log("Uji asap lolos:", JSON.stringify(r));
  await browser.close();
})().catch((e) => fail("Uji asap error: " + (e && e.message || e)));
