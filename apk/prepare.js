// Menyiapkan folder www/ untuk dibungkus jadi APK.
// Mengambil ../index.html, mengganti sumber pustaka dari CDN ke file lokal,
// lalu menyalin pustaka (Excel, PDF, OCR) dari node_modules.
const fs = require("fs");
const path = require("path");

const root = __dirname;
const www = path.join(root, "www");
const nm = path.join(root, "node_modules");

function must(p) {
  if (!fs.existsSync(p)) throw new Error("File tidak ditemukan: " + p);
  return p;
}
function copy(src, dst) {
  fs.mkdirSync(path.dirname(dst), { recursive: true });
  fs.copyFileSync(must(src), dst);
}
function replaceOnce(s, from, to) {
  const n = s.split(from).length - 1;
  if (n !== 1) throw new Error("Harus ada tepat 1 kecocokan, ditemukan " + n + ": " + from.slice(0, 70));
  return s.replace(from, () => to);
}

fs.rmSync(www, { recursive: true, force: true });
fs.mkdirSync(www, { recursive: true });

let html = fs.readFileSync(must(path.join(root, "..", "index.html")), "utf8");
html = replaceOnce(
  html,
  '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>',
  '<script src="libs/xlsx.full.min.js"></script>'
);
html = replaceOnce(
  html,
  '<meta charset="utf-8">',
  '<meta charset="utf-8">\n<script>window.RK_APK=1;window.RK_PDFJS="libs/pdfjs/";window.RK_TESS="libs/tesseract/";</script>'
);
fs.writeFileSync(path.join(www, "index.html"), html);

copy(path.join(nm, "xlsx/dist/xlsx.full.min.js"), path.join(www, "libs/xlsx.full.min.js"));
copy(path.join(nm, "pdfjs-dist/build/pdf.min.js"), path.join(www, "libs/pdfjs/pdf.min.js"));
copy(path.join(nm, "pdfjs-dist/build/pdf.worker.min.js"), path.join(www, "libs/pdfjs/pdf.worker.min.js"));
copy(path.join(nm, "tesseract.js/dist/tesseract.min.js"), path.join(www, "libs/tesseract/tesseract.min.js"));
copy(path.join(nm, "tesseract.js/dist/worker.min.js"), path.join(www, "libs/tesseract/worker.min.js"));

const coreDir = must(path.join(nm, "tesseract.js-core"));
const cores = fs.readdirSync(coreDir).filter((f) => /^tesseract-core.*\.(wasm\.js|wasm)$/.test(f));
if (!cores.length) throw new Error("Berkas tesseract-core tidak ditemukan di " + coreDir);
cores.forEach((f) => copy(path.join(coreDir, f), path.join(www, "libs/tesseract", f)));

const dataRoot = must(path.join(nm, "@tesseract.js-data/eng"));
const cands = ["4.0.0_best_int", "4.0.0", "4.0.0_fast"].map((d) => path.join(dataRoot, d, "eng.traineddata.gz"));
const lang = cands.find((p) => fs.existsSync(p));
if (!lang) throw new Error("eng.traineddata.gz tidak ditemukan di " + dataRoot);
copy(lang, path.join(www, "libs/tesseract/lang/eng.traineddata.gz"));

console.log("www siap. Berkas inti OCR:", cores.join(", "));
