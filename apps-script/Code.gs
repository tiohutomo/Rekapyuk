// rekapyuk -> Google Sheets (Merkuri). Tempel di Extensions > Apps Script pada spreadsheet-mu.
// Ganti KODE_RAHASIA dengan kata bebas, lalu isi kata yang sama di aplikasi (Setup > Google Sheets).
var KODE_RAHASIA = 'ganti-kode-ini';
var CONFIG = { SHEET: 'Transaksi', HEADER_ROW: 7, COL_TANGGAL: 4, COL_DESKRIPSI: 51 };

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.token !== KODE_RAHASIA) return json_({ ok: false, error: 'Kode rahasia salah.' });
  if (p.aksi === 'saldo') {
    try { return json_(saldoRekening_()); }
    catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
  }
  if (p.aksi === 'periksa') {
    try { return json_(periksa_(JSON.parse(p.names || '[]'), p.tgl)); }
    catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
  }
  if (p.rows) {
    try { return json_(tambah_(JSON.parse(p.rows))); }
    catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
  }
  var sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET);
  return json_(sh ? { ok: true, versi: 6, sheet: SpreadsheetApp.getActive().getName() } : { ok: false, error: 'Tab "' + CONFIG.SHEET + '" tidak ditemukan.' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.token !== KODE_RAHASIA) return json_({ ok: false, error: 'Kode rahasia salah.' });
    return json_(tambah_(body.rows || []));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function tambah_(rows) {
  if (!rows.length) return { ok: false, error: 'Tidak ada transaksi.' };
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(CONFIG.SHEET);
  if (!sh) return { ok: false, error: 'Tab "' + CONFIG.SHEET + '" tidak ditemukan.' };
  var tz = ss.getSpreadsheetTimeZone();

  var from = CONFIG.HEADER_ROW + 1;
  var count = sh.getMaxRows() - CONFIG.HEADER_ROW;
  var vals = sh.getRange(from, CONFIG.COL_TANGGAL, count, 5).getValues();
  var last = 0, have = {};
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (v[0] === '' || v[3] === '') continue; // baris dianggap terisi kalau Tanggal dan Jumlah ada; baris setengah jadi akan ditimpa
    last = i + 1;
    if (v[0] instanceof Date) {
      var k = [Utilities.formatDate(v[0], tz, 'yyyy-MM-dd'), String(v[1]).trim(), Number(v[3]), String(v[4]).trim()].join('|');
      have[k] = (have[k] || 0) + 1;
    }
  }

  var send = [], skipped = 0;
  rows.forEach(function (r) {
    var k = [r.t, r.j, Number(r.n), r.r].join('|');
    if (have[k]) { have[k]--; skipped++; } else send.push(r);
  });

  var start = from + last;
  if (send.length) {
    if (start + send.length - 1 > sh.getMaxRows()) return { ok: false, error: 'Baris di tab Transaksi tidak cukup.' };
    var ab = [], cc = [], desc = [];
    send.forEach(function (r) {
      var p = r.t.split('-');
      ab.push([new Date(+p[0], +p[1] - 1, +p[2]), r.j]);
      cc.push([r.k || '', Number(r.n), r.r]);
      desc.push([r.d || '']);
    });
    try {
      // Dropdown Kategori bergantung pada kolom Transaksi, jadi tulis dua tahap:
      // 1) Tanggal + Transaksi, 2) Kategori + Jumlah + Rekening (setelah daftar kategori baris itu terbentuk).
      sh.getRange(start, CONFIG.COL_TANGGAL, send.length, 2).setValues(ab);
      SpreadsheetApp.flush();
      var tulis2 = function () {
        sh.getRange(start, CONFIG.COL_TANGGAL + 2, send.length, 3).setValues(cc);
        SpreadsheetApp.flush();
      };
      try { tulis2(); } catch (e1) { Utilities.sleep(1500); tulis2(); }
      sh.getRange(start, CONFIG.COL_DESKRIPSI, send.length, 1).setValues(desc);
      SpreadsheetApp.flush();
    } catch (err) {
      return { ok: false, error: 'Gagal menulis mulai baris ' + start + ': ' + String(err && err.message || err) };
    }
  }
  return { ok: true, sent: send.length, skipped: skipped, firstRow: send.length ? start : 0, lastRow: send.length ? start + send.length - 1 : 0 };
}

// Mencari baris judul tabel REKENING (kolom REKENING dan SALDO SAAT INI) di tab Rekening.
function cariTabelRekening_(vals) {
  for (var i = 0; i < vals.length; i++) {
    for (var c = 0; c < vals[i].length - 1; c++) {
      if (String(vals[i][c]).trim().toUpperCase() !== 'REKENING') continue;
      for (var d = c + 1; d <= c + 3 && d < vals[i].length; d++) {
        if (String(vals[i][d]).trim().toUpperCase() === 'SALDO SAAT INI') return { hr: i, hc: c };
      }
    }
  }
  return null;
}

// Membaca daftar "SALDO SAAT INI" per rekening dari tab Rekening.
function saldoRekening_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Rekening');
  if (!sh) return { ok: false, error: 'Tab "Rekening" tidak ditemukan.' };
  var vals = sh.getDataRange().getValues(), t = cariTabelRekening_(vals), i, c;
  if (!t) return { ok: false, error: 'Tabel REKENING / SALDO SAAT INI tidak ditemukan di tab Rekening.' };
  var out = [];
  for (i = t.hr + 1; i < vals.length; i++) {
    var nama = String(vals[i][t.hc] || '').trim();
    if (!nama) break;
    var saldo = null;
    for (c = t.hc + 1; c <= t.hc + 4 && c < vals[i].length; c++) {
      if (typeof vals[i][c] === 'number') { saldo = vals[i][c]; break; }
    }
    out.push({ n: nama, saldo: saldo === null ? 0 : saldo });
  }
  return { ok: true, rekening: out };
}

// Menulis tanggal ke kolom "TERAKHIR DIPERIKSA" untuk rekening yang saldonya sama.
// Sel yang berisi catatan teks (bukan tanggal) dilewati supaya catatanmu tidak tertimpa.
function periksa_(names, tgl) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Rekening');
  if (!sh) return { ok: false, error: 'Tab "Rekening" tidak ditemukan.' };
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(tgl || '');
  if (!m) return { ok: false, error: 'Tanggal tidak valid.' };
  var tanggal = new Date(+m[1], +m[2] - 1, +m[3]);
  var vals = sh.getDataRange().getValues(), t = cariTabelRekening_(vals);
  if (!t) return { ok: false, error: 'Tabel REKENING tidak ditemukan di tab Rekening.' };
  var kol = -1, c;
  for (c = 0; c < vals[t.hr].length; c++) {
    if (String(vals[t.hr][c]).trim().toUpperCase() === 'TERAKHIR DIPERIKSA') { kol = c; break; }
  }
  if (kol < 0) return { ok: false, error: 'Kolom TERAKHIR DIPERIKSA tidak ditemukan.' };
  var want = {}; names.forEach(function (n) { want[String(n).trim().toLowerCase()] = 1; });
  var updated = 0, dilewati = [];
  for (var i = t.hr + 1; i < vals.length; i++) {
    var nama = String(vals[i][t.hc] || '').trim();
    if (!nama) break;
    if (!want[nama.toLowerCase()]) continue;
    var cur = vals[i][kol];
    if (cur !== '' && !(cur instanceof Date)) { dilewati.push(nama); continue; }
    sh.getRange(i + 1, kol + 1).setValue(tanggal);
    updated++;
  }
  SpreadsheetApp.flush();
  return { ok: true, updated: updated, dilewati: dilewati };
}
