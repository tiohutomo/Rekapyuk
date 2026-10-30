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
  if (p.rows) {
    try { return json_(tambah_(JSON.parse(p.rows))); }
    catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
  }
  var sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET);
  return json_(sh ? { ok: true, versi: 4, sheet: SpreadsheetApp.getActive().getName() } : { ok: false, error: 'Tab "' + CONFIG.SHEET + '" tidak ditemukan.' });
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
