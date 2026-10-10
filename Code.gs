/**
 * JEMBATAN APLIKASI <-> GOOGLE SPREADSHEET (KPU Kab. Pacitan)
 * Cara pasang: lihat petunjuk di chat. Ganti TOKEN di bawah dengan kode rahasia buatan Anda sendiri.
 */
const TOKEN = 'GANTI-DENGAN-KODE-RAHASIA-ANDA';
const SHEET = 'PERUBAHAN DATA';
const IDS = {
  '2026-IV':  '1CKcw637pVvZlRQwzOPzx-88FUTqyTXq9Zm72v2dlS0s',
  '2027-I':   '1DAU6e-maYxQgdjh6QNgTfUkOV25vwvdQXVgGqpU_6_g',
  '2027-II':  '1B_Y74wmRgYrok5CizH2IUslsQcws6jC1y9SdpaLUErc',
  '2027-III': '1u2Fx6r1Xx78_mmlaSxXU7F5I3_wl5AW2h8QcylAdRis',
  '2027-IV':  '16dnqD8NOetU45sLL4L1wPAhiBztfek6RUgPBUWsikfY'
};
// Kolom aplikasi 0..19 -> kolom sheet A..P (0-15), lewati Q, lalu R..U (16-19)
const col_ = c => c < 16 ? c + 1 : c + 2;

function sh_(q) {
  const id = IDS[q]; if (!id) throw new Error('Triwulan tidak dikenal: ' + q);
  const s = SpreadsheetApp.openById(id).getSheetByName(SHEET);
  if (!s) throw new Error('Sheet "' + SHEET + '" tidak ditemukan');
  return s;
}
function read_(s) {              // baca A:U mulai baris 2, buang baris kosong paling bawah
  const mr = s.getMaxRows(); if (mr < 2) return [];
  const v = s.getRange(2, 1, mr - 1, 21).getDisplayValues();
  let n = v.length;
  while (n > 0 && v[n - 1].every((x, i) => i === 16 || x === '')) n--;
  return v.slice(0, n).map(r => r.slice(0, 16).concat(r.slice(17, 21)));
}
const h_ = rows => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(rows)));
function pad_(r, a, b) { const x = r.slice(a, b); while (x.length < b - a) x.push(''); return x; }
function write_(s, rows, old) {  // hanya menyentuh A:P dan R:U (kolom Q dan V dst. tidak diubah)
  const m = Math.max(rows.length, old, 1), need = m + 1 - s.getMaxRows();
  if (need > 0) s.insertRowsAfter(s.getMaxRows(), need);
  const va = [], vb = [];
  for (let i = 0; i < m; i++) { const r = rows[i] || []; va.push(pad_(r, 0, 16)); vb.push(pad_(r, 16, 20)); }
  s.getRange(2, 1, m, 16).setNumberFormat('@').setValues(va);
  s.getRange(2, 18, m, 4).setNumberFormat('@').setValues(vb);
}
function doGet() { return ContentService.createTextOutput('Web App aktif'); }
function doPost(e) {
  let out;
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.t !== TOKEN) throw new Error('Kode rahasia salah');
    const s = sh_(b.q || '');
    if (b.a === 'get') { const r = read_(s); out = { ok: true, rows: r, h: h_(r) }; }
    else if (b.a === 'ver') { out = { ok: true, h: h_(read_(s)) }; }
    else if (b.a === 'write') {
      const L = LockService.getScriptLock(); L.waitLock(20000);
      try {
        const cur = read_(s);
        if (b.base && b.base !== h_(cur)) { out = { ok: false, conflict: true }; }
        else {
          b.ops.forEach(o => {
            if (o.op === 'cell') { const c = s.getRange(o.r + 2, col_(o.c)); c.setNumberFormat('@'); c.setValue(o.v); }
            else if (o.op === 'all') write_(s, o.rows, cur.length);
          });
          SpreadsheetApp.flush();
          out = { ok: true, h: h_(read_(s)) };
        }
      } finally { L.releaseLock(); }
    } else throw new Error('Aksi tidak dikenal');
  } catch (err) { out = { ok: false, err: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
