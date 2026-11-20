// Tempel di Google Sheet > Extensions > Apps Script. Ganti PIN, lalu Deploy > Web app (Anyone).
const PIN = '1234';
const DEFAULT_KAT = [['keluar','Makan'],['keluar','Transport'],['keluar','Listrik & Air'],['keluar','Belanja'],['keluar','Lainnya'],['masuk','Gaji'],['masuk','Bonus'],['masuk','Lainnya']];
const HEAD = {Transaksi:['id','tanggal','tipe','kategori','jumlah','catatan','pencatat'], Kategori:['tipe','nama'], Anggaran:['kategori','batas']};

function sh(n) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n); s.appendRow(HEAD[n]);
    if (n === 'Kategori') DEFAULT_KAT.forEach(r => s.appendRow(r));
  } else if (n === 'Transaksi' && !s.getRange(1,7).getValue()) s.getRange(1,7).setValue('pencatat');
  return s;
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function rows(n) { return sh(n).getDataRange().getValues().slice(1); }

function doGet(e) {
  if (e.parameter.pin !== PIN) return out({ok:false, error:'PIN salah'});
  const tz = Session.getScriptTimeZone();
  const trx = rows('Transaksi').map(r => ({
    id: r[0], tanggal: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
    tipe: r[2], kategori: r[3], jumlah: Number(r[4]), catatan: r[5], pencatat: r[6] || ''
  }));
  const kat = rows('Kategori').map(r => ({tipe: r[0], nama: r[1]}));
  const bud = {}; rows('Anggaran').forEach(r => bud[r[0]] = Number(r[1]));
  return out({ok:true, trx, kat, bud});
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.pin !== PIN) return out({ok:false, error:'PIN salah'});
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (b.action === 'add') {
      sh('Transaksi').appendRow([Utilities.getUuid(), "'" + b.tanggal, b.tipe, b.kategori, b.jumlah, b.catatan || '', b.pencatat || '']);
    } else if (b.action === 'edit' || b.action === 'del') {
      const s = sh('Transaksi'), ids = s.getRange(1,1,s.getLastRow(),1).getValues().flat(), i = ids.indexOf(b.id);
      if (i > 0) {
        if (b.action === 'del') s.deleteRow(i + 1);
        else {
          s.getRange(i+1, 2).setNumberFormat('@').setValue(b.tanggal);
          s.getRange(i+1, 3, 1, 5).setValues([[b.tipe, b.kategori, b.jumlah, b.catatan || '', b.pencatat || '']]);
        }
      }
    } else if (b.action === 'addKat') {
      sh('Kategori').appendRow([b.tipe, b.nama]);
    } else if (b.action === 'delKat') {
      const s = sh('Kategori'), v = s.getDataRange().getValues();
      for (let i = v.length - 1; i > 0; i--) if (v[i][0] === b.tipe && v[i][1] === b.nama) { s.deleteRow(i + 1); break; }
    } else if (b.action === 'setBud') {
      const s = sh('Anggaran'), v = s.getDataRange().getValues();
      let i = 0; for (let j = 1; j < v.length; j++) if (v[j][0] === b.kategori) { i = j + 1; break; }
      if (!b.batas) { if (i) s.deleteRow(i); }
      else if (i) s.getRange(i, 2).setValue(b.batas);
      else s.appendRow([b.kategori, b.batas]);
    }
    return out({ok:true});
  } finally { lock.releaseLock(); }
}
