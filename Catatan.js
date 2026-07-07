/**
 * Catatan.js — Catatan Khusus Siswa
 * Guru dapat menulis catatan per siswa (perilaku, prestasi, dll).
 * Sheet CATATAN: [0]=id, [1]=email_guru, [2]=kelas, [3]=nis, [4]=nama, [5]=catatan, [6]=tgl
 */

var CATATAN_SHEET = 'CATATAN';

function _ensureCatatanSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(CATATAN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CATATAN_SHEET);
    sh.appendRow(['id', 'email_guru', 'kelas', 'nis', 'nama', 'catatan', 'tgl']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * getCatatanSiswa(kelas) — Ambil catatan siswa untuk kelas tertentu milik guru aktif
 */
function getCatatanSiswa(kelas) {
  assertLicenseActive();
  var email = getLoginEmail();
  var sh = _ensureCatatanSheet_();
  var data = sh.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[1] || '').toLowerCase() !== email) continue;
    if (kelas && String(row[2] || '') !== String(kelas)) continue;
    result.push({
      id      : String(row[0] || ''),
      kelas   : String(row[2] || ''),
      nis     : String(row[3] || ''),
      nama    : String(row[4] || ''),
      catatan : String(row[5] || ''),
      tgl     : String(row[6] || '')
    });
  }
  return result;
}

/**
 * saveCatatanSiswa(obj) — Simpan atau update catatan siswa
 * obj: {kelas, nis, nama, catatan}
 */
function saveCatatanSiswa(obj) {
  assertLicenseActive();
  var email = getLoginEmail();

  if (!obj || !obj.kelas) throw new Error('Kelas wajib diisi');
  if (!obj.nis) throw new Error('NIS wajib diisi');
  if (!obj.catatan || !obj.catatan.trim()) throw new Error('Catatan tidak boleh kosong');

  var sh    = _ensureCatatanSheet_();
  var data  = sh.getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  var tglNow = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Cari baris yang sudah ada (email_guru + kelas + nis)
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').toLowerCase() !== email) continue;
    if (String(data[i][2]) !== String(obj.kelas)) continue;
    if (String(data[i][3]) !== String(obj.nis)) continue;
    // Update
    sh.getRange(i + 1, 5).setValue(obj.nama || data[i][4]);
    sh.getRange(i + 1, 6).setValue(obj.catatan.trim());
    sh.getRange(i + 1, 7).setValue(tglNow);
    logAudit('UPDATE_CATATAN', email, 'NIS: ' + obj.nis + ' Kelas: ' + obj.kelas);
    return { success: true, action: 'updated' };
  }

  // Insert baru
  var id = 'CAT_' + new Date().getTime();
  sh.appendRow([id, email, obj.kelas, String(obj.nis), obj.nama || '', obj.catatan.trim(), tglNow]);
  logAudit('SAVE_CATATAN', email, 'NIS: ' + obj.nis + ' Kelas: ' + obj.kelas);
  return { success: true, action: 'created' };
}

/**
 * deleteCatatanSiswa(kelas, nis) — Hapus catatan siswa
 */
function deleteCatatanSiswa(kelas, nis) {
  assertLicenseActive();
  var email = getLoginEmail();
  var sh   = _ensureCatatanSheet_();
  var data = sh.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1] || '').toLowerCase() !== email) continue;
    if (String(data[i][2]) !== String(kelas)) continue;
    if (String(data[i][3]) !== String(nis)) continue;
    sh.deleteRow(i + 1);
    logAudit('DELETE_CATATAN', email, 'NIS: ' + nis + ' Kelas: ' + kelas);
    return { success: true };
  }
  return { success: false, message: 'Catatan tidak ditemukan' };
}

/**
 * getAllCatatanKepsek() — Kepsek: lihat semua catatan di sekolah
 */
function getAllCatatanKepsek() {
  assertKepsek_();
  var sh   = _ensureCatatanSheet_();
  var data = sh.getDataRange().getValues();
  var result = [];

  // Mapping setting untuk nama guru
  var setData  = sheet('SETTING');
  var namaMap  = {};
  if (setData) {
    var sd = setData.getDataRange().getValues();
    for (var si = 1; si < sd.length; si++) {
      namaMap[String(sd[si][0] || '').toLowerCase()] = String(sd[si][4] || '');
    }
  }

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var gEmail = String(data[i][1] || '').toLowerCase();
    result.push({
      id          : String(data[i][0]),
      guru        : namaMap[gEmail] || gEmail,
      email_guru  : gEmail,
      kelas       : String(data[i][2] || ''),
      nis         : String(data[i][3] || ''),
      nama        : String(data[i][4] || ''),
      catatan     : String(data[i][5] || ''),
      tgl         : String(data[i][6] || '')
    });
  }
  result.sort(function(a, b) { return b.tgl > a.tgl ? 1 : -1; });
  return result;
}
