/**
 * Catatan.js — Catatan Khusus Siswa
 * Guru dapat menulis catatan per siswa (perilaku, prestasi, dll).
 * Sheet CATATAN: [0]=id, [1]=email_guru, [2]=kelas, [3]=nis, [4]=nama, [5]=catatan, [6]=tgl, [7]=tipe
 *
 * tipe: 'individu' (default, siswa dengan masalah serius/khusus) atau
 * 'kelompok' (catatan yang sama ditulis sekali lalu diterapkan ke
 * beberapa siswa binaan sekaligus — supaya guru wali tidak wajib
 * mengetik ulang catatan yang sama untuk tiap siswa satu-satu kalau
 * memang bukan kasus individual).
 */

var CATATAN_SHEET = 'CATATAN';

function _ensureCatatanSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(CATATAN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CATATAN_SHEET);
    sh.appendRow(['id', 'email_guru', 'kelas', 'nis', 'nama', 'catatan', 'tgl', 'tipe']);
    sh.setFrozenRows(1);
  } else {
    var lastCol = sh.getLastColumn();
    var header = lastCol > 0
      ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||'').toLowerCase().trim(); })
      : [];
    if (header.indexOf('tipe') === -1) {
      sh.getRange(1, lastCol + 1).setValue('tipe');
    }
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
      tgl     : String(row[6] || ''),
      tipe    : String(row[7] || 'individu') || 'individu'
    });
  }
  return result;
}

/**
 * saveCatatanSiswa(obj) — Simpan atau update catatan SATU siswa (dipakai
 * untuk kasus serius/individual). obj: {kelas, nis, nama, catatan}
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
    sh.getRange(i + 1, 8).setValue('individu');
    logAudit('UPDATE_CATATAN', email, 'NIS: ' + obj.nis + ' Kelas: ' + obj.kelas);
    return { success: true, action: 'updated' };
  }

  // Insert baru
  var id = 'CAT_' + new Date().getTime();
  sh.appendRow([id, email, obj.kelas, String(obj.nis), obj.nama || '', obj.catatan.trim(), tglNow, 'individu']);
  logAudit('SAVE_CATATAN', email, 'NIS: ' + obj.nis + ' Kelas: ' + obj.kelas);
  return { success: true, action: 'created' };
}

/**
 * saveCatatanSiswaGrup(obj) — Simpan SATU catatan yang sama untuk BEBERAPA
 * siswa binaan sekaligus, supaya guru wali tidak perlu mengetik ulang
 * catatan yang sama satu-satu kalau memang bukan kasus individual (mis.
 * "kelompok ini aktif diskusi hari ini"). Tetap menyimpan satu baris per
 * siswa (skema tidak berubah) agar konsisten dengan getCatatanSiswa() dan
 * saveCatatanSiswa() — hanya cara pengisiannya yang di-batch, ditandai
 * tipe='kelompok'. Kalau siswa itu nanti dapat catatan individual, baris
 * kelompoknya tertimpa seperti update biasa (perilaku upsert yang sama
 * seperti saveCatatanSiswa()).
 * obj: {kelas, siswa: [{nis, nama}, ...], catatan}
 */
function saveCatatanSiswaGrup(obj) {
  assertLicenseActive();
  var email = getLoginEmail();

  if (!obj || !obj.kelas) throw new Error('Kelas wajib diisi');
  var siswaList = Array.isArray(obj.siswa) ? obj.siswa.filter(function(s){ return s && s.nis; }) : [];
  if (!siswaList.length) throw new Error('Pilih minimal satu siswa untuk catatan kelompok');
  if (!obj.catatan || !obj.catatan.trim()) throw new Error('Catatan tidak boleh kosong');

  var sh     = _ensureCatatanSheet_();
  var data   = sh.getDataRange().getValues();
  var tz     = Session.getScriptTimeZone();
  var tglNow = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var teks   = obj.catatan.trim();

  // Index baris yang sudah ada (email_guru + kelas + nis) supaya update
  // di-batch (bukan appendRow/deleteRow satu-satu di dalam loop).
  var existingRowIdx = {}; // nis -> row index (0-based, termasuk header)
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').toLowerCase() !== email) continue;
    if (String(data[i][2]) !== String(obj.kelas)) continue;
    existingRowIdx[String(data[i][3])] = i;
  }

  var newRows = [];
  var updated = 0, created = 0;
  siswaList.forEach(function(s) {
    var nis = String(s.nis);
    if (existingRowIdx.hasOwnProperty(nis)) {
      var rowIdx = existingRowIdx[nis];
      sh.getRange(rowIdx + 1, 5).setValue(s.nama || data[rowIdx][4]);
      sh.getRange(rowIdx + 1, 6).setValue(teks);
      sh.getRange(rowIdx + 1, 7).setValue(tglNow);
      sh.getRange(rowIdx + 1, 8).setValue('kelompok');
      updated++;
    } else {
      var id = 'CAT_' + new Date().getTime() + '_' + created;
      newRows.push([id, email, obj.kelas, nis, s.nama || '', teks, tglNow, 'kelompok']);
      created++;
    }
  });

  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
  }

  logAudit('SAVE_CATATAN_GRUP', email, 'Kelas: ' + obj.kelas + ' | ' + siswaList.length + ' siswa');
  return { success: true, updated: updated, created: created, total: siswaList.length };
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
