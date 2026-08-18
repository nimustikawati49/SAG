/**
 * Kalender.js — Kalender Hari Libur (input manual SuperAdmin)
 *
 * Sheet HARI_LIBUR (central-only, lihat _isCentralOnlySheet_ di Code.js)
 * menyimpan tanggal yang tidak efektif untuk mengajar — dipakai untuk
 * menandai kartu "Jadwal Mengajar" di Dashboard guru supaya jelas kalau
 * suatu hari itu libur, bukan cuma tidak ada jadwal.
 *
 * CATATAN: sebelumnya ada juga sinkronisasi otomatis dari kalender publik
 * Google ("Holidays in Indonesia") lewat CalendarApp — DIHAPUS karena
 * akun Google Workspace deployment ini (domain @guru.smp.belajar.id)
 * menolak izin scope Calendar terus-menerus (kemungkinan dibatasi
 * kebijakan admin Workspace sekolah, di luar kendali SuperAdmin aplikasi
 * ini) — dicoba re-otorisasi manual pun tetap gagal. Semua hari libur
 * sekarang HANYA lewat input manual SuperAdmin (addHariLiburManual),
 * kolom 'sumber' tetap dipertahankan di sheet untuk kompatibilitas data
 * lama (baris ber-sumber 'nasional' dari percobaan sync sebelumnya,
 * kalau ada, tetap tampil apa adanya — cuma tidak ada lagi cara
 * menambah baris baru ber-sumber itu).
 */

var HARI_LIBUR_SHEET_ = 'HARI_LIBUR';

function _ensureHariLiburSheet_() {
  var ss = getCentralSpreadsheet_();
  var sh = ss.getSheetByName(HARI_LIBUR_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(HARI_LIBUR_SHEET_);
    sh.appendRow(['tanggal', 'keterangan', 'sumber', 'created_at']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e5e7eb');
  }
  return sh;
}

function _formatTanggalISO_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * addHariLiburManual(tanggal, keterangan)
 * SuperAdmin only — satu-satunya cara mengisi HARI_LIBUR (lihat catatan
 * di header file ini soal sinkronisasi otomatis yang dihapus). Mencakup
 * hari libur nasional MAUPUN khusus sekolah — semuanya diinput manual di
 * sini. Kalau tanggal itu sudah ada, keterangannya diganti.
 */
function addHariLiburManual(tanggal, keterangan) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var tgl = String(tanggal || '').trim();
  var ket = String(keterangan || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tgl)) throw new Error('Format tanggal tidak valid (yyyy-mm-dd)');
  if (!ket) throw new Error('Keterangan tidak boleh kosong');

  var sh = _ensureHariLiburSheet_();
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === tgl) {
      sh.getRange(i + 1, 2).setValue(ket);
      sh.getRange(i + 1, 3).setValue('sekolah');
      logAudit('UPDATE_HARI_LIBUR', getLoginEmail(), tgl + ' | ' + ket);
      return { success: true, action: 'updated' };
    }
  }

  sh.appendRow([tgl, ket, 'sekolah', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')]);
  logAudit('ADD_HARI_LIBUR', getLoginEmail(), tgl + ' | ' + ket);
  return { success: true, action: 'created' };
}

/**
 * deleteHariLibur(tanggal)
 * SuperAdmin only.
 */
function deleteHariLibur(tanggal) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var tgl = String(tanggal || '').trim();
  var sh = _ensureHariLiburSheet_();
  var rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0] || '').trim() === tgl) {
      sh.deleteRow(i + 1);
      logAudit('DELETE_HARI_LIBUR', getLoginEmail(), tgl);
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * getHariLiburList()
 * SuperAdmin only — daftar lengkap untuk panel manajemen, urut tanggal.
 */
function getHariLiburList() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh = _ensureHariLiburSheet_();
  var rows = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var tgl = String(rows[i][0] || '').trim();
    if (!tgl) continue;
    result.push({
      tanggal: tgl,
      keterangan: String(rows[i][1] || ''),
      sumber: String(rows[i][2] || 'sekolah').toLowerCase().trim()
    });
  }
  result.sort(function (a, b) { return a.tanggal.localeCompare(b.tanggal); });
  return result;
}

/**
 * _getHariLiburMap_(dariISO, sampaiISO)
 * Internal — { 'yyyy-MM-dd': keterangan } untuk rentang tanggal INKLUSIF.
 * Dipakai getDashboardJadwal() (Jadwal.js) untuk menandai kartu Jadwal
 * Mengajar. Tidak ada gate role di sini (dipakai dari fungsi yang sudah
 * di-gate di pemanggilnya) — read-only, tidak sensitif.
 */
function _getHariLiburMap_(dariISO, sampaiISO) {
  var map = {};
  try {
    var sh = _ensureHariLiburSheet_();
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var tgl = String(rows[i][0] || '').trim();
      if (!tgl) continue;
      if (dariISO && tgl < dariISO) continue;
      if (sampaiISO && tgl > sampaiISO) continue;
      map[tgl] = String(rows[i][1] || '');
    }
  } catch (e) { /* fail-soft — jangan sampai dashboard gagal gara-gara ini */ }
  return map;
}

/**
 * getHariLiburMingguIni()
 * Untuk SEMUA role login — peta hari libur Senin-Sabtu minggu berjalan
 * (mengikuti timezone script). Dipanggil terpisah dari client kalau
 * suatu tampilan butuh info libur tanpa menarik jadwal lengkap.
 */
function getHariLiburMingguIni() {
  var auth = getAuth();
  if (!auth || !auth.email) return {};

  var now = new Date();
  var day = now.getDay(); // 0=Minggu
  var mondayOffset = day === 0 ? -6 : 1 - day;
  var monday = new Date(now.getFullYear(), now.getMonth(), now.getDate() + mondayOffset);
  var saturday = new Date(monday.getTime() + 5 * 24 * 60 * 60 * 1000);

  return _getHariLiburMap_(_formatTanggalISO_(monday), _formatTanggalISO_(saturday));
}
