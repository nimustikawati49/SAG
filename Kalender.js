/**
 * Kalender.js — Kalender Hari Libur / Kegiatan Sekolah (input manual)
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
 * sekarang HANYA lewat input manual (addHariLiburManual), kolom 'sumber'
 * tetap dipertahankan di sheet untuk kompatibilitas data lama (baris
 * ber-sumber 'nasional' dari percobaan sync sebelumnya, kalau ada, tetap
 * tampil apa adanya — cuma tidak ada lagi cara menambah baris baru
 * ber-sumber itu).
 *
 * SCOPING PER SEKOLAH: awalnya fitur ini SuperAdmin-only dan kalendernya
 * GLOBAL (tidak dipisah per sekolah) — aman selama cuma dipakai SuperAdmin
 * yang memang mengawasi semua sekolah. Begitu admin biasa (guru) ikut
 * diberi akses Tambah/Hapus (lewat menu Setting), kalender WAJIB dipisah
 * per kode_sekolah — supaya guru di sekolah A tidak bisa menambah/hapus
 * hari libur yang kepakai guru sekolah B. Baris LAMA (sebelum migrasi
 * kolom kode_sekolah ini) sengaja dibiarkan TANPA kode_sekolah dan tetap
 * tampil ke SEMUA sekolah — supaya data lama tidak mendadak hilang,
 * sampai ada yang menambahkan ulang secara scoped.
 */

var HARI_LIBUR_SHEET_ = 'HARI_LIBUR';

function _ensureHariLiburSheet_() {
  var ss = getCentralSpreadsheet_();
  var sh = ss.getSheetByName(HARI_LIBUR_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(HARI_LIBUR_SHEET_);
    sh.appendRow(['tanggal', 'keterangan', 'sumber', 'created_at', 'kode_sekolah']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e5e7eb');
  } else {
    var lastCol = sh.getLastColumn();
    var header = lastCol > 0
      ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').toLowerCase().trim(); })
      : [];
    if (header.indexOf('kode_sekolah') === -1) {
      sh.getRange(1, lastCol + 1).setValue('kode_sekolah');
    }
  }
  return sh;
}

/** Peta nama kolom -> index (0-based), dicari lewat header — tahan migrasi kolom baru. */
function _hariLiburColIdx_(sh) {
  var lastCol = sh.getLastColumn();
  var header = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').toLowerCase().trim(); })
    : [];
  var idx = {};
  header.forEach(function (h, i) { if (h) idx[h] = i; });
  return idx;
}

function _formatTanggalISO_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * _assertHariLiburAkses_()
 * admin (guru) DAN superadmin boleh kelola Hari Libur — admin lewat menu
 * Setting, superadmin lewat panel SuperAdmin. Guest/role lain ditolak.
 */
function _assertHariLiburAkses_() {
  var auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin') {
    throw new Error('AKSES_DITOLAK');
  }
  return auth;
}

/**
 * kode_sekolah milik akun yang login — dicari dari USERS sheet, sama
 * seperti _kepsekActiveGuruEmails_() di Kepsek.js. Kosong kalau akun ini
 * belum di-assign sekolah oleh SuperAdmin (kolom "Kode/Nama Sekolah").
 */
function _callerKodeSekolahHariLibur_(email) {
  var shUsers = typeof _ensureUsersKodeSekolahColumn_ === 'function'
    ? _ensureUsersKodeSekolahColumn_()
    : _getCentralSheetByName_('USERS');
  if (!shUsers) return '';
  var data = shUsers.getDataRange().getValues();
  var target = String(email || '').toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').toLowerCase().trim() === target) {
      return String(data[i][4] || '').trim();
    }
  }
  return '';
}

/** true kalau baris ini boleh dilihat/diubah oleh akun dengan kodeSekolah tsb (lihat catatan scoping di atas file ini). */
function _hariLiburRowVisible_(rowKode, kodeSekolah) {
  if (!rowKode) return true; // baris lama tanpa kode_sekolah — tampil ke semua
  if (!kodeSekolah) return false; // baris sudah ber-sekolah, caller belum di-assign — jangan bocor
  return rowKode === kodeSekolah;
}

/**
 * addHariLiburManual(tanggal, keterangan)
 * Satu tanggal saja — dipertahankan sebagai wrapper tipis di atas
 * addHariLiburRangeManual() (dari=sampai=tanggal yang sama) supaya
 * caller lama tetap jalan tanpa logika dobel.
 */
function addHariLiburManual(tanggal, keterangan) {
  var tgl = String(tanggal || '').trim();
  return addHariLiburRangeManual(tgl, tgl, keterangan);
}

/**
 * addHariLiburRangeManual(dariISO, sampaiISO, keterangan)
 * Isi hari libur untuk SATU tanggal ATAU RENTANG tanggal sekaligus
 * (kegiatan sekolah biasanya berlangsung beberapa hari) — bikin/timpa 1
 * baris per tanggal di rentang itu, keterangan yang sama untuk semua
 * tanggal. Dibatasi 31 hari sekali input supaya salah ketik tahun tidak
 * bikin ribuan baris tak sengaja.
 */
function addHariLiburRangeManual(dariISO, sampaiISO, keterangan) {
  var auth = _assertHariLiburAkses_();
  assertLicenseActive();

  var dari   = String(dariISO   || '').trim();
  var sampai = String(sampaiISO || dari || '').trim();
  var ket    = String(keterangan || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dari) || !/^\d{4}-\d{2}-\d{2}$/.test(sampai)) {
    throw new Error('Format tanggal tidak valid (yyyy-mm-dd)');
  }
  if (!ket) throw new Error('Keterangan tidak boleh kosong');
  if (sampai < dari) throw new Error('Sampai Tanggal tidak boleh sebelum Dari Tanggal');

  var dDari   = new Date(dari   + 'T00:00:00');
  var dSampai = new Date(sampai + 'T00:00:00');
  var totalHari = Math.round((dSampai - dDari) / 86400000) + 1;
  if (totalHari > 31) throw new Error('Rentang maksimal 31 hari sekali input — pecah jadi beberapa kali kalau lebih panjang');

  var kodeSekolah = _callerKodeSekolahHariLibur_(auth.email);
  if (!kodeSekolah) throw new Error('Akun Anda belum di-assign ke sekolah (kode_sekolah) — hubungi SuperAdmin untuk mengisi kolom "Kode/Nama Sekolah" dulu.');

  var sh  = _ensureHariLiburSheet_();
  var idx = _hariLiburColIdx_(sh);
  var rows = sh.getDataRange().getValues();

  // tanggal|kode_sekolah -> nomor baris (1-based) untuk baris yang SUDAH
  // ada, supaya tanggal yang sudah tercatat cukup ditimpa keterangannya
  // (bukan dobel baris) — sisanya dikumpulkan untuk sekali batch-append.
  var existingRow = {};
  for (var i = 1; i < rows.length; i++) {
    var key = String(rows[i][idx.tanggal] || '').trim() + '|' + String(rows[i][idx.kode_sekolah] || '').trim();
    existingRow[key] = i + 1;
  }

  var newRows = [];
  var jumlah  = 0;
  for (var d = new Date(dDari); d <= dSampai; d.setDate(d.getDate() + 1)) {
    var tgl = _formatTanggalISO_(d);
    var key = tgl + '|' + kodeSekolah;
    var rNum = existingRow[key];
    if (rNum) {
      sh.getRange(rNum, idx.keterangan + 1).setValue(ket);
      sh.getRange(rNum, idx.sumber + 1).setValue('sekolah');
    } else {
      var newRow = new Array(sh.getLastColumn()).fill('');
      newRow[idx.tanggal]      = tgl;
      newRow[idx.keterangan]   = ket;
      newRow[idx.sumber]       = 'sekolah';
      newRow[idx.created_at]   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      newRow[idx.kode_sekolah] = kodeSekolah;
      newRows.push(newRow);
    }
    jumlah++;
  }
  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, sh.getLastColumn()).setValues(newRows);
  }

  logAudit('ADD_HARI_LIBUR', auth.email, dari + (sampai !== dari ? (' s/d ' + sampai) : '') + ' | ' + ket + ' | ' + jumlah + ' hari');
  return { success: true, jumlah: jumlah };
}

/**
 * deleteHariLibur(tanggal)
 * Cuma bisa menghapus baris milik SEKOLAH SENDIRI (atau baris lama tanpa
 * kode_sekolah) — tidak bisa menghapus kalender sekolah lain.
 */
function deleteHariLibur(tanggal) {
  var auth = _assertHariLiburAkses_();
  var tgl = String(tanggal || '').trim();
  var kodeSekolah = _callerKodeSekolahHariLibur_(auth.email);

  var sh  = _ensureHariLiburSheet_();
  var idx = _hariLiburColIdx_(sh);
  var rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idx.tanggal] || '').trim() !== tgl) continue;
    var rowKode = String(rows[i][idx.kode_sekolah] || '').trim();
    if (!_hariLiburRowVisible_(rowKode, kodeSekolah)) continue;
    sh.deleteRow(i + 1);
    logAudit('DELETE_HARI_LIBUR', auth.email, tgl);
    return { success: true };
  }
  return { success: false };
}

/**
 * getHariLiburList()
 * Daftar lengkap Hari Libur milik sekolah akun yang login, untuk panel
 * manajemen (SuperAdmin maupun Setting admin) — urut tanggal.
 */
function getHariLiburList() {
  var auth = _assertHariLiburAkses_();
  var kodeSekolah = _callerKodeSekolahHariLibur_(auth.email);

  var sh  = _ensureHariLiburSheet_();
  var idx = _hariLiburColIdx_(sh);
  var rows = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var tgl = String(rows[i][idx.tanggal] || '').trim();
    if (!tgl) continue;
    var rowKode = String(rows[i][idx.kode_sekolah] || '').trim();
    if (!_hariLiburRowVisible_(rowKode, kodeSekolah)) continue;
    result.push({
      tanggal: tgl,
      keterangan: String(rows[i][idx.keterangan] || ''),
      sumber: String(rows[i][idx.sumber] || 'sekolah').toLowerCase().trim()
    });
  }
  result.sort(function (a, b) { return a.tanggal.localeCompare(b.tanggal); });
  return result;
}

/**
 * _getHariLiburMap_(dariISO, sampaiISO)
 * Internal — { 'yyyy-MM-dd': keterangan } untuk rentang tanggal INKLUSIF,
 * dipersempit ke kode_sekolah akun yang sedang login. Dipakai
 * getDashboardJadwal() (Jadwal.js) untuk menandai kartu Jadwal Mengajar.
 * Tidak ada gate role di sini (dipakai dari fungsi yang sudah di-gate di
 * pemanggilnya) — read-only, tidak sensitif.
 */
function _getHariLiburMap_(dariISO, sampaiISO) {
  var map = {};
  try {
    var auth = getAuth();
    var kodeSekolah = _callerKodeSekolahHariLibur_(auth.email);

    var sh  = _ensureHariLiburSheet_();
    var idx = _hariLiburColIdx_(sh);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var tgl = String(rows[i][idx.tanggal] || '').trim();
      if (!tgl) continue;
      if (dariISO && tgl < dariISO) continue;
      if (sampaiISO && tgl > sampaiISO) continue;
      var rowKode = String(rows[i][idx.kode_sekolah] || '').trim();
      if (!_hariLiburRowVisible_(rowKode, kodeSekolah)) continue;
      map[tgl] = String(rows[i][idx.keterangan] || '');
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
