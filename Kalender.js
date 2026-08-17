/**
 * Kalender.js — Kalender Hari Libur (Nasional + Sekolah)
 *
 * Sheet HARI_LIBUR (central-only, lihat _isCentralOnlySheet_ di Code.js)
 * menyimpan tanggal yang tidak efektif untuk mengajar — dipakai untuk
 * menandai kartu "Jadwal Mengajar" di Dashboard guru supaya jelas kalau
 * suatu hari itu libur, bukan cuma tidak ada jadwal.
 *
 * Dua sumber data:
 *  1. 'nasional' — ditarik otomatis dari kalender publik resmi Google
 *     ("Holidays in Indonesia", ID: en.indonesian#holiday@group.v.calendar.google.com)
 *     lewat syncKalenderNasional() (SuperAdmin only, manual trigger —
 *     BUKAN otomatis tiap hari, supaya tidak membebani kuota Calendar API
 *     tanpa perlu). Kalender publik ini TIDAK mencakup hal spesifik
 *     sekolah (libur semester, hari raya lokal, dst).
 *  2. 'sekolah' — ditambah manual oleh SuperAdmin lewat addHariLiburManual()
 *     untuk hari libur yang tidak ada di kalender nasional. Entry
 *     'sekolah' TIDAK PERNAH ditimpa oleh sync nasional (lihat
 *     syncKalenderNasional — hanya menulis/update baris ber-sumber
 *     'nasional').
 */

var HARI_LIBUR_SHEET_ = 'HARI_LIBUR';
var KALENDER_NASIONAL_ID_ = 'en.indonesian#holiday@group.v.calendar.google.com';

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
 * syncKalenderNasional()
 * SuperAdmin only — tarik hari libur nasional Indonesia dari kalender
 * publik Google untuk rentang -30 hari s.d. +400 hari dari sekarang
 * (cukup untuk menutup 1 tahun ajaran penuh + sisa semester berjalan).
 * Upsert per tanggal: kalau sudah ada baris 'sekolah' untuk tanggal itu,
 * DILEWATI (tidak ditimpa) — entry manual sekolah selalu menang. Kalau
 * sudah ada baris 'nasional' lama, keterangannya diperbarui kalau beda.
 */
function syncKalenderNasional() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');

  var cal;
  try {
    cal = CalendarApp.getCalendarById(KALENDER_NASIONAL_ID_);
  } catch (e) {
    throw new Error('Gagal membuka kalender nasional: ' + e.message);
  }
  if (!cal) {
    throw new Error('Kalender publik "Holidays in Indonesia" tidak ditemukan/tidak bisa diakses.');
  }

  var now = new Date();
  var start = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  var end = new Date(now.getTime() + 400 * 24 * 60 * 60 * 1000);

  var events;
  try {
    events = cal.getEvents(start, end);
  } catch (e) {
    throw new Error('Gagal membaca event kalender nasional: ' + e.message);
  }

  var sh = _ensureHariLiburSheet_();
  var rows = sh.getDataRange().getValues();
  var existingByTanggal = {}; // tanggal -> { rowNum, sumber }
  for (var i = 1; i < rows.length; i++) {
    var t = String(rows[i][0] || '').trim();
    if (!t) continue;
    existingByTanggal[t] = { rowNum: i + 1, sumber: String(rows[i][2] || '').toLowerCase().trim(), keterangan: String(rows[i][1] || '') };
  }

  var now_ = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var added = 0, updated = 0, skipped = 0;
  var appendBatch = [];

  events.forEach(function (ev) {
    var tanggal;
    try { tanggal = _formatTanggalISO_(ev.getAllDayStartDate ? ev.getAllDayStartDate() : ev.getStartTime()); }
    catch (e) { return; }
    var keterangan = String(ev.getTitle() || '').trim();
    if (!tanggal || !keterangan) return;

    var existing = existingByTanggal[tanggal];
    if (existing) {
      if (existing.sumber === 'sekolah') { skipped++; return; } // entry sekolah menang, jangan ditimpa
      if (existing.keterangan !== keterangan) {
        sh.getRange(existing.rowNum, 2).setValue(keterangan);
        updated++;
      }
      return;
    }

    appendBatch.push([tanggal, keterangan, 'nasional', now_]);
    existingByTanggal[tanggal] = { rowNum: -1, sumber: 'nasional', keterangan: keterangan }; // cegah duplikat dalam batch yang sama
    added++;
  });

  if (appendBatch.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appendBatch.length, 4).setValues(appendBatch);
  }

  logAudit('SYNC_KALENDER_NASIONAL', getLoginEmail(), 'ditambah: ' + added + ' | diperbarui: ' + updated + ' | dilewati (sudah ada entry sekolah): ' + skipped);
  return { success: true, added: added, updated: updated, skipped: skipped };
}

/**
 * addHariLiburManual(tanggal, keterangan)
 * SuperAdmin only — tambah/ubah 1 hari libur khusus sekolah (sumber
 * 'sekolah'), mis. libur semester, hari raya lokal yang tidak ada di
 * kalender nasional, dst. Kalau tanggal itu sudah ada (dari sumber mana
 * pun), keterangannya diganti dan sumbernya jadi 'sekolah'.
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
