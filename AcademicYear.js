/**
 * AcademicYear.js
 * Backward-compatible multi-year academic data layer.
 */

/**
 * _getCalendarBasedPeriod_() — fallback terakhir jika setting belum ada.
 * SMP/SMA Indonesia: Juli–Des = Ganjil, Jan–Jun = Genap.
 */
function _getCalendarBasedPeriod_() {
  var now   = new Date();
  var month = now.getMonth() + 1;
  var year  = now.getFullYear();
  return {
    tahun_pelajaran: month >= 7 ? year + '/' + (year + 1) : (year - 1) + '/' + year,
    semester: month >= 7 ? 'Ganjil' : 'Genap'
  };
}

function ensureAcademicSchema_() {
  var cacheKey = 'ACADEMIC_SCHEMA_READY';
  try {
    var mode = (typeof getStorageMode_ === 'function' ? getStorageMode_() : 'central');
    var email = Session.getEffectiveUser().getEmail().toLowerCase().trim();
    if (mode === 'per_guru' && email) cacheKey += '_' + email.replace(/[^a-z0-9]/gi, '_');
  } catch (e) {}

  const cache = CacheService.getScriptCache();
  const hit = cache.get(cacheKey);
  if (hit) return;

  const lock = LockService.getScriptLock();
  let gotLock = false;
  try {
    gotLock = lock.tryLock(10000);
  } catch (e) {
    gotLock = false;
  }
  if (!gotLock) {
    // Skenario multi-guru: proses lain sedang memegang lock (mis. sedang
    // ensure schema juga). Jangan lempar error ke pengguna — schema ini
    // idempoten dan sudah pernah berhasil dibuat sebelumnya di 99% kasus,
    // jadi lebih aman lanjut tanpa re-check daripada memblokir aksi guru
    // (simpan setting, upload logo, dll) dengan "Lock timeout".
    try { logError_('ENSURE_ACADEMIC_SCHEMA_LOCK_BUSY', new Error('Lock sedang dipakai proses lain, schema check dilewati')); } catch (e2) {}
    return;
  }
  try {
    const hit2 = cache.get(cacheKey);
    if (hit2) return;

    const ss = getSpreadsheet_();

    ensureSheetWithHeader_(ss, 'MasterTahunPelajaran', [
      'id', 'tahun_pelajaran', 'semester', 'status', 'created_at'
    ]);
    ensureSheetWithHeader_(ss, 'MasterSiswa', [
      'id', 'nis', 'nisn', 'nama', 'jk', 'ttl', 'alamat', 'orang_tua', 'kontak', 'status', 'created_at', 'updated_at'
    ]);
    ensureSheetWithHeader_(ss, 'RiwayatKelas', [
      'id', 'tahun_pelajaran', 'semester', 'siswa_id', 'kelas', 'status', 'created_at', 'updated_at'
    ]);
    ensureSheetWithHeader_(ss, 'GuruMengajar', [
      'id', 'guru', 'tahun_pelajaran', 'semester', 'kelas', 'mapel', 'created_at'
    ]);

    ensureSettingAcademicColumns_(ss);
    ensureJurnalAcademicColumns_(ss);
    migrateLegacyData_();

    cache.put(cacheKey, '1', 3600);
  } finally {
    lock.releaseLock();
  }
}

function ensureSheetWithHeader_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const lastCol = Math.max(sh.getLastColumn(), header.length);
  const firstRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const isEmptyHeader = firstRow.every(function(v) { return String(v || '').trim() === ''; });
  if (isEmptyHeader) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#e5e7eb');
    return;
  }

  const existing = firstRow.map(function(v) { return String(v || '').toLowerCase().trim(); });
  let updated = false;
  header.forEach(function(col) {
    if (existing.indexOf(String(col).toLowerCase()) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(col);
      updated = true;
    }
  });
  if (updated) {
    const lc = sh.getLastColumn();
    sh.getRange(1, 1, 1, lc).setFontWeight('bold').setBackground('#e5e7eb');
  }
}

function ensureSettingAcademicColumns_(ss) {
  const sh = ss.getSheetByName('SETTING');
  if (!sh || sh.getLastRow() < 1) return;
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h || '').toLowerCase().trim(); });
  if (header.indexOf('tahun_pelajaran_aktif') === -1) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('tahun_pelajaran_aktif');
  }
  if (header.indexOf('semester_aktif') === -1) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('semester_aktif');
  }
}

function ensureJurnalAcademicColumns_(ss) {
  const sh = ss.getSheetByName('JURNAL');
  if (!sh || sh.getLastRow() < 1) return;
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h || '').toLowerCase().trim(); });
  if (header.indexOf('mapel') === -1) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('mapel');
  }
}

function getLegacyDefaultYear_() {
  return PropertiesService.getScriptProperties().getProperty('LEGACY_DEFAULT_TAHUN') || 'Default Tahun Lama';
}

function migrateLegacyData_() {
  const cache = CacheService.getScriptCache();
  var cacheKey = 'ACADEMIC_MIGRATED';
  try {
    var mode = (typeof getStorageMode_ === 'function' ? getStorageMode_() : 'central');
    var email = Session.getEffectiveUser().getEmail().toLowerCase().trim();
    if (mode === 'per_guru' && email) cacheKey += '_' + email.replace(/[^a-z0-9]/gi, '_');
  } catch (e) {}
  if (cache.get(cacheKey)) return;

  const ss = getSpreadsheet_();
  const shSiswa = ss.getSheetByName('SISWA');
  const shMaster = ss.getSheetByName('MasterSiswa');
  const shRiwayat = ss.getSheetByName('RiwayatKelas');
  const shSet = ss.getSheetByName('SETTING');
  const shTahun = ss.getSheetByName('MasterTahunPelajaran');
  const shJurnal = ss.getSheetByName('JURNAL');

  const defaultTahun = getLegacyDefaultYear_();
  const defaultSemester = 'Ganjil';

  // Pakai versi RAW (tanpa ensureAcademicSchema_ lagi) -- kita SUDAH di
  // dalam ensureAcademicSchema_() sekarang (memegang lock-nya), jadi
  // memanggil ensureAcademicYearExists_()/setGlobalAcademicPeriod() biasa
  // di sini bikin ensureAcademicSchema_ terpanggil ULANG 2 LAPIS secara
  // rekursif SAAT LOCK MASIH DIPEGANG -- salah satu penyebab lock
  // "dipakai proses lain" bertumpuk yang terlihat di log timing.
  _ensureAcademicYearExistsRaw_(defaultTahun, defaultSemester, true);

  if (shSiswa && shMaster && shRiwayat) {
    const masterRows = shMaster.getDataRange().getValues();
    const masterByNis = {};
    for (let i = 1; i < masterRows.length; i++) {
      const nis = String(masterRows[i][1] || '').trim();
      const nisn = String(masterRows[i][2] || '').trim();
      if (nis) masterByNis['NIS:' + nis] = masterRows[i][0];
      if (nisn) masterByNis['NISN:' + nisn] = masterRows[i][0];
    }

    const riwayatRows = shRiwayat.getDataRange().getValues();
    const riwayatKey = {};
    for (let i = 1; i < riwayatRows.length; i++) {
      const k = [riwayatRows[i][1], riwayatRows[i][2], riwayatRows[i][3], riwayatRows[i][4]].join('|');
      riwayatKey[k] = true;
    }

    // Kumpulkan baris baru dulu, tulis SEKALI di akhir lewat setValues()
    // batch -- bukan appendRow() satu-satu per siswa (bisa ratusan baris
    // untuk sekolah dengan banyak siswa, sama kelasnya dengan bug JURNAL
    // di atas: N+1 write saat memegang lock global).
    const siswaRows = shSiswa.getDataRange().getValues();
    const now = new Date();
    const newMasterRows = [];
    const newRiwayatRows = [];
    for (let i = 1; i < siswaRows.length; i++) {
      const kelas = String(siswaRows[i][0] || '').trim();
      const nis = String(siswaRows[i][2] || '').trim();
      const nama = String(siswaRows[i][3] || '').trim();
      const jk = String(siswaRows[i][4] || '').trim();
      if (!nis || !nama) continue;

      let siswaId = masterByNis['NIS:' + nis] || '';
      if (!siswaId) {
        siswaId = 'SIS-' + Utilities.getUuid().slice(0, 8).toUpperCase();
        newMasterRows.push([siswaId, nis, '', nama, jk, '', '', '', '', 'AKTIF', now, now]);
        masterByNis['NIS:' + nis] = siswaId;
      }

      const k = [defaultTahun, defaultSemester, siswaId, kelas].join('|');
      if (!riwayatKey[k]) {
        newRiwayatRows.push([
          'RK-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
          defaultTahun,
          defaultSemester,
          siswaId,
          kelas,
          'AKTIF',
          now,
          now
        ]);
        riwayatKey[k] = true;
      }
    }
    if (newMasterRows.length) {
      shMaster.getRange(shMaster.getLastRow() + 1, 1, newMasterRows.length, newMasterRows[0].length).setValues(newMasterRows);
    }
    if (newRiwayatRows.length) {
      shRiwayat.getRange(shRiwayat.getLastRow() + 1, 1, newRiwayatRows.length, newRiwayatRows[0].length).setValues(newRiwayatRows);
    }
  }

  if (shSet && shSet.getLastRow() > 1) {
    const all = shSet.getDataRange().getValues();
    const header = all[0].map(function(h) { return String(h || '').toLowerCase().trim(); });
    const idxTahun = header.indexOf('tahun');
    const idxSem = header.indexOf('semester');
    const idxTaAktif = header.indexOf('tahun_pelajaran_aktif');
    const idxSemAktif = header.indexOf('semester_aktif');
    const numRows = all.length - 1;

    // Batch per kolom, bukan setValue() satu-satu per baris guru.
    if (idxTaAktif > -1) {
      let changed = false;
      const col = [];
      for (let i = 1; i < all.length; i++) {
        const cur = String(all[i][idxTaAktif] || '').trim();
        const tahunOld = idxTahun > -1 ? String(all[i][idxTahun] || '').trim() : '';
        if (!cur) { col.push([tahunOld || defaultTahun]); changed = true; }
        else col.push([cur]);
      }
      if (changed) shSet.getRange(2, idxTaAktif + 1, numRows, 1).setValues(col);
    }
    if (idxSemAktif > -1) {
      let changed = false;
      const col = [];
      for (let i = 1; i < all.length; i++) {
        const cur = String(all[i][idxSemAktif] || '').trim();
        const semOld = idxSem > -1 ? String(all[i][idxSem] || '').trim() : '';
        if (!cur) { col.push([semOld || defaultSemester]); changed = true; }
        else col.push([cur]);
      }
      if (changed) shSet.getRange(2, idxSemAktif + 1, numRows, 1).setValues(col);
    }
  }

  // Batch write, BUKAN setValue() satu-satu per baris. Sheet JURNAL bisa
  // punya ratusan/ribuan baris (data gabungan semua guru di sheet pusat)
  // -- versi lama di sini melakukan sampai 2 panggilan Sheets API
  // TERPISAH per baris SAAT MEMEGANG LOCK GLOBAL, gampang memakan
  // puluhan detik sampai menit dan memblokir semua guru lain.
  if (shJurnal && shJurnal.getLastRow() > 1) {
    const numRows = shJurnal.getLastRow() - 1;
    const tahunCol = shJurnal.getRange(2, 19, numRows, 1).getValues();
    const semCol = shJurnal.getRange(2, 14, numRows, 1).getValues();
    let tahunChanged = false;
    let semChanged = false;
    for (let i = 0; i < numRows; i++) {
      if (!tahunCol[i][0]) { tahunCol[i][0] = defaultTahun; tahunChanged = true; }
      if (!semCol[i][0]) { semCol[i][0] = defaultSemester; semChanged = true; }
    }
    if (tahunChanged) shJurnal.getRange(2, 19, numRows, 1).setValues(tahunCol);
    if (semChanged) shJurnal.getRange(2, 14, numRows, 1).setValues(semCol);
  }

  cache.put(cacheKey, '1', 21600);
}

/**
 * _ensureAcademicYearExistsRaw_(tahun, semester, setActive)
 * Versi TANPA ensureAcademicSchema_()/auth-check/audit-log, dipakai HANYA
 * dari dalam migrateLegacyData_() yang sudah berjalan di dalam
 * ensureAcademicSchema_()'s lock -- memanggil versi publik di sana bikin
 * ensureAcademicSchema_() terpanggil ulang 2 lapis secara rekursif SAAT
 * LOCK MASIH DIPEGANG, salah satu penyebab log "lock dipakai proses lain"
 * bertumpuk. Batched (bukan setValue() satu-satu per baris) untuk kasus
 * setActive juga.
 */
function _ensureAcademicYearExistsRaw_(tahun, semester, setActive) {
  tahun = String(tahun || '').trim();
  semester = String(semester || '').trim();
  if (!tahun || !semester) return;

  const sh = sheet('MasterTahunPelajaran');
  const rows = sh.getDataRange().getValues();

  if (setActive && rows.length > 1) {
    const statusCol = rows.slice(1).map(function(r) {
      const isTarget = String(r[1] || '') === tahun && String(r[2] || '') === semester;
      return [isTarget ? 'AKTIF' : 'NONAKTIF'];
    });
    sh.getRange(2, 4, statusCol.length, 1).setValues(statusCol);
  }

  let foundRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || '') === tahun && String(rows[i][2] || '') === semester) { foundRow = i + 1; break; }
  }
  if (foundRow === -1) {
    sh.appendRow(['TP-' + Utilities.getUuid().slice(0, 8).toUpperCase(), tahun, semester, setActive ? 'AKTIF' : 'NONAKTIF', new Date()]);
  }
}

function ensureAcademicYearExists_(tahun, semester, setActive) {
  ensureAcademicSchema_();
  tahun = String(tahun || '').trim();
  semester = String(semester || '').trim();
  if (!tahun || !semester) return;

  const sh = sheet('MasterTahunPelajaran');
  const rows = sh.getDataRange().getValues();
  let foundRow = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || '') === tahun && String(rows[i][2] || '') === semester) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) {
    sh.appendRow([
      'TP-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      tahun,
      semester,
      setActive ? 'AKTIF' : 'NONAKTIF',
      new Date()
    ]);
    foundRow = sh.getLastRow();
  }

  if (setActive) {
    setGlobalAcademicPeriod(tahun, semester);
  }
}

function listAcademicYears() {
  ensureAcademicSchema_();
  const sh = sheet('MasterTahunPelajaran');
  const rows = sh.getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return {
      id: r[0],
      tahun_pelajaran: r[1],
      semester: r[2],
      status: r[3],
      created_at: r[4]
    };
  });
}

function getGlobalAcademicPeriod_() {
  ensureAcademicSchema_();
  const rows = sheet('MasterTahunPelajaran').getDataRange().getValues();
  let active = null;
  for (let i = 1; i < rows.length; i++) {
    const st = String(rows[i][3] || '').toUpperCase();
    if (st === 'AKTIF' || st === 'ACTIVE') {
      active = { tahun_pelajaran: rows[i][1], semester: rows[i][2], row: i + 1 };
      break;
    }
  }
  return active;
}

function setGlobalAcademicPeriod(tahun, semester) {
  const auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin' && auth.role !== 'kepsek') {
    throw new Error('AKSES_DITOLAK');
  }
  ensureAcademicSchema_();
  tahun = String(tahun || '').trim();
  semester = String(semester || '').trim();
  if (!tahun || !semester) throw new Error('Tahun pelajaran dan semester wajib diisi');

  const sh = sheet('MasterTahunPelajaran');
  const rows = sh.getDataRange().getValues();
  let found = -1;

  if (rows.length > 1) {
    const statusCol = rows.slice(1).map(function(r, idx) {
      const isTarget = String(r[1] || '') === tahun && String(r[2] || '') === semester;
      if (isTarget) found = idx + 2;
      return [isTarget ? 'AKTIF' : 'NONAKTIF'];
    });
    sh.getRange(2, 4, statusCol.length, 1).setValues(statusCol);
  }

  if (found === -1) {
    sh.appendRow(['TP-' + Utilities.getUuid().slice(0, 8).toUpperCase(), tahun, semester, 'AKTIF', new Date()]);
  }

  logAudit('SET_GLOBAL_TAHUN', auth.email, tahun + ' | ' + semester);
  return { success: true, tahun_pelajaran: tahun, semester: semester };
}

function getUserAcademicPeriod(email) {
  ensureAcademicSchema_();
  const target = String(email || getLoginEmail()).toLowerCase().trim();
  const sh = sheet('SETTING');
  if (!sh || sh.getLastRow() < 2) {
    const g = getGlobalAcademicPeriod_();
    return g || { tahun_pelajaran: getLegacyDefaultYear_(), semester: 'Ganjil' };
  }

  const rows = sh.getDataRange().getValues();
  const header = rows[0].map(function(h) { return String(h || '').toLowerCase().trim(); });
  const idxEmail = header.indexOf('email');
  const idxTaAktif = header.indexOf('tahun_pelajaran_aktif');
  const idxSemAktif = header.indexOf('semester_aktif');
  const idxTahunOld = header.indexOf('tahun');
  const idxSemOld = header.indexOf('semester');

  for (let i = 1; i < rows.length; i++) {
    const em = String(rows[i][idxEmail > -1 ? idxEmail : 0] || '').toLowerCase().trim();
    if (em !== target) continue;

    const tahun = idxTaAktif > -1 ? String(rows[i][idxTaAktif] || '').trim() : '';
    const sem = idxSemAktif > -1 ? String(rows[i][idxSemAktif] || '').trim() : '';
    const tahunOld = idxTahunOld > -1 ? String(rows[i][idxTahunOld] || '').trim() : '';
    const semOld = idxSemOld > -1 ? String(rows[i][idxSemOld] || '').trim() : '';

    return {
      tahun_pelajaran: tahun || tahunOld || (getGlobalAcademicPeriod_() || {}).tahun_pelajaran || getLegacyDefaultYear_(),
      semester: sem || semOld || (getGlobalAcademicPeriod_() || {}).semester || 'Ganjil'
    };
  }

  const g = getGlobalAcademicPeriod_();
  return g || { tahun_pelajaran: getLegacyDefaultYear_(), semester: 'Ganjil' };
}

function setUserAcademicPeriod(tahun, semester) {
  ensureAcademicSchema_();
  const auth = getAuth();
  if (!auth.email || auth.role === 'guest' || String(auth.status || '').toLowerCase() !== 'active') {
    throw new Error('AKSES_DITOLAK');
  }
  tahun = String(tahun || '').trim();
  semester = String(semester || '').trim();
  if (!tahun || !semester) throw new Error('Tahun pelajaran dan semester wajib diisi');

  const sh = sheet('SETTING');
  const rows = sh.getDataRange().getValues();
  const header = rows[0].map(function(h) { return String(h || '').toLowerCase().trim(); });
  const idxEmail = header.indexOf('email');
  const idxTaAktif = header.indexOf('tahun_pelajaran_aktif');
  const idxSemAktif = header.indexOf('semester_aktif');
  if (idxTaAktif === -1 || idxSemAktif === -1) throw new Error('Kolom setting akademik belum tersedia');

  for (let i = 1; i < rows.length; i++) {
    const em = String(rows[i][idxEmail > -1 ? idxEmail : 0] || '').toLowerCase().trim();
    if (em !== auth.email) continue;
    sh.getRange(i + 1, idxTaAktif + 1).setValue(tahun);
    sh.getRange(i + 1, idxSemAktif + 1).setValue(semester);
    invalidateCache_('SETTING');
    invalidateDashboardCache_();
    logAudit('SET_USER_TAHUN', auth.email, tahun + ' | ' + semester);
    trySyncGuruSummaryAfterMutation_(auth.email, 'SET_USER_TAHUN');
    return { success: true, tahun_pelajaran: tahun, semester: semester };
  }

  throw new Error('Data setting guru belum ada');
}

function getKelasDiampuAktifForUser_(email, precomputedPeriod) {
  ensureAcademicSchema_();
  const authEmailNorm = String(email || getLoginEmail()).toLowerCase().trim();
  // Terima period yang sudah dihitung caller (mis. getDashboardAllData)
  // supaya tidak baca ulang sheet SETTING — SETTING sudah dibaca sekali
  // di sana, membaca ulang di sini murni redundan dan memperlambat load.
  const period = precomputedPeriod || getUserAcademicPeriod(authEmailNorm);

  const sh = sheet('GuruMengajar');
  const rows = sh.getDataRange().getValues();
  const kelasSet = new Set();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || '').toLowerCase().trim() !== authEmailNorm) continue;
    if (String(rows[i][2] || '') !== String(period.tahun_pelajaran || '')) continue;
    if (String(rows[i][3] || '').toLowerCase().trim() !== String(period.semester || '').toLowerCase().trim()) continue;
    const kelas = String(rows[i][4] || '').trim();
    if (kelas) kelasSet.add(kelas);
  }
  return Array.from(kelasSet).sort();
}

function getMapelDiampuAktifForUser_(email) {
  ensureAcademicSchema_();
  const authEmailNorm = String(email || getLoginEmail()).toLowerCase().trim();
  const period = getUserAcademicPeriod(authEmailNorm);

  const sh = sheet('GuruMengajar');
  const rows = sh.getDataRange().getValues();
  const mapelSet = new Set();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || '').toLowerCase().trim() !== authEmailNorm) continue;
    if (String(rows[i][2] || '') !== String(period.tahun_pelajaran || '')) continue;
    if (String(rows[i][3] || '').toLowerCase().trim() !== String(period.semester || '').toLowerCase().trim()) continue;
    const mapel = String(rows[i][5] || '').trim();
    if (mapel) mapelSet.add(mapel);
  }
  return Array.from(mapelSet).sort();
}

function saveGuruMengajarSetting(payload) {
  const _t0 = Date.now();
  ensureAcademicSchema_();
  logTiming_('saveGuruMengajarSetting > ensureAcademicSchema_', _t0);
  const auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin' && auth.role !== 'kepsek') throw new Error('AKSES_DITOLAK');

  payload = payload || {};
  const tahun = String(payload.tahun_pelajaran || getUserAcademicPeriod(auth.email).tahun_pelajaran || '').trim();
  const semester = String(payload.semester || getUserAcademicPeriod(auth.email).semester || '').trim();
  const kelasList = Array.isArray(payload.kelas) ? payload.kelas : [];
  const mapelList = Array.isArray(payload.mapel) ? payload.mapel : [];

  const _tSheet = Date.now();
  const sh = sheet('GuruMengajar');
  const rows = sh.getDataRange().getValues();
  logTiming_('saveGuruMengajarSetting > sheet(GuruMengajar)+getDataRange', _tSheet);

  // Hapus baris lama milik guru ini di periode ini. deleteRow() satu-satu
  // memang N panggilan API terpisah, tapi jumlah baris yang dihapus di sini
  // wajar kecil (kombinasi kelas x mapel guru itu sendiri), beda dari
  // masalah penulisan di bawah yang bisa puluhan baris sekaligus.
  const _tDelete = Date.now();
  let deletedCount = 0;
  for (let i = rows.length; i >= 2; i--) {
    const r = rows[i - 1];
    if (String(r[1] || '').toLowerCase().trim() !== auth.email) continue;
    if (String(r[2] || '') !== tahun) continue;
    if (String(r[3] || '').toLowerCase().trim() !== semester.toLowerCase()) continue;
    sh.deleteRow(i);
    deletedCount++;
  }
  logTiming_('saveGuruMengajarSetting > delete ' + deletedCount + ' baris lama', _tDelete);

  const now = new Date();
  if (kelasList.length === 0 && mapelList.length === 0) return { success: true, inserted: 0 };

  const combos = [];
  if (kelasList.length && mapelList.length) {
    kelasList.forEach(function(k) {
      mapelList.forEach(function(m) { combos.push([k, m]); });
    });
  } else if (kelasList.length) {
    kelasList.forEach(function(k) { combos.push([k, '']); });
  } else {
    mapelList.forEach(function(m) { combos.push(['', m]); });
  }

  // Tulis SEMUA kombinasi kelas x mapel dalam SATU panggilan setValues(),
  // bukan appendRow() satu-satu per baris. appendRow() per-baris adalah
  // N round-trip Sheets API terpisah — untuk guru dengan banyak kelas
  // (mis. 10 kelas x beberapa mapel = puluhan baris) ini bisa memakan
  // 20-40+ detik. Batch write ini turunkan jadi 1 round-trip.
  const _tWrite = Date.now();
  const rowsToWrite = combos.map(function(c) {
    return [
      'GM-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      auth.email,
      tahun,
      semester,
      c[0],
      c[1],
      now
    ];
  });
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rowsToWrite.length, 7).setValues(rowsToWrite);
  logTiming_('saveGuruMengajarSetting > batch write ' + combos.length + ' baris', _tWrite);

  logAudit('SAVE_GURU_MENGAJAR', auth.email, tahun + ' | ' + semester + ' | ' + combos.length + ' baris');
  invalidateDashboardCache_();
  const _tSync = Date.now();
  trySyncGuruSummaryAfterMutation_(auth.email, 'SAVE_GURU_MENGAJAR');
  logTiming_('saveGuruMengajarSetting > trySyncGuruSummaryAfterMutation_', _tSync);
  logTiming_('saveGuruMengajarSetting > TOTAL', _t0);
  return { success: true, inserted: combos.length };
}

/**
 * getGuruMengajarForPeriod_(email, tahun, semester)
 * Ambil kombinasi kelas+mapel milik SATU guru untuk SATU periode spesifik.
 */
function getGuruMengajarForPeriod_(email, tahun, semester) {
  ensureAcademicSchema_();
  const targetEmail = String(email || '').toLowerCase().trim();
  const targetTahun = String(tahun || '').trim();
  const targetSem = String(semester || '').toLowerCase().trim();

  const sh = sheet('GuruMengajar');
  const rows = sh.getDataRange().getValues();
  const combos = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || '').toLowerCase().trim() !== targetEmail) continue;
    if (String(rows[i][2] || '') !== targetTahun) continue;
    if (String(rows[i][3] || '').toLowerCase().trim() !== targetSem) continue;
    combos.push({ kelas: String(rows[i][4] || '').trim(), mapel: String(rows[i][5] || '').trim() });
  }
  return combos;
}

/**
 * findLatestGuruMengajarPeriod_(email, excludeTahun, excludeSemester)
 * Cari periode (tahun+semester) TERAKHIR milik guru ini yang punya jadwal
 * mengajar tercatat, selain periode yang dikecualikan (biasanya periode
 * yang baru saja dipilih). Dipakai untuk menawarkan "lanjutkan jadwal
 * sebelumnya" saat guru pindah tahun/semester.
 */
function findLatestGuruMengajarPeriod_(email, excludeTahun, excludeSemester) {
  ensureAcademicSchema_();
  const targetEmail = String(email || '').toLowerCase().trim();
  const exTahun = String(excludeTahun || '').trim();
  const exSem = String(excludeSemester || '').toLowerCase().trim();

  const candidates = [];

  const sh = sheet('GuruMengajar');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || '').toLowerCase().trim() !== targetEmail) continue;
    const tahun = String(rows[i][2] || '').trim();
    const semester = String(rows[i][3] || '').trim();
    if (!tahun || !semester) continue;
    candidates.push({ tahun_pelajaran: tahun, semester: semester, created_at: rows[i][6] ? new Date(rows[i][6]) : null });
  }

  // Ikut sertakan periode yang punya jadwal hari/jam (JADWAL_SEMESTER) walau
  // guru itu belum sempat mengisi Kelas/Mapel Diampu — supaya "lanjutkan
  // jadwal" tetap terdeteksi dari kedua sumber data.
  if (typeof getJadwalSemesterCandidatePeriods_ === 'function') {
    getJadwalSemesterCandidatePeriods_(targetEmail).forEach(function(c) { candidates.push(c); });
  }

  const seen = {};
  let latest = null;
  candidates.forEach(function(c) {
    if (c.tahun_pelajaran === exTahun && c.semester.toLowerCase() === exSem) return;
    const key = c.tahun_pelajaran + '|' + c.semester;
    if (seen[key]) return;
    seen[key] = true;
    if (!latest || (c.created_at && (!latest.created_at || c.created_at > latest.created_at))) {
      latest = c;
    }
  });
  return latest;
}

/**
 * checkGuruMengajarTransition(newTahun, newSemester)
 * Dipanggil FE tepat setelah guru ganti tahun/semester aktif (setUserAcademicPeriod).
 * Memberi tahu apakah periode baru sudah punya jadwal (kelas+mapel diampu)
 * — kalau belum, apakah ada jadwal periode lain milik guru ini yang bisa
 * "dilanjutkan" ke periode baru, supaya FE bisa menawarkan pilihan
 * Lanjutkan vs Mulai Baru (reset).
 */
function checkGuruMengajarTransition(newTahun, newSemester) {
  ensureAcademicSchema_();
  const auth = getAuth();
  if (!auth.email || auth.role === 'guest') throw new Error('AKSES_DITOLAK');

  const tahun = String(newTahun || '').trim();
  const semester = String(newSemester || '').trim();
  if (!tahun || !semester) throw new Error('Tahun & semester wajib diisi');

  const current = getGuruMengajarForPeriod_(auth.email, tahun, semester);
  const currentJadwal = (typeof getJadwalSemesterForPeriod_ === 'function') ? getJadwalSemesterForPeriod_(auth.email, tahun, semester) : [];
  if (current.length || currentJadwal.length) {
    return {
      hasCurrent: true,
      hasPrevious: false,
      kelas: Array.from(new Set(current.map(function(c) { return c.kelas; }).filter(Boolean))),
      mapel: Array.from(new Set(current.map(function(c) { return c.mapel; }).filter(Boolean))),
      jadwal_count: currentJadwal.length
    };
  }

  const prev = findLatestGuruMengajarPeriod_(auth.email, tahun, semester);
  if (!prev) return { hasCurrent: false, hasPrevious: false };

  const prevCombos = getGuruMengajarForPeriod_(auth.email, prev.tahun_pelajaran, prev.semester);
  const prevJadwal = (typeof getJadwalSemesterForPeriod_ === 'function') ? getJadwalSemesterForPeriod_(auth.email, prev.tahun_pelajaran, prev.semester) : [];
  return {
    hasCurrent: false,
    hasPrevious: true,
    previous_period: { tahun_pelajaran: prev.tahun_pelajaran, semester: prev.semester },
    kelas: Array.from(new Set(prevCombos.map(function(c) { return c.kelas; }).filter(Boolean))),
    mapel: Array.from(new Set(prevCombos.map(function(c) { return c.mapel; }).filter(Boolean))),
    total_kombinasi: prevCombos.length,
    jadwal_count: prevJadwal.length
  };
}

/**
 * continueGuruMengajarToPeriod(payload)
 * "Lanjutkan jadwal mengajar" — salin kombinasi kelas+mapel milik guru
 * yang login dari periode sumber ke periode tujuan. Hanya menyentuh baris
 * milik guru yang login sendiri (beda dari cloneAcademicYearWizard yang
 * bersifat school-wide & superadmin-only).
 */
function continueGuruMengajarToPeriod(payload) {
  ensureAcademicSchema_();
  const auth = getAuth();
  if (!auth.email || auth.role === 'guest') throw new Error('AKSES_DITOLAK');

  payload = payload || {};
  const fromTahun = String(payload.from_tahun || '').trim();
  const fromSemester = String(payload.from_semester || '').trim();
  const toTahun = String(payload.to_tahun || '').trim();
  const toSemester = String(payload.to_semester || '').trim();
  if (!fromTahun || !fromSemester || !toTahun || !toSemester) {
    throw new Error('Periode sumber dan tujuan wajib diisi');
  }

  const sourceCombos = getGuruMengajarForPeriod_(auth.email, fromTahun, fromSemester);
  let copiedGM = 0;
  if (sourceCombos.length) {
    const sh = sheet('GuruMengajar');
    const rows = sh.getDataRange().getValues();

    // Hapus dulu baris milik guru ini di periode tujuan (idempoten, hindari
    // duplikat kalau tombol "Lanjutkan" ini sampai terpencet dua kali).
    for (let i = rows.length; i >= 2; i--) {
      const r = rows[i - 1];
      if (String(r[1] || '').toLowerCase().trim() !== auth.email) continue;
      if (String(r[2] || '') !== toTahun) continue;
      if (String(r[3] || '').toLowerCase().trim() !== toSemester.toLowerCase()) continue;
      sh.deleteRow(i);
    }

    const now = new Date();
    sourceCombos.forEach(function(c) {
      sh.appendRow(['GM-' + Utilities.getUuid().slice(0, 8).toUpperCase(), auth.email, toTahun, toSemester, c.kelas, c.mapel, now]);
    });
    copiedGM = sourceCombos.length;
  }

  let copiedJadwal = 0;
  if (typeof continueJadwalSemesterToPeriod_ === 'function') {
    copiedJadwal = continueJadwalSemesterToPeriod_(auth.email, fromTahun, fromSemester, toTahun, toSemester);
  }

  logAudit('CONTINUE_GURU_MENGAJAR', auth.email, fromTahun + '/' + fromSemester + ' -> ' + toTahun + '/' + toSemester + ' | GM=' + copiedGM + ' | Jadwal=' + copiedJadwal);
  invalidateDashboardCache_();
  invalidateCache_('JADWAL_SEMESTER');
  trySyncGuruSummaryAfterMutation_(auth.email, 'CONTINUE_GURU_MENGAJAR');
  return { success: true, copied: copiedGM, copied_jadwal: copiedJadwal };
}

/**
 * resetGuruMengajarForPeriod(tahun, semester)
 * "Reset jadwal mengajar" milik guru yang login untuk SATU periode —
 * hapus semua kombinasi kelas+mapel diampu miliknya di periode itu,
 * supaya guru bisa mulai isi ulang dari kosong. Tidak menyentuh guru lain.
 */
function resetGuruMengajarForPeriod(tahun, semester) {
  ensureAcademicSchema_();
  const auth = getAuth();
  if (!auth.email || auth.role === 'guest') throw new Error('AKSES_DITOLAK');

  const targetTahun = String(tahun || '').trim();
  const targetSem = String(semester || '').trim();
  if (!targetTahun || !targetSem) throw new Error('Tahun & semester wajib diisi');

  const sh = sheet('GuruMengajar');
  const rows = sh.getDataRange().getValues();
  let deleted = 0;
  for (let i = rows.length; i >= 2; i--) {
    const r = rows[i - 1];
    if (String(r[1] || '').toLowerCase().trim() !== auth.email) continue;
    if (String(r[2] || '') !== targetTahun) continue;
    if (String(r[3] || '').toLowerCase().trim() !== targetSem.toLowerCase()) continue;
    sh.deleteRow(i);
    deleted++;
  }

  let deletedJadwal = 0;
  if (typeof resetJadwalSemesterForPeriod_ === 'function') {
    deletedJadwal = resetJadwalSemesterForPeriod_(auth.email, targetTahun, targetSem);
  }

  logAudit('RESET_GURU_MENGAJAR', auth.email, targetTahun + '/' + targetSem + ' | GM=' + deleted + ' | Jadwal=' + deletedJadwal);
  invalidateDashboardCache_();
  invalidateCache_('JADWAL_SEMESTER');
  trySyncGuruSummaryAfterMutation_(auth.email, 'RESET_GURU_MENGAJAR');
  return { success: true, deleted: deleted, deleted_jadwal: deletedJadwal };
}

function getAcademicConfig() {
  ensureAcademicSchema_();
  const auth = getAuth();
  const period = getUserAcademicPeriod(auth.email);
  return {
    years: listAcademicYears(),
    global: getGlobalAcademicPeriod_(),
    user: period,
    kelas_diampu: getKelasDiampuAktifForUser_(auth.email),
    mapel_diampu: getMapelDiampuAktifForUser_(auth.email)
  };
}

function getOrCreateMasterSiswa_(rowObj) {
  ensureAcademicSchema_();
  const sh = sheet('MasterSiswa');
  const rows = sh.getDataRange().getValues();
  const nis = String(rowObj.nis || '').trim();
  const nisn = String(rowObj.nisn || '').trim();

  for (let i = 1; i < rows.length; i++) {
    const rNis = String(rows[i][1] || '').trim();
    const rNisn = String(rows[i][2] || '').trim();
    if ((nis && rNis && rNis === nis) || (nisn && rNisn && rNisn === nisn)) {
      return { id: rows[i][0], created: false };
    }
  }

  const id = 'SIS-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  const now = new Date();
  sh.appendRow([
    id,
    nis,
    nisn,
    rowObj.nama || '',
    rowObj.jk || '',
    rowObj.ttl || '',
    rowObj.alamat || '',
    rowObj.orang_tua || '',
    rowObj.kontak || '',
    rowObj.status || 'AKTIF',
    now,
    now
  ]);
  return { id: id, created: true };
}

function addRiwayatKelas_(siswaId, tahun, semester, kelas, status) {
  ensureAcademicSchema_();
  const sh = sheet('RiwayatKelas');
  const rows = sh.getDataRange().getValues();
  const st = String(status || 'AKTIF').toUpperCase();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || '') !== tahun) continue;
    if (String(rows[i][2] || '').toLowerCase().trim() !== String(semester || '').toLowerCase().trim()) continue;
    if (String(rows[i][3] || '') !== String(siswaId || '')) continue;
    if (String(rows[i][4] || '') !== String(kelas || '')) continue;
    return false;
  }

  const now = new Date();
  sh.appendRow([
    'RK-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
    tahun,
    semester,
    siswaId,
    kelas,
    st,
    now,
    now
  ]);
  return true;
}

function importSiswaBaru(payload) {
  ensureAcademicSchema_();
  const auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin') throw new Error('AKSES_DITOLAK');

  payload = payload || {};
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const tahun = String(payload.tahun_pelajaran || getUserAcademicPeriod(auth.email).tahun_pelajaran || '').trim();
  const semester = String(payload.semester || getUserAcademicPeriod(auth.email).semester || '').trim();

  let createdMaster = 0;
  let createdRiwayat = 0;

  rows.forEach(function(r) {
    const m = getOrCreateMasterSiswa_({
      nis: r.nis,
      nisn: r.nisn,
      nama: r.nama,
      jk: r.jk,
      ttl: r.ttl,
      alamat: r.alamat,
      orang_tua: r.orang_tua,
      kontak: r.kontak,
      status: r.status || 'AKTIF'
    });
    if (m.created) createdMaster++;

    if (r.kelas) {
      if (addRiwayatKelas_(m.id, tahun, semester, r.kelas, r.status || 'AKTIF')) {
        createdRiwayat++;
      }
    }
  });

  invalidateCache_('SISWA');
  return { success: true, created_master: createdMaster, created_riwayat: createdRiwayat };
}

function getSiswaAktifByKelasForUser_(kelas, email) {
  ensureAcademicSchema_();
  const authEmailNorm = String(email || getLoginEmail()).toLowerCase().trim();
  const period = getUserAcademicPeriod(authEmailNorm);

  const gm = sheet('GuruMengajar').getDataRange().getValues().slice(1);
  const hasAccess = gm.some(function(r) {
    return String(r[1] || '').toLowerCase().trim() === authEmailNorm &&
      String(r[2] || '') === String(period.tahun_pelajaran || '') &&
      String(r[3] || '').toLowerCase().trim() === String(period.semester || '').toLowerCase().trim() &&
      String(r[4] || '').trim() === String(kelas || '').trim();
  });

  if (!hasAccess) return [];

  const riwayat = sheet('RiwayatKelas').getDataRange().getValues().slice(1).filter(function(r) {
    return String(r[1] || '') === String(period.tahun_pelajaran || '') &&
      String(r[2] || '').toLowerCase().trim() === String(period.semester || '').toLowerCase().trim() &&
      String(r[4] || '').trim() === String(kelas || '').trim() &&
      !['ALUMNI', 'MUTASI_KELUAR'].includes(String(r[5] || '').toUpperCase());
  });

  if (!riwayat.length) return [];

  const master = sheet('MasterSiswa').getDataRange().getValues().slice(1);
  const byId = {};
  master.forEach(function(m) { byId[String(m[0])] = m; });

  return riwayat.map(function(r, idx) {
    const s = byId[String(r[3])] || [];
    return {
      no_absen: idx + 1,
      nis: s[1] || '',
      nama: s[3] || '-',
      jk: s[4] || '-'
    };
  }).sort(function(a, b) { return String(a.nama).localeCompare(String(b.nama)); });
}

function getMasterSiswaMapByNis_() {
  ensureAcademicSchema_();
  const rows = sheet('MasterSiswa').getDataRange().getValues().slice(1);
  const map = {};
  rows.forEach(function(r) {
    const nis = String(r[1] || '').trim();
    if (!nis) return;
    map[nis] = {
      id: r[0],
      nis: nis,
      nisn: r[2] || '',
      nama: r[3] || '-',
      jk: r[4] || '-',
      status: String(r[9] || 'AKTIF').toUpperCase()
    };
  });
  return map;
}

function getNamaSiswaByNis_(nis) {
  const key = String(nis || '').trim();
  if (!key) return '-';
  const map = getMasterSiswaMapByNis_();
  return map[key] ? map[key].nama : key;
}

function cloneAcademicYearWizard(sourceTahun, destinationTahun) {
  ensureAcademicSchema_();
  const auth = getAuth();
  // School-wide: menyalin GuruMengajar SEMUA guru sekaligus (bukan hanya
  // guru yang login), jadi wajib superadmin. Untuk guru biasa yang cuma
  // mau melanjutkan jadwal mengajarnya sendiri, pakai
  // continueGuruMengajarToPeriod() (scoped ke email sendiri).
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');

  sourceTahun = String(sourceTahun || '').trim();
  destinationTahun = String(destinationTahun || '').trim();
  if (!sourceTahun || !destinationTahun) throw new Error('Source dan destination wajib diisi');
  if (sourceTahun === destinationTahun) throw new Error('Source dan destination tidak boleh sama');

  const sems = ['Ganjil', 'Genap'];
  sems.forEach(function(sem) { ensureAcademicYearExists_(destinationTahun, sem, false); });

  const shGM = sheet('GuruMengajar');
  const rows = shGM.getDataRange().getValues();
  const now = new Date();
  let clonedGM = 0;

  const existingKey = {};
  for (let i = 1; i < rows.length; i++) {
    const k = [rows[i][1], rows[i][2], rows[i][3], rows[i][4], rows[i][5]].join('|');
    existingKey[k] = true;
  }

  for (let i = 1; i < rows.length; i++) {
    const guru = String(rows[i][1] || '').toLowerCase().trim();
    const tahun = String(rows[i][2] || '');
    const semester = String(rows[i][3] || '');
    const kelas = String(rows[i][4] || '');
    const mapel = String(rows[i][5] || '');
    if (tahun !== sourceTahun) continue;

    const newKey = [guru, destinationTahun, semester, kelas, mapel].join('|');
    if (existingKey[newKey]) continue;

    shGM.appendRow([
      'GM-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      guru,
      destinationTahun,
      semester,
      kelas,
      mapel,
      now
    ]);
    existingKey[newKey] = true;
    clonedGM++;
  }

  logAudit('CLONE_TAHUN', auth.email, sourceTahun + ' -> ' + destinationTahun + ' | GM=' + clonedGM);
  return { success: true, source: sourceTahun, destination: destinationTahun, cloned_guru_mengajar: clonedGM };
}

function promoteSiswaWizard(payload) {
  ensureAcademicSchema_();
  const auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin') throw new Error('AKSES_DITOLAK');

  payload = payload || {};
  const fromYear = String(payload.from_tahun || '').trim();
  const fromSem  = String(payload.from_semester || '').trim();
  const toYear   = String(payload.to_tahun || '').trim();
  const toSem    = String(payload.to_semester || '').trim();
  const mapping  = Array.isArray(payload.mapping) ? payload.mapping : [];
  // Kelas baru di tahun tujuan yang tidak berasal dari promosi (siswa angkatan baru)
  const newKelas  = Array.isArray(payload.new_kelas)
    ? payload.new_kelas.map(function(k){ return String(k||'').trim(); }).filter(Boolean)
    : [];
  // Mapel yang akan diajar di tahun baru (opsional; default: sama dengan tahun lama)
  const mapelBaru = Array.isArray(payload.mapel_baru)
    ? payload.mapel_baru.map(function(m){ return String(m||'').trim(); }).filter(Boolean)
    : null;

  if (!fromYear || !fromSem || !toYear || !toSem) {
    throw new Error('Periode lama dan baru wajib diisi');
  }

  const mapObj = {};
  mapping.forEach(function(m) { mapObj[String(m.from || '').trim()] = String(m.to || '').trim(); });

  // Batasi promosi hanya pada kelas yang diajar guru ini di periode lama
  const guruKelasLama = new Set(
    getGuruMengajarForPeriod_(auth.email, fromYear, fromSem)
      .map(function(c){ return c.kelas; }).filter(Boolean)
  );

  const riwayat = sheet('RiwayatKelas').getDataRange().getValues().slice(1).filter(function(r) {
    return String(r[1] || '') === fromYear &&
           String(r[2] || '').toLowerCase().trim() === fromSem.toLowerCase() &&
           guruKelasLama.has(String(r[4] || '').trim());
  });

  let processed = 0;
  riwayat.forEach(function(r) {
    const fromKelas = String(r[4] || '').trim();
    const toKelas   = mapObj[fromKelas];
    if (!toKelas) return;

    var status = 'AKTIF';
    var kelasTujuan = toKelas;
    if (String(toKelas).toUpperCase() === 'ALUMNI') {
      status = 'ALUMNI';
      kelasTujuan = 'ALUMNI';
    }
    if (addRiwayatKelas_(r[3], toYear, toSem, kelasTujuan, status)) processed++;
  });

  // Fallback: promosi di sheet SISWA lama, untuk guru yang rosternya belum
  // pernah masuk RiwayatKelas. Sejak kolom tahun_pelajaran ditambahkan,
  // promosi TIDAK LAGI menimpa baris lama in-place — sebaliknya baris baru
  // ditambahkan untuk tahun tujuan, supaya riwayat kelas 7/8/9 (atau
  // jenjang lain) tiap siswa tetap tersimpan lengkap per tahun ajaran,
  // sama seperti promoteSiswaBinaan() di Wali.js.
  let processedLegacy = 0;
  try {
    const shSiswaProm = sheet('SISWA');
    if (shSiswaProm && mapping.length) {
      ensureSiswaTahunKolom_();
      const emailNormProm = String(auth.email).toLowerCase().trim();

      const targetKelasSet = new Set(
        Object.keys(mapObj)
          .map(function(k) { return mapObj[k]; })
          .filter(function(t) { return String(t).toUpperCase() !== 'ALUMNI'; })
      );

      // Hapus dulu baris promosi LAMA guru ini di tahun TUJUAN (idempoten
      // kalau wizard ini sampai dijalankan dua kali untuk periode yang
      // sama) — hanya untuk kelas tujuan yang memang bagian dari mapping
      // kali ini, supaya kelas lain di tahun yang sama tidak ikut terhapus.
      let rows = shSiswaProm.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        const r = rows[i];
        if (String(r[5] || '').toLowerCase().trim() !== emailNormProm) continue;
        if (String(r[SISWA_TAHUN_COL_] || '').trim() !== toYear) continue;
        if (!targetKelasSet.has(String(r[0] || '').trim())) continue;
        shSiswaProm.deleteRow(i + 1);
      }

      // Baca ulang, lalu tambahkan baris baru untuk siswa di tahun ASAL
      // yang kelasnya termasuk mapping. Baris lama TIDAK diubah/dihapus.
      rows = shSiswaProm.getDataRange().getValues();
      const newRows = [];
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (String(r[5] || '').toLowerCase().trim() !== emailNormProm) continue;
        // Baris lama tanpa tahun_pelajaran (sebelum migrasi ini) dianggap
        // milik periode manapun, supaya tetap bisa dipromosikan.
        const rowTahun = String(r[SISWA_TAHUN_COL_] || '').trim();
        if (rowTahun && rowTahun !== fromYear) continue;

        const target = mapObj[String(r[0] || '').trim()];
        if (!target) continue;

        const kelasTujuan = String(target).toUpperCase() === 'ALUMNI' ? 'ALUMNI' : target;
        newRows.push([kelasTujuan, r[1], r[2], r[3], r[4], auth.email, toYear]);
        processedLegacy++;
      }
      if (newRows.length) {
        shSiswaProm.getRange(shSiswaProm.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
        invalidateCache_('SISWA');
      }
    }
  } catch (e) {
    try { logError_('PROMOTE_SISWA_LEGACY', e); } catch (e2) {}
  }

  // Pastikan tahun tujuan terdaftar di MasterTahunPelajaran
  ensureAcademicYearExists_(toYear, toSem, false);

  // Kumpulkan semua kelas yang akan diajar di tahun baru:
  // - kelas tujuan dari mapping (yang bukan ALUMNI)
  // - kelas baru yang disuplai guru (angkatan baru)
  const destKelasSet = new Set();
  mapping.forEach(function(m) {
    const to = String(m.to || '').trim();
    if (to && to.toUpperCase() !== 'ALUMNI') destKelasSet.add(to);
  });
  newKelas.forEach(function(k) { destKelasSet.add(k); });
  const allNewKelas = Array.from(destKelasSet);

  // Perbarui GuruMengajar untuk periode tujuan
  var guruMengajarUpdated = 0;
  if (allNewKelas.length) {
    const oldCombos = getGuruMengajarForPeriod_(auth.email, fromYear, fromSem);
    const mapelToUse = (mapelBaru !== null)
      ? mapelBaru
      : Array.from(new Set(oldCombos.map(function(c){ return c.mapel; }).filter(Boolean)));

    const shGM = sheet('GuruMengajar');
    const gmRows = shGM.getDataRange().getValues();
    // Hapus entri guru ini di periode tujuan agar tidak duplikat
    for (var i = gmRows.length; i >= 2; i--) {
      const r = gmRows[i - 1];
      if (String(r[1]||'').toLowerCase().trim() !== auth.email) continue;
      if (String(r[2]||'') !== toYear) continue;
      if (String(r[3]||'').toLowerCase().trim() !== toSem.toLowerCase()) continue;
      shGM.deleteRow(i);
    }

    const now = new Date();
    const rowsToWrite = [];
    allNewKelas.forEach(function(k) {
      if (mapelToUse.length) {
        mapelToUse.forEach(function(m) {
          rowsToWrite.push(['GM-' + Utilities.getUuid().slice(0,8).toUpperCase(), auth.email, toYear, toSem, k, m, now]);
        });
      } else {
        rowsToWrite.push(['GM-' + Utilities.getUuid().slice(0,8).toUpperCase(), auth.email, toYear, toSem, k, '', now]);
      }
    });
    if (rowsToWrite.length) {
      shGM.getRange(shGM.getLastRow() + 1, 1, rowsToWrite.length, 7).setValues(rowsToWrite);
      guruMengajarUpdated = rowsToWrite.length;
    }
  }

  logAudit('PROMOSI_SISWA', auth.email,
    fromYear + '/' + fromSem + ' -> ' + toYear + '/' + toSem +
    ' | Siswa(riwayat)=' + processed + ' | Siswa(legacy)=' + processedLegacy + ' | GM=' + guruMengajarUpdated);
  invalidateDashboardCache_();
  trySyncGuruSummaryAfterMutation_(auth.email, 'PROMOSI_SISWA');
  return {
    success: true,
    processed: processed + processedLegacy,
    processed_riwayat: processed,
    processed_legacy: processedLegacy,
    guru_mengajar_updated: guruMengajarUpdated,
    all_new_kelas: allNewKelas,
    kelas_perlu_siswa_baru: newKelas
  };
}

/**
 * getKelasListForPeriod(tahun, semester)
 * Wrapper tipis untuk UI Promosi Siswa — daftar kelas yang diampu GURU
 * YANG LOGIN di satu periode tertentu, supaya guru tinggal pilih dari
 * kelas yang benar-benar ia ajar (bukan mengetik nama kelas manual).
 */
function getKelasListForPeriod(tahun, semester) {
  ensureAcademicSchema_();
  const auth = getAuth();
  if (!auth.email || auth.role === 'guest') throw new Error('AKSES_DITOLAK');
  const combos = getGuruMengajarForPeriod_(auth.email, tahun, semester);
  return Array.from(new Set(combos.map(function(c) { return c.kelas; }).filter(Boolean))).sort();
}

/**
 * getAcademicHistorySummary(tahun, semester)
 * Rekap BACA-SAJA untuk periode manapun milik guru yang login — kelas
 * diampu, mapel diampu, jumlah siswa aktif per kelas, jumlah entri
 * jadwal hari/jam. Tidak menyentuh/menulis apapun, dan tidak mengubah
 * periode aktif guru (beda dari setUserAcademicPeriod).
 */
function getAcademicHistorySummary(tahun, semester) {
  ensureAcademicSchema_();
  const auth = getAuth();
  if (!auth.email || auth.role === 'guest') throw new Error('AKSES_DITOLAK');

  tahun = String(tahun || '').trim();
  semester = String(semester || '').trim();
  if (!tahun || !semester) throw new Error('Tahun & semester wajib diisi');

  const combos = getGuruMengajarForPeriod_(auth.email, tahun, semester);
  let kelasList = Array.from(new Set(combos.map(function(c) { return c.kelas; }).filter(Boolean))).sort();
  const mapelList = Array.from(new Set(combos.map(function(c) { return c.mapel; }).filter(Boolean))).sort();
  const jadwal = (typeof getJadwalSemesterForPeriod_ === 'function')
    ? getJadwalSemesterForPeriod_(auth.email, tahun, semester) : [];

  const riwayat = sheet('RiwayatKelas').getDataRange().getValues().slice(1).filter(function(r) {
    return String(r[1] || '') === tahun &&
      String(r[2] || '').toLowerCase().trim() === semester.toLowerCase() &&
      kelasList.indexOf(String(r[4] || '').trim()) > -1;
  });
  const siswaPerKelas = {};
  kelasList.forEach(function(k) { siswaPerKelas[k] = 0; });
  riwayat.forEach(function(r) {
    const k = String(r[4] || '').trim();
    const status = String(r[5] || '').toUpperCase();
    if (status === 'ALUMNI' || status === 'MUTASI_KELUAR') return;
    if (siswaPerKelas.hasOwnProperty(k)) siswaPerKelas[k]++;
  });

  // Fallback: guru ini belum pernah isi "Kelas Diampu" (GuruMengajar) untuk
  // periode ini — ambil langsung dari roster SISWA lama (sheet legacy tanpa
  // kolom tahun_pelajaran/semester), sama seperti getKelas() di Jurnal.js
  // saat GuruMengajar kosong. Tanpa ini, Promosi Siswa tampak "tidak ada
  // kelas" walau guru punya siswa aktif, karena rosternya memang belum
  // pernah dicatat lewat sistem multi-tahun.
  let fromLegacySiswa = false;
  if (kelasList.length === 0) {
    const shSiswaHist = sheet('SISWA');
    if (shSiswaHist && shSiswaHist.getLastRow() > 1) {
      const emailNormHist = String(auth.email).toLowerCase().trim();
      const siswaRows = shSiswaHist.getDataRange().getValues().slice(1);
      const counts = {};
      siswaRows.forEach(function(r) {
        if (String(r[5] || '').toLowerCase().trim() !== emailNormHist) return;
        // Periode yang diminta di sini (tahun/semester parameter fungsi ini)
        // bisa periode manapun, bukan cuma periode aktif — jadi dicocokkan
        // ke kolom tahun_pelajaran baris ini, bukan setting global.
        if (!_siswaRowMatchesPeriode_(r, tahun)) return;
        const k = String(r[0] || '').trim();
        if (!k) return;
        counts[k] = (counts[k] || 0) + 1;
      });
      kelasList = Object.keys(counts).sort();
      kelasList.forEach(function(k) { siswaPerKelas[k] = counts[k]; });
      fromLegacySiswa = kelasList.length > 0;
    }
  }

  return {
    tahun_pelajaran: tahun,
    semester: semester,
    kelas_diampu: kelasList,
    mapel_diampu: mapelList,
    siswa_per_kelas: siswaPerKelas,
    jadwal_count: jadwal.length,
    from_legacy_siswa: fromLegacySiswa
  };
}

