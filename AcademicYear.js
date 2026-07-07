/**
 * AcademicYear.js
 * Backward-compatible multi-year academic data layer.
 */

function ensureAcademicSchema_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('ACADEMIC_SCHEMA_READY');
  if (hit) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const hit2 = cache.get('ACADEMIC_SCHEMA_READY');
    if (hit2) return;

    const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || SPREADSHEET_ID;
    const ss = SpreadsheetApp.openById(ssId);

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

    cache.put('ACADEMIC_SCHEMA_READY', '1', 3600);
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
  if (cache.get('ACADEMIC_MIGRATED')) return;

  const ss = getSpreadsheet_();
  const shSiswa = ss.getSheetByName('SISWA');
  const shMaster = ss.getSheetByName('MasterSiswa');
  const shRiwayat = ss.getSheetByName('RiwayatKelas');
  const shSet = ss.getSheetByName('SETTING');
  const shTahun = ss.getSheetByName('MasterTahunPelajaran');
  const shJurnal = ss.getSheetByName('JURNAL');

  const defaultTahun = getLegacyDefaultYear_();
  const defaultSemester = 'Ganjil';

  ensureAcademicYearExists_(defaultTahun, defaultSemester, true);

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

    const siswaRows = shSiswa.getDataRange().getValues();
    const now = new Date();
    for (let i = 1; i < siswaRows.length; i++) {
      const kelas = String(siswaRows[i][0] || '').trim();
      const nis = String(siswaRows[i][2] || '').trim();
      const nama = String(siswaRows[i][3] || '').trim();
      const jk = String(siswaRows[i][4] || '').trim();
      if (!nis || !nama) continue;

      let siswaId = masterByNis['NIS:' + nis] || '';
      if (!siswaId) {
        siswaId = 'SIS-' + Utilities.getUuid().slice(0, 8).toUpperCase();
        shMaster.appendRow([siswaId, nis, '', nama, jk, '', '', '', '', 'AKTIF', now, now]);
        masterByNis['NIS:' + nis] = siswaId;
      }

      const k = [defaultTahun, defaultSemester, siswaId, kelas].join('|');
      if (!riwayatKey[k]) {
        shRiwayat.appendRow([
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
  }

  if (shSet && shSet.getLastRow() > 1) {
    const all = shSet.getDataRange().getValues();
    const header = all[0].map(function(h) { return String(h || '').toLowerCase().trim(); });
    const idxTahun = header.indexOf('tahun');
    const idxSem = header.indexOf('semester');
    const idxTaAktif = header.indexOf('tahun_pelajaran_aktif');
    const idxSemAktif = header.indexOf('semester_aktif');

    for (let i = 1; i < all.length; i++) {
      const tahunOld = idxTahun > -1 ? String(all[i][idxTahun] || '').trim() : '';
      const semOld = idxSem > -1 ? String(all[i][idxSem] || '').trim() : '';
      if (idxTaAktif > -1 && !String(all[i][idxTaAktif] || '').trim()) {
        shSet.getRange(i + 1, idxTaAktif + 1).setValue(tahunOld || defaultTahun);
      }
      if (idxSemAktif > -1 && !String(all[i][idxSemAktif] || '').trim()) {
        shSet.getRange(i + 1, idxSemAktif + 1).setValue(semOld || defaultSemester);
      }
    }
  }

  if (shJurnal && shJurnal.getLastRow() > 1) {
    const rows = shJurnal.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][18]) shJurnal.getRange(i + 1, 19).setValue(defaultTahun);
      if (!rows[i][13]) shJurnal.getRange(i + 1, 14).setValue(defaultSemester);
    }
  }

  cache.put('ACADEMIC_MIGRATED', '1', 21600);
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

  for (let i = 1; i < rows.length; i++) {
    const isTarget = String(rows[i][1] || '') === tahun && String(rows[i][2] || '') === semester;
    sh.getRange(i + 1, 4).setValue(isTarget ? 'AKTIF' : 'NONAKTIF');
    if (isTarget) found = i + 1;
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
    return { success: true, tahun_pelajaran: tahun, semester: semester };
  }

  throw new Error('Data setting guru belum ada');
}

function getKelasDiampuAktifForUser_(email) {
  ensureAcademicSchema_();
  const authEmailNorm = String(email || getLoginEmail()).toLowerCase().trim();
  const period = getUserAcademicPeriod(authEmailNorm);

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
  ensureAcademicSchema_();
  const auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin' && auth.role !== 'kepsek') throw new Error('AKSES_DITOLAK');

  payload = payload || {};
  const tahun = String(payload.tahun_pelajaran || getUserAcademicPeriod(auth.email).tahun_pelajaran || '').trim();
  const semester = String(payload.semester || getUserAcademicPeriod(auth.email).semester || '').trim();
  const kelasList = Array.isArray(payload.kelas) ? payload.kelas : [];
  const mapelList = Array.isArray(payload.mapel) ? payload.mapel : [];

  const sh = sheet('GuruMengajar');
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length; i >= 2; i--) {
    const r = rows[i - 1];
    if (String(r[1] || '').toLowerCase().trim() !== auth.email) continue;
    if (String(r[2] || '') !== tahun) continue;
    if (String(r[3] || '').toLowerCase().trim() !== semester.toLowerCase()) continue;
    sh.deleteRow(i);
  }

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

  combos.forEach(function(c) {
    sh.appendRow([
      'GM-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      auth.email,
      tahun,
      semester,
      c[0],
      c[1],
      now
    ]);
  });

  logAudit('SAVE_GURU_MENGAJAR', auth.email, tahun + ' | ' + semester + ' | ' + combos.length + ' baris');
  invalidateDashboardCache_();
  return { success: true, inserted: combos.length };
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
  if (auth.role !== 'admin' && auth.role !== 'superadmin') throw new Error('AKSES_DITOLAK');

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
  const fromSem = String(payload.from_semester || '').trim();
  const toYear = String(payload.to_tahun || '').trim();
  const toSem = String(payload.to_semester || '').trim();
  const mapping = Array.isArray(payload.mapping) ? payload.mapping : [];

  if (!fromYear || !fromSem || !toYear || !toSem) {
    throw new Error('Periode lama dan baru wajib diisi');
  }

  const mapObj = {};
  mapping.forEach(function(m) { mapObj[String(m.from || '').trim()] = String(m.to || '').trim(); });

  const riwayat = sheet('RiwayatKelas').getDataRange().getValues().slice(1).filter(function(r) {
    return String(r[1] || '') === fromYear && String(r[2] || '').toLowerCase().trim() === fromSem.toLowerCase();
  });

  let processed = 0;
  riwayat.forEach(function(r) {
    const fromKelas = String(r[4] || '').trim();
    const toKelas = mapObj[fromKelas];
    if (!toKelas) return;

    let status = 'AKTIF';
    let kelasTujuan = toKelas;
    if (String(toKelas).toUpperCase() === 'ALUMNI') {
      status = 'ALUMNI';
      kelasTujuan = 'ALUMNI';
    }

    if (addRiwayatKelas_(r[3], toYear, toSem, kelasTujuan, status)) processed++;
  });

  logAudit('PROMOSI_SISWA', auth.email, fromYear + '/' + fromSem + ' -> ' + toYear + '/' + toSem + ' | ' + processed);
  return { success: true, processed: processed };
}
