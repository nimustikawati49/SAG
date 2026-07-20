/**
 * ModulAjar.js — CRUD Modul Ajar + Upload Drive + Relasi Jadwal
 * Sheet: ModulAjar          → id, mapel, kelas, bab, judul, file_url, deskripsi, created_at, email
 * Sheet: Relasi_Modul_Jadwal → id, id_modul, id_jadwal
 */

const FOLDER_MODUL_AJAR = 'MODUL_AJAR_GURU';

// =========================================================
// SHEET SETUP
// =========================================================

function ensureModulAjarSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('ModulAjar');
  if (!sh) {
    sh = ss.insertSheet('ModulAjar');
    sh.appendRow(['id', 'mapel', 'kelas', 'bab', 'judul', 'file_url', 'deskripsi', 'created_at', 'email']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureRelasiModulJadwalSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('Relasi_Modul_Jadwal');
  if (!sh) {
    sh = ss.insertSheet('Relasi_Modul_Jadwal');
    sh.appendRow(['id', 'id_modul', 'id_jadwal']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// =========================================================
// AUTO INCREMENT
// =========================================================

function nextModulId_() {
  const sh = ensureModulAjarSheet_();
  const vals = sh.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < vals.length; i++) {
    const v = Number(vals[i][0]);
    if (v > max) max = v;
  }
  return max + 1;
}

function nextRelasiId_() {
  const sh = ensureRelasiModulJadwalSheet_();
  const vals = sh.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < vals.length; i++) {
    const v = Number(vals[i][0]);
    if (v > max) max = v;
  }
  return max + 1;
}

// =========================================================
// CRUD MODUL AJAR
// =========================================================

/**
 * Ambil semua modul milik guru yang sedang login (dengan cache 30s).
 * Superadmin dapat melihat semua modul.
 */
function getAllModul() {
  assertLicenseActive();
  assertMinTier_('PRO');
  const auth = getAuth();

  // Gunakan cache per email agar tidak konflik antar guru
  const cacheKey = 'MODUL_' + auth.email;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  const sh = ensureModulAjarSheet_();
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];

  const headers = vals[0];
  const result = [];

  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][0] && !vals[i][1]) continue;
    const row = {};
    headers.forEach((h, j) => {
      const v = vals[i][j];
      row[h] = v instanceof Date ? v.toISOString() : String(v === null || v === undefined ? '' : v);
    });
    if (auth.role !== 'superadmin') {
      const rowEmail = String(row.email || '').toLowerCase().trim();
      if (rowEmail && rowEmail !== auth.email) continue;
    }
    result.push(row);
  }

  cache.put(cacheKey, JSON.stringify(result), 30);
  return result;
}

function invalidateModulCache_() {
  try {
    const auth = getAuth();
    CacheService.getScriptCache().remove('MODUL_' + auth.email);
  } catch(e) {}
}

/**
 * Filter modul berdasarkan mapel dan kelas.
 */
function getModulByMapelKelas(mapel, kelas) {
  assertLicenseActive();
  return getAllModul().filter(m => m.mapel === mapel && m.kelas === kelas);
}

/**
 * Tambah modul baru. Bila ada file (base64), upload ke Drive lebih dulu.
 * data = { mapel, kelas, bab, judul, deskripsi, file_base64, file_name, file_mime }
 */
function addModul(data) {
  assertMinTier_('PRO');
  assertLicenseActive();
  const auth = getAuth();

  if (!data.mapel || !data.kelas || !data.bab || !data.judul) {
    throw new Error('Mapel, kelas, bab, dan judul wajib diisi');
  }

  // Cek duplikat (judul + kelas + bab)
  const existing = getAllModul();
  const isDuplicate = existing.some(function(m) {
    return String(m.judul).toLowerCase().trim() === String(data.judul).toLowerCase().trim()
        && m.kelas === data.kelas
        && String(m.bab).trim() === String(data.bab).trim();
  });
  if (isDuplicate) {
    throw new Error('Modul dengan judul, kelas, dan bab yang sama sudah ada');
  }

  let file_url = '';
  if (data.file_base64 && data.file_name) {
    const uploaded = uploadModulFileToDrive(data.file_base64, data.file_name, data.file_mime);
    file_url = uploaded.url;
  }

  const sh = ensureModulAjarSheet_();
  const id = nextModulId_();
  sh.appendRow([
    id,
    data.mapel,
    data.kelas,
    data.bab,
    data.judul,
    file_url,
    data.deskripsi || '',
    new Date(),
    auth.email
  ]);

  invalidateModulCache_();
  logAudit('ADD_MODUL', auth.email, data.judul + ' | ' + data.kelas);
  return { status: true, id: id };
}

/**
 * Update modul ajar.
 * data = { mapel, kelas, bab, judul, deskripsi, file_base64?, file_name?, file_mime? }
 */
function updateModul(id, data) {
  assertLicenseActive();
  const auth = getAuth();

  if (!data.mapel || !data.kelas || !data.bab || !data.judul) {
    throw new Error('Mapel, kelas, bab, dan judul wajib diisi');
  }

  const sh = ensureModulAjarSheet_();
  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][0]) !== Number(id)) continue;
    const ownerEmail = String(vals[i][8] || '').toLowerCase().trim();
    if (auth.role !== 'superadmin' && ownerEmail !== auth.email) {
      throw new Error('AKSES_DITOLAK');
    }

    // Jika ada file baru, upload dulu
    let file_url = String(vals[i][5] || ''); // pertahankan file lama
    if (data.file_base64 && data.file_name) {
      const uploaded = uploadModulFileToDrive(data.file_base64, data.file_name, data.file_mime);
      file_url = uploaded.url;
    }

    const rowNum = i + 1;
    sh.getRange(rowNum, 2).setValue(data.mapel);
    sh.getRange(rowNum, 3).setValue(data.kelas);
    sh.getRange(rowNum, 4).setValue(data.bab);
    sh.getRange(rowNum, 5).setValue(data.judul);
    sh.getRange(rowNum, 6).setValue(file_url);
    sh.getRange(rowNum, 7).setValue(data.deskripsi || '');

    invalidateModulCache_();
    logAudit('UPDATE_MODUL', auth.email, 'id=' + id + ' | ' + data.judul);
    return true;
  }
  throw new Error('Modul tidak ditemukan');
}

/**
 * Hapus modul berdasarkan id.
 * Hanya pemilik atau superadmin yang bisa menghapus.
 */
function deleteModul(id) {
  assertLicenseActive();
  const auth = getAuth();
  const sh = ensureModulAjarSheet_();
  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][0]) !== Number(id)) continue;
    const ownerEmail = String(vals[i][8] || '').toLowerCase().trim();
    if (auth.role !== 'superadmin' && ownerEmail !== auth.email) {
      throw new Error('AKSES_DITOLAK');
    }
    sh.deleteRow(i + 1);
    invalidateModulCache_();
    logAudit('DELETE_MODUL', auth.email, 'id=' + id);
    return true;
  }
  throw new Error('Modul tidak ditemukan');
}

// =========================================================
// UPLOAD FILE KE GOOGLE DRIVE
// =========================================================

/**
 * Menerima file dalam format base64, simpan ke folder MODUL_AJAR_GURU di Drive.
 * Kembalikan { url, id, name }.
 */
function uploadModulFileToDrive(base64Data, fileName, mimeType) {
  if (!base64Data || !fileName) throw new Error('Data file tidak lengkap');

  // Sanitasi nama file
  const safeName = fileName.replace(/[^a-zA-Z0-9._\- ]/g, '_');

  // Cari atau buat folder
  const folderIt = DriveApp.getFoldersByName(FOLDER_MODUL_AJAR);
  const folder = folderIt.hasNext()
    ? folderIt.next()
    : DriveApp.createFolder(FOLDER_MODUL_AJAR);

  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', safeName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    url: file.getUrl(),
    id:  file.getId(),
    name: file.getName()
  };
}

// =========================================================
// RELASI MODUL — JADWAL
// =========================================================

/**
 * Hubungkan modul ke jadwal mengajar.
 */
function assignModulToJadwal(id_modul, id_jadwal) {
  assertLicenseActive();
  const sh = ensureRelasiModulJadwalSheet_();
  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][1]) === Number(id_modul) && Number(vals[i][2]) === Number(id_jadwal)) {
      throw new Error('Modul sudah terhubung ke jadwal ini');
    }
  }

  const id = nextRelasiId_();
  sh.appendRow([id, Number(id_modul), Number(id_jadwal)]);
  logAudit('ASSIGN_MODUL_JADWAL', getLoginEmail(), 'modul=' + id_modul + ' jadwal=' + id_jadwal);
  return { status: true, id: id };
}

/**
 * Ambil semua modul yang terhubung ke jadwal tertentu (by row id jadwal).
 */
function getModulByJadwal(id_jadwal) {
  assertLicenseActive();
  const shRelasi = ensureRelasiModulJadwalSheet_();
  const shModul  = ensureModulAjarSheet_();

  const relasiVals = shRelasi.getDataRange().getValues();
  const modulVals  = shModul.getDataRange().getValues();
  if (modulVals.length <= 1) return [];

  const modulHeaders = modulVals[0];
  const modulMap = {};
  for (let i = 1; i < modulVals.length; i++) {
    if (!modulVals[i][0]) continue;
    const row = {};
    modulHeaders.forEach((h, j) => {
      const v = modulVals[i][j];
      row[h] = v instanceof Date ? v.toISOString() : String(v === null || v === undefined ? '' : v);
    });
    modulMap[Number(row.id)] = row;
  }

  const result = [];
  for (let i = 1; i < relasiVals.length; i++) {
    if (!relasiVals[i][0]) continue;
    if (Number(relasiVals[i][2]) !== Number(id_jadwal)) continue;
    const m = modulMap[Number(relasiVals[i][1])];
    if (m) result.push(Object.assign({}, m, { relasi_id: Number(relasiVals[i][0]) }));
  }
  return result;
}

/**
 * Lepas relasi modul dari jadwal.
 */
function removeModulFromJadwal(relasi_id) {
  assertLicenseActive();
  const sh = ensureRelasiModulJadwalSheet_();
  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][0]) !== Number(relasi_id)) continue;
    sh.deleteRow(i + 1);
    logAudit('REMOVE_MODUL_JADWAL', getLoginEmail(), 'relasi_id=' + relasi_id);
    return true;
  }
  throw new Error('Relasi tidak ditemukan');
}

// =========================================================
// JADWAL WITH ROW ID (dipakai untuk relasi modul–jadwal)
// =========================================================

// =========================================================
// MULTI-MAPEL: Ambil daftar mapel dari setting guru
// =========================================================

/**
 * Kembalikan array mapel yang diajar oleh guru yang sedang login.
 * Mapel disimpan sebagai comma-separated di sheet SETTING.
 */
function getMapelListGuru() {
  const setting = getSetting();
  const raw = String(setting.mata_pelajaran || '').trim();
  if (!raw) return [];
  return raw.split(/,\s*/).map(function(m) { return m.trim(); }).filter(Boolean);
}

// =========================================================
// JADWAL WITH ROW ID
// =========================================================

/**
 * Sama seperti getJadwalMengajar() namun menyertakan id baris sebagai id jadwal.
 */
function getJadwalWithId() {
  const auth = getAuth();
  if (!auth || !auth.email) return [];

  const emailLogin    = String(auth.email).toLowerCase().trim();
  const setting       = getSetting();
  const semesterAktif = String(setting.semester || '').trim().toLowerCase();
  const tahunAktif     = String(setting.tahun_pelajaran || '').trim();
  const sh            = getJadwalSheet_();
  if (!sh) return [];

  const tIdx = (typeof ensureJadwalSchemaAcademic_ === 'function') ? ensureJadwalSchemaAcademic_() : -1;
  const values = sh.getDataRange().getValues();
  const result = [];

  // TIDAK ADA fallback "tampilkan semua kalau kosong" — sengaja dihapus,
  // sama seperti getDashboardAllData()/getJadwalMengajar(): fallback itu
  // bikin jadwal periode LAMA nyangkut tampil lagi tiap kali periode aktif
  // memang benar-benar kosong (mis. baru saja di-reset).
  for (let i = 1; i < values.length; i++) {
    const emailRow    = String(values[i][0] || '').toLowerCase().trim();
    const semesterRow = String(values[i][1] || '').trim().toLowerCase();
    if (emailRow !== emailLogin) continue;
    if (semesterAktif && semesterRow !== semesterAktif) continue;
    if (tIdx > -1 && tahunAktif) {
      const rowTahun = String(values[i][tIdx] || '').trim();
      if (rowTahun && rowTahun !== tahunAktif) continue;
    }
    result.push({
      id:          i + 1,
      hari:        String(values[i][2] || '').trim().toUpperCase(),
      kelas:       values[i][3],
      mapel:       values[i][4],
      jam_mulai:   formatJam(values[i][5]),
      jam_selesai: formatJam(values[i][6])
    });
  }

  return (typeof _sortJadwalFlat_ === 'function') ? _sortJadwalFlat_(result) : result;
}
