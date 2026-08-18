// JURNAL GURU DIGITAL VERSION: v3.2 (STABLE)
// STATUS : LOCKED : YES
// DATE   : 2026-01-05
// DEVELOPER : DWI M.

/**************** WEB APP ENTRY POINT ****************/
function doGet(e) {
  const tmpl = HtmlService.createTemplateFromFile('index');
  return tmpl.evaluate()
    .setTitle('Sistem Akademik Guru')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/** Server-side include helper � dipanggil dari template index.html */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =========================================================
// SUPERADMIN EMAIL ? disimpan di ScriptProperties
// Set awal via: File > Project Settings > Script properties
//               Key: SUPERADMIN_EMAIL
// =========================================================
const _SUPERADMIN_FALLBACK = 'nimustikawati49@guru.smp.belajar.id';
let   _superAdminEmailCache = null;

function getSuperAdminEmail_() {
  if (_superAdminEmailCache) return _superAdminEmailCache;
  try {
    const p = PropertiesService.getScriptProperties().getProperty('SUPERADMIN_EMAIL');
    _superAdminEmailCache = (p || _SUPERADMIN_FALLBACK).toLowerCase().trim();
  } catch(e) {
    _superAdminEmailCache = _SUPERADMIN_FALLBACK.toLowerCase();
  }
  return _superAdminEmailCache;
}

// =========================================================
// KONSTANTA GLOBAL
// =========================================================
const FOLDER_DOKUMENTASI  = 'JURNAL_DOKUMENTASI';
const MAX_IMAGE_SIZE_KB   = 900;
const THUMB_SIZE_KB       = 100;
const APP_STORAGE_MODE_PROPERTY = 'APP_STORAGE_MODE';
const DEFAULT_APP_STORAGE_MODE  = 'per_guru';

// =========================================================
// SPREADSHEET CONNECTION
// =========================================================
const SPREADSHEET_ID = '1fsdShDdm7ULvaiWE0QOTRfspi7U64HrL5ZPUX0l_As4';

function getCentralSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}

function getStorageMode_() {
  const props = PropertiesService.getScriptProperties();
  let raw = props.getProperty(APP_STORAGE_MODE_PROPERTY);
  if (!raw) {
    // Migrasi satu kali: belum pernah di-set eksplisit lewat panel SuperAdmin,
    // jadi pakai default & simpan supaya panel Storage Mode menampilkan
    // pilihan yang benar-benar aktif (bukan diam-diam ganti tiap deploy).
    raw = DEFAULT_APP_STORAGE_MODE;
    try { props.setProperty(APP_STORAGE_MODE_PROPERTY, raw); } catch (e) {}
  }
  const mode = String(raw).toLowerCase().trim();
  return mode === 'per_guru' ? 'per_guru' : 'central';
}

function _isCentralOnlySheet_(name) {
  const centralSheets = {
    'USERS': true,
    'LICENSES': true,
    'AUDIT_LOG': true,
    '_LOG_ERROR_': true,
    'DEPLOYMENTS': true,
    'RESOURCE_MAP': true,
    'APP_RELEASES': true,
    'UPDATE_LOG': true,
    'HARI_LIBUR': true
  };
  return !!centralSheets[String(name || '').trim().toUpperCase()];
}

function _getCentralSheetByName_(name) {
  return getCentralSpreadsheet_().getSheetByName(name);
}

function _ensureOperationalSheetFromCentral_(targetSpreadsheet, name) {
  if (!targetSpreadsheet) return null;
  let localSheet = targetSpreadsheet.getSheetByName(name);
  if (localSheet) return localSheet;

  const centralSheet = _getCentralSheetByName_(name);
  if (!centralSheet) return null;

  localSheet = targetSpreadsheet.insertSheet(name);
  const lastCol = centralSheet.getLastColumn();
  if (lastCol > 0) {
    const header = centralSheet.getRange(1, 1, 1, lastCol).getValues();
    localSheet.getRange(1, 1, 1, lastCol).setValues(header);
    localSheet.setFrozenRows(1);
    try {
      const bg = centralSheet.getRange(1, 1, 1, lastCol).getBackgrounds();
      const fw = centralSheet.getRange(1, 1, 1, lastCol).getFontWeights();
      localSheet.getRange(1, 1, 1, lastCol).setBackgrounds(bg).setFontWeights(fw);
    } catch (e) {}
  }
  return localSheet;
}

function _getResourceMapEntryForUser_(email, resourceType) {
  try {
    const targetEmail = String(email || '').toLowerCase().trim();
    if (!targetEmail) return null;
    const sh = _getCentralSheetByName_('RESOURCE_MAP');
    if (!sh || sh.getLastRow() < 2) return null;

    const rows = sh.getDataRange().getValues();
    const header = rows[0].map(h => String(h || '').toLowerCase().trim());
    const idx = {};
    header.forEach((h, i) => idx[h] = i);
    const typeNeed = String(resourceType || '').toLowerCase().trim();

    for (let i = rows.length - 1; i >= 1; i--) {
      const rowEmail = String(rows[i][idx.email_guru] || '').toLowerCase().trim();
      const rowType = String(rows[i][idx.resource_type] || '').toLowerCase().trim();
      const rowStatus = String(rows[i][idx.status] || 'active').toLowerCase().trim();
      if (rowEmail !== targetEmail) continue;
      if (rowType !== typeNeed) continue;
      if (rowStatus !== 'active') continue;
      return {
        id: String(rows[i][idx.id] || ''),
        deployment_id: String(rows[i][idx.deployment_id] || ''),
        email_guru: rowEmail,
        resource_type: rowType,
        resource_id: String(rows[i][idx.resource_id] || ''),
        resource_name: String(rows[i][idx.resource_name] || ''),
        owner_email: String(rows[i][idx.owner_email] || ''),
        status: rowStatus,
        catatan: String(rows[i][idx.catatan] || '')
      };
    }
  } catch (e) {}
  return null;
}

function _getDeploymentEntryForUser_(email) {
  try {
    const targetEmail = String(email || '').toLowerCase().trim();
    if (!targetEmail) return null;
    const sh = _getCentralSheetByName_('DEPLOYMENTS');
    if (!sh || sh.getLastRow() < 2) return null;
    const rows = sh.getDataRange().getValues();
    const header = rows[0].map(h => String(h || '').toLowerCase().trim());
    const idx = {};
    header.forEach((h, i) => idx[h] = i);

    for (let i = rows.length - 1; i >= 1; i--) {
      const ownerEmail = String(rows[i][idx.owner_email] || '').toLowerCase().trim();
      const saEmail = String(rows[i][idx.email_sa] || '').toLowerCase().trim();
      const status = String(rows[i][idx.status] || 'aktif').toLowerCase().trim();
      if (status !== 'aktif' && status !== 'active') continue;
      if (ownerEmail !== targetEmail && saEmail !== targetEmail) continue;
      return {
        id: String(rows[i][idx.id] || ''),
        spreadsheet_id: String(rows[i][idx.spreadsheet_id] || ''),
        folder_data_id: String(rows[i][idx.folder_data_id] || ''),
        folder_export_id: String(rows[i][idx.folder_export_id] || ''),
        owner_email: ownerEmail,
        email_sa: saEmail
      };
    }
  } catch (e) {}
  return null;
}

function resolveSpreadsheetIdForUser_(email) {
  const targetEmail = String(email || '').toLowerCase().trim();
  if (!targetEmail) return '';
  const fromResource = _getResourceMapEntryForUser_(targetEmail, 'data_spreadsheet');
  if (fromResource && fromResource.resource_id) return fromResource.resource_id;
  const deployment = _getDeploymentEntryForUser_(targetEmail);
  if (deployment && deployment.spreadsheet_id) return deployment.spreadsheet_id;
  return '';
}

function resolveFolderIdForUser_(email, resourceType) {
  const targetEmail = String(email || '').toLowerCase().trim();
  const targetType = String(resourceType || '').toLowerCase().trim();
  if (!targetEmail || !targetType) return '';

  const fromResource = _getResourceMapEntryForUser_(targetEmail, targetType);
  if (fromResource && fromResource.resource_id) return fromResource.resource_id;

  const deployment = _getDeploymentEntryForUser_(targetEmail);
  if (!deployment) return '';

  if ((targetType === 'arsip_folder' || targetType === 'dokumentasi_folder' || targetType === 'wali_dokumentasi_folder' || targetType === 'logo_folder') && deployment.folder_data_id) {
    return deployment.folder_data_id;
  }
  if ((targetType === 'export_folder' || targetType === 'backup_folder') && deployment.folder_export_id) {
    return deployment.folder_export_id;
  }
  return '';
}

function _getOrCreateFolderByName_(name, parentFolder) {
  const targetName = String(name || '').trim();
  if (!targetName) return parentFolder || null;
  let iter = parentFolder ? parentFolder.getFoldersByName(targetName) : DriveApp.getFoldersByName(targetName);
  if (iter.hasNext()) return iter.next();
  return parentFolder ? parentFolder.createFolder(targetName) : DriveApp.createFolder(targetName);
}

function getUserResourceFolder_(email, resourceType, fallbackName) {
  const mode = getStorageMode_();
  const targetEmail = String(email || '').toLowerCase().trim();
  const mappedId = mode === 'per_guru' ? resolveFolderIdForUser_(targetEmail, resourceType) : '';
  if (mappedId) {
    try { return DriveApp.getFolderById(mappedId); } catch (e) {}
  }
  return _getOrCreateFolderByName_(fallbackName || 'APP_DATA');
}

function getUserNestedFolder_(email, resourceType, fallbackRootName, segments) {
  let folder = getUserResourceFolder_(email, resourceType, fallbackRootName);
  const parts = Array.isArray(segments) ? segments : [];
  parts.forEach(function(segment) {
    const part = String(segment || '').trim();
    if (!part) return;
    folder = _getOrCreateFolderByName_(part, folder);
  });
  return folder;
}

function getSpreadsheet_(options) {
  const opts = options || {};
  if (opts.forceCentral) return getCentralSpreadsheet_();
  if (opts.sheetName && _isCentralOnlySheet_(opts.sheetName)) return getCentralSpreadsheet_();
  if (getStorageMode_() !== 'per_guru') return getCentralSpreadsheet_();

  let targetEmail = String(opts.email || '').toLowerCase().trim();
  if (!targetEmail) {
    try {
      targetEmail = Session.getEffectiveUser().getEmail().toLowerCase().trim();
    } catch (e) {
      targetEmail = '';
    }
  }
  if (!targetEmail) return getCentralSpreadsheet_();

  const userSpreadsheetId = resolveSpreadsheetIdForUser_(targetEmail);
  if (!userSpreadsheetId) {
    try {
      if (typeof _autoProvisionUserSpreadsheet_ === 'function') {
        _autoProvisionUserSpreadsheet_(targetEmail);
      }
    } catch (e) {}
  }

  const refreshedSpreadsheetId = resolveSpreadsheetIdForUser_(targetEmail);
  if (!refreshedSpreadsheetId) return getCentralSpreadsheet_();

  try {
    return SpreadsheetApp.openById(refreshedSpreadsheetId);
  } catch (e) {
    return getCentralSpreadsheet_();
  }
}

// =========================================================
// AUTO-PROVISIONING SPREADSHEET PER GURU
// =========================================================

// Sheet yang dibuat di spreadsheet pribadi guru (bukan sheet central).
var _GURU_OPERATIONAL_SHEETS_ = [
  'SETTING', 'JURNAL', 'ABSENSI', 'SISWA', 'NILAI_SISWA',
  'SETTING_NILAI', 'MasterSiswa', 'MasterTahunPelajaran', 'RiwayatKelas'
];

/**
 * _autoProvisionUserSpreadsheet_(email)
 * Buat spreadsheet pribadi untuk guru jika belum ada di RESOURCE_MAP.
 * Hanya berjalan saat storage mode = per_guru. Idempoten — aman dipanggil
 * berkali-kali. Spreadsheet dibuat di Drive script owner, bukan Drive guru.
 *
 * CATATAN: sebelumnya ada DUA definisi fungsi ini (satu di sini, satu di
 * SuperAdmin.js) — di Apps Script semua file berbagi satu namespace
 * global, jadi cuma salah satu yang benar-benar aktif (yang lain
 * ke-shadow diam-diam) dan masing-masing punya guard yang beda. Sudah
 * digabung jadi satu di sini dengan SEMUA guard:
 *  - SuperAdmin selalu ke central (bukan spreadsheet baru).
 *  - Lisensi sekolah harus aktif.
 *  - Guru yang SUDAH punya data central dilewati, bukan diarahkan ke
 *    spreadsheet baru yang kosong (root cause insiden dashboard kosong).
 *  - Klaim ringan per-email lewat cache, bukan LockService global —
 *    supaya proses provisioning satu guru (bisa >10 detik) tidak
 *    memblokir guru lain yang kebetulan butuh lock yang sama.
 */
function _autoProvisionUserSpreadsheet_(email) {
  if (getStorageMode_() !== 'per_guru') return;

  var targetEmail = String(email || '').toLowerCase().trim();
  if (!targetEmail) return;

  var cache = CacheService.getScriptCache();
  var cacheKey = 'PROV_' + targetEmail.replace(/[^a-z0-9]/g, '_');
  if (cache.get(cacheKey)) return;

  // Cek dulu tanpa klaim supaya path yang sudah ter-provisioning cepat.
  if (resolveSpreadsheetIdForUser_(targetEmail)) {
    cache.put(cacheKey, '1', 21600);
    return;
  }

  // SuperAdmin (pemilik script) selalu dipetakan ke spreadsheet CENTRAL —
  // jangan pernah dibuatkan spreadsheet baru, supaya saat login untuk
  // kelola/testing, data guru-guru lain di central tetap terlihat.
  if (targetEmail === getSuperAdminEmail_()) {
    try {
      if (typeof _upsertResourceMapEntry_ === 'function') {
        _upsertResourceMapEntry_({
          email_guru: targetEmail,
          resource_type: 'data_spreadsheet',
          resource_id: getCentralSpreadsheet_().getId(),
          resource_name: 'Central (SuperAdmin)',
          owner_email: targetEmail,
          status: 'active',
          catatan: 'SuperAdmin selalu memakai central'
        });
      }
      cache.put(cacheKey, '1', 21600);
    } catch (e) {}
    return;
  }

  try {
    if (!_readSchoolLicense_().isActive) return;
  } catch (e) { return; }

  // Jangan provisioning kalau email ini SUDAH punya riwayat data di
  // spreadsheet central (SETTING) — itu tandanya akun lama, bukan guru
  // baru. Memindahkannya diam-diam ke spreadsheet pribadi yang kosong
  // bikin seolah semua datanya hilang (padahal masih ada di central,
  // cuma tidak lagi dibaca dari sana).
  try {
    if (typeof _hasExistingCentralTeacherData_ === 'function' && _hasExistingCentralTeacherData_(targetEmail)) {
      cache.put(cacheKey, '1', 21600);
      try { logError_('AUTO_PROVISION_SKIPPED_EXISTING_DATA', new Error(targetEmail + ' sudah punya data central, provisioning dilewati')); } catch (e2) {}
      return;
    }
  } catch (e) { return; }

  var claimKey = 'PROV_CLAIM_' + targetEmail.replace(/[^a-z0-9]/gi, '_');
  try {
    if (cache.get(claimKey)) return; // proses lain sedang mengerjakan, jangan ikut nunggu
    cache.put(claimKey, '1', 60);
  } catch (e) {}

  try {
    // Cek ulang setelah klaim, siapa tahu sudah diselesaikan proses lain barusan.
    if (resolveSpreadsheetIdForUser_(targetEmail)) {
      cache.put(cacheKey, '1', 21600);
      return;
    }

    var guruSlug = targetEmail.split('@')[0].replace(/[^a-z0-9]/gi, '_');
    var ss = SpreadsheetApp.create('Data_Guru_' + guruSlug);
    var ssId = ss.getId();
    try {
      var defaultSh = ss.getSheetByName('Sheet1') || ss.getSheetByName('Lembar1');
      if (defaultSh && ss.getSheets().length > 1) ss.deleteSheet(defaultSh);
    } catch (e) {}
    _GURU_OPERATIONAL_SHEETS_.forEach(function(name) {
      try { _ensureOperationalSheetFromCentral_(ss, name); } catch (e) {}
    });

    if (typeof _upsertResourceMapEntry_ === 'function') {
      _upsertResourceMapEntry_({
        email_guru: targetEmail,
        resource_type: 'data_spreadsheet',
        resource_id: ssId,
        resource_name: 'Data_Guru_' + guruSlug,
        owner_email: getSuperAdminEmail_(),
        status: 'active',
        catatan: 'Auto-provisioned ' + new Date().toISOString()
      });
    }

    try { logAudit('AUTO_PROVISION_SPREADSHEET', targetEmail, ssId); } catch (e) {}
    cache.put(cacheKey, '1', 21600);

  } catch (provErr) {
    try { logError_('AUTO_PROVISION_SPREADSHEET', provErr); } catch (e) {}
  } finally {
    try { cache.remove(claimKey); } catch (e) {}
  }
}

// =========================================================
// HELPER UMUM
// =========================================================
function daysBetween(d1, d2) {
  if (!d1 || !d2) return 0;
  return Math.floor((d2.getTime() - d1.getTime()) / (1000 * 60 * 60 * 24));
}

function sheet(name, options) {
  const opts = options || {};
  opts.sheetName = name;
  const ss = getSpreadsheet_(opts);
  let sh = ss.getSheetByName(name);
  if (!sh && ss.getId() !== getCentralSpreadsheet_().getId() && !_isCentralOnlySheet_(name)) {
    sh = _ensureOperationalSheetFromCentral_(ss, name);
  }
  if (!sh && ss.getId() !== getCentralSpreadsheet_().getId()) {
    sh = getCentralSpreadsheet_().getSheetByName(name);
  }
  return sh;
}

function authEmail() {
  return Session.getEffectiveUser().getEmail().toLowerCase();
}

function getSheetCached(sheetName, ttl = 60) {
  const cache  = CacheService.getScriptCache();
  // Include email in cache key to prevent cross-user data leakage
  const email  = (function(){ try { return Session.getEffectiveUser().getEmail().toLowerCase(); } catch(e){ return 'anon'; } })();
  const key    = 'SHEET_' + sheetName + '_' + email;
  const cached = cache.get(key);
  if (cached) {
    return JSON.parse(cached);
  }
  const targetSheet = sheet(sheetName);
  const data = targetSheet ? targetSheet.getDataRange().getValues() : [];
  try { cache.put(key, JSON.stringify(data), ttl); } catch(e) { /* ignore quota errors */ }
  return data;
}

function invalidateCache_(sheetName) {
  const cache = CacheService.getScriptCache();
  const email = (function(){ try { return Session.getEffectiveUser().getEmail().toLowerCase(); } catch(e){ return 'anon'; } })();
  cache.remove('SHEET_' + sheetName + '_' + email);
  // Also remove legacy key (without email) for backward compatibility
  cache.remove('SHEET_' + sheetName);
  // Also invalidate dashboard bundle cache when sheet data changes
  try { CacheService.getUserCache().remove('DASH_ALL'); } catch(e) {}
}

/**
 * getSheetValuesCached_(sheetName, sh, ttl)
 * Sama seperti getSheetCached(), tapi menerima objek Sheet yang sudah
 * di-resolve oleh caller (mis. lewat ensureXSheet_() yang juga menangani
 * pembuatan sheet baru + migrasi header) — dipakai saat caller butuh
 * caching di lapisan baca data tapi resolusi sheet-nya sendiri lebih
 * rumit dari sekadar sheet(name). Pakai cache key persis sama dengan
 * getSheetCached()/invalidateCache_() supaya saling kompatibel —
 * invalidateCache_(sheetName) akan menghapus cache ini juga.
 */
function getSheetValuesCached_(sheetName, sh, ttl = 120) {
  const cache  = CacheService.getScriptCache();
  const email  = (function(){ try { return Session.getEffectiveUser().getEmail().toLowerCase(); } catch(e){ return 'anon'; } })();
  const key    = 'SHEET_' + sheetName + '_' + email;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) { /* cache korup, baca ulang */ }
  }
  const data = sh ? sh.getDataRange().getValues() : [];
  try { cache.put(key, JSON.stringify(data), ttl); } catch(e) { /* > 100KB, lewati cache */ }
  return data;
}

/**
 * invalidateDashboardCache_()
 * Hapus cache dashboard user saat ada perubahan data
 * (jadwal, setting, jurnal). Dipanggil dari Save* functions.
 */
function invalidateDashboardCache_() {
  try { CacheService.getUserCache().remove('DASH_ALL'); } catch(e) {}
}

/**
 * logError_(context, err)
 * Catat error ke sheet _LOG_ERROR_ untuk monitoring.
 * Tidak pernah throw agar tidak mengganggu caller.
 */
/**
 * trimLogSheetIfNeeded_(sh, maxRows)
 * Kalau jumlah baris data (tidak termasuk header) di sheet log/audit
 * melebihi maxRows, hapus baris PALING LAMA (persis di bawah header)
 * sampai sisa tepat maxRows baris terbaru. Dipakai untuk sheet yang terus
 * bertambah tanpa batas (_LOG_ERROR_, AUDIT_LOG, UPDATE_LOG) supaya tidak
 * membengkak dan memperlambat baca/tulis sheet lain di spreadsheet yang
 * sama. Gagal diam-diam (jangan sampai proses utama ikut gagal cuma
 * karena trim gagal).
 */
function trimLogSheetIfNeeded_(sh, maxRows) {
  try {
    if (!sh) return;
    var dataRows = sh.getLastRow() - 1; // exclude header
    if (dataRows <= maxRows) return;
    var excess = dataRows - maxRows;
    sh.deleteRows(2, excess);
  } catch (e) { /* jangan sampai trim gagal mengganggu proses utama */ }
}

/**
 * logTiming_(label, startMs)
 * Diagnostik performa sementara — catat berapa lama sebuah tahap makan
 * waktu, ditulis ke sheet _LOG_ERROR_ yang sama (context diberi prefix
 * ⏱️ TIMING supaya gampang dibedakan dari error asli saat ditinjau).
 */
function logTiming_(label, startMs) {
  try {
    var elapsed = Date.now() - startMs;
    logError_('⏱️ TIMING: ' + label, new Error(elapsed + 'ms'));
  } catch (e) {}
}

function logError_(context, err) {
  try {
    var ss = getCentralSpreadsheet_();
    var sh = ss.getSheetByName('_LOG_ERROR_');
    if (!sh) {
      sh = ss.insertSheet('_LOG_ERROR_');
      sh.appendRow(['timestamp','email','context','error','stack']);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#fee2e2');
    }
    var email = 'system';
    try { email = Session.getEffectiveUser().getEmail(); } catch(e2) {}
    var msg   = String(err && err.message ? err.message : err);
    var stack = String(err && err.stack  ? err.stack  : '');
    sh.appendRow([new Date(), email, String(context), msg, stack.substring(0, 500)]);
    trimLogSheetIfNeeded_(sh, 300);
  } catch(innerErr) { /* must not throw */ }
}

/**
 * clearServerCache()
 * Dulu dimaksudkan untuk "bersihkan semua cache" setelah deploy baru,
 * tapi CacheService GAS TIDAK punya API untuk menghapus SEMUA key
 * sekaligus tanpa tahu key-nya lebih dulu — Cache.removeAll() WAJIB diisi
 * array key, bukan dipanggil tanpa argumen (itu sebab error "The
 * parameters () don't match the method signature" yang selalu muncul
 * tiap fungsi ini dijalankan — sudah dari awal begini, bukan regresi
 * baru). Semua cache di aplikasi ini sudah pakai TTL pendek (60-180
 * detik untuk data sheet, lihat getSheetCached/invalidateCache_ di file
 * ini) jadi otomatis kadaluwarsa sendiri — tidak butuh wipe manual.
 * Fungsi ini disederhanakan supaya aman dijalankan (tidak lagi error),
 * tapi memang tidak melakukan apa-apa selain melapor status.
 *
 * CATATAN DEPLOYMENT (tetap berlaku):
 * - URL /dev  → selalu menampilkan kode terbaru (versi HEAD).
 * - URL produksi (Manage Deployments) → menampilkan versi snapshot.
 * - Agar kode baru tampil di URL produksi:
 *     Manage Deployments → Edit (✏️) → Version: "New version" → Deploy
 *   ATAU jalankan: clasp deploy --description "v$(Get-Date -f yyyy-MM-dd)"
 */
function clearServerCache() {
  return {
    status: 'noop',
    note: 'CacheService tidak mendukung hapus semua key sekaligus — cache di aplikasi ini pakai TTL pendek dan kadaluwarsa otomatis, tidak perlu wipe manual.',
    ts: new Date().toISOString()
  };
}

function generateLicenseToken() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
}