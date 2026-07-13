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
const DEFAULT_APP_STORAGE_MODE  = 'central';

// =========================================================
// SPREADSHEET CONNECTION
// =========================================================
const SPREADSHEET_ID = '1fsdShDdm7ULvaiWE0QOTRfspi7U64HrL5ZPUX0l_As4';

function getCentralSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}

function getStorageMode_() {
  const mode = String(PropertiesService.getScriptProperties().getProperty(APP_STORAGE_MODE_PROPERTY) || DEFAULT_APP_STORAGE_MODE).toLowerCase().trim();
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
    'UPDATE_LOG': true
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
  if (!userSpreadsheetId) return getCentralSpreadsheet_();

  try {
    return SpreadsheetApp.openById(userSpreadsheetId);
  } catch (e) {
    return getCentralSpreadsheet_();
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
  } catch(innerErr) { /* must not throw */ }
}

/**
 * clearServerCache()
 * Bersihkan GAS Script Cache & User Cache setelah deploy baru.
 * Jalankan via: clasp run clearServerCache
 * atau buka Apps Script Editor → Run → clearServerCache
 *
 * CATATAN DEPLOYMENT:
 * - URL /dev  → selalu menampilkan kode terbaru (versi HEAD).
 * - URL produksi (Manage Deployments) → menampilkan versi snapshot.
 * - Agar kode baru tampil di URL produksi:
 *     Manage Deployments → Edit (✏️) → Version: "New version" → Deploy
 *   ATAU jalankan: clasp deploy --description "v$(Get-Date -f yyyy-MM-dd)"
 */
function clearServerCache() {
  CacheService.getScriptCache().removeAll();
  CacheService.getUserCache().removeAll();
  return { status: 'cleared', ts: new Date().toISOString() };
}

function generateLicenseToken() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
}