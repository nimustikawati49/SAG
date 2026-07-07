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

// =========================================================
// SPREADSHEET CONNECTION
// =========================================================
const SPREADSHEET_ID = '1fsdShDdm7ULvaiWE0QOTRfspi7U64HrL5ZPUX0l_As4';

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}

// =========================================================
// HELPER UMUM
// =========================================================
function daysBetween(d1, d2) {
  if (!d1 || !d2) return 0;
  return Math.floor((d2.getTime() - d1.getTime()) / (1000 * 60 * 60 * 24));
}

function sheet(name) {
  return getSpreadsheet_().getSheetByName(name);
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
  const data = sheet(sheetName).getDataRange().getValues();
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
    var ss = getSpreadsheet_();
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