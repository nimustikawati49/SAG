/**
 * License.js — Sistem Lisensi Per Sekolah (school-wide)
 *
 * Lisensi disimpan di ScriptProperties (bukan per user).
 * 1 project GAS = 1 sekolah = 1 lisensi.
 * SuperAdmin yang mengelola (set / renew / deactivate / tier).
 *
 * TIER SYSTEM:
 *   LITE  — default gratis, maks 6 kelas, 192 siswa, tanpa fitur premium
 *   PRO   — unlock semua fitur 1 guru (modul ajar, RPP, export, katrol nilai)
 *   SCHOOL — PRO + multi-guru + rekap sekolah (dihandle via role kepsek)
 */

var LIC_ = {
  KEY         : 'SCHOOL_LICENSE_KEY',
  EXPIRES     : 'SCHOOL_LICENSE_EXPIRES',   // ISO date string
  STATUS      : 'SCHOOL_LICENSE_STATUS',    // 'active' | 'inactive'
  INSTALL_DATE: 'SCHOOL_INSTALL_DATE',      // legacy
  TRIAL       : 'SCHOOL_LICENSE_IS_TRIAL',
  TIER        : 'SCHOOL_TIER'              // 'LITE' | 'PRO' | 'SCHOOL'
};

/* Batas fitur tier LITE */
var LITE_LIMITS_ = {
  MAX_KELAS : 6,
  MAX_SISWA : 192
};

/* ─────────────────────────────────────────────────────────────
   INTERNAL HELPERS
───────────────────────────────────────────────────────────── */

function _readSchoolLicense_() {
  var props   = PropertiesService.getScriptProperties();
  var key     = props.getProperty(LIC_.KEY)     || '';
  var expires = props.getProperty(LIC_.EXPIRES) || '';
  var status  = props.getProperty(LIC_.STATUS)  || '';
  var trialFlag = String(props.getProperty(LIC_.TRIAL) || '').toLowerCase() === 'true';
  var isTrial = false;
  var isLifetime = false;
  var daysLeft = null;
  var isActive = false;

  // Default baru: jika belum pernah diset lisensi aplikasi, anggap lifetime aktif.
  if (!key && !expires && !status) {
    key = 'LIFETIME';
    status = 'active';
    isLifetime = true;
    isActive = true;
  } else if (expires) {
    var exp  = new Date(expires);
    var now  = new Date();
    daysLeft = Math.ceil((exp - now) / 86400000);
    isActive = (status === 'active') && (daysLeft > 0);
    isTrial  = trialFlag;
  } else {
    isLifetime = (status === 'active');
    isActive   = (status === 'active');
  }

  return {
    key     : key,
    expires : expires ? expires.split('T')[0] : '',
    status  : status,
    daysLeft: daysLeft,
    isActive: isActive,
    isTrial : isTrial,
    isLifetime: isLifetime
  };
}

/**
 * Dipakai oleh assertLicenseActive() di Auth.js.
 * Melempar error jika lisensi sekolah tidak aktif / expired.
 */
function assertSchoolLicenseActive_() {
  var lic = _readSchoolLicense_();
  if (!lic.isActive) {
    if (!lic.expires || lic.daysLeft <= 0) {
      throw new Error('LISENSI_EXPIRED');
    }
    throw new Error('INPUT_LICENSE');
  }
  return true;
}

/* ─────────────────────────────────────────────────────────────
   PUBLIC – dipanggil dari frontend via google.script.run
───────────────────────────────────────────────────────────── */

/** Ambil info lisensi sekolah (SuperAdmin only) */
function getSchoolLicense() {
  var auth = getAuth();
  if (auth.role !== 'superadmin' && auth.role !== 'admin') throw new Error('AKSES_DITOLAK');
  return _readSchoolLicense_();
}

/**
 * Set / aktivasi lisensi sekolah (SuperAdmin only).
 * @param {string} key         Kode lisensi unik
 * @param {string} expiredDate 'YYYY-MM-DD'
 */
function setSchoolLicense(key, expiredDate) {
  var auth = getAuth();
  if (auth.role !== 'superadmin' && auth.role !== 'admin') throw new Error('AKSES_DITOLAK');
  if (!key || !expiredDate) throw new Error('Key dan tanggal expired wajib diisi');

  key = String(key).trim().toUpperCase();

  var expDate = new Date(expiredDate);
  if (isNaN(expDate.getTime())) throw new Error('Format tanggal tidak valid (YYYY-MM-DD)');

  var props = PropertiesService.getScriptProperties();
  props.setProperty(LIC_.KEY,     key);
  props.setProperty(LIC_.EXPIRES, expDate.toISOString());
  props.setProperty(LIC_.STATUS,  'active');
  props.setProperty(LIC_.TRIAL,   key.indexOf('TRIAL') === 0 ? 'true' : 'false');

  logAudit('SET_SCHOOL_LICENSE', getLoginEmail(), 'Key: ' + key + ' | Expires: ' + expiredDate);

  return _readSchoolLicense_();
}

/**
 * Perpanjang lisensi sekolah (SuperAdmin only).
 * @param {number} years Jumlah tahun perpanjangan (default: 1)
 */
function renewSchoolLicense(years) {
  var auth = getAuth();
  if (auth.role !== 'superadmin' && auth.role !== 'admin') throw new Error('AKSES_DITOLAK');
  years = parseInt(years) || 1;
  if (years < 1 || years > 10) throw new Error('Tahun perpanjangan harus antara 1-10');

  var props   = PropertiesService.getScriptProperties();
  var current = props.getProperty(LIC_.EXPIRES);

  var base = (current && new Date(current) > new Date())
    ? new Date(current)
    : new Date();

  base.setFullYear(base.getFullYear() + years);

  props.setProperty(LIC_.EXPIRES, base.toISOString());
  props.setProperty(LIC_.STATUS,  'active');
  props.setProperty(LIC_.TRIAL,   'false');

  logAudit('RENEW_SCHOOL_LICENSE', getLoginEmail(), '+' + years + ' tahun → ' + base.toISOString().split('T')[0]);

  return _readSchoolLicense_();
}

/** Non-aktifkan lisensi sekolah (SuperAdmin only) */
function deactivateSchoolLicense() {
  var auth = getAuth();
  if (auth.role !== 'superadmin' && auth.role !== 'admin') throw new Error('AKSES_DITOLAK');

  var props = PropertiesService.getScriptProperties();
  props.setProperty(LIC_.STATUS, 'inactive');
  props.setProperty(LIC_.TRIAL,  'false');

  logAudit('DEACTIVATE_SCHOOL_LICENSE', getLoginEmail(), '');
  return { status: 'inactive' };
}

/**
 * Buat kode lisensi baru (SuperAdmin only).
 * Format: SEKOLAH-YYYYMMDD-XXXX
 */
function generateSchoolLicenseKey() {
  var auth = getAuth();
  if (auth.role !== 'superadmin' && auth.role !== 'admin') throw new Error('AKSES_DITOLAK');

  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var rand  = '';
  for (var i = 0; i < 8; i++) {
    rand += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  return 'SCH-' + today + '-' + rand;
}

/**
 * Dipanggil dari getLicenseBadge() di Auth.js agar UI bisa
 * menampilkan badge expiry yang benar.
 */
function getSchoolLicenseBadge_() {
  var lic = _readSchoolLicense_();
  if (!lic.isActive) {
    return { level: 'expired', days: lic.daysLeft };
  }
  if (lic.isLifetime) {
    return { level: 'ok', days: null, isLifetime: true };
  }
  if (lic.daysLeft <= 30) return { level: 'danger',  days: lic.daysLeft };
  if (lic.daysLeft <= 60) return { level: 'warning', days: lic.daysLeft };
  return { level: 'ok', days: lic.daysLeft };
}

/**
 * Ambil status lisensi sekolah — bisa dipanggil oleh SEMUA role (bukan SA only).
 * Dipakai oleh checkLicenseStatusBadge() dan checkLicenseExpiryWarning_() di frontend.
 */
function getSchoolLicenseInfo() {
  var auth = getAuth();
  if (auth.role !== 'superadmin' && auth.role !== 'admin') return null;
  var lic = _readSchoolLicense_();
  return {
    isActive : lic.isActive,
    daysLeft : lic.daysLeft,
    expires  : lic.expires,
    status   : lic.status,
    isTrial  : lic.isTrial,
    isLifetime: lic.isLifetime
  };
}

/* ─────────────────────────────────────────────────────────────
   TIER SYSTEM — LITE / PRO / SCHOOL
───────────────────────────────────────────────────────────── */

/**
 * Ambil tier aktif. Default LITE jika belum di-set.
 * SuperAdmin & kepsek selalu dapat tier tertinggi (SCHOOL).
 */
function getTier_() {
  var props = PropertiesService.getScriptProperties();
  var tier  = props.getProperty(LIC_.TIER) || 'LITE';
  // Normalise
  tier = tier.toUpperCase();
  if (tier !== 'PRO' && tier !== 'SCHOOL') tier = 'LITE';
  return tier;
}

/**
 * Cek apakah tier saat ini memenuhi minimum yang dibutuhkan.
 * Order: LITE < PRO < SCHOOL
 */
var TIER_ORDER_ = { LITE: 0, PRO: 1, SCHOOL: 2 };

function _tierMeetsMin_(current, minimum) {
  return (TIER_ORDER_[current] || 0) >= (TIER_ORDER_[minimum] || 0);
}

/**
 * Lempar error UPGRADE_REQUIRED jika tier tidak memenuhi minimum.
 * Selalu bypass untuk superadmin dan kepsek.
 */
function assertMinTier_(minTier) {
  var auth = getAuth();
  if (auth.role === 'superadmin' || auth.role === 'kepsek') return true;
  var tier = getTier_();
  if (!_tierMeetsMin_(tier, minTier)) {
    throw new Error('UPGRADE_REQUIRED:' + minTier);
  }
  return true;
}

/**
 * Cek batas kelas untuk tier LITE.
 * @param {number} currentCount — jumlah kelas yang sudah ada
 */
function assertLiteKelasLimit_(currentCount) {
  var auth = getAuth();
  if (auth.role === 'superadmin' || auth.role === 'kepsek') return;
  if (_tierMeetsMin_(getTier_(), 'PRO')) return;
  if (currentCount >= LITE_LIMITS_.MAX_KELAS) {
    throw new Error('LITE_LIMIT_KELAS:' + LITE_LIMITS_.MAX_KELAS);
  }
}

/**
 * Cek batas siswa untuk tier LITE.
 * @param {number} currentCount — jumlah siswa yang sudah terdaftar
 */
function assertLiteSiswaLimit_(currentCount) {
  var auth = getAuth();
  if (auth.role === 'superadmin' || auth.role === 'kepsek') return;
  if (_tierMeetsMin_(getTier_(), 'PRO')) return;
  if (currentCount >= LITE_LIMITS_.MAX_SISWA) {
    throw new Error('LITE_LIMIT_SISWA:' + LITE_LIMITS_.MAX_SISWA);
  }
}

/**
 * Set tier sekolah (SuperAdmin only).
 * @param {string} tier — 'LITE' | 'PRO' | 'SCHOOL'
 * @param {string} key  — kode lisensi (diperlukan untuk PRO/SCHOOL)
 */
function setSchoolTier(tier, key) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  tier = String(tier || '').toUpperCase();
  if (tier !== 'LITE' && tier !== 'PRO' && tier !== 'SCHOOL') {
    throw new Error('Tier tidak valid. Pilih: LITE, PRO, atau SCHOOL');
  }

  if (tier !== 'LITE') {
    // Validasi kode lisensi untuk upgrade
    key = String(key || '').trim().toUpperCase();
    if (!key) throw new Error('Kode lisensi wajib diisi untuk upgrade ke ' + tier);
    if (!_isValidTierKey_(key, tier)) {
      throw new Error('Kode lisensi tidak valid untuk tier ' + tier);
    }
    // Simpan juga sebagai school license key
    var props = PropertiesService.getScriptProperties();
    props.setProperty(LIC_.KEY, key);
    props.setProperty(LIC_.STATUS, 'active');
    var exp = new Date();
    exp.setFullYear(exp.getFullYear() + 99); // lifetime
    props.setProperty(LIC_.EXPIRES, exp.toISOString());
  }

  PropertiesService.getScriptProperties().setProperty(LIC_.TIER, tier);
  logAudit('SET_TIER', getLoginEmail(), 'Tier: ' + tier + (key ? ' | Key: ' + key : ''));
  return { tier: tier, success: true };
}

/**
 * Validasi kode lisensi tier.
 * Format kode: PRO-XXXX-XXXX atau SCH-XXXX-XXXX (8 karakter alfanumerik unik).
 * Validasi sederhana berbasis format + checksum — tidak butuh server eksternal.
 */
function _isValidTierKey_(key, tier) {
  var prefix = (tier === 'PRO') ? 'PRO-' : 'SCH-';
  if (!key.startsWith(prefix)) return false;
  var parts = key.split('-');
  if (parts.length !== 3) return false;
  var a = parts[1], b = parts[2];
  if (a.length !== 4 || b.length !== 4) return false;
  // Checksum sederhana: sum char codes % 36 harus sama antara a dan b
  var sumA = 0, sumB = 0;
  for (var i = 0; i < a.length; i++) sumA += a.charCodeAt(i);
  for (var i = 0; i < b.length; i++) sumB += b.charCodeAt(i);
  return (sumA % 36) === (sumB % 36);
}

/**
 * Generate kode lisensi PRO atau SCHOOL (SuperAdmin only).
 * @param {string} tierType — 'PRO' | 'SCHOOL'
 */
function generateTierLicenseKey(tierType) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  tierType = String(tierType || 'PRO').toUpperCase();
  if (tierType !== 'PRO' && tierType !== 'SCHOOL') throw new Error('Tier tidak valid');

  var prefix = (tierType === 'PRO') ? 'PRO' : 'SCH';
  var chars  = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

  // Generate part A (4 chars)
  var a = '';
  for (var i = 0; i < 4; i++) a += chars.charAt(Math.floor(Math.random() * chars.length));

  // Generate part B yang checksumnya cocok dengan A
  var sumA = 0;
  for (var i = 0; i < a.length; i++) sumA += a.charCodeAt(i);
  var targetMod = sumA % 36;

  var b = '';
  while (true) {
    b = '';
    for (var i = 0; i < 4; i++) b += chars.charAt(Math.floor(Math.random() * chars.length));
    var sumB = 0;
    for (var i = 0; i < b.length; i++) sumB += b.charCodeAt(i);
    if (sumB % 36 === targetMod) break;
  }

  return prefix + '-' + a + '-' + b;
}

/**
 * Ambil info tier untuk frontend (semua user bisa akses).
 * Dipakai oleh bootApp() untuk menyesuaikan UI.
 */
function getSchoolTierInfo() {
  var auth = getAuth();
  // SA dan kepsek tidak dibatasi
  if (auth.role === 'superadmin' || auth.role === 'kepsek') {
    return { tier: 'SCHOOL', limits: null };
  }
  var tier = getTier_();
  return {
    tier  : tier,
    limits: tier === 'LITE' ? LITE_LIMITS_ : null
  };
}

/**
 * Kirim email reminder ke SuperAdmin saat lisensi sekolah hampir habis.
 * Dipanggil oleh trigger harian (runDailyLicenseReminder_ di Trigger.js).
 * Hanya mengirim email pada H-60, H-30, H-7.
 */
function checkSchoolLicenseExpiryReminder() {
  var lic = _readSchoolLicense_();
  if (!lic.expires || lic.isTrial) return; // Jangan kirim saat masa trial

  var days = lic.daysLeft;
  if (days > 60 || days <= 0) return; // Di luar rentang reminder

  // Kirim hanya pada H-60, H-30, H-7 (toleransi ±1 hari)
  var REMINDER_DAYS = [60, 30, 7];
  var shouldSend = REMINDER_DAYS.some(function(d) { return Math.abs(days - d) <= 1; });
  if (!shouldSend) return;

  var recipients = _getLicenseReminderRecipients_();
  if (!recipients.length) return;

  var appUrl  = ScriptApp.getService().getUrl();
  var bgColor = days <= 7 ? '#dc2626' : '#d97706';
  var subject = '⚠️ Lisensi Sekolah Berakhir ' + days + ' Hari Lagi — Sistem Akademik Guru';

  var htmlBody =
    '<div style="font-family:sans-serif;max-width:520px;margin:auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:' + bgColor + ';padding:20px 24px">' +
    '<h2 style="margin:0;color:#fff;font-size:18px">🔑 Reminder Lisensi Sekolah</h2>' +
    '</div>' +
    '<div style="padding:24px">' +
    '<p>Halo SuperAdmin,</p>' +
    '<p>Lisensi sekolah akan <b>berakhir dalam ' + days + ' hari</b>.</p>' +
    '<p>Tanggal expired: <b>' + lic.expires + '</b></p>' +
    '<p>Segera perpanjang agar semua guru dapat terus menggunakan Sistem Akademik Guru.</p>' +
    '<a href="' + appUrl + '" style="display:inline-block;background:#6C63FF;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:700;margin:8px 0">⚙️ Buka SuperAdmin Panel</a>' +
    '</div></div>';

  recipients.forEach(function(email) {
    GmailApp.sendEmail(email, subject, '', { htmlBody: htmlBody });
  });
  logAudit('SCHOOL_LICENSE_REMINDER', 'SYSTEM', 'H-' + days + ' | expired: ' + lic.expires);
}

function _getLicenseReminderRecipients_() {
  var recipients = [];
  try {
    var sa = getSuperAdminEmail_();
    if (sa) recipients.push(String(sa).toLowerCase().trim());
  } catch (e) {}

  try {
    var rows = sheet('USERS').getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var email = String(rows[i][0] || '').toLowerCase().trim();
      var role = String(rows[i][1] || '').toLowerCase().trim();
      var status = String(rows[i][2] || '').toLowerCase().trim();
      if (email && status === 'active' && (role === 'admin' || role === 'superadmin')) {
        recipients.push(email);
      }
    }
  } catch (e) {}

  return recipients.filter(function(email, idx, arr) {
    return email && arr.indexOf(email) === idx;
  });
}
