/**
 * License.js — Sistem Lisensi Per Deployment (school-wide)
 *
 * Lisensi disimpan di ScriptProperties (bukan per user) — 1 project/
 * deployment GAS = 1 lisensi, dikelola SuperAdmin (set / renew /
 * deactivate). CATATAN: satu deployment TIDAK LAGI berarti satu
 * sekolah — sejak fitur multi-sekolah ditambahkan, satu deployment
 * bisa melayani beberapa sekolah sekaligus, dikelompokkan lewat
 * kolom kode_sekolah di USERS (lihat updateUserSekolah() di Users.js
 * dan _kepsekActiveGuruEmails_() di Kepsek.js). Lisensi ini tetap satu
 * untuk seluruh deployment, bukan per kode_sekolah.
 *
 * Akses guru sendiri diatur lewat status akun per-akun (lihat Users.js):
 * guru baru otomatis dapat masa uji coba 30 hari, lalu SuperAdmin
 * mengaktifkan penuh lewat panel "Aktivasi Akun Guru". Tidak ada lagi
 * pembagian fitur per tier (LITE/PRO/SCHOOL) — begitu aktif, semua
 * fitur terbuka.
 */

var LIC_ = {
  KEY         : 'SCHOOL_LICENSE_KEY',
  EXPIRES     : 'SCHOOL_LICENSE_EXPIRES',   // ISO date string
  STATUS      : 'SCHOOL_LICENSE_STATUS',    // 'active' | 'inactive'
  INSTALL_DATE: 'SCHOOL_INSTALL_DATE',      // legacy
  TRIAL       : 'SCHOOL_LICENSE_IS_TRIAL'
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
