/**
 * License.js — Sistem Lisensi Per Deployment (school-wide)
 *
 * Lisensi disimpan di ScriptProperties (bukan per user) — 1 project/
 * deployment GAS = 1 lisensi untuk SELURUH deployment. CATATAN PENTING:
 * properti ini (SCHOOL_LICENSE_KEY/EXPIRES/STATUS/IS_TRIAL) HANYA PERNAH
 * DIBACA (_readSchoolLicense_ di bawah) — tidak ada satu pun jalur di
 * aplikasi ini (UI maupun server) yang pernah MENULISNYA. Jadi kalau
 * belum pernah diisi manual lewat Apps Script Editor > Project Settings,
 * lisensi SELALU default lifetime-aktif (lihat cabang "belum pernah
 * diset" di _readSchoolLicense_). Reminder H-60/30/7 berbasis
 * expired date yang dulu ada di sini sudah dihapus karena tidak akan
 * pernah benar-benar terpakai selama tidak ada jalur untuk mengisi
 * SCHOOL_LICENSE_EXPIRES.
 *
 * CATATAN: satu deployment TIDAK LAGI berarti satu sekolah — sejak
 * fitur multi-sekolah ditambahkan, satu deployment bisa melayani
 * beberapa sekolah sekaligus, dikelompokkan lewat kolom kode_sekolah di
 * USERS (lihat updateUserSekolah() di Users.js dan
 * _kepsekActiveGuruEmails_() di Kepsek.js). Lisensi ini tetap satu untuk
 * seluruh deployment, bukan per kode_sekolah.
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

// NOTE: checkSchoolLicenseExpiryReminder()/_getLicenseReminderRecipients_()
// dulu ada di sini (dipanggil dari runDailyLicenseReminder_ di Trigger.js,
// dihapus) — keduanya dihapus karena TIDAK ADA satu pun jalur di seluruh
// kode yang pernah menulis SCHOOL_LICENSE_EXPIRES ke ScriptProperties,
// jadi fungsi ini tidak akan pernah benar-benar mengirim apa pun (selalu
// keluar di baris pertama: `if (!lic.expires...) return;`). Lisensi
// sekolah di sistem ini SELALU default ke lifetime (lihat
// _readSchoolLicense_ di bawah) sejak model aktivasi lifetime per-akun
// guru diperkenalkan — tidak ada lagi jalur untuk memasang lisensi
// bertanggal-expired lewat aplikasi ini.
