/**
 * Users.js — Manajemen Pengguna & Lisensi
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

/* ─────────────────────────────────────────────────────────────
   TRIAL 30 HARI — akun guru baru otomatis bisa mencoba aplikasi
   tanpa perlu didaftarkan manual dulu oleh SuperAdmin. Statusnya
   'trial' di USERS sampai SuperAdmin klik "Aktifkan Penuh", yang
   mengubahnya jadi 'active' secara permanen tanpa pernah menyentuh
   sheet operasional guru (SETTING/JURNAL/SISWA/dll).
───────────────────────────────────────────────────────────── */
var TRIAL_DAYS_ = 30;

/**
 * _ensureUsersKodeSekolahColumn_()
 * Sheet USERS (central) dibuat manual di awal (bukan lewat kode), jadi
 * migrasi kolom baru harus idempoten & defensif seperti pola ensureXSheet_
 * lain di codebase ini — cek header dulu, baru tambahkan di ujung kalau
 * belum ada. Kolom ini dipakai untuk mengelompokkan guru per sekolah
 * (multi-sekolah dalam satu deployment yang sama) — SuperAdmin yang
 * mengisi lewat updateUserSekolah(), bukan diisi guru sendiri.
 */
function _ensureUsersKodeSekolahColumn_() {
  var sh = sheet('USERS');
  if (!sh) return null;
  var lastCol = sh.getLastColumn();
  var header = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||'').toLowerCase().trim(); })
    : [];
  if (header.indexOf('kode_sekolah') === -1) {
    sh.getRange(1, lastCol + 1).setValue('kode_sekolah');
  }
  return sh;
}

function _computeTrialInfo_(dibuatDate) {
  var created = dibuatDate ? new Date(dibuatDate) : new Date();
  var elapsedDays = Math.floor((Date.now() - created.getTime()) / 86400000);
  return {
    daysLeft: Math.max(0, TRIAL_DAYS_ - elapsedDays),
    expired : elapsedDays >= TRIAL_DAYS_
  };
}

/**
 * _autoRegisterTrialUser_(email)
 * Dipanggil dari getAuth() saat email belum ada di USERS sama sekali.
 * Daftarkan sebagai guru role 'admin' status 'trial' supaya bisa
 * langsung memakai aplikasi. Dikunci (LockService) supaya request
 * paralel dari akun yang sama tidak membuat baris dobel.
 */
function _autoRegisterTrialUser_(email) {
  var sh = _ensureUsersKodeSekolahColumn_();
  if (!sh) return null;

  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(5000); } catch (e) { gotLock = false; }
  if (!gotLock) return null;

  try {
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase().trim() === email) {
        return { dibuat: data[i][3] }; // sudah didaftarkan proses lain, jangan dobel
      }
    }
    var now = new Date();
    sh.appendRow([email, 'admin', 'trial', now]);
    try { logAudit('AUTO_REGISTER_TRIAL', email, 'Trial 30 hari dimulai'); } catch (e2) {}
    return { dibuat: now };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * getGuruActivationList()
 * SuperAdmin only — daftar SEMUA akun guru (role admin/kepsek, bukan
 * superadmin) beserta status aktivasinya (trial/active/inactive).
 * Menggantikan kartu "Lisensi Aplikasi" lama sebagai satu-satunya
 * tempat SuperAdmin mengaktifkan/menonaktifkan akun guru. Urutan:
 * trial (paling mendesak dulu) > active > inactive.
 */
function getGuruActivationList() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh = _ensureUsersKodeSekolahColumn_();
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var role = String(data[i][1] || 'admin').toLowerCase();
    var email = String(data[i][0] || '').toLowerCase().trim();
    if (!email) continue;
    var isSelfSA = (role === 'superadmin');
    // SuperAdmin SENGAJA ikut ditampilkan (dulu di-skip) — kalau SuperAdmin
    // juga mengajar, dia perlu lihat status koneksi Drive pribadinya sendiri
    // di tabel yang sama seperti guru lain (lihat DriveConnect.js, role
    // superadmin sudah diizinkan connectOwnSpreadsheet()). Statusnya
    // dipaksa 'active' (lifetime) di sini APAPUN isi kolom status di sheet
    // — SuperAdmin selalu bypass gerbang lisensi/trial (lihat
    // assertLicenseActive di Auth.js), jadi tampilannya harus konsisten
    // dengan itu, bukan ikut nilai mentah yang mungkin belum tentu 'active'.
    var status = isSelfSA ? 'active' : String(data[i][2] || 'inactive').toLowerCase().trim();

    var entry = {
      email: email,
      role: role,
      status: status,
      dibuat: data[i][3] ? Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '-',
      kode_sekolah: String(data[i][4] || '').trim(),
      days_left: null,
      expired: false,
      drive_self_owned: null,
      isSelfSA: isSelfSA
    };
    if (status === 'trial') {
      var info = _computeTrialInfo_(data[i][3]);
      entry.days_left = info.daysLeft;
      entry.expired = info.expired;
    }
    // Kepemilikan Drive database — relevan untuk guru (role admin) DAN
    // SuperAdmin sendiri (kalau juga mengajar) di mode per_guru, lihat
    // DriveConnect.js. null = tidak relevan (kepsek, mode central, dll).
    if ((role === 'admin' || isSelfSA) && typeof getStorageMode_ === 'function' && getStorageMode_() === 'per_guru') {
      try { entry.drive_self_owned = typeof _isDriveSelfOwned_ === 'function' ? _isDriveSelfOwned_(email) : null; } catch (e) { entry.drive_self_owned = null; }
    }
    result.push(entry);
  }

  var statusOrder = { trial: 0, active: 1, inactive: 2 };
  result.sort(function(a, b) {
    var oa = statusOrder[a.status] !== undefined ? statusOrder[a.status] : 3;
    var ob = statusOrder[b.status] !== undefined ? statusOrder[b.status] : 3;
    if (oa !== ob) return oa - ob;
    if (a.status === 'trial') {
      if (a.expired === b.expired) return a.days_left - b.days_left;
      return a.expired ? -1 : 1;
    }
    return a.email.localeCompare(b.email);
  });
  return result;
}

/**
 * _sendAccountEmail_(email, subject, headerColor, headerEmoji, headerTitle, bodyHtml)
 * Helper kirim email "bukti autentik" perubahan akun — SATU-SATUNYA
 * tempat email masih dipakai untuk notifikasi akun guru (lihat komentar
 * runDailyReminderCheck_ di Trigger.js: reminder harian & backup sudah
 * pindah ke notifikasi in-app, BUKAN email — email cuma dipertahankan
 * untuk 2 momen yang butuh bukti tertulis: aktivasi lifetime & perubahan
 * role, keduanya HANYA dipicu aksi eksplisit SuperAdmin, bukan trigger
 * terjadwal, jadi tidak akan pernah spam berulang). Gagal kirim tidak
 * boleh menggagalkan aksi utamanya (fail-soft, dibungkus try/catch di
 * pemanggil).
 */
function _sendAccountEmail_(email, subject, headerColor, headerEmoji, headerTitle, bodyHtml) {
  var appUrl = ScriptApp.getService().getUrl();
  var htmlBody =
    '<div style="font-family:sans-serif;max-width:520px;margin:auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:' + headerColor + ';padding:20px 24px">' +
    '<h2 style="margin:0;color:#fff;font-size:18px">' + headerEmoji + ' ' + headerTitle + '</h2>' +
    '</div>' +
    '<div style="padding:24px">' + bodyHtml +
    '<a href="' + appUrl + '" style="display:inline-block;background:#6C63FF;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:700;margin:8px 0">🚀 Buka Aplikasi</a>' +
    '</div></div>';
  GmailApp.sendEmail(email, subject, '', { htmlBody: htmlBody });
}

/**
 * _sendLifetimeActivationEmail_(email)
 * Dipanggil dari approveTrialAccount() DAN updateUser() (saat status
 * diset 'active') — dua jalur berbeda yang sama-sama berarti "SuperAdmin
 * mengaktifkan akun ini permanen" dari sudut pandang guru.
 */
function _sendLifetimeActivationEmail_(email) {
  try {
    _sendAccountEmail_(
      email,
      '✅ Akun Diaktifkan Permanen — Sistem Akademik Guru',
      '#16a34a',
      '✅',
      'Akun Anda Aktif Permanen',
      '<p>Halo,</p>' +
      '<p>Akun Anda (<b>' + email + '</b>) telah diaktifkan <b>permanen (lifetime)</b> oleh SuperAdmin — tidak lagi terikat batas masa uji coba.</p>' +
      '<p style="color:#6b7280;font-size:12px">Email ini adalah bukti autentik aktivasi akun Anda.</p>'
    );
  } catch (e) {
    try { logError_('SEND_LIFETIME_ACTIVATION_EMAIL', e); } catch (e2) {}
  }
}

/**
 * _sendRoleChangeEmail_(email, newRole)
 * Dipanggil dari updateUser() saat payload.role diubah.
 */
function _sendRoleChangeEmail_(email, newRole) {
  try {
    var roleLabel = { admin: 'Guru', kepsek: 'Kepala Sekolah', superadmin: 'SuperAdmin' }[newRole] || newRole;
    _sendAccountEmail_(
      email,
      '👤 Role Akun Diubah — Sistem Akademik Guru',
      '#6C63FF',
      '👤',
      'Role Akun Anda Diubah',
      '<p>Halo,</p>' +
      '<p>Role akun Anda (<b>' + email + '</b>) telah diubah oleh SuperAdmin menjadi <b>' + roleLabel + '</b>.</p>' +
      '<p style="color:#6b7280;font-size:12px">Email ini adalah bukti autentik perubahan role akun Anda.</p>'
    );
  } catch (e) {
    try { logError_('SEND_ROLE_CHANGE_EMAIL', e); } catch (e2) {}
  }
}

/**
 * approveTrialAccount(email)
 * SuperAdmin only — ubah status akun jadi 'active' permanen.
 * HANYA mengubah baris USERS (dan memastikan ada baris LICENSES aktif
 * untuk kompatibilitas fitur lama) — tidak pernah menyentuh sheet
 * operasional guru (SETTING/JURNAL/SISWA/NILAI/dll), jadi data yang
 * sudah dibuat guru selama masa trial tetap utuh.
 */
function approveTrialAccount(email) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  email = String(email || '').toLowerCase().trim();
  if (!email) throw new Error('Email tidak valid');

  var sh = sheet('USERS');
  var data = sh.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').toLowerCase().trim() !== email) continue;
    sh.getRange(i + 1, 3).setValue('active');
    found = true;
    break;
  }
  if (!found) throw new Error('Akun tidak ditemukan');

  try {
    if (typeof _autoProvisionUserSpreadsheet_ === 'function') {
      _autoProvisionUserSpreadsheet_(email);
    }
  } catch (e) {}

  var shLic = sheet('LICENSES');
  var lics = shLic.getDataRange().getValues();
  var licFound = false;
  for (var j = 1; j < lics.length; j++) {
    if (String(lics[j][1] || '').toLowerCase().trim() !== email) continue;
    shLic.getRange(j + 1, 3).setValue('');
    shLic.getRange(j + 1, 4).setValue('active');
    licFound = true;
    break;
  }
  if (!licFound) {
    var key = 'JGD-' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
    shLic.appendRow([key, email, '', 'active', new Date(), new Date(), '']);
  }

  if (typeof _invalidateAuthCache_ === 'function') _invalidateAuthCache_(email);
  logAudit('APPROVE_TRIAL_ACCOUNT', email, 'Aktivasi lifetime oleh SuperAdmin');
  _sendLifetimeActivationEmail_(email);
  return { success: true, email: email };
}

function updateUser(email, payload){ 
  if(!isSuperAdmin()) throw new Error('Akses ditolak'); 
  const sh = sheet('USERS'); 
  const data = sh.getDataRange().getValues(); 
  email = email.toLowerCase().trim(); 
  for(let i=1;i<data.length;i++){ 
    if(String(data[i][0]).toLowerCase().trim() === email){ 
      if(payload.role){
        sh.getRange(i+1,2).setValue(payload.role);
        logAudit('UPDATE_ROLE', email, payload.role);
        _sendRoleChangeEmail_(email, payload.role);
      }
      if(payload.status){
        sh.getRange(i+1,3).setValue(payload.status);
        const licSheet = sheet('LICENSES');
        const lic = licSheet.getDataRange().getValues();
        let licFound = false;
        for(let j=1;j<lic.length;j++){
          if(String(lic[j][1]).toLowerCase() === email){
            licSheet.getRange(j+1,4).setValue(payload.status);
            if(payload.status === 'active') licSheet.getRange(j+1,3).setValue('');
            licFound = true;
          }
        }
        if(!licFound && (payload.status === 'active' || payload.status === 'inactive')){
          const key = 'JGD-' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
          licSheet.appendRow([key, email, payload.status === 'active' ? '' : '', payload.status, new Date(), '', '']);
        }
        if (payload.status === 'active') {
          try {
            if (typeof _autoProvisionUserSpreadsheet_ === 'function') {
              _autoProvisionUserSpreadsheet_(email);
            }
          } catch (e) {}
          _sendLifetimeActivationEmail_(email);
        }
        logAudit('UPDATE_STATUS', email, payload.status);
      }
      if (typeof _invalidateAuthCache_ === 'function') _invalidateAuthCache_(email);
      return true;
    }
  }
  throw new Error('User tidak ditemukan');
}

/**
 * updateUserSekolah(email, kodeSekolah)
 * SuperAdmin only — kelompokkan akun guru/kepsek ke sekolah tertentu.
 * Dipakai supaya satu deployment aplikasi ini bisa melayani beberapa
 * sekolah sekaligus: Rekap Sekolah & Rekap Guru Wali milik seorang
 * Kepsek cuma menghitung guru dengan kode_sekolah yang SAMA dengan
 * Kepsek itu sendiri (lihat _kepsekActiveGuruEmails_ di Kepsek.js).
 * kodeSekolah boleh nama sekolah biasa (bukan kode formal) — cuma
 * dipakai sebagai label pengelompokan, bukan divalidasi ke database
 * sekolah resmi manapun.
 */
function updateUserSekolah(email, kodeSekolah) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  email = String(email || '').toLowerCase().trim();
  if (!email) throw new Error('Email tidak valid');
  var kode = String(kodeSekolah || '').trim();

  var sh = _ensureUsersKodeSekolahColumn_();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').toLowerCase().trim() !== email) continue;
    sh.getRange(i + 1, 5).setValue(kode);
    logAudit('UPDATE_KODE_SEKOLAH', email, kode || '(dikosongkan)');
    if (typeof _invalidateAuthCache_ === 'function') _invalidateAuthCache_(email);
    return { success: true, email: email, kode_sekolah: kode };
  }
  throw new Error('User tidak ditemukan');
}

function getUsers(){
  if(!isSuperAdmin()) throw new Error('Akses ditolak'); 
  const users = sheet('USERS').getDataRange().getValues(); 
  if(users.length <= 1) return []; 
  const licenses = sheet('LICENSES').getDataRange().getValues(); 
  const licMap = {}; 
  for(let i=1;i<licenses.length;i++){ 
    const email = String(licenses[i][1] || '').toLowerCase(); 
    if(!email) continue; 
    licMap[email] = { 
      key: licenses[i][0], 
      expired: licenses[i][2] ? Utilities.formatDate(new Date(licenses[i][2]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null 
    }; 
  } 
  return users.slice(1).map((u,i)=>{ 
    const email = String(u[0] || '').toLowerCase(); 
    const role = String(u[1] || '-').toLowerCase();
    const lic = licMap[email] || {}; 
    let sisaHari = null; 
    if(role !== 'superadmin' && lic.expired){ 
      sisaHari = Math.ceil((new Date(lic.expired) - new Date()) / 86400000); 
    } 
    return { 
      no: i+1, email: email, role: u[1] || '-', status: u[2] || 'inactive', 
      dibuat: u[3] ? Utilities.formatDate(new Date(u[3]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '-', 
      nama_guru: '-', sekolah: '-', licenseKey: lic.key || null, expired: role === 'superadmin' ? '-' : (lic.expired || '-'), sisaHari: role === 'superadmin' ? null : sisaHari 
    }; 
  }); 
} 

function logAudit(action, target, detail){
  var sh = sheet('AUDIT_LOG');
  sh.appendRow([ new Date(), getLoginEmail(), action, target, detail ]);
  trimLogSheetIfNeeded_(sh, 300);
}

function ensureUserOnce(email){ 
  const sh = sheet('USERS'); 
  const data = sh.getDataRange().getValues(); 
  email = String(email).toLowerCase().trim(); 
  for(let i=1;i<data.length;i++){ 
    if(String(data[i][0]).toLowerCase().trim() === email){ 
      return i + 1; 
    } 
  } 
  throw new Error('Akun Anda belum terdaftar.\nSilakan hubungi Admin di WA 081916527525'); 
} 

function deleteUser(email){ 
  const auth = getAuth(); 
  if(!isSuperAdmin()){ throw new Error('Akses ditolak'); } 
  if(!email){ throw new Error('Email tidak valid'); } 
  const emailNorm = String(email).toLowerCase().trim(); 
  if(emailNorm === auth.email){ throw new Error('Tidak boleh menghapus akun sendiri'); } 
  const sh = sheet('USERS'); 
  const data = sh.getDataRange().getValues(); 
  for(let i=1;i<data.length;i++){ 
    const rowEmail = String(data[i][0] || '').toLowerCase().trim(); 
    const role = String(data[i][1] || '').toLowerCase(); 
    if(rowEmail === emailNorm){ 
      if(role === 'superadmin'){ throw new Error('Akun superadmin tidak boleh dihapus'); } 
      logAudit('DELETE_USER', emailNorm, `deleted_by=${auth.email}`); 
      sh.deleteRow(i+1); 
      return true; 
    } 
  } 
  throw new Error('User tidak ditemukan');
}

/**
 * deleteInactiveUsers()
 * Hapus SEMUA akun berstatus 'inactive' sekaligus (SuperAdmin only).
 * Sama seperti deleteUser(): akun superadmin dan akun yang sedang login
 * tidak pernah ikut terhapus, apapun statusnya.
 */
function deleteInactiveUsers(){
  const auth = getAuth();
  if(!isSuperAdmin()){ throw new Error('Akses ditolak'); }

  const sh = sheet('USERS');
  const data = sh.getDataRange().getValues();
  const deletedEmails = [];

  for(let i = data.length - 1; i >= 1; i--){
    const rowEmail = String(data[i][0] || '').toLowerCase().trim();
    const role = String(data[i][1] || '').toLowerCase().trim();
    const status = String(data[i][2] || '').toLowerCase().trim();
    if(status !== 'inactive') continue;
    if(role === 'superadmin') continue;
    if(rowEmail === auth.email) continue;
    sh.deleteRow(i + 1);
    deletedEmails.push(rowEmail);
  }

  if(deletedEmails.length){
    logAudit('DELETE_INACTIVE_USERS', auth.email, deletedEmails.length + ' user: ' + deletedEmails.join(', '));
  }
  return { success: true, deleted: deletedEmails.length, emails: deletedEmails };
}
