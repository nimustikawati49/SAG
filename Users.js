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
  var sh = sheet('USERS');
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
 * getPendingTrialAccounts()
 * SuperAdmin only — daftar akun guru yang masih berstatus 'trial',
 * diurutkan yang paling mendesak (sudah expired) di atas.
 */
function getPendingTrialAccounts() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh = sheet('USERS');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][2] || '').toLowerCase().trim();
    if (status !== 'trial') continue;
    var email = String(data[i][0] || '').toLowerCase().trim();
    if (!email) continue;
    var info = _computeTrialInfo_(data[i][3]);
    result.push({
      email: email,
      role: String(data[i][1] || 'admin').toLowerCase(),
      dibuat: data[i][3] ? Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '-',
      days_left: info.daysLeft,
      expired: info.expired
    });
  }
  result.sort(function(a, b) {
    if (a.expired === b.expired) return a.days_left - b.days_left;
    return a.expired ? -1 : 1;
  });
  return result;
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

  var shLic = sheet('LICENSES');
  var lics = shLic.getDataRange().getValues();
  var licFound = false;
  var exp = new Date();
  exp.setFullYear(exp.getFullYear() + 1);
  for (var j = 1; j < lics.length; j++) {
    if (String(lics[j][1] || '').toLowerCase().trim() !== email) continue;
    shLic.getRange(j + 1, 3).setValue(exp);
    shLic.getRange(j + 1, 4).setValue('active');
    licFound = true;
    break;
  }
  if (!licFound) {
    var key = 'JGD-' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
    shLic.appendRow([key, email, exp, 'active', new Date(), new Date(), '']);
  }

  if (typeof _invalidateAuthCache_ === 'function') _invalidateAuthCache_(email);
  logAudit('APPROVE_TRIAL_ACCOUNT', email, 'Aktivasi penuh oleh SuperAdmin');
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
      } 
      if(payload.status){ 
        sh.getRange(i+1,3).setValue(payload.status); 
        const lic = sheet('LICENSES').getDataRange().getValues(); 
        for(let j=1;j<lic.length;j++){ 
          if(String(lic[j][1]).toLowerCase() === email){ 
            sheet('LICENSES').getRange(j+1,4).setValue(payload.status); 
          } 
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
 * Set tanggal expired lisensi untuk user tertentu.
 * expiredDate: string 'yyyy-MM-dd' atau '' (lifetime/kosong)
 */
function setUserExpiry(email, expiredDate) {
  if (!isSuperAdmin()) throw new Error('Akses ditolak');
  email = String(email).trim().toLowerCase();
  if (!email) throw new Error('Email tidak valid');

  const shLic = sheet('LICENSES');
  const lics  = shLic.getDataRange().getValues();
  for (let i = 1; i < lics.length; i++) {
    if (String(lics[i][1] || '').toLowerCase().trim() === email) {
      const expVal = expiredDate ? new Date(expiredDate) : '';
      shLic.getRange(i + 1, 3).setValue(expVal);
      // Pastikan status active
      shLic.getRange(i + 1, 4).setValue('active');
      logAudit('SET_EXPIRY', email,
        expiredDate ? ('expired=' + expiredDate) : 'lifetime');
      return true;
    }
  }
  throw new Error('Lisensi untuk ' + email + ' tidak ditemukan');
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

function createLicense(email){ 
  const sh = sheet('LICENSES'); 
  email = String(email).toLowerCase().trim(); 
  if(!email.includes('@')){ throw new Error('Email tidak valid'); } 
  const key = 'JGD-' + Utilities.getUuid().slice(0,8).toUpperCase(); 
  const now = new Date(); 
  const expired = new Date(now); 
  expired.setFullYear(expired.getFullYear() + 1); 
  sh.appendRow([ key, email, expired, 'inactive', now, '', '' ]); 
  logAudit('CREATE_LICENSE', email, 'Lisensi dibuat (belum aktif)'); 
  return { key, expired }; 
} 

function generateLicense(email, years){
  if(!isSuperAdmin()) throw new Error('Akses ditolak');
  email = String(email).toLowerCase().trim();
  if(!email.includes('@')) throw new Error('Email tidak valid');
  const sh = sheet('LICENSES');
  const rows = sh.getDataRange().getValues();
  const key = 'JGD-' + Utilities.getUuid().slice(0,8).toUpperCase();
  const now = new Date();
  const expired = new Date(now);
  expired.setFullYear(expired.getFullYear() + (Number(years) || 1));
  const expiredStr = Utilities.formatDate(expired, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  for(let i=1; i<rows.length; i++){
    if(String(rows[i][1]).toLowerCase().trim() === email){
      sh.getRange(i+1,1).setValue(key);
      sh.getRange(i+1,3).setValue(expired);
      sh.getRange(i+1,4).setValue('active');
      sh.getRange(i+1,6).setValue(now);
      logAudit('GENERATE_LICENSE', email, key + ' exp:' + expiredStr);
      return { key, expired: expiredStr };
    }
  }
  sh.appendRow([ key, email, expired, 'active', now, now, '' ]);
  logAudit('GENERATE_LICENSE', email, key + ' exp:' + expiredStr);
  return { key, expired: expiredStr };
}

function regenerateLicense(email){

  const auth = getAuth();
  if(auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }

  email = String(email).toLowerCase().trim();

  const sh = sheet('LICENSES');
  const data = sh.getDataRange().getValues();

  for(let i=1;i<data.length;i++){

    if(String(data[i][1]).toLowerCase().trim() === email){

      const newKey = 'JGD-' +
        Utilities.getUuid()
          .replace(/-/g,'')
          .substring(0,10)
          .toUpperCase();

      sh.getRange(i+1,1).setValue(newKey);
      sh.getRange(i+1,3).setValue('');
      sh.getRange(i+1,4).setValue('inactive');
      sh.getRange(i+1,6).setValue('');

      logAudit('REGENERATE_LICENSE', email, newKey);

      return {
        key:newKey
      };
    }
  }

  throw new Error('Lisensi tidak ditemukan');
}

function addAdminUser(payload){

  const auth = getAuth();
  if(auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }

  let email      = '';
  let masaAktif  = '1year';  // default: 1 tahun
  let role       = 'admin';

  if(typeof payload === 'object'){
    email     = payload.email     || '';
    masaAktif = payload.masaAktif || '1year';
    role      = payload.role      || 'admin';
  }else{
    email = payload;
  }

  email = String(email).trim().toLowerCase();

  if(!email){
    throw new Error('Email tidak boleh kosong');
  }

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if(!emailRegex.test(email)){
    throw new Error('Format email tidak valid');
  }

  // Validasi role
  const allowedRoles = ['admin','kepsek'];
  if(!allowedRoles.includes(role)) role = 'admin';

  const shUser = sheet('USERS');
  const shLic  = sheet('LICENSES');

  const users = shUser.getDataRange().getValues();
  for(let i=1;i<users.length;i++){
    if(String(users[i][0]).toLowerCase().trim() === email){
      throw new Error('User sudah terdaftar');
    }
  }

  shUser.appendRow([
    email,
    role,
    'active',
    new Date()
  ]);

  const licenseKey = 'JGD-' +
    Utilities.getUuid()
      .replace(/-/g,'')
      .substring(0,10)
      .toUpperCase();

  // Hitung expired berdasarkan masaAktif
  let expiredValue = '';
  let expiredStr   = 'Seumur Hidup';
  if(masaAktif === '1year'){
    const exp = new Date();
    exp.setFullYear(exp.getFullYear() + 1);
    expiredValue = exp;
    expiredStr   = Utilities.formatDate(exp, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  shLic.appendRow([
    licenseKey,
    email,
    expiredValue,
    'active',
    new Date(),
    new Date(),
    ''
  ]);

  logAudit('ADD_ADMIN_WITH_LICENSE', email,
    licenseKey + ' | role=' + role + ' | masaAktif=' + masaAktif + ' | exp=' + expiredStr);

  return {
    status:  true,
    key:     licenseKey,
    expired: expiredStr,
    role:    role
  };
}
