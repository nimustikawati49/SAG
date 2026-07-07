/**
 * Auth.js — Autentikasi, Otorisasi & Lisensi
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

/** Ubah superadmin email (hanya bisa dipanggil oleh superadmin) */
function setSuperAdminEmail(newEmail) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  if (!newEmail || !newEmail.includes('@')) throw new Error('Email tidak valid');
  PropertiesService.getScriptProperties().setProperty('SUPERADMIN_EMAIL', newEmail.toLowerCase().trim());
  _superAdminEmailCache = null; // reset cache
  logAudit('SET_SUPERADMIN_EMAIL', getLoginEmail(), 'Email baru: ' + newEmail);
  return { status: true, email: newEmail.toLowerCase().trim() };
}

/** Ambil superadmin email aktif untuk ditampilkan di UI */
function getSuperAdminEmailForUI() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  return getSuperAdminEmail_();
}

function getLicenseBadge(){
  const auth = getAuth();
  if(auth.role === 'superadmin' || auth.role === 'kepsek') return null;
  // Delegasi ke License.js (school-wide)
  return getSchoolLicenseBadge_();
} 

function getLoginEmail(){ 
  const email = Session.getEffectiveUser().getEmail(); 
  if(!email){ 
    throw new Error('Tidak dapat mendeteksi email login'); 
  } 
  return email.toLowerCase(); 
} 

function isSuperAdmin() {
  return Session.getEffectiveUser().getEmail().toLowerCase() === getSuperAdminEmail_();
}

function getActiveLicenseByEmail(email){ 
  const lic = getLicenseByEmail(email); 
  if(!lic) return null; 
  if(lic.status !== 'active') return null; 
  if(!lic.expired) return null; 
  if(new Date(lic.expired) < new Date()) return null; 
  return lic; 
} 

function assertLicenseActive(){
  const auth = getAuth();
  // Superadmin & kepsek bypass lisensi
  if(auth.role === 'superadmin' || auth.role === 'kepsek') return true;
  // Cek lisensi per sekolah (school-wide)
  return assertSchoolLicenseActive_();
}

function assertCanWrite_(){ 
  const auth = getAuth(); 
  if(auth.role === 'guest'){ 
    throw new Error('Akun Anda belum terdaftar.\nSilakan hubungi Admin di WA 081916527525'); 
  } 
} 

function getLicenseInfo(){

  const auth = getAuth();
  if(auth.role === 'superadmin') return null;

  const lic = getLicenseByEmail(auth.email);
  if(!lic || !lic.expired) return null;

  const now = new Date();
  const diff = Math.ceil((new Date(lic.expired) - now) / 86400000);

  return {
    status: lic.status,
    expired: lic.expired,
    daysLeft: diff
  };
}

function activateLicense(inputKey){

  const auth = getAuth();
  const email = String(auth.email).toLowerCase().trim();

  if(!inputKey){
    throw new Error('Kode lisensi kosong');
  }

  const shLic = sheet('LICENSES');
  const shUser = sheet('USERS');

  const licData = shLic.getDataRange().getValues();
  const userData = shUser.getDataRange().getValues();

  inputKey = String(inputKey).trim().toUpperCase();

  let found = false;

  for(let i=1;i<licData.length;i++){

    const key = String(licData[i][0] || '').trim().toUpperCase();
    const licEmail = String(licData[i][1] || '').toLowerCase().trim();

    if(key === inputKey && licEmail === email){

      found = true;

      const now = new Date();
      const expired = new Date();
      expired.setFullYear(now.getFullYear() + 1);

      shLic.getRange(i+1,3).setValue(expired);
      shLic.getRange(i+1,4).setValue('active');
      shLic.getRange(i+1,6).setValue(now);

      for(let u=1;u<userData.length;u++){
        if(String(userData[u][0]).toLowerCase().trim() === email){
          shUser.getRange(u+1,3).setValue('active');
          break;
        }
      }

      logAudit('ACTIVATE_LICENSE', email, inputKey);

      return {
        status:true,
        expired: Utilities.formatDate(
          expired,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        )
      };
    }
  }

  if(!found){
    throw new Error('Kode lisensi tidak cocok dengan email login');
  }
}

function renewLicense(email, years = 1){ 
  if(!isSuperAdmin()) throw new Error('Akses ditolak'); 
  email = String(email).toLowerCase().trim(); 
  const lic = getLicenseByEmail(email); 
  if(!lic) throw new Error('Lisensi tidak ditemukan'); 
  const base = lic.expired && lic.expired > new Date() ? new Date(lic.expired) : new Date(); 
  base.setFullYear(base.getFullYear() + years); 
  const sh = sheet('LICENSES'); 
  sh.getRange(lic.row, 3).setValue(base); 
  sh.getRange(lic.row, 4).setValue('active'); 
  sh.getRange(lic.row, 7).setValue(new Date()); 
  logAudit('RENEW_LICENSE', email, `+${years} tahun`); 
  return { expired: base }; 
} 

function sendActivationEmail(email, expired){ 
  const subject = 'Lisensi Jurnal Guru Digital Aktif'; 
  const body = `Halo Ibu/Bapak, \n\nLisensi Jurnal Guru Digital Anda telah berhasil diaktifkan.\n\nDetail:\n- Email : ${email}\n- Status : Aktif\n- Berlaku : sampai ${expired}\n\nSilakan lanjutkan penggunaan aplikasi.\n\nSalam, Jurnal Guru Digital`; 
  GmailApp.sendEmail(email, subject, body); 
} 

function getLicenseByEmail(email){ 
  const sh = sheet('LICENSES'); 
  const rows = sh.getDataRange().getValues(); 
  email = String(email).toLowerCase().trim(); 
  for(let i=1;i<rows.length;i++){ 
    if(String(rows[i][1]).toLowerCase() === email){ 
      return { 
        row: i + 1, key: rows[i][0], email: rows[i][1], 
        expired: rows[i][2] ? new Date(rows[i][2]) : null, 
        status: rows[i][3], created_at: rows[i][4], 
        activated_at: rows[i][5], renewed_at: rows[i][6] 
      }; 
    } 
  } 
  return null; 
} 

function getAuth(){
  try{
    const email = Session.getEffectiveUser().getEmail().toLowerCase();

    if(!email){
      return { email:null, role:'guest', status:'inactive' };
    }

// Rate limiting: max 60 calls per minute per email
try {
  const cache   = CacheService.getScriptCache();
  const rlKey   = 'RL_' + email.replace(/[^a-z0-9]/g, '_');
  const current = parseInt(cache.get(rlKey) || '0', 10);
  if (current >= 60) {
    return { email, role:'guest', status:'rate_limited' };
  }
  cache.put(rlKey, String(current + 1), 60);
} catch(rlErr) { /* cache non-critical, silently ignore */ }

    const sh = sheet('USERS');
    if(!sh){
      return { email, role:'guest', status:'inactive' };
    }

    const data = sh.getDataRange().getValues();

    for(let i=1;i<data.length;i++){
      if(String(data[i][0]).toLowerCase().trim() === email){
        var role = String(data[i][1] || 'admin').toLowerCase();
        // Tier: SA & kepsek selalu SCHOOL, admin tergantung license
        var tier = (role === 'superadmin' || role === 'kepsek') ? 'SCHOOL' : getTier_();
        return {
          email,
          role,
          status: String(data[i][2] || 'inactive').toLowerCase(),
          tier
        };
      }
    }

    return { email, role:'guest', status:'inactive' };

  }catch(e){
    return { email:null, role:'guest', status:'inactive' };
  }
}

function findUserByEmail(email){
  try{
    if(!email) return null;

    const sh = sheet('USERS');
    if(!sh) return null;

    const data = sh.getDataRange().getValues();
    if(data.length < 2) return null;

    email = String(email).toLowerCase().trim();

    for(let i=1;i<data.length;i++){
      const rowEmail = String(data[i][0] || '').toLowerCase().trim();
      if(rowEmail === email){
        return {
          email  : rowEmail,
          role   : data[i][1] || 'admin',
          status : data[i][2] || 'inactive',
          dibuat : data[i][3] || null
        };
      }
    }
    return null;

  }catch(err){
    return null;
  }
}

function checkLicenseExpiryReminder(){ 
  const sh = sheet('LICENSES'); 
  const data = sh.getDataRange().getValues(); 
  const HARI_REMINDER = 60; 
  const now = new Date(); 
  let listAdmin = []; 
  for(let i=1;i<data.length;i++){ 
    const email = String(data[i][1] || '').toLowerCase().trim(); 
    const expired = data[i][2]; 
    const status = data[i][3]; 
    if(status !== 'active' || !expired) continue; 
    const sisaHari = daysBetween(now, new Date(expired)); 
    if(sisaHari > 0 && sisaHari <= HARI_REMINDER){ 
      listAdmin.push({ email, expired, sisaHari }); 
      MailApp.sendEmail({ 
        to: email, 
        subject: '\u23F0 Lisensi Jurnal Guru Digital Akan Berakhir', 
        htmlBody: `<p>Lisensi Anda akan berakhir dalam <b>${sisaHari} hari</b>.</p><p>Tanggal: <b>${Utilities.formatDate(new Date(expired), Session.getScriptTimeZone(),'yyyy-MM-dd')}</b></p>` 
      }); 
    } 
  } 
  if(listAdmin.length){ 
    let html = `<h3>\u23F0 Reminder Lisensi</h3><table border="1">`; 
    listAdmin.forEach(l => { html += `<tr><td>${l.email}</td><td>${l.sisaHari} hari</td></tr>`; }); 
    html += `</table>`; 
    MailApp.sendEmail({ to: getSuperAdminEmail_(), subject: '\u23F0 Rekap Reminder Lisensi', htmlBody: html }); 
  } 
}

function checkLicenseAccess(){

  const auth = getAuth();

  if(auth.role === 'superadmin'){
    return { allowed:true };
  }

  if(auth.role === 'guest'){
    return {
      allowed:false,
      reason:'AKUN_TIDAK_TERDAFTAR'
    };
  }

  const lic = getLicenseByEmail(auth.email);

  if(!lic){
    return {
      allowed:false,
      reason:'INPUT_LICENSE'
    };
  }

  if(lic.status !== 'active'){
    return {
      allowed:false,
      reason:'INPUT_LICENSE'
    };
  }

  if(!lic.expired || new Date(lic.expired) < new Date()){
    return {
      allowed:false,
      reason:'EXPIRED'
    };
  }

  return { allowed:true };
}
