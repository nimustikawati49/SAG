/**
 * Setting.js — Pengaturan Profil Guru
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

function saveSetting(data){

  ensureAcademicSchema_();

  const auth = getAuth();
  if(auth.role !== 'admin' && auth.role !== 'superadmin' && auth.role !== 'kepsek'){
    throw new Error('AKSES_DITOLAK');
  }

  if(auth.role !== 'superadmin'){

    const lic = getLicenseByEmail(auth.email);

    if(!lic){
      throw new Error('INPUT_LICENSE');
    }

    if(lic.status !== 'active'){
      throw new Error('INPUT_LICENSE');
    }

    if(!lic.expired || new Date(lic.expired) < new Date()){
      throw new Error('LISENSI_EXPIRED');
    }
  }

  const email = auth.email;
  const sh = sheet('SETTING');
  const values = sh.getDataRange().getValues();

  if(values.length === 0){
    throw new Error('Header SETTING belum ada');
  }

  const header = values[0].map(h => String(h).toLowerCase().trim());
  const idx = {};
  header.forEach((h,i)=> idx[h] = i);

  const payload = {
    sekolah     : data.sekolah || '',
    tahun       : data.tahun || data.tahun_pelajaran || '',
    semester    : data.semester || '',
    guru        : data.nama_guru || data.guru || '',
    nip         : data.nip_guru || data.nip || '',
    mapel       : Array.isArray(data.mata_pelajaran)
                    ? data.mata_pelajaran.join(',')
                    : (data.mata_pelajaran || data.mapel || ''),
    logo        : data.app_logo_url || data.logo || ''
  };

  if(!payload.tahun){
    throw new Error('Field tahun belum diisi');
  }

  // Pastikan master tahun pelajaran tersedia dan dapat dipakai lintas tahun.
  ensureAcademicYearExists_(payload.tahun, payload.semester || 'Ganjil', false);

  let rowIndex = -1;

  for(let i=1;i<values.length;i++){
    if(String(values[i][0]).toLowerCase().trim() === email){
      rowIndex = i + 1;
      break;
    }
  }

  const rowData = header.map(h=>{
    if(h === 'email') return email;
    if(h === 'lock_status') return 'LOCKED';
    return payload[h] ?? '';
  });

  if(rowIndex > 0){
    sh.getRange(rowIndex,1,1,rowData.length).setValues([rowData]);

    const taAktifCol = header.indexOf('tahun_pelajaran_aktif');
    const semAktifCol = header.indexOf('semester_aktif');
    if (taAktifCol > -1) sh.getRange(rowIndex, taAktifCol + 1).setValue(payload.tahun);
    if (semAktifCol > -1) sh.getRange(rowIndex, semAktifCol + 1).setValue(payload.semester || 'Ganjil');
  }else{
    sh.appendRow(rowData);

    const newRow = sh.getLastRow();
    const taAktifCol = header.indexOf('tahun_pelajaran_aktif');
    const semAktifCol = header.indexOf('semester_aktif');
    if (taAktifCol > -1) sh.getRange(newRow, taAktifCol + 1).setValue(payload.tahun);
    if (semAktifCol > -1) sh.getRange(newRow, semAktifCol + 1).setValue(payload.semester || 'Ganjil');
  }

  logAudit('SAVE_SETTING', email, 'Setting disimpan & dikunci');
  invalidateCache_('SETTING');
  invalidateDashboardCache_();
  trySyncGuruSummaryAfterMutation_(email, 'SAVE_SETTING');

  return {
    success:true,
    locked:true
  };
}

function getSetting(){

  ensureAcademicSchema_();

  const email = getLoginEmail();
  // Use cached sheet data (TTL 60s) — invalidated on saveSetting
  const values = getSheetCached('SETTING', 60);

  if (values.length < 2) return {};

  const header = values[0].map(h => String(h).toLowerCase().trim());
  const idx = {};
  header.forEach((h,i)=> idx[h] = i);

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][idx.email]).toLowerCase().trim() === email) {

      const lockIdx = idx.lock_status !== undefined ? idx.lock_status : (header.length - 1);
      const lockStatus = String(values[i][lockIdx] || '').toUpperCase() === 'LOCKED';
      const mapelRaw   = String(values[i][idx.mapel] || '');
      const mapelList  = mapelRaw ? mapelRaw.split(',').map(function(m){ return m.trim(); }).filter(Boolean) : [];

      return {
        sekolah : values[i][idx.sekolah] || '',
        tahun_pelajaran : (idx.tahun_pelajaran_aktif > -1 ? values[i][idx.tahun_pelajaran_aktif] : '') || values[i][idx.tahun] || '',
        semester : (idx.semester_aktif > -1 ? values[i][idx.semester_aktif] : '') || values[i][idx.semester] || '',
        nama_guru : values[i][idx.guru] || '',
        nip_guru : values[i][idx.nip] || '',
        mata_pelajaran : mapelRaw,
        mapel_list : mapelList,
        app_logo_url : values[i][idx.logo] || '',
        lock_status : lockStatus,
        tahun_pelajaran_aktif : idx.tahun_pelajaran_aktif > -1 ? values[i][idx.tahun_pelajaran_aktif] : (values[i][idx.tahun] || ''),
        semester_aktif : idx.semester_aktif > -1 ? values[i][idx.semester_aktif] : (values[i][idx.semester] || '')
      };
    }
  }

  const p = getUserAcademicPeriod(email);
  return {
    tahun_pelajaran : p.tahun_pelajaran || '',
    semester : p.semester || '',
    tahun_pelajaran_aktif : p.tahun_pelajaran || '',
    semester_aktif : p.semester || ''
  };
}

function isSettingLocked(){

  const auth = getAuth();
  const sh = sheet('SETTING');
  const data = sh.getDataRange().getValues();

  if(data.length < 2) return false;

  const header = data[0].map(h=>String(h).toLowerCase().trim());
  const lockIndex = header.indexOf('lock_status');

  if(lockIndex === -1) return false;

  for(let i=1;i<data.length;i++){
    if(String(data[i][0]).toLowerCase().trim() === auth.email){
      return data[i][lockIndex] === 'LOCKED';
    }
  }

  return false;
}

function unlockSetting(){

  const auth = getAuth();

  if(auth.role !== 'admin' && auth.role !== 'superadmin'){
    throw new Error('Akses ditolak');
  }

  const sh = sheet('SETTING');
  const data = sh.getDataRange().getValues();
  const header = data[0].map(h=>String(h).toLowerCase().trim());

  const lockIndex = header.indexOf('lock_status');
  if(lockIndex === -1){
    throw new Error('Kolom lock_status tidak ditemukan');
  }

  for(let i=1;i<data.length;i++){
    if(String(data[i][0]).toLowerCase().trim() === auth.email){
      sh.getRange(i+1, lockIndex+1).setValue('UNLOCKED');
      logAudit('UNLOCK_SETTING', auth.email, 'Setting dibuka');
      return true;
    }
  }

  throw new Error('Setting tidak ditemukan');
}

function updateSemesterAktif(semester){
  ensureAcademicSchema_();
  const auth = getAuth();
  if(auth.role !== 'admin' && auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }
  if(!semester) throw new Error('Semester tidak boleh kosong');
  const sh = sheet('SETTING');
  const data = sh.getDataRange().getValues();
  if(data.length === 0) throw new Error('Sheet SETTING kosong');
  const header = data[0].map(h=>String(h).toLowerCase().trim());
  const semIndex = header.indexOf('semester');
  const semAktifIndex = header.indexOf('semester_aktif');
  if(semIndex === -1) throw new Error('Kolom semester tidak ditemukan di SETTING');
  const email = auth.email;
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]).toLowerCase().trim() === email){
      sh.getRange(i+1, semIndex+1).setValue(semester);
      if (semAktifIndex > -1) sh.getRange(i+1, semAktifIndex+1).setValue(semester);
      logAudit('UPDATE_SEMESTER', email, semester);
      trySyncGuruSummaryAfterMutation_(email, 'UPDATE_SEMESTER');
      return true;
    }
  }
  throw new Error('Setting tidak ditemukan untuk akun ini');
}

function cekAutoLockStatus(){
  const isLocked = isSettingLocked();
  return { shouldLock: isLocked };
}

function uploadLogoSekolah(payload){

  assertLicenseActive();

  const auth = getAuth();
  const email = auth.email;

  const folder = getUserResourceFolder_(email, 'logo_folder', 'LOGO_JURNAL');

  const setting = getSetting();

  if(setting.app_logo_url){
    try{
      const match = setting.app_logo_url.match(/[-\w]{25,}/);
      if(match){
        DriveApp.getFileById(match[0]).setTrashed(true);
      }
    }catch(e){ console.error('[JGD] trash logo lama:', e.message||e); }
  }

  const base64 = payload.data.split(',')[1];

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    payload.type,
    payload.name
  );

  const file = folder.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  const fileId = file.getId();

  const stableUrl = "https://drive.google.com/thumbnail?id=" 
                    + fileId + "&sz=w300";

  logAudit('UPLOAD_LOGO', email, fileId);

  return stableUrl;
}
