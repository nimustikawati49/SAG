/**
 * SuperAdmin.js — Fitur Khusus SuperAdmin
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

function getSuperAdminDashboard(){
  const users = sheet('USERS').getDataRange().getValues().slice(1);
  const jurnal = sheet('JURNAL').getDataRange().getValues().slice(1);

  return {
    totalUser: users.length,
    guruAktif: users.filter(u=>u[2]==='active').length,
    guruNonaktif: users.filter(u=>u[2]!=='active').length,
    totalJurnal: jurnal.length,
    totalKelas: [...new Set(jurnal.map(r=>r[2]))].length
  };
}

function getStatistikGlobal(){
  if(!isSuperAdmin()) throw new Error('Akses ditolak');

  const jurnalData = sheet('JURNAL').getDataRange().getValues().slice(1);
  const settingData = sheet('SETTING').getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  const nameMap = {};
  if(settingData.length > 1){
    const hdr = settingData[0].map(h => String(h).toLowerCase().trim());
    const emailIdx = hdr.indexOf('email');
    const namaIdx  = hdr.indexOf('guru');
    settingData.slice(1).forEach(r => {
      const e = String(r[emailIdx] || '').toLowerCase().trim();
      if(e) nameMap[e] = String(r[namaIdx] || '');
    });
  }

  const map = {};
  jurnalData.forEach(r => {
    const e = String(r[12] || '').toLowerCase().trim();
    if(!e) return;
    if(!map[e]) map[e] = { kelas: new Set(), count: 0, last: null };
    map[e].count++;
    if(r[2]) map[e].kelas.add(String(r[2]));
    const d = r[1] ? new Date(r[1]) : null;
    if(d && (!map[e].last || d > map[e].last)) map[e].last = d;
  });

  const perGuru = Object.keys(map).map((e, i) => ({
    no          : i + 1,
    email       : e,
    nama_guru   : nameMap[e] || '-',
    totalJurnal : map[e].count,
    totalKelas  : map[e].kelas.size,
    lastActivity: map[e].last
      ? Utilities.formatDate(map[e].last, tz, 'dd/MM/yyyy')
      : '-'
  }));

  perGuru.sort((a, b) => b.totalJurnal - a.totalJurnal);
  perGuru.forEach((g, i) => g.no = i + 1);

  return {
    perGuru,
    ringkasan: {
      totalGuru: perGuru.length,
      guruAktif: perGuru.filter(function(g) { return g.totalJurnal > 0; }).length,
      topGuru: perGuru.slice(0, 5)
    }
  };
}

function getDaftarArsip(){
  const auth = getAuth();
  const result = [];
  try{
    const root = _getArsipRootFolder_();
    const safeEmail = auth.email.replace(/[@.]/g, '_');
    const guruIt = root.getFoldersByName(safeEmail);
    if(!guruIt.hasNext()) return result;
    const guruFolder = guruIt.next();

    const subIt = guruFolder.getFolders();
    while(subIt.hasNext()){
      const tahunFolder = subIt.next();
      const tahunName   = tahunFolder.getName();
      const filesIt = tahunFolder.getFiles();
      while(filesIt.hasNext()){
        const f = filesIt.next();
        try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
        result.push({ tahun: tahunName, nama: f.getName(), url: f.getDownloadUrl() });
      }
    }

    const directIt = guruFolder.getFiles();
    while(directIt.hasNext()){
      const f = directIt.next();
      const name = f.getName();
      let tahun = name.replace(safeEmail+'_','').replace(/\.\w+$/,'');
      if(tahun.startsWith('JURNAL_WALI_')) tahun = tahun.replace('JURNAL_WALI_','');
      try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
      result.push({ tahun, nama: name, url: f.getDownloadUrl() });
    }
  }catch(e){ console.error('[JGD] getDaftarArsip:', e.message||e); }
  result.sort((a,b) => b.tahun.localeCompare(a.tahun));
  return result;
}

function getDaftarArsipSemuaUser(){
  if(!isSuperAdmin()) throw new Error('Akses ditolak');
  const result = [];
  try{
    const root = DriveApp.getFoldersByName('JURNAL_ARSIP');
    if(!root.hasNext()) return result;
    const subfolders = root.next().getFolders();
    while(subfolders.hasNext()){
      const folder = subfolders.next();
      const safeEmail = folder.getName();
      const email = safeEmail.replace(/_([^_]+)$/, (m, p) => '@' + p).replace(/_/g, '.');
      const files = folder.getFiles();
      while(files.hasNext()){
        const f = files.next();
        const name = f.getName();
        const tahun = name.replace(safeEmail + '_', '').replace('.zip', '');
        try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
        result.push({ email: safeEmail, tahun, url: f.getDownloadUrl() });
      }
    }
  }catch(e){ console.error('[JGD] getDaftarArsipSemuaUser:', e.message||e); }
  return result.reverse();
}

function backupSemesterToSheet(){
  assertLicenseActive();
  const auth    = getAuth();
  const setting = getSetting();
  const tahun   = setting.tahun_pelajaran || '-';
  const semester= setting.semester || '-';
  const email   = auth.email;
  const now     = new Date();
  const tz      = Session.getScriptTimeZone();

  const jurnalData = sheet('JURNAL').getDataRange().getValues().slice(1);
  const filtered = jurnalData.filter(r => r[12] === email);

  if(filtered.length === 0){
    throw new Error('Tidak ada jurnal untuk dibackup semester ini');
  }

  const ss  = getSpreadsheet_();
  const sheetName = 'ARSIP';
  let archiveSh = ss.getSheetByName(sheetName);
  if(!archiveSh){
    archiveSh = ss.insertSheet(sheetName);
    archiveSh.appendRow([
      'backup_time','email_guru','tahun','semester',
      'tanggal','kelas','materi','jumlah_siswa'
    ]);
  }

  const backupTime = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
  filtered.forEach(r => {
    const tgl = r[1] ? Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM-dd') : '-';
    archiveSh.appendRow([
      backupTime, email, tahun, semester,
      tgl, r[2]||'-', r[5]||'-', r[11]||0
    ]);
  });

  logAudit('BACKUP_SEMESTER', email, semester + ' ' + tahun + ' (' + filtered.length + ' jurnal)');
  return { status: true, backed: filtered.length, semester, tahun };
}

function getAuditLog(limit, filterText){
  if(!isSuperAdmin()) throw new Error('Akses ditolak');
  const sh = sheet('AUDIT_LOG');
  if(!sh) return [];
  const data = sh.getDataRange().getValues();
  if(data.length < 2) return [];
  const max = Number(limit) || 200;
  const q = filterText ? String(filterText).toLowerCase().trim() : '';
  let rows = data.slice(1);
  if(q){
    rows = rows.filter(r =>
      String(r[1]||'').toLowerCase().includes(q) ||
      String(r[2]||'').toLowerCase().includes(q) ||
      String(r[3]||'').toLowerCase().includes(q) ||
      String(r[4]||'').toLowerCase().includes(q)
    );
  }
  return rows.slice(-max).reverse().map(r => ({
    waktu  : r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '-',
    actor  : r[1] || '-',
    action : r[2] || '-',
    target : r[3] || '-',
    detail : r[4] || '-'
  }));
}

/**
 * getAuditLogAll() — Semua baris AUDIT_LOG untuk export (SuperAdmin only)
 */
function getAuditLogAll(){
  if(!isSuperAdmin()) throw new Error('Akses ditolak');
  const sh = sheet('AUDIT_LOG');
  if(!sh) return [];
  const data = sh.getDataRange().getValues();
  if(data.length < 2) return [];
  const tz = Session.getScriptTimeZone();
  return data.slice(1).map(r => ({
    waktu  : r[0] ? Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd HH:mm:ss') : '-',
    actor  : String(r[1] || '-'),
    action : String(r[2] || '-'),
    target : String(r[3] || '-'),
    detail : String(r[4] || '-')
  }));
}

function getAuditByEmail(email) {
  if (!isSuperAdmin()) throw new Error('Akses ditolak');
  const sh = sheet('AUDIT_LOG');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  email = String(email).toLowerCase().trim();
  const rows = data.slice(1).filter(r =>
    String(r[1] || '').toLowerCase().trim() === email ||
    String(r[3] || '').toLowerCase().trim() === email
  );
  return rows.slice(-100).reverse().map(r => ({
    waktu : r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '-',
    action: r[2] || '-',
    detail: r[4] || '-'
  }));
}

function arsipTahunPelajaran(){

  assertLicenseActive();

  const auth    = getAuth();
  const setting = getSetting();
  const tahun   = setting.tahun_pelajaran;
  const email   = auth.email;
  const tz      = Session.getScriptTimeZone();

  if(!tahun) throw new Error('Tahun pelajaran belum diset di pengaturan');

  const jurnalSheet = sheet('JURNAL');
  const absSheet    = sheet('ABSENSI');
  const jurnalData  = jurnalSheet.getDataRange().getValues();
  const absRaw      = absSheet ? absSheet.getDataRange().getValues() : [];

  const jurnalGanjil = [];
  const jurnalGenap  = [];
  const jurnalIdSet  = new Set();

  for(let i = 1; i < jurnalData.length; i++){
    const r = jurnalData[i];
    if(r[12] !== email) continue;
    const rowTahun = r[18] || tahun;
    if(rowTahun !== tahun) continue;
    jurnalIdSet.add(String(r[0]));
    const sem = String(r[13]||'').toLowerCase().trim();
    if(sem === 'genap' || sem === 'ii' || sem === '2')
      jurnalGenap.push(r);
    else
      jurnalGanjil.push(r);
  }

  if(jurnalIdSet.size === 0)
    throw new Error('Tidak ada jurnal untuk tahun pelajaran ' + tahun);

  const absMap = {};
  absRaw.slice(1).forEach(a => {
    const jid = String(a[0]);
    if(!jurnalIdSet.has(jid)) return;
    if(!absMap[jid]) absMap[jid] = [];
    absMap[jid].push(a);
  });

  const safeTahun = tahun.replace(/[\/\\:*?\[\]]/g, '-');
  const tempSS    = SpreadsheetApp.create('TEMP_ARSIP_' + safeTahun);
  const tempId    = tempSS.getId();

  try{
    const shJG = tempSS.getActiveSheet();
    shJG.setName('Jurnal_Ganjil');
    fillArsipJurnalSheet_(shJG, jurnalGanjil, 'GANJIL', setting, tahun, tz);

    const shAG = tempSS.insertSheet('Absensi_Ganjil');
    fillArsipAbsensiSheet_(shAG, jurnalGanjil, 'GANJIL', setting, tahun, tz, absMap);

    if(jurnalGenap.length > 0){
      const shJGn = tempSS.insertSheet('Jurnal_Genap');
      fillArsipJurnalSheet_(shJGn, jurnalGenap, 'GENAP', setting, tahun, tz);

      const shAGn = tempSS.insertSheet('Absensi_Genap');
      fillArsipAbsensiSheet_(shAGn, jurnalGenap, 'GENAP', setting, tahun, tz, absMap);
    }

    SpreadsheetApp.flush();

    const xlsBlob = DriveApp.getFileById(tempId)
      .getBlob()
      .setName('Rekap_' + safeTahun + '.xlsx');

    const rootArsip = _getArsipRootFolder_();
    const safeEmail   = email.replace(/[@.]/g, '_');
    const guruFolder  = rootArsip.getFoldersByName(safeEmail).hasNext()
      ? rootArsip.getFoldersByName(safeEmail).next()
      : rootArsip.createFolder(safeEmail);
    const tahunFolder = guruFolder.getFoldersByName(safeTahun).hasNext()
      ? guruFolder.getFoldersByName(safeTahun).next()
      : guruFolder.createFolder(safeTahun);

    const xlsFile = tahunFolder.createFile(xlsBlob);
    xlsFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  }finally{
    try{ DriveApp.getFileById(tempId).setTrashed(true); }catch(e){}
  }

  const absAll = absSheet ? absSheet.getDataRange().getValues() : [];
  for(let i = absAll.length - 1; i >= 1; i--){
    if(jurnalIdSet.has(String(absAll[i][0]))) absSheet.deleteRow(i + 1);
  }

  const jAll = jurnalSheet.getDataRange().getValues();
  for(let i = jAll.length - 1; i >= 1; i--){
    const r = jAll[i];
    if(r[12] !== email) continue;
    if((r[18] || tahun) === tahun) jurnalSheet.deleteRow(i + 1);
  }

  const settingSheet = sheet('SETTING');
  const settingData  = settingSheet.getDataRange().getValues();
  const hdr    = settingData[0].map(h => String(h).toLowerCase().trim());
  const tIdx   = hdr.indexOf('tahun');
  const semIdx = hdr.indexOf('semester');
  const lkIdx  = hdr.indexOf('lock_status');
  for(let i = 1; i < settingData.length; i++){
    if(String(settingData[i][0]).toLowerCase().trim() !== email) continue;
    if(tIdx   > -1) settingSheet.getRange(i+1, tIdx+1).setValue('');
    if(semIdx > -1) settingSheet.getRange(i+1, semIdx+1).setValue('');
    if(lkIdx  > -1) settingSheet.getRange(i+1, lkIdx+1).setValue('UNLOCK');
    break;
  }

  invalidateCache_('SETTING');
  invalidateCache_('JURNAL');
  invalidateCache_('ABSENSI');

  logAudit('ARSIP_TAHUN', email, tahun + ' | ' + jurnalIdSet.size + ' jurnal');

  return {
    success     : true,
    tahun,
    jumlahJurnal: jurnalIdSet.size
  };
}

function getStatusArsipTahunan(){

  const auth        = getAuth();
  const setting     = getSetting();
  const tahun       = setting.tahun_pelajaran;
  const email       = auth.email;
  const activeTahun = tahun || '';

  const jurnal = sheet('JURNAL').getDataRange().getValues().slice(1);

  let adaGanjil = false;
  let adaGenap  = false;

  jurnal.forEach(r => {
    if(r[12] !== email) return;
    if(activeTahun && r[18] && String(r[18]) !== activeTahun) return;
    const sem = String(r[13]||'').toLowerCase().trim();
    if(sem === 'ganjil' || sem === 'i'  || sem === '1') adaGanjil = true;
    if(sem === 'genap'  || sem === 'ii' || sem === '2') adaGenap  = true;
  });

  const safeEmail = email.replace(/[@.]/g,'_');
  const safeTahun = tahun ? tahun.replace(/[\/\\:*?\[\]]/g,'-') : '';
  let sudahArsip = false;

  try{
    const rootIt = _getArsipRootFolder_();
    const guruIt = rootIt.getFoldersByName(safeEmail);
    if(guruIt.hasNext()){
      const gf = guruIt.next();
      if(safeTahun && gf.getFoldersByName(safeTahun).hasNext()) sudahArsip = true;
      if(!sudahArsip && gf.getFilesByName(safeEmail+'_'+tahun+'.zip').hasNext()) sudahArsip = true;
    }
  }catch(e){ console.error('[JGD] cek status arsip:', e.message||e); }

  return { tahun, adaGanjil, adaGenap, lengkap: adaGanjil && adaGenap, sudahArsip };
}

function simpanTahunAjaranBaru(tahun, semester){
  const auth = getAuth();
  if(!tahun || !semester) throw new Error('Tahun pelajaran dan semester wajib diisi');
  tahun    = String(tahun).trim();
  semester = String(semester).trim();
  const sh   = sheet('SETTING');
  const data = sh.getDataRange().getValues();
  const hdr  = data[0].map(h => String(h).toLowerCase().trim());
  const tIdx   = hdr.indexOf('tahun');
  const semIdx = hdr.indexOf('semester');
  const lkIdx  = hdr.indexOf('lock_status');
  const email  = auth.email;
  let found = false;
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]).toLowerCase().trim() !== email) continue;
    if(tIdx   > -1) sh.getRange(i+1, tIdx+1).setValue(tahun);
    if(semIdx > -1) sh.getRange(i+1, semIdx+1).setValue(semester);
    if(lkIdx  > -1) sh.getRange(i+1, lkIdx+1).setValue('LOCKED');
    found = true;
    break;
  }
  if(!found) throw new Error('Setting tidak ditemukan untuk akun ini');
  invalidateCache_('SETTING');
  logAudit('SET_TAHUN_BARU', email, tahun + ' ' + semester);
  return { success: true, tahun, semester };
}

function fillArsipJurnalSheet_(sh, rows, semLabel, setting, tahun, tz){
  const nama  = setting.nama_guru        || '-';
  const nip   = setting.nip_guru         || '-';
  const mapel = setting.mata_pelajaran   || '-';
  sh.getRange('A1').setValue('REKAP JURNAL ' + semLabel + ' \u2013 ' + tahun)
    .setFontSize(13).setFontWeight('bold');
  sh.getRange('A2').setValue('Guru: '+nama+' | NIP: '+nip+' | Mapel: '+mapel);
  sh.getRange('A4:H4').setValues([[
    'No','Tanggal','Kelas','Jam ke-','Pertemuan ke-','Materi Pembelajaran','Asesmen','Refleksi'
  ]]).setFontWeight('bold').setBackground('#4338ca').setFontColor('#ffffff');
  if(rows.length > 0){
    // ✅ BATCH: collect all rows first, then single setValues call
    const allRows = rows.map((r, i) => {
      const tgl = r[1] ? Utilities.formatDate(new Date(r[1]), tz, 'dd/MM/yyyy') : '-';
      return [i+1, tgl, r[2]||'-', r[3]||'-', r[4]||'-', r[5]||'-', r[7]||'-', r[15]||'-'];
    });
    sh.getRange(5, 1, allRows.length, 8).setValues(allRows);
    sh.getRange('A4:H'+(4+rows.length)).setBorder(true,true,true,true,true,true);
    sh.setColumnWidths(1,1,35); sh.setColumnWidth(2,80);  sh.setColumnWidth(3,60);
    sh.setColumnWidth(4,55);    sh.setColumnWidth(5,80);  sh.setColumnWidth(6,220);
    sh.setColumnWidth(7,160);   sh.setColumnWidth(8,220);
    sh.getRange('A5:H'+(4+rows.length)).setWrap(true).setVerticalAlignment('top');
  }
}

function fillArsipAbsensiSheet_(sh, rows, semLabel, setting, tahun, tz, absMap){
  const nama = setting.nama_guru || '-';
  sh.getRange('A1').setValue('REKAP ABSENSI ' + semLabel + ' \u2013 ' + tahun)
    .setFontSize(13).setFontWeight('bold');
  sh.getRange('A2').setValue('Guru: '+nama+' | Jumlah Pertemuan: '+rows.length);
  sh.getRange('A4:H4').setValues([[
    'No','Tanggal','Kelas','Pertemuan ke-','Hadir (H)','Sakit (S)','Izin (I)','Alpha (A)'
  ]]).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  if(rows.length > 0){
    // ✅ BATCH: collect all rows first, then single setValues call
    const allRows = rows.map((r, i) => {
      const tgl = r[1] ? Utilities.formatDate(new Date(r[1]), tz, 'dd/MM/yyyy') : '-';
      const jid = String(r[0]);
      const abs = absMap[jid] || [];
      const H   = abs.filter(a => a[3]==='H').length;
      const S   = abs.filter(a => a[3]==='S').length;
      const Izn = abs.filter(a => a[3]==='I').length;
      const Alp = abs.filter(a => a[3]==='A').length;
      return [i+1, tgl, r[2]||'-', r[4]||'-', H, S, Izn, Alp];
    });
    sh.getRange(5, 1, allRows.length, 8).setValues(allRows);
    sh.getRange('A4:H'+(4+rows.length)).setBorder(true,true,true,true,true,true);
    sh.setColumnWidths(1,1,35); sh.setColumnWidth(2,80); sh.setColumnWidth(3,60);
    sh.setColumnWidth(4,80);    sh.setColumnWidth(5,55); sh.setColumnWidth(6,55);
    sh.setColumnWidth(7,55);    sh.setColumnWidth(8,55);
  }
}

/**************** INFO STORAGE DRIVE (BARU) ****************/
function getDriveStorageInfo() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  try {
    const token = ScriptApp.getOAuthToken();
    const res   = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/about?fields=storageQuota',
      { headers: { Authorization: 'Bearer ' + token } }
    );
    const q       = JSON.parse(res.getContentText()).storageQuota;
    const usedMB  = Math.round(parseInt(q.usage  || 0) / 1048576);
    const limitGB = q.limit ? (parseInt(q.limit) / 1073741824).toFixed(1) : null;
    const pct     = limitGB
      ? Math.round((usedMB / 1024) / parseFloat(limitGB) * 100)
      : null;
    return { usedMB, limitGB, pct };
  } catch(e) {
    return { usedMB: null, limitGB: null, pct: null, error: e.message };
  }
}

/**************** NOTIFIKASI GURU dari SuperAdmin ****************/
/**
 * Mengembalikan notifikasi yang relevan untuk guru yang sedang login.
 * Diambil dari baris AUDIT_LOG yang menarget email guru tersebut,
 * lakukan oleh SuperAdmin (CREATE_LICENSE, GENERATE_LICENSE, UPDATE_ROLE,
 * UPDATE_STATUS, DELETE_USER).
 * Hanya notif dalam 30 hari terakhir, max 20 entri.
 */
function getNotifikasiGuru() {
  const auth  = getAuth();
  const email = auth.email;
  const log   = sheet('AUDIT_LOG');
  const rows  = log.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  const relevantActions = ['CREATE_LICENSE','GENERATE_LICENSE','UPDATE_ROLE','UPDATE_STATUS','DELETE_USER'];
  const result = [];

  for(let i = rows.length - 1; i >= 1 && result.length < 20; i--){
    const ts       = rows[i][0];
    const actor    = String(rows[i][1] || '');
    const action   = String(rows[i][2] || '');
    const target   = String(rows[i][3] || '').toLowerCase().trim();
    const detail   = String(rows[i][4] || '');

    if(!ts || new Date(ts) < cutoff) continue;
    if(!relevantActions.includes(action)) continue;
    if(target !== email) continue;

    const label = {
      CREATE_LICENSE  : '🔑 Lisensi dibuat untuk akun Anda',
      GENERATE_LICENSE: '✅ Lisensi diaktifkan/diperpanjang: ' + detail,
      UPDATE_ROLE     : '👤 Role akun diubah menjadi: ' + detail,
      UPDATE_STATUS   : '🔄 Status akun diubah menjadi: ' + detail,
      DELETE_USER     : '⚠️ Akun Anda telah dihapus oleh SuperAdmin'
    }[action] || action + ': ' + detail;

    result.push({
      waktu : Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      pesan : label,
      actor : actor,
      action: action
    });
  }

  return result;
}
// =====================================================================
// CENTRAL REGISTRY HUB
// DEPLOYMENTS   : identitas deployment per sekolah/tenant
// RESOURCE_MAP  : resource penyimpanan per guru/deployment
// APP_RELEASES  : versi rilis aplikasi yang boleh disebar
// UPDATE_LOG    : histori update, approval, dan migrasi
// =====================================================================

var DEPLOYMENTS_SHEET  = 'DEPLOYMENTS';
var RESOURCE_MAP_SHEET = 'RESOURCE_MAP';
var APP_RELEASES_SHEET = 'APP_RELEASES';
var UPDATE_LOG_SHEET   = 'UPDATE_LOG';
var SUMMARY_SYNC_SHEET = 'SUMMARY_SYNC';

var DEPLOYMENTS_HEADERS = [
  'id', 'nama_sekolah', 'email_sa', 'script_id', 'spreadsheet_id', 'owner_email',
  'folder_data_id', 'folder_export_id', 'webapp_url', 'app_version', 'schema_version',
  'maintenance_until', 'allow_update', 'release_channel', 'status', 'last_sync',
  'last_update', 'tanggal', 'catatan'
];

var RESOURCE_MAP_HEADERS = [
  'id', 'deployment_id', 'email_guru', 'resource_type', 'resource_id', 'resource_name',
  'owner_email', 'status', 'created_at', 'updated_at', 'catatan'
];

var APP_RELEASES_HEADERS = [
  'id', 'version', 'schema_version', 'channel', 'status', 'allow_migration',
  'manifest_json', 'created_at', 'created_by', 'catatan'
];

var UPDATE_LOG_HEADERS = [
  'id', 'deployment_id', 'target_email', 'from_version', 'to_version', 'from_schema',
  'to_schema', 'action', 'status', 'approved_by', 'executed_by', 'executed_at', 'catatan'
];

var SUMMARY_SYNC_HEADERS = [
  'id', 'deployment_id', 'email_guru', 'nama_guru', 'sekolah', 'tahun_pelajaran', 'semester',
  'total_jurnal', 'total_kelas', 'total_siswa', 'rata_hadir', 'kelas_favorit',
  'ketuntasan_terbaik', 'license_status', 'app_version', 'schema_version', 'last_sync', 'catatan'
];

function _ensureRegistrySheet_(sheetName, headers) {
  var ss = getCentralSpreadsheet_();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    return sh;
  }

  var lastCol = sh.getLastColumn();
  var current = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) {
    return String(v || '').trim();
  }) : [];

  if (!current.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return sh;
  }

  headers.forEach(function(header) {
    if (current.indexOf(header) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(header);
      current.push(header);
    }
  });

  sh.setFrozenRows(1);
  return sh;
}

function _ensureCentralRegistrySchema_() {
  _ensureRegistrySheet_(DEPLOYMENTS_SHEET, DEPLOYMENTS_HEADERS);
  _ensureRegistrySheet_(RESOURCE_MAP_SHEET, RESOURCE_MAP_HEADERS);
  _ensureRegistrySheet_(APP_RELEASES_SHEET, APP_RELEASES_HEADERS);
  _ensureRegistrySheet_(UPDATE_LOG_SHEET, UPDATE_LOG_HEADERS);
  _ensureRegistrySheet_(SUMMARY_SYNC_SHEET, SUMMARY_SYNC_HEADERS);
}

function _getHeaderIndexMap_(sheetObj) {
  var lastCol = sheetObj.getLastColumn();
  var header = lastCol > 0 ? sheetObj.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var map = {};
  header.forEach(function(h, i) { map[String(h || '').trim()] = i; });
  return map;
}

function _ensureDeploymentsSheet_() {
  _ensureCentralRegistrySchema_();
  return getSpreadsheet_().getSheetByName(DEPLOYMENTS_SHEET);
}

function getDeployments() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh   = _ensureDeploymentsSheet_();
  var data = sh.getDataRange().getValues();
  var idx  = _getHeaderIndexMap_(sh);
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][idx.id]) continue;
    result.push({
      id                : String(data[i][idx.id] || ''),
      nama_sekolah      : String(data[i][idx.nama_sekolah] || ''),
      email_sa          : String(data[i][idx.email_sa] || ''),
      script_id         : String(data[i][idx.script_id] || ''),
      spreadsheet_id    : String(data[i][idx.spreadsheet_id] || ''),
      owner_email       : String(data[i][idx.owner_email] || ''),
      folder_data_id    : String(data[i][idx.folder_data_id] || ''),
      folder_export_id  : String(data[i][idx.folder_export_id] || ''),
      webapp_url        : String(data[i][idx.webapp_url] || ''),
      app_version       : String(data[i][idx.app_version] || ''),
      schema_version    : String(data[i][idx.schema_version] || ''),
      maintenance_until : String(data[i][idx.maintenance_until] || ''),
      allow_update      : String(data[i][idx.allow_update] || 'true'),
      release_channel   : String(data[i][idx.release_channel] || 'stable'),
      tanggal           : String(data[i][idx.tanggal] || ''),
      last_sync         : String(data[i][idx.last_sync] || ''),
      last_update       : String(data[i][idx.last_update] || ''),
      status            : String(data[i][idx.status] || 'aktif'),
      catatan           : String(data[i][idx.catatan] || '')
    });
  }
  return result;
}

function saveDeployment(obj) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  if (!obj || !obj.nama_sekolah) throw new Error('Nama sekolah wajib diisi');
  if (!obj.email_sa) throw new Error('Email SA wajib diisi');

  var sh   = _ensureDeploymentsSheet_();
  var data = sh.getDataRange().getValues();
  var idx  = _getHeaderIndexMap_(sh);
  var tz   = Session.getScriptTimeZone();

  function setDeploymentRow_(rowNumber, payload, createdAt) {
    sh.getRange(rowNumber, idx.nama_sekolah + 1).setValue(payload.nama_sekolah || '');
    sh.getRange(rowNumber, idx.email_sa + 1).setValue(payload.email_sa || '');
    sh.getRange(rowNumber, idx.script_id + 1).setValue(payload.script_id || '');
    sh.getRange(rowNumber, idx.spreadsheet_id + 1).setValue(payload.spreadsheet_id || '');
    sh.getRange(rowNumber, idx.owner_email + 1).setValue(payload.owner_email || payload.email_sa || '');
    sh.getRange(rowNumber, idx.folder_data_id + 1).setValue(payload.folder_data_id || '');
    sh.getRange(rowNumber, idx.folder_export_id + 1).setValue(payload.folder_export_id || '');
    sh.getRange(rowNumber, idx.webapp_url + 1).setValue(payload.webapp_url || '');
    sh.getRange(rowNumber, idx.app_version + 1).setValue(payload.app_version || '');
    sh.getRange(rowNumber, idx.schema_version + 1).setValue(payload.schema_version || '');
    sh.getRange(rowNumber, idx.maintenance_until + 1).setValue(payload.maintenance_until || '');
    sh.getRange(rowNumber, idx.allow_update + 1).setValue(String(payload.allow_update !== false));
    sh.getRange(rowNumber, idx.release_channel + 1).setValue(payload.release_channel || 'stable');
    sh.getRange(rowNumber, idx.status + 1).setValue(payload.status || 'aktif');
    sh.getRange(rowNumber, idx.last_sync + 1).setValue(payload.last_sync || '');
    sh.getRange(rowNumber, idx.last_update + 1).setValue(payload.last_update || '');
    sh.getRange(rowNumber, idx.tanggal + 1).setValue(createdAt || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));
    sh.getRange(rowNumber, idx.catatan + 1).setValue(payload.catatan || '');
  }

  if (obj.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idx.id] || '') !== String(obj.id)) continue;
      setDeploymentRow_(i + 1, obj, String(data[i][idx.tanggal] || '') || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));
      logAudit('UPDATE_DEPLOYMENT', getLoginEmail(), obj.nama_sekolah);
      return { success: true, action: 'updated' };
    }
  }

  var id      = 'DEP_' + new Date().getTime();
  var tanggal = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  sh.appendRow(new Array(DEPLOYMENTS_HEADERS.length).fill(''));
  var newRow = sh.getLastRow();
  sh.getRange(newRow, idx.id + 1).setValue(id);
  setDeploymentRow_(newRow, obj, tanggal);
  logAudit('ADD_DEPLOYMENT', getLoginEmail(), obj.nama_sekolah);
  return { success: true, action: 'created', id: id };
}

function deleteDeployment(id) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh   = _ensureDeploymentsSheet_();
  var data = sh.getDataRange().getValues();
  var idx  = _getHeaderIndexMap_(sh);

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idx.id] || '') !== String(id)) continue;
    sh.deleteRow(i + 1);
    logAudit('DELETE_DEPLOYMENT', getLoginEmail(), 'ID: ' + id);
    return { success: true };
  }
  return { success: false };
}

function getCentralRegistryOverview() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  _ensureCentralRegistrySchema_();

  var deployments = getDeployments();
  var resourceSh  = getCentralSpreadsheet_().getSheetByName(RESOURCE_MAP_SHEET);
  var releaseSh   = getCentralSpreadsheet_().getSheetByName(APP_RELEASES_SHEET);
  var updateSh    = getCentralSpreadsheet_().getSheetByName(UPDATE_LOG_SHEET);

  var resources = resourceSh ? Math.max(resourceSh.getLastRow() - 1, 0) : 0;
  var releases  = releaseSh ? Math.max(releaseSh.getLastRow() - 1, 0) : 0;
  var updates   = updateSh ? Math.max(updateSh.getLastRow() - 1, 0) : 0;

  var latestRelease = null;
  if (releaseSh && releaseSh.getLastRow() > 1) {
    var releaseRows = releaseSh.getDataRange().getValues();
    var ridx = _getHeaderIndexMap_(releaseSh);
    var last = releaseRows[releaseRows.length - 1];
    latestRelease = {
      version: String(last[ridx.version] || ''),
      schema_version: String(last[ridx.schema_version] || ''),
      channel: String(last[ridx.channel] || 'stable'),
      status: String(last[ridx.status] || ''),
      created_at: String(last[ridx.created_at] || '')
    };
  }

  return {
    counts: {
      deployments: deployments.length,
      allow_update: deployments.filter(function(d) { return String(d.allow_update).toLowerCase() === 'true'; }).length,
      resources: resources,
      releases: releases,
      updates: updates,
      summary_sync: Math.max(((getCentralSpreadsheet_().getSheetByName(SUMMARY_SYNC_SHEET) || {}).getLastRow || function(){ return 1; })() - 1, 0)
    },
    latestRelease: latestRelease,
    storageMode: (typeof getStorageMode_ === 'function' ? getStorageMode_() : 'central'),
    sheets: {
      deployments: DEPLOYMENTS_HEADERS,
      resource_map: RESOURCE_MAP_HEADERS,
      app_releases: APP_RELEASES_HEADERS,
      update_log: UPDATE_LOG_HEADERS,
      summary_sync: SUMMARY_SYNC_HEADERS
    }
  };
}

function _ensureSummarySyncSheet_() {
  _ensureCentralRegistrySchema_();
  return getCentralSpreadsheet_().getSheetByName(SUMMARY_SYNC_SHEET);
}

function buildGuruSummarySync_(email) {
  if (!email) throw new Error('Email guru wajib diisi');
  email = String(email).toLowerCase().trim();

  var setting = getSetting();
  var dash = getDashboardAllData();
  var license = getSchoolLicenseInfo() || {};
  var deployment = _getDeploymentEntryForUser_(email) || {};
  var favClass = '-';
  var mastery = '-';

  try {
    var recap = getRekapSekolah();
    favClass = (recap.kelasFavorit && recap.kelasFavorit[0] && recap.kelasFavorit[0].kelas) || '-';
    mastery = (recap.kelasKetuntasan && recap.kelasKetuntasan[0])
      ? (recap.kelasKetuntasan[0].kelas + ' (' + recap.kelasKetuntasan[0].persenTuntas + '%)')
      : '-';
  } catch (e) {}

  return {
    deployment_id: deployment.id || '',
    email_guru: email,
    nama_guru: setting.nama_guru || '-',
    sekolah: setting.sekolah || '-',
    tahun_pelajaran: setting.tahun_pelajaran || '-',
    semester: setting.semester || '-',
    total_jurnal: Number(dash.totalJurnal || 0),
    total_kelas: Number(dash.totalKelas || 0),
    total_siswa: Number(dash.totalSiswa || 0),
    rata_hadir: Number(dash.rata2 || 0),
    kelas_favorit: favClass,
    ketuntasan_terbaik: mastery,
    license_status: license.isLifetime ? 'lifetime' : (license.isTrial ? 'trial' : (license.isActive ? 'active' : 'inactive')),
    app_version: deployment.app_version || '',
    schema_version: deployment.schema_version || '',
    catatan: ''
  };
}

function syncGuruSummaryToCentral() {
  const auth = getAuth();
  if (!auth.email || (auth.role !== 'admin' && auth.role !== 'superadmin' && auth.role !== 'kepsek')) {
    throw new Error('AKSES_DITOLAK');
  }

  const sh = _ensureSummarySyncSheet_();
  const rows = sh.getDataRange().getValues();
  const idx = _getHeaderIndexMap_(sh);
  const summary = buildGuruSummarySync_(auth.email);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idx.email_guru] || '').toLowerCase().trim() !== auth.email) continue;
    sh.getRange(i + 1, idx.deployment_id + 1).setValue(summary.deployment_id);
    sh.getRange(i + 1, idx.nama_guru + 1).setValue(summary.nama_guru);
    sh.getRange(i + 1, idx.sekolah + 1).setValue(summary.sekolah);
    sh.getRange(i + 1, idx.tahun_pelajaran + 1).setValue(summary.tahun_pelajaran);
    sh.getRange(i + 1, idx.semester + 1).setValue(summary.semester);
    sh.getRange(i + 1, idx.total_jurnal + 1).setValue(summary.total_jurnal);
    sh.getRange(i + 1, idx.total_kelas + 1).setValue(summary.total_kelas);
    sh.getRange(i + 1, idx.total_siswa + 1).setValue(summary.total_siswa);
    sh.getRange(i + 1, idx.rata_hadir + 1).setValue(summary.rata_hadir);
    sh.getRange(i + 1, idx.kelas_favorit + 1).setValue(summary.kelas_favorit);
    sh.getRange(i + 1, idx.ketuntasan_terbaik + 1).setValue(summary.ketuntasan_terbaik);
    sh.getRange(i + 1, idx.license_status + 1).setValue(summary.license_status);
    sh.getRange(i + 1, idx.app_version + 1).setValue(summary.app_version);
    sh.getRange(i + 1, idx.schema_version + 1).setValue(summary.schema_version);
    sh.getRange(i + 1, idx.last_sync + 1).setValue(now);
    sh.getRange(i + 1, idx.catatan + 1).setValue(summary.catatan || '');
    logAudit('SYNC_GURU_SUMMARY', auth.email, summary.sekolah + ' | ' + now);
    return { success: true, action: 'updated', last_sync: now };
  }

  sh.appendRow([
    'SUM_' + new Date().getTime(),
    summary.deployment_id,
    summary.email_guru,
    summary.nama_guru,
    summary.sekolah,
    summary.tahun_pelajaran,
    summary.semester,
    summary.total_jurnal,
    summary.total_kelas,
    summary.total_siswa,
    summary.rata_hadir,
    summary.kelas_favorit,
    summary.ketuntasan_terbaik,
    summary.license_status,
    summary.app_version,
    summary.schema_version,
    now,
    summary.catatan || ''
  ]);
  logAudit('SYNC_GURU_SUMMARY', auth.email, summary.sekolah + ' | ' + now);
  return { success: true, action: 'created', last_sync: now };
}

function getCentralSummarySyncRows() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const sh = _ensureSummarySyncSheet_();
  const rows = sh.getDataRange().getValues();
  const idx = _getHeaderIndexMap_(sh);
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][idx.id] || '').trim();
    if (!id) continue;
    result.push({
      id: id,
      deployment_id: String(rows[i][idx.deployment_id] || ''),
      email_guru: String(rows[i][idx.email_guru] || ''),
      nama_guru: String(rows[i][idx.nama_guru] || ''),
      sekolah: String(rows[i][idx.sekolah] || ''),
      tahun_pelajaran: String(rows[i][idx.tahun_pelajaran] || ''),
      semester: String(rows[i][idx.semester] || ''),
      total_jurnal: Number(rows[i][idx.total_jurnal] || 0),
      total_kelas: Number(rows[i][idx.total_kelas] || 0),
      total_siswa: Number(rows[i][idx.total_siswa] || 0),
      rata_hadir: Number(rows[i][idx.rata_hadir] || 0),
      kelas_favorit: String(rows[i][idx.kelas_favorit] || '-'),
      ketuntasan_terbaik: String(rows[i][idx.ketuntasan_terbaik] || '-'),
      license_status: String(rows[i][idx.license_status] || '-'),
      app_version: String(rows[i][idx.app_version] || ''),
      schema_version: String(rows[i][idx.schema_version] || ''),
      last_sync: String(rows[i][idx.last_sync] || ''),
      catatan: String(rows[i][idx.catatan] || '')
    });
  }
  result.sort(function(a, b) { return String(b.last_sync || '').localeCompare(String(a.last_sync || '')); });
  return result;
}

function getStorageModeConfig() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  return {
    mode: (typeof getStorageMode_ === 'function' ? getStorageMode_() : 'central')
  };
}

function setStorageModeConfig(mode) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var target = String(mode || '').toLowerCase().trim();
  if (target !== 'central' && target !== 'per_guru') throw new Error('Mode storage tidak valid');
  PropertiesService.getScriptProperties().setProperty('APP_STORAGE_MODE', target);
  logAudit('SET_STORAGE_MODE', getLoginEmail(), target);
  return { success: true, mode: target };
}

function _ensureResourceMapSheet_() {
  _ensureCentralRegistrySchema_();
  return getSpreadsheet_({ forceCentral: true }).getSheetByName(RESOURCE_MAP_SHEET);
}

function getResourceMapEntries(deploymentId) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh = _ensureResourceMapSheet_();
  var rows = sh.getDataRange().getValues();
  var idx = _getHeaderIndexMap_(sh);
  var targetDeployment = String(deploymentId || '').trim();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][idx.id] || '').trim();
    if (!id) continue;
    var depId = String(rows[i][idx.deployment_id] || '').trim();
    if (targetDeployment && depId !== targetDeployment) continue;
    result.push({
      id: id,
      deployment_id: depId,
      email_guru: String(rows[i][idx.email_guru] || ''),
      resource_type: String(rows[i][idx.resource_type] || ''),
      resource_id: String(rows[i][idx.resource_id] || ''),
      resource_name: String(rows[i][idx.resource_name] || ''),
      owner_email: String(rows[i][idx.owner_email] || ''),
      status: String(rows[i][idx.status] || 'active'),
      created_at: String(rows[i][idx.created_at] || ''),
      updated_at: String(rows[i][idx.updated_at] || ''),
      catatan: String(rows[i][idx.catatan] || '')
    });
  }
  return result;
}

function saveResourceMapEntry(obj) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  obj = obj || {};
  if (!obj.deployment_id) throw new Error('Deployment wajib dipilih');
  if (!obj.email_guru) throw new Error('Email guru wajib diisi');
  if (!obj.resource_type) throw new Error('Tipe resource wajib diisi');
  if (!obj.resource_id) throw new Error('Resource ID wajib diisi');

  var sh = _ensureResourceMapSheet_();
  var rows = sh.getDataRange().getValues();
  var idx = _getHeaderIndexMap_(sh);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var normalizedType = String(obj.resource_type).toLowerCase().trim();

  for (var i = 1; i < rows.length; i++) {
    var rowId = String(rows[i][idx.id] || '').trim();
    if (obj.id && rowId === String(obj.id).trim()) {
      sh.getRange(i + 1, idx.deployment_id + 1).setValue(obj.deployment_id);
      sh.getRange(i + 1, idx.email_guru + 1).setValue(String(obj.email_guru).toLowerCase().trim());
      sh.getRange(i + 1, idx.resource_type + 1).setValue(normalizedType);
      sh.getRange(i + 1, idx.resource_id + 1).setValue(obj.resource_id);
      sh.getRange(i + 1, idx.resource_name + 1).setValue(obj.resource_name || '');
      sh.getRange(i + 1, idx.owner_email + 1).setValue((obj.owner_email || obj.email_guru || '').toLowerCase().trim());
      sh.getRange(i + 1, idx.status + 1).setValue(obj.status || 'active');
      sh.getRange(i + 1, idx.updated_at + 1).setValue(now);
      sh.getRange(i + 1, idx.catatan + 1).setValue(obj.catatan || '');
      logAudit('UPDATE_RESOURCE_MAP', getLoginEmail(), obj.deployment_id + ' | ' + normalizedType + ' | ' + obj.email_guru);
      return { success: true, action: 'updated', id: rowId };
    }
  }

  var id = 'RES_' + new Date().getTime();
  sh.appendRow([
    id,
    obj.deployment_id,
    String(obj.email_guru).toLowerCase().trim(),
    normalizedType,
    obj.resource_id,
    obj.resource_name || '',
    (obj.owner_email || obj.email_guru || '').toLowerCase().trim(),
    obj.status || 'active',
    now,
    now,
    obj.catatan || ''
  ]);
  logAudit('ADD_RESOURCE_MAP', getLoginEmail(), obj.deployment_id + ' | ' + normalizedType + ' | ' + obj.email_guru);
  return { success: true, action: 'created', id: id };
}

function deleteResourceMapEntry(id) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh = _ensureResourceMapSheet_();
  var rows = sh.getDataRange().getValues();
  var idx = _getHeaderIndexMap_(sh);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idx.id] || '').trim() !== String(id || '').trim()) continue;
    sh.deleteRow(i + 1);
    logAudit('DELETE_RESOURCE_MAP', getLoginEmail(), String(id || ''));
    return { success: true };
  }
  return { success: false };
}

/* =====================================================
   AUDIT LOG — Kosongkan
   ===================================================== */

/**
 * clearAuditLog() — Hapus semua baris AUDIT_LOG kecuali header.
 * Hanya SuperAdmin.
 */
function clearAuditLog() {
  if (!isSuperAdmin()) throw new Error('Akses ditolak');
  const sh = sheet('AUDIT_LOG');
  if (!sh) return { success: true, deleted: 0 };
  const total = sh.getLastRow();
  if (total <= 1) return { success: true, deleted: 0 };
  sh.deleteRows(2, total - 1);
  logAudit('CLEAR_AUDIT_LOG', getLoginEmail(), (total - 1) + ' baris dihapus');
  return { success: true, deleted: total - 1 };
}

/* =====================================================
   ARSIP DRIVE — Info ukuran + Delete + Forward Email
   ===================================================== */

/**
 * getArsipDriveUsage() — Ukuran folder JURNAL_ARSIP per guru (bytes & MB).
 * SuperAdmin only.
 */
function getArsipDriveUsage() {
  if (!isSuperAdmin()) throw new Error('Akses ditolak');
  const result = [];
  try {
    if (typeof getStorageMode_ === 'function' && getStorageMode_() === 'per_guru') {
      const entries = getResourceMapEntries('').filter(function(r) {
        return String(r.resource_type || '').toLowerCase() === 'arsip_folder' &&
          String(r.status || '').toLowerCase() === 'active' &&
          String(r.resource_id || '').trim();
      });
      let totalBytesPerGuru = 0;
      entries.forEach(function(entry) {
        try {
          const folder = DriveApp.getFolderById(entry.resource_id);
          let guruBytes = 0;
          const files = [];

          const directFiles = folder.getFiles();
          while (directFiles.hasNext()) {
            const f = directFiles.next();
            const sz = f.getSize();
            guruBytes += sz;
            files.push({ id: f.getId(), nama: f.getName(), tahun: '-', size_mb: (sz / 1048576).toFixed(2), url: f.getDownloadUrl() });
          }

          const subIt = folder.getFolders();
          while (subIt.hasNext()) {
            const tahunFolder = subIt.next();
            const tahunName = tahunFolder.getName();
            const tahunFiles = tahunFolder.getFiles();
            while (tahunFiles.hasNext()) {
              const f = tahunFiles.next();
              const sz = f.getSize();
              guruBytes += sz;
              files.push({ id: f.getId(), nama: f.getName(), tahun: tahunName, size_mb: (sz / 1048576).toFixed(2), url: f.getDownloadUrl() });
            }
          }

          totalBytesPerGuru += guruBytes;
          result.push({ guru: entry.email_guru || entry.owner_email || folder.getName(), size_mb: (guruBytes / 1048576).toFixed(2), files });
        } catch (e) {}
      });
      return { total_mb: (totalBytesPerGuru / 1048576).toFixed(2), entries: result };
    }

    const root = _getArsipRootFolder_();
    let totalBytes = 0;

    const guruIt = root.getFolders();
    while (guruIt.hasNext()) {
      const guruFolder = guruIt.next();
      const guruName = guruFolder.getName();
      let guruBytes = 0;
      const files = [];

      // File langsung di folder guru
      const directFiles = guruFolder.getFiles();
      while (directFiles.hasNext()) {
        const f = directFiles.next();
        const sz = f.getSize();
        guruBytes += sz;
        files.push({ id: f.getId(), nama: f.getName(), tahun: '-', size_mb: (sz / 1048576).toFixed(2), url: f.getDownloadUrl() });
      }

      // File di subfolder tahun
      const subIt = guruFolder.getFolders();
      while (subIt.hasNext()) {
        const tahunFolder = subIt.next();
        const tahunName = tahunFolder.getName();
        const tahunFiles = tahunFolder.getFiles();
        while (tahunFiles.hasNext()) {
          const f = tahunFiles.next();
          const sz = f.getSize();
          guruBytes += sz;
          files.push({ id: f.getId(), nama: f.getName(), tahun: tahunName, size_mb: (sz / 1048576).toFixed(2), url: f.getDownloadUrl() });
        }
      }

      totalBytes += guruBytes;
      result.push({ guru: guruName, size_mb: (guruBytes / 1048576).toFixed(2), files });
    }
    return { total_mb: (totalBytes / 1048576).toFixed(2), entries: result };
  } catch (e) {
    return { total_mb: 0, entries: [], error: e.message };
  }
}

/**
 * deleteArsipFromDrive(fileId) — Pindahkan file arsip ke Trash Drive.
 * (File bisa dipulihkan dalam 30 hari dari Google Drive Trash)
 * SuperAdmin only.
 */
function deleteArsipFromDrive(fileId) {
  if (!isSuperAdmin()) throw new Error('Akses ditolak');
  if (!fileId) throw new Error('fileId wajib diisi');
  try {
    const file = DriveApp.getFileById(fileId);
    const nama = file.getName();
    file.setTrashed(true);
    logAudit('TRASH_ARSIP_DRIVE', getLoginEmail(), nama);
    return { success: true, nama };
  } catch (e) {
    throw new Error('Gagal menghapus file: ' + e.message);
  }
}

/**
 * forwardArsipToEmail(fileId, targetEmail) — Kirim link download arsip ke email lain via Gmail.
 * SuperAdmin only.
 */
function forwardArsipToEmail(fileId, targetEmail) {
  if (!isSuperAdmin()) throw new Error('Akses ditolak');
  if (!fileId) throw new Error('fileId wajib diisi');
  if (!targetEmail || !targetEmail.trim()) throw new Error('Email tujuan wajib diisi');
  targetEmail = targetEmail.trim().toLowerCase();
  // Validasi format email sederhana
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(targetEmail)) throw new Error('Format email tidak valid');

  try {
    const file = DriveApp.getFileById(fileId);
    const nama = file.getName();
    const url  = file.getDownloadUrl();

    const subject = '[Backup Jurnal] ' + nama;
    const body = 'Halo,\n\nBerikut link download arsip jurnal:\n\nNama File: ' + nama +
                 '\nLink: ' + url +
                 '\n\nLink berlaku selama akun yang membagikan masih aktif.\n\nSalam,\nSistem Akademik Guru';
    GmailApp.sendEmail(targetEmail, subject, body);
    logAudit('FORWARD_ARSIP', getLoginEmail(), nama + ' → ' + targetEmail);
    return { success: true, nama, targetEmail };
  } catch (e) {
    throw new Error('Gagal mengirim email: ' + e.message);
  }
}

/* =====================================================
   KONFIGURASI FOLDER ARSIP
   ===================================================== */

var PROP_ARSIP_FOLDER = 'ARSIP_FOLDER_ID';
var PROP_ARSIP_WARN_MB = 'ARSIP_WARN_MB';

/**
 * getArsipFolderConfig() — Info folder arsip yang dikonfigurasi + threshold warning.
 */
function getArsipFolderConfig() {
  if (!isSuperAdmin()) throw new Error('Akses ditolak');
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty(PROP_ARSIP_FOLDER) || '';
  const warnMb   = Number(props.getProperty(PROP_ARSIP_WARN_MB) || 500);
  let folderName = '';
  let folderUrl  = '';
  if (folderId) {
    try {
      const f = DriveApp.getFolderById(folderId);
      folderName = f.getName();
      folderUrl  = 'https://drive.google.com/drive/folders/' + folderId;
    } catch(e) { folderName = '⚠️ Folder tidak ditemukan'; }
  }
  return { folderId, folderName, folderUrl, warnMb };
}

/**
 * setArsipFolderConfig(input, warnMb) — Simpan folder ID arsip dan threshold warning.
 * input: URL folder Google Drive atau folder ID langsung
 */
function setArsipFolderConfig(input, warnMb) {
  if (!isSuperAdmin()) throw new Error('Akses ditolak');
  const props = PropertiesService.getScriptProperties();

  if (warnMb !== undefined && warnMb !== null) {
    const mb = Number(warnMb);
    if (!isNaN(mb) && mb > 0) props.setProperty(PROP_ARSIP_WARN_MB, String(mb));
  }

  if (!input || !String(input).trim()) {
    // Hapus config — kembali ke default JURNAL_ARSIP
    props.deleteProperty(PROP_ARSIP_FOLDER);
    logAudit('ARSIP_FOLDER_RESET', getLoginEmail(), 'Kembali ke JURNAL_ARSIP default');
    return { success: true, folderId: '', folderName: 'JURNAL_ARSIP (default)' };
  }

  // Ekstrak folder ID dari URL atau pakai langsung sebagai ID
  const raw = String(input).trim();
  let folderId = raw;
  const m = raw.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) folderId = m[1];

  try {
    const f = DriveApp.getFolderById(folderId);
    const folderName = f.getName();
    props.setProperty(PROP_ARSIP_FOLDER, folderId);
    logAudit('ARSIP_FOLDER_SET', getLoginEmail(), folderName + ' (' + folderId + ')');
    return { success: true, folderId, folderName };
  } catch(e) {
    throw new Error('Folder tidak ditemukan atau tidak bisa diakses: ' + folderId);
  }
}

/**
 * _getArsipRootFolder_() — Helper: ambil folder root arsip (dari config atau default).
 * Dipakai oleh arsipTahunPelajaran, getStatusArsipTahunan, getDaftarArsip.
 */
function _getArsipRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const configId = props.getProperty(PROP_ARSIP_FOLDER);
  if (configId) {
    try { return DriveApp.getFolderById(configId); } catch(e) { /* fallback */ }
  }
  // Default: folder bernama "JURNAL_ARSIP"
  const it = DriveApp.getFoldersByName('JURNAL_ARSIP');
  return it.hasNext() ? it.next() : DriveApp.createFolder('JURNAL_ARSIP');
}