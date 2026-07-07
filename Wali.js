/**
 * Wali.js — Jurnal Guru Wali & Siswa Binaan
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

function ensureSiswaBinaanSheet_(){
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('SISWA_BINAAN');
  if(!sh){
    sh = ss.insertSheet('SISWA_BINAAN');
    sh.appendRow([
      'id','nama_siswa','nis','kelas',
      'guru_wali','tahun_masuk','status'
    ]);
  }
  return sh;
}

function ensureJurnalGuruWaliSheet_(){
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('JURNAL_GURU_WALI');
  if(!sh){
    sh = ss.insertSheet('JURNAL_GURU_WALI');
    sh.appendRow([
      'id','tanggal','hari','waktu',
      'fokus_pendampingan','topik_pendampingan',
      'catatan','tindak_lanjut','dokumentasi',
      'guru_wali','nip','tahun_pelajaran'
    ]);
  }
  return sh;
}

function getInfoGuruWali(){
  const auth    = getAuth();
  const setting = getSetting();
  const sh      = ensureSiswaBinaanSheet_();
  const data    = sh.getDataRange().getValues();
  const email   = auth.email;

  let jumlah = 0;
  for(let i = 1; i < data.length; i++){
    const guruWali = String(data[i][4] || '').toLowerCase().trim();
    const status   = String(data[i][6] || '').toLowerCase().trim();
    if(guruWali === email && status !== 'nonaktif') jumlah++;
  }

  return {
    nama_guru   : setting.nama_guru || '-',
    nip         : setting.nip_guru  || '-',
    jumlah_siswa: jumlah
  };
}

function getSiswaBinaan(page, pageSize){
  const auth  = getAuth();
  const sh    = ensureSiswaBinaanSheet_();
  const data  = sh.getDataRange().getValues();
  const email = auth.email;

  const PAGE_SIZE = Number(pageSize) || 20;
  const PAGE      = Math.max(1, Number(page) || 1);

  const all = [];
  for(let i = 1; i < data.length; i++){
    const guruWali = String(data[i][4] || '').toLowerCase().trim();
    const status   = String(data[i][6] || '').toLowerCase().trim();
    if(guruWali !== email) continue;
    if(status === 'nonaktif') continue;
    all.push({
      nama_siswa: String(data[i][1] || '-'),
      nis       : String(data[i][2] || '-'),
      kelas     : String(data[i][3] || '-')
    });
  }

  const total  = all.length;
  const start  = (PAGE - 1) * PAGE_SIZE;
  const sliced = all.slice(start, start + PAGE_SIZE);

  return {
    data    : sliced.map((s, i) => ({ no: start + i + 1, ...s })),
    total,
    page    : PAGE,
    pageSize: PAGE_SIZE,
    totalPages: Math.ceil(total / PAGE_SIZE)
  };
}

function simpanJurnalGuruWali(data){

  assertLicenseActive();

  const auth    = getAuth();
  const setting = getSetting();

  if(!data.tanggal)            throw new Error('Tanggal wajib diisi');
  if(!data.hari)               throw new Error('Hari wajib diisi');
  if(!data.waktu_mulai)        throw new Error('Waktu mulai wajib diisi');
  if(!data.waktu_selesai)      throw new Error('Waktu selesai wajib diisi');
  if(!data.fokus_pendampingan) throw new Error('Fokus pendampingan wajib dipilih');
  if(!data.topik_pendampingan) throw new Error('Topik pendampingan wajib diisi');

  const sh     = ensureJurnalGuruWaliSheet_();
  const id     = Date.now().toString();

  const wMulai   = String(data.waktu_mulai).replace(':','.');
  const wSelesai = String(data.waktu_selesai).replace(':','.');
  const waktu    = wMulai + ' - ' + wSelesai;

  // ── Upload photos to Drive if provided ──
  let dokumentasi = data.dokumentasi || '-';
  if(data.fotos && Array.isArray(data.fotos) && data.fotos.length > 0){
    try {
      const urls = _uploadFotoWali_(auth.email, id, data.fotos);
      if(urls.length > 0) dokumentasi = urls.join(',');
    } catch(e){
      Logger.log('Foto upload warning: ' + e.message);
    }
  }

  sh.appendRow([
    id,
    new Date(data.tanggal + 'T00:00:00'),
    String(data.hari).toUpperCase(),
    waktu,
    data.fokus_pendampingan,
    data.topik_pendampingan,
    data.catatan     || '-',
    data.tindak_lanjut || '-',
    dokumentasi,
    auth.email,
    setting.nip_guru || '-',
    setting.tahun_pelajaran || ''
  ]);

  logAudit('SIMPAN_JURNAL_WALI', auth.email, data.topik_pendampingan);

  return { status: true, id };
}

/**
 * Upload 1–2 compressed photos to Drive, return public view URLs.
 * @param {string} email
 * @param {string} id  – jurnal ID (used as filename prefix)
 * @param {Array}  fotos – [{data:'data:image/jpeg;base64,...', type, name}]
 */
function _uploadFotoWali_(email, id, fotos){
  const setting = getSetting();
  const tahun   = setting.tahun_pelajaran || 'Tanpa_Tahun';

  const root = DriveApp.getFoldersByName('WALI_DOKUMENTASI').hasNext()
    ? DriveApp.getFoldersByName('WALI_DOKUMENTASI').next()
    : DriveApp.createFolder('WALI_DOKUMENTASI');

  const safeEmail = email.replace(/[@.]/g,'_');
  const guruFolder = root.getFoldersByName(safeEmail).hasNext()
    ? root.getFoldersByName(safeEmail).next()
    : root.createFolder(safeEmail);
  const tahunFolder = guruFolder.getFoldersByName(tahun).hasNext()
    ? guruFolder.getFoldersByName(tahun).next()
    : guruFolder.createFolder(tahun);

  const urls = [];
  fotos.slice(0, 2).forEach(function(f, idx){
    const base64str = String(f.data || '').split(',').pop();
    if(!base64str) return;
    const bytes = Utilities.base64Decode(base64str);
    const blob  = Utilities.newBlob(
      bytes,
      f.type || 'image/jpeg',
      id + '_foto' + (idx + 1) + '.jpg'
    );
    const file = tahunFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls.push('https://drive.google.com/uc?id=' + file.getId());
  });
  return urls;
}


function getRiwayatJurnalGuruWali(){

  const auth        = getAuth();
  const setting     = getSetting();
  const activeTahun = setting.tahun_pelajaran || '';
  const sh    = ensureJurnalGuruWaliSheet_();
  const data  = sh.getDataRange().getValues();
  const email = auth.email;
  const tz    = Session.getScriptTimeZone();

  let result = [];

  for(let i = 1; i < data.length; i++){

    const guruWali = String(data[i][9] || '').toLowerCase().trim();
    if(guruWali !== email) continue;

    const rowTahun = String(data[i][11]||'').trim();
    if(activeTahun && rowTahun && rowTahun !== activeTahun) continue;

    const tgl       = data[i][1] ? new Date(data[i][1]) : null;
    const isoTanggal= tgl ? Utilities.formatDate(tgl, tz, 'yyyy-MM-dd') : '';
    const hariStr   = String(data[i][2] || '');
    const tglFmt    = tgl ? Utilities.formatDate(tgl, tz, 'dd MMMM yyyy') : '-';

    result.push({
      id                : String(data[i][0]),
      tanggal           : hariStr ? hariStr + ', ' + tglFmt : tglFmt,
      iso_tanggal       : isoTanggal,
      waktu             : String(data[i][3]  || '-'),
      fokus_pendampingan: String(data[i][4]  || '-'),
      topik_pendampingan: String(data[i][5]  || '-'),
      catatan           : String(data[i][6]  || '-'),
      tindak_lanjut     : String(data[i][7]  || '-'),
      dokumentasi       : String(data[i][8]  || '-')
    });
  }

  return result.reverse();
}

function hapusJurnalGuruWali(id){

  assertLicenseActive();

  const auth = getAuth();
  const sh   = ensureJurnalGuruWaliSheet_();
  const rows = sh.getDataRange().getValues();

  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) !== String(id)) continue;

    const owner = String(rows[i][9] || '').toLowerCase().trim();
    if(auth.role !== 'superadmin' && owner !== auth.email){
      throw new Error('AKSES_DITOLAK');
    }

    sh.deleteRow(i + 1);
    logAudit('HAPUS_JURNAL_WALI', auth.email, id);
    return true;
  }

  throw new Error('Jurnal tidak ditemukan');
}

function getStatusArsipJurnalWali(){
  const auth        = getAuth();
  const setting     = getSetting();
  const activeTahun = setting.tahun_pelajaran || '';
  const email       = auth.email;
  const safeEmail   = email.replace(/[@.]/g,'_');

  const sh   = ensureJurnalGuruWaliSheet_();
  const data = sh.getDataRange().getValues();

  const tahunSet = new Set();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][9]||'').toLowerCase().trim() !== email) continue;
    const t = String(data[i][11]||'').trim();
    if(t) tahunSet.add(t);
  }

  let sudahArsip = false;
  try{
    const rootIt = DriveApp.getFoldersByName('JURNAL_ARSIP');
    if(activeTahun && rootIt.hasNext()){
      const guruIt = rootIt.next().getFoldersByName(safeEmail);
      if(guruIt.hasNext()){
        const safeTahun = activeTahun.replace(/[\/\\:*?\[\]]/g,'-');
        const f = guruIt.next().getFilesByName('JURNAL_WALI_'+safeTahun+'.xlsx');
        if(f.hasNext()) sudahArsip = true;
      }
    }
  }catch(e){ console.error('[JGD] cek arsip wali:', e.message||e); }

  const tahunList = Array.from(tahunSet).sort();
  return {
    activeTahun,
    totalTahun: tahunSet.size,
    tahunList,
    sudahArsip,
    perluArsip: tahunSet.size >= 3
  };
}

function arsipJurnalGuruWali(){
  assertLicenseActive();
  const auth    = getAuth();
  const setting = getSetting();
  const tahun   = setting.tahun_pelajaran;
  const email   = auth.email;
  const tz      = Session.getScriptTimeZone();

  if(!tahun) throw new Error('Tahun pelajaran belum diset');

  const sh   = ensureJurnalGuruWaliSheet_();
  const data = sh.getDataRange().getValues();

  const toArchive = [];
  for(let i = 1; i < data.length; i++){
    if(String(data[i][9]||'').toLowerCase().trim() !== email) continue;
    const rowTahun = String(data[i][11]||'').trim() || tahun;
    if(rowTahun !== tahun) continue;
    toArchive.push({ rowIdx: i, row: data[i] });
  }

  if(toArchive.length === 0)
    throw new Error('Tidak ada jurnal guru wali untuk tahun ' + tahun);

  const safeTahun = tahun.replace(/[\/\\:*?\[\]]/g,'-');
  const tempSS    = SpreadsheetApp.create('TEMP_ARSIP_WALI_' + safeTahun);
  const tempId    = tempSS.getId();

  try{
    const sh2 = tempSS.getActiveSheet();
    sh2.setName('Jurnal_Guru_Wali');
    sh2.getRange('A1').setValue('REKAP JURNAL GURU WALI \u2013 ' + tahun)
      .setFontSize(13).setFontWeight('bold');
    sh2.getRange('A2').setValue(
      'Guru: '+(setting.nama_guru||email)+' | NIP: '+(setting.nip_guru||'-'));
    sh2.getRange('A4:J4').setValues([[
      'No','Tanggal','Hari','Waktu','Fokus Pendampingan',
      'Topik Pendampingan','Catatan','Tindak Lanjut','Dokumentasi','NIP'
    ]]).setFontWeight('bold').setBackground('#7c3aed').setFontColor('#fff');
    toArchive.forEach((item, i) => {
      const r   = item.row;
      const tgl = r[1] ? Utilities.formatDate(new Date(r[1]), tz, 'dd/MM/yyyy') : '-';
      sh2.getRange(5+i, 1, 1, 10).setValues([[
        i+1, tgl, r[2]||'-', r[3]||'-', r[4]||'-',
        r[5]||'-', r[6]||'-', r[7]||'-', r[8]||'-', r[10]||'-'
      ]]);
    });
    if(toArchive.length > 0){
      sh2.getRange('A4:J'+(4+toArchive.length)).setBorder(true,true,true,true,true,true);
      sh2.getRange('A5:J'+(4+toArchive.length)).setWrap(true).setVerticalAlignment('top');
    }
    SpreadsheetApp.flush();

    const xlsBlob = DriveApp.getFileById(tempId)
      .getBlob()
      .setName('JURNAL_WALI_'+safeTahun+'.xlsx');

    const rootArsip = DriveApp.getFoldersByName('JURNAL_ARSIP').hasNext()
      ? DriveApp.getFoldersByName('JURNAL_ARSIP').next()
      : DriveApp.createFolder('JURNAL_ARSIP');
    const safeEmail  = email.replace(/[@.]/g,'_');
    const guruFolder = rootArsip.getFoldersByName(safeEmail).hasNext()
      ? rootArsip.getFoldersByName(safeEmail).next()
      : rootArsip.createFolder(safeEmail);
    const xlsFile = guruFolder.createFile(xlsBlob);
    xlsFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  }finally{
    try{ DriveApp.getFileById(tempId).setTrashed(true); }catch(e){}
  }

  for(let i = toArchive.length - 1; i >= 0; i--){
    sh.deleteRow(toArchive[i].rowIdx + 1);
  }

  logAudit('ARSIP_JURNAL_WALI', email, tahun + ' | ' + toArchive.length + ' entri');
  return { success: true, tahun, jumlah: toArchive.length };
}

function importSiswaBinaan(rows){

  assertLicenseActive();

  const auth  = getAuth();
  if(auth.role !== 'admin' && auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }

  if(!Array.isArray(rows) || rows.length === 0){
    throw new Error('Data siswa kosong');
  }

  const valid = rows.filter(r => r.nama_siswa && String(r.nama_siswa).trim());

  if(valid.length === 0){
    throw new Error('Tidak ada data valid yang bisa diimport');
  }

  if(valid.length > 20){
    throw new Error(
      'Maksimal 20 siswa binaan. Anda mencoba import ' + valid.length + ' siswa.'
    );
  }

  const sh    = ensureSiswaBinaanSheet_();
  const data  = sh.getDataRange().getValues();
  const email = auth.email;

  for(let i = data.length - 1; i >= 1; i--){
    const guruWali = String(data[i][4] || '').toLowerCase().trim();
    if(guruWali === email){
      sh.deleteRow(i + 1);
    }
  }

  const now = new Date();
  valid.forEach((r, idx) => {
    const id = now.getTime().toString() + idx;
    sh.appendRow([
      id,
      String(r.nama_siswa || '').trim(),
      String(r.nis        || '').trim(),
      String(r.kelas      || '').trim(),
      email,
      String(r.tahun_masuk || '').trim(),
      String(r.status     || 'aktif').trim().toLowerCase()
    ]);
  });

  logAudit('IMPORT_SISWA_BINAAN', email, valid.length + ' siswa');
  invalidateCache_('SISWA_BINAAN');

  return { status: true, imported: valid.length };
}

/**
 * getKelasSiswaBinaan() — Ambil daftar kelas unik dari siswa binaan guru ini.
 * Dipakai di tab Catatan Siswa agar hanya menampilkan kelas siswa binaan.
 */
function getKelasSiswaBinaan(){
  const auth  = getAuth();
  const email = String(auth.email || '').toLowerCase().trim();
  const sh    = ensureSiswaBinaanSheet_();
  const data  = sh.getDataRange().getValues();
  const kelasSet = {};
  for(let i = 1; i < data.length; i++){
    if(String(data[i][4] || '').toLowerCase().trim() !== email) continue;
    if(String(data[i][6] || '').toLowerCase().trim() === 'nonaktif') continue;
    const k = String(data[i][3] || '').trim();
    if(k) kelasSet[k] = true;
  }
  return Object.keys(kelasSet).sort();
}

/**
 * getSiswaBinaanByKelas(kelas) — Ambil daftar siswa binaan untuk kelas tertentu.
 * Dipakai di form Catatan Siswa (menggantikan getSiswaByKelas untuk guru wali).
 */
function getSiswaBinaanByKelas(kelas){
  const auth  = getAuth();
  const email = String(auth.email || '').toLowerCase().trim();
  const sh    = ensureSiswaBinaanSheet_();
  const data  = sh.getDataRange().getValues();
  const result = [];
  for(let i = 1; i < data.length; i++){
    if(String(data[i][4] || '').toLowerCase().trim() !== email) continue;
    if(String(data[i][6] || '').toLowerCase().trim() === 'nonaktif') continue;
    if(kelas && String(data[i][3] || '').trim() !== String(kelas)) continue;
    result.push({
      nis  : String(data[i][2] || ''),
      nama : String(data[i][1] || ''),
      kelas: String(data[i][3] || '')
    });
  }
  return result;
}
