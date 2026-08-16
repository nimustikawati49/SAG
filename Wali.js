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
      'guru_wali','tahun_masuk','status','tahun_pelajaran'
    ]);
  } else {
    // Migrasi: sheet lama tidak punya kolom tahun_pelajaran sama sekali,
    // jadi kelas siswa binaan tidak pernah dibedakan per tahun ajaran —
    // setiap import ulang menimpa/menghapus semua riwayat lama. Tambahkan
    // kolom ini di akhir (index kolom lama tidak berubah).
    const lastCol = sh.getLastColumn();
    const header = lastCol > 0
      ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||'').toLowerCase().trim(); })
      : [];
    if (header.indexOf('tahun_pelajaran') === -1) {
      sh.getRange(1, lastCol + 1).setValue('tahun_pelajaran');
    }
  }
  return sh;
}

/**
 * _siswaBinaanTahunColIdx_(sh)
 * Index (0-based) kolom tahun_pelajaran di SISWA_BINAAN — dicari lewat
 * header, bukan hardcode, supaya tahan kalau kolomnya sudah pernah
 * dimigrasi ke posisi manapun.
 */
function _siswaBinaanTahunColIdx_(sh){
  const lastCol = sh.getLastColumn();
  const header = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||'').toLowerCase().trim(); })
    : [];
  const idx = header.indexOf('tahun_pelajaran');
  return idx === -1 ? 7 : idx;
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
  const tIdx    = _siswaBinaanTahunColIdx_(sh);
  const data    = sh.getDataRange().getValues();
  const email   = auth.email;
  const tahunAktif = setting.tahun_pelajaran || '';

  let jumlah = 0;
  for(let i = 1; i < data.length; i++){
    const guruWali = String(data[i][4] || '').toLowerCase().trim();
    const status   = String(data[i][6] || '').toLowerCase().trim();
    if(guruWali !== email || status === 'nonaktif') continue;
    const rowTahun = String(data[i][tIdx] || '').trim();
    if(rowTahun && tahunAktif && rowTahun !== tahunAktif) continue;
    jumlah++;
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
  const tIdx  = _siswaBinaanTahunColIdx_(sh);
  const data  = sh.getDataRange().getValues();
  const email = auth.email;
  const tahunAktif = getSetting().tahun_pelajaran || '';

  const PAGE_SIZE = Number(pageSize) || 20;
  const PAGE      = Math.max(1, Number(page) || 1);

  const all = [];
  for(let i = 1; i < data.length; i++){
    const guruWali = String(data[i][4] || '').toLowerCase().trim();
    const status   = String(data[i][6] || '').toLowerCase().trim();
    if(guruWali !== email) continue;
    if(status === 'nonaktif') continue;
    // Baris lama (sebelum kolom tahun_pelajaran ada) dianggap cocok ke
    // periode manapun — baris baru yang sudah punya tahun WAJIB cocok,
    // supaya siswa binaan tahun ajaran lalu tidak nyangkut tampil lagi.
    const rowTahun = String(data[i][tIdx] || '').trim();
    if(rowTahun && tahunAktif && rowTahun !== tahunAktif) continue;
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
  trySyncGuruSummaryAfterMutation_(auth.email, 'SIMPAN_JURNAL_WALI');

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

  const safeEmail = email.replace(/[@.]/g,'_');
  const tahunFolder = getUserNestedFolder_(email, 'wali_dokumentasi_folder', 'WALI_DOKUMENTASI', [safeEmail, tahun]);

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
  const tIdx  = _siswaBinaanTahunColIdx_(sh);
  const data  = sh.getDataRange().getValues();
  const email = auth.email;
  const tahunAktif = getSetting().tahun_pelajaran || '';

  // Hanya hapus baris guru ini di TAHUN AJARAN AKTIF yang sama — riwayat
  // siswa binaan tahun-tahun sebelumnya (mis. saat siswa masih kelas 7/8)
  // tetap utuh, tidak ikut terhapus setiap kali import ulang/tahun baru.
  // Baris lama tanpa tahun_pelajaran (sebelum migrasi ini) dianggap milik
  // periode manapun sehingga tetap ikut ter-replace seperti perilaku lama.
  for(let i = data.length - 1; i >= 1; i--){
    const guruWali = String(data[i][4] || '').toLowerCase().trim();
    if(guruWali !== email) continue;
    const rowTahun = String(data[i][tIdx] || '').trim();
    if(rowTahun && tahunAktif && rowTahun !== tahunAktif) continue;
    sh.deleteRow(i + 1);
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
    sh.getRange(sh.getLastRow(), tIdx + 1).setValue(tahunAktif);
  });

  logAudit('IMPORT_SISWA_BINAAN', email, valid.length + ' siswa | tahun=' + tahunAktif);
  invalidateCache_('SISWA_BINAAN');

  return { status: true, imported: valid.length };
}

/**
 * promoteSiswaBinaan(payload)
 * Naikkan kelas siswa binaan ke tahun ajaran baru TANPA upload ulang dan
 * TANPA menghapus riwayat tahun-tahun sebelumnya — setiap tahun ajaran
 * punya barisnya sendiri (dibedakan kolom tahun_pelajaran), jadi histori
 * "siswa ini kelas 7 di tahun X, kelas 8 di tahun Y, kelas 9 di tahun Z"
 * tetap tersimpan dan bisa dilihat lewat getRiwayatKelasSiswaBinaan().
 */
function promoteSiswaBinaan(payload){
  assertLicenseActive();
  const auth = getAuth();
  if(auth.role !== 'admin' && auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }

  payload = payload || {};
  const fromTahun = String(payload.from_tahun || '').trim();
  const toTahun   = String(payload.to_tahun || '').trim();
  const mapping   = Array.isArray(payload.mapping) ? payload.mapping : [];
  if(!fromTahun || !toTahun) throw new Error('Tahun lama dan tahun baru wajib diisi');
  if(fromTahun === toTahun) throw new Error('Tahun lama dan tahun baru tidak boleh sama');
  if(!mapping.length) return { success: true, processed: 0 };

  const mapObj = {};
  mapping.forEach(function(m){
    const from = String(m.from || '').trim();
    const to   = String(m.to || '').trim();
    if(from && to) mapObj[from] = to;
  });

  const email = String(auth.email || '').toLowerCase().trim();
  const sh    = ensureSiswaBinaanSheet_();
  const tIdx  = _siswaBinaanTahunColIdx_(sh);

  // Hapus dulu baris guru ini di tahun TUJUAN — idempoten kalau promosi
  // ini sampai dijalankan dua kali untuk periode yang sama.
  let data = sh.getDataRange().getValues();
  for(let i = data.length - 1; i >= 1; i--){
    if(String(data[i][4] || '').toLowerCase().trim() !== email) continue;
    if(String(data[i][tIdx] || '').trim() !== toTahun) continue;
    sh.deleteRow(i + 1);
  }

  data = sh.getDataRange().getValues();
  const now = new Date();
  let processed = 0;
  let lulus = 0;
  for(let i = 1; i < data.length; i++){
    if(String(data[i][4] || '').toLowerCase().trim() !== email) continue;
    if(String(data[i][tIdx] || '').trim() !== fromTahun) continue;
    if(String(data[i][6] || '').toLowerCase().trim() === 'nonaktif') continue;

    const kelasLama = String(data[i][3] || '').trim();
    const kelasBaru = mapObj[kelasLama];
    if(!kelasBaru) continue;

    if(String(kelasBaru).toUpperCase() === 'ALUMNI'){
      lulus++;
      continue; // lulus — tidak dibawa ke tahun baru, riwayat lama tetap ada
    }

    const id = now.getTime().toString() + processed;
    sh.appendRow([
      id,
      String(data[i][1] || ''), // nama_siswa
      String(data[i][2] || ''), // nis
      kelasBaru,
      email,
      String(data[i][5] || ''), // tahun_masuk tetap sama
      'aktif',
      toTahun
    ]);
    processed++;
  }

  logAudit('PROMOSI_SISWA_BINAAN', email, fromTahun + ' -> ' + toTahun + ' | naik=' + processed + ' | lulus=' + lulus);
  invalidateCache_('SISWA_BINAAN');
  return { success: true, processed: processed, lulus: lulus };
}

/**
 * getRiwayatKelasSiswaBinaan(nis)
 * Riwayat kelas SATU siswa binaan di semua tahun ajaran yang tercatat
 * (guru wali yang login) — mis. "kelas 7 (2024/2025) → kelas 8
 * (2025/2026) → kelas 9 (2026/2027)". Berlaku untuk jenjang manapun
 * (SD/SMP/SMA/SMK) karena nilai kelas murni teks bebas.
 */
function getRiwayatKelasSiswaBinaan(nis){
  const auth  = getAuth();
  const email = String(auth.email || '').toLowerCase().trim();
  const nisNorm = String(nis || '').trim();
  if(!nisNorm) return [];

  const sh   = ensureSiswaBinaanSheet_();
  const tIdx = _siswaBinaanTahunColIdx_(sh);
  const data = sh.getDataRange().getValues();

  const result = [];
  for(let i = 1; i < data.length; i++){
    if(String(data[i][2] || '').trim() !== nisNorm) continue;
    if(String(data[i][4] || '').toLowerCase().trim() !== email) continue;
    result.push({
      tahun_pelajaran: String(data[i][tIdx] || '').trim() || '-',
      kelas: String(data[i][3] || ''),
      status: String(data[i][6] || 'aktif')
    });
  }
  result.sort(function(a, b){ return String(a.tahun_pelajaran).localeCompare(String(b.tahun_pelajaran)); });
  return result;
}

/**
 * getKelasSiswaBinaan() — Ambil daftar kelas unik dari siswa binaan guru ini
 * di tahun ajaran AKTIF. Dipakai di tab Catatan Siswa agar hanya menampilkan
 * kelas siswa binaan.
 */
function getKelasSiswaBinaan(){
  const auth  = getAuth();
  const email = String(auth.email || '').toLowerCase().trim();
  const sh    = ensureSiswaBinaanSheet_();
  const tIdx  = _siswaBinaanTahunColIdx_(sh);
  const data  = sh.getDataRange().getValues();
  const tahunAktif = getSetting().tahun_pelajaran || '';
  const kelasSet = {};
  for(let i = 1; i < data.length; i++){
    if(String(data[i][4] || '').toLowerCase().trim() !== email) continue;
    if(String(data[i][6] || '').toLowerCase().trim() === 'nonaktif') continue;
    const rowTahun = String(data[i][tIdx] || '').trim();
    if(rowTahun && tahunAktif && rowTahun !== tahunAktif) continue;
    const k = String(data[i][3] || '').trim();
    if(k) kelasSet[k] = true;
  }
  return Object.keys(kelasSet).sort();
}

/**
 * getSiswaBinaanByKelas(kelas) — Ambil daftar siswa binaan untuk kelas
 * tertentu di tahun ajaran AKTIF. Dipakai di form Catatan Siswa (menggantikan
 * getSiswaByKelas untuk guru wali).
 */
function getSiswaBinaanByKelas(kelas){
  const auth  = getAuth();
  const email = String(auth.email || '').toLowerCase().trim();
  const sh    = ensureSiswaBinaanSheet_();
  const tIdx  = _siswaBinaanTahunColIdx_(sh);
  const data  = sh.getDataRange().getValues();
  const tahunAktif = getSetting().tahun_pelajaran || '';
  const result = [];
  for(let i = 1; i < data.length; i++){
    if(String(data[i][4] || '').toLowerCase().trim() !== email) continue;
    if(String(data[i][6] || '').toLowerCase().trim() === 'nonaktif') continue;
    if(kelas && String(data[i][3] || '').trim() !== String(kelas)) continue;
    const rowTahun = String(data[i][tIdx] || '').trim();
    if(rowTahun && tahunAktif && rowTahun !== tahunAktif) continue;
    result.push({
      nis  : String(data[i][2] || ''),
      nama : String(data[i][1] || ''),
      kelas: String(data[i][3] || '')
    });
  }
  return result;
}
