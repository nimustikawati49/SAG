/**
 * Jurnal.js — Manajemen Jurnal Pembelajaran
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

/* ===================================================================
   SISWA (legacy roster) — riwayat kelas per tahun ajaran
   Kolom tetap (semua pembaca lain pakai indeks tetap, bukan header):
   0=Kelas, 1=no_absen, 2=NIS, 3=nama_lengkap, 4=jk, 5=owner guru,
   6=tahun_pelajaran (baru).
   =================================================================== */
const SISWA_TAHUN_COL_ = 6;

/** Tambahkan kolom header 'tahun_pelajaran' kalau sheet SISWA sudah punya
 * header lama tapi belum punya kolom ini (migrasi idempoten). */
function ensureSiswaTahunKolom_() {
  const sh = sheet('SISWA');
  if (!sh) return;
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return;
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').toLowerCase().trim(); });
  if (header[0] !== 'kelas') return; // bukan sheet dengan header standar, jangan sentuh
  if (header.indexOf('tahun_pelajaran') === -1) {
    sh.getRange(1, lastCol + 1).setValue('tahun_pelajaran');
  }
}

/** true jika baris SISWA cocok tahun ajaran aktif. Baris lama tanpa
 * tahun_pelajaran (blank, sebelum migrasi ini) dianggap cocok periode
 * manapun supaya data yang sudah ada tetap tampil seperti biasa. */
function _siswaRowMatchesPeriode_(row, tahunAktif) {
  const rowTahun = String(row[SISWA_TAHUN_COL_] || '').trim();
  return !rowTahun || !tahunAktif || rowTahun === String(tahunAktif).trim();
}

function getKelas(){
  ensureAcademicSchema_();
  const auth = getAuth();

  if(!auth.email) return [];

  const kelasAktif = getKelasDiampuAktifForUser_(auth.email);
  if (kelasAktif && kelasAktif.length) return kelasAktif;

  if(!sheet('SISWA')) return [];

  const tahunAktif = getSetting().tahun_pelajaran || '';

  // Cache SISWA sheet (TTL 120s) — invalidated on upload siswa
  const data = getSheetCached('SISWA', 120).slice(1);

  const kelas = data
    .filter(r => String(r[5]).toLowerCase().trim() === auth.email)
    .filter(r => _siswaRowMatchesPeriode_(r, tahunAktif))
    .map(r => r[0])
    .filter(Boolean)
    // Siswa yang sudah "lulus" lewat Promosi Siswa ditandai Kelas=ALUMNI —
    // jangan tampil sebagai pilihan kelas aktif untuk isi jurnal baru.
    .filter(k => String(k).toUpperCase() !== 'ALUMNI');

  return [...new Set(kelas)].sort();
}

function getSiswaByKelas(kelas){
  ensureAcademicSchema_();
  const auth = getAuth();
  if(!auth.email) return [];

  const fromRiwayat = getSiswaAktifByKelasForUser_(kelas, auth.email);
  if (fromRiwayat && fromRiwayat.length) return fromRiwayat;

  if(!sheet('SISWA')) return [];

  const tahunAktif = getSetting().tahun_pelajaran || '';

  // Cache SISWA sheet (TTL 120s)
  const data = getSheetCached('SISWA', 120).slice(1);

  return data
    .filter(r =>
      r[0] === kelas &&
      String(r[5]).toLowerCase().trim() === auth.email &&
      _siswaRowMatchesPeriode_(r, tahunAktif)
    )
    .map(r => ({
      no_absen: r[1],
      nis     : r[2],
      nama    : r[3],
      jk      : r[4]
    }));
}

function importSiswa(rows){
  ensureAcademicSchema_();
  const auth = getAuth();

  if(auth.role !== 'admin' && auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }

  const sh = sheet('SISWA');
  if(!sh) return []; // belum ada data siswa, kembalikan kosong (bukan error)

  ensureSiswaTahunKolom_();

  const setting = getSetting();
  const tahunAktif = setting.tahun_pelajaran || getLegacyDefaultYear_();
  const semesterAktif = setting.semester || 'Ganjil';

  // Parser fleksibel: tetap menerima template lama, plus kolom baru untuk multi-tahun.
  const header = (rows[0] || []).map(function(h){ return String(h || '').toLowerCase().trim(); });
  const idx = {
    kelas: header.indexOf('kelas'),
    no_absen: header.indexOf('no_absen'),
    nis: header.indexOf('nis'),
    nisn: header.indexOf('nisn'),
    nama: header.indexOf('nama'),
    jk: header.indexOf('jk'),
    ttl: header.indexOf('ttl'),
    alamat: header.indexOf('alamat'),
    orang_tua: header.indexOf('orang_tua'),
    kontak: header.indexOf('kontak'),
    status: header.indexOf('status')
  };

  const normalized = rows.slice(1).map(function(r){
    return {
      kelas: String((idx.kelas > -1 ? r[idx.kelas] : r[0]) || '').trim(),
      no_absen: String((idx.no_absen > -1 ? r[idx.no_absen] : r[1]) || '').trim(),
      nis: String((idx.nis > -1 ? r[idx.nis] : r[2]) || '').trim(),
      nisn: String((idx.nisn > -1 ? r[idx.nisn] : '') || '').trim(),
      nama: String((idx.nama > -1 ? r[idx.nama] : r[3]) || '').trim(),
      jk: String((idx.jk > -1 ? r[idx.jk] : r[4]) || '').trim(),
      ttl: String((idx.ttl > -1 ? r[idx.ttl] : '') || '').trim(),
      alamat: String((idx.alamat > -1 ? r[idx.alamat] : '') || '').trim(),
      orang_tua: String((idx.orang_tua > -1 ? r[idx.orang_tua] : '') || '').trim(),
      kontak: String((idx.kontak > -1 ? r[idx.kontak] : '') || '').trim(),
      status: String((idx.status > -1 ? r[idx.status] : 'AKTIF') || 'AKTIF').trim()
    };
  }).filter(function(r){ return r.kelas && r.nama && (r.nis || r.nisn); });

  // Hapus dulu baris LAMA milik guru ini untuk kelas yang sedang diupload
  // ulang, TAPI hanya di tahun ajaran aktif — supaya CSV yang diupload
  // ulang (mis. perbaiki typo) menggantikan data lama, tanpa duplikat, dan
  // TANPA menghapus riwayat kelas siswa di tahun-tahun ajaran sebelumnya.
  //
  // Ditulis sebagai SATU batch read + SATU batch write (bukan deleteRow/
  // appendRow per baris di dalam loop) — dengan siswa yang banyak (30-40+),
  // ratusan panggilan Sheets API satu-satu itulah yang bikin proses ini
  // makan waktu sampai bermenit-menit dan progress bar di UI terlihat
  // "macet" karena responsnya belum juga kembali ke browser.
  const kelasDiupload = new Set(normalized.map(function(r){ return r.kelas; }));
  const emailNorm = String(auth.email).toLowerCase().trim();
  const numCols = Math.max(sh.getLastColumn(), SISWA_TAHUN_COL_ + 1);
  const lastRowBefore = sh.getLastRow();
  const allRows = lastRowBefore > 1 ? sh.getRange(2, 1, lastRowBefore - 1, numCols).getValues() : [];

  const keepRows = allRows.filter(function(r){
    const owner = String(r[5] || '').toLowerCase().trim();
    if (owner !== emailNorm) return true;
    if (!kelasDiupload.has(String(r[0] || '').trim())) return true;
    if (!_siswaRowMatchesPeriode_(r, tahunAktif)) return true;
    return false; // baris lama yang sedang diganti — buang
  });

  const rowsForMaster = [];
  const newSiswaRows = normalized.map(function(r){
    rowsForMaster.push({
      nis: r.nis,
      nisn: r.nisn,
      nama: r.nama,
      jk: r.jk,
      ttl: r.ttl,
      alamat: r.alamat,
      orang_tua: r.orang_tua,
      kontak: r.kontak,
      kelas: r.kelas,
      status: r.status || 'AKTIF'
    });
    return [r.kelas, r.no_absen, r.nis, r.nama, r.jk, auth.email, tahunAktif];
  });

  const combined = keepRows.concat(newSiswaRows).map(function(r){
    if (r.length >= numCols) return r.slice(0, numCols);
    return r.concat(Array(numCols - r.length).fill(''));
  });

  if (lastRowBefore > 1) {
    sh.getRange(2, 1, lastRowBefore - 1, numCols).clearContent();
  }
  if (combined.length) {
    sh.getRange(2, 1, combined.length, numCols).setValues(combined);
  }

  // Sinkronkan ke arsitektur baru (MasterSiswa + RiwayatKelas) tanpa merusak format lama.
  if (rowsForMaster.length) {
    importSiswaBaru({
      tahun_pelajaran: tahunAktif,
      semester: semesterAktif,
      rows: rowsForMaster
    });
  }

  invalidateCache_('SISWA');
  trySyncGuruSummaryAfterMutation_(auth.email, 'IMPORT_SISWA');
  return true;
}

function getRekapKelasPerGuru(){
  const auth = getAuth();
  if(auth.role !== 'superadmin'){
    throw new Error('Akses ditolak');
  }

  const d = sheet('SISWA').getDataRange().getValues();
  const map = {};

  for(let i=1;i<d.length;i++){
    const kelas = String(d[i][0] || '').trim();
    const owner = String(d[i][5] || '').toLowerCase().trim();

    if(!kelas || !owner) continue;

    if(!map[owner]) map[owner] = new Set();
    map[owner].add(kelas);
  }

  return Object.keys(map).map(email => ({
    email,
    jumlah_kelas: map[email].size,
    kelas: Array.from(map[email]).sort()
  }));
}

function getGuruFolder_(email){

  const setting = getSetting();
  const tahun = setting.tahun_pelajaran || 'Tanpa_Tahun';

  const safeEmail = email.replace(/[@.]/g,'_');
  return getUserNestedFolder_(email, 'dokumentasi_folder', FOLDER_DOKUMENTASI, [safeEmail, tahun]);
}

function compressImage_(base64, type){

  return Utilities.newBlob(
    Utilities.base64Decode(base64),
    type || 'image/jpeg'
  );

}

/**
 * getNextPertemuanKe(kelas)
 * Sarankan nomor pertemuan BERIKUTNYA untuk kelas ini — dihitung dari
 * pertemuan TERBESAR yang sudah ada untuk guru+kelas+tahun/semester aktif
 * yang sama, +1 (bukan sekadar jumlah baris, supaya tidak salah saran
 * kalau ada pertemuan yang pernah dihapus di tengah). Dipanggil client
 * (loadSiswa di scripts-jurnal.html) saat kelas dipilih untuk entri BARU
 * — dropdown "Pertemuan Ke" sebelumnya statis 1-24 tanpa saran apa pun,
 * jadi diam-diam nyangkut di pilihan pertama (1) kalau guru lupa
 * menggantinya manual. Guru tetap bisa override manual kalau perlu.
 */
function getNextPertemuanKe(kelas) {
  const auth = getAuth();
  if (!auth.email || !kelas) return 1;

  const setting = getSetting();
  const activeTahun = setting.tahun_pelajaran || '';
  const activeSemester = setting.semester || '';

  const sh = sheet('JURNAL');
  if (!sh) return 1;
  const rows = sh.getDataRange().getValues();

  let maxPertemuan = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[12] || '').toLowerCase().trim() !== auth.email) continue;
    if (String(row[2] || '').trim() !== String(kelas).trim()) continue;
    if (activeTahun && row[18] && String(row[18]) !== activeTahun) continue;
    if (activeSemester && row[13] && String(row[13]) !== activeSemester) continue;
    const p = parseInt(row[4], 10);
    if (!isNaN(p) && p > maxPertemuan) maxPertemuan = p;
  }

  return maxPertemuan + 1;
}

/**
 * _restFindOrCreateFolder_(token, name, parentId)
 * Cari/buat folder lewat Drive API v3 MURNI (UrlFetchApp + OAuth token),
 * bukan lewat service DriveApp bawaan — jalur kode yang BENAR-BENAR
 * berbeda. parentId null = di My Drive root.
 */
function _restFindOrCreateFolder_(token, name, parentId) {
  var safeName = String(name).replace(/'/g, "\\'");
  var q = "name = '" + safeName + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false" +
    (parentId ? " and '" + parentId + "' in parents" : " and 'root' in parents");
  var searchRes = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(q) + '&fields=files(id)',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  var searchJson = JSON.parse(searchRes.getContentText() || '{}');
  if (searchJson.files && searchJson.files.length) return searchJson.files[0].id;

  var metadata = { name: name, mimeType: 'application/vnd.google-apps.folder' };
  if (parentId) metadata.parents = [parentId];
  var createRes = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files?fields=id',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(metadata),
      muteHttpExceptions: true
    }
  );
  var createJson = JSON.parse(createRes.getContentText() || '{}');
  if (!createJson.id) throw new Error('Gagal buat folder "' + name + '" via REST API: ' + createRes.getContentText());
  return createJson.id;
}

/**
 * _uploadFotoViaRestApi_(email, jurnalId, f)
 * FALLBACK saat DriveApp (service bawaan) gagal dengan "Access denied" —
 * upload lewat Drive API v3 langsung via UrlFetchApp + OAuth token,
 * jalur kode yang BERBEDA dari DriveApp (kadang lolos pembatasan yang
 * spesifik memblokir service DriveApp tapi tidak panggilan REST API
 * murni, tergantung kebijakan domain Google Workspace). Folder yang
 * dipakai sama seperti getGuruFolder_ (FOLDER_DOKUMENTASI/emailSlug/
 * tahun) supaya tetap terorganisir sama, terlepas dari jalur mana yang
 * berhasil.
 */
function _uploadFotoViaRestApi_(email, jurnalId, f) {
  var token = ScriptApp.getOAuthToken();
  var setting = getSetting();
  var tahun = setting.tahun_pelajaran || 'Tanpa_Tahun';
  var safeEmail = email.replace(/[@.]/g, '_');

  var folderId = _restFindOrCreateFolder_(token, FOLDER_DOKUMENTASI, null);
  folderId = _restFindOrCreateFolder_(token, safeEmail, folderId);
  folderId = _restFindOrCreateFolder_(token, tahun, folderId);

  var base64 = f.data.split(',')[1];
  var mimeType = f.type || 'image/jpeg';
  var fileName = jurnalId + '_' + (f.name || 'foto');

  var boundary = 'jgd_' + jurnalId + '_' + Math.random().toString(36).slice(2);
  var metadata = { name: fileName, parents: [folderId] };
  var body =
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: ' + mimeType + '\r\n' +
    'Content-Transfer-Encoding: base64\r\n\r\n' +
    base64 + '\r\n' +
    '--' + boundary + '--';

  var uploadRes = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id',
    {
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      headers: { Authorization: 'Bearer ' + token },
      payload: body,
      muteHttpExceptions: true
    }
  );
  var uploadJson = JSON.parse(uploadRes.getContentText() || '{}');
  if (!uploadJson.id) throw new Error('Upload REST API gagal: ' + uploadRes.getContentText());

  try {
    UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + uploadJson.id + '/permissions',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ role: 'reader', type: 'anyone' }),
        muteHttpExceptions: true
      }
    );
  } catch (eShare) { /* file tetap terupload walau share publiknya gagal */ }

  return { full: "https://drive.google.com/uc?id=" + uploadJson.id };
}

/**
 * _uploadJurnalFotoList_(email, jurnalId, dokumentasiArr)
 * Upload SEMUA foto di dokumentasiArr — coba DriveApp dulu (jalur normal,
 * lewat getGuruFolder_), kalau itu melempar error APA PUN, coba lagi
 * SATU KALI lewat _uploadFotoViaRestApi_ (jalur REST API murni) sebelum
 * benar-benar menyerah. Melempar error kalau KEDUA jalur gagal — dipakai
 * simpanJurnal/updateJurnal yang membungkusnya sendiri dengan try/catch
 * supaya jurnal tetap tersimpan walau upload foto akhirnya tetap gagal.
 */
function _uploadJurnalFotoList_(email, jurnalId, dokumentasiArr) {
  var result = [];
  var usedFallback = false;

  var folder = null;
  try { folder = getGuruFolder_(email); } catch (e) { folder = null; }

  dokumentasiArr.forEach(function (f) {
    if (!f || !f.data) return;

    if (folder) {
      try {
        var base64 = f.data.split(',')[1];
        var blob = Utilities.newBlob(
          Utilities.base64Decode(base64),
          f.type || 'image/jpeg',
          jurnalId + '_' + (f.name || 'foto')
        );
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        result.push({ full: "https://drive.google.com/uc?id=" + file.getId() });
        return;
      } catch (eDriveApp) {
        try { logError_('uploadFoto/DriveApp_gagal', eDriveApp); } catch (e2) {}
        // lanjut ke fallback REST API di bawah, jangan return dulu
      }
    }

    // Fallback REST API — kalau ini JUGA gagal, error-nya dilempar ke
    // atas (dibiarkan, tidak ditangkap di sini) supaya pemanggil
    // (simpanJurnal/updateJurnal) tahu upload benar-benar gagal total.
    usedFallback = true;
    result.push(_uploadFotoViaRestApi_(email, jurnalId, f));
  });

  if (usedFallback) {
    try { logAudit('UPLOAD_FOTO_VIA_REST_FALLBACK', email, jurnalId + ' | DriveApp gagal, berhasil lewat REST API'); } catch (e) {}
  }

  return result;
}

function simpanJurnal(data){

  ensureAcademicSchema_();

  assertLicenseActive();

  const sh = sheet('JURNAL');
  const absSheet = sheet('ABSENSI');
  if (!sh || !absSheet) throw new Error('Sheet operasional jurnal belum tersedia');

  const jurnalId = Date.now().toString();
  const now = new Date();

  const setting = getSetting();
  const period = getUserAcademicPeriod(getAuth().email);
  const auth = getAuth();
  const email = auth.email;

  const jamKe = String(
    data.jam_ke ?? data.jamKe ?? data.jam ?? ''
  ).trim();

  if(!data.kelas){
    throw new Error("Kelas belum dipilih");
  }

  // ✅ DUPLICATE VALIDATION: cek jurnal sama (email + kelas + pertemuan + hari yang sama)
  const existingRows = sh.getDataRange().getValues();
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  for(let i = 1; i < existingRows.length; i++){
    const r = existingRows[i];
    const rowEmail    = String(r[12] || '').toLowerCase().trim();
    const rowKelas    = String(r[2]  || '').trim();
    const rowPert     = String(r[4]  || '').trim();
    const rowTanggal  = r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    if(
      rowEmail   === email &&
      rowKelas   === String(data.kelas   || '').trim() &&
      rowPert    === String(data.pertemuan || '').trim() &&
      rowTanggal === todayStr
    ){
      throw new Error('Jurnal duplikat: kelas ' + data.kelas + ' pertemuan ' + data.pertemuan + ' sudah ada hari ini.');
    }
  }

  let fotoArr = [];
  let fotoUrls = [];
  let fotoGagal = false;

  // Upload foto SENGAJA dibungkus try/catch terpisah dari sisa proses simpan
  // — kalau DriveApp diblokir (mis. kebijakan domain Google Workspace yang
  // membatasi akses Drive untuk web app, di luar kendali SuperAdmin
  // aplikasi ini — pernah terjadi persis begini di deployment ini), jurnal
  // TETAP tersimpan (cuma butuh akses Sheets, bukan Drive) alih-alih gagal
  // total cuma karena fotonya tidak bisa diupload. fotoGagal dikirim balik
  // ke client supaya guru diberi tahu jelas, bukan diam-diam kehilangan foto.
  if(data.dokumentasi && data.dokumentasi.length){
    try {
      fotoArr = _uploadJurnalFotoList_(email, jurnalId, data.dokumentasi);
      fotoUrls = fotoArr.map(f => f.full).filter(Boolean);
    } catch (fotoErr) {
      fotoGagal = true;
      fotoArr = [];
      fotoUrls = [];
      try { logError_('simpanJurnal/uploadFoto', fotoErr); } catch (e2) {}
    }
  }

  const dokumentasi_json = JSON.stringify(fotoArr);
  const foto_urls = fotoUrls.join(',');

  const jumlah = data.absensi ? data.absensi.length : 0;

  sh.appendRow([
    jurnalId,
    now,
    data.kelas || '',
    jamKe,
    data.pertemuan || '',
    data.materi || '',
    data.tujuan || '',
    data.asesmen || '-',
    '',
    now,
    'OPEN',
    0,
    email,
    setting.semester || '',
    dokumentasi_json,
    data.refleksi || '-',
    jumlah,
    foto_urls,
    period.tahun_pelajaran || setting.tahun_pelajaran || '',
    (data.mapel || (setting.mapel_list && setting.mapel_list[0]) || setting.mata_pelajaran || '')
  ]);

  if(data.absensi && data.absensi.length){
    data.absensi.forEach(a=>{
      absSheet.appendRow([
        jurnalId,
        data.kelas,
        a.nis,
        a.status
      ]);
    });
  }

  try {
    // Sinkron summary jangan memblokir status sukses save jurnal.
    setTimeout(function(){
      trySyncGuruSummaryAfterMutation_(email, 'SIMPAN_JURNAL');
    }, 0);
  } catch (syncErr) {
    logError_('simpanJurnal/asyncSync', syncErr);
  }

  return { status:true, fotoGagal: fotoGagal };
}

function getJurnalDetail(id){

  const auth = getAuth();
  const j = sheet('JURNAL').getDataRange().getValues();
  const a = sheet('ABSENSI').getDataRange().getValues();

  for(let i=1;i<j.length;i++){

    if(auth.role !== 'superadmin' && j[i][12] !== auth.email) continue;

    if(j[i][0] == id){

      let fotoArray = [];
      try{
        if(j[i][14]){
          fotoArray = JSON.parse(j[i][14]);
        }
      }catch(e){
        fotoArray = [];
      }

      return {
        jurnalId : id,
        kelas    : j[i][2] || '',
        jam_ke   : j[i][3] || '',
        pertemuan: j[i][4] || '',
        materi   : j[i][5] || '',
        tujuan   : j[i][6] || '',
        asesmen  : j[i][7] || '-',
        refleksi : j[i][15] || '-',
        foto     : fotoArray,
        edit_count: Number(j[i][11]) || 0,
        absensi  : a
          .filter(x=>x[0]==id)
          .map(x=>({nis:x[2],status:x[3]}))
      };
    }
  }

  return null;
}

function updateJurnal(data){

  assertLicenseActive();

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('JURNAL');
  const values = sh.getDataRange().getValues();

  const rowIndex = values.findIndex(r => r[0] == data.jurnalId);
  if(rowIndex === -1){
    return { status:false, msg:'Jurnal tidak ditemukan' };
  }

  const rowNumber = rowIndex + 1;

  sh.getRange(rowNumber, 3).setValue(data.kelas);
  sh.getRange(rowNumber, 4).setValue(data.jam_ke);
  sh.getRange(rowNumber, 5).setValue(data.pertemuan);
  sh.getRange(rowNumber, 6).setValue(data.materi);
  sh.getRange(rowNumber, 7).setValue(data.tujuan);
  sh.getRange(rowNumber, 8).setValue(data.asesmen);
  sh.getRange(rowNumber, 16).setValue(data.refleksi);

  // edit_count (kolom 12) — dulu TIDAK PERNAH bertambah sama sekali di
  // sini (bug lama), jadi badge "0/3" di Riwayat Jurnal tidak pernah naik
  // walau sudah diedit berkali-kali. Sekarang benar-benar dinaikkan tiap
  // edit berhasil.
  const prevEditCount = Number(values[rowIndex][11]) || 0;
  sh.getRange(rowNumber, 12).setValue(prevEditCount + 1);

  // Foto — CATATAN PENTING: client (scripts-jurnal.html) SELALU mengirim
  // foto baru lewat key 'dokumentasi' (array base64), sama persis seperti
  // simpanJurnal — TIDAK PERNAH mengirim 'keepFotos'/'replaceFotos'/
  // 'newFotos' (versi lama fungsi ini salah asumsi menunggu key-key itu,
  // yang tidak pernah ada isinya → finalFotos selalu [] → foto lama
  // KETIMPA KOSONG tiap kali diedit, itu sebab laporan "foto hilang saat
  // edit"). Sekarang: kalau ada foto BARU di 'dokumentasi', upload &
  // GANTIKAN foto lama sepenuhnya (sama seperti simpanJurnal). Kalau
  // 'dokumentasi' kosong (guru tidak menyentuh foto saat edit), foto LAMA
  // dipertahankan apa adanya — TIDAK dihapus cuma karena tidak diisi ulang.
  let existingFotos = [];
  try{
    const raw = values[rowIndex][14];
    if(raw){
      existingFotos = JSON.parse(raw);
      if(!Array.isArray(existingFotos)) existingFotos = [];
    }
  }catch(e){
    existingFotos = [];
  }

  let finalFotos = existingFotos;
  let fotoGagal = false;

  if(data.dokumentasi && data.dokumentasi.length){
    try {
      // Upload sukses SEMUA baru menggantikan foto lama — kalau ada yang
      // gagal di tengah jalan (exception dilempar dari dalam), foto lama
      // TETAP dipertahankan (tidak separuh terganti separuh hilang).
      finalFotos = _uploadJurnalFotoList_(getAuth().email, data.jurnalId, data.dokumentasi);
    } catch (fotoErr) {
      fotoGagal = true;
      finalFotos = existingFotos; // gagal upload — foto lama tetap dipertahankan, bukan dihapus
      try { logError_('updateJurnal/uploadFoto', fotoErr); } catch (e2) {}
    }
  }

  sh.getRange(rowNumber, 15)
    .setValue(JSON.stringify(finalFotos));
  sh.getRange(rowNumber, 18)
    .setValue(finalFotos.map(f => f.full).filter(Boolean).join(','));

  // Absensi — pengaman: kalau data.absensi kosong/tidak terkirim (mis.
  // tabel siswa belum sempat termuat penuh saat klik Simpan — race
  // condition async), JANGAN hapus data absensi yang sudah ada. Dulu
  // baris lama selalu dihapus dulu baru diisi ulang TANPA cek ini, jadi
  // kalau payload-nya kebetulan kosong, hasilnya "NIHIL" walau guru tidak
  // benar-benar mengubah absensi (laporan bug ini).
  if(data.absensi && data.absensi.length){
    const shAbs = ss.getSheetByName('ABSENSI');
    const absData = shAbs.getDataRange().getValues();

    for(let i = absData.length - 1; i > 0; i--){
      if(absData[i][0] == data.jurnalId){
        shAbs.deleteRow(i + 1);
      }
    }

    data.absensi.forEach(a => {
      shAbs.appendRow([
        data.jurnalId,
        data.kelas,
        a.nis,
        a.status
      ]);
    });
  }

  return { status:true, fotoGagal: fotoGagal };
}

function hapusJurnal(id){

  assertLicenseActive();
  const auth = getAuth();

  const j = sheet('JURNAL');
  const rows = j.getDataRange().getValues();

  for(let i=1;i<rows.length;i++){

    if(auth.role !== 'superadmin' && rows[i][12] !== auth.email) continue;

    if(rows[i][0]==id){

      try{
        const foto = rows[i][14] ? JSON.parse(rows[i][14]) : [];
        foto.forEach(f=>{
          if(f.id){
            try{
              DriveApp.getFileById(f.id).setTrashed(true);
            }catch(e){ console.error('[JGD] trash foto:', e.message||e); }
          }
        });
      }catch(e){ console.error('[JGD] parse foto hapus:', e.message||e); }

      j.deleteRow(i+1);
      break;
    }
  }

  const a = sheet('ABSENSI');
  for(let i=a.getLastRow();i>1;i--){
    if(a.getRange(i,1).getValue()==id){
      a.deleteRow(i);
    }
  }
}

/**
 * Paginated riwayat jurnal — returns {rows, total, page, pageSize}
 * so the client does NOT need to download the full dataset.
 */
function getRiwayatJurnalPaged(page, pageSize, filters) {
  page     = Math.max(1, parseInt(page)     || 1);
  pageSize = Math.max(5, parseInt(pageSize) || 20);
  filters  = filters || {};

  const fDari   = filters.dari   ? new Date(filters.dari)   : null;
  const fSampai = filters.sampai ? new Date(filters.sampai + 'T23:59:59') : null;
  const fKelas  = (filters.kelas || '').trim().toLowerCase();
  const fCari   = (filters.cari  || '').trim().toLowerCase();

  const auth        = getAuth();
  const setting     = getSetting();
  const activeTahun = setting.tahun_pelajaran || '';
  const ss = getSpreadsheet_();

  const shJurnal = ss.getSheetByName('JURNAL');
  const shAbs    = ss.getSheetByName('ABSENSI');
  const shSiswa  = ss.getSheetByName('SISWA');

  if (!shJurnal) return { rows: [], total: 0, page, pageSize };

  const jurnalData = shJurnal.getDataRange().getValues();

  // 1. Lightweight filter to get matching row indices (newest first)
  const matchIdx = [];
  for (let i = 1; i < jurnalData.length; i++) {
    const row = jurnalData[i];
    if (row[12] !== auth.email) continue;
    if (activeTahun && row[18] && String(row[18]) !== activeTahun) continue;
    // --- client filters ---
    if (fDari || fSampai) {
      const tgl = new Date(row[1]);
      if (fDari   && tgl < fDari)   continue;
      if (fSampai && tgl > fSampai) continue;
    }
    if (fKelas && String(row[2] || '').toLowerCase() !== fKelas) continue;
    if (fCari  && !String(row[5] || '').toLowerCase().includes(fCari)) continue;
    matchIdx.push(i);
  }
  // Reverse: newest first
  matchIdx.reverse();

  const total  = matchIdx.length;
  const start  = (page - 1) * pageSize;
  const end    = Math.min(start + pageSize, total);
  const pageIdx = matchIdx.slice(start, end);

  if (pageIdx.length === 0) return { rows: [], total, page, pageSize };

  // 2. Heavy work only for current page rows
  const absData   = shAbs   ? shAbs.getDataRange().getValues()   : [];
  const siswaData = shSiswa ? shSiswa.getDataRange().getValues() : [];

  const rows = [];
  pageIdx.forEach((i, localIdx) => {
    const row      = jurnalData[i];
    const jurnalId = row[0];
    const kelas    = row[2] || '-';
    const absRows  = absData.filter(a => a[0] == jurnalId);

    let sakit = [], izin = [], alpha = [];
    absRows.forEach(a => {
      if (a[3] === 'H') return;
      const nis   = String(a[2] || '').trim();
      const siswa = siswaData.find(s => s[0] == kelas && String(s[2]) == nis);
      const nama  = siswa ? siswa[3] : getNamaSiswaByNis_(nis);
      if (a[3] === 'S') sakit.push(nama);
      if (a[3] === 'I') izin.push(nama);
      if (a[3] === 'A') alpha.push(nama);
    });

    let absHtml = '✅ NIHIL';
    if (sakit.length || izin.length || alpha.length) {
      absHtml = '';
      if (sakit.length) absHtml += `🟠 <b>S (Sakit)</b><br>${sakit.join('<br>')}<br><br>`;
      if (izin.length)  absHtml += `🔵 <b>I (Izin)</b><br>${izin.join('<br>')}<br><br>`;
      if (alpha.length) absHtml += `🔴 <b>A (Alpha)</b><br>${alpha.join('<br>')}`;
    }

    // foto — thumbnail visual (bukan cuma teks link), pakai endpoint
    // thumbnail resmi Drive (lh3.googleusercontent.com/d/ID=wN-hN) supaya
    // gambar kecil benar-benar tampil di tabel, bukan cuma tulisan
    // "📷 Foto 1" yang harus diklik dulu untuk lihat isinya. Tetap
    // dibungkus <a> supaya klik tetap buka versi penuh di tab baru.
    let fotoHtml = '-';
    try {
      let fotoThumbs = [];
      const pushThumb = (fileId) => {
        if (!fileId) return;
        fotoThumbs.push(
          `<a href="https://drive.google.com/file/d/${fileId}/view" target="_blank" title="Sorot untuk perbesar, klik untuk buka penuh" ` +
          `onmouseenter="_fotoZoomShow_(event,'${fileId}')" onmousemove="_fotoZoomMove_(event)" onmouseleave="_fotoZoomHide_()" ` +
          `ontouchstart="_fotoZoomShow_(event,'${fileId}')" ontouchend="_fotoZoomHide_()" ontouchcancel="_fotoZoomHide_()">` +
          `<img src="https://lh3.googleusercontent.com/d/${fileId}=w80-h80" ` +
          `style="width:44px;height:44px;object-fit:cover;border-radius:8px;border:1px solid #e5e7eb;margin:2px" ` +
          `loading="lazy" alt="Dokumentasi"></a>`
        );
      };
      if (row[14]) {
        const parsed = JSON.parse(row[14]);
        if (Array.isArray(parsed)) {
          parsed.forEach((f) => {
            if (!f || !f.full) return;
            const m = f.full.match(/[-\w]{25,}/);
            if (m) pushThumb(m[0]);
          });
        }
      }
      if (fotoThumbs.length === 0 && row[17]) {
        String(row[17]).trim().split(',').filter(u=>u.trim()).forEach((u) => {
          const m = u.match(/[-\w]{25,}/);
          if (m) pushThumb(m[0]);
        });
      }
      if (fotoThumbs.length) fotoHtml = `<div style="display:flex;flex-wrap:wrap;gap:2px">${fotoThumbs.join('')}</div>`;
    } catch(e) { fotoHtml = '-'; }

    rows.push({
      no        : start + localIdx + 1,
      jurnalId  : jurnalId,
      tanggal   : Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      kelas, jam_ke: row[3] || '', pertemuan: row[4] || '',
      materi    : row[5] || '', asesmen: row[7] || '-',
      foto      : fotoHtml, refleksi: row[15] || '-',
      absensi   : absHtml, edit_count: row[11] || 0,
      locked    : false
    });
  });

  return { rows, total, page, pageSize };
}

function getRiwayatJurnal(){

  const auth = getAuth();
  const setting     = getSetting();
  const activeTahun = setting.tahun_pelajaran || '';
  const ss = getSpreadsheet_();

  const shJurnal = ss.getSheetByName('JURNAL');
  const shAbs    = ss.getSheetByName('ABSENSI');
  const shSiswa  = ss.getSheetByName('SISWA');

  if(!shJurnal) return [];

  const jurnalData = shJurnal.getDataRange().getValues();
  const absData    = shAbs ? shAbs.getDataRange().getValues() : [];
  const siswaData  = shSiswa ? shSiswa.getDataRange().getValues() : [];

  let result = [];

  for(let i = 1; i < jurnalData.length; i++){

    const row = jurnalData[i];

    if(auth.role !== 'superadmin' && row[12] !== auth.email) continue;
    if(activeTahun && row[18] && String(row[18]) !== activeTahun) continue;

    const jurnalId = row[0];
    const kelas    = row[2] || '-';

    const absRows = absData.filter(a => a[0] == jurnalId);

    let sakit = [];
    let izin  = [];
    let alpha = [];

    absRows.forEach(a => {

      const nis = String(a[2] || '').trim();
      const status = a[3];

      if(status === 'H') return;

      const siswa = siswaData.find(s =>
        s[0] == kelas && String(s[2]) == nis
      );

      const nama = siswa ? siswa[3] : getNamaSiswaByNis_(nis);

      if(status === 'S') sakit.push(nama);
      if(status === 'I') izin.push(nama);
      if(status === 'A') alpha.push(nama);
    });

    let absHtml = '\u2705 NIHIL';

    if(sakit.length || izin.length || alpha.length){
      absHtml = '';

      if(sakit.length){
        absHtml += `\uD83D\uDFE0 <b>S (Sakit)</b><br>${sakit.join('<br>')}<br><br>`;
      }
      if(izin.length){
        absHtml += `\uD83D\uDD35 <b>I (Izin)</b><br>${izin.join('<br>')}<br><br>`;
      }
      if(alpha.length){
        absHtml += `\uD83D\uDD34 <b>A (Alpha)</b><br>${alpha.join('<br>')}`;
      }
    }

    let fotoHtml = '-';

    try{

      let fotoLinks = [];

      if(row[14]){

        let parsed = [];

        try{
          parsed = JSON.parse(row[14]);
        }catch(e){
          parsed = [];
        }

        if(Array.isArray(parsed) && parsed.length){

          parsed.forEach((f,idx)=>{

            if(!f || !f.full) return;

            const match = f.full.match(/[-\w]{25,}/);
            if(!match) return;

            const fileId = match[0];

            const stableUrl =
              "https://drive.google.com/file/d/" +
              fileId +
              "/view";

            fotoLinks.push(
              '<a href="' + stableUrl + '" target="_blank" ' +
              'style="font-weight:600;color:#4f46e5;text-decoration:none;">' +
              '\uD83D\uDCF7 Foto ' + (idx+1) +
              '</a>'
            );
          });
        }
      }

      if(fotoLinks.length === 0 && row[17]){

        const raw = String(row[17]).trim();

        if(raw){

          const urls = raw.split(',').filter(u=>u.trim());

          urls.forEach((u,idx)=>{

            const match = u.match(/[-\w]{25,}/);
            if(!match) return;

            const fileId = match[0];

            const stableUrl =
              "https://drive.google.com/file/d/" +
              fileId +
              "/view";

            fotoLinks.push(
              '<a href="' + stableUrl + '" target="_blank" ' +
              'style="font-weight:600;color:#4f46e5;text-decoration:none;">' +
              '\uD83D\uDCF7 Foto ' + (idx+1) +
              '</a>'
            );
          });
        }
      }

      if(fotoLinks.length > 0){
        fotoHtml = fotoLinks.join('<br>');
      }

    }catch(e){
      fotoHtml = '-';
    }

    result.push({
      no        : result.length + 1,
      jurnalId  : jurnalId,
      tanggal   : Utilities.formatDate(
                    new Date(row[1]),
                    Session.getScriptTimeZone(),
                    "yyyy-MM-dd"
                  ),
      kelas     : kelas,
      jam_ke    : row[3] || '',
      pertemuan : row[4] || '',
      materi    : row[5] || '',
      asesmen   : row[7] || '-',
      foto      : fotoHtml,
      refleksi  : row[15] || '-',
      absensi   : absHtml,
      edit_count: row[11] || 0,
      locked    : false
    });
  }

  return result.reverse();
}

function getRekapAbsensi(dari, sampai){ 
  const auth = getAuth(); 
  const jurnal = sheet('JURNAL').getDataRange().getValues(); 
  const absensi = sheet('ABSENSI').getDataRange().getValues(); 
  const siswa = sheet('SISWA').getDataRange().getValues(); 
  const siswaMap = {}; 
  for(let i=1;i<siswa.length;i++){ 
    const nis = String(siswa[i][2] || '').trim(); 
    if(nis){ siswaMap[nis] = siswa[i]; } 
  } 
  const d1 = new Date(dari + 'T00:00:00'); 
  const d2 = new Date(sampai + 'T23:59:59'); 
  let hasil = []; 
  let no = 1; 
  for(let i=1;i<jurnal.length;i++){ 
    if(auth.role !== 'superadmin' && jurnal[i][12] !== auth.email) continue; 
    const tgl = new Date(jurnal[i][1]); 
    if(tgl < d1 || tgl > d2) continue; 
    const jurnalId = jurnal[i][0]; 
    let list = []; 
    for(let a=1;a<absensi.length;a++){ 
      if(absensi[a][0] == jurnalId && absensi[a][3] !== 'H'){ 
        const s = siswaMap[String(absensi[a][2])]; 
        const icon = absensi[a][3] === 'S' ? '\uD83D\uDFE1' : absensi[a][3] === 'I' ? '\uD83D\uDD35' : '\uD83D\uDD34'; 
        const nama = s ? s[3] : getNamaSiswaByNis_(absensi[a][2]);
        list.push(icon + ' ' + nama + ' (' + absensi[a][3] + ')'); 
      } 
    } 
    hasil.push({ 
      no: no++, kelas: jurnal[i][2], 
      tanggal: Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd'), 
      pertemuan: jurnal[i][4], asesmen: jurnal[i][7] || '-', 
      keterangan: list.length ? list.join('<br>') : '\u2705 NIHIL' 
    }); 
  } 
  hasil.sort((a,b)=>a.tanggal.localeCompare(b.tanggal)); 
  return hasil; 
} 

/**
 * Laporan absensi per siswa dalam 1 kelas — bisa per bulan / semester / tahun pelajaran
 * mode: 'bulan'    → period = 'YYYY-MM'       (e.g. '2025-03')
 * mode: 'semester' → period = 'ganjil'|'genap'
 * mode: 'tahun'    → period = ''  (gunakan tahun pelajaran dari setting)
 */
function getLaporanAbsensiSiswa(mode, kelas, period) {
  const auth    = getAuth();
  const setting = getSetting();
  const tz      = Session.getScriptTimeZone();
  const ss      = getSpreadsheet_();

  const shJurnal = ss.getSheetByName('JURNAL');
  const shAbs    = ss.getSheetByName('ABSENSI');
  const shSiswa  = ss.getSheetByName('SISWA');
  if (!shJurnal || !shAbs || !shSiswa) return [];

  const jurnalData = shJurnal.getDataRange().getValues().slice(1);
  const absData    = shAbs.getDataRange().getValues().slice(1);
  const siswaData  = shSiswa.getDataRange().getValues().slice(1);

  // Filter jurnal by email, kelas, and period
  const activeTahun = setting.tahun_pelajaran || '';

  // Filter siswa by kelas, owner email & tahun ajaran aktif — tanpa filter
  // tahun, nama kelas yang dipakai ulang tiap tahun (mis. "8A") akan
  // menggabung siswa dari beberapa angkatan berbeda ke satu laporan.
  let siswaList = siswaData.filter(s =>
    s[0] == kelas && String(s[5] || '').toLowerCase().trim() === auth.email
    && _siswaRowMatchesPeriode_(s, activeTahun)
  );
  if (!siswaList.length) {
    const fromRiwayat = getSiswaAktifByKelasForUser_(kelas, auth.email);
    siswaList = fromRiwayat.map(function(s){
      return ['', '', s.nis, s.nama, s.jk, auth.email];
    });
  }
  if (!siswaList.length) return [];
  const filteredJurnal = jurnalData.filter(r => {
    if (String(r[12] || '').toLowerCase().trim() !== auth.email) return false;
    if (r[2] != kelas) return false;
    const tgl = new Date(r[1]);
    if (mode === 'bulan') {
      const ym = Utilities.formatDate(tgl, tz, 'yyyy-MM');
      return ym === period;
    } else if (mode === 'semester') {
      const sem = String(r[13] || '').toLowerCase().trim();
      const p   = String(period || '').toLowerCase().trim();
      return sem === p ||
        (p === 'ganjil' && (sem === 'i'  || sem === '1')) ||
        (p === 'genap'  && (sem === 'ii' || sem === '2'));
    } else { // tahun
      if (!activeTahun) return true;
      return !r[18] || String(r[18]) === activeTahun;
    }
  });

  const totalPertemuan = filteredJurnal.length;
  const jurnalIds = new Set(filteredJurnal.map(r => String(r[0])));

  // Count per siswa
  const result = siswaList.map(s => {
    const nis  = String(s[2] || '').trim();
    const nama = s[3] || '-';
    const jk   = s[4] || '-';

    let hadir = 0, sakit = 0, izin = 0, alpha = 0;

    absData.forEach(a => {
      if (!jurnalIds.has(String(a[0]))) return;
      if (String(a[2] || '').trim() !== nis) return;
      const st = a[3];
      if (st === 'H') hadir++;
      else if (st === 'S') sakit++;
      else if (st === 'I') izin++;
      else if (st === 'A') alpha++;
    });

    // Siswa yang tidak ada di absensi dianggap hadir
    const totalCatat = hadir + sakit + izin + alpha;
    if (totalCatat === 0 && totalPertemuan > 0) hadir = totalPertemuan;

    const persen = totalPertemuan > 0
      ? Math.round((hadir / totalPertemuan) * 100)
      : 0;

    return { nis, nama, jk, hadir, sakit, izin, alpha, totalPertemuan, persen };
  });

  return result.sort((a,b) => a.nama.localeCompare(b.nama));
}

function cekJurnalSemester(){
  const auth        = getAuth();
  const setting     = getSetting();
  const activeTahun = setting.tahun_pelajaran || '';
  const jurnal = sheet('JURNAL').getDataRange().getValues().slice(1);
  let adaGanjil = false;
  let adaGenap  = false;
  jurnal.forEach(r => {
    if(r[12] !== auth.email) return;
    if(activeTahun && r[18] && String(r[18]) !== activeTahun) return;
    const sem = String(r[13] || '').toLowerCase().trim();
    if(sem === 'ganjil' || sem === 'i'  || sem === '1') adaGanjil = true;
    if(sem === 'genap'  || sem === 'ii' || sem === '2') adaGenap  = true;
  });
  return { adaGanjil, adaGenap };
}

