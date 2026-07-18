/**
 * Jurnal.js — Manajemen Jurnal Pembelajaran
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

function getKelas(){
  ensureAcademicSchema_();
  const auth = getAuth();

  if(!auth.email) return [];

  const kelasAktif = getKelasDiampuAktifForUser_(auth.email);
  if (kelasAktif && kelasAktif.length) return kelasAktif;

  if(!sheet('SISWA')) return [];

  // Cache SISWA sheet (TTL 120s) — invalidated on upload siswa
  const data = getSheetCached('SISWA', 120).slice(1);

  const kelas = data
    .filter(r => String(r[5]).toLowerCase().trim() === auth.email)
    .map(r => r[0])
    .filter(Boolean);

  return [...new Set(kelas)].sort();
}

function getSiswaByKelas(kelas){
  ensureAcademicSchema_();
  const auth = getAuth();
  if(!auth.email) return [];

  const fromRiwayat = getSiswaAktifByKelasForUser_(kelas, auth.email);
  if (fromRiwayat && fromRiwayat.length) return fromRiwayat;

  if(!sheet('SISWA')) return [];

  // Cache SISWA sheet (TTL 120s)
  const data = getSheetCached('SISWA', 120).slice(1);

  return data
    .filter(r =>
      r[0] === kelas &&
      String(r[5]).toLowerCase().trim() === auth.email
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
  if(!sh) throw new Error('Sheet SISWA tidak ditemukan');

  const setting = getSetting();
  const tahunAktif = setting.tahun_pelajaran || getLegacyDefaultYear_();
  const semesterAktif = setting.semester || 'Ganjil';

  // Hitung jumlah kelas & siswa milik guru ini sebelum import
  const existingData = sh.getDataRange().getValues().slice(1)
    .filter(r => String(r[5] || '').toLowerCase().trim() === auth.email);
  const existingKelas = new Set(existingData.map(r => String(r[0] || '').trim()).filter(Boolean));
  const existingSiswa = existingData.length;

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

  // Cek batas tier LITE per kelas baru
  const newKelas = new Set(normalized.map(r => String(r.kelas || '').trim()).filter(Boolean));
  newKelas.forEach(function(k) {
    if (!existingKelas.has(k)) {
      assertLiteKelasLimit_(existingKelas.size);
      existingKelas.add(k); // prevent re-checking same new kelas
    }
  });
  // Cek batas siswa
  assertLiteSiswaLimit_(existingSiswa + normalized.length - 1);

  const rowsForMaster = [];
  normalized.forEach(r=>{
    sh.appendRow([
      r.kelas,
      r.no_absen,
      r.nis,
      r.nama,
      r.jk,
      auth.email
    ]);

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
  });

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

  if(data.dokumentasi && data.dokumentasi.length){

    const folder = getGuruFolder_(email);

    data.dokumentasi.forEach(f=>{

      const base64 = f.data.split(',')[1];

      const blob = Utilities.newBlob(
        Utilities.base64Decode(base64),
        f.type || 'image/jpeg',
        jurnalId + '_' + f.name
      );

      const file = folder.createFile(blob);
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );

      const url = "https://drive.google.com/uc?id=" + file.getId();

      fotoArr.push({ full: url });
      fotoUrls.push(url);
    });
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

  trySyncGuruSummaryAfterMutation_(email, 'SIMPAN_JURNAL');

  return { status:true };
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
  sh.getRange(rowNumber, 6).setValue(data.materi);
  sh.getRange(rowNumber, 7).setValue(data.tujuan);
  sh.getRange(rowNumber, 8).setValue(data.asesmen);
  sh.getRange(rowNumber, 16).setValue(data.refleksi);

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

  let finalFotos = [];

  if(data.keepFotos && Array.isArray(data.keepFotos)){
    finalFotos = data.keepFotos;
  }

  const folder = getGuruFolder_(getAuth().email);

  if(data.replaceFotos && Array.isArray(data.replaceFotos)){

    data.replaceFotos.forEach(r => {

      if(!r.index && r.index !== 0) return;
      if(!r.file || !r.file.data) return;

      const base64 = r.file.data.split(',')[1];
      const blob = compressImage_(base64, r.file.type);
      const file = folder.createFile(blob);

      finalFotos[r.index] = {
        full: "https://drive.google.com/uc?id=" + file.getId(),
        thumb: "https://drive.google.com/uc?id=" + file.getId()
      };

    });
  }

  if(data.newFotos && Array.isArray(data.newFotos)){

    data.newFotos.forEach(f => {

      if(!f.data) return;

      const base64 = f.data.split(',')[1];
      const blob = compressImage_(base64, f.type);
      const file = folder.createFile(blob);

      finalFotos.push({
        full: "https://drive.google.com/uc?id=" + file.getId(),
        thumb: "https://drive.google.com/uc?id=" + file.getId()
      });

    });
  }

  sh.getRange(rowNumber, 15)
    .setValue(JSON.stringify(finalFotos));

  const shAbs = ss.getSheetByName('ABSENSI');
  const absData = shAbs.getDataRange().getValues();

  for(let i = absData.length - 1; i > 0; i--){
    if(absData[i][0] == data.jurnalId){
      shAbs.deleteRow(i + 1);
    }
  }

  if(data.absensi && data.absensi.length){
    data.absensi.forEach(a => {
      shAbs.appendRow([
        data.jurnalId,
        data.kelas,
        a.nis,
        a.status
      ]);
    });
  }

  return { status:true };
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

    // foto
    let fotoHtml = '-';
    try {
      let fotoLinks = [];
      if (row[14]) {
        const parsed = JSON.parse(row[14]);
        if (Array.isArray(parsed)) {
          parsed.forEach((f, idx) => {
            if (!f || !f.full) return;
            const m = f.full.match(/[-\w]{25,}/);
            if (!m) return;
            fotoLinks.push(`<a href="https://drive.google.com/file/d/${m[0]}/view" target="_blank" style="font-weight:600;color:#4f46e5;text-decoration:none;">📷 Foto ${idx+1}</a>`);
          });
        }
      }
      if (fotoLinks.length === 0 && row[17]) {
        String(row[17]).trim().split(',').filter(u=>u.trim()).forEach((u, idx) => {
          const m = u.match(/[-\w]{25,}/);
          if (!m) return;
          fotoLinks.push(`<a href="https://drive.google.com/file/d/${m[0]}/view" target="_blank" style="font-weight:600;color:#4f46e5;text-decoration:none;">📷 Foto ${idx+1}</a>`);
        });
      }
      if (fotoLinks.length) fotoHtml = fotoLinks.join('<br>');
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

  // Filter siswa by kelas & owner email
  let siswaList = siswaData.filter(s =>
    s[0] == kelas && String(s[5] || '').toLowerCase().trim() === auth.email
  );
  if (!siswaList.length) {
    const fromRiwayat = getSiswaAktifByKelasForUser_(kelas, auth.email);
    siswaList = fromRiwayat.map(function(s){
      return ['', '', s.nis, s.nama, s.jk, auth.email];
    });
  }
  if (!siswaList.length) return [];

  // Filter jurnal by email, kelas, and period
  const activeTahun = setting.tahun_pelajaran || '';
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

