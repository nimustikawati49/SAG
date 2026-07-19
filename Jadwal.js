/**
 * Jadwal.js — Jadwal Semester Guru
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

/**
 * Cari sheet jadwal dengan case-insensitive.
 * Coba lowercase ('jadwal_semester') dulu, lalu uppercase ('JADWAL_SEMESTER').
 * Kembalikan null jika tidak ada → TIDAK membuat sheet baru.
 */
function getJadwalSheet_(){
  const ss = getSpreadsheet_();
  return ss.getSheetByName('jadwal_semester')
    || ss.getSheetByName('JADWAL_SEMESTER')
    || null;
}

/**
 * Pastikan sheet jadwal ada. Buat baru jika belum ada.
 * Hanya dipanggil dari operasi TULIS (save/update/delete).
 */
function ensureJadwalSheet_(){
  let sh = getJadwalSheet_();
  if(!sh){
    const ss = getSpreadsheet_();
    sh = ss.insertSheet('JADWAL_SEMESTER');
    sh.appendRow([
      'email',
      'semester',
      'hari',
      'kelas',
      'mapel',
      'jam_mulai',
      'jam_selesai',
      'created_at'
    ]);
  }
  return sh;
}

function normalizeJam_(val){

  if(val === null || val === undefined || val === '') return '';

  // Google Sheets time stored as Date object — use Utilities.formatDate for timezone safety
  if(Object.prototype.toString.call(val) === '[object Date]'){
    try{
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
    }catch(e){
      // fallback: getHours in local timezone
      const h = val.getHours().toString().padStart(2,'0');
      const m = val.getMinutes().toString().padStart(2,'0');
      return h + ':' + m;
    }
  }

  // Google Sheets time stored as decimal fraction (0.333... = 8:00)
  if(typeof val === 'number'){
    const totalMin = Math.round(val * 24 * 60);
    const h = Math.floor(totalMin / 60) % 24;
    const m = totalMin % 60;
    return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  }

  if(typeof val === 'string'){
    // ISO datetime from JSON cache — extract LOCAL hour from offset
    // e.g. "1899-12-30T00:00:00.000Z" for 08:00 UTC+8 is WRONG, skip ISO for Date case
    // Only handle ISO if it has clear local indicator
    const isoMatch = val.match(/T(\d{2}):(\d{2})/);
    if(isoMatch){
      // Parse as Date and re-format with timezone
      try{
        const d = new Date(val);
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
      }catch(e){
        return isoMatch[1] + ':' + isoMatch[2];
      }
    }

    // AM/PM format
    if(val.includes('AM') || val.includes('PM')){
      const d = new Date('1970-01-01 ' + val);
      if(!isNaN(d.getTime())){
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
      }
    }

    // Already HH:MM format
    const hmMatch = val.match(/^(\d{1,2}):(\d{2})/);
    if(hmMatch){
      return String(Number(hmMatch[1])).padStart(2,'0') + ':' + hmMatch[2];
    }

    return '';
  }

  return '';
}

function toMinutes_(val){

  const t = normalizeJam_(val);
  if(!t) return 0;

  const parts = t.split(':');
  return (Number(parts[0]) * 60) + Number(parts[1]);
}

/**
 * ensureJadwalSchemaAcademic_()
 * JADWAL_SEMESTER awalnya hanya dikunci per (email, semester) — TANPA
 * tahun_pelajaran. Akibatnya jadwal hari/jam semester "Ganjil" tahun lalu
 * bisa nyangkut/campur dengan "Ganjil" tahun ajaran baru. Tambahkan kolom
 * tahun_pelajaran di akhir (index tetap, tidak mengubah kolom lama yang
 * sudah dipakai dengan index tetap di seluruh file ini), lalu backfill
 * baris lama dengan tahun default supaya tidak silang sama sekali.
 * Return: index kolom (0-based) tahun_pelajaran.
 */
function ensureJadwalSchemaAcademic_(){
  const sh = ensureJadwalSheet_();
  const lastCol = sh.getLastColumn();
  const header = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||'').toLowerCase().trim(); })
    : [];
  let idx = header.indexOf('tahun_pelajaran');
  if (idx === -1) {
    idx = lastCol; // 0-based index kolom baru
    sh.getRange(1, lastCol + 1).setValue('tahun_pelajaran');
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const defaultTahun = (typeof getLegacyDefaultYear_ === 'function') ? getLegacyDefaultYear_() : '';
      if (defaultTahun) {
        const fillRange = sh.getRange(2, idx + 1, lastRow - 1, 1);
        const vals = [];
        for (let r = 0; r < lastRow - 1; r++) vals.push([defaultTahun]);
        fillRange.setValues(vals);
      }
    }
  }
  return idx;
}

/**
 * getJadwalSemesterForPeriod_(email, tahun, semester)
 * Ambil entri jadwal (hari/kelas/mapel/jam) milik satu guru untuk satu
 * periode (tahun_pelajaran + semester) spesifik. Baris lama tanpa
 * tahun_pelajaran (legacy, sebelum migrasi) dianggap cocok ke periode manapun.
 */
function getJadwalSemesterForPeriod_(email, tahun, semester){
  const targetEmail = String(email || '').toLowerCase().trim();
  const targetTahun = String(tahun || '').trim();
  const targetSem = String(semester || '').trim();

  const sh = getJadwalSheet_();
  if (!sh) return [];
  const tIdx = ensureJadwalSchemaAcademic_();
  const values = sh.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').toLowerCase().trim() !== targetEmail) continue;
    if (String(values[i][1] || '').trim() !== targetSem) continue;
    const rowTahun = String(values[i][tIdx] || '').trim();
    if (rowTahun && rowTahun !== targetTahun) continue;
    result.push({
      hari: values[i][2],
      kelas: values[i][3],
      mapel: values[i][4],
      jam_mulai: normalizeJam_(values[i][5]),
      jam_selesai: normalizeJam_(values[i][6])
    });
  }
  return result;
}

/**
 * getJadwalSemesterCandidatePeriods_(email)
 * Daftar periode (tahun+semester) unik yang punya entri jadwal milik guru
 * ini — dipakai untuk mendeteksi "periode terakhir yang punya jadwal"
 * saat guru pindah tahun/semester.
 */
function getJadwalSemesterCandidatePeriods_(email){
  const targetEmail = String(email || '').toLowerCase().trim();
  const sh = getJadwalSheet_();
  if (!sh) return [];
  const tIdx = ensureJadwalSchemaAcademic_();
  const values = sh.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').toLowerCase().trim() !== targetEmail) continue;
    const semester = String(values[i][1] || '').trim();
    const tahun = String(values[i][tIdx] || '').trim();
    if (!semester || !tahun) continue;
    result.push({
      tahun_pelajaran: tahun,
      semester: semester,
      created_at: values[i][7] ? new Date(values[i][7]) : null
    });
  }
  return result;
}

/**
 * continueJadwalSemesterToPeriod_(email, fromTahun, fromSemester, toTahun, toSemester)
 * Salin entri jadwal (hari/jam) milik satu guru dari periode sumber ke tujuan.
 * Return: jumlah baris yang disalin.
 */
function continueJadwalSemesterToPeriod_(email, fromTahun, fromSemester, toTahun, toSemester){
  const targetEmail = String(email || '').toLowerCase().trim();
  const source = getJadwalSemesterForPeriod_(targetEmail, fromTahun, fromSemester);
  if (!source.length) return 0;

  // Hapus dulu entri guru ini di periode tujuan supaya idempoten (hindari duplikat).
  resetJadwalSemesterForPeriod_(targetEmail, toTahun, toSemester);

  const sh = ensureJadwalSheet_();
  const tIdx = ensureJadwalSchemaAcademic_();
  const now = new Date();
  source.forEach(function(j){
    sh.appendRow([targetEmail, toSemester, j.hari, j.kelas, j.mapel, j.jam_mulai, j.jam_selesai, now]);
    sh.getRange(sh.getLastRow(), tIdx + 1).setValue(toTahun);
  });
  return source.length;
}

/**
 * resetJadwalSemesterForPeriod_(email, tahun, semester)
 * Hapus semua entri jadwal (hari/jam) milik satu guru untuk satu periode.
 * Return: jumlah baris yang dihapus.
 */
function resetJadwalSemesterForPeriod_(email, tahun, semester){
  const targetEmail = String(email || '').toLowerCase().trim();
  const targetTahun = String(tahun || '').trim();
  const targetSem = String(semester || '').trim();

  const sh = getJadwalSheet_();
  if (!sh) return 0;
  const tIdx = ensureJadwalSchemaAcademic_();
  const values = sh.getDataRange().getValues();
  let deleted = 0;
  for (let i = values.length; i >= 2; i--) {
    const r = values[i - 1];
    if (String(r[0] || '').toLowerCase().trim() !== targetEmail) continue;
    if (String(r[1] || '').trim() !== targetSem) continue;
    const rowTahun = String(r[tIdx] || '').trim();
    if (rowTahun && rowTahun !== targetTahun) continue;
    sh.deleteRow(i);
    deleted++;
  }
  return deleted;
}

function isJamBentrok_(data, existing){

  const startNew = toMinutes_(data.jam_mulai);
  const endNew   = toMinutes_(data.jam_selesai);

  const startOld = toMinutes_(existing.jam_mulai);
  const endOld   = toMinutes_(existing.jam_selesai);

  return (
    startNew < endOld &&
    endNew > startOld
  );
}

function saveJadwalSemester(data){

  assertLicenseActive();

  const auth = getAuth();
  if(auth.role !== 'admin' && auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }

  // ── Sanitize & validate server-side ──
  var VALID_HARI = ['SENIN','SELASA','RABU','KAMIS','JUMAT','SABTU'];
  var JAM_RE     = /^\d{2}:\d{2}$/;
  data.hari       = String(data.hari       || '').trim().toUpperCase();
  data.kelas      = String(data.kelas      || '').trim().substring(0, 50);
  data.mapel      = String(data.mapel      || '').trim().substring(0, 100);
  data.jam_mulai  = String(data.jam_mulai  || '').trim();
  data.jam_selesai= String(data.jam_selesai|| '').trim();

  if(VALID_HARI.indexOf(data.hari) === -1){
    throw new Error('Hari tidak valid: ' + data.hari);
  }
  if(!data.kelas){
    throw new Error('Kelas wajib diisi');
  }
  if(!JAM_RE.test(data.jam_mulai)){
    throw new Error('Format jam mulai tidak valid (gunakan HH:MM)');
  }
  if(!JAM_RE.test(data.jam_selesai)){
    throw new Error('Format jam selesai tidak valid (gunakan HH:MM)');
  }

  const sh = ensureJadwalSheet_();
  const tIdx = ensureJadwalSchemaAcademic_();
  const values = sh.getDataRange().getValues();

  const setting = getSetting();
  const semesterAktif = setting.semester || '';
  const tahunAktif = setting.tahun_pelajaran || '';
  const email = auth.email;

  if(!data.hari || !data.kelas || !data.jam_mulai || !data.jam_selesai){
    throw new Error('Data jadwal belum lengkap');
  }

  if(data.jam_mulai >= data.jam_selesai){
    throw new Error('Jam selesai harus lebih besar dari jam mulai');
  }

  for(let i=1;i<values.length;i++){

    const rowEmail = String(values[i][0]).toLowerCase();
    const rowSemester = values[i][1];
    const rowHari = values[i][2];
    const rowTahun = String(values[i][tIdx] || '').trim();

    if(rowEmail !== email) continue;
    if(rowSemester !== semesterAktif) continue;
    if(rowTahun && tahunAktif && rowTahun !== tahunAktif) continue;
    if(rowHari !== data.hari) continue;

    const existing = {
      jam_mulai: values[i][5],
      jam_selesai: values[i][6]
    };

    if(isJamBentrok_(data, existing)){
      throw new Error(
        'Jam bentrok dengan jadwal lain pada hari ' + data.hari
      );
    }
  }
  if(!isSettingLocked()){
    throw new Error('Setting harus disimpan & terkunci sebelum membuat jadwal');
  }

  sh.appendRow([
    email,
    semesterAktif,
    data.hari,
    data.kelas,
    data.mapel || setting.mata_pelajaran || '',
    data.jam_mulai,
    data.jam_selesai,
    new Date()
  ]);
  sh.getRange(sh.getLastRow(), tIdx + 1).setValue(tahunAktif);

  logAudit(
    'ADD_JADWAL',
    email,
    data.hari + ' ' + data.kelas + ' ' + data.jam_mulai
  );
  invalidateCache_('JADWAL_SEMESTER');
  invalidateDashboardCache_();
  return true;
}

function saveBulkJadwalSemester(list){

  assertLicenseActive();

  const auth = getAuth();
  if(auth.role !== 'admin' && auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }

  if(!Array.isArray(list) || list.length === 0){
    throw new Error('Data jadwal kosong');
  }

  if(!isSettingLocked()){
    throw new Error('Setting harus disimpan & terkunci sebelum membuat jadwal');
  }

  // ── Sanitize & validate setiap item bulk ──
  var VALID_HARI = ['SENIN','SELASA','RABU','KAMIS','JUMAT','SABTU'];
  var JAM_RE = /^\d{2}:\d{2}$/;
  list = list.map(function(item) {
    return {
      hari      : String(item.hari       || '').trim().toUpperCase(),
      kelas     : String(item.kelas      || '').trim().substring(0, 50),
      mapel     : String(item.mapel      || '').trim().substring(0, 100),
      jam_mulai : String(item.jam_mulai  || '').trim(),
      jam_selesai: String(item.jam_selesai|| '').trim()
    };
  });
  list.forEach(function(item, idx) {
    if(VALID_HARI.indexOf(item.hari) === -1)
      throw new Error('Baris ' + (idx+1) + ': Hari tidak valid – ' + item.hari);
    if(!item.kelas)
      throw new Error('Baris ' + (idx+1) + ': Kelas wajib diisi');
    if(!JAM_RE.test(item.jam_mulai))
      throw new Error('Baris ' + (idx+1) + ': Format jam mulai tidak valid (HH:MM)');
    if(!JAM_RE.test(item.jam_selesai))
      throw new Error('Baris ' + (idx+1) + ': Format jam selesai tidak valid (HH:MM)');
  });

  const sh = ensureJadwalSheet_();
  const tIdx = ensureJadwalSchemaAcademic_();
  const values = sh.getDataRange().getValues();
  const settingBulk = getSetting();
  const semesterAktif = settingBulk.semester || '';
  const tahunAktif = settingBulk.tahun_pelajaran || '';
  const email = auth.email;

  for(let i=0;i<list.length;i++){
    for(let j=i+1;j<list.length;j++){

      if(list[i].hari !== list[j].hari) continue;

      const s1 = toMinutes_(list[i].jam_mulai);
      const e1 = toMinutes_(list[i].jam_selesai);
      const s2 = toMinutes_(list[j].jam_mulai);
      const e2 = toMinutes_(list[j].jam_selesai);

      if(s1 < e2 && e1 > s2){
        throw new Error('Bentrok internal pada hari ' + list[i].hari);
      }
    }
  }

  for(let i=1;i<values.length;i++){

    if(String(values[i][0]).toLowerCase() !== email) continue;
    if(values[i][1] !== semesterAktif) continue;
    const rowTahunBulk = String(values[i][tIdx] || '').trim();
    if(rowTahunBulk && tahunAktif && rowTahunBulk !== tahunAktif) continue;

    for(let item of list){

      if(values[i][2] !== item.hari) continue;

      const startOld = toMinutes_(values[i][5]);
      const endOld   = toMinutes_(values[i][6]);
      const startNew = toMinutes_(item.jam_mulai);
      const endNew   = toMinutes_(item.jam_selesai);

      if(startNew < endOld && endNew > startOld){
        throw new Error('Bentrok dengan jadwal lama hari ' + item.hari);
      }
    }
  }

  list.forEach(item=>{
    sh.appendRow([
      email,
      semesterAktif,
      item.hari,
      item.kelas,
      item.mapel || settingBulk.mata_pelajaran || '',
      normalizeJam_(item.jam_mulai),
      normalizeJam_(item.jam_selesai),
      new Date()
    ]);
    sh.getRange(sh.getLastRow(), tIdx + 1).setValue(tahunAktif);
  });

  logAudit('ADD_BULK_JADWAL', email, 'Total: ' + list.length);
  invalidateCache_('JADWAL_SEMESTER');
  invalidateDashboardCache_();
  return true;
}

function getJadwalMengajar(){

  const auth = getAuth();
  if(!auth || !auth.email) return [];

  const emailLogin = String(auth.email).toLowerCase().trim();

  const setting = getSetting();
  const semesterAktif = String(setting.semester || '').trim();
  const tahunAktif = String(setting.tahun_pelajaran || '').trim();

  const sh = getJadwalSheet_();
  if(!sh) return [];
  const tIdx = ensureJadwalSchemaAcademic_();
  const values = sh.getDataRange().getValues();

  let result = [];

  for(let i=1;i<values.length;i++){

    const emailRow = String(values[i][0] || '').toLowerCase().trim();
    const semesterRow = String(values[i][1] || '').trim();
    const tahunRow = String(values[i][tIdx] || '').trim();

    if(emailRow !== emailLogin) continue;
    if(semesterRow !== semesterAktif) continue;
    if(tahunRow && tahunAktif && tahunRow !== tahunAktif) continue;

    result.push({
      hari: values[i][2],
      kelas: values[i][3],
      mapel: values[i][4],
      jam_mulai: formatJam(values[i][5]),
      jam_selesai: formatJam(values[i][6])
    });
  }

  return result;
}

function deleteJadwalSemester(id){

  assertLicenseActive();

  const auth = getAuth();
  const sh = ensureJadwalSheet_();

  const row = Number(id);

  if(row < 2){
    throw new Error('ID tidak valid');
  }

  const rowEmail = sh.getRange(row,1).getValue();

  if(rowEmail !== auth.email && auth.role !== 'superadmin'){
    throw new Error('AKSES_DITOLAK');
  }

  sh.deleteRow(row);

  logAudit(
    'DELETE_JADWAL',
    auth.email,
    'Row ' + row
  );
  invalidateCache_('JADWAL_SEMESTER');
  invalidateDashboardCache_();

  return true;
}

function getDashboardJadwal(){
  try {
    const auth = getAuth();
    if(!auth || !auth.email) return { _error: 'NOT_AUTH' };

    // READ ONLY — gunakan getJadwalSheet_() agar tidak membuat sheet baru
    const ss = getSpreadsheet_();
    if(!ss) return { _error: 'NO_SPREADSHEET' };

    // Cari sheet dengan berbagai nama yang mungkin
    const sh = ss.getSheetByName('JADWAL_SEMESTER')
      || ss.getSheetByName('jadwal_semester')
      || ss.getSheetByName('Jadwal_Semester')
      || null;

    if(!sh) return { _error: 'NO_SHEET' };

    const tIdx = (typeof ensureJadwalSchemaAcademic_ === 'function') ? ensureJadwalSchemaAcademic_() : -1;
    const values = sh.getDataRange().getValues();
    if(values.length < 2) return {};

    const email = String(auth.email).toLowerCase().trim();

    var semesterAktif = '';
    var tahunAktif = '';
    try {
      const setting = getSetting();
      semesterAktif = String(setting.semester || '').trim().toLowerCase();
      tahunAktif = String(setting.tahun_pelajaran || '').trim();
    } catch(e) {
      semesterAktif = '';
      tahunAktif = '';
    }

    let grouped = {};

    for(let i = 1; i < values.length; i++){
      const rowEmail = String(values[i][0] || '').toLowerCase().trim();
      if(rowEmail !== email) continue;

      if(semesterAktif){
        const rowSemester = String(values[i][1] || '').trim().toLowerCase();
        if(rowSemester !== semesterAktif) continue;
      }

      // Baris lama (sebelum migrasi tahun_pelajaran) tidak punya nilai di
      // kolom ini — dianggap cocok ke periode manapun. Baris baru yang
      // sudah punya tahun_pelajaran WAJIB cocok dengan periode aktif,
      // supaya jadwal tahun ajaran lama tidak nyangkut ke tahun baru.
      if(tIdx > -1 && tahunAktif){
        const rowTahun = String(values[i][tIdx] || '').trim();
        if(rowTahun && rowTahun !== tahunAktif) continue;
      }

      const hari = String(values[i][2] || '').trim().toUpperCase();
      if(!hari) continue;
      if(!grouped[hari]) grouped[hari] = [];

      grouped[hari].push({
        kelas:       String(values[i][3] || ''),
        mapel:       String(values[i][4] || ''),
        jam_mulai:   normalizeJam_(values[i][5]) || '--:--',
        jam_selesai: normalizeJam_(values[i][6]) || '--:--'
      });
    }

    return grouped;

  } catch(e) {
    return { _error: e.message || 'UNKNOWN_ERROR' };
  }
}

function previewBentrokJadwal(data){

  const auth = getAuth();
  if(!auth.email) return { bentrok:false };

  const sh = getJadwalSheet_();
  if(!sh) return { bentrok:false };

  const tIdxPreview = ensureJadwalSchemaAcademic_();
  const settingPreview = getSetting();
  const semesterAktif = settingPreview.semester || '';
  const tahunAktifPreview = settingPreview.tahun_pelajaran || '';
  const values = sh.getDataRange().getValues();

  const startNew = toMinutes_(data.jam_mulai);
  const endNew   = toMinutes_(data.jam_selesai);

  for(let i=1;i<values.length;i++){

    if(String(values[i][0]).toLowerCase().trim() !== auth.email) continue;
    if(values[i][1] !== semesterAktif) continue;
    const rowTahunPreview = String(values[i][tIdxPreview] || '').trim();
    if(rowTahunPreview && tahunAktifPreview && rowTahunPreview !== tahunAktifPreview) continue;
    if(values[i][2] !== data.hari) continue;

    const startOld = toMinutes_(values[i][5]);
    const endOld   = toMinutes_(values[i][6]);

    if(startNew < endOld && endNew > startOld){
      return {
        bentrok:true,
        konflik:{
          hari: values[i][2],
          kelas: values[i][3],
          jam: normalizeJam_(values[i][5]) + ' - ' + normalizeJam_(values[i][6])
        }
      };
    }
  }

  return { bentrok:false };
}

function getJadwalForSetting(){

  const auth = getAuth();
  if(!auth || !auth.email) return [];

  const emailLogin = String(auth.email).toLowerCase().trim();
  const settingForList = getSetting();
  const semesterAktif = String(settingForList.semester || '').trim();
  const tahunAktif = String(settingForList.tahun_pelajaran || '').trim();

  const sh = getJadwalSheet_();
  if(!sh) return [];
  const tIdx = ensureJadwalSchemaAcademic_();
  const values = sh.getDataRange().getValues();

  let result = [];

  for(let i=1;i<values.length;i++){

    const emailRow = String(values[i][0] || '').toLowerCase().trim();
    const semesterRow = String(values[i][1] || '').trim();
    const tahunRow = String(values[i][tIdx] || '').trim();

    if(emailRow !== emailLogin) continue;
    if(semesterRow !== semesterAktif) continue;
    if(tahunRow && tahunAktif && tahunRow !== tahunAktif) continue;

    result.push({
      row: i+1,
      hari: values[i][2],
      kelas: values[i][3],
      mapel: values[i][4],
      jam_mulai: formatJam(values[i][5]),
      jam_selesai: formatJam(values[i][6])
    });
  }

  return result;
}

function formatJam(value){

  if(value instanceof Date){
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }

  // Google Sheets time as decimal fraction (0.333... = 8:00)
  if(typeof value === 'number' && value >= 0 && value < 1){
    const totalMin = Math.round(value * 24 * 60);
    const h = Math.floor(totalMin / 60) % 24;
    const m = totalMin % 60;
    return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  }

  // HH:MM string — return as-is (trim to 5 chars)
  const s = String(value || '');
  if(s.match(/^\d{1,2}:\d{2}/)) return s.substring(0, 5);

  return s;
}

function updateJadwalSemester(row,data){

  assertLicenseActive();

  const auth = getAuth();
  const sh = ensureJadwalSheet_();
  const tIdxUpd = ensureJadwalSchemaAcademic_();
  const settingUpd = getSetting();
  const semesterAktif = settingUpd.semester || '';
  const tahunAktifUpd = settingUpd.tahun_pelajaran || '';

  if(sh.getRange(row,1).getValue() !== auth.email){
    throw new Error('AKSES_DITOLAK');
  }

  const values = sh.getDataRange().getValues();

  const startNew = toMinutes_(data.jam_mulai);
  const endNew   = toMinutes_(data.jam_selesai);

  for(let i=1;i<values.length;i++){

    if(i+1 === row) continue;
    if(String(values[i][0]).toLowerCase() !== auth.email) continue;
    if(values[i][1] !== semesterAktif) continue;
    const rowTahunUpd = String(values[i][tIdxUpd] || '').trim();
    if(rowTahunUpd && tahunAktifUpd && rowTahunUpd !== tahunAktifUpd) continue;
    if(values[i][2] !== data.hari) continue;

    const startOld = toMinutes_(values[i][5]);
    const endOld   = toMinutes_(values[i][6]);

    if(startNew < endOld && endNew > startOld){
      throw new Error('Jam bentrok dengan jadwal lain');
    }
  }

  sh.getRange(row,3).setValue(data.hari);
  sh.getRange(row,4).setValue(data.kelas);
  sh.getRange(row,5).setValue(data.mapel);
  sh.getRange(row,6).setValue(normalizeJam_(data.jam_mulai));
  sh.getRange(row,7).setValue(normalizeJam_(data.jam_selesai));

  logAudit('UPDATE_JADWAL', auth.email, data.hari);
  invalidateCache_('JADWAL_SEMESTER');
  invalidateDashboardCache_();

  return true;
}
