/**
 * Dashboard.js — Data Dashboard & Meta
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

/**
 * _computeV4FromRows_(jurnal, absen, email, period, dari, sampai)
 * Inti perhitungan V4 (rekap jurnal+absensi per kelas), dipisah dari
 * getDashboardV4Data() supaya baris JURNAL/ABSENSI yang SUDAH dibaca oleh
 * getDashboardAllData() bisa dipakai ulang di sini tanpa fetch sheet lagi.
 */
function _computeV4FromRows_(jurnal, absen, email, period, dari, sampai){

  if(jurnal.length < 2){
    return emptyDashboard();
  }

  const from = dari ? new Date(dari + 'T00:00:00') : null;
  const to   = sampai ? new Date(sampai + 'T23:59:59') : null;

  const absensiMap = {};
  for (let ai = 1; ai < absen.length; ai++) {
    const jurnalId = absen[ai][0];
    const status = String(absen[ai][3] || '').trim();
    if (!jurnalId || !status) continue;
    if (!absensiMap[jurnalId]) {
      absensiMap[jurnalId] = { H: 0, S: 0, I: 0, A: 0, total: 0 };
    }
    if (absensiMap[jurnalId][status] !== undefined) {
      absensiMap[jurnalId][status]++;
      absensiMap[jurnalId].total++;
    }
  }

  let totalJurnal = 0;
  let kelasSet = new Set();
  let kelasData = {};

  for(let i=1;i<jurnal.length;i++){

    const row = jurnal[i];

    if(row[12] !== email) continue;

    const tgl = new Date(row[1]);

    if (period.tahun_pelajaran && row[18] && String(row[18]) !== String(period.tahun_pelajaran)) continue;
    if (period.semester && row[13] && String(row[13]).toLowerCase().trim() !== String(period.semester).toLowerCase().trim()) continue;

    if(from && tgl < from) continue;
    if(to && tgl > to) continue;

    totalJurnal++;
    const kelas = row[2];
    kelasSet.add(kelas);

    if(!kelasData[kelas]){
      kelasData[kelas] = {H:0,S:0,I:0,A:0,total:0};
    }

    const jurnalId = row[0];
    const absSummary = absensiMap[jurnalId];
    if (absSummary) {
      kelasData[kelas].H += absSummary.H;
      kelasData[kelas].S += absSummary.S;
      kelasData[kelas].I += absSummary.I;
      kelasData[kelas].A += absSummary.A;
      kelasData[kelas].total += absSummary.total;
    }
  }

  if(totalJurnal === 0){
    return {
      totalJurnal:0,
      totalKelas:0,
      rata2:0,
      risiko:0,
      kelasData:{},
      ranking:[],
      insight:['Belum ada jurnal dalam rentang tanggal ini.']
    };
  }

  let totalPersen = 0;
  let risiko = 0;
  let ranking = [];

  for(let k in kelasData){
    const d = kelasData[k];
    const persen = d.total > 0
      ? Math.round((d.H / d.total) * 100)
      : 0;

    totalPersen += persen;
    if(persen < 80) risiko++;

    ranking.push({kelas:k,persentase:persen});
  }

  ranking.sort((a,b)=>b.persentase-a.persentase);

  return {
    totalJurnal,
    totalKelas: kelasSet.size,
    rata2: Math.round(totalPersen / ranking.length),
    risiko,
    kelasData,
    ranking,
    insight: [
      "Total jurnal: " + totalJurnal,
      "Rata-rata hadir: " + Math.round(totalPersen / ranking.length) + "%",
      "Kelas risiko: " + risiko
    ]
  };
}

function getDashboardV4Data(dari, sampai){
  ensureAcademicSchema_();
  const auth = getAuth();
  const email = auth.email;
  const period = getUserAcademicPeriod(email);
  const jurnal = sheet('JURNAL').getDataRange().getValues();
  const absen  = sheet('ABSENSI').getDataRange().getValues();
  return _computeV4FromRows_(jurnal, absen, email, period, dari, sampai);
}

function emptyDashboard(){
  return {
    totalJurnal:0,
    totalKelas:0,
    rata2:0,
    risiko:0,
    kelasData:{},
    ranking:[],
    insight:['Belum ada data.']
  };
}

function buildInsightV4(ranking, risiko){

  let arr = [];

  if(ranking.length === 0){
    arr.push("Belum ada jurnal dalam rentang tanggal ini.");
    return arr;
  }

  if(risiko > 0){
    arr.push(risiko + " kelas memiliki rata-rata < 80%.");
  }

  arr.push("Kelas terbaik: " + ranking[0].kelas + 
           " (" + ranking[0].persentase + "%)");

  if(ranking[ranking.length-1].persentase < 80){
    arr.push("Perlu evaluasi kehadiran siswa.");
  }

  return arr;
}

/**
 * _countSiswaPerKelasBatch_(tahun, semester, kelasList)
 * Hitung jumlah siswa aktif per kelas dalam SATU pembacaan RiwayatKelas,
 * bukan satu pembacaan per kelas (dulu dipanggil lewat
 * getSiswaAktifByKelasForUser_ di dalam loop — untuk guru dengan 6 kelas
 * itu = 6 x 3 pembacaan sheet penuh, jadi bottleneck utama dashboard).
 */
function _countSiswaPerKelasBatch_(tahun, semester, kelasList){
  const counts = {};
  kelasList.forEach(function(k){ counts[k] = 0; });
  if (!kelasList.length) return counts;

  const targetTahun = String(tahun || '');
  const targetSem = String(semester || '').toLowerCase().trim();
  const kelasSet = {};
  kelasList.forEach(function(k){ kelasSet[String(k).trim()] = true; });

  const rows = sheet('RiwayatKelas').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[1] || '') !== targetTahun) continue;
    if (String(r[2] || '').toLowerCase().trim() !== targetSem) continue;
    const kelas = String(r[4] || '').trim();
    if (!kelasSet[kelas]) continue;
    const status = String(r[5] || '').toUpperCase();
    if (status === 'ALUMNI' || status === 'MUTASI_KELUAR') continue;
    counts[kelas] = (counts[kelas] || 0) + 1;
  }
  return counts;
}

/**
 * _computeDashboardMeta_(email, period, jurnalRows)
 * Inti getDashboardMetaData(), menerima baris JURNAL yang sudah dibaca
 * supaya getDashboardAllData() tidak perlu fetch sheet JURNAL dua kali
 * (dulu: sekali di sini, sekali lagi di getDashboardV4Data).
 */
function _computeDashboardMeta_(email, period, jurnalRows){
  const shSiswa = sheet('SISWA');

  let totalKelas = 0;
  let totalSiswa = 0;
  let jadwal     = '-';
  let firstDate  = '';
  let lastDate   = '';

  if(shSiswa){
    const data = shSiswa.getDataRange().getValues();
    const kelasSet = new Set();

    for(let i=1;i<data.length;i++){
      const kelas = data[i][0];
      const owner = String(data[i][5] || '').toLowerCase().trim();

      if(owner !== email) continue;

      if(kelas){
        kelasSet.add(kelas);
        totalSiswa++;
      }
    }

    totalKelas = kelasSet.size;
  }

  // Override total kelas/siswa dari arsitektur multi-tahun jika data tersedia.
  try {
    const kelasAktif = getKelasDiampuAktifForUser_(email);
    if (kelasAktif.length) {
      totalKelas = kelasAktif.length;
      const counts = _countSiswaPerKelasBatch_(period.tahun_pelajaran, period.semester, kelasAktif);
      let sum = 0;
      kelasAktif.forEach(function(k) { sum += (counts[k] || 0); });
      totalSiswa = sum;
    }
  } catch(e) {}

  if(jurnalRows && jurnalRows.length){
    const tanggalList = [];

    for(let i=1;i<jurnalRows.length;i++){
      const owner = String(jurnalRows[i][12] || '').toLowerCase().trim();
      if(owner !== email) continue;
      if(period.tahun_pelajaran && jurnalRows[i][18] && String(jurnalRows[i][18]) !== String(period.tahun_pelajaran)) continue;
      if(period.semester && jurnalRows[i][13] && String(jurnalRows[i][13]).toLowerCase().trim() !== String(period.semester).toLowerCase().trim()) continue;

      const tgl = jurnalRows[i][1];
      if(tgl){
        tanggalList.push(new Date(tgl));
      }
    }

    if(tanggalList.length > 0){
      tanggalList.sort((a,b)=>a-b);
      firstDate = Utilities.formatDate(
        tanggalList[0],
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );
      lastDate = Utilities.formatDate(
        tanggalList[tanggalList.length-1],
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );
    }
  }

  if(totalKelas > 0){
    jadwal = totalKelas + ' Kelas Aktif';
  }

  return {
    totalKelas,
    totalSiswa,
    jadwal,
    firstDate,
    lastDate
  };
}

function getDashboardMetaData(){
  ensureAcademicSchema_();

  const auth = getAuth();
  if(!auth || !auth.email){
    return {
      totalKelas:0,
      totalSiswa:0,
      jadwal:'-',
      firstDate:'',
      lastDate:''
    };
  }

  const email = String(auth.email).toLowerCase().trim();
  const period = getUserAcademicPeriod(email);
  const shJurnal = sheet('JURNAL');
  const jurnalRows = shJurnal ? shJurnal.getDataRange().getValues() : [];

  return _computeDashboardMeta_(email, period, jurnalRows);
}

/**
 * getDashboardAllData()
 * Menggabungkan getDashboardMetaData() + getDashboardV4Data() dalam SATU
 * GAS call sehingga latensi berkurang dari ~12-16s (2 round-trip) menjadi
 * ~6-8s (1 round-trip).
 *
 * Return shape:
 *   { kelasDiampu, totalSiswa, firstDate, lastDate,    ← dari meta
 *     totalJurnal, totalKelas, rata2, risiko,
 *     kelasData, ranking, insight }                    ← dari V4
 */
function getDashboardAllData() {
  // ── CacheService: serve stale while fresh data loads (server-side, 5 min) ──
  var _uc  = CacheService.getUserCache();
  var _hit = _uc.get('DASH_ALL');
  if (_hit) {
    try { return JSON.parse(_hit); } catch(e) { /* parse failed — fall through */ }
  }

  var _auth = getAuth();
  var _email = String(_auth.email || '').toLowerCase().trim();
  var _period = getUserAcademicPeriod(_email);

  var setting = {};
  try { setting = getSetting(); } catch(e) { logError_('getDashboardAllData/getSetting', e); }

  // Baca jadwal LANGSUNG — dengan column-header mapping agar tahan perubahan urutan kolom
  var jadwal = {};
  try {
    if (_email) {
      var _ss = getSpreadsheet_();
      var _shJ = _ss.getSheetByName('JADWAL_SEMESTER')
               || _ss.getSheetByName('jadwal_semester')
               || _ss.getSheetByName('Jadwal_Semester');
      if (_shJ) {
        var _jv = _shJ.getDataRange().getValues();
        // Build column map from header row (row 0)
        var _colMap = {};
        if (_jv.length > 0) {
          _jv[0].forEach(function(h, idx) { _colMap[String(h).toLowerCase().trim()] = idx; });
        }
        var _cEmail    = _colMap['email']       !== undefined ? _colMap['email']       : 0;
        var _cSem      = _colMap['semester']    !== undefined ? _colMap['semester']    : 1;
        var _cHari     = _colMap['hari']        !== undefined ? _colMap['hari']        : 2;
        var _cKelas    = _colMap['kelas']       !== undefined ? _colMap['kelas']       : 3;
        var _cMapel    = _colMap['mapel']       !== undefined ? _colMap['mapel']       : 4;
        var _cMulai    = _colMap['jam_mulai']   !== undefined ? _colMap['jam_mulai']   : 5;
        var _cSelesai  = _colMap['jam_selesai'] !== undefined ? _colMap['jam_selesai'] : 6;
        var _cTahun    = _colMap['tahun_pelajaran'] !== undefined ? _colMap['tahun_pelajaran'] : -1;

        var _semAktif = String(setting.semester || '').trim().toLowerCase();
        var _tahunAktif = String(setting.tahun_pelajaran || '').trim();

        // TIDAK ADA fallback "tampilkan semua kalau kosong" — sengaja dihapus.
        // Fallback itu dulu bikin jadwal PERIODE LAMA nyangkut tampil lagi
        // setiap kali periode aktif memang benar-benar kosong (mis. baru
        // saja di-reset), padahal dashboard harus menunjukkan kosong.
        for (var i = 1; i < _jv.length; i++) {
          if (String(_jv[i][_cEmail] || '').toLowerCase().trim() !== _email) continue;
          if (_semAktif && String(_jv[i][_cSem] || '').trim().toLowerCase() !== _semAktif) continue;
          if (_cTahun !== -1 && _tahunAktif) {
            var _rowTahun = String(_jv[i][_cTahun] || '').trim();
            // Baris lama sebelum migrasi tahun_pelajaran (kosong) dianggap
            // cocok ke periode manapun — baris baru WAJIB cocok tahunnya.
            if (_rowTahun && _rowTahun !== _tahunAktif) continue;
          }
          var hari = String(_jv[i][_cHari] || '').trim().toUpperCase();
          if (!hari) continue;
          if (!jadwal[hari]) jadwal[hari] = [];
          jadwal[hari].push({
            kelas      : String(_jv[i][_cKelas]   || ''),
            mapel      : String(_jv[i][_cMapel]   || ''),
            jam_mulai  : normalizeJam_(_jv[i][_cMulai])   || '--:--',
            jam_selesai: normalizeJam_(_jv[i][_cSelesai]) || '--:--'
          });
        }
        if (typeof _sortJadwalGrouped_ === 'function') jadwal = _sortJadwalGrouped_(jadwal);
      }
    }
  } catch(e) {
    logError_('getDashboardAllData/jadwal', e);
    jadwal = {};
  }

  // Baca JURNAL & ABSENSI SEKALI di sini, dipakai ulang untuk meta + v4 —
  // dulu masing-masing dibaca sendiri-sendiri lewat getDashboardMetaData()
  // dan getDashboardV4Data(), jadi JURNAL kebaca 2x per load dashboard.
  var _shJurnal = _email ? sheet('JURNAL') : null;
  var _shAbsen  = _email ? sheet('ABSENSI') : null;
  var jurnalRows = _shJurnal ? _shJurnal.getDataRange().getValues() : [];
  var absenRows  = _shAbsen ? _shAbsen.getDataRange().getValues() : [];

  var meta = { totalKelas: 0, totalSiswa: 0, jadwal: '-', firstDate: '', lastDate: '' };
  try {
    if (_email) meta = _computeDashboardMeta_(_email, _period, jurnalRows);
  } catch(e) { logError_('getDashboardAllData/meta', e); }

  var dari   = meta.firstDate || '';
  var sampai = meta.lastDate  || '';
  var v4 = emptyDashboard();
  try {
    v4 = _computeV4FromRows_(jurnalRows, absenRows, _email, _period, dari, sampai);
  } catch(e) { logError_('getDashboardAllData/v4', e); }

  var result = {
    setting     : setting,
    jadwal      : jadwal,
    kelasDiampu : meta.totalKelas,
    totalSiswa  : meta.totalSiswa,
    firstDate   : meta.firstDate,
    lastDate    : meta.lastDate,
    totalJurnal : v4.totalJurnal,
    totalKelas  : v4.totalKelas,
    rata2       : v4.rata2,
    risiko      : v4.risiko,
    kelasData   : v4.kelasData,
    ranking     : v4.ranking,
    insight     : v4.insight
  };

  // Store in user cache — 5-minute TTL
  try { _uc.put('DASH_ALL', JSON.stringify(result), 300); } catch(e) { /* quota — ignore */ }

  return result;
}

function getOrCreateFolder(name){

  const folders = DriveApp.getFoldersByName(name);

  if(folders.hasNext()){
    return folders.next();
  }

  return DriveApp.createFolder(name);
}
