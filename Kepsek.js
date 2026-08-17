/**
 * Kepsek.js — Dashboard Kepala Sekolah
 * Menyediakan data rekap seluruh guru di satu sekolah (read-only).
 *
 * Role yang diizinkan: kepsek, superadmin
 */

/**
 * Pastikan caller adalah kepsek atau superadmin.
 */
function assertKepsek_() {
  var auth = getAuth();
  if (auth.role !== 'kepsek' && auth.role !== 'superadmin') {
    throw new Error('AKSES_DITOLAK: Hanya Kepala Sekolah yang dapat mengakses fitur ini.');
  }
}

/**
 * ================================================================
 * MULTI-SPREADSHEET HELPERS (mode per_guru)
 * ================================================================
 * Di mode storage 'central' semua guru berbagi SATU spreadsheet, jadi
 * baca sheet USERS/JURNAL/SISWA/dst sekali saja sudah cukup — begitu
 * getRekapSekolah() aslinya dibuat.
 *
 * Tapi kalau mode aktifnya 'per_guru' (tiap guru punya spreadsheet
 * sendiri, supaya beban Drive tidak numpuk di akun SuperAdmin), data
 * operasional (JURNAL, SISWA, SETTING, NILAI_SISWA, JURNAL_GURU_WALI,
 * dst) TERSEBAR di spreadsheet masing-masing guru — cuma sheet USERS
 * yang tetap selalu di spreadsheet central (lihat _isCentralOnlySheet_
 * di Code.js). Akun Kepsek sendiri TIDAK PERLU di-mapping ke
 * spreadsheet guru manapun — dia cukup dikenali dari role='kepsek' di
 * USERS (central), lalu kode di bawah ini yang membuka SATU-SATU
 * spreadsheet tiap guru aktif dan menggabungkan barisnya. Baris di
 * setiap sheet operasional sudah menyimpan email pemiliknya sendiri
 * (kolom owner/guru_wali), jadi logika agregasi lama yang mengelompokkan
 * per-email tetap jalan tanpa perubahan — cuma sumber barisnya yang
 * sekarang bisa datang dari banyak spreadsheet sekaligus.
 */

/**
 * Daftar email guru/admin berstatus aktif, dari sheet USERS (selalu central).
 * Satu deployment ini bisa melayani BEBERAPA sekolah sekaligus — SuperAdmin
 * mengelompokkan tiap akun ke kode_sekolah lewat updateUserSekolah(). Kalau
 * pemanggilnya seorang Kepsek, daftar guru otomatis DIPERSEMPIT ke guru
 * dengan kode_sekolah yang SAMA dengan Kepsek itu — kalau Kepsek-nya sendiri
 * belum di-assign sekolah, hasilnya sengaja dikosongkan (bukan malah
 * menampilkan guru yang juga belum di-assign) supaya tidak bocor data
 * sekolah lain. SuperAdmin (bukan kepsek) tetap melihat semua sekolah
 * sekaligus seperti sebelumnya.
 */
function _kepsekActiveGuruEmails_() {
  var auth  = getAuth();
  var shUsers = _ensureUsersKodeSekolahColumn_() || _getCentralSheetByName_('USERS');
  var guruEmails = [];
  if (!shUsers) return guruEmails;

  var userData = shUsers.getDataRange().getValues();

  var scopeToSchool = auth.role === 'kepsek';
  var callerKode = '';
  if (scopeToSchool) {
    for (var ci = 1; ci < userData.length; ci++) {
      if (String(userData[ci][0] || '').toLowerCase().trim() === auth.email) {
        callerKode = String(userData[ci][4] || '').trim();
        break;
      }
    }
    if (!callerKode) return guruEmails; // Kepsek belum di-assign sekolah — kosongkan, jangan tebak
  }

  for (var ui = 1; ui < userData.length; ui++) {
    var uEmail = String(userData[ui][0] || '').toLowerCase().trim();
    var uRole  = String(userData[ui][1] || '').toLowerCase().trim();
    var uStatus= String(userData[ui][2] || '').toLowerCase().trim();
    if (!uEmail || !(uRole === 'admin' || uRole === 'guru') || uStatus !== 'active') continue;
    if (scopeToSchool && String(userData[ui][4] || '').trim() !== callerKode) continue;
    guruEmails.push(uEmail);
  }
  return guruEmails;
}

/**
 * Buka spreadsheet operasional untuk tiap guru di guruEmails.
 * Mode central -> semua guru menunjuk ke objek spreadsheet central yang sama
 * (dibuka sekali). Mode per_guru -> resolve spreadsheet_id masing-masing
 * lewat RESOURCE_MAP/DEPLOYMENTS; guru yang belum pernah login/belum
 * ter-provisioning dilewati saja (bukan error) supaya satu guru bermasalah
 * tidak menggagalkan rekap guru lain.
 * @returns {Object} map email -> Spreadsheet (guru yang gagal dibuka tidak ada di map)
 */
function _kepsekOpenSpreadsheetsMap_(guruEmails) {
  var mode = getStorageMode_();
  var map = {};
  if (mode !== 'per_guru') {
    var centralSS = getCentralSpreadsheet_();
    guruEmails.forEach(function(email) { map[email] = centralSS; });
    return map;
  }

  var ssById = {}; // cache biar spreadsheet yang sama tidak dibuka berkali-kali
  guruEmails.forEach(function(email) {
    var sid = resolveSpreadsheetIdForUser_(email);
    if (!sid) return; // belum ter-provisioning — dilewati, bukan dianggap error
    if (ssById[sid] === undefined) {
      try { ssById[sid] = SpreadsheetApp.openById(sid); }
      catch (e) { ssById[sid] = null; }
    }
    if (ssById[sid]) map[email] = ssById[sid];
  });
  return map;
}

/**
 * Gabungkan baris dari sheet bernama sama di banyak spreadsheet guru
 * jadi satu array getValues()-style (header sekali di baris 0, lalu
 * semua baris data). Sheet yang tidak ada di spreadsheet guru tertentu
 * (mis. guru itu belum pernah pakai fitur Guru Wali) dilewati saja.
 */
function _kepsekReadSheetRowsMulti_(ssMap, guruEmails, sheetName) {
  var header = null;
  var dataRows = [];
  var seenSS = {}; // mode central: semua guru share 1 objek SS, jangan baca berkali-kali
  guruEmails.forEach(function(email) {
    var ss = ssMap[email];
    if (!ss) return;
    var ssKey = ss.getId();
    if (seenSS[ssKey]) return; // sudah dibaca (mode central atau kebetulan share spreadsheet)
    seenSS[ssKey] = true;

    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    var vals = sh.getDataRange().getValues();
    if (!vals.length) return;
    if (!header) header = vals[0];
    dataRows = dataRows.concat(vals.slice(1));
  });
  return header ? [header].concat(dataRows) : [];
}

/**
 * getRekapSekolah()
 * Mengembalikan rekap seluruh guru di sekolah yang sama
 * (semua user bertipe 'admin' di USERS sheet).
 *
 * Return shape:
 * {
 *   sekolah          : string,
 *   tahun_pelajaran  : string,
 *   semester         : string,
 *   totalGuru        : number,
 *   guruAktif        : number,   // punya jurnal bulan ini
 *   totalJurnalBulan : number,
 *   totalSiswa       : number,
 *   guruList         : [{
 *     email, nama, totalJurnal, totalJurnalBulan,
 *     totalSiswa, totalKelas, lastJurnal,
 *     statusJurnal: 'aktif'|'tidak_aktif'|'belum_mulai'
 *   }]
 * }
 */
function getRekapSekolah(forceRefresh) {
  assertKepsek_();

  var cacheKey = 'KEPSEK_REKAP_SEKOLAH';
  if (!forceRefresh) {
    try {
      var cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) { /* cache miss/corrupt, lanjut hitung ulang */ }
  }

  var guruEmails = _kepsekActiveGuruEmails_();
  var ssMap      = _kepsekOpenSpreadsheetsMap_(guruEmails);

  var jurnalRows = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'JURNAL');
  var siswaRows  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'SISWA');
  var setAll     = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'SETTING');
  var nilaiRows  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'NILAI_SISWA');
  var settingRowsNilai = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'SETTING_NILAI');

  // ── Setting sekolah (ambil baris pertama yang terisi) ──
  var sekolah         = '';
  var tahun_pelajaran = '';
  var semester_aktif  = '';
  for (var si = 1; si < setAll.length; si++) {
    if (setAll[si][1]) { sekolah         = String(setAll[si][1] || ''); }
    if (setAll[si][2]) { tahun_pelajaran = String(setAll[si][2] || ''); }
    if (setAll[si][3]) { semester_aktif  = String(setAll[si][3] || ''); break; }
  }

  // ── Baca setting per guru (nama_guru) ──
  var namaMap = {};
  for (var ni = 1; ni < setAll.length; ni++) {
    var nEmail = String(setAll[ni][0] || '').toLowerCase().trim();
    var nNama  = String(setAll[ni][4] || '').trim(); // kolom nama_guru
    if (nEmail && nNama) namaMap[nEmail] = nNama;
  }

  // ── Tanggal awal bulan ini ──
  var now       = new Date();
  var bulanIni  = new Date(now.getFullYear(), now.getMonth(), 1);

  // ── Hitung per guru ──
  var guruMap = {};
  guruEmails.forEach(function(email) {
    guruMap[email] = {
      email            : email,
      nama             : namaMap[email] || email,
      totalJurnal      : 0,
      totalJurnalBulan : 0,
      totalSiswa       : 0,
      totalKelas       : 0,
      lastJurnal       : null
    };
  });

  for (var ji = 1; ji < jurnalRows.length; ji++) {
    var jEmail = String(jurnalRows[ji][12] || '').toLowerCase().trim();
    if (!guruMap[jEmail]) continue;

    guruMap[jEmail].totalJurnal++;

    var tgl = jurnalRows[ji][1] ? new Date(jurnalRows[ji][1]) : null;
    if (tgl) {
      if (!guruMap[jEmail].lastJurnal || tgl > guruMap[jEmail].lastJurnal) {
        guruMap[jEmail].lastJurnal = tgl;
      }
      if (tgl >= bulanIni) {
        guruMap[jEmail].totalJurnalBulan++;
      }
    }
  }

  // Siswa & kelas per guru (dibatasi tahun ajaran aktif sekolah, kalau ada)
  for (var ssi = 1; ssi < siswaRows.length; ssi++) {
    var sOwner = String(siswaRows[ssi][5] || '').toLowerCase().trim();
    if (!guruMap[sOwner]) continue;
    if (!_siswaRowMatchesPeriode_(siswaRows[ssi], tahun_pelajaran)) continue;
    guruMap[sOwner].totalSiswa++;
  }

  // Kelas unik per guru dari jurnal
  var kelasPerGuru = {};
  for (var kji = 1; kji < jurnalRows.length; kji++) {
    var kEmail = String(jurnalRows[kji][12] || '').toLowerCase().trim();
    if (!guruMap[kEmail]) continue;
    if (!kelasPerGuru[kEmail]) kelasPerGuru[kEmail] = new Set();
    if (jurnalRows[kji][2]) kelasPerGuru[kEmail].add(jurnalRows[kji][2]);
  }
  guruEmails.forEach(function(email) {
    guruMap[email].totalKelas = kelasPerGuru[email] ? kelasPerGuru[email].size : 0;
  });

  // ── Status jurnal ──
  var tz = Session.getScriptTimeZone();
  var guruList = guruEmails.map(function(email) {
    var g = guruMap[email];
    var status = 'belum_mulai';
    if (g.totalJurnal > 0) {
      status = g.totalJurnalBulan > 0 ? 'aktif' : 'tidak_aktif';
    }
    return {
      email            : g.email,
      nama             : g.nama,
      totalJurnal      : g.totalJurnal,
      totalJurnalBulan : g.totalJurnalBulan,
      totalSiswa       : g.totalSiswa,
      totalKelas       : g.totalKelas,
      lastJurnal       : g.lastJurnal
        ? Utilities.formatDate(g.lastJurnal, tz, 'dd MMM yyyy')
        : '-',
      statusJurnal     : status
    };
  });

  // Sort: aktif dulu, lalu tidak_aktif, lalu belum_mulai; per group sort by totalJurnalBulan desc
  var ORDER = { aktif: 0, tidak_aktif: 1, belum_mulai: 2 };
  guruList.sort(function(a, b) {
    if (ORDER[a.statusJurnal] !== ORDER[b.statusJurnal]) {
      return ORDER[a.statusJurnal] - ORDER[b.statusJurnal];
    }
    return b.totalJurnalBulan - a.totalJurnalBulan;
  });

  var guruAktif        = guruList.filter(function(g) { return g.statusJurnal === 'aktif'; }).length;
  var totalJurnalBulan = guruList.reduce(function(s, g) { return s + g.totalJurnalBulan; }, 0);
  var totalSiswa       = guruList.reduce(function(s, g) { return s + g.totalSiswa; }, 0);

  // ── Kelas favorit berdasarkan frekuensi jurnal ──
  var kelasCount = {};
  for (var kj = 1; kj < jurnalRows.length; kj++) {
    var kelasKey = String(jurnalRows[kj][2] || '').trim();
    var emailKey = String(jurnalRows[kj][12] || '').toLowerCase().trim();
    if (!kelasKey || !guruMap[emailKey]) continue;
    kelasCount[kelasKey] = (kelasCount[kelasKey] || 0) + 1;
  }
  var kelasFavorit = Object.keys(kelasCount).map(function(kelas) {
    return { kelas: kelas, totalJurnal: kelasCount[kelas] };
  }).sort(function(a, b) { return b.totalJurnal - a.totalJurnal; }).slice(0, 5);

  // ── Ketuntasan siswa per kelas berdasarkan nilai akhir ──
  var kelasKetuntasan = [];
  if (nilaiRows.length > 1) {
      var nh = nilaiRows[0].map(function(h) { return String(h || '').toLowerCase().trim(); });
      var nidx = {};
      nh.forEach(function(h, i) { nidx[h] = i; });

      var settingRows = settingRowsNilai;
      var kkmMap = {};
      if (settingRows.length > 1) {
        var shh = settingRows[0].map(function(h) { return String(h || '').toLowerCase().trim(); });
        var sidx = {};
        shh.forEach(function(h, i) { sidx[h] = i; });
        for (var si = 1; si < settingRows.length; si++) {
          var klsSet = String(settingRows[si][sidx.kelas] || '').trim();
          var thnSet = String(settingRows[si][sidx.tahun] || '').trim();
          var semSet = String(settingRows[si][sidx.semester] || '').trim();
          var kkmSet = Number(settingRows[si][sidx.kkm]);
          if (!klsSet || !thnSet || !semSet || isNaN(kkmSet)) continue;
          if (String(thnSet) !== String(tahun_pelajaran || '')) continue;
          if (String(semSet).toLowerCase() !== String(semester_aktif || '').toLowerCase()) continue;
          if (!kkmMap[klsSet]) kkmMap[klsSet] = [];
          kkmMap[klsSet].push(kkmSet);
        }
      }

      var classStudents = {};
      for (var ni = 1; ni < nilaiRows.length; ni++) {
        var rowTahun = String(nilaiRows[ni][nidx.tahun] || '').trim();
        var rowSem = String(nilaiRows[ni][nidx.semester] || '').toLowerCase().trim();
        if (tahun_pelajaran && rowTahun !== String(tahun_pelajaran)) continue;
        if (semester_aktif && rowSem !== String(semester_aktif).toLowerCase()) continue;

        var kelasNilai = String(nilaiRows[ni][nidx.kelas] || '').trim();
        var nisNilai = String(nilaiRows[ni][nidx.nis] || '').trim();
        var akhirRaw = nilaiRows[ni][nidx.nilai_akhir];
        var akhir = akhirRaw === '' || akhirRaw === null || akhirRaw === undefined ? null : Number(akhirRaw);
        if (!kelasNilai || !nisNilai || isNaN(akhir)) continue;
        if (!classStudents[kelasNilai]) classStudents[kelasNilai] = {};
        if (!classStudents[kelasNilai][nisNilai]) classStudents[kelasNilai][nisNilai] = [];
        classStudents[kelasNilai][nisNilai].push(akhir);
      }

      Object.keys(classStudents).forEach(function(kelas) {
        var siswaMap = classStudents[kelas];
        var nisList = Object.keys(siswaMap);
        if (!nisList.length) return;
        var kkmList = kkmMap[kelas] || [70];
        var kkm = Math.round((kkmList.reduce(function(a, b) { return a + b; }, 0) / kkmList.length) * 100) / 100;
        var tuntas = 0;
        nisList.forEach(function(nis) {
          var avg = siswaMap[nis].reduce(function(a, b) { return a + b; }, 0) / siswaMap[nis].length;
          if (avg >= kkm) tuntas++;
        });
        kelasKetuntasan.push({
          kelas: kelas,
          totalSiswa: nisList.length,
          tuntas: tuntas,
          belumTuntas: nisList.length - tuntas,
          kkm: kkm,
          persenTuntas: Math.round((tuntas / nisList.length) * 100)
        });
      });

      kelasKetuntasan.sort(function(a, b) {
        if (b.persenTuntas !== a.persenTuntas) return b.persenTuntas - a.persenTuntas;
        return b.totalSiswa - a.totalSiswa;
      });
      kelasKetuntasan = kelasKetuntasan.slice(0, 10);
    }

  var result = {
    sekolah         : sekolah,
    tahun_pelajaran : tahun_pelajaran,
    semester        : semester_aktif,
    bulanLabel      : Utilities.formatDate(bulanIni, tz, 'MMMM yyyy'),
    totalGuru       : guruEmails.length,
    guruAktif       : guruAktif,
    totalJurnalBulan: totalJurnalBulan,
    totalSiswa      : totalSiswa,
    kelasFavorit    : kelasFavorit,
    kelasKetuntasan : kelasKetuntasan,
    guruList        : guruList
  };

  // Di mode per_guru fungsi ini membuka spreadsheet tiap guru satu-satu
  // (bisa berat kalau guru banyak) — cache 3 menit supaya buka tab/refresh
  // beruntun tidak mengulang proses berat itu. Tombol "🔄 Refresh" di UI
  // memanggil dengan forceRefresh=true untuk melewati cache.
  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 180); } catch (e) { /* > 100KB, lewati cache */ }

  return result;
}

/**
 * getRekapGuruWali(forceRefresh)
 * Rekap Jurnal Guru Wali (pendampingan) semua guru di sekolah — Kepsek bisa
 * lihat berapa guru sudah mengisi jurnal wali dan ringkasannya per guru.
 * Isi lengkap tiap entri jurnal diambil terpisah lewat
 * getRekapGuruWaliDetail(email) supaya payload awal ini tetap ringkas.
 */
function getRekapGuruWali(forceRefresh) {
  assertKepsek_();

  var cacheKey = 'KEPSEK_REKAP_GURU_WALI';
  if (!forceRefresh) {
    try {
      var cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) { /* cache miss/corrupt, lanjut hitung ulang */ }
  }

  var guruEmails = _kepsekActiveGuruEmails_();
  var ssMap      = _kepsekOpenSpreadsheetsMap_(guruEmails);

  var jgwRows = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'JURNAL_GURU_WALI');
  var setAll  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'SETTING');

  var namaMap = {};
  var activeTahun = '';
  for (var si = 1; si < setAll.length; si++) {
    var nEmail = String(setAll[si][0] || '').toLowerCase().trim();
    var nNama  = String(setAll[si][4] || '').trim();
    if (nEmail && nNama) namaMap[nEmail] = nNama;
    if (!activeTahun && setAll[si][2]) activeTahun = String(setAll[si][2] || '');
  }

  var now      = new Date();
  var bulanIni = new Date(now.getFullYear(), now.getMonth(), 1);
  var tz       = Session.getScriptTimeZone();

  var guruMap = {};
  guruEmails.forEach(function(email) {
    guruMap[email] = { email: email, nama: namaMap[email] || email, totalEntri: 0, entriBulanIni: 0, lastEntri: null };
  });

  for (var i = 1; i < jgwRows.length; i++) {
    var jEmail = String(jgwRows[i][9] || '').toLowerCase().trim();
    if (!guruMap[jEmail]) continue;
    var rowTahun = String(jgwRows[i][11] || '').trim();
    if (activeTahun && rowTahun && rowTahun !== activeTahun) continue;

    guruMap[jEmail].totalEntri++;
    var tgl = jgwRows[i][1] ? new Date(jgwRows[i][1]) : null;
    if (tgl) {
      if (!guruMap[jEmail].lastEntri || tgl > guruMap[jEmail].lastEntri) guruMap[jEmail].lastEntri = tgl;
      if (tgl >= bulanIni) guruMap[jEmail].entriBulanIni++;
    }
  }

  var guruList = guruEmails.map(function(email) {
    var g = guruMap[email];
    var status = 'belum_mulai';
    if (g.totalEntri > 0) status = g.entriBulanIni > 0 ? 'aktif' : 'tidak_aktif';
    return {
      email        : g.email,
      nama         : g.nama,
      totalEntri   : g.totalEntri,
      entriBulanIni: g.entriBulanIni,
      lastEntri    : g.lastEntri ? Utilities.formatDate(g.lastEntri, tz, 'dd MMM yyyy') : '-',
      statusJurnal : status
    };
  });

  var ORDER = { aktif: 0, tidak_aktif: 1, belum_mulai: 2 };
  guruList.sort(function(a, b) {
    if (ORDER[a.statusJurnal] !== ORDER[b.statusJurnal]) return ORDER[a.statusJurnal] - ORDER[b.statusJurnal];
    return b.entriBulanIni - a.entriBulanIni;
  });

  var result = {
    activeTahun     : activeTahun,
    totalGuru       : guruEmails.length,
    guruSudahMengisi: guruList.filter(function(g) { return g.totalEntri > 0; }).length,
    guruList        : guruList
  };

  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 180); } catch (e) { /* > 100KB, lewati cache */ }
  return result;
}

/**
 * getRekapGuruWaliDetail(email)
 * Isi lengkap jurnal pendampingan wali SATU guru — dipanggil saat Kepsek
 * klik/expand nama guru tertentu di tabel Rekap Guru Wali.
 */
function getRekapGuruWaliDetail(email) {
  assertKepsek_();
  var targetEmail = String(email || '').toLowerCase().trim();
  if (!targetEmail) return [];

  var ssMap = _kepsekOpenSpreadsheetsMap_([targetEmail]);
  var ss = ssMap[targetEmail];
  if (!ss) return [];

  var sh = ss.getSheetByName('JURNAL_GURU_WALI');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();

  var activeTahun = '';
  var shSet = ss.getSheetByName('SETTING');
  if (shSet) {
    var setRows = shSet.getDataRange().getValues();
    for (var si = 1; si < setRows.length; si++) {
      if (String(setRows[si][0] || '').toLowerCase().trim() === targetEmail) {
        activeTahun = String(setRows[si][2] || '');
        break;
      }
    }
  }

  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][9] || '').toLowerCase().trim();
    if (rowEmail !== targetEmail) continue;
    var rowTahun = String(data[i][11] || '').trim();
    if (activeTahun && rowTahun && rowTahun !== activeTahun) continue;

    var tgl = data[i][1] ? new Date(data[i][1]) : null;
    var tglFmt = tgl ? Utilities.formatDate(tgl, tz, 'dd MMMM yyyy') : '-';
    var hariStr = String(data[i][2] || '');

    result.push({
      tanggal           : hariStr ? hariStr + ', ' + tglFmt : tglFmt,
      waktu             : String(data[i][3] || '-'),
      fokus_pendampingan: String(data[i][4] || '-'),
      topik_pendampingan: String(data[i][5] || '-'),
      catatan           : String(data[i][6] || '-'),
      tindak_lanjut     : String(data[i][7] || '-')
    });
  }
  return result.reverse();
}

/**
 * getRekapAbsensiSekolah(dari, sampai)
 * Rekap kehadiran semua kelas semua guru dalam rentang tanggal.
 * Return: array of { kelas, guru, email, H, S, I, A, total, persen }
 */
function getRekapAbsensiSekolah(dari, sampai) {
  assertKepsek_();

  var cacheKey = 'KEPSEK_REKAP_ABSENSI_' + String(dari || '') + '_' + String(sampai || '');
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* cache miss/corrupt, lanjut hitung ulang */ }

  var guruEmails = _kepsekActiveGuruEmails_();
  var ssMap      = _kepsekOpenSpreadsheetsMap_(guruEmails);
  var guruSet    = new Set(guruEmails);

  // Nama guru map
  var setAll = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'SETTING');
  var namaMap = {};
  for (var si = 1; si < setAll.length; si++) {
    var em = String(setAll[si][0] || '').toLowerCase().trim();
    var nm = String(setAll[si][4] || '').trim();
    if (em && nm) namaMap[em] = nm;
  }

  var from = dari   ? new Date(dari   + 'T00:00:00') : null;
  var to   = sampai ? new Date(sampai + 'T23:59:59') : null;

  var jurnalData = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'JURNAL');
  var absenData  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'ABSENSI');

  // Build jurnal id → {email, kelas}
  var jurnalMap = {};
  for (var ji = 1; ji < jurnalData.length; ji++) {
    var jEmail = String(jurnalData[ji][12] || '').toLowerCase().trim();
    if (!guruSet.has(jEmail)) continue;
    var tgl = jurnalData[ji][1] ? new Date(jurnalData[ji][1]) : null;
    if (from && tgl < from) continue;
    if (to   && tgl > to)   continue;
    jurnalMap[jurnalData[ji][0]] = {
      email : jEmail,
      kelas : String(jurnalData[ji][2] || '')
    };
  }

  // Accumulate absensi
  var kelasMap = {};
  for (var ai = 1; ai < absenData.length; ai++) {
    var jId = absenData[ai][0];
    var j   = jurnalMap[jId];
    if (!j) continue;
    var key = j.kelas + '___' + j.email;
    if (!kelasMap[key]) {
      kelasMap[key] = { kelas: j.kelas, email: j.email, H: 0, S: 0, I: 0, A: 0, total: 0 };
    }
    var status = String(absenData[ai][3] || '').toUpperCase();
    if (kelasMap[key][status] !== undefined) kelasMap[key][status]++;
    kelasMap[key].total++;
  }

  var result = Object.values(kelasMap).map(function(d) {
    return {
      kelas  : d.kelas,
      guru   : namaMap[d.email] || d.email,
      email  : d.email,
      H      : d.H, S: d.S, I: d.I, A: d.A,
      total  : d.total,
      persen : d.total > 0 ? Math.round(d.H / d.total * 100) : 0
    };
  });

  result.sort(function(a, b) { return a.persen - b.persen; }); // risiko terbesar dulu

  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 180); } catch (e) { /* > 100KB, lewati cache */ }
  return result;
}
/**
 * getEarlyWarningSiswa(threshold)
 * Identifikasi siswa dengan kehadiran di bawah threshold (default 75%).
 * Returns: array of {kelas, nis, nama, H, S, I, A, total, persen}
 */
function getEarlyWarningSiswa(threshold) {
  assertKepsek_();
  var limit = (typeof threshold === 'number') ? threshold : 75;

  var cacheKey = 'KEPSEK_EARLY_WARNING_' + limit;
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* cache miss/corrupt, lanjut hitung ulang */ }

  var guruEmails = _kepsekActiveGuruEmails_();
  var ssMap      = _kepsekOpenSpreadsheetsMap_(guruEmails);
  var guruSet    = new Set(guruEmails);

  var jurnalData = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'JURNAL');
  var absenData  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'ABSENSI');
  var siswaData  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'SISWA');

  if (!jurnalData.length || !absenData.length || !siswaData.length) return [];

  // Build jurnal id → kelas (hanya guru sekolah ini)
  var jurnalMap  = {};
  for (var ji = 1; ji < jurnalData.length; ji++) {
    var jEmail = String(jurnalData[ji][12] || '').toLowerCase().trim();
    if (!guruSet.has(jEmail)) continue;
    jurnalMap[jurnalData[ji][0]] = {
      kelas: String(jurnalData[ji][2] || ''),
      email: jEmail
    };
  }

  // Build siswa map: kelas → {nis → nama}
  var siswaMap  = {};
  for (var si = 1; si < siswaData.length; si++) {
    var sOwner = String(siswaData[si][5] || '').toLowerCase().trim();
    if (!guruSet.has(sOwner)) continue;
    var kelas = String(siswaData[si][0] || '');
    var nis   = String(siswaData[si][2] || '');
    var nama  = String(siswaData[si][3] || '');
    if (!siswaMap[kelas]) siswaMap[kelas] = {};
    if (nis) siswaMap[kelas][nis] = nama;
  }

  // Akumulasi absensi per siswa per kelas
  // ABSENSI: [0]=jurnal_id, [1]=no_absen, [2]=nis, [3]=status, [4]=keterangan, [5]=nama
  var siswaAbsen  = {}; // key: kelas+nis

  for (var ai = 1; ai < absenData.length; ai++) {
    var jId = absenData[ai][0];
    if (!jurnalMap[jId]) continue;

    var kls    = jurnalMap[jId].kelas;
    var nis_ab = String(absenData[ai][2] || '').trim();
    var nama_ab = String(absenData[ai][5] || absenData[ai][1] || '').trim();
    var status  = String(absenData[ai][3] || '').toUpperCase();
    if (!nis_ab) continue;

    var sKey = kls + '__' + nis_ab;
    if (!siswaAbsen[sKey]) {
      siswaAbsen[sKey] = { kelas: kls, nis: nis_ab, nama: nama_ab, H: 0, S: 0, I: 0, A: 0, total: 0 };
    }
    if (['H','S','I','A'].indexOf(status) > -1) {
      siswaAbsen[sKey][status]++;
      siswaAbsen[sKey].total++;
    }
    if (nama_ab && !siswaAbsen[sKey].nama) siswaAbsen[sKey].nama = nama_ab;
  }

  var result = [];
  var keys = Object.keys(siswaAbsen);
  for (var k = 0; k < keys.length; k++) {
    var d = siswaAbsen[keys[k]];
    var persen = d.total > 0 ? Math.round(d.H / d.total * 100) : 0;
    if (persen < limit) {
      result.push({
        kelas  : d.kelas,
        nis    : d.nis,
        nama   : d.nama,
        H      : d.H,
        S      : d.S,
        I      : d.I,
        A      : d.A,
        total  : d.total,
        persen : persen
      });
    }
  }
  result.sort(function(a, b) { return a.persen - b.persen; });

  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 180); } catch (e) { /* > 100KB, lewati cache */ }
  return result;
}

/**
 * getRaportAbsensiSiswa()
 * Raport absensi per siswa (semua kelas, semua guru di sekolah ini).
 * Returns: array of {kelas, nis, nama, H, S, I, A, total, persen}
 */
function getRaportAbsensiSiswa() {
  assertKepsek_();

  var cacheKey = 'KEPSEK_RAPORT_ABSENSI';
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* cache miss/corrupt, lanjut hitung ulang */ }

  var guruEmails = _kepsekActiveGuruEmails_();
  var ssMap      = _kepsekOpenSpreadsheetsMap_(guruEmails);
  var guruSet    = new Set(guruEmails);

  var jurnalData = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'JURNAL');
  var absenData  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'ABSENSI');

  if (!jurnalData.length || !absenData.length) return [];

  var jurnalMap  = {};
  for (var ji = 1; ji < jurnalData.length; ji++) {
    var jEmail = String(jurnalData[ji][12] || '').toLowerCase().trim();
    if (!guruSet.has(jEmail)) continue;
    jurnalMap[jurnalData[ji][0]] = String(jurnalData[ji][2] || '');
  }

  var siswaMap   = {};

  for (var ai = 1; ai < absenData.length; ai++) {
    var jId = absenData[ai][0];
    if (!jurnalMap[jId]) continue;
    var kls     = jurnalMap[jId];
    var nisAb   = String(absenData[ai][2] || '').trim();
    var namaAb  = String(absenData[ai][5] || '').trim() || String(absenData[ai][1] || '').trim();
    var status  = String(absenData[ai][3] || '').toUpperCase();
    if (!nisAb) continue;
    var sKey = kls + '__' + nisAb;
    if (!siswaMap[sKey]) {
      siswaMap[sKey] = { kelas: kls, nis: nisAb, nama: namaAb, H:0, S:0, I:0, A:0, total:0 };
    }
    if (['H','S','I','A'].indexOf(status) > -1) {
      siswaMap[sKey][status]++;
      siswaMap[sKey].total++;
    }
    if (namaAb && !siswaMap[sKey].nama) siswaMap[sKey].nama = namaAb;
  }

  var result = Object.values(siswaMap).map(function(d) {
    return {
      kelas  : d.kelas,
      nis    : d.nis,
      nama   : d.nama,
      H      : d.H, S: d.S, I: d.I, A: d.A,
      total  : d.total,
      persen : d.total > 0 ? Math.round(d.H / d.total * 100) : 0
    };
  });
  result.sort(function(a, b) {
    if (a.kelas < b.kelas) return -1;
    if (a.kelas > b.kelas) return 1;
    return a.nama < b.nama ? -1 : 1;
  });

  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 180); } catch (e) { /* > 100KB, lewati cache */ }
  return result;
}

/**
 * getJurnalChartData()
 * Data chart untuk kepsek:
 * - Tren jurnal mingguan (4 minggu terakhir) per sekolah
 * - Distribusi status kehadiran (H/S/I/A) secara total
 */
function getJurnalChartData() {
  assertKepsek_();

  var cacheKey = 'KEPSEK_JURNAL_CHART';
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* cache miss/corrupt, lanjut hitung ulang */ }

  var guruEmails = _kepsekActiveGuruEmails_();
  var ssMap      = _kepsekOpenSpreadsheetsMap_(guruEmails);
  var guruSet    = new Set(guruEmails);

  var jurnalData = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'JURNAL');
  var absenData  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'ABSENSI');

  if (!jurnalData.length) return { weekly: [], absence: {H:0,S:0,I:0,A:0}, labels: [] };

  // Tren jurnal per minggu (4 minggu terakhir)
  var now      = new Date();
  var weekMs   = 7 * 24 * 60 * 60 * 1000;
  var weeks    = [];
  var labels   = [];
  var tz       = Session.getScriptTimeZone();
  for (var w = 3; w >= 0; w--) {
    var start  = new Date(now.getTime() - (w + 1) * weekMs);
    var end    = new Date(now.getTime() - w * weekMs);
    weeks.push({ start: start, end: end, count: 0 });
    var label  = 'Mg ' + (4 - w) + ' (' + Utilities.formatDate(start, tz, 'dd/MM') + ')';
    labels.push(label);
  }

  var jurnalIds  = new Set();
  for (var ji = 1; ji < jurnalData.length; ji++) {
    var jEmail = String(jurnalData[ji][12] || '').toLowerCase().trim();
    if (!guruSet.has(jEmail)) continue;
    var tgl = jurnalData[ji][1] ? new Date(jurnalData[ji][1]) : null;
    if (!tgl) continue;
    jurnalIds.add(String(jurnalData[ji][0]));
    for (var wk = 0; wk < weeks.length; wk++) {
      if (tgl >= weeks[wk].start && tgl < weeks[wk].end) {
        weeks[wk].count++;
        break;
      }
    }
  }

  // Distribusi absensi total
  var absTotals = { H:0, S:0, I:0, A:0 };
  for (var ai = 1; ai < absenData.length; ai++) {
    if (!jurnalIds.has(String(absenData[ai][0]))) continue;
    var st = String(absenData[ai][3] || '').toUpperCase();
    if (absTotals[st] !== undefined) absTotals[st]++;
  }

  var chartResult = {
    labels : labels,
    weekly : weeks.map(function(w) { return w.count; }),
    absence: absTotals
  };
  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(chartResult), 180); } catch (e) { /* > 100KB, lewati cache */ }
  return chartResult;
}

// Fitur AI Summary dihapus (quota Gemini API habis)
