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
function getRekapSekolah() {
  assertKepsek_();

  var ss        = getSpreadsheet_();
  var shUsers   = ss.getSheetByName('USERS');
  var shJurnal  = ss.getSheetByName('JURNAL');
  var shSiswa   = ss.getSheetByName('SISWA');
  var shSetting = ss.getSheetByName('SETTING');

  // ── Setting sekolah ──
  var sekolah         = '';
  var tahun_pelajaran = '';
  var semester_aktif  = '';
  if (shSetting) {
    var setRows = shSetting.getDataRange().getValues();
    // Find any row (setting is per-user, take the most common sekolah)
    for (var si = 1; si < setRows.length; si++) {
      if (setRows[si][1]) { sekolah         = String(setRows[si][1] || ''); }
      if (setRows[si][2]) { tahun_pelajaran = String(setRows[si][2] || ''); }
      if (setRows[si][3]) { semester_aktif  = String(setRows[si][3] || ''); break; }
    }
  }

  // ── Daftar semua guru (role = admin) ──
  var guruEmails = [];
  if (shUsers) {
    var userData = shUsers.getDataRange().getValues();
    for (var ui = 1; ui < userData.length; ui++) {
      var uEmail = String(userData[ui][0] || '').toLowerCase().trim();
      var uRole  = String(userData[ui][1] || '').toLowerCase().trim();
      var uStatus= String(userData[ui][2] || '').toLowerCase().trim();
      if (uEmail && (uRole === 'admin' || uRole === 'guru') && uStatus === 'active') {
        guruEmails.push(uEmail);
      }
    }
  }

  // ── Baca jurnal semua guru ──
  var jurnalRows = [];
  if (shJurnal) { jurnalRows = shJurnal.getDataRange().getValues(); }

  // ── Baca siswa semua guru ──
  var siswaRows = [];
  if (shSiswa) { siswaRows = shSiswa.getDataRange().getValues(); }

  // ── Baca setting per guru (nama_guru) ──
  var namaMap = {};
  if (shSetting) {
    var setAll = shSetting.getDataRange().getValues();
    for (var ni = 1; ni < setAll.length; ni++) {
      var nEmail = String(setAll[ni][0] || '').toLowerCase().trim();
      var nNama  = String(setAll[ni][4] || '').trim(); // kolom nama_guru
      if (nEmail && nNama) namaMap[nEmail] = nNama;
    }
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

  // Siswa & kelas per guru
  for (var ssi = 1; ssi < siswaRows.length; ssi++) {
    var sOwner = String(siswaRows[ssi][5] || '').toLowerCase().trim();
    if (!guruMap[sOwner]) continue;
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

  return {
    sekolah         : sekolah,
    tahun_pelajaran : tahun_pelajaran,
    semester        : semester_aktif,
    bulanLabel      : Utilities.formatDate(bulanIni, tz, 'MMMM yyyy'),
    totalGuru       : guruEmails.length,
    guruAktif       : guruAktif,
    totalJurnalBulan: totalJurnalBulan,
    totalSiswa      : totalSiswa,
    guruList        : guruList
  };
}

/**
 * getRekapAbsensiSekolah(dari, sampai)
 * Rekap kehadiran semua kelas semua guru dalam rentang tanggal.
 * Return: array of { kelas, guru, email, H, S, I, A, total, persen }
 */
function getRekapAbsensiSekolah(dari, sampai) {
  assertKepsek_();

  var ss       = getSpreadsheet_();
  var shJurnal = ss.getSheetByName('JURNAL');
  var shAbsen  = ss.getSheetByName('ABSENSI');
  var shUsers  = ss.getSheetByName('USERS');
  var shSetting= ss.getSheetByName('SETTING');

  if (!shJurnal || !shAbsen) return [];

  // Guru aktif set
  var guruSet = new Set();
  if (shUsers) {
    var ud = shUsers.getDataRange().getValues();
    for (var ui = 1; ui < ud.length; ui++) {
      if (String(ud[ui][1] || '').toLowerCase() === 'admin' &&
          String(ud[ui][2] || '').toLowerCase() === 'active') {
        guruSet.add(String(ud[ui][0] || '').toLowerCase().trim());
      }
    }
  }

  // Nama guru map
  var namaMap = {};
  if (shSetting) {
    var sd = shSetting.getDataRange().getValues();
    for (var si = 1; si < sd.length; si++) {
      var em = String(sd[si][0] || '').toLowerCase().trim();
      var nm = String(sd[si][4] || '').trim();
      if (em && nm) namaMap[em] = nm;
    }
  }

  var from = dari   ? new Date(dari   + 'T00:00:00') : null;
  var to   = sampai ? new Date(sampai + 'T23:59:59') : null;

  var jurnalData = shJurnal.getDataRange().getValues();
  var absenData  = shAbsen.getDataRange().getValues();

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

  var ss       = getSpreadsheet_();
  var shJurnal = ss.getSheetByName('JURNAL');
  var shAbsen  = ss.getSheetByName('ABSENSI');
  var shSiswa  = ss.getSheetByName('SISWA');

  if (!shJurnal || !shAbsen || !shSiswa) return [];

  // Guru aktif di sekolah ini
  var shUsers = ss.getSheetByName('USERS');
  var guruSet = new Set();
  if (shUsers) {
    var ud = shUsers.getDataRange().getValues();
    for (var ui = 1; ui < ud.length; ui++) {
      if (String(ud[ui][1] || '').toLowerCase() === 'admin' &&
          String(ud[ui][2] || '').toLowerCase() === 'active') {
        guruSet.add(String(ud[ui][0] || '').toLowerCase().trim());
      }
    }
  }

  // Build jurnal id → kelas (hanya guru sekolah ini)
  var jurnalData = shJurnal.getDataRange().getValues();
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
  var siswaData = shSiswa.getDataRange().getValues();
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
  var absenData   = shAbsen.getDataRange().getValues();
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
  return result;
}

/**
 * getRaportAbsensiSiswa()
 * Raport absensi per siswa (semua kelas, semua guru di sekolah ini).
 * Returns: array of {kelas, nis, nama, H, S, I, A, total, persen}
 */
function getRaportAbsensiSiswa() {
  assertKepsek_();

  var ss       = getSpreadsheet_();
  var shJurnal = ss.getSheetByName('JURNAL');
  var shAbsen  = ss.getSheetByName('ABSENSI');

  if (!shJurnal || !shAbsen) return [];

  var shUsers = ss.getSheetByName('USERS');
  var guruSet = new Set();
  if (shUsers) {
    var ud = shUsers.getDataRange().getValues();
    for (var ui = 1; ui < ud.length; ui++) {
      if (String(ud[ui][1] || '').toLowerCase() === 'admin' &&
          String(ud[ui][2] || '').toLowerCase() === 'active') {
        guruSet.add(String(ud[ui][0] || '').toLowerCase().trim());
      }
    }
  }

  var jurnalData = shJurnal.getDataRange().getValues();
  var jurnalMap  = {};
  for (var ji = 1; ji < jurnalData.length; ji++) {
    var jEmail = String(jurnalData[ji][12] || '').toLowerCase().trim();
    if (!guruSet.has(jEmail)) continue;
    jurnalMap[jurnalData[ji][0]] = String(jurnalData[ji][2] || '');
  }

  var absenData  = shAbsen.getDataRange().getValues();
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

  var ss       = getSpreadsheet_();
  var shJurnal = ss.getSheetByName('JURNAL');
  var shAbsen  = ss.getSheetByName('ABSENSI');

  if (!shJurnal) return { weekly: [], absence: {H:0,S:0,I:0,A:0}, labels: [] };

  var shUsers = ss.getSheetByName('USERS');
  var guruSet = new Set();
  if (shUsers) {
    var ud = shUsers.getDataRange().getValues();
    for (var ui = 1; ui < ud.length; ui++) {
      if (String(ud[ui][1] || '').toLowerCase() === 'admin' &&
          String(ud[ui][2] || '').toLowerCase() === 'active') {
        guruSet.add(String(ud[ui][0] || '').toLowerCase().trim());
      }
    }
  }

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

  var jurnalData = shJurnal.getDataRange().getValues();
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
  if (shAbsen) {
    var absenData = shAbsen.getDataRange().getValues();
    for (var ai = 1; ai < absenData.length; ai++) {
      if (!jurnalIds.has(String(absenData[ai][0]))) continue;
      var st = String(absenData[ai][3] || '').toUpperCase();
      if (absTotals[st] !== undefined) absTotals[st]++;
    }
  }

  return {
    labels : labels,
    weekly : weeks.map(function(w) { return w.count; }),
    absence: absTotals
  };
}

// Fitur AI Summary dihapus (quota Gemini API habis)
