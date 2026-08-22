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
 * KKTP (Kriteria Ketercapaian Tujuan Pembelajaran) PER FASE
 * ================================================================
 * Rentang ketuntasan resmi Kurikulum Merdeka, beda-beda tiap fase (jenjang
 * kelas) — dipakai di getRekapSekolah() untuk "Ketuntasan Siswa per Kelas"
 * supaya sesuai standar per fase, bukan satu KKM tunggal untuk semua
 * jenjang. nilai_akhir siswa (skala 0-100) diperlakukan sebagai persentase
 * ketercapaian tujuan pembelajaran.
 */
var KKTP_BANDS_A_ = [
  { min: 0,  max: 50,  status: 'Belum Mencapai Tujuan', tindak: 'Remedial total (bimbingan penuh dari guru dan keterlibatan orang tua).' },
  { min: 51, max: 70,  status: 'Belum Mencapai Tujuan', tindak: 'Remedial parsial pada indikator tujuan pembelajaran yang belum dikuasai.' },
  { min: 71, max: 85,  status: 'Sudah Mencapai Tujuan', tindak: 'Tuntas. Lanjutkan ke materi/tujuan pembelajaran berikutnya.' },
  { min: 86, max: 100, status: 'Sudah Mencapai Tujuan', tindak: 'Tuntas. Berikan aktivitas pengayaan atau tantangan belajar mandiri.' }
];
var KKTP_BANDS_BC_ = [
  { min: 0,  max: 45,  status: 'Belum Mencapai Tujuan', tindak: 'Remedial terstruktur pada seluruh bagian materi.' },
  { min: 46, max: 65,  status: 'Belum Mencapai Tujuan', tindak: 'Remedial mandiri melalui tugas tambahan atau tutor sebaya pada bagian yang kurang.' },
  { min: 66, max: 85,  status: 'Sudah Mencapai Tujuan', tindak: 'Tuntas. Siswa direkomendasikan mempertahankan konsistensi belajar.' },
  { min: 86, max: 100, status: 'Sudah Mencapai Tujuan', tindak: 'Tuntas. Sistem otomatis merekomendasikan materi pengayaan tingkat lanjut.' }
];
var KKTP_BANDS_D_ = [
  { min: 0,  max: 40,  status: 'Belum Mencapai Tujuan', tindak: 'Intervensi khusus (remedial terbimbing tatap muka dengan guru).' },
  { min: 41, max: 65,  status: 'Belum Mencapai Tujuan', tindak: 'Perbaikan portofolio atau revisi tugas pada indikator yang lemah.' },
  { min: 66, max: 85,  status: 'Sudah Mencapai Tujuan', tindak: 'Tuntas. Siswa siap menghadapi asesmen sumatif lingkup materi berikutnya.' },
  { min: 86, max: 100, status: 'Sudah Mencapai Tujuan', tindak: 'Tuntas. Diberikan peran sebagai tutor sebaya atau proyek eksplorasi mandiri.' }
];
var KKTP_BANDS_EF_ = [
  { min: 0,  max: 40,  status: 'Belum Mencapai Tujuan', tindak: 'Remedial komprehensif (teori dan praktik ulang).' },
  { min: 41, max: 70,  status: 'Belum Mencapai Tujuan', tindak: 'Remedial terfokus melalui pengerjaan ulang instrumen asesmen yang gagal.' },
  { min: 71, max: 88,  status: 'Sudah Mencapai Tujuan', tindak: 'Tuntas. Kompetensi esensial telah terpenuhi dengan baik.' },
  { min: 89, max: 100, status: 'Sudah Mencapai Tujuan', tindak: 'Tuntas. Siswa diarahkan ke pendalaman materi berbasis riset atau aplikasi projek nyata.' }
];

// Fase B & C dan Fase E & F masing-masing berbagi rentang persen yang
// sama persis (cuma beda label/kelompok kelas) — jadi kunci berbeda
// tapi menunjuk ke array band yang sama.
var KKTP_FASE_BANDS_ = {
  A: KKTP_BANDS_A_,
  B: KKTP_BANDS_BC_,
  C: KKTP_BANDS_BC_,
  D: KKTP_BANDS_D_,
  E: KKTP_BANDS_EF_,
  F: KKTP_BANDS_EF_
};

var KKTP_FASE_LABEL_ = {
  A: 'Fase A (Kelas 1–2 SD)',
  B: 'Fase B (Kelas 3–4 SD)',
  C: 'Fase C (Kelas 5–6 SD)',
  D: 'Fase D (Kelas 7–9 SMP)',
  E: 'Fase E (Kelas 10 SMA/SMK)',
  F: 'Fase F (Kelas 11–12 SMA/SMK)'
};

/**
 * Tentukan fase dari nama kelas (angka di depan, mis. "7A" -> 7).
 * Penomoran kelas di Indonesia unik per jenjang (SD 1-6, SMP 7-9,
 * SMA/SMK 10-12, tidak tumpang tindih) jadi angka kelas saja sudah
 * cukup — tidak perlu baca teks kode_sekolah segala. Kalau angka
 * kelasnya tidak jelas, fallback ke Fase D karena sistem ini paling
 * banyak dipakai di jenjang SMP.
 */
function _kepsekFaseDariKelas_(kelas) {
  var m = String(kelas || '').match(/(\d+)/);
  var n = m ? parseInt(m[1], 10) : null;
  if (n === null) return 'D';
  if (n <= 2) return 'A';
  if (n <= 4) return 'B';
  if (n <= 6) return 'C';
  if (n <= 9) return 'D';
  if (n <= 10) return 'E';
  return 'F';
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
 * Daftar email guru/admin berstatus aktif ATAU trial, dari sheet USERS
 * (selalu central) — guru trial sengaja ikut disertakan supaya Kepsek/
 * SuperAdmin bisa memantau guru sejak masa coba 30 hari, sebelum
 * diaktifkan penuh lewat "Aktifkan Penuh" di panel SuperAdmin.
 * Satu deployment ini bisa melayani BEBERAPA sekolah sekaligus — SuperAdmin
 * mengelompokkan tiap akun ke kode_sekolah lewat updateUserSekolah(). Baik
 * Kepsek MAUPUN SuperAdmin daftar gurunya otomatis DIPERSEMPIT ke guru
 * dengan kode_sekolah yang SAMA dengan akun pemanggil itu sendiri — kalau
 * pemanggilnya sendiri belum di-assign sekolah, hasilnya sengaja
 * dikosongkan (bukan malah menampilkan guru yang juga belum di-assign)
 * supaya tidak bocor data sekolah lain. SuperAdmin diperlakukan sama
 * seperti Kepsek di sini karena satu akun SuperAdmin sudah dipetakan ke
 * SATU kode_sekolah lewat panel SuperAdmin (kolom "Kode/Nama Sekolah") —
 * kalau kelak ada kebutuhan SuperAdmin memantau lebih dari satu sekolah
 * sekaligus, itu perlu fitur "pilih sekolah" terpisah, bukan default view
 * gabungan semua sekolah seperti sebelumnya.
 */
function _kepsekActiveGuruEmails_() {
  var auth  = getAuth();
  var shUsers = _ensureUsersKodeSekolahColumn_() || _getCentralSheetByName_('USERS');
  var guruEmails = [];
  if (!shUsers) return guruEmails;

  var userData = shUsers.getDataRange().getValues();

  var scopeToSchool = auth.role === 'kepsek' || auth.role === 'superadmin';
  var callerKode = '';
  if (scopeToSchool) {
    for (var ci = 1; ci < userData.length; ci++) {
      if (String(userData[ci][0] || '').toLowerCase().trim() === auth.email) {
        callerKode = String(userData[ci][4] || '').trim();
        break;
      }
    }
    if (!callerKode) return guruEmails; // Belum di-assign sekolah — kosongkan, jangan tebak
  }

  for (var ui = 1; ui < userData.length; ui++) {
    var uEmail = String(userData[ui][0] || '').toLowerCase().trim();
    var uRole  = String(userData[ui][1] || '').toLowerCase().trim();
    var uStatus= String(userData[ui][2] || '').toLowerCase().trim();
    if (!uEmail || !(uRole === 'admin' || uRole === 'guru') || (uStatus !== 'active' && uStatus !== 'trial')) continue;
    if (scopeToSchool && String(userData[ui][4] || '').trim() !== callerKode) continue;
    guruEmails.push(uEmail);
  }

  // SuperAdmin kadang juga mengajar sendiri (isi Jurnal/Nilai lewat akun
  // SuperAdmin-nya sendiri, bukan cuma mengelola sistem) — role-nya
  // 'superadmin' jadi tidak lolos filter role di atas. Sertakan email
  // SuperAdmin sendiri supaya data mengajarnya ikut terhitung di Rekap
  // Sekolah/Rekap Guru Wali miliknya sendiri maupun milik Kepsek di
  // sekolah yang sama. Kepsek TIDAK disertakan di sini karena perannya
  // read-only, tidak pernah mengisi Jurnal/Nilai sendiri.
  if (auth.role === 'superadmin' && guruEmails.indexOf(auth.email) === -1) {
    guruEmails.push(auth.email);
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
 * _kepsekSettingIdx_(headerRow)
 * Bangun index kolom SETTING dari HEADER ASLI — sama seperti getSetting()
 * di Setting.js — bukan indeks tetap (setAll[si][2] dst). Sheet SETTING
 * dulu dibaca dengan indeks tetap di sini (setAll[si][2] dianggap selalu
 * "tahun_pelajaran"), padahal urutan kolomnya tergantung urutan header
 * ASLI di tiap spreadsheet (bisa beda-beda hasil migrasi antar guru) —
 * ini akar penyebab "Tahun Pelajaran: -" kosong di Rekap Guru Wali
 * walau datanya sebenarnya ada.
 */
function _kepsekSettingIdx_(headerRow) {
  var header = (headerRow || []).map(function(h) { return String(h || '').toLowerCase().trim(); });
  var idx = {};
  header.forEach(function(h, i) { idx[h] = i; });
  return idx;
}

/**
 * _kepsekTahunFromRow_/_kepsekSemesterFromRow_
 * Prioritaskan kolom tahun_pelajaran_aktif/semester_aktif (nilai yang
 * benar-benar dipakai guru sekarang, lihat getSetting() di Setting.js)
 * — baru fallback ke kolom tahun/semester lama kalau kolom aktif itu
 * belum ada/kosong.
 */
function _kepsekTahunFromRow_(row, idx) {
  if (idx.tahun_pelajaran_aktif > -1 && row[idx.tahun_pelajaran_aktif]) return String(row[idx.tahun_pelajaran_aktif]);
  if (idx.tahun > -1) return String(row[idx.tahun] || '');
  return '';
}
function _kepsekSemesterFromRow_(row, idx) {
  if (idx.semester_aktif > -1 && row[idx.semester_aktif]) return String(row[idx.semester_aktif]);
  if (idx.semester > -1) return String(row[idx.semester] || '');
  return '';
}

/**
 * _kepsekAktifTahunSemester_(ssMap, guruEmails)
 * Tahun pelajaran & semester aktif sekolah ini (baris SETTING pertama
 * yang terisi punya guru sekolah ini), fallback ke Setting akun pemanggil
 * sendiri kalau belum ada guru yang mengisi. Dipakai buat menyaring data
 * absensi/nilai supaya hanya periode ajaran yang sedang berjalan yang
 * dihitung, bukan akumulasi dari tahun-tahun lama.
 */
function _kepsekAktifTahunSemester_(ssMap, guruEmails) {
  var setAll = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'SETTING');
  var sIdx = setAll.length ? _kepsekSettingIdx_(setAll[0]) : {};
  var tahun_pelajaran = '';
  var semester_aktif  = '';
  for (var si = 1; si < setAll.length; si++) {
    var rowTahun = _kepsekTahunFromRow_(setAll[si], sIdx);
    if (rowTahun) tahun_pelajaran = rowTahun;
    var rowSem = _kepsekSemesterFromRow_(setAll[si], sIdx);
    if (rowSem) { semester_aktif = rowSem; break; }
  }
  if (!tahun_pelajaran || !semester_aktif) {
    try {
      var ownSetting = getSetting();
      if (!tahun_pelajaran) tahun_pelajaran = ownSetting.tahun_pelajaran_aktif || ownSetting.tahun_pelajaran || '';
      if (!semester_aktif) semester_aktif = ownSetting.semester_aktif || ownSetting.semester || '';
    } catch (e) { /* getSetting gagal — biarkan tetap kosong */ }
  }
  return { tahun: tahun_pelajaran, semester: semester_aktif };
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
 *     totalSiswa, totalKelas, kelasList, mapelList, lastJurnal,
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

  // ── Setting sekolah (ambil baris pertama yang terisi) ──
  var sIdx = setAll.length ? _kepsekSettingIdx_(setAll[0]) : {};
  var sekolah         = '';
  var tahun_pelajaran = '';
  var semester_aktif  = '';
  for (var si = 1; si < setAll.length; si++) {
    if (sIdx.sekolah > -1 && setAll[si][sIdx.sekolah]) sekolah = String(setAll[si][sIdx.sekolah] || '');
    var rowTahun = _kepsekTahunFromRow_(setAll[si], sIdx);
    if (rowTahun) tahun_pelajaran = rowTahun;
    var rowSem = _kepsekSemesterFromRow_(setAll[si], sIdx);
    if (rowSem) { semester_aktif = rowSem; break; }
  }
  // Fallback: kalau belum ada guru yang punya baris SETTING terisi (mis.
  // sekolah baru/belum ada guru mengisi Setting), pakai tahun/semester
  // aktif milik akun pemanggil sendiri (Kepsek/SuperAdmin) supaya kartu
  // ini tidak nongol "-" walau periode ajaran sekolah sudah ditentukan.
  if (!tahun_pelajaran || !semester_aktif) {
    try {
      var ownSetting = getSetting();
      if (!tahun_pelajaran) tahun_pelajaran = ownSetting.tahun_pelajaran_aktif || ownSetting.tahun_pelajaran || '';
      if (!semester_aktif) semester_aktif = ownSetting.semester_aktif || ownSetting.semester || '';
    } catch (e) { /* getSetting gagal — biarkan tetap kosong */ }
  }

  // ── Baca setting per guru (nama_guru) ──
  var namaMap = {};
  for (var ni = 1; ni < setAll.length; ni++) {
    var nEmail = sIdx.email > -1 ? String(setAll[ni][sIdx.email] || '').toLowerCase().trim() : '';
    var nNama  = sIdx.guru > -1 ? String(setAll[ni][sIdx.guru] || '').trim() : '';
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

  // Kelas & mapel unik per guru dari jurnal (yang BENAR-BENAR diajar,
  // bukan sekadar didaftarkan di Setting) — dipakai utk kolom "Kelas
  // Diampu" & "Mata Pelajaran" di tabel Guru supaya Kepsek tahu guru
  // tsb mengajar kelas berapa saja dan mapel apa.
  var kelasPerGuru = {};
  var mapelPerGuru = {};
  for (var kji = 1; kji < jurnalRows.length; kji++) {
    var kEmail = String(jurnalRows[kji][12] || '').toLowerCase().trim();
    if (!guruMap[kEmail]) continue;
    if (!kelasPerGuru[kEmail]) kelasPerGuru[kEmail] = new Set();
    if (jurnalRows[kji][2]) kelasPerGuru[kEmail].add(jurnalRows[kji][2]);
    if (!mapelPerGuru[kEmail]) mapelPerGuru[kEmail] = new Set();
    if (jurnalRows[kji][19]) mapelPerGuru[kEmail].add(jurnalRows[kji][19]);
  }
  guruEmails.forEach(function(email) {
    guruMap[email].totalKelas = kelasPerGuru[email] ? kelasPerGuru[email].size : 0;
    guruMap[email].kelasList  = kelasPerGuru[email] ? Array.from(kelasPerGuru[email]).sort() : [];
    guruMap[email].mapelList  = mapelPerGuru[email] ? Array.from(mapelPerGuru[email]).sort() : [];
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
      kelasList        : g.kelasList,
      mapelList        : g.mapelList,
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

  // ── Ketuntasan siswa per kelas berdasarkan KKTP (rentang per Fase) ──
  // Skor per siswa = rata-rata (Rata Harian + PTS + SAS/ASAT) digabung
  // dari SEMUA mapel yang diampu guru-guru sekolah ini, lalu dibandingkan
  // ke rentang KKTP sesuai fase jenjang kelasnya (lihat KKTP_FASE_BANDS_
  // di atas) — bukan satu KKM tunggal, supaya sesuai standar Kurikulum
  // Merdeka yang berbeda tiap fase.
  var kelasKetuntasan = [];
  if (nilaiRows.length > 1) {
      var nh = nilaiRows[0].map(function(h) { return String(h || '').toLowerCase().trim(); });
      var nidx = {};
      nh.forEach(function(h, i) { nidx[h] = i; });

      // Nama siswa (kelas -> nis -> nama) utk daftar "Belum Mencapai Tujuan".
      var namaMapKtp = {};
      for (var nsi = 1; nsi < siswaRows.length; nsi++) {
        var kelasNs = String(siswaRows[nsi][0] || '');
        var nisNs   = String(siswaRows[nsi][2] || '').trim();
        var namaNs  = String(siswaRows[nsi][3] || '');
        if (!namaMapKtp[kelasNs]) namaMapKtp[kelasNs] = {};
        if (nisNs) namaMapKtp[kelasNs][nisNs] = namaNs;
      }

      // Ketuntasan KKTP dihitung dari rata-rata (Rata Harian + PTS +
      // SAS/ASAT) SAJA — bukan nilai_akhir/nilai_asli yang sudah ada
      // (itu formula campuran 40% [Harian+Tugas] + 30% PTS + 30% UAS).
      // Tugas Mandiri/Kelompok sengaja TIDAK ikut. Komponen yang kosong
      // dilewati, rata-rata dihitung dari yang tersedia saja.
      var _parseKtpNum_ = function(v) {
        if (v === '' || v === null || v === undefined) return null;
        var n = Number(v);
        return isNaN(n) ? null : n;
      };

      var classStudents = {};
      for (var ni = 1; ni < nilaiRows.length; ni++) {
        var rowTahun = String(nilaiRows[ni][nidx.tahun] || '').trim();
        var rowSem = String(nilaiRows[ni][nidx.semester] || '').toLowerCase().trim();
        if (tahun_pelajaran && rowTahun !== String(tahun_pelajaran)) continue;
        if (semester_aktif && rowSem !== String(semester_aktif).toLowerCase()) continue;

        var kelasNilai = String(nilaiRows[ni][nidx.kelas] || '').trim();
        var nisNilai = String(nilaiRows[ni][nidx.nis] || '').trim();
        if (!kelasNilai || !nisNilai) continue;

        var rataHarian = _parseKtpNum_(nilaiRows[ni][nidx.rata_uh]);
        var ptsVal     = _parseKtpNum_(nilaiRows[ni][nidx.pts]);
        var isGenapRow = rowSem.indexOf('genap') > -1;
        var uasVal     = _parseKtpNum_(nilaiRows[ni][isGenapRow ? nidx.pat : nidx.pas]);

        var komponen = [rataHarian, ptsVal, uasVal].filter(function(v) { return v !== null; });
        if (!komponen.length) continue;
        var skor = komponen.reduce(function(a, b) { return a + b; }, 0) / komponen.length;

        if (!classStudents[kelasNilai]) classStudents[kelasNilai] = {};
        if (!classStudents[kelasNilai][nisNilai]) classStudents[kelasNilai][nisNilai] = [];
        classStudents[kelasNilai][nisNilai].push(skor);
      }

      Object.keys(classStudents).forEach(function(kelas) {
        var siswaMap = classStudents[kelas];
        var nisList = Object.keys(siswaMap);
        if (!nisList.length) return;

        var fase = _kepsekFaseDariKelas_(kelas);
        var bandDefs = KKTP_FASE_BANDS_[fase];
        var siswaBelum = []; // {nis, nama, nilai} — nilai dipakai buat urutan, tidak ditampilkan

        var tuntasCount = 0;
        nisList.forEach(function(nis) {
          var avg = siswaMap[nis].reduce(function(a, b) { return a + b; }, 0) / siswaMap[nis].length;
          var band = bandDefs[bandDefs.length - 1];
          for (var bi = 0; bi < bandDefs.length; bi++) {
            if (avg >= bandDefs[bi].min && avg <= bandDefs[bi].max) { band = bandDefs[bi]; break; }
          }

          if (band.status === 'Sudah Mencapai Tujuan') {
            tuntasCount++;
          } else {
            siswaBelum.push({ nama: (namaMapKtp[kelas] && namaMapKtp[kelas][nis]) || nis, nilai: avg });
          }
        });

        siswaBelum.sort(function(a, b) { return a.nilai - b.nilai; });
        var persenTuntas = Math.round((tuntasCount / nisList.length) * 100);

        kelasKetuntasan.push({
          kelas               : kelas,
          faseLabel           : KKTP_FASE_LABEL_[fase],
          totalSiswa          : nisList.length,
          persenTuntas        : persenTuntas,
          persenPerluBimbingan: 100 - persenTuntas,
          siswaBelum          : siswaBelum.map(function(s) { return s.nama; })
        });
      });

      // Semua kelas yang punya data nilai ditampilkan (tidak dibatasi top-N)
      // — satu kode_sekolah bisa punya banyak kelas dari gabungan semua
      // guru/mapel yang pakai sistem ini, jadi tabelnya harus lengkap.
      kelasKetuntasan.sort(function(a, b) {
        return a.kelas < b.kelas ? -1 : (a.kelas > b.kelas ? 1 : 0);
      });
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

  var sIdx = setAll.length ? _kepsekSettingIdx_(setAll[0]) : {};
  var namaMap = {};
  var activeTahun = '';
  for (var si = 1; si < setAll.length; si++) {
    var nEmail = sIdx.email > -1 ? String(setAll[si][sIdx.email] || '').toLowerCase().trim() : '';
    var nNama  = sIdx.guru > -1 ? String(setAll[si][sIdx.guru] || '').trim() : '';
    if (nEmail && nNama) namaMap[nEmail] = nNama;
    if (!activeTahun) {
      var rowTahun = _kepsekTahunFromRow_(setAll[si], sIdx);
      if (rowTahun) activeTahun = rowTahun;
    }
  }
  // Fallback sama seperti getRekapSekolah(): kalau belum ada guru dengan
  // baris SETTING terisi, pakai tahun aktif milik akun pemanggil sendiri.
  if (!activeTahun) {
    try {
      var ownSettingGw = getSetting();
      activeTahun = ownSettingGw.tahun_pelajaran_aktif || ownSettingGw.tahun_pelajaran || '';
    } catch (e) { /* getSetting gagal — biarkan tetap kosong */ }
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
    if (setRows.length) {
      var dIdx = _kepsekSettingIdx_(setRows[0]);
      for (var si = 1; si < setRows.length; si++) {
        if (dIdx.email > -1 && String(setRows[si][dIdx.email] || '').toLowerCase().trim() === targetEmail) {
          activeTahun = _kepsekTahunFromRow_(setRows[si], dIdx);
          break;
        }
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
 * getEarlyWarningSiswa()
 * Top 3 siswa dengan Sakit+Izin+Alpa terbanyak di tiap kelas (gabungan
 * semua mapel/guru di sekolah ini), tampil otomatis begitu Rekap Sekolah
 * dibuka.
 * Returns: array of {kelas, siswa: [{nis, nama, S, I, A, totalSIA, persen, dominant}, ...]}
 */
function getEarlyWarningSiswa() {
  assertKepsek_();

  var cacheKey = 'KEPSEK_EARLY_WARNING';
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* cache miss/corrupt, lanjut hitung ulang */ }

  var siswaRows = _kepsekHitungAbsensiPerSiswa_();
  var perKelas  = {};
  siswaRows.forEach(function(d) {
    var totalSIA = d.S + d.I + d.A;
    if (totalSIA <= 0) return; // tidak ada yang perlu disorot
    if (!perKelas[d.kelas]) perKelas[d.kelas] = [];
    perKelas[d.kelas].push({
      nis: d.nis, nama: d.nama, S: d.S, I: d.I, A: d.A, totalSIA: totalSIA,
      persen: d.total > 0 ? Math.round(d.H / d.total * 100) : 0,
      dominant: _kepsekDominanSIA_(d.S, d.I, d.A)
    });
  });

  var result = Object.keys(perKelas).sort().map(function(kelas) {
    var siswa = perKelas[kelas]
      .sort(function(a, b) { return b.totalSIA - a.totalSIA; })
      .slice(0, 3);
    return { kelas: kelas, siswa: siswa };
  });

  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 180); } catch (e) { /* > 100KB, lewati cache */ }
  return result;
}

/**
 * _kepsekHitungAbsensiPerSiswa_()
 * Helper bersama: akumulasi H/S/I/A per siswa (semua kelas, semua guru
 * di sekolah ini) dari sheet JURNAL+ABSENSI+SISWA, dibatasi ke tahun
 * pelajaran & semester AKTIF saja (bukan akumulasi dari semester/tahun
 * lama). ABSENSI sendiri TIDAK punya kolom nama (cuma
 * jurnal_id/kelas/nis/status) — nama diambil dari sheet SISWA lewat
 * namaMap.
 *
 * PENTING: siswa bisa dapat beberapa jurnal DI HARI YANG SAMA (satu per
 * mapel/guru — bisa 3-4 mapel sehari), jadi absensi DIRINGKAS DULU per
 * (siswa, tanggal) sebelum diakumulasi — kalau dalam 1 hari statusnya
 * sama di beberapa mapel, dihitung 1 hari saja (bukan 3-4x). Kalau
 * beda hari, baru dihitung terpisah. Kalau dalam 1 hari ada mapel yang
 * beda status (jarang, tapi bisa), diambil yang paling parah (prioritas
 * A > I > S > H).
 *
 * Dipakai oleh getRaportAbsensiSiswa() (rekap per kelas) dan
 * _kepsekSiaTerbanyakPerKelas_() (siswa paling perlu perhatian).
 * Returns: array of {kelas, nis, nama, H, S, I, A, total}
 */
function _kepsekHitungAbsensiPerSiswa_() {
  var guruEmails = _kepsekActiveGuruEmails_();
  var ssMap      = _kepsekOpenSpreadsheetsMap_(guruEmails);
  var guruSet    = new Set(guruEmails);
  var periode    = _kepsekAktifTahunSemester_(ssMap, guruEmails);
  var tz         = Session.getScriptTimeZone();

  var jurnalData = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'JURNAL');
  var absenData  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'ABSENSI');
  var siswaData  = _kepsekReadSheetRowsMulti_(ssMap, guruEmails, 'SISWA');

  if (!jurnalData.length || !absenData.length) return [];

  // jurnal id -> {kelas, tanggal} — dibatasi guru sekolah ini DAN
  // tahun/semester aktif saja.
  var jurnalMap = {};
  for (var ji = 1; ji < jurnalData.length; ji++) {
    var jEmail = String(jurnalData[ji][12] || '').toLowerCase().trim();
    if (!guruSet.has(jEmail)) continue;
    var jTahun = String(jurnalData[ji][18] || '').trim();
    var jSem   = String(jurnalData[ji][13] || '').toLowerCase().trim();
    if (periode.tahun && jTahun !== String(periode.tahun)) continue;
    if (periode.semester && jSem !== String(periode.semester).toLowerCase()) continue;

    var tglRaw = jurnalData[ji][1];
    if (!tglRaw) continue;
    var tglKey = Utilities.formatDate(new Date(tglRaw), tz, 'yyyy-MM-dd');
    jurnalMap[jurnalData[ji][0]] = { kelas: String(jurnalData[ji][2] || ''), tanggal: tglKey };
  }

  var namaMap = {}; // kelas -> nis -> nama
  for (var si = 1; si < siswaData.length; si++) {
    var sOwner = String(siswaData[si][5] || '').toLowerCase().trim();
    if (!guruSet.has(sOwner)) continue;
    var kelasS = String(siswaData[si][0] || '');
    var nisS   = String(siswaData[si][2] || '');
    var namaS  = String(siswaData[si][3] || '');
    if (!namaMap[kelasS]) namaMap[kelasS] = {};
    if (nisS) namaMap[kelasS][nisS] = namaS;
  }

  // Ringkas dulu per (siswa, tanggal) — 1 hari = 1 status, prioritas
  // A > I > S > H kalau beda mapel beda status di hari yang sama.
  var STATUS_PRIORITAS_ = { A: 3, I: 2, S: 1, H: 0 };
  var perHari = {}; // key: kelas__nis__tanggal

  for (var ai = 1; ai < absenData.length; ai++) {
    var jId = absenData[ai][0];
    var j = jurnalMap[jId];
    if (!j) continue;
    var nisAb   = String(absenData[ai][2] || '').trim();
    var status  = String(absenData[ai][3] || '').toUpperCase();
    if (!nisAb || STATUS_PRIORITAS_[status] === undefined) continue;

    var hKey = j.kelas + '__' + nisAb + '__' + j.tanggal;
    if (!perHari[hKey] || STATUS_PRIORITAS_[status] > STATUS_PRIORITAS_[perHari[hKey].status]) {
      perHari[hKey] = { kelas: j.kelas, nis: nisAb, status: status };
    }
  }

  // Baru akumulasi per siswa dari ringkasan per-hari (1 hari = 1 hitungan).
  var siswaMap = {};
  Object.values(perHari).forEach(function(d) {
    var sKey = d.kelas + '__' + d.nis;
    if (!siswaMap[sKey]) {
      siswaMap[sKey] = {
        kelas: d.kelas, nis: d.nis,
        nama: (namaMap[d.kelas] && namaMap[d.kelas][d.nis]) || '',
        H:0, S:0, I:0, A:0, total:0
      };
    }
    siswaMap[sKey][d.status]++;
    siswaMap[sKey].total++;
  });

  return Object.values(siswaMap);
}

/**
 * _kepsekDominanSIA_(S, I, A)
 * Kategori S/I/A yang paling parah buat satu siswa — dipakai supaya Early
 * Warning cukup tampilkan 1 angka (mis. "I: 3") bukan tiga-tiganya
 * sekaligus. Kalau seri, prioritas A (Alpa) > I (Izin) > S (Sakit) karena
 * makin ke situ makin perlu perhatian.
 */
function _kepsekDominanSIA_(S, I, A) {
  var opts = [{ kategori: 'A', jumlah: A }, { kategori: 'I', jumlah: I }, { kategori: 'S', jumlah: S }];
  opts.sort(function(a, b) { return b.jumlah - a.jumlah; });
  return opts[0];
}

/**
 * getRaportAbsensiSiswa()
 * Rekap absensi UMUM per kelas (bukan per-siswa satu-satu — dulu daftar
 * seluruh siswa satu sekolah, terlalu panjang & duplikat rapor Dapodik).
 * Returns: array of {kelas, totalSiswa, H, S, I, A, total, persenHadir}
 */
function getRaportAbsensiSiswa() {
  assertKepsek_();

  var cacheKey = 'KEPSEK_RAPORT_ABSENSI_KELAS';
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* cache miss/corrupt, lanjut hitung ulang */ }

  var siswaRows = _kepsekHitungAbsensiPerSiswa_();
  if (!siswaRows.length) return [];

  var perKelas = {};
  siswaRows.forEach(function(d) {
    if (!perKelas[d.kelas]) {
      perKelas[d.kelas] = { kelas: d.kelas, totalSiswa: 0, H:0, S:0, I:0, A:0, total:0 };
    }
    var k = perKelas[d.kelas];
    k.totalSiswa++;
    k.H += d.H; k.S += d.S; k.I += d.I; k.A += d.A; k.total += d.total;
  });

  var result = Object.values(perKelas).map(function(k) {
    return {
      kelas       : k.kelas,
      totalSiswa  : k.totalSiswa,
      H: k.H, S: k.S, I: k.I, A: k.A,
      total       : k.total,
      persenHadir : k.total > 0 ? Math.round(k.H / k.total * 100) : 0
    };
  });
  result.sort(function(a, b) { return a.kelas < b.kelas ? -1 : (a.kelas > b.kelas ? 1 : 0); });

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
