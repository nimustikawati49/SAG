/**
 * DriveConnect.js — Migrasi Database Guru ke Drive Pribadi
 *
 * Di mode storage per_guru, spreadsheet data tiap guru memang SATU FILE
 * TERPISAH per guru (data logically dipisah), tapi file itu dibuat lewat
 * SpreadsheetApp.create() yang berjalan di identitas SuperAdmin (script
 * owner via deployment "Execute as: Me") — jadi semua file itu tetap
 * memakai KUOTA Drive SuperAdmin, bukan Drive masing-masing guru (lihat
 * komentar _autoProvisionUserSpreadsheet_ di Code.js).
 *
 * File ini memberi jalan keluar TANPA perlu ubah pengaturan deployment
 * (Execute as) atau Domain-Wide Delegation (yang butuh akses Admin
 * Console Google Workspace, biasanya di luar jangkauan SuperAdmin
 * aplikasi ini): guru membuat SENDIRI 1 spreadsheet kosong di Drive-nya,
 * membagikannya (share) ke email SuperAdmin sebagai Editor, lalu
 * menghubungkannya lewat connectOwnSpreadsheet(). Karena guru sendiri
 * yang membuat file itu, dia tetap PEMILIK file itu dan kuota Drive-nya
 * sendiri yang terpakai — share ke SuperAdmin cuma memberi AKSES (perlu
 * supaya script yang berjalan sebagai SuperAdmin tetap bisa baca/tulis,
 * termasuk dari trigger backup/reminder terjadwal di Trigger.js), BUKAN
 * memindahkan kepemilikan/kuota.
 *
 * Data lama (kalau ada, mis. dari masa trial yang sudah lebih dulu
 * auto-provisioned di Drive SuperAdmin) otomatis disalin ke spreadsheet
 * baru ini sekali saat pertama connect, supaya tidak ada data hilang.
 *
 * ⚠️ PENTING UNTUK SUPERADMIN — GURU "GRANDFATHERED" & SUPERADMIN SENDIRI:
 * fitur ini JUGA menangani guru yang datanya masih nyangkut di
 * spreadsheet CENTRAL (didaftarkan sebelum mode per_guru aktif, lihat
 * _hasExistingCentralTeacherData_ di SuperAdmin.js) — dan SuperAdmin
 * sendiri, yang sebelumnya SELALU dipetakan ke central biar bisa lihat
 * data guru lain (lihat cabang khusus SuperAdmin di
 * _autoProvisionUserSpreadsheet_, Code.js). Untuk kasus ini, migrasinya
 * TIDAK menyalin/menghapus seisi sheet central (dipakai bersama banyak
 * akun!) — cuma baris milik email yang bersangkutan, disaring lewat
 * kolom pemilik (owner_email/email_guru/email/guru, ATAU kolom pemilik
 * tetap SISWA — lihat _migrateLegacyCentralDataForUser_ di bawah).
 *
 * MasterSiswa/RiwayatKelas TIDAK dimigrasi langsung (sheet itu memang
 * TIDAK punya kolom pemilik sama sekali) — begitu SISWA guru ini selesai
 * dipindah, connectOwnSpreadsheet() memicu ensureAcademicSchema_() di
 * spreadsheet BARU-nya, yang otomatis MEMBANGUN ULANG MasterSiswa/
 * RiwayatKelas miliknya sendiri dari SISWA yang baru dipindah itu (lewat
 * migrateLegacyData_() di AcademicYear.js — fungsi yang sama yang dulu
 * membuat baris-baris itu di central). MasterSiswa/RiwayatKelas/
 * MasterTahunPelajaran di CENTRAL sendiri baru jadi benar-benar redundan
 * (aman dibersihkan) setelah SISWA central kosong sepenuhnya (semua guru
 * grandfathered + SuperAdmin sudah pindah) — dicek lewat
 * getCentralMasterRosterCleanupStatus()/clearRedundantCentralMasterRoster()
 * di bawah, harus dikonfirmasi manual oleh SuperAdmin, tidak otomatis.
 *
 * ⚠️ PENTING UNTUK SUPERADMIN — pembersihan spreadsheet lama OTOMATIS:
 * begitu data berhasil disalin, connectOwnSpreadsheet() MEMVERIFIKASI dulu
 * (jumlah sheet, baris, & kolom di spreadsheet baru harus persis cocok
 * dengan yang lama) — kalau cocok, spreadsheet LAMA (bekas auto-provision)
 * langsung DIPINDAH KE TRASH Drive SuperAdmin secara otomatis, tidak
 * menunggu ditinjau manual. Ini AMAN karena Trash Google punya masa
 * retensi standar (biasanya ~30 hari) sebelum benar-benar terhapus
 * permanen — kalau ternyata ada yang keliru, masih bisa dipulihkan lewat
 * Trash di Drive akun SuperAdmin sendiri selama masih dalam masa itu.
 * Kalau verifikasi TIDAK cocok (jarang — mis. ada sheet yang gagal
 * tersalin), spreadsheet lama TIDAK ditrash otomatis — cuma ditandai
 * prefix "[SUDAH PINDAH]" dan masuk antrean panel SuperAdmin "🗑️
 * Spreadsheet Lama (Sudah Dipindah Guru)" untuk ditinjau & dihapus manual
 * (lihat getMigratedOldSpreadsheets/trashOldMigratedSpreadsheet di bawah).
 * Jadi panel itu SEHARUSNYA jarang/hampir selalu kosong — kalau ada isinya,
 * itu tandanya ada migrasi yang perlu dicek lebih teliti, bukan hal rutin.
 */

/**
 * _isDriveSelfOwned_(email)
 * True kalau spreadsheet data guru ini SUDAH dimiliki (owner_email di
 * RESOURCE_MAP) oleh guru itu sendiri, bukan lagi hasil auto-provisioning
 * di Drive SuperAdmin. Dipakai assertLicenseActive() (Auth.js) untuk
 * mewajibkan koneksi Drive pribadi begitu akun guru diaktifkan penuh.
 *
 * CATATAN: owner_email cocok SAJA tidak cukup — akun SuperAdmin sendiri
 * punya entry RESOURCE_MAP dengan owner_email = email-nya sendiri TAPI
 * resource_id-nya tetap spreadsheet CENTRAL (lihat cabang khusus
 * SuperAdmin di _autoProvisionUserSpreadsheet_, Code.js). Makanya di sini
 * juga dipastikan resource_id-nya BUKAN central, supaya SuperAdmin yang
 * belum connect Drive pribadi tidak salah kedeteksi "sudah terhubung".
 */
function _isDriveSelfOwned_(email) {
  var targetEmail = String(email || '').toLowerCase().trim();
  if (!targetEmail) return false;
  var entry = _getResourceMapEntryForUser_(targetEmail, 'data_spreadsheet');
  if (!entry) return false;
  if (String(entry.owner_email || '').toLowerCase().trim() !== targetEmail) return false;
  try {
    if (String(entry.resource_id || '').trim() === getCentralSpreadsheet_().getId()) return false;
  } catch (e) {}
  return true;
}

/**
 * _extractSpreadsheetId_(urlOrId)
 * Terima link lengkap Google Sheets ATAU ID mentah, kembalikan ID-nya.
 */
function _extractSpreadsheetId_(urlOrId) {
  var s = String(urlOrId || '').trim();
  if (!s) return '';
  var m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m && m[1]) return m[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(s)) return s;
  return '';
}

/**
 * getDriveConnectInfo()
 * Status koneksi Drive pribadi untuk guru yang sedang login — dipakai
 * kartu "Google Drive Pribadi Saya" di halaman Pengaturan.
 */
function getDriveConnectInfo() {
  var auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin') return { applicable: false };

  var perGuru = getStorageMode_() === 'per_guru';
  var connected = false;
  try { connected = perGuru && _isDriveSelfOwned_(auth.email); } catch (e) {}

  return {
    applicable: perGuru,
    connected: connected,
    required: !!(perGuru && auth.status === 'active' && !connected),
    shareWithEmail: getSuperAdminEmail_(),
    myEmail: auth.email
  };
}

/**
 * _getResourceMapEntryOwnedByOther_(spreadsheetId, myEmail)
 * Cegah dua guru berbeda menghubungkan spreadsheet yang sama.
 */
function _getResourceMapEntryOwnedByOther_(spreadsheetId, myEmail) {
  try {
    var sh = _getCentralSheetByName_('RESOURCE_MAP');
    if (!sh || sh.getLastRow() < 2) return null;
    var rows = sh.getDataRange().getValues();
    var header = rows[0].map(function (h) { return String(h || '').toLowerCase().trim(); });
    var idx = {};
    header.forEach(function (h, i) { idx[h] = i; });
    var target = String(myEmail || '').toLowerCase().trim();
    for (var i = 1; i < rows.length; i++) {
      var rowType = String(rows[i][idx.resource_type] || '').toLowerCase().trim();
      var rowId = String(rows[i][idx.resource_id] || '').trim();
      var rowStatus = String(rows[i][idx.status] || 'active').toLowerCase().trim();
      var rowEmail = String(rows[i][idx.email_guru] || '').toLowerCase().trim();
      if (rowType !== 'data_spreadsheet') continue;
      if (rowId !== spreadsheetId) continue;
      if (rowStatus !== 'active') continue;
      if (rowEmail === target) continue;
      return rowEmail;
    }
  } catch (e) {}
  return null;
}

/**
 * _LEGACY_CENTRAL_CANDIDATE_SHEETS_
 * Sheet non-central-only yang MUNGKIN masih menyimpan baris guru
 * "grandfathered" (atau SuperAdmin sendiri) langsung di central — dicek
 * satu-satu oleh _migrateLegacyCentralDataForUser_. Sengaja TIDAK termasuk
 * MasterSiswa/RiwayatKelas/MasterTahunPelajaran — sheet-sheet itu TIDAK
 * PUNYA kolom pemilik sama sekali. Tapi bukan berarti tidak tertangani:
 * begitu SISWA (yang PUNYA kolom pemilik, lihat _LEGACY_CENTRAL_FIXED_OWNER_COLUMN_
 * di bawah) berhasil dimigrasi, connectOwnSpreadsheet() memicu
 * ensureAcademicSchema_() di spreadsheet BARU guru itu — yang otomatis
 * MEMBANGUN ULANG MasterSiswa/RiwayatKelas miliknya sendiri dari SISWA
 * yang baru dipindah (lihat migrateLegacyData_() di AcademicYear.js, ini
 * memang fungsi yang sama yang tadinya membuat baris-baris itu di
 * central). Jadi tidak perlu logika pemisahan terpisah untuk 2 sheet itu.
 */
var _LEGACY_CENTRAL_CANDIDATE_SHEETS_ = [
  'JURNAL', 'ABSENSI', 'SISWA', 'NILAI_SISWA', 'SETTING_NILAI',
  'JADWAL_SEMESTER', 'GuruMengajar', 'SETTING', 'ModulAjar', 'Relasi_Modul_Jadwal'
];

/**
 * _LEGACY_CENTRAL_FIXED_OWNER_COLUMN_
 * Sheet SISWA (roster legacy) SEBENARNYA punya kolom pemilik — index
 * tetap 5 ("owner guru", lihat komentar SISWA_TAHUN_COL_ di Jurnal.js) —
 * tapi sheet ini dibuat manual dulu (bukan lewat kode), jadi teks header
 * barisnya tidak selalu cocok dengan kandidat generik di
 * _findOwnerColumnIndex_ (SuperAdmin.js), yang mencari berdasarkan TEKS
 * header, bukan posisi. Dipetakan manual di sini supaya SISWA tetap bisa
 * ikut dimigrasi otomatis per-guru seperti sheet lain.
 */
var _LEGACY_CENTRAL_FIXED_OWNER_COLUMN_ = {
  'SISWA': 5
};

/**
 * _migrateLegacyCentralDataForUser_(email, newSs)
 * Bagian dari connectOwnSpreadsheet() — menangani guru "grandfathered"
 * yang datanya masih nyangkut di spreadsheet CENTRAL (didaftarkan sebelum
 * mode per_guru aktif, lihat _hasExistingCentralTeacherData_ di
 * SuperAdmin.js) ATAU SuperAdmin sendiri (yang sebelumnya selalu
 * dipetakan ke central). Cuma menyasar sheet yang PUNYA kolom pemilik
 * (owner_email/email_guru/email/guru — lihat _findOwnerColumnIndex_ di
 * SuperAdmin.js, dipakai juga oleh getCentralDataSizeReport supaya
 * konsisten — atau kolom tetap khusus, lihat
 * _LEGACY_CENTRAL_FIXED_OWNER_COLUMN_). Untuk tiap sheet: salin HANYA
 * baris milik email ini ke spreadsheet baru, verifikasi jumlah baris yang
 * berhasil ditambahkan cocok, baru hapus baris ASLINYA (satu-satu, dari
 * bawah ke atas supaya nomor baris tidak bergeser) dari central — baris
 * guru LAIN di sheet yang sama sengaja tidak tersentuh sama sekali.
 */
function _migrateLegacyCentralDataForUser_(email, newSs) {
  var centralSs = getCentralSpreadsheet_();
  var migrated = [];

  _LEGACY_CENTRAL_CANDIDATE_SHEETS_.forEach(function (name) {
    try {
      var centralSh = centralSs.getSheetByName(name);
      if (!centralSh || centralSh.getLastRow() < 2) return;

      var lastCol = centralSh.getLastColumn();
      var header = centralSh.getRange(1, 1, 1, lastCol).getValues()[0];
      var ownerIdx = _LEGACY_CENTRAL_FIXED_OWNER_COLUMN_.hasOwnProperty(name)
        ? _LEGACY_CENTRAL_FIXED_OWNER_COLUMN_[name]
        : _findOwnerColumnIndex_(header);
      if (ownerIdx === -1 || ownerIdx >= lastCol) return; // tidak ada kolom pemilik, lewati

      var allValues = centralSh.getRange(2, 1, centralSh.getLastRow() - 1, lastCol).getValues();
      var myRowIdxs = [];
      allValues.forEach(function (row, i) {
        if (String(row[ownerIdx] || '').toLowerCase().trim() === email) myRowIdxs.push(i);
      });
      if (!myRowIdxs.length) return;

      var targetSh = newSs.getSheetByName(name);
      if (!targetSh) {
        targetSh = newSs.insertSheet(name);
        targetSh.getRange(1, 1, 1, lastCol).setValues([header]);
        targetSh.setFrozenRows(1);
      }

      var rowsToCopy = myRowIdxs.map(function (i) { return allValues[i]; });
      var beforeCount = targetSh.getLastRow();
      targetSh.getRange(beforeCount + 1, 1, rowsToCopy.length, lastCol).setValues(rowsToCopy);
      var actuallyAdded = targetSh.getLastRow() - beforeCount;

      // Verifikasi dulu sebelum hapus dari central — kalau jumlah yang
      // benar-benar tertulis tidak cocok dengan yang seharusnya, JANGAN
      // hapus apa pun dari central (biar tetap aman, baris aslinya masih
      // utuh, bisa dicoba lagi lain waktu).
      if (actuallyAdded !== rowsToCopy.length) return;

      myRowIdxs.slice().reverse().forEach(function (i) {
        centralSh.deleteRow(i + 2); // +1 header, +1 karena 1-indexed
      });

      migrated.push(name + ' (' + rowsToCopy.length + ' baris dari central)');
    } catch (eSheet) {
      try { logError_('MIGRATE_LEGACY_CENTRAL_DATA', eSheet); } catch (e2) {}
    }
  });

  return migrated;
}

/**
 * connectOwnSpreadsheet(urlOrId)
 * Guru (atau SuperAdmin sendiri, kalau dia juga mengajar) menghubungkan
 * spreadsheet MILIK SENDIRI (sudah dibuat + dishare Editor ke SuperAdmin)
 * sebagai database pribadinya. Data lama otomatis disalin sekali ke
 * spreadsheet baru ini lewat DUA jalur berbeda tergantung dari mana asal
 * data lamanya:
 *  1. Spreadsheet per-guru EKSKLUSIF (hasil auto-provision, cuma dipakai
 *     satu orang) → disalin UTUH sheet demi sheet (lihat blok di bawah).
 *  2. Baris yang masih nyangkut di spreadsheet CENTRAL — dari guru
 *     "grandfathered" (sudah punya data central sebelum mode per_guru
 *     aktif) ATAU dari SuperAdmin sendiri (yang sebelumnya SELALU
 *     dipetakan ke central, lihat _autoProvisionUserSpreadsheet_ di
 *     Code.js) → disalin PER-BARIS lewat _migrateLegacyCentralDataForUser_
 *     di bawah, karena central dipakai BERSAMA banyak akun sekaligus,
 *     tidak boleh disalin/dihapus utuh seperti jalur 1.
 */
function connectOwnSpreadsheet(urlOrId) {
  var auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin') throw new Error('Fitur ini khusus akun guru/SuperAdmin.');
  if (getStorageMode_() !== 'per_guru') throw new Error('Mode storage aplikasi saat ini bukan per-guru, tidak perlu menghubungkan spreadsheet.');

  var email = auth.email;
  var id = _extractSpreadsheetId_(urlOrId);
  if (!id) throw new Error('Link atau ID spreadsheet tidak valid. Tempel link lengkap dari address bar Google Sheets.');

  var saEmail = getSuperAdminEmail_();
  var newSs;
  try {
    newSs = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Spreadsheet tidak ditemukan, atau belum dibagikan (Share) ke ' + saEmail + ' sebagai Editor.');
  }

  var editors = [];
  try { editors = newSs.getEditors().map(function (u) { return u.getEmail().toLowerCase().trim(); }); } catch (e) {}
  if (editors.indexOf(saEmail) === -1) {
    throw new Error('Spreadsheet ini belum dibagikan (Share) ke ' + saEmail + ' sebagai Editor. Bagikan dulu, lalu coba lagi.');
  }

  var collision = _getResourceMapEntryOwnedByOther_(id, email);
  if (collision) {
    throw new Error('Spreadsheet ini sudah terhubung ke akun guru lain (' + collision + '). Gunakan spreadsheet lain.');
  }

  var centralId = getCentralSpreadsheet_().getId();
  var oldId = resolveSpreadsheetIdForUser_(email);
  var migratedSheets = [];
  var autoTrashed = false;
  // Jalur salin-utuh di bawah ini HANYA aman untuk spreadsheet yang
  // EKSKLUSIF milik satu guru (hasil auto-provision per_guru) — bukan
  // untuk central, yang dipakai bersama banyak akun (mis. SuperAdmin
  // selalu dipetakan ke central). Kalau oldId ternyata central, migrasinya
  // WAJIB lewat _migrateLegacyCentralDataForUser_ (filter per-email) di
  // bawah, bukan disalin/di-trash utuh — makanya oldId === centralId
  // sengaja dikecualikan di sini.
  if (oldId && oldId !== id && oldId !== centralId) {
    var defaultNames = ['Sheet1', 'Lembar1'];
    var alreadyHasData = newSs.getSheets().some(function (sh) {
      return defaultNames.indexOf(sh.getName()) === -1;
    });

    var oldSs = null;
    try { oldSs = SpreadsheetApp.openById(oldId); } catch (e) {}

    if (!alreadyHasData && oldSs) {
      var candidateNames = [];
      oldSs.getSheets().forEach(function (sh) {
        var name = sh.getName();
        if (_isCentralOnlySheet_(name)) return;
        candidateNames.push(name);
        if (newSs.getSheetByName(name)) return; // sudah ada, tidak perlu disalin ulang
        try {
          var copied = sh.copyTo(newSs);
          copied.setName(name);
          migratedSheets.push(name);
        } catch (e2) {}
      });
      try {
        var defaultSh = newSs.getSheetByName('Sheet1') || newSs.getSheetByName('Lembar1');
        if (defaultSh && newSs.getSheets().length > 1) newSs.deleteSheet(defaultSh);
      } catch (e3) {}

      // Verifikasi sebelum boleh auto-trash: SEMUA sheet non-central dari
      // spreadsheet lama harus ada di spreadsheet baru dengan jumlah baris
      // & kolom PERSIS SAMA. Kalau ada yang tidak cocok (mis. copyTo() di
      // atas gagal diam-diam kena try/catch), JANGAN auto-trash — supaya
      // tidak berisiko membuang spreadsheet yang datanya belum benar-benar
      // tersalin sempurna. Fallback-nya turun ke jalur "tandai + tinjau
      // manual" di bawah, sama seperti sebelum auto-trash ini ada.
      var verified = candidateNames.length > 0 && candidateNames.every(function (name) {
        var oldSh = oldSs.getSheetByName(name);
        var newSh = newSs.getSheetByName(name);
        return oldSh && newSh &&
          oldSh.getLastRow() === newSh.getLastRow() &&
          oldSh.getLastColumn() === newSh.getLastColumn();
      });

      if (verified) {
        try {
          DriveApp.getFileById(oldId).setTrashed(true);
          autoTrashed = true;
        } catch (eTrash) {}
      }
    }

    if (!autoTrashed) {
      // Spreadsheet lama TIDAK bisa dipastikan aman untuk dibuang otomatis
      // (verifikasi di atas gagal / gagal dibuka / spreadsheet baru sudah
      // ada isinya sebelum connect, jadi tidak ada yang disalin sama
      // sekali) — sengaja dibiarkan, cuma ditandai prefix nama supaya
      // gampang dikenali, lalu masuk antrean review manual SuperAdmin
      // lewat panel "Spreadsheet Lama (Sudah Dipindah Guru)" (lihat
      // getMigratedOldSpreadsheets/trashOldMigratedSpreadsheet di bawah).
      try {
        var oldFile = DriveApp.getFileById(oldId);
        var oldName = oldFile.getName();
        if (oldName.indexOf('[SUDAH PINDAH]') !== 0) {
          oldFile.setName('[SUDAH PINDAH] ' + oldName);
        }
      } catch (eRename) {}
    }
  }

  // Baris milik email ini yang masih nyangkut di CENTRAL — dari guru
  // grandfathered ATAU SuperAdmin sendiri. Berjalan TANPA syarat oldId di
  // atas (guru grandfathered biasanya tidak punya entry RESOURCE_MAP sama
  // sekali, jadi oldId kosong), difilter per-baris, tidak menyalin/menyentuh
  // baris akun lain di sheet central yang sama.
  var legacyCentralMigrated = _migrateLegacyCentralDataForUser_(email, newSs);
  migratedSheets = migratedSheets.concat(legacyCentralMigrated);

  var oldEntry = _getResourceMapEntryForUser_(email, 'data_spreadsheet');
  if (oldEntry && oldEntry.id) {
    // PENTING: kalau resource_id entry lama ini adalah spreadsheet CENTRAL
    // (kasus SuperAdmin, yang sebelumnya selalu dipetakan ke central),
    // JANGAN PERNAH ditandai 'migrated_to_own_drive'/'trashed_auto' — dua
    // status itu dibaca getMigratedOldSpreadsheets()/trashOldMigratedSpreadsheet()
    // sebagai "spreadsheet ini boleh ditinjau lalu DIHAPUS", dan central
    // TIDAK BOLEH PERNAH masuk kategori itu (bisa menghapus seluruh data
    // sekolah). Dipakai status khusus yang tidak pernah dibaca fungsi trash
    // mana pun — central sendiri tidak disentuh sama sekali, cuma mapping
    // lamanya yang diganti.
    var oldEntryIsCentral = (String(oldEntry.resource_id || '').trim() === centralId);
    try {
      _upsertResourceMapEntry_({
        id: oldEntry.id,
        deployment_id: oldEntry.deployment_id,
        email_guru: email,
        resource_type: 'data_spreadsheet',
        resource_id: oldEntry.resource_id,
        resource_name: oldEntry.resource_name,
        owner_email: oldEntry.owner_email,
        status: oldEntryIsCentral ? 'central_mapping_replaced' : (autoTrashed ? 'trashed_auto' : 'migrated_to_own_drive'),
        catatan: (oldEntryIsCentral
          ? 'Mapping central digantikan spreadsheet pribadi guru pada '
          : (autoTrashed
            ? 'Otomatis dipindah ke Trash SuperAdmin setelah migrasi terverifikasi cocok pada '
            : 'Digantikan spreadsheet pribadi guru pada ')) + new Date().toISOString()
      });
    } catch (e) {}
  }

  _upsertResourceMapEntry_({
    email_guru: email,
    resource_type: 'data_spreadsheet',
    resource_id: id,
    resource_name: newSs.getName(),
    owner_email: email,
    status: 'active',
    catatan: 'Dihubungkan guru sendiri pada ' + new Date().toISOString() +
      (migratedSheets.length ? (' | disalin: ' + migratedSheets.join(', ')) : ' | tidak ada data lama untuk disalin')
  });

  // Bersihkan cache schema/migrasi akademik SEBELUM memicu ulang di bawah
  // — cache key-nya per-email (bukan per-spreadsheet, lihat
  // ensureAcademicSchema_/migrateLegacyData_ di AcademicYear.js), jadi
  // kalau tidak dibersihkan dulu, sistem bisa menganggap "sudah pernah
  // jalan" dari sesi sebelumnya (saat masih di spreadsheet/central lama)
  // dan TIDAK regenerasi MasterSiswa/RiwayatKelas dari SISWA yang baru
  // saja dipindah.
  try {
    var emailSlug = email.replace(/[^a-z0-9]/gi, '_');
    CacheService.getScriptCache().remove('ACADEMIC_SCHEMA_READY_' + emailSlug);
    CacheService.getScriptCache().remove('ACADEMIC_MIGRATED_' + emailSlug);
  } catch (eCacheClear) {}

  // Picu ulang ensureAcademicSchema_() SEKARANG (bukan menunggu interaksi
  // guru berikutnya) — resolve otomatis ke spreadsheet BARU (RESOURCE_MAP
  // sudah diperbarui di atas), dan migrateLegacyData_() di dalamnya akan
  // membangun ulang MasterSiswa/RiwayatKelas milik guru ini dari SISWA
  // yang baru saja dipindah (lihat AcademicYear.js) — jadi TIDAK perlu
  // logika pemisahan terpisah untuk 2 sheet tanpa kolom pemilik itu.
  try {
    if (typeof ensureAcademicSchema_ === 'function') ensureAcademicSchema_();
  } catch (eSchema) {
    try { logError_('CONNECT_OWN_SPREADSHEET_REBUILD_ROSTER', eSchema); } catch (e2) {}
  }

  invalidateCache_('SETTING');
  invalidateCache_('JURNAL');
  invalidateDashboardCache_();
  if (typeof _invalidateAuthCache_ === 'function') _invalidateAuthCache_(email);
  logAudit('CONNECT_OWN_SPREADSHEET', email, id +
    (migratedSheets.length ? (' | migrated: ' + migratedSheets.join(',')) : ' | no migration') +
    (autoTrashed ? ' | old spreadsheet auto-trashed' : ''));

  return { success: true, spreadsheetId: id, spreadsheetName: newSs.getName(), migratedSheets: migratedSheets, autoTrashed: autoTrashed };
}

/**
 * getMigratedOldSpreadsheets()
 * SuperAdmin only — daftar spreadsheet BEKAS auto-provisioning yang sudah
 * digantikan guru dengan spreadsheet pribadinya sendiri (lihat
 * connectOwnSpreadsheet di atas). File-file ini masih fisik ada di Drive
 * SuperAdmin (sengaja tidak dihapus otomatis saat migrasi) — panel ini
 * dipakai SuperAdmin untuk meninjau lalu menghapusnya manual satu-satu
 * lewat trashOldMigratedSpreadsheet() begitu yakin datanya sudah aman di
 * spreadsheet baru guru.
 */
function getMigratedOldSpreadsheets() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var centralId = getCentralSpreadsheet_().getId();
  var sh = _ensureResourceMapSheet_();
  var rows = sh.getDataRange().getValues();
  var idx = _getHeaderIndexMap_(sh);
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][idx.id] || '').trim();
    if (!id) continue;
    var type = String(rows[i][idx.resource_type] || '').toLowerCase().trim();
    var status = String(rows[i][idx.status] || '').toLowerCase().trim();
    if (type !== 'data_spreadsheet' || status !== 'migrated_to_own_drive') continue;

    var resourceId = String(rows[i][idx.resource_id] || '');
    // Jaga-jaga (defense-in-depth): spreadsheet CENTRAL tidak boleh pernah
    // muncul di panel ini apapun yang terjadi — status 'migrated_to_own_drive'
    // seharusnya memang tidak pernah dipasang untuk resource_id central
    // (lihat connectOwnSpreadsheet), tapi dicek ulang di sini juga.
    if (resourceId === centralId) continue;
    var sizeMB = null;
    try {
      var f = DriveApp.getFileById(resourceId);
      sizeMB = Math.round((f.getSize() / (1024 * 1024)) * 100) / 100;
    } catch (e) {}

    result.push({
      id: id,
      email_guru: String(rows[i][idx.email_guru] || ''),
      resource_id: resourceId,
      resource_name: String(rows[i][idx.resource_name] || ''),
      updated_at: String(rows[i][idx.updated_at] || ''),
      catatan: String(rows[i][idx.catatan] || ''),
      url: 'https://docs.google.com/spreadsheets/d/' + resourceId,
      sizeMB: sizeMB
    });
  }
  result.sort(function (a, b) { return String(b.updated_at || '').localeCompare(String(a.updated_at || '')); });
  return result;
}

/**
 * trashOldMigratedSpreadsheet(resourceMapId)
 * SuperAdmin only — pindahkan SATU spreadsheet lama (bekas auto-provision,
 * sudah digantikan guru dengan Drive pribadinya) ke Trash. Masih bisa
 * dipulihkan lewat Trash Drive SuperAdmin (retensi standar Google, biasa
 * ~30 hari) kalau ternyata ada yang keliru. Dibatasi cuma boleh menyasar
 * entry berstatus 'migrated_to_own_drive' — bukan sembarang resource_id —
 * supaya tidak mungkin salah trash spreadsheet yang masih aktif dipakai.
 */
function trashOldMigratedSpreadsheet(resourceMapId) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var centralId = getCentralSpreadsheet_().getId();
  var sh = _ensureResourceMapSheet_();
  var rows = sh.getDataRange().getValues();
  var idx = _getHeaderIndexMap_(sh);
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][idx.id] || '').trim();
    if (id !== String(resourceMapId || '').trim()) continue;

    var type = String(rows[i][idx.resource_type] || '').toLowerCase().trim();
    var status = String(rows[i][idx.status] || '').toLowerCase().trim();
    if (type !== 'data_spreadsheet' || status !== 'migrated_to_own_drive') {
      throw new Error('Entry ini bukan spreadsheet lama hasil migrasi — dibatalkan demi keamanan.');
    }

    var resourceId = String(rows[i][idx.resource_id] || '');
    // Jaga-jaga (defense-in-depth): TIDAK PERNAH boleh trash spreadsheet
    // CENTRAL lewat jalur ini apapun yang terjadi — lihat catatan yang
    // sama di getMigratedOldSpreadsheets().
    if (resourceId === centralId) {
      throw new Error('Spreadsheet ini adalah database CENTRAL — tidak boleh dihapus lewat panel ini.');
    }
    var emailGuru = String(rows[i][idx.email_guru] || '');
    try {
      DriveApp.getFileById(resourceId).setTrashed(true);
    } catch (e) {
      throw new Error('Gagal memindahkan ke Trash: ' + e.message);
    }

    sh.getRange(i + 1, idx.status + 1).setValue('trashed_by_sa');
    sh.getRange(i + 1, idx.updated_at + 1).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
    logAudit('TRASH_OLD_MIGRATED_SPREADSHEET', emailGuru, resourceId);
    return { success: true };
  }
  throw new Error('Entry tidak ditemukan');
}

/**
 * _MASTER_ROSTER_CENTRAL_SHEETS_
 * Sheet tanpa kolom pemilik yang jadi redundan (bisa dibangun ulang) begitu
 * SISWA central sudah kosong — lihat getCentralMasterRosterCleanupStatus.
 */
var _MASTER_ROSTER_CENTRAL_SHEETS_ = ['MasterSiswa', 'RiwayatKelas', 'MasterTahunPelajaran'];

/**
 * getCentralMasterRosterCleanupStatus()
 * SuperAdmin only, READ-ONLY — cek apakah MasterSiswa/RiwayatKelas/
 * MasterTahunPelajaran di CENTRAL sudah aman dibersihkan. Sheet-sheet ini
 * TIDAK punya kolom pemilik (tidak bisa dipisah per-guru seperti SISWA),
 * tapi begitu SISWA central benar-benar kosong (semua guru grandfathered
 * + SuperAdmin sudah connect Drive pribadi masing-masing — lihat
 * connectOwnSpreadsheet), sheet-sheet ini jadi REDUNDAN: setiap guru
 * sudah punya salinannya sendiri, dibangun ulang otomatis dari SISWA
 * miliknya sendiri lewat ensureAcademicSchema_()/migrateLegacyData_()
 * (AcademicYear.js) saat mereka connect.
 */
function getCentralMasterRosterCleanupStatus() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var centralSs = getCentralSpreadsheet_();
  var shSiswa = centralSs.getSheetByName('SISWA');
  var siswaRows = shSiswa ? Math.max(shSiswa.getLastRow() - 1, 0) : 0;

  var rosterInfo = _MASTER_ROSTER_CENTRAL_SHEETS_.map(function (name) {
    var sh = centralSs.getSheetByName(name);
    return { sheet: name, rows: sh ? Math.max(sh.getLastRow() - 1, 0) : 0 };
  });

  return {
    siswaRowsRemaining: siswaRows,
    safeToClean: siswaRows === 0,
    roster: rosterInfo
  };
}

/**
 * clearRedundantCentralMasterRoster()
 * SuperAdmin only — hapus SEMUA baris data (bukan header) di
 * MasterSiswa/RiwayatKelas/MasterTahunPelajaran CENTRAL. DIBATALKAN kalau
 * SISWA central masih ada isinya (berarti masih ada guru yang belum
 * connect Drive pribadi dan masih bergantung pada roster central ini) —
 * jangan pernah dipanggil tanpa lolos getCentralMasterRosterCleanupStatus()
 * dulu (safeToClean === true). Ini BUKAN operasi per-baris seperti
 * _migrateLegacyCentralDataForUser_ — sengaja menghapus SELURUH isi
 * (bukan disaring per-guru) karena tahap ini cuma boleh jalan kalau sudah
 * dipastikan tidak ada satu pun guru yang masih bergantung padanya.
 */
function clearRedundantCentralMasterRoster() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var status = getCentralMasterRosterCleanupStatus();
  if (!status.safeToClean) {
    throw new Error('Belum aman dibersihkan — SISWA central masih punya ' + status.siswaRowsRemaining + ' baris (masih ada guru yang belum pindah Drive pribadi).');
  }

  var centralSs = getCentralSpreadsheet_();
  var cleared = [];
  _MASTER_ROSTER_CENTRAL_SHEETS_.forEach(function (name) {
    var sh = centralSs.getSheetByName(name);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    if (lastRow > 1) {
      sh.deleteRows(2, lastRow - 1);
      cleared.push(name);
    }
  });

  logAudit('CLEAR_REDUNDANT_CENTRAL_MASTER_ROSTER', getLoginEmail(), cleared.join(', ') || '(sudah kosong)');
  return { success: true, cleared: cleared };
}
