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
 */

/**
 * _isDriveSelfOwned_(email)
 * True kalau spreadsheet data guru ini SUDAH dimiliki (owner_email di
 * RESOURCE_MAP) oleh guru itu sendiri, bukan lagi hasil auto-provisioning
 * di Drive SuperAdmin. Dipakai assertLicenseActive() (Auth.js) untuk
 * mewajibkan koneksi Drive pribadi begitu akun guru diaktifkan penuh.
 */
function _isDriveSelfOwned_(email) {
  var targetEmail = String(email || '').toLowerCase().trim();
  if (!targetEmail) return false;
  var entry = _getResourceMapEntryForUser_(targetEmail, 'data_spreadsheet');
  if (!entry) return false;
  return String(entry.owner_email || '').toLowerCase().trim() === targetEmail;
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
  if (auth.role !== 'admin') return { applicable: false };

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
 * connectOwnSpreadsheet(urlOrId)
 * Guru menghubungkan spreadsheet MILIK SENDIRI (sudah dibuat + dishare
 * Editor ke SuperAdmin) sebagai database pribadinya. Data lama (kalau
 * ada di spreadsheet yang sebelumnya dipakai, mis. hasil auto-provision
 * saat trial) otomatis disalin sekali ke spreadsheet baru ini.
 */
function connectOwnSpreadsheet(urlOrId) {
  var auth = getAuth();
  if (auth.role !== 'admin') throw new Error('Fitur ini khusus akun guru.');
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

  var oldId = resolveSpreadsheetIdForUser_(email);
  var migratedSheets = [];
  if (oldId && oldId !== id) {
    var defaultNames = ['Sheet1', 'Lembar1'];
    var alreadyHasData = newSs.getSheets().some(function (sh) {
      return defaultNames.indexOf(sh.getName()) === -1;
    });
    if (!alreadyHasData) {
      var oldSs = null;
      try { oldSs = SpreadsheetApp.openById(oldId); } catch (e) {}
      if (oldSs) {
        oldSs.getSheets().forEach(function (sh) {
          var name = sh.getName();
          if (_isCentralOnlySheet_(name)) return;
          if (newSs.getSheetByName(name)) return;
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
      }
    }
  }

  var oldEntry = _getResourceMapEntryForUser_(email, 'data_spreadsheet');
  if (oldEntry && oldEntry.id) {
    try {
      _upsertResourceMapEntry_({
        id: oldEntry.id,
        deployment_id: oldEntry.deployment_id,
        email_guru: email,
        resource_type: 'data_spreadsheet',
        resource_id: oldEntry.resource_id,
        resource_name: oldEntry.resource_name,
        owner_email: oldEntry.owner_email,
        status: 'migrated_to_own_drive',
        catatan: 'Digantikan spreadsheet pribadi guru pada ' + new Date().toISOString()
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

  invalidateCache_('SETTING');
  invalidateCache_('JURNAL');
  invalidateDashboardCache_();
  if (typeof _invalidateAuthCache_ === 'function') _invalidateAuthCache_(email);
  logAudit('CONNECT_OWN_SPREADSHEET', email, id + (migratedSheets.length ? (' | migrated: ' + migratedSheets.join(',')) : ' | no migration'));

  return { success: true, spreadsheetId: id, spreadsheetName: newSs.getName(), migratedSheets: migratedSheets };
}
