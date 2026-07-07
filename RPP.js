/**
 * RPP.js — Generate RPP dari Jurnal Mengajar atau Modul Ajar
 *
 * Alur:
 *  1. Admin/SA set template Google Docs ID + folder tujuan + data kepsek
 *     via setRppTemplateConfig()
 *  2. Guru klik "Generate RPP" dari Jurnal atau Modul Ajar
 *  3. Sistem copy template → replace placeholder → simpan di folder guru
 *  4. Return link (Google Docs) atau URL export DOCX
 *
 * Placeholders dalam template:
 *  {{mapel}}            — Mata pelajaran
 *  {{kelas}}            — Kelas
 *  {{bab}}              — Nomor Bab (dari Modul Ajar)
 *  {{judul_modul}}      — Judul Modul Ajar
 *  {{materi}}           — Materi (dari Jurnal)
 *  {{tujuan_pembelajaran}} — Tujuan (dari Jurnal / Modul Ajar)
 *  {{asesmen}}          — Asesmen/Penilaian (dari Jurnal)
 *  {{refleksi}}         — Refleksi guru (dari Jurnal)
 *  {{pertemuan}}        — Pertemuan ke-
 *  {{jam_ke}}           — Jam ke-
 *  {{semester}}         — Semester
 *  {{tahun_pelajaran}}  — Tahun pelajaran
 *  {{tanggal}}          — Tanggal jurnal dibuat
 *  {{nama_guru}}        — Nama guru
 *  {{nip_guru}}         — NIP guru
 *  {{sekolah}}          — Nama sekolah
 *  {{kepala_sekolah}}   — Nama kepala sekolah
 *  {{nip_kepsek}}       — NIP kepala sekolah
 *  {{deskripsi_modul}}  — Deskripsi modul ajar
 */

// =========================================================
// CONSTANTS
// =========================================================

var RPP_PROP_TEMPLATE_ID  = 'RPP_TEMPLATE_DOC_ID';
var RPP_PROP_FOLDER_ID    = 'RPP_OUTPUT_FOLDER_ID';
var RPP_PROP_KEPSEK_NAMA  = 'RPP_KEPSEK_NAMA';
var RPP_PROP_KEPSEK_NIP   = 'RPP_KEPSEK_NIP';

// =========================================================
// KONFIGURASI
// =========================================================

/**
 * Ambil konfigurasi template RPP.
 */
function getRppTemplateConfig() {
  assertLicenseActive();
  const auth = getAuth();
  if (auth.role !== 'superadmin' && auth.role !== 'admin') throw new Error('Akses ditolak');

  const props = PropertiesService.getScriptProperties();
  const templateId = props.getProperty(RPP_PROP_TEMPLATE_ID) || '';
  const folderId   = props.getProperty(RPP_PROP_FOLDER_ID) || '';
  const kepsekNama = props.getProperty(RPP_PROP_KEPSEK_NAMA) || '';
  const kepsekNip  = props.getProperty(RPP_PROP_KEPSEK_NIP) || '';

  var templateName = '';
  var templateUrl  = '';
  if (templateId) {
    try {
      var f = DriveApp.getFileById(templateId);
      templateName = f.getName();
      templateUrl  = 'https://docs.google.com/document/d/' + templateId + '/edit';
    } catch(e) { templateName = '⚠️ File tidak ditemukan'; }
  }

  var folderName = '';
  var folderUrl  = '';
  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      folderName = folder.getName();
      folderUrl  = 'https://drive.google.com/drive/folders/' + folderId;
    } catch(e) { folderName = '⚠️ Folder tidak ditemukan'; }
  }

  return { templateId, templateName, templateUrl, folderId, folderName, folderUrl, kepsekNama, kepsekNip };
}

/**
 * Simpan konfigurasi template RPP.
 * @param {Object} data — { templateInput, folderInput, kepsekNama, kepsekNip }
 */
function setRppTemplateConfig(data) {
  assertLicenseActive();
  const auth = getAuth();
  if (auth.role !== 'superadmin' && auth.role !== 'admin') throw new Error('Akses ditolak');

  const props = PropertiesService.getScriptProperties();

  // Template Doc ID (bisa URL atau ID langsung)
  if (data.templateInput !== undefined) {
    var tRaw = String(data.templateInput || '').trim();
    if (!tRaw) {
      props.deleteProperty(RPP_PROP_TEMPLATE_ID);
    } else {
      var tId = tRaw;
      var tMatch = tRaw.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
      if (tMatch) tId = tMatch[1];
      DriveApp.getFileById(tId); // validasi akses
      props.setProperty(RPP_PROP_TEMPLATE_ID, tId);
    }
  }

  // Output Folder ID (bisa URL atau ID)
  if (data.folderInput !== undefined) {
    var fRaw = String(data.folderInput || '').trim();
    if (!fRaw) {
      props.deleteProperty(RPP_PROP_FOLDER_ID);
    } else {
      var fId = fRaw;
      var fMatch = fRaw.match(/\/folders\/([a-zA-Z0-9_-]+)/);
      if (fMatch) fId = fMatch[1];
      DriveApp.getFolderById(fId); // validasi akses
      props.setProperty(RPP_PROP_FOLDER_ID, fId);
    }
  }

  // Kepsek
  if (data.kepsekNama !== undefined) props.setProperty(RPP_PROP_KEPSEK_NAMA, String(data.kepsekNama || ''));
  if (data.kepsekNip  !== undefined) props.setProperty(RPP_PROP_KEPSEK_NIP,  String(data.kepsekNip  || ''));

  logAudit('SET_RPP_CONFIG', auth.email, 'Konfigurasi template RPP diperbarui');
  return { success: true };
}

// =========================================================
// GENERATE RPP
// =========================================================

/**
 * getModulForJurnal(jurnalId) — Ambil daftar Modul Ajar yang cocok dengan
 * kelas & email guru dari jurnal. Dipakai untuk menampilkan pilihan modul
 * sebelum generate RPP.
 */
function getModulForJurnal(jurnalId) {
  assertLicenseActive();
  const auth = getAuth();
  const jurnalData = _getJurnalRow_(jurnalId);
  if (!jurnalData) throw new Error('Jurnal tidak ditemukan');

  // Guru hanya bisa akses jurnal sendiri
  if (auth.role !== 'superadmin' && auth.role !== 'admin' && auth.role !== 'kepsek') {
    if (jurnalData.email !== auth.email) throw new Error('Akses ditolak');
  }

  const sh = ensureModulAjarSheet_();
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];

  const headers = vals[0];
  const result = [];
  for (var i = 1; i < vals.length; i++) {
    var row = {};
    headers.forEach(function(h, j){ row[h] = String(vals[i][j] === null || vals[i][j] === undefined ? '' : vals[i][j]); });
    // Cocokkan berdasarkan email + kelas
    if (String(row.email || '').toLowerCase().trim() !== jurnalData.email) continue;
    if (String(row.kelas || '').trim() !== jurnalData.kelas) continue;
    result.push({ id: row.id, mapel: row.mapel, kelas: row.kelas, bab: row.bab, judul: row.judul, deskripsi: row.deskripsi });
  }
  return result;
}

/**
 * generateRppFromJurnal(jurnalId, modulId, outputFormat)
 * outputFormat: 'gdocs' | 'docx'
 * Returns: { url, name }
 */
function generateRppFromJurnal(jurnalId, modulId, outputFormat) {
  assertLicenseActive();
  assertMinTier_('PRO');
  const auth = getAuth();

  var jurnalRow = _getJurnalRow_(jurnalId);
  if (!jurnalRow) throw new Error('Jurnal tidak ditemukan');
  if (auth.role !== 'superadmin' && auth.role !== 'admin' && auth.role !== 'kepsek') {
    if (jurnalRow.email !== auth.email) throw new Error('Akses ditolak');
  }

  var modulRow = modulId ? _getModulRow_(modulId) : null;
  var setting  = _getSettingByEmail_(jurnalRow.email);
  var props    = PropertiesService.getScriptProperties();

  var placeholders = {
    '{{mapel}}'               : modulRow ? modulRow.mapel : (setting.mata_pelajaran || '-'),
    '{{kelas}}'               : jurnalRow.kelas,
    '{{bab}}'                 : modulRow ? modulRow.bab : '-',
    '{{judul_modul}}'         : modulRow ? modulRow.judul : '-',
    '{{deskripsi_modul}}'     : modulRow ? modulRow.deskripsi : '-',
    '{{materi}}'              : jurnalRow.materi,
    '{{tujuan_pembelajaran}}' : jurnalRow.tujuan || (modulRow ? modulRow.deskripsi : '-'),
    '{{asesmen}}'             : jurnalRow.asesmen || '-',
    '{{refleksi}}'            : jurnalRow.refleksi || '-',
    '{{pertemuan}}'           : jurnalRow.pertemuan || '-',
    '{{jam_ke}}'              : jurnalRow.jam_ke || '-',
    '{{semester}}'            : jurnalRow.semester || setting.semester || '-',
    '{{tahun_pelajaran}}'     : jurnalRow.tahun_pelajaran || setting.tahun_pelajaran || '-',
    '{{tanggal}}'             : jurnalRow.tanggal,
    '{{nama_guru}}'           : setting.nama_guru || jurnalRow.email,
    '{{nip_guru}}'            : setting.nip_guru || '-',
    '{{sekolah}}'             : setting.sekolah || '-',
    '{{kepala_sekolah}}'      : props.getProperty(RPP_PROP_KEPSEK_NAMA) || '-',
    '{{nip_kepsek}}'          : props.getProperty(RPP_PROP_KEPSEK_NIP) || '-'
  };

  var fileName = 'RPP_' + placeholders['{{mapel}}'] + '_' + jurnalRow.kelas + '_' + jurnalRow.tanggal;
  var result = _fillRppTemplate_(placeholders, fileName, outputFormat);
  logAudit('GENERATE_RPP', auth.email, fileName + ' (from jurnal:' + jurnalId + ')');
  return result;
}

/**
 * generateRppFromModul(modulId, outputFormat)
 * outputFormat: 'gdocs' | 'docx'
 * Returns: { url, name }
 */
function generateRppFromModul(modulId, outputFormat) {
  assertLicenseActive();
  assertMinTier_('PRO');
  const auth = getAuth();

  var modulRow = _getModulRow_(modulId);
  if (!modulRow) throw new Error('Modul Ajar tidak ditemukan');
  if (auth.role !== 'superadmin' && auth.role !== 'admin' && auth.role !== 'kepsek') {
    if (String(modulRow.email || '').toLowerCase() !== auth.email) throw new Error('Akses ditolak');
  }

  var setting = _getSettingByEmail_(modulRow.email);
  var props   = PropertiesService.getScriptProperties();

  var placeholders = {
    '{{mapel}}'               : modulRow.mapel || '-',
    '{{kelas}}'               : modulRow.kelas || '-',
    '{{bab}}'                 : modulRow.bab || '-',
    '{{judul_modul}}'         : modulRow.judul || '-',
    '{{deskripsi_modul}}'     : modulRow.deskripsi || '-',
    '{{materi}}'              : modulRow.judul || '-',
    '{{tujuan_pembelajaran}}' : modulRow.deskripsi || '-',
    '{{asesmen}}'             : '-',
    '{{refleksi}}'            : '-',
    '{{pertemuan}}'           : '-',
    '{{jam_ke}}'              : '-',
    '{{semester}}'            : setting.semester || '-',
    '{{tahun_pelajaran}}'     : setting.tahun_pelajaran || '-',
    '{{tanggal}}'             : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy'),
    '{{nama_guru}}'           : setting.nama_guru || modulRow.email,
    '{{nip_guru}}'            : setting.nip_guru || '-',
    '{{sekolah}}'             : setting.sekolah || '-',
    '{{kepala_sekolah}}'      : props.getProperty(RPP_PROP_KEPSEK_NAMA) || '-',
    '{{nip_kepsek}}'          : props.getProperty(RPP_PROP_KEPSEK_NIP) || '-'
  };

  var fileName = 'RPP_' + modulRow.mapel + '_' + modulRow.kelas + '_' + modulRow.bab + '_' + modulRow.judul;
  var result = _fillRppTemplate_(placeholders, fileName, outputFormat);
  logAudit('GENERATE_RPP', auth.email, fileName + ' (from modul:' + modulId + ')');
  return result;
}

// =========================================================
// HELPERS INTERNAL
// =========================================================

/**
 * _fillRppTemplate_(placeholders, fileName, outputFormat)
 * Copy template → isi placeholder → simpan → return URL
 */
function _fillRppTemplate_(placeholders, fileName, outputFormat) {
  var props      = PropertiesService.getScriptProperties();
  var templateId = props.getProperty(RPP_PROP_TEMPLATE_ID) || '';
  var folderId   = props.getProperty(RPP_PROP_FOLDER_ID) || '';

  if (!templateId) throw new Error('Template RPP belum dikonfigurasi. Hubungi Admin/SuperAdmin untuk menyiapkan template Google Docs.');

  // Tentukan folder output
  var outputFolder;
  if (folderId) {
    try { outputFolder = DriveApp.getFolderById(folderId); } catch(e) { outputFolder = DriveApp.getRootFolder(); }
  } else {
    // Default: subfolder "RPP_HASIL" di My Drive
    var rootIt = DriveApp.getFoldersByName('RPP_HASIL');
    outputFolder = rootIt.hasNext() ? rootIt.next() : DriveApp.createFolder('RPP_HASIL');
  }

  // Bersihkan nama file dari karakter tidak valid
  var safeName = fileName.replace(/[\\/:*?"<>|]/g, '_').substring(0, 100);

  // Copy template
  var templateFile = DriveApp.getFileById(templateId);
  var newFile      = templateFile.makeCopy(safeName, outputFolder);
  var doc          = DocumentApp.openById(newFile.getId());
  var body         = doc.getBody();

  // Ganti semua placeholder
  Object.keys(placeholders).forEach(function(key) {
    var val = String(placeholders[key] || '-');
    body.replaceText(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), val);
  });

  doc.saveAndClose();

  if (outputFormat === 'docx') {
    // Return URL download DOCX langsung dari Google Docs
    var docxUrl = 'https://docs.google.com/document/d/' + newFile.getId() + '/export?format=docx';
    return { url: docxUrl, docUrl: 'https://docs.google.com/document/d/' + newFile.getId() + '/edit', name: safeName, format: 'docx', fileId: newFile.getId() };
  }

  // Default: return link Google Docs
  return { url: 'https://docs.google.com/document/d/' + newFile.getId() + '/edit', name: safeName, format: 'gdocs', fileId: newFile.getId() };
}

/**
 * _getJurnalRow_(jurnalId) — Ambil satu baris jurnal by ID.
 * Kolom: [0]=id, [1]=created_at, [2]=kelas, [3]=jam_ke, [4]=pertemuan,
 *        [5]=materi, [6]=tujuan, [7]=asesmen, [12]=email,
 *        [13]=semester, [15]=refleksi, [18]=tahun_pelajaran
 */
function _getJurnalRow_(jurnalId) {
  var sh = sheet('JURNAL');
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(jurnalId)) {
      var tz = Session.getScriptTimeZone();
      var tgl = data[i][1] ? Utilities.formatDate(new Date(data[i][1]), tz, 'dd MMMM yyyy') : '-';
      return {
        id             : String(data[i][0]),
        tanggal        : tgl,
        kelas          : String(data[i][2] || ''),
        jam_ke         : String(data[i][3] || ''),
        pertemuan      : String(data[i][4] || ''),
        materi         : String(data[i][5] || ''),
        tujuan         : String(data[i][6] || ''),
        asesmen        : String(data[i][7] || ''),
        email          : String(data[i][12] || '').toLowerCase().trim(),
        semester       : String(data[i][13] || ''),
        refleksi       : String(data[i][15] || ''),
        tahun_pelajaran: String(data[i][18] || '')
      };
    }
  }
  return null;
}

/**
 * _getModulRow_(modulId) — Ambil satu baris Modul Ajar by ID.
 */
function _getModulRow_(modulId) {
  var sh = ensureModulAjarSheet_();
  var vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return null;
  var headers = vals[0];
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(modulId)) {
      var row = {};
      headers.forEach(function(h, j) {
        var v = vals[i][j];
        row[h] = v instanceof Date ? v.toISOString() : String(v === null || v === undefined ? '' : v);
      });
      return row;
    }
  }
  return null;
}

/**
 * _getSettingByEmail_(email) — Ambil setting guru berdasarkan email.
 */
function _getSettingByEmail_(email) {
  var values = getSheetCached('SETTING', 60);
  if (values.length < 2) return {};
  var header = values[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var idx = {};
  header.forEach(function(h, i) { idx[h] = i; });
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idx.email] || '').toLowerCase().trim() === email.toLowerCase().trim()) {
      return {
        sekolah        : String(values[i][idx.sekolah] || ''),
        tahun_pelajaran: String(values[i][idx.tahun]   || ''),
        semester       : String(values[i][idx.semester] || ''),
        nama_guru      : String(values[i][idx.guru]     || ''),
        nip_guru       : String(values[i][idx.nip]      || ''),
        mata_pelajaran : String(values[i][idx.mapel]    || '')
      };
    }
  }
  return {};
}
