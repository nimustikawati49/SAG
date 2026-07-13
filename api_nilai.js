// =============================================================
// api_nilai.js  –  Modul Input & Pengolahan Nilai Siswa
// Jurnal Guru Digital v3.2  |  Modular – tidak ubah Code.js
// =============================================================
'use strict';

/* ==================== PRIVATE HELPERS ==================== */

/** Pastikan sheet NILAI_SISWA ada + header lengkap */
function sheetNilai_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('NILAI_SISWA');
  if (!sh) {
    sh = ss.insertSheet('NILAI_SISWA');
    sh.appendRow([
      'id', 'nis', 'nama_siswa', 'kelas', 'mapel',
      'uh1', 'uh2', 'uh3',
      'tgs1', 'tgs2', 'tgs3',
      'pts', 'pas', 'pat',
      'rata_uh', 'rata_tugas', 'nilai_fix', 'nilai_asli', 'nilai_katrol',
      'nilai_remedial', 'nilai_akhir',
      'tahun', 'semester', 'owner_email'
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Pastikan sheet SETTING_NILAI ada + header lengkap */
function sheetSettingNilai_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('SETTING_NILAI');
  if (!sh) {
    sh = ss.insertSheet('SETTING_NILAI');
    sh.appendRow([
      'kelas', 'mapel', 'tahun', 'semester',
      'nilai_min_target', 'nilai_max_target', 'kkm', 'owner_email'
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Cek hak akses admin/superadmin, kembalikan objek auth */
function authAdmin_() {
  const auth = getAuth();
  if (auth.role !== 'admin' && auth.role !== 'superadmin') {
    throw new Error('AKSES_DITOLAK');
  }
  return auth;
}

/**
 * Hitung nilai turunan untuk 1 baris.
 * @param {object} r  – raw row: {uh1,uh2,uh3,tgs1,tgs2,tgs3,pts,pas,pat,semester}
 * @returns {object}  – {rata_uh, rata_tugas, nilai_fix, nilai_asli}
 */
function hitungNilaiRow_(r) {
  const parse = (v) => {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    return isNaN(n) ? null : n;
  };

  const avg = (arr) => {
    const vals = arr.map(parse).filter(v => v !== null);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  };

  const rata_uh     = avg([r.uh1, r.uh2, r.uh3]);
  const rata_tugas  = avg([r.tgs1, r.tgs2, r.tgs3]);

  let nilai_fix = null;
  if (rata_uh !== null && rata_tugas !== null) {
    nilai_fix = (rata_uh + rata_tugas) / 2;
  } else if (rata_uh !== null) {
    nilai_fix = rata_uh;
  } else if (rata_tugas !== null) {
    nilai_fix = rata_tugas;
  }

  // Ganjil → PAS  |  Genap → PAT
  const isGenap   = String(r.semester || '').toLowerCase().includes('genap');
  const uas       = isGenap ? parse(r.pat) : parse(r.pas);
  const pts_num   = parse(r.pts);

  let nilai_asli = null;
  if (nilai_fix !== null && pts_num !== null && uas !== null) {
    nilai_asli = (0.4 * nilai_fix) + (0.3 * pts_num) + (0.3 * uas);
  }

  return { rata_uh, rata_tugas, nilai_fix, nilai_asli };
}

/**
 * Normalisasi nilai katrol untuk seluruh batch.
 * @param {Array}  rows       – input rows dengan field nilai_asli
 * @param {number} minTarget  – nilai katrol terendah yg diinginkan
 * @param {number} maxTarget  – nilai katrol tertinggi yg diinginkan
 * @returns {Array} rows + nilai_katrol + nilai_akhir
 */
function hitungKatrolBatch_(rows, minTarget, maxTarget) {
  const validAsli = rows
    .map(r => r.nilai_asli)
    .filter(v => v !== null && v !== undefined && !isNaN(v));

  if (!validAsli.length) return rows;

  const minAsli = Math.min(...validAsli);
  const maxAsli = Math.max(...validAsli);

  return rows.map(r => {
    const asli = r.nilai_asli;
    if (asli === null || asli === undefined || isNaN(asli)) {
      return { ...r, nilai_katrol: null, nilai_akhir: null };
    }

    let katrol;
    if (maxAsli === minAsli) {
      // Edge case: semua siswa nilai sama → semua dapat minTarget
      katrol = minTarget;
    } else {
      katrol = ((asli - minAsli) / (maxAsli - minAsli)) * (maxTarget - minTarget) + minTarget;
    }
    katrol = Math.round(katrol * 100) / 100;

    const remedial = r.nilai_remedial !== null && r.nilai_remedial !== undefined && r.nilai_remedial !== ''
      ? Number(r.nilai_remedial) : null;
    const nilai_akhir = remedial !== null
      ? Math.round(Math.max(katrol, remedial) * 100) / 100
      : katrol;

    return { ...r, nilai_katrol: katrol, nilai_akhir };
  });
}

/** Helper: round 2 desimal atau kosong */
function rnd2_(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(v);
  return isNaN(n) ? '' : Math.round(n * 100) / 100;
}

/** Helper: parse number atau empty string */
function num_(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}

/** Versi data untuk deteksi konflik edit paralel */
function getNilaiDataVersion_(kelas, mapel, tahun, semester, ownerEmail, sheetRef) {
  const sh   = sheetRef || sheetNilai_();
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return 0;

  const h   = rows[0];
  const idx = (n) => h.indexOf(n);

  let latest = 0;
  rows.slice(1).forEach(r => {
    const ownerOk = String(r[idx('owner_email')]).toLowerCase().trim() === String(ownerEmail || '').toLowerCase().trim();
    const match = ownerOk
      && String(r[idx('kelas')]).trim()    === String(kelas).trim()
      && String(r[idx('mapel')]).trim()    === String(mapel).trim()
      && String(r[idx('tahun')]).trim()    === String(tahun).trim()
      && String(r[idx('semester')]).trim() === String(semester).trim();
    if (!match) return;

    const id = String(r[idx('id')] || '');
    const m  = id.match(/^NL_(\d+)_/);
    const ts = m ? Number(m[1]) : 0;
    if (!isNaN(ts) && ts > latest) latest = ts;
  });

  return latest;
}


/* ==================== PUBLIC API ==================== */

/**
 * Daftar kelas milik user saat ini (berdasarkan email uploader di sheet SISWA).
 * Superadmin maupun Admin hanya melihat kelas yang mereka upload sendiri.
 */
function getAllKelasUntukNilai() {
  const auth = authAdmin_();
  const sh = sheet('SISWA');
  if (!sh) return [];

  const email = auth.email.toLowerCase().trim();
  const data  = sh.getDataRange().getValues().slice(1);
  const filtered = data.filter(r => String(r[5]).toLowerCase().trim() === email);
  const kelas = [...new Set(filtered.map(r => String(r[0]).trim()).filter(Boolean))];
  return kelas.sort();
}

/**
 * Data siswa pada kelas tertentu — hanya siswa yang diupload oleh user ini.
 * @param {string} kelas
 * @returns {Array} [{no_absen, nis, nama, jk}]
 */
function getDataSiswaUntukNilai(kelas) {
  const auth = authAdmin_();
  const sh = sheet('SISWA');
  if (!sh) return [];

  const email = auth.email.toLowerCase().trim();
  const data  = sh.getDataRange().getValues().slice(1);
  return data
    .filter(r =>
      String(r[0]).trim() === String(kelas).trim()
      && String(r[5]).toLowerCase().trim() === email
    )
    .map(r => ({
      no_absen: r[1],
      nis     : r[2],
      nama    : r[3],
      jk      : r[4]
    }))
    .sort((a, b) => Number(a.no_absen) - Number(b.no_absen));
}

/**
 * Ambil nilai tersimpan untuk kombinasi kelas/mapel/tahun/semester.
 * @returns {Array} array of row objects
 */
function getNilaiSiswa(kelas, mapel, tahun, semester) {
  const auth = authAdmin_();
  const sh = sheetNilai_();
  const rows = sh.getDataRange().getValues();
  const h = rows[0];
  const idx = (n) => h.indexOf(n);

  return rows.slice(1)
    .filter(r => {
      const ownerOk = String(r[idx('owner_email')]).toLowerCase().trim() === auth.email;
      return ownerOk
        && String(r[idx('kelas')]).trim()    === String(kelas).trim()
        && String(r[idx('mapel')]).trim()    === String(mapel).trim()
        && String(r[idx('tahun')]).trim()    === String(tahun).trim()
        && String(r[idx('semester')]).trim() === String(semester).trim();
    })
    .map(r => ({
      nis           : r[idx('nis')],
      nama_siswa    : r[idx('nama_siswa')],
      uh1           : r[idx('uh1')],
      uh2           : r[idx('uh2')],
      uh3           : r[idx('uh3')],
      tgs1          : r[idx('tgs1')],
      tgs2          : r[idx('tgs2')],
      tgs3          : r[idx('tgs3')],
      pts           : r[idx('pts')],
      pas           : r[idx('pas')],
      pat           : r[idx('pat')],
      rata_uh       : r[idx('rata_uh')],
      rata_tugas    : r[idx('rata_tugas')],
      nilai_fix     : r[idx('nilai_fix')],
      nilai_asli    : r[idx('nilai_asli')],
      nilai_katrol  : r[idx('nilai_katrol')],
      nilai_remedial: r[idx('nilai_remedial')],
      nilai_akhir   : r[idx('nilai_akhir')]
    }));
}

/**
 * Ambil data nilai + versi dataset untuk proteksi konflik edit.
 */
function getNilaiBundle(kelas, mapel, tahun, semester) {
  const auth = authAdmin_();
  const sh   = sheetNilai_();
  return {
    rows   : getNilaiSiswa(kelas, mapel, tahun, semester),
    version: getNilaiDataVersion_(kelas, mapel, tahun, semester, auth.email, sh)
  };
}

/**
 * Simpan seluruh nilai siswa (upsert per kelas/mapel/tahun/semester).
 * Menghapus data lama terlebih dahulu, kemudian tulis ulang.
 * @param {object} payload  { kelas, mapel, tahun, semester, setting:{...}, rows:[...] }
 */
function simpanNilaiSiswa(payload) {
  const auth  = authAdmin_();
  const sh    = sheetNilai_();
  const setting   = payload.setting   || {};
  const minTarget = Number(setting.nilai_min_target ?? 60);
  const maxTarget = Number(setting.nilai_max_target ?? 95);
  const semester  = String(payload.semester || '');

  // 0. Cek konflik edit paralel (optimistic locking)
  const currentVersion = getNilaiDataVersion_(payload.kelas, payload.mapel, payload.tahun, semester, auth.email, sh);
  const baseVersion    = Number(payload.base_version ?? 0);
  if (baseVersion !== currentVersion) {
    throw new Error('KONFLIK_DATA: Data nilai sudah diubah user lain/perangkat lain. Muat ulang data terbaru sebelum menyimpan.');
  }

  // 1. Validasi nilai (0–100)
  for (const r of payload.rows) {
    const fields = ['uh1','uh2','uh3','tgs1','tgs2','tgs3','pts','pas','pat','nilai_remedial'];
    for (const f of fields) {
      const v = r[f];
      if (v !== null && v !== undefined && v !== '') {
        const n = Number(v);
        if (isNaN(n) || n < 0 || n > 100) {
          throw new Error(`Nilai ${f} untuk siswa "${r.nama_siswa}" tidak valid (${v}). Harus 0–100.`);
        }
      }
    }
  }

  // 2. Hitung nilai turunan per baris
  let rows = payload.rows.map(r => {
    const calc = hitungNilaiRow_({ ...r, semester });
    return { ...r, ...calc };
  });

  // 3. Hitung katrol + nilai_akhir (butuh semua nilai_asli)
  rows = hitungKatrolBatch_(rows, minTarget, maxTarget);

  // 4. Baca semua data lama, pisahkan yang bukan milik filter ini
  const allRows = sh.getDataRange().getValues();
  const hdr     = allRows[0];
  const ci      = (n) => hdr.indexOf(n);

  const keepRows = allRows.slice(1).filter(r => {
    const ownerOk = String(r[ci('owner_email')]).toLowerCase().trim() === auth.email;
    const isTarget = ownerOk
      && String(r[ci('kelas')]).trim()    === String(payload.kelas).trim()
      && String(r[ci('mapel')]).trim()    === String(payload.mapel).trim()
      && String(r[ci('tahun')]).trim()    === String(payload.tahun).trim()
      && String(r[ci('semester')]).trim() === semester;
    return !isTarget;
  });

  // 5. Bangun baris baru
  const ts = Date.now();
  const newRows = rows.map((r, i) => [
    'NL_' + ts + '_' + i,
    r.nis          || '',
    r.nama_siswa   || '',
    payload.kelas,
    payload.mapel,
    num_(r.uh1),  num_(r.uh2),  num_(r.uh3),
    num_(r.tgs1), num_(r.tgs2), num_(r.tgs3),
    num_(r.pts),  num_(r.pas),  num_(r.pat),
    rnd2_(r.rata_uh),  rnd2_(r.rata_tugas),
    rnd2_(r.nilai_fix), rnd2_(r.nilai_asli),  rnd2_(r.nilai_katrol),
    num_(r.nilai_remedial),
    rnd2_(r.nilai_akhir),
    payload.tahun, semester, auth.email
  ]);

  // 6. Tulis ulang sheet (clear lama, set semua)
  const numCols = hdr.length;
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, numCols).clearContent();
  }

  const allNew = [...keepRows, ...newRows];
  if (allNew.length > 0) {
    const normalized = allNew.map(r => {
      if (r.length >= numCols) return r.slice(0, numCols);
      return [...r, ...Array(numCols - r.length).fill('')];
    });
    sh.getRange(2, 1, normalized.length, numCols).setValues(normalized);
  }

  // 7. Audit trail
  appendAuditNilai_(auth.email, 'SAVE', payload.kelas, payload.mapel, payload.tahun, semester, newRows.length);

  return { success: true, jumlah: newRows.length, version: ts };
}

/**
 * Ambil pengaturan nilai (KKM + target katrol) untuk filter tertentu.
 */
function getSettingNilai(kelas, mapel, tahun, semester) {
  const auth = authAdmin_();
  const sh   = sheetSettingNilai_();
  const rows = sh.getDataRange().getValues();
  const h    = rows[0];
  const idx  = (n) => h.indexOf(n);

  const found = rows.slice(1).find(r =>
    String(r[idx('kelas')]).trim()    === String(kelas).trim()
    && String(r[idx('mapel')]).trim()    === String(mapel).trim()
    && String(r[idx('tahun')]).trim()    === String(tahun).trim()
    && String(r[idx('semester')]).trim() === String(semester).trim()
    && String(r[idx('owner_email')]).toLowerCase().trim() === auth.email
  );

  if (!found) return { nilai_min_target: 60, nilai_max_target: 95, kkm: 70 };

  return {
    nilai_min_target: num_(found[idx('nilai_min_target')]) || 60,
    nilai_max_target: num_(found[idx('nilai_max_target')]) || 95,
    kkm             : num_(found[idx('kkm')]) || 70
  };
}

/**
 * Simpan pengaturan nilai (upsert).
 */
function simpanSettingNilai(obj) {
  const auth = authAdmin_();
  const sh   = sheetSettingNilai_();
  const rows = sh.getDataRange().getValues();
  const h    = rows[0];
  const idx  = (n) => h.indexOf(n);

  // Hapus existing (bottom-to-top)
  for (let i = rows.length - 1; i >= 1; i--) {
    const r      = rows[i];
    const ownerOk = String(r[idx('owner_email')]).toLowerCase().trim() === auth.email;
    if (ownerOk
      && String(r[idx('kelas')]).trim()    === String(obj.kelas).trim()
      && String(r[idx('mapel')]).trim()    === String(obj.mapel).trim()
      && String(r[idx('tahun')]).trim()    === String(obj.tahun).trim()
      && String(r[idx('semester')]).trim() === String(obj.semester).trim()) {
      sh.deleteRow(i + 1);
    }
  }

  sh.appendRow([
    obj.kelas, obj.mapel, obj.tahun, obj.semester,
    Number(obj.nilai_min_target ?? 60),
    Number(obj.nilai_max_target ?? 95),
    Number(obj.kkm ?? 70),
    auth.email
  ]);

  return { success: true };
}

/**
 * Parse data impor dari Excel (sudah dikonversi ke array oleh SheetJS di browser).
 * @param {Array} rows  – array of arrays (baris 0 = header)
 * @returns {object}    – { success, data: [...], errors: [...] }
 */
function importNilaiData(rows) {
  authAdmin_();
  if (!rows || rows.length < 2) return { success: false, message: 'Data kosong' };

  const rawHeader = rows[0].map(h => String(h).toLowerCase().trim().replace(/\s+/g, '_'));

  // Flexible header mapping
  const colMap = {
    nis            : ['nis', 'no_induk', 'nisn'],
    nama_siswa     : ['nama', 'nama_siswa', 'nama_lengkap'],
    kelas          : ['kelas', 'class'],
    uh1            : ['uh1', 'uh_1', 'ulangan_harian_1', 'hasil_uh1'],
    uh2            : ['uh2', 'uh_2', 'ulangan_harian_2'],
    uh3            : ['uh3', 'uh_3', 'ulangan_harian_3'],
    tgs1           : ['tgs1', 'tugas1', 'tugas_1', 'nilai_tugas_1'],
    tgs2           : ['tgs2', 'tugas2', 'tugas_2'],
    tgs3           : ['tgs3', 'tugas3', 'tugas_3'],
    pts            : ['pts', 'penilaian_tengah_semester', 'uts'],
    pas            : ['pas', 'penilaian_akhir_semester', 'uas'],
    pat            : ['pat', 'penilaian_akhir_tahun'],
    nilai_remedial : ['remedial', 'nilai_remedial', 'remidi']
  };

  const findIdx = (aliases) => {
    for (const a of aliases) {
      const i = rawHeader.indexOf(a);
      if (i >= 0) return i;
    }
    return -1;
  };

  const idxMap = {};
  Object.keys(colMap).forEach(k => { idxMap[k] = findIdx(colMap[k]); });

  const getVal = (row, key) => {
    const i = idxMap[key];
    return (i >= 0 && row[i] !== undefined && row[i] !== null) ? row[i] : '';
  };

  const data = [];
  const errors = [];

  rows.slice(1).forEach((row, ri) => {
    const nama = String(getVal(row, 'nama_siswa') || '').trim();
    if (!nama) return; // skip baris kosong

    // Validasi nilai numerik
    const numFields = ['uh1','uh2','uh3','tgs1','tgs2','tgs3','pts','pas','pat','nilai_remedial'];
    let rowError = null;
    for (const f of numFields) {
      const v = getVal(row, f);
      if (v !== '' && v !== null) {
        const n = Number(v);
        if (isNaN(n) || n < 0 || n > 100) {
          rowError = `Baris ${ri + 2}: ${f} = "${v}" tidak valid (0–100)`;
          break;
        }
      }
    }
    if (rowError) { errors.push(rowError); return; }

    data.push({
      nis           : String(getVal(row, 'nis') || '').trim(),
      nama_siswa    : nama,
      kelas         : String(getVal(row, 'kelas') || '').trim(),
      uh1           : getVal(row, 'uh1'),
      uh2           : getVal(row, 'uh2'),
      uh3           : getVal(row, 'uh3'),
      tgs1          : getVal(row, 'tgs1'),
      tgs2          : getVal(row, 'tgs2'),
      tgs3          : getVal(row, 'tgs3'),
      pts           : getVal(row, 'pts'),
      pas           : getVal(row, 'pas'),
      pat           : getVal(row, 'pat'),
      nilai_remedial: getVal(row, 'nilai_remedial')
    });
  });

  return { success: true, data, errors };
}

/**
 * Ambil data nilai + ranking untuk keperluan cetak/ekspor.
 * @returns {Array} rows diurutkan berdasarkan nilai_akhir desc, dengan field ranking
 */
function getRankingNilai(kelas, mapel, tahun, semester) {
  authAdmin_();
  const rows = getNilaiSiswa(kelas, mapel, tahun, semester);
  if (!rows.length) return [];

  // Ranking berdasarkan nilai_akhir
  const sorted = [...rows]
    .filter(r => r.nilai_akhir !== null && r.nilai_akhir !== '')
    .sort((a, b) => Number(b.nilai_akhir) - Number(a.nilai_akhir));

  let rank = 1;
  sorted.forEach((r, i) => {
    if (i > 0 && Number(r.nilai_akhir) < Number(sorted[i - 1].nilai_akhir)) rank = i + 1;
    r.ranking = rank;
  });

  // Siswa tanpa nilai_akhir → ranking terakhir
  const noVal = rows.filter(r => r.nilai_akhir === null || r.nilai_akhir === '');
  noVal.forEach(r => { r.ranking = '-'; });

  return [...sorted, ...noVal];
}

/**
 * Daftar kelas yang memiliki data nilai untuk mapel/tahun/semester tertentu.
 * Digunakan untuk multi-kelas cetak pada modal.
 */
function getKelasYangAdaNilai(mapel, tahun, semester) {
  const auth = authAdmin_();
  const sh   = sheetNilai_();
  const rows = sh.getDataRange().getValues();
  const h    = rows[0];
  const idx  = (n) => h.indexOf(n);

  const kelasList = new Set();
  rows.slice(1).forEach(r => {
    const ownerOk = String(r[idx('owner_email')]).toLowerCase().trim() === auth.email;
    if (ownerOk
      && String(r[idx('mapel')]).trim()    === String(mapel).trim()
      && String(r[idx('tahun')]).trim()    === String(tahun).trim()
      && String(r[idx('semester')]).trim() === String(semester).trim()) {
      kelasList.add(String(r[idx('kelas')]).trim());
    }
  });
  return [...kelasList].sort();
}

/**
 * Recalculate semua nilai turunan untuk baris yang dikirim (tanpa simpan).
 * Digunakan saat user ubah setting katrol lalu minta recalc.
 * @param {object} payload { rows, setting:{min,max}, semester }
 */
function recalcNilai(payload) {
  authAdmin_();
  assertMinTier_('PRO'); // Katrol nilai hanya PRO ke atas
  const minTarget = Number(payload.setting?.nilai_min_target ?? 60);
  const maxTarget = Number(payload.setting?.nilai_max_target ?? 95);
  const semester  = String(payload.semester || '');

  let rows = payload.rows.map(r => {
    const calc = hitungNilaiRow_({ ...r, semester });
    return { ...r, ...calc };
  });
  rows = hitungKatrolBatch_(rows, minTarget, maxTarget);

  return rows.map(r => ({
    nis           : r.nis,
    nama_siswa    : r.nama_siswa,
    rata_uh       : rnd2_(r.rata_uh),
    rata_tugas    : rnd2_(r.rata_tugas),
    nilai_fix     : rnd2_(r.nilai_fix),
    nilai_asli    : rnd2_(r.nilai_asli),
    nilai_katrol  : rnd2_(r.nilai_katrol),
    nilai_akhir   : rnd2_(r.nilai_akhir)
  }));
}

/* =========================================================
   AUDIT TRAIL — log setiap perubahan nilai
   ========================================================= */

function sheetAuditNilai_() {
  const ss = getSpreadsheet_();
  let sh   = ss.getSheetByName('AUDIT_NILAI');
  if (!sh) {
    sh = ss.insertSheet('AUDIT_NILAI');
    sh.appendRow(['timestamp','email','aksi','kelas','mapel','tahun','semester','jumlah_siswa']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function appendAuditNilai_(email, aksi, kelas, mapel, tahun, semester, jumlah) {
  try {
    const sh = sheetAuditNilai_();
    const tz = Session.getScriptTimeZone();
    const ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    sh.appendRow([ts, email, aksi, kelas, mapel, tahun, semester, jumlah]);
  } catch(e) {
    console.error('[AUDIT_NILAI] Gagal catat audit:', e.message || e);
  }
}

/**
 * Ambil audit trail nilai milik user sendiri (max 100 entri terbaru)
 */
function getAuditNilai() {
  const auth = authAdmin_();
  const sh   = sheetAuditNilai_();
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1)
    .filter(r => String(r[1]||'').toLowerCase().trim() === auth.email)
    .reverse()
    .slice(0, 100)
    .map(r => ({
      timestamp : r[0] || '-',
      aksi      : r[2] || '-',
      kelas     : r[3] || '-',
      mapel     : r[4] || '-',
      tahun     : r[5] || '-',
      semester  : r[6] || '-',
      jumlah    : r[7] || 0
    }));
}
