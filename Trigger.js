/**
 * Trigger.js — Backup Otomatis & Trigger Harian
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

// NOTE: Trigger reminder renewal lisensi (setupLicenseReminderTrigger/
// removeLicenseReminderTrigger/getLicenseReminderTriggerStatus/
// runDailyLicenseReminder_) dulu ada di sini, dihapus — diverifikasi
// TIDAK ADA satu pun jalur di seluruh kode (UI maupun server) yang
// pernah menulis SCHOOL_LICENSE_EXPIRES ke ScriptProperties, jadi
// checkSchoolLicenseExpiryReminder() (License.js) tidak akan pernah
// benar-benar mengirim apa pun — lisensi sekolah di sistem ini SELALU
// default ke lifetime (lihat _readSchoolLicense_ di License.js) sejak
// model aktivasi lifetime per-akun guru diperkenalkan. Trigger ini
// tidak sedang aktif saat dihapus (dicek dulu lewat panel SuperAdmin).

function setupBackupTrigger() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyBackup_') {
      return { status: 'already_set', message: 'Trigger sudah aktif' };
    }
  }
  ScriptApp.newTrigger('runDailyBackup_')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  logAudit('SETUP_BACKUP_TRIGGER', getLoginEmail(), 'Backup harian diaktifkan (01:00)');
  return { status: true, message: 'Trigger backup harian berhasil diaktifkan (pukul 01:00)' };
}

function removeBackupTrigger() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyBackup_') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  }
  logAudit('REMOVE_BACKUP_TRIGGER', getLoginEmail(), removed + ' trigger dihapus');
  return { status: true, removed };
}

function getBackupTriggerStatus() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyBackup_') return { active: true };
  }
  return { active: false };
}

/**
 * Dipanggil otomatis oleh trigger — bukan oleh google.script.run.
 *
 * PERBAIKAN: sebelumnya cuma backup getSpreadsheet_() — di mode per_guru
 * ini resolve ke spreadsheet milik PEMILIK SCRIPT (trigger terjadwal
 * berjalan sebagai pemilik script, bukan sebagai guru manapun), jadi
 * hanya spreadsheet SuperAdmin sendiri (biasanya nyaris kosong karena
 * SuperAdmin tidak mengajar) yang ke-backup — data guru yang sebenarnya
 * tidak pernah tersentuh. Sekarang: mode central tetap backup spreadsheet
 * central seperti biasa; mode per_guru membackup spreadsheet PRIBADI
 * tiap guru aktif satu per satu.
 */
function runDailyBackup_() {
  const tz    = Session.getScriptTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm');

  if (getStorageMode_() !== 'per_guru') {
    const ss = getCentralSpreadsheet_();
    const name = 'Backup_' + stamp + '_' + ss.getName();
    let backupEmail = '';
    try { backupEmail = Session.getEffectiveUser().getEmail().toLowerCase().trim(); } catch (e) {}
    const folder = getUserResourceFolder_(backupEmail, 'backup_folder', 'BACKUP_JURNAL');
    const copy = ss.copy(name);
    DriveApp.getFileById(copy.getId()).moveTo(folder);
    logAudit('DAILY_BACKUP', 'SYSTEM', name);
    return;
  }

  const shUsersCentral = _getCentralSheetByName_('USERS');
  if (!shUsersCentral) return;
  const userData = shUsersCentral.getDataRange().getValues();
  let backedUp = 0;
  const failed = [];

  for (let ui = 1; ui < userData.length; ui++) {
    const email  = String(userData[ui][0] || '').toLowerCase().trim();
    const role   = String(userData[ui][1] || '').toLowerCase().trim();
    const status = String(userData[ui][2] || '').toLowerCase().trim();
    if (!email || role === 'superadmin' || status !== 'active') continue;

    const sid = resolveSpreadsheetIdForUser_(email);
    if (!sid) continue; // belum ter-provisioning, tidak ada apa pun untuk di-backup

    try {
      const ss     = SpreadsheetApp.openById(sid);
      const name   = 'Backup_' + stamp + '_' + ss.getName();
      const folder = getUserResourceFolder_(email, 'backup_folder', 'BACKUP_JURNAL');
      const copy   = ss.copy(name);
      DriveApp.getFileById(copy.getId()).moveTo(folder);
      backedUp++;
    } catch (e) {
      failed.push(email);
      console.error('[DAILY_BACKUP] Gagal backup ' + email + ': ' + (e.message || e));
    }
  }

  logAudit('DAILY_BACKUP', 'SYSTEM', backedUp + ' spreadsheet guru di-backup' + (failed.length ? ' | gagal: ' + failed.join(', ') : ''));
}

function getOrCreateFolder_(name) {
  let backupEmail = '';
  try { backupEmail = Session.getEffectiveUser().getEmail().toLowerCase().trim(); } catch (e) {}
  return getUserResourceFolder_(backupEmail, 'backup_folder', name);
}

/* =========================================================
   REMINDER TRIGGER — email pengingat pengisian jurnal harian
   ========================================================= */

function setupReminderTrigger() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyReminderCheck_') {
      return { status: 'already_set', message: 'Trigger reminder sudah aktif' };
    }
  }
  ScriptApp.newTrigger('runDailyReminderCheck_')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  logAudit('SETUP_REMINDER_TRIGGER', getLoginEmail(), 'Reminder harian diaktifkan (07:00)');
  return { status: true, message: 'Trigger reminder berhasil diaktifkan (pukul 07:00)' };
}

function removeReminderTrigger() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyReminderCheck_') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  }
  logAudit('REMOVE_REMINDER_TRIGGER', getLoginEmail(), removed + ' trigger dihapus');
  return { status: true, removed };
}

function getReminderTriggerStatus() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyReminderCheck_') return { active: true };
  }
  return { active: false };
}

/**
 * Dipanggil otomatis oleh trigger — cek guru yang belum isi jurnal 3 hari
 * terakhir dan kirim PENGINGAT IN-APP (bukan email — permintaan user:
 * notifikasi cukup lewat sistem, email cuma dipakai untuk hal yang
 * butuh bukti otentik seperti aktivasi akun lifetime).
 *
 * PERBAIKAN: sebelumnya cuma baca getSpreadsheet_() — di mode per_guru
 * ini resolve ke spreadsheet PEMILIK SCRIPT (trigger berjalan sebagai
 * pemilik script, bukan sebagai guru manapun), jadi JURNAL yang dicek
 * cuma milik SuperAdmin sendiri untuk SEMUA guru — hampir pasti salah
 * kirim reminder ke semua orang meski mereka sudah rajin mengisi jurnal
 * di spreadsheet mereka sendiri. Sekarang tiap guru aktif dibuka
 * spreadsheet-nya sendiri satu per satu.
 */
function runDailyReminderCheck_() {
  const shUsersCentral = _getCentralSheetByName_('USERS');
  if (!shUsersCentral) return;
  const userData = shUsersCentral.getDataRange().getValues();

  const now    = new Date();
  const cutoff = new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000); // 3 hari lalu
  const tz     = Session.getScriptTimeZone();
  const modePerGuru = getStorageMode_() === 'per_guru';

  for (let ui = 1; ui < userData.length; ui++) {
    const email  = String(userData[ui][0] || '').toLowerCase().trim();
    const role   = String(userData[ui][1] || '').toLowerCase().trim();
    const status = String(userData[ui][2] || '').toLowerCase().trim();
    if (!email || role === 'superadmin' || status !== 'active') continue;

    let ss;
    if (modePerGuru) {
      const sid = resolveSpreadsheetIdForUser_(email);
      if (!sid) continue; // belum ter-provisioning, belum ada jurnal untuk dicek
      try { ss = SpreadsheetApp.openById(sid); } catch (e) { continue; }
    } else {
      ss = getCentralSpreadsheet_();
    }

    const shJurnal = ss.getSheetByName('JURNAL');
    const jurnal = (shJurnal && shJurnal.getLastRow() > 1) ? shJurnal.getDataRange().getValues().slice(1) : [];

    const adaJurnal = jurnal.some(function(r) {
      if (String(r[12] || '').toLowerCase().trim() !== email) return false;
      const tgl = new Date(r[1]);
      return tgl >= cutoff;
    });

    if (adaJurnal) continue; // sudah isi, lewati

    try {
      const tanggal = Utilities.formatDate(now, tz, 'EEEE, d MMMM yyyy');
      _appendNotifToOpenSpreadsheet_(
        ss,
        '🔔 Pengingat: Jurnal Mengajar Belum Diisi',
        'Anda belum mengisi jurnal mengajar selama 3 hari terakhir (per ' + tanggal + '). Mohon segera dilengkapi supaya rekap dan laporan tetap akurat.',
        email
      );
      logAudit('REMINDER_SENT', 'SYSTEM', email);
    } catch(e) {
      console.error('[REMINDER] Gagal kirim notif ke ' + email + ': ' + (e.message || e));
    }
  }
}

/* =========================================================
   JADWAL REMINDER — email pengingat jadwal mengajar besok (H-1)
   ========================================================= */

function setupJadwalReminderTrigger() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runJadwalReminderCheck_') {
      return { status: 'already_set', message: 'Trigger jadwal reminder sudah aktif' };
    }
  }
  ScriptApp.newTrigger('runJadwalReminderCheck_')
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .create();
  logAudit('SETUP_JADWAL_REMINDER', getLoginEmail(), 'Reminder jadwal H-1 diaktifkan (15:00)');
  return { status: true, message: 'Trigger reminder jadwal berhasil diaktifkan (pukul 15:00)' };
}

function removeJadwalReminderTrigger() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runJadwalReminderCheck_') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  }
  logAudit('REMOVE_JADWAL_REMINDER', getLoginEmail(), removed + ' trigger dihapus');
  return { status: true, removed };
}

function getJadwalReminderStatus() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runJadwalReminderCheck_') return { active: true };
  }
  return { active: false };
}

/**
 * Dipanggil otomatis pukul 15:00 — kirim email ke setiap guru
 * tentang jadwal mengajar mereka besok hari.
 *
 * PERBAIKAN (sebelumnya kemungkinan besar tidak pernah benar-benar
 * mengirim apa pun): fungsi lama mencari sheet 'JADWAL_MENGAJAR', yang
 * tidak pernah dibuat oleh kode manapun — sheet jadwal yang sebenarnya
 * aktif dan dipakai di seluruh aplikasi bernama JADWAL_SEMESTER (lihat
 * ensureJadwalSheet_() di Jadwal.js). Ia juga membaca kolom 'jam'
 * (tidak ada di skema — yang ada jam_mulai & jam_selesai terpisah) dan
 * membandingkan nama hari dengan casing campuran ('Senin') padahal
 * kolom 'hari' selalu disimpan UPPERCASE ('SENIN', lihat baris tulis
 * jadwal di Jadwal.js). Sekaligus diperbaiki supaya benar di mode
 * per_guru: trigger terjadwal berjalan sebagai PEMILIK SCRIPT, bukan
 * sebagai guru yang bersangkutan, jadi tidak bisa pakai getSpreadsheet_()/
 * sheet() biasa (itu akan resolve ke spreadsheet pemilik script untuk
 * SEMUA guru) — di sini tiap guru aktif dibuka spreadsheet-nya sendiri
 * satu per satu lewat resolveSpreadsheetIdForUser_().
 */
function runJadwalReminderCheck_() {
  const tz      = Session.getScriptTimeZone();
  const now     = new Date();
  // Tentukan "besok" berdasarkan nama hari — kolom 'hari' di JADWAL_SEMESTER
  // selalu UPPERCASE (lihat data.hari = ....toUpperCase() di Jadwal.js).
  const tomorrowDate = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  const hariNames    = ['MINGGU','SENIN','SELASA','RABU','KAMIS','JUMAT','SABTU'];
  const tomorrowHari = hariNames[tomorrowDate.getDay()];
  const tomorrowStr  = Utilities.formatDate(tomorrowDate, tz, 'EEEE, d MMMM yyyy');

  const appName = 'Sistem Akademik Guru';
  const appUrl  = ScriptApp.getService().getUrl();
  const modePerGuru = getStorageMode_() === 'per_guru';

  const shUsersCentral = _getCentralSheetByName_('USERS');
  if (!shUsersCentral) return;
  const userData = shUsersCentral.getDataRange().getValues();

  for (let ui = 1; ui < userData.length; ui++) {
    const email  = String(userData[ui][0] || '').toLowerCase().trim();
    const role   = String(userData[ui][1] || '').toLowerCase().trim();
    const status = String(userData[ui][2] || '').toLowerCase().trim();
    if (!email || role === 'superadmin' || status !== 'active') continue;

    let ss;
    if (modePerGuru) {
      const sid = resolveSpreadsheetIdForUser_(email);
      if (!sid) continue; // belum ter-provisioning, lewati
      try { ss = SpreadsheetApp.openById(sid); } catch (e) { continue; }
    } else {
      ss = getCentralSpreadsheet_();
    }

    const shJdwl = ss.getSheetByName('JADWAL_SEMESTER') || ss.getSheetByName('jadwal_semester');
    if (!shJdwl || shJdwl.getLastRow() < 2) continue;

    // Semester/tahun ajaran aktif guru ini, dari SETTING di spreadsheet-nya sendiri.
    let semesterAktif = '';
    let tahunAktif    = '';
    const shSetting = ss.getSheetByName('SETTING');
    if (shSetting && shSetting.getLastRow() > 1) {
      const setRows   = shSetting.getDataRange().getValues();
      const setHeader = setRows[0].map(h => String(h || '').toLowerCase().trim());
      const idxEmail    = setHeader.indexOf('email');
      const idxTaAktif  = setHeader.indexOf('tahun_pelajaran_aktif');
      const idxSemAktif = setHeader.indexOf('semester_aktif');
      const idxTahunOld = setHeader.indexOf('tahun');
      const idxSemOld   = setHeader.indexOf('semester');
      for (let si = 1; si < setRows.length; si++) {
        if (idxEmail > -1 && String(setRows[si][idxEmail] || '').toLowerCase().trim() !== email) continue;
        tahunAktif = String((idxTaAktif > -1 && setRows[si][idxTaAktif]) || (idxTahunOld > -1 && setRows[si][idxTahunOld]) || '').trim();
        semesterAktif = normalizeSemesterLabel_(String((idxSemAktif > -1 && setRows[si][idxSemAktif]) || (idxSemOld > -1 && setRows[si][idxSemOld]) || ''));
        break;
      }
    }

    // Header JADWAL_SEMESTER: email, semester, hari, kelas, mapel, jam_mulai, jam_selesai, [tahun_pelajaran]
    const jdwlLastCol = shJdwl.getLastColumn();
    const jdwlHeader  = shJdwl.getRange(1, 1, 1, jdwlLastCol).getValues()[0]
      .map(h => String(h || '').toLowerCase().trim());
    const tIdx = jdwlHeader.indexOf('tahun_pelajaran');

    const rows = shJdwl.getDataRange().getValues();
    const jadwals = [];
    for (let i = 1; i < rows.length; i++) {
      const rowEmail = String(rows[i][0] || '').toLowerCase().trim();
      if (rowEmail !== email) continue;
      const hari = String(rows[i][2] || '').trim().toUpperCase();
      if (hari !== tomorrowHari) continue;
      if (semesterAktif && normalizeSemesterLabel_(String(rows[i][1] || '')) !== semesterAktif) continue;
      if (tahunAktif && tIdx > -1) {
        const tahunRow = String(rows[i][tIdx] || '').trim();
        if (tahunRow && tahunRow !== tahunAktif) continue;
      }
      jadwals.push({
        jam_mulai: formatJam(rows[i][5]),
        jam_selesai: formatJam(rows[i][6]),
        mapel: rows[i][4] || '',
        kelas: rows[i][3] || ''
      });
    }
    if (!jadwals.length) continue;

    jadwals.sort((a, b) => toMinutes_(a.jam_mulai) - toMinutes_(b.jam_mulai));

    try {
      const rows_html = jadwals.map(j =>
        `<tr>
          <td style="padding:8px 12px;border-bottom:1px solid #f3f4f6">${j.jam_mulai} - ${j.jam_selesai}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #f3f4f6">${j.mapel}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #f3f4f6">${j.kelas}</td>
        </tr>`
      ).join('');

      GmailApp.sendEmail(
        email,
        `[${appName}] Pengingat Jadwal Mengajar Besok — ${tomorrowStr}`,
        '',
        {
          htmlBody: `
<div style="font-family:sans-serif;max-width:520px;margin:auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">
  <div style="background:#6C63FF;padding:20px 24px">
    <h2 style="margin:0;color:#fff;font-size:18px">📅 Pengingat Jadwal Mengajar</h2>
  </div>
  <div style="padding:24px">
    <p>Halo, Bapak/Ibu Guru,</p>
    <p>Berikut jadwal mengajar Anda <b>besok, ${tomorrowStr}</b>:</p>
    <table style="width:100%;border-collapse:collapse;margin:12px 0;font-size:14px">
      <thead>
        <tr style="background:#f5f3ff;color:#5b21b6">
          <th style="padding:8px 12px;text-align:left">Jam</th>
          <th style="padding:8px 12px;text-align:left">Mata Pelajaran</th>
          <th style="padding:8px 12px;text-align:left">Kelas</th>
        </tr>
      </thead>
      <tbody>${rows_html}</tbody>
    </table>
    <p>Pastikan perangkat dan materi sudah disiapkan. Semangat mengajar! 💪</p>
    <a href="${appUrl}" style="display:inline-block;background:#6C63FF;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:700;margin:8px 0">
      📘 Buka Aplikasi
    </a>
    <hr style="margin:20px 0;border:none;border-top:1px solid #e5e7eb">
    <p style="font-size:12px;color:#9ca3af">Pesan ini dikirim otomatis oleh ${appName}. Jangan balas email ini.</p>
  </div>
</div>`
        }
      );
      logAudit('JADWAL_REMINDER_SENT', 'SYSTEM', email + ' | ' + jadwals.length + ' jadwal');
    } catch(e) {
      console.error('[JADWAL_REMINDER] Gagal kirim ke ' + email + ': ' + (e.message || e));
    }
  }
}
