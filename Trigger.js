/**
 * Trigger.js — Backup Otomatis & Trigger Harian
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

/* =========================================================
  LICENSE RENEWAL REMINDER TRIGGER
  Kirim email ke SA saat lisensi aplikasi hampir habis (H-60/30/7)
  ========================================================= */

function setupLicenseReminderTrigger() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyLicenseReminder_') {
      return { status: 'already_set', message: 'Trigger reminder lisensi sudah aktif' };
    }
  }
  ScriptApp.newTrigger('runDailyLicenseReminder_')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  logAudit('SETUP_LICENSE_REMINDER_TRIGGER', getLoginEmail(), 'Reminder lisensi aplikasi aktif (08:00)');
  return { status: true, message: 'Trigger reminder lisensi berhasil diaktifkan (08:00)' };
}

function removeLicenseReminderTrigger() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyLicenseReminder_') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  }
  logAudit('REMOVE_LICENSE_REMINDER_TRIGGER', getLoginEmail(), removed + ' trigger dihapus');
  return { status: true, removed };
}

function getLicenseReminderTriggerStatus() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyLicenseReminder_') return { active: true };
  }
  return { active: false };
}

/** Dipanggil otomatis oleh trigger — delegasi ke License.js */
function runDailyLicenseReminder_() {
  checkSchoolLicenseExpiryReminder();
}

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

/** Dipanggil otomatis oleh trigger — bukan oleh google.script.run */
function runDailyBackup_() {
  const ss     = getSpreadsheet_();
  const tz     = Session.getScriptTimeZone();
  const stamp  = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm');
  const name   = 'Backup_' + stamp + '_' + ss.getName();
  const folder = getOrCreateFolder_('BACKUP_JURNAL');
  const copy   = ss.copy(name);
  DriveApp.getFileById(copy.getId()).moveTo(folder);
  logAudit('DAILY_BACKUP', 'SYSTEM', name);
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
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
 * Dipanggil otomatis oleh trigger — cek guru yang belum isi jurnal 3 hari terakhir
 * dan kirim email pengingat.
 */
function runDailyReminderCheck_() {
  const ss       = getSpreadsheet_();
  const shUsers  = ss.getSheetByName('USERS');
  const shJurnal = ss.getSheetByName('JURNAL');
  if (!shUsers || !shJurnal) return;

  const users  = shUsers.getDataRange().getValues().slice(1);
  const jurnal = shJurnal.getDataRange().getValues().slice(1);

  const now     = new Date();
  const cutoff  = new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000); // 3 hari lalu
  const tz      = Session.getScriptTimeZone();
  const appName = 'Sistem Akademik Guru';
  const appUrl  = ScriptApp.getService().getUrl();

  users.forEach(u => {
    const email  = String(u[0] || '').toLowerCase().trim();
    const role   = String(u[1] || '').toLowerCase().trim();
    const status = String(u[2] || '').toLowerCase().trim();
    if (!email || role === 'superadmin' || status !== 'active') return;

    // Cek apakah ada jurnal dalam 3 hari terakhir
    const adaJurnal = jurnal.some(r => {
      if (String(r[12] || '').toLowerCase().trim() !== email) return false;
      const tgl = new Date(r[1]);
      return tgl >= cutoff;
    });

    if (adaJurnal) return; // Sudah isi, skip

    // Kirim email pengingat
    try {
      const tanggal = Utilities.formatDate(now, tz, 'EEEE, d MMMM yyyy');
      GmailApp.sendEmail(
        email,
        `[${appName}] Pengingat: Jurnal Mengajar Belum Diisi`,
        '',
        {
          htmlBody: `
            <div style="font-family:sans-serif;max-width:520px;margin:auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">
              <div style="background:#6C63FF;padding:20px 24px">
                <h2 style="margin:0;color:#fff;font-size:18px">🔔 Pengingat Jurnal Mengajar</h2>
              </div>
              <div style="padding:24px">
                <p>Halo, Bapak/Ibu Guru,</p>
                <p>Kami mendeteksi bahwa Anda <b>belum mengisi jurnal mengajar</b> selama 3 hari terakhir.</p>
                <p>Tanggal hari ini: <b>${tanggal}</b></p>
                <p>Mohon segera melengkapi jurnal melalui aplikasi:</p>
                <a href="${appUrl}" style="display:inline-block;background:#6C63FF;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:700;margin:8px 0">
                  📘 Buka Aplikasi
                </a>
                <hr style="margin:20px 0;border:none;border-top:1px solid #e5e7eb">
                <p style="font-size:12px;color:#9ca3af">Pesan ini dikirim otomatis oleh ${appName}. Jangan balas email ini.</p>
              </div>
            </div>
          `
        }
      );
      logAudit('REMINDER_SENT', 'SYSTEM', email);
    } catch(e) {
      console.error('[REMINDER] Gagal kirim ke ' + email + ': ' + (e.message || e));
    }
  });
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
 */
function runJadwalReminderCheck_() {
  const ss      = getSpreadsheet_();
  const shJdwl  = ss.getSheetByName('JADWAL_MENGAJAR');
  if (!shJdwl) return;

  const tz      = Session.getScriptTimeZone();
  const now     = new Date();
  // Tentukan "besok" berdasarkan nama hari
  const tomorrowDate = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  const hariNames    = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const tomorrowHari = hariNames[tomorrowDate.getDay()];
  const tomorrowStr  = Utilities.formatDate(tomorrowDate, tz, 'EEEE, d MMMM yyyy');

  const appName = 'Sistem Akademik Guru';
  const appUrl  = ScriptApp.getService().getUrl();

  const rows  = shJdwl.getDataRange().getValues();
  const header = rows[0];
  const idx = {
    hari:  header.indexOf('hari'),
    jam:   header.indexOf('jam'),
    mapel: header.indexOf('mapel'),
    kelas: header.indexOf('kelas'),
    email: header.indexOf('email'),
  };

  // Kumpulkan jadwal besok per guru
  const byEmail = {};
  rows.slice(1).forEach(r => {
    const hari  = String(r[idx.hari]  || '').trim();
    const email = String(r[idx.email] || '').toLowerCase().trim();
    if (!email || hari !== tomorrowHari) return;
    if (!byEmail[email]) byEmail[email] = [];
    byEmail[email].push({
      jam:   r[idx.jam]   || '',
      mapel: r[idx.mapel] || '',
      kelas: r[idx.kelas] || '',
    });
  });

  Object.keys(byEmail).forEach(email => {
    const jadwals = byEmail[email];
    if (!jadwals.length) return;
    try {
      const rows_html = jadwals.map(j =>
        `<tr>
          <td style="padding:8px 12px;border-bottom:1px solid #f3f4f6">${j.jam}</td>
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
  });
}
