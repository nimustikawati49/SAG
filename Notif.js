/**
 * Notif.js — Sistem Notifikasi In-App
 * Superadmin dapat broadcast notifikasi ke semua guru.
 * Sheet NOTIF: [0]=id, [1]=waktu, [2]=from, [3]=judul, [4]=pesan, [5]=target(all|email), [6]=read_by(csv email)
 */

var NOTIF_SHEET = 'NOTIF';

function _ensureNotifSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(NOTIF_SHEET);
  if (!sh) {
    sh = ss.insertSheet(NOTIF_SHEET);
    sh.appendRow(['id', 'waktu', 'from', 'judul', 'pesan', 'target', 'read_by']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * broadcastNotif(judul, pesan) — SA only
 * Mengirim notifikasi ke semua guru aktif.
 */
function broadcastNotif(judul, pesan) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  if (!judul || !judul.trim()) throw new Error('Judul tidak boleh kosong');
  if (!pesan || !pesan.trim()) throw new Error('Pesan tidak boleh kosong');

  var sh = _ensureNotifSheet_();
  var id = 'NOTIF_' + new Date().getTime();
  var tz = Session.getScriptTimeZone();
  var waktu = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");

  sh.appendRow([id, waktu, getLoginEmail(), judul.trim(), pesan.trim(), 'all', '']);
  logAudit('BROADCAST_NOTIF', getLoginEmail(), 'Judul: ' + judul.trim());
  invalidateCache_(NOTIF_SHEET);
  return { success: true, id: id };
}

/**
 * getNotifList() — Ambil semua notif yang belum dibaca oleh user saat ini
 * Returns array of {id, waktu, judul, pesan, dibaca}
 */
function getNotifList() {
  var email = getLoginEmail();
  var sh = _ensureNotifSheet_();
  var data = sh.getDataRange().getValues();

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id      = String(row[0] || '');
    if (!id) continue;

    var target  = String(row[5] || 'all').toLowerCase().trim();
    var readBy  = String(row[6] || '');

    // Target: 'all' atau email spesifik
    var forMe = (target === 'all') || (target === email);
    if (!forMe) continue;

    var readArr = readBy ? readBy.split(',') : [];
    var sudahDibaca = readArr.indexOf(email) > -1;

    result.push({
      id      : id,
      waktu   : String(row[1] || ''),
      judul   : String(row[3] || ''),
      pesan   : String(row[4] || ''),
      dibaca  : sudahDibaca
    });
  }

  // Urutkan terbaru di atas
  result.reverse();
  return result;
}

/**
 * markNotifRead(notifId) — Tandai notif sudah dibaca oleh user saat ini
 */
function markNotifRead(notifId) {
  var email = getLoginEmail();
  var sh = _ensureNotifSheet_();
  var data = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(notifId)) continue;

    var readBy  = String(data[i][6] || '');
    var readArr = readBy ? readBy.split(',') : [];
    if (readArr.indexOf(email) === -1) {
      readArr.push(email);
      sh.getRange(i + 1, 7).setValue(readArr.join(','));
    }
    invalidateCache_(NOTIF_SHEET);
    return { success: true };
  }
  return { success: false };
}

/**
 * markAllNotifRead() — Tandai semua notif sudah dibaca
 */
function markAllNotifRead() {
  var email = getLoginEmail();
  var sh = _ensureNotifSheet_();
  var data = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0] || '');
    if (!id) continue;

    var target = String(data[i][5] || 'all').toLowerCase().trim();
    var forMe  = (target === 'all') || (target === email);
    if (!forMe) continue;

    var readBy  = String(data[i][6] || '');
    var readArr = readBy ? readBy.split(',') : [];
    if (readArr.indexOf(email) === -1) {
      readArr.push(email);
      sh.getRange(i + 1, 7).setValue(readArr.join(','));
    }
  }
  invalidateCache_(NOTIF_SHEET);
  return { success: true };
}

/**
 * getUnreadNotifCount() — Jumlah notif belum dibaca untuk badge
 */
function getUnreadNotifCount() {
  var email = getLoginEmail();
  var sh = _ensureNotifSheet_();
  var data = sh.getDataRange().getValues();

  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0] || '');
    if (!id) continue;

    var target = String(data[i][5] || 'all').toLowerCase().trim();
    var forMe  = (target === 'all') || (target === email);
    if (!forMe) continue;

    var readBy  = String(data[i][6] || '');
    var readArr = readBy ? readBy.split(',') : [];
    if (readArr.indexOf(email) === -1) count++;
  }
  return count;
}

/**
 * deleteNotif(notifId) — SA only, hapus notif
 */
function deleteNotif(notifId) {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh = _ensureNotifSheet_();
  var data = sh.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(notifId)) {
      sh.deleteRow(i + 1);
      invalidateCache_(NOTIF_SHEET);
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * getAllNotifForSA() — SA only, semua notif yang pernah dikirim
 */
function getAllNotifForSA() {
  if (!isSuperAdmin()) throw new Error('AKSES_DITOLAK');
  var sh = _ensureNotifSheet_();
  var data = sh.getDataRange().getValues();

  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var readBy  = String(data[i][6] || '');
    var dibacaCount = readBy ? readBy.split(',').filter(function(e){ return e.trim(); }).length : 0;
    result.push({
      id          : String(data[i][0]),
      waktu       : String(data[i][1] || ''),
      from        : String(data[i][2] || ''),
      judul       : String(data[i][3] || ''),
      pesan       : String(data[i][4] || ''),
      target      : String(data[i][5] || 'all'),
      dibacaCount : dibacaCount
    });
  }
  result.reverse();
  return result;
}
