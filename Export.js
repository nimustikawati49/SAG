/**
 * Export.js — Export Excel & PDF
 * Dipecah dari Code.js untuk kemudahan pemeliharaan.
 */

function dashVal_(v) {
  if (v === null || v === undefined) return '-';
  const s = String(v).trim();
  return s === '' ? '-' : v;
}

function exportJurnalExcel(){ 
  assertMinTier_('PRO');
  const auth = getAuth(); 
  const ss = getSpreadsheet_(); 
  let sh = ss.getSheetByName('EXPORT_JURNAL'); 
  if(sh) ss.deleteSheet(sh); 
  sh = ss.insertSheet('EXPORT_JURNAL'); 
  const setting = getSetting(); 
  const data = getRiwayatJurnal(); 
  sh.getRange('A1:J1').merge().setValue('Jurnal Pembelajaran').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center'); 
  sh.getRange('A2:J2').merge().setValue(setting.sekolah || '').setHorizontalAlignment('center'); 
  sh.getRange('A3:J3').merge().setValue(`Guru: ${setting.nama_guru || auth.email} | Mapel: ${setting.mata_pelajaran || '-'} | Tahun: ${setting.tahun_pelajaran || '-'} | Semester: ${setting.semester || '-'}`).setFontSize(10).setHorizontalAlignment('center'); 
  sh.getRange('A5:J5').setValues([[ 'No','Tanggal','Kelas','Pertemuan', 'Materi','Asesmen','Absensi', 'Refleksi','Foto','Edit Ke' ]]).setFontWeight('bold').setHorizontalAlignment('center'); 
  let row = 6; 
  data.forEach((d,i)=>{ 
    sh.getRange(row,1,1,10).setValues([[ 
      i + 1,
      dashVal_(d.tanggal),
      dashVal_(d.kelas),
      dashVal_(d.pertemuan),
      dashVal_(d.materi),
      dashVal_(d.asesmen),
      dashVal_(String(d.absensi || '').replace(/<br>/g, ', ')),
      dashVal_(d.refleksi),
      (d.foto && d.foto.length) ? d.foto.map((f,idx)=>`Foto ${idx+1}: ${f.full}`).join('\n') : '-',
      (d.edit_count ?? '-')
    ]]); 
    row++; 
  }); 
  const lastRow = row - 1; 
  sh.setColumnWidths(1, 1, 40); sh.setColumnWidths(2, 1, 80); sh.setColumnWidths(3, 1, 60); sh.setColumnWidths(4, 1, 80); sh.setColumnWidths(5, 1, 200); sh.setColumnWidths(6, 1, 160); sh.setColumnWidths(7, 1, 200); sh.setColumnWidths(8, 1, 200); sh.setColumnWidths(9, 1, 260); sh.setColumnWidths(10,1, 60); 
  sh.getRange(`A6:J${lastRow}`).setWrap(true).setVerticalAlignment('top'); 
  sh.getRange(`A5:J${lastRow}`).setBorder(true,true,true,true,true,true); 
  const gid = sh.getSheetId(); 
  return ss.getUrl().replace(/edit$/, `export?format=xlsx&gid=${gid}`); 
} 

function exportJurnalExcelBySemester(semester){
  assertLicenseActive();
  const auth = getAuth();
  const semFilter = String(semester || '').toLowerCase().trim();
  const ss = getSpreadsheet_();
  const jurnalSheet = ss.getSheetByName('JURNAL');
  const absSheet    = ss.getSheetByName('ABSENSI');
  const jurnalData  = jurnalSheet.getDataRange().getValues();
  const absData     = absSheet ? absSheet.getDataRange().getValues() : [];
  const setting     = getSetting();
  const filtered = jurnalData.slice(1).filter(r => {
    if(r[12] !== auth.email) return false;
    const sem = String(r[13] || '').toLowerCase().trim();
    return sem === semFilter ||
           (semFilter === 'ganjil' && (sem === 'i'  || sem === '1')) ||
           (semFilter === 'genap'  && (sem === 'ii' || sem === '2'));
  });
  if(filtered.length === 0) throw new Error('Tidak ada jurnal untuk semester ' + semester);
  const sheetName = 'EXPORT_' + semester.toUpperCase();
  let sh = ss.getSheetByName(sheetName);
  if(sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);
  sh.getRange('A1:J1').merge().setValue('Jurnal Pembelajaran')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2:J2').merge().setValue(setting.sekolah || '').setHorizontalAlignment('center');
  sh.getRange('A3:J3').merge()
    .setValue('Guru: '+(setting.nama_guru||auth.email)+
              ' | Mapel: '+(setting.mata_pelajaran||'-')+
              ' | Tahun: '+(setting.tahun_pelajaran||'-')+
              ' | Semester: '+semester)
    .setFontSize(10).setHorizontalAlignment('center');
  sh.getRange('A5:J5').setValues([[
    'No','Tanggal','Kelas','Pertemuan','Materi',
    'Asesmen','Refleksi','Foto','Absensi (Tidak Hadir)','Edit Ke'
  ]]).setFontWeight('bold').setHorizontalAlignment('center');
  let row = 6;
  filtered.forEach((r, i) => {
    const jurnalId = r[0];
    let fotoText = '-';
    try{
      if(r[14]){
        const fotos = JSON.parse(r[14]);
        if(Array.isArray(fotos) && fotos.length > 0){
          fotoText = fotos.map((f, idx) => 'Foto '+(idx+1)+': '+f.full).join('\n');
        }
      }
    }catch(e){ console.error('[JGD] parse foto export:', e.message||e); }
    const absRows = absData.filter(a => a[0] == jurnalId);
    const tidakHadir = absRows.filter(a => a[3] !== 'H').map(a => a[2]+' ('+a[3]+')');
    const absText = tidakHadir.length ? tidakHadir.join(', ') : 'NIHIL';
    sh.getRange(row, 1, 1, 10).setValues([[
      i+1,
      r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '-',
      r[2] || '-', r[4] || '-', r[5] || '-', r[7] || '-',
      r[15] || '-', fotoText, absText, r[11] || 0
    ]]);
    row++;
  });
  sh.setColumnWidths(1,1,40); sh.setColumnWidths(2,1,80);  sh.setColumnWidths(3,1,60);
  sh.setColumnWidths(4,1,80); sh.setColumnWidths(5,1,200); sh.setColumnWidths(6,1,160);
  sh.setColumnWidths(7,1,200); sh.setColumnWidths(8,1,260); sh.setColumnWidths(9,1,200);
  sh.setColumnWidths(10,1,60);
  sh.getRange('A5:J'+(row-1)).setBorder(true,true,true,true,true,true);
  sh.getRange('A6:J'+(row-1)).setWrap(true).setVerticalAlignment('top');
  const gid = sh.getSheetId();
  return ss.getUrl().replace(/edit$/, 'export?format=xlsx&gid='+gid);
}

/**
 * Export laporan absensi per siswa ke Excel
 */
function exportLaporanAbsensiSiswa(mode, kelas, period) {
  assertLicenseActive();
  const auth    = getAuth();
  const setting = getSetting();
  const data    = getLaporanAbsensiSiswa(mode, kelas, period);
  if (!data || data.length === 0) throw new Error('Tidak ada data untuk diekspor');

  const labelPeriode = mode === 'bulan'
    ? 'Bulan ' + period
    : mode === 'semester'
    ? 'Semester ' + period.charAt(0).toUpperCase() + period.slice(1)
    : 'Tahun Pelajaran ' + (setting.tahun_pelajaran || '-');

  const ss        = getSpreadsheet_();
  const sheetName = 'LAPORAN_ABS_' + kelas.replace(/\s/g,'_').toUpperCase();
  let   sh        = ss.getSheetByName(sheetName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);

  sh.getRange('A1:I1').merge().setValue('Laporan Absensi Siswa — ' + kelas)
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2:I2').merge()
    .setValue('Guru: '+(setting.nama_guru||auth.email)+'  |  '+labelPeriode+'  |  '+data[0].totalPertemuan+' pertemuan')
    .setFontSize(10).setHorizontalAlignment('center');
  sh.getRange('A4:I4').setValues([['No','NIS','Nama Siswa','JK','Hadir','Sakit','Izin','Alpha','% Hadir']])
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#6C63FF').setFontColor('#ffffff');

  const rows = data.map((s,i) =>
    [
      i+1,
      dashVal_(s.nis),
      dashVal_(s.nama),
      dashVal_(s.jk),
      dashVal_(s.hadir),
      dashVal_(s.sakit),
      dashVal_(s.izin),
      dashVal_(s.alpha),
      (s.persen === null || s.persen === undefined || s.persen === '') ? '-' : (s.persen + '%')
    ]
  );
  sh.getRange(5, 1, rows.length, 9).setValues(rows).setWrap(true).setVerticalAlignment('top');
  sh.getRange('A4:I'+(rows.length+4)).setBorder(true,true,true,true,true,true);
  sh.setColumnWidths(1,1,35); sh.setColumnWidths(2,1,90); sh.setColumnWidths(3,1,200);
  sh.setColumnWidths(4,1,40); sh.setColumnWidths(5,1,60); sh.setColumnWidths(6,1,60);
  sh.setColumnWidths(7,1,60); sh.setColumnWidths(8,1,60); sh.setColumnWidths(9,1,70);
  const gid = sh.getSheetId();
  return ss.getUrl().replace(/edit$/, 'export?format=xlsx&gid='+gid);
}

/**
 * Export jurnal per bulan tertentu (1-12) dan tahun (e.g. 2025)
 * Returns download URL (xlsx)
 */
function exportJurnalByBulan(bulan, tahun) {
  assertLicenseActive();
  const auth    = getAuth();
  bulan  = parseInt(bulan);
  tahun  = parseInt(tahun);
  if (!bulan || bulan < 1 || bulan > 12) throw new Error('Bulan tidak valid');
  if (!tahun || tahun < 2000)            throw new Error('Tahun tidak valid');

  const ss          = getSpreadsheet_();
  const jurnalSheet = ss.getSheetByName('JURNAL');
  const absSheet    = ss.getSheetByName('ABSENSI');
  const jurnalData  = jurnalSheet.getDataRange().getValues();
  const absData     = absSheet ? absSheet.getDataRange().getValues() : [];
  const setting     = getSetting();
  const tz          = Session.getScriptTimeZone();

  const filtered = jurnalData.slice(1).filter(r => {
    if (r[12] !== auth.email) return false;
    const tgl = new Date(r[1]);
    return tgl.getFullYear() === tahun && (tgl.getMonth() + 1) === bulan;
  });

  if (filtered.length === 0) throw new Error('Tidak ada jurnal untuk bulan ' + bulan + '/' + tahun);

  const namaBulan = ['Januari','Februari','Maret','April','Mei','Juni',
                     'Juli','Agustus','September','Oktober','November','Desember'][bulan - 1];
  const sheetName = 'EXP_' + namaBulan.substring(0,3).toUpperCase() + '_' + tahun;
  let sh = ss.getSheetByName(sheetName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);

  sh.getRange('A1:J1').merge().setValue('Jurnal Pembelajaran — ' + namaBulan + ' ' + tahun)
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2:J2').merge().setValue(setting.sekolah || '').setHorizontalAlignment('center');
  sh.getRange('A3:J3').merge()
    .setValue('Guru: '+(setting.nama_guru||auth.email)+
              ' | Mapel: '+(setting.mata_pelajaran||'-')+
              ' | Tahun: '+(setting.tahun_pelajaran||'-')+
              ' | Semester: '+(setting.semester||'-'))
    .setFontSize(10).setHorizontalAlignment('center');
  sh.getRange('A5:J5').setValues([[
    'No','Tanggal','Kelas','Pertemuan','Materi',
    'Asesmen','Refleksi','Foto','Absensi (Tidak Hadir)','Edit Ke'
  ]]).setFontWeight('bold').setHorizontalAlignment('center');

  let row = 6;
  filtered.forEach((r, i) => {
    const jurnalId  = r[0];
    let fotoText = '-';
    try {
      if (r[14]) {
        const fotos = JSON.parse(r[14]);
        if (Array.isArray(fotos) && fotos.length) {
          fotoText = fotos.map((f, idx) => 'Foto '+(idx+1)+': '+f.full).join('\n');
        }
      }
    } catch(e) {}
    const absRows    = absData.filter(a => a[0] == jurnalId);
    const tidakHadir = absRows.filter(a => a[3] !== 'H').map(a => a[2]+' ('+a[3]+')');
    const absText    = tidakHadir.length ? tidakHadir.join(', ') : 'NIHIL';
    sh.getRange(row, 1, 1, 10).setValues([[
      i+1,
      r[1] ? Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM-dd') : '-',
      r[2]||'-', r[4]||'-', r[5]||'-', r[7]||'-',
      r[15]||'-', fotoText, absText, r[11]||0
    ]]);
    row++;
  });
  sh.setColumnWidths(1,1,40);  sh.setColumnWidths(2,1,80);  sh.setColumnWidths(3,1,60);
  sh.setColumnWidths(4,1,80);  sh.setColumnWidths(5,1,200); sh.setColumnWidths(6,1,160);
  sh.setColumnWidths(7,1,200); sh.setColumnWidths(8,1,260); sh.setColumnWidths(9,1,200);
  sh.setColumnWidths(10,1,60);
  sh.getRange('A5:J'+(row-1)).setBorder(true,true,true,true,true,true);
  sh.getRange('A6:J'+(row-1)).setWrap(true).setVerticalAlignment('top');
  const gid = sh.getSheetId();
  return ss.getUrl().replace(/edit$/, 'export?format=xlsx&gid='+gid);
}

function exportRekapPDF(dari, sampai) {
  assertLicenseActive();
  const auth    = getAuth();
  const setting = getSetting();
  const ss      = getSpreadsheet_();

  let sh = ss.getSheetByName('_PDF_REKAP');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('_PDF_REKAP');

  const cols = 6;
  sh.getRange(1, 1, 1, cols).merge()
    .setValue('Rekap Absensi \u2013 ' + (setting.sekolah || ''))
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, cols).merge()
    .setValue('Guru: ' + (setting.nama_guru || auth.email) + '  |  Mapel: ' + (setting.mata_pelajaran || '-'))
    .setFontSize(10).setHorizontalAlignment('center');
  sh.getRange(3, 1, 1, cols).merge()
    .setValue('Periode: ' + (dari || '-') + ' s/d ' + (sampai || '-'))
    .setFontSize(10).setHorizontalAlignment('center');

  sh.getRange(5, 1, 1, cols)
    .setValues([['No', 'Kelas', 'Tanggal', 'Pertemuan', 'Asesmen', 'Keterangan']])
    .setFontWeight('bold').setBackground('#6366f1').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const rekap = getRekapAbsensi(dari, sampai);
  if (rekap.length) {
    const rows = rekap.map((r, i) => [
      i + 1,
      r.kelas      || '-',
      r.tanggal    || '-',
      r.pertemuan  || '-',
      r.asesmen    || '-',
      String(r.keterangan || '-').replace(/<br>/g, ', ')
    ]);
    sh.getRange(6, 1, rows.length, cols).setValues(rows).setWrap(true).setVerticalAlignment('top');
    sh.getRange(5, 1, rows.length + 1, cols).setBorder(true, true, true, true, true, true);
  } else {
    sh.getRange(6, 1, 1, cols).merge().setValue('Tidak ada data untuk periode ini')
      .setHorizontalAlignment('center').setFontColor('#6b7280');
  }

  sh.setColumnWidths(1, 1, 35);
  sh.setColumnWidths(2, 1, 65);
  sh.setColumnWidths(3, 1, 85);
  sh.setColumnWidths(4, 1, 130);
  sh.setColumnWidths(5, 1, 130);
  sh.setColumnWidths(6, 1, 210);

  const gid     = sh.getSheetId();
  const baseUrl = ss.getUrl().replace(/\/edit.*$/, '');
  return baseUrl + '/export?format=pdf&gid=' + gid +
    '&portrait=true&fitw=true&gridlines=false&printtitle=false' +
    '&sheetnames=false&pagenum=UNDEFINED&attachment=false';
}
