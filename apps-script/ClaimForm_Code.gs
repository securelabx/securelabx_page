const SPREADSHEET_ID = '1b0kkDm-AnlFYUbP_hl3i90fQT3zR09irCRWxZrEmDNA';
const CLAIM_FOLDER_ID = '1D7ITJh7eP1aiPXbBVYjeUVA2O7aN38LH';
const TIMEZONE = 'Asia/Kuala_Lumpur';
const FORM_URL = ''; // Set to your form URL e.g. 'https://yoursite.com/index.html'

// ── ROUTER ──
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = p.action || '';
  try {
    if (action === 'login') {
      const username = (p.username || '').trim().toLowerCase();
      const password = (p.password || '').trim();
      const requiredRole = (p.role || '').trim().toLowerCase();
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Password');
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const u = String(rows[i][0] || '').trim().toLowerCase();
        const pw = String(rows[i][1] || '').trim();
        if (u === username && pw === password) {
          if (u !== requiredRole) return jsonOut({ success: false, error: 'Akses tidak dibenarkan untuk peranan ini.' });
          return jsonOut({ success: true, role: u, displayName: String(rows[i][0]).trim() });
        }
      }
      return jsonOut({ success: false, error: 'Username atau password tidak sah.' });
    }
    if (action === 'getStaff') {
      const r = getStaffInfo((p.staffId || '').trim());
      return jsonOut(r
        ? { success: true, nama: r.nama, jabatan: r.jabatan, jawatan: r.jawatan, email: r.email }
        : { success: false, error: 'Staff ID tidak dijumpai.' });
    }
    if (action === 'getClaims') {
      const role = p.role || 'admin';
      return jsonOut({ success: true, claims: role === 'all' ? getAllClaims_() : getPendingClaims_(role) });
    }
    if (action === 'getClaimDetails') {
      const claim = getClaimDetails_((p.noClaim || '').trim());
      return jsonOut(claim ? { success: true, claim } : { success: false, error: 'Claim tidak dijumpai.' });
    }
    if (action === 'adminReview') {
      adminReviewClaim_((p.noClaim || '').trim(), (p.status || '').trim(), (p.catatan || '').trim());
      return jsonOut({ success: true });
    }
    if (action === 'omDecision') {
      omDecisionClaim_((p.noClaim || '').trim(), (p.keputusan || '').trim(), (p.catatan || '').trim());
      return jsonOut({ success: true });
    }
    if (action === 'generatePDF') {
      return jsonOut(generateClaimPDF_((p.noClaim || '').trim()));
    }
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
  return jsonOut({ status: 'SLX Claim API running' });
}

function doPost(e) {
  try {
    return jsonOut(submitClaim_(JSON.parse(e.postData.contents)));
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── DATE HELPER ──
function fmt_(val, format) {
  if (val === null || val === undefined || val === '') return '';
  try {
    const d = val instanceof Date ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, TIMEZONE, format || 'dd/MM/yyyy');
  } catch (e) { return String(val || ''); }
}

// ── SUBMIT CLAIM ──
function submitClaim_(data) {
  validateClaimData_(data);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName('CLAIM_MASTER');
  const details = ss.getSheetByName('CLAIM_DETAILS');
  const noClaim = generateClaimNo_();
  const timestamp = new Date();
  const folder = createClaimFolder_(noClaim, data.nama || data.staffId || 'Staff');
  let total = 0;

  data.items.forEach(function (item, i) {
    let jumlah = Number(item.jumlah || 0);
    if (item.kategori === 'Mileage') {
      const km = Number(item.km || 0);
      jumlah = km > 20 ? km : 0;
    }
    total += jumlah;
    const links = saveItemFiles_(folder, noClaim, i + 1, item.files || []);
    details.appendRow([
      timestamp, noClaim, i + 1, item.tarikh || '', item.kategori || '', item.keterangan || '',
      item.dari || '', item.ke || '', item.tujuan || '', item.km || '', item.kadarKm || '',
      item.noResit || '', item.dokumenDiperlukan || '', jumlah, links.join('\n')
    ]);
  });

  master.appendRow([
    timestamp, noClaim, data.tarikhPermohonan || '', data.staffId || '', data.nama || '',
    data.jabatan || '', data.jawatan || '', data.email || '', total,
    'Pending Admin Review', '', '', 'Pending OM', '', '', folder.getUrl(), '', data.catatan || ''
  ]);

  return { success: true, noClaim, totalClaim: total, folderUrl: folder.getUrl() };
}

// ── STAFF LOOKUP ──
function getStaffInfo(staffId) {
  if (!staffId) return null;
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('STAFF_DATABASE').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === staffId.toUpperCase()) {
      return { staffId: data[i][0], nama: data[i][1], jabatan: data[i][2], jawatan: data[i][3], email: data[i][4] };
    }
  }
  return null;
}

// ── GET CLAIMS ──
function getAllClaims_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('CLAIM_MASTER')
    .getDataRange().getValues().slice(1).map(toObj_);
}

function getPendingClaims_(role) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('CLAIM_MASTER')
    .getDataRange().getValues().slice(1).filter(function (r) {
      if (role === 'admin') return (r[9] || '') === 'Pending Admin Review';
      if (role === 'om') return (r[9] || '') === 'Pending OM Approval' && (r[12] || '') === 'Pending OM';
      return false;
    }).map(toObj_);
}

function getClaimDetails_(noClaim) {
  if (!noClaim) return null;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterRows = ss.getSheetByName('CLAIM_MASTER').getDataRange().getValues().slice(1);
  let claim = null;
  for (let i = 0; i < masterRows.length; i++) {
    if (String(masterRows[i][1]).trim() === noClaim) { claim = toObj_(masterRows[i]); break; }
  }
  if (!claim) return null;
  claim.items = ss.getSheetByName('CLAIM_DETAILS').getDataRange().getValues().slice(1)
    .filter(function (r) { return String(r[1]).trim() === noClaim; })
    .map(function (r) {
      return {
        bil: r[2] || '', tarikh: fmt_(r[3]), kategori: r[4] || '', keterangan: r[5] || '',
        dari: r[6] || '', ke: r[7] || '', tujuan: r[8] || '', km: r[9] || '',
        noResit: r[11] || '', dokumenDiperlukan: r[12] || '',
        jumlah: Number(r[13] || 0).toFixed(2), fileLinks: r[14] || ''
      };
    });
  return claim;
}

// ── ADMIN REVIEW ──
function adminReviewClaim_(noClaim, status, catatan) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('CLAIM_MASTER');
  const row = findRow_(sheet, noClaim);
  if (!row) throw new Error('No Claim tidak dijumpai.');
  const next = status === 'Dokumen Lengkap / Disyorkan' ? 'Pending OM Approval' : status;
  sheet.getRange(row, 10).setValue(next);
  sheet.getRange(row, 11).setValue(new Date());
  sheet.getRange(row, 12).setValue(catatan || '');
  if (next === 'Pending OM Approval') sheet.getRange(row, 13).setValue('Pending OM');
  sendStatusEmail_(sheet, row, noClaim, next, catatan);
}

// ── OM DECISION ──
function omDecisionClaim_(noClaim, keputusan, catatan) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('CLAIM_MASTER');
  const row = findRow_(sheet, noClaim);
  if (!row) throw new Error('No Claim tidak dijumpai.');
  sheet.getRange(row, 13).setValue(keputusan);
  sheet.getRange(row, 14).setValue(new Date());
  sheet.getRange(row, 15).setValue(catatan || '');
  sendStatusEmail_(sheet, row, noClaim, keputusan, catatan);
}

// ── GENERATE PDF ──
function generateClaimPDF_(noClaim) {
  const claim = getClaimDetails_(noClaim);
  if (!claim) throw new Error('Claim tidak dijumpai: ' + noClaim);

  const doc = DocumentApp.create('_PDF_' + noClaim);
  const body = doc.getBody();
  body.setMarginTop(50).setMarginBottom(50).setMarginLeft(50).setMarginRight(50);

  const h1 = body.appendParagraph('SECURE LABX SDN. BHD.');
  h1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  h1.editAsText().setFontSize(16).setBold(true);

  const h2 = body.appendParagraph('BORANG TUNTUTAN BAYARAN BALIK WANG');
  h2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  h2.editAsText().setFontSize(12).setBold(true);

  const h3 = body.appendParagraph('No Claim: ' + noClaim + '   |   Tarikh: ' + (claim.tarikhPermohonan || ''));
  h3.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  h3.editAsText().setFontSize(11);
  body.appendHorizontalRule();

  // Staff info
  const infoTable = body.appendTable([
    ['Staff ID', claim.staffId || ''], ['Nama', claim.nama || ''],
    ['Jabatan', claim.jabatan || ''], ['Jawatan', claim.jawatan || ''],
    ['Email', claim.email || '']
  ]);
  for (let r = 0; r < 5; r++) infoTable.getRow(r).getCell(0).editAsText().setBold(true);

  body.appendParagraph('');
  body.appendParagraph('BUTIRAN TUNTUTAN').editAsText().setBold(true).setFontSize(12);
  body.appendParagraph('');

  if (claim.items && claim.items.length) {
    const rows = [['Bil', 'Tarikh', 'Kategori', 'Keterangan', 'No Resit', 'KM', 'Jumlah (RM)']];
    claim.items.forEach(function (item) {
      rows.push([
        String(item.bil || ''), String(item.tarikh || ''), String(item.kategori || ''),
        String(item.keterangan || ''), String(item.noResit || ''),
        item.kategori === 'Mileage' ? String(item.km || '') : '-',
        'RM ' + String(item.jumlah || '0.00')
      ]);
    });
    const tbl = body.appendTable(rows);
    for (let c = 0; c < 7; c++) tbl.getRow(0).getCell(c).editAsText().setBold(true);
  }

  body.appendParagraph('');
  const tot = body.appendParagraph('JUMLAH TUNTUTAN: RM ' + claim.jumlah);
  tot.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  tot.editAsText().setBold(true).setFontSize(13);

  body.appendParagraph('');
  body.appendHorizontalRule();
  body.appendParagraph('MAKLUMAT KELULUSAN').editAsText().setBold(true);
  body.appendTable([
    ['Status Admin', claim.statusAdmin || ''], ['Catatan Admin', claim.catatanAdmin || '-'],
    ['Keputusan OM', claim.keputusanOM || ''], ['Tarikh OM', claim.tarikhOM || ''],
    ['Catatan OM', claim.catatanOM || '-']
  ]);

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  pdfBlob.setName(noClaim + '_Tuntutan.pdf');

  let targetFolder = DriveApp.getFolderById(CLAIM_FOLDER_ID);
  const subs = targetFolder.getFolders();
  while (subs.hasNext()) {
    const f = subs.next();
    if (f.getName().indexOf(noClaim) === 0) { targetFolder = f; break; }
  }

  const pdfFile = targetFolder.createFile(pdfBlob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('CLAIM_MASTER');
  const row = findRow_(sheet, noClaim);
  if (row) sheet.getRange(row, 17).setValue(pdfFile.getUrl());

  return { success: true, pdfUrl: pdfFile.getUrl() };
}

// ── SHARED HELPERS ──
function toObj_(r) {
  return {
    timestamp: fmt_(r[0], 'dd/MM/yyyy HH:mm'), noClaim: String(r[1] || ''),
    tarikhPermohonan: fmt_(r[2]), staffId: String(r[3] || ''), nama: String(r[4] || ''),
    jabatan: String(r[5] || ''), jawatan: String(r[6] || ''), email: String(r[7] || ''),
    jumlah: Number(r[8] || 0).toFixed(2), statusAdmin: String(r[9] || ''),
    tarikhAdmin: fmt_(r[10]), catatanAdmin: String(r[11] || ''),
    keputusanOM: String(r[12] || ''), tarikhOM: fmt_(r[13]), catatanOM: String(r[14] || ''),
    folderUrl: String(r[15] || ''), pdfLink: String(r[16] || ''), catatanStaff: String(r[17] || '')
  };
}

function findRow_(sheet, noClaim) {
  const vals = sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(noClaim)) return i + 2;
  }
  return null;
}

function generateClaimNo_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('CLAIM_MASTER');
  const yr = Utilities.formatDate(new Date(), TIMEZONE, 'yy');
  const mo = Utilities.formatDate(new Date(), TIMEZONE, 'MM');
  const prefix = 'CLM-' + yr + mo + '-';
  let max = 0;
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat().forEach(function (v) {
      const t = String(v || '');
      if (t.indexOf(prefix) === 0) { const n = Number(t.replace(prefix, '')); if (n > max) max = n; }
    });
  }
  return prefix + String(max + 1).padStart(4, '0');
}

function createClaimFolder_(noClaim, name) {
  return DriveApp.getFolderById(CLAIM_FOLDER_ID)
    .createFolder(noClaim + ' - ' + String(name).replace(/[\\/:*?"<>|]/g, '-'));
}

function saveItemFiles_(folder, noClaim, bil, files) {
  return files.reduce(function (links, file, idx) {
    if (!file || !file.data || !file.name) return links;
    const bytes = Utilities.base64Decode(file.data.split(',').pop());
    const blob = Utilities.newBlob(bytes, file.type || MimeType.BINARY,
      (noClaim + '_Item' + bil + '_' + (idx + 1) + '_' + file.name).replace(/[\\/:*?"<>|]/g, '-'));
    links.push(folder.createFile(blob).getUrl());
    return links;
  }, []);
}

function validateClaimData_(data) {
  if (!data) throw new Error('Tiada data claim.');
  if (!data.staffId) throw new Error('Staff ID wajib diisi.');
  if (!data.nama) throw new Error('Nama wajib diisi.');
  if (!data.items || !data.items.length) throw new Error('Sekurang-kurangnya satu item claim diperlukan.');
  data.items.forEach(function (item, i) {
    if (!item.kategori) throw new Error('Jenis claim baris ' + (i + 1) + ' wajib dipilih.');
    if (!item.tarikh) throw new Error('Tarikh baris ' + (i + 1) + ' wajib diisi.');
    if (item.kategori === 'Mileage' && (!item.dari || !item.ke || !item.km))
      throw new Error('Mileage baris ' + (i + 1) + ': Dari, Ke dan KM wajib diisi.');
  });
}

// ── EMAIL NOTIFICATION ──
function sendStatusEmail_(sheet, row, noClaim, eventType, catatan) {
  try {
    const email  = String(sheet.getRange(row, 8).getValue() || '').trim();
    const nama   = String(sheet.getRange(row, 5).getValue() || '');
    const jumlah = Number(sheet.getRange(row, 9).getValue() || 0).toFixed(2);
    if (!email) return;
    const resubmitUrl = FORM_URL ? (FORM_URL + '?resubmit=' + noClaim) : '';
    const sig = '\n\nTerima kasih.\n\n\nSecure Labx Sdn. Bhd.';
    let subject, body;
    if (eventType === 'Dokumen Tidak Lengkap / Dikembalikan') {
      subject = '[' + noClaim + '] Tuntutan Dikembalikan – Dokumen Tidak Lengkap';
      body = 'Salam ' + nama + ',\n\n' +
        'Tuntutan anda (' + noClaim + ') telah DIKEMBALIKAN oleh Admin kerana dokumen tidak lengkap.\n\n' +
        'Catatan Admin: ' + (catatan || '-') + '\n\n' +
        'Sila kemukakan semula tuntutan dengan dokumen yang lengkap.' +
        (resubmitUrl ? '\nPautan kemukakan semula: ' + resubmitUrl : '') + sig;
    } else if (eventType === 'Pending OM Approval') {
      subject = '[' + noClaim + '] Tuntutan Dalam Semakan Operations Manager';
      body = 'Salam ' + nama + ',\n\n' +
        'Tuntutan anda (' + noClaim + ') telah disemak Admin dan kini menunggu kelulusan Operations Manager.\n\n' +
        'Jumlah Tuntutan: RM ' + jumlah + '\n' +
        'Catatan Admin: ' + (catatan || '-') + sig;
    } else if (eventType === 'Lulus') {
      subject = '[' + noClaim + '] Tuntutan LULUS Diluluskan';
      body = 'Salam ' + nama + ',\n\n' +
        'Tuntutan anda (' + noClaim + ') telah LULUS diluluskan oleh Operations Manager.\n\n' +
        'Jumlah Tuntutan: RM ' + jumlah + '\n' +
        'Catatan OM: ' + (catatan || '-') + sig;
    } else if (eventType === 'Tolak') {
      subject = '[' + noClaim + '] Tuntutan DITOLAK';
      body = 'Salam ' + nama + ',\n\n' +
        'Tuntutan anda (' + noClaim + ') telah DITOLAK oleh Operations Manager.\n\n' +
        'Catatan OM: ' + (catatan || '-') +
        (resubmitUrl ? '\n\nJika ingin kemukakan semula: ' + resubmitUrl : '') + sig;
    }
    if (subject && body) MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('Email error: ' + e.message);
  }
}
