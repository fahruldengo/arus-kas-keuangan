// ============================================================
//  EXPENSE TRACKER — Google Apps Script (API Backend)
//  Deploy sebagai Web App:
//    Execute as  : Me
//    Who can access : Anyone
// ============================================================

var SHEET_NAME = 'Transaksi';

// ── CORS HELPER ──────────────────────────────────────────────
function cors(output) {
  return output; // Apps Script otomatis menambah CORS header untuk Anyone
}

// ── ROUTER GET ───────────────────────────────────────────────
function doGet(e) {
  var action = e.parameter.action || '';
  var result;

  try {
    if (action === 'getTransactions') result = getTransactions();
    else if (action === 'getSummary')  result = getSummary();
    else if (action === 'deleteTransaction') result = deleteTransaction(e.parameter.id);
    else result = { success: false, message: 'Unknown action: ' + action };
  } catch (err) {
    result = { success: false, message: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ROUTER POST ──────────────────────────────────────────────
function doPost(e) {
  var result;
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || '';

    if (action === 'saveTransaction') result = saveTransaction(data);
    else result = { success: false, message: 'Unknown action: ' + action };
  } catch (err) {
    result = { success: false, message: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── INISIALISASI SPREADSHEET & FOLDER ────────────────────────
function getOrCreateResources() {
  var props    = PropertiesService.getScriptProperties();
  var sheetId  = props.getProperty('SHEET_ID');
  var folderId = props.getProperty('FOLDER_ID');

  if (!sheetId) {
    var ss    = SpreadsheetApp.create('Expense Tracker — Data');
    var sheet = ss.getActiveSheet();
    sheet.setName(SHEET_NAME);

    var headers = ['ID','Tanggal','Keterangan','Kategori','Jenis','Jumlah (Rp)','Foto (Link)','Waktu Input'];
    sheet.appendRow(headers);

    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#1a73e8');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    hRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    var widths = [60, 110, 260, 130, 80, 130, 300, 160];
    widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

    sheetId = ss.getId();
    props.setProperty('SHEET_ID', sheetId);
  }

  if (!folderId) {
    var folder = DriveApp.createFolder('Expense Tracker — Foto Nota');
    folderId   = folder.getId();
    props.setProperty('FOLDER_ID', folderId);
  }

  return { sheetId: sheetId, folderId: folderId };
}

// ── SIMPAN TRANSAKSI ─────────────────────────────────────────
function saveTransaction(data) {
  var res      = getOrCreateResources();
  var ss       = SpreadsheetApp.openById(res.sheetId);
  var sheet    = ss.getSheetByName(SHEET_NAME);
  var newId    = sheet.getLastRow(); // baris 1 = header
  var fotoLink = '';

  if (data.fotoBase64 && data.fotoBase64.length > 10) {
    fotoLink = uploadFoto(res.folderId, data.fotoBase64, data.fotoNama, newId);
  }

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  sheet.appendRow([
    newId,
    data.tanggal,
    data.keterangan,
    data.kategori,
    data.jenis,
    parseFloat(data.jumlah) || 0,
    fotoLink,
    now
  ]);

  var row      = sheet.getLastRow();
  var rowRange = sheet.getRange(row, 1, 1, 8);
  rowRange.setBackground(data.jenis === 'Masuk' ? '#e6f4ea' : '#fce8e6');
  sheet.getRange(row, 6).setNumberFormat('#,##0');

  return {
    success: true,
    message: 'Transaksi berhasil disimpan!',
    id: newId,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + res.sheetId
  };
}

// ── UPLOAD FOTO ───────────────────────────────────────────────
function uploadFoto(folderId, base64Data, fileName, rowId) {
  var parts    = base64Data.split(',');
  var mime     = (parts[0].match(/:(.*?);/) || [])[1] || 'image/jpeg';
  var raw      = parts.length > 1 ? parts[1] : parts[0];
  var blob     = Utilities.newBlob(Utilities.base64Decode(raw), mime, fileName || ('nota_' + rowId + '.jpg'));
  var folder   = DriveApp.getFolderById(folderId);
  var file     = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ── AMBIL TRANSAKSI ───────────────────────────────────────────
function getTransactions() {
  var res   = getOrCreateResources();
  var ss    = SpreadsheetApp.openById(res.sheetId);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + res.sheetId;
  if (data.length <= 1) return { success: true, rows: [], sheetUrl: sheetUrl };

  var headers = data[0];
  var rows = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  }).reverse();

  return { success: true, rows: rows, sheetUrl: sheetUrl };
}

// ── HAPUS TRANSAKSI ───────────────────────────────────────────
function deleteTransaction(rowId) {
  var res   = getOrCreateResources();
  var ss    = SpreadsheetApp.openById(res.sheetId);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(rowId)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Data tidak ditemukan' };
}

// ── SUMMARY ───────────────────────────────────────────────────
function getSummary() {
  var res   = getOrCreateResources();
  var ss    = SpreadsheetApp.openById(res.sheetId);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();

  var totalMasuk = 0, totalKeluar = 0;
  for (var i = 1; i < data.length; i++) {
    var j = parseFloat(data[i][5]) || 0;
    if (data[i][4] === 'Masuk')  totalMasuk  += j;
    if (data[i][4] === 'Keluar') totalKeluar += j;
  }

  return {
    success: true,
    totalMasuk:  totalMasuk,
    totalKeluar: totalKeluar,
    saldo: totalMasuk - totalKeluar,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + res.sheetId
  };
}
