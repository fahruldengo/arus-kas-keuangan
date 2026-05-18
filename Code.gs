// ============================================================
//  EXPENSE TRACKER — Google Apps Script Backend
//  Pasang di: script.google.com → Project baru
// ============================================================

// ── KONFIGURASI ──────────────────────────────────────────────
// Kosongkan dulu. Saat pertama deploy, ID akan dibuat otomatis.
var SHEET_ID   = '';   // Isi setelah spreadsheet pertama kali dibuat
var FOLDER_ID  = '';   // Isi setelah folder Drive pertama kali dibuat
var SHEET_NAME = 'Transaksi';

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Expense Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── INISIALISASI SPREADSHEET & FOLDER ────────────────────────
function getOrCreateResources() {
  var props  = PropertiesService.getScriptProperties();
  var sheetId  = props.getProperty('SHEET_ID');
  var folderId = props.getProperty('FOLDER_ID');

  // Buat Spreadsheet jika belum ada
  if (!sheetId) {
    var ss = SpreadsheetApp.create('Expense Tracker — Data');
    var sheet = ss.getActiveSheet();
    sheet.setName(SHEET_NAME);

    // Header
    var headers = ['ID','Tanggal','Keterangan','Kategori','Jenis','Jumlah (Rp)','Foto (Link)','Waktu Input'];
    sheet.appendRow(headers);

    // Format header
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // Lebar kolom
    sheet.setColumnWidth(1, 60);
    sheet.setColumnWidth(2, 110);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 130);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 130);
    sheet.setColumnWidth(7, 300);
    sheet.setColumnWidth(8, 150);

    sheetId = ss.getId();
    props.setProperty('SHEET_ID', sheetId);
  }

  // Buat Folder Drive jika belum ada
  if (!folderId) {
    var folder = DriveApp.createFolder('Expense Tracker — Foto Nota');
    folderId = folder.getId();
    props.setProperty('FOLDER_ID', folderId);
  }

  return { sheetId: sheetId, folderId: folderId };
}

// ── SIMPAN TRANSAKSI ─────────────────────────────────────────
function saveTransaction(data) {
  try {
    var res    = getOrCreateResources();
    var ss     = SpreadsheetApp.openById(res.sheetId);
    var sheet  = ss.getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();
    var newId  = lastRow; // baris 1 = header, jadi ID = lastRow

    var fotoLink = '';

    // Upload foto jika ada
    if (data.fotoBase64 && data.fotoBase64.length > 0) {
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

    // Warnai baris berdasarkan jenis
    var row = sheet.getLastRow();
    var rowRange = sheet.getRange(row, 1, 1, 8);
    if (data.jenis === 'Masuk') {
      rowRange.setBackground('#e6f4ea');
    } else {
      rowRange.setBackground('#fce8e6');
    }

    // Format kolom jumlah
    sheet.getRange(row, 6).setNumberFormat('#,##0');

    return {
      success: true,
      message: 'Transaksi berhasil disimpan!',
      id: newId,
      sheetUrl: 'https://docs.google.com/spreadsheets/d/' + res.sheetId
    };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ── UPLOAD FOTO KE DRIVE ─────────────────────────────────────
function uploadFoto(folderId, base64Data, fileName, rowId) {
  try {
    // Pisahkan header base64 (data:image/jpeg;base64,...)
    var parts     = base64Data.split(',');
    var mimeMatch = parts[0].match(/:(.*?);/);
    var mimeType  = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    var raw       = parts.length > 1 ? parts[1] : parts[0];

    var blob   = Utilities.newBlob(Utilities.base64Decode(raw), mimeType, fileName || ('nota_' + rowId + '.jpg'));
    var folder = DriveApp.getFolderById(folderId);
    var file   = folder.createFile(blob);

    // Jadikan bisa diakses siapapun yang punya link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();
  } catch (e) {
    return 'Upload gagal: ' + e.message;
  }
}

// ── AMBIL SEMUA TRANSAKSI ────────────────────────────────────
function getTransactions() {
  try {
    var res   = getOrCreateResources();
    var ss    = SpreadsheetApp.openById(res.sheetId);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data  = sheet.getDataRange().getValues();

    if (data.length <= 1) return { success: true, rows: [], sheetUrl: 'https://docs.google.com/spreadsheets/d/' + res.sheetId };

    var headers = data[0];
    var rows = data.slice(1).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });

    return {
      success: true,
      rows: rows.reverse(), // terbaru di atas
      sheetUrl: 'https://docs.google.com/spreadsheets/d/' + res.sheetId
    };
  } catch (e) {
    return { success: false, message: e.message, rows: [] };
  }
}

// ── HAPUS TRANSAKSI ──────────────────────────────────────────
function deleteTransaction(rowId) {
  try {
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
    return { success: false, message: 'Baris tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── RINGKASAN / SUMMARY ──────────────────────────────────────
function getSummary() {
  try {
    var res   = getOrCreateResources();
    var ss    = SpreadsheetApp.openById(res.sheetId);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data  = sheet.getDataRange().getValues();

    var totalMasuk  = 0;
    var totalKeluar = 0;

    for (var i = 1; i < data.length; i++) {
      var jenis  = data[i][4];
      var jumlah = parseFloat(data[i][5]) || 0;
      if (jenis === 'Masuk')  totalMasuk  += jumlah;
      if (jenis === 'Keluar') totalKeluar += jumlah;
    }

    return {
      success: true,
      totalMasuk:  totalMasuk,
      totalKeluar: totalKeluar,
      saldo: totalMasuk - totalKeluar,
      sheetUrl: 'https://docs.google.com/spreadsheets/d/' + res.sheetId
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
