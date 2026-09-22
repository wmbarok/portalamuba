/**
 * Portal Digital Terpadu Yayasan Al-Mubarok - Backend (Code.gs)
 * Handles doGet, database setup, and CRUD operations.
 */

function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Portal Digital Yayasan Al-Mubarok')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0');
  } catch (err) {
    return HtmlService.createHtmlOutput('<h3>Terjadi kesalahan sistem: ' + err.message + '</h3>');
  }
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan script ini terikat (bound) ke Google Spreadsheet.');
  }
  var schemas = {
    'Pustaka': ['ID', 'Judul', 'Kategori', 'Penulis', 'Deskripsi', 'FileUrl', 'CoverUrl', 'Tanggal'],
    'Acara': ['ID', 'Judul', 'Tanggal', 'Kategori', 'Deskripsi', 'ImageUrl'],
    'Galeri': ['ID', 'Judul', 'Kategori', 'ImageUrl', 'Tanggal'],
    'PPDB': ['ID', 'NamaLengkap', 'JenisKelamin', 'TempatTanggalLahir', 'NamaWali', 'NoTelepon', 'Alamat', 'JenjangPendidikan', 'TanggalDaftar']
  };
  for (var sheetName in schemas) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(schemas[sheetName]);
      sheet.getRange(1, 1, 1, schemas[sheetName].length).setFontWeight('bold');
    }
  }
  Logger.log('Setup sheets completed successfully.');
  return 'Inisialisasi Database Google Sheets Berhasil!';
}

function adminLogin(username, password) {
  if (username === 'admin' && password === 'almubarok2026') {
    return { success: true, token: 'token_' + new Date().getTime() };
  }
  throw new Error('Username atau Password admin salah!');
}

function getPortalData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      pustaka: getDataFromSheet(ss, 'Pustaka'),
      acara: getDataFromSheet(ss, 'Acara'),
      galeri: getDataFromSheet(ss, 'Galeri')
    };
  } catch (err) {
    throw new Error('Gagal mengambil data portal: ' + err.message);
  }
}

function getDataFromSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = rows[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[headers[j]] = val;
    }
    data.push(obj);
  }
  return data;
}

function addPustaka(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pustaka');
  if (!sheet) throw new Error('Sheet Pustaka tidak ditemukan.');
  if (!data.judul || !data.fileUrl) throw new Error('Judul dan URL PDF wajib diisi.');
  var id = 'PST-' + new Date().getTime();
  var tanggal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([id, data.judul, data.kategori, data.penulis, data.deskripsi, data.fileUrl, data.coverUrl || '', tanggal]);
  return { success: true, id: id };
}

function deletePustaka(id) { deleteRowById('Pustaka', id); return { success: true }; }

function addAcara(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Acara');
  if (!sheet) throw new Error('Sheet Acara tidak ditemukan.');
  if (!data.judul || !data.tanggal) throw new Error('Judul dan tanggal wajib diisi.');
  var id = 'ACR-' + new Date().getTime();
  sheet.appendRow([id, data.judul, data.tanggal, data.kategori, data.deskripsi, data.imageUrl || '']);
  return { success: true, id: id };
}

function deleteAcara(id) { deleteRowById('Acara', id); return { success: true }; }

function addGaleri(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Galeri');
  if (!sheet) throw new Error('Sheet Galeri tidak ditemukan.');
  if (!data.judul || !data.imageUrl) throw new Error('Judul dan gambar wajib diisi.');
  var id = 'GLR-' + new Date().getTime();
  var tanggal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([id, data.judul, data.kategori, data.imageUrl, tanggal]);
  return { success: true, id: id };
}

function deleteGaleri(id) { deleteRowById('Galeri', id); return { success: true }; }

function submitPPDB(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PPDB');
  if (!sheet) throw new Error('Sheet PPDB tidak ditemukan.');
  if (!data.namaLengkap || !data.noTelepon) throw new Error('Nama lengkap dan nomor telepon wajib diisi!');
  var id = 'PPDB-' + new Date().getTime();
  var tanggalDaftar = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sheet.appendRow([
    id, data.namaLengkap, data.jenisKelamin, data.tempatTanggalLahir,
    data.namaWali, data.noTelepon, data.alamat, data.jenjangPendidikan, tanggalDaftar
  ]);
  return { success: true, id: id, message: 'Pendaftaran berhasil dikirim!' };
}

function getPPDBList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return getDataFromSheet(ss, 'PPDB');
}

function deletePPDB(id) { deleteRowById('PPDB', id); return { success: true }; }

function deleteRowById(sheetName, id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet ' + sheetName + ' tidak ditemukan.');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error('Data dengan ID ' + id + ' tidak ditemukan.');
}
