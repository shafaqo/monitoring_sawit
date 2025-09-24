// Ganti dengan ID Spreadsheet Anda
const SPREADSHEET_ID = '1tRIxUQLwhMW2D7YX9b_sKsCtbqahUbQ41xBjtvM7GtA';

// Konfigurasi untuk setiap sheet: nama dan header kolom
const SHEETS_CONFIG = {
  logTombol: {
    name: 'LogTombol',
    headers: ['Timestamp', 'ButtonID']
  },
  logKegiatan: {
    name: 'LogKegiatan',
    headers: ['Timestamp', 'Tanggal', 'Jam', 'NamaKegiatan', 'Keterangan']
  },
  catatan: {
    name: 'Catatan',
    headers: ['ID', 'Konten', 'Timestamp']
  }
};

/**
 * Fungsi helper untuk mendapatkan sheet.
 * Jika sheet tidak ada, fungsi ini akan membuatnya secara otomatis beserta header-nya.
 * @param {string} sheetName - Nama sheet yang ingin diakses (e.g., 'LogTombol').
 * @returns {Sheet} Objek sheet dari Spreadsheet.
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);

  // Jika sheet tidak ditemukan, buat baru
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Sheet "${sheetName}" telah dibuat.`);
    
    // Cari konfigurasi header yang sesuai dan tambahkan ke sheet baru
    for (const key in SHEETS_CONFIG) {
      if (SHEETS_CONFIG[key].name === sheetName) {
        sheet.appendRow(SHEETS_CONFIG[key].headers);
        Logger.log(`Header untuk "${sheetName}" telah ditambahkan.`);
        break;
      }
    }
  }
  
  return sheet;
}

// Fungsi utama yang menerima request dari web app
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.payload;
    let result;

    switch (action) {
      case 'logButtonPress':
        result = logButtonPress(payload);
        break;
      case 'addActivity':
        result = addActivity(payload);
        break;
      case 'getLogs':
        result = getLogs();
        break;
      case 'addNote':
        result = addNote(payload);
        break;
      case 'getNotes':
        result = getNotes();
        break;
      case 'updateNote':
        result = updateNote(payload);
        break;
      case 'deleteNote':
        result = deleteNote(payload);
        break;
      default:
        throw new Error('Aksi tidak valid');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- FUNGSI-FUNGSI AKSI (sudah dimodifikasi untuk otomatis membuat sheet) ---

function logButtonPress(payload) {
  const sheet = getSheet(SHEETS_CONFIG.logTombol.name);
  const timestamp = new Date().toISOString();
  sheet.appendRow([timestamp, payload.buttonId]);
  return 'Log berhasil dicatat.';
}

function addActivity(payload) {
  const sheet = getSheet(SHEETS_CONFIG.logKegiatan.name);
  const timestamp = new Date().toISOString();
  // Pastikan urutannya sesuai dengan header di SHEETS_CONFIG
  sheet.appendRow([timestamp, payload.date, payload.time, payload.name, payload.description]);
  return 'Aktivitas berhasil dicatat.';
}

function getLogs() {
  const sheet = getSheet(SHEETS_CONFIG.logTombol.name);
  // Mulai ambil data dari baris kedua untuk menghindari header
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return data;
}

// --- FUNGSI-FUNGSI CATATAN (NOTES) ---

function addNote(payload) {
  const sheet = getSheet(SHEETS_CONFIG.catatan.name);
  const id = Utilities.getUuid();
  const timestamp = new Date().toISOString();
  sheet.appendRow([id, payload.content, timestamp]);
  return 'Catatan berhasil ditambahkan.';
}

function getNotes() {
  const sheet = getSheet(SHEETS_CONFIG.catatan.name);
  // Mulai ambil data dari baris kedua untuk menghindari header
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return data;
}

function updateNote(payload) {
  const sheet = getSheet(SHEETS_CONFIG.catatan.name);
  const data = sheet.getDataRange().getValues();
  // rowIndex + 1 karena array index dimulai dari 0, sedangkan baris sheet dari 1
  const rowIndex = data.findIndex(row => row[0] === payload.id) + 1;

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 2).setValue(payload.content); // Kolom 2 (B) untuk konten
    return 'Catatan berhasil diperbarui.';
  }
  throw new Error('Catatan tidak ditemukan.');
}

function deleteNote(payload) {
  const sheet = getSheet(SHEETS_CONFIG.catatan.name);
  const data = sheet.getDataRange().getValues();
  // rowIndex + 1 karena array index dimulai dari 0, sedangkan baris sheet dari 1
  const rowIndex = data.findIndex(row => row[0] === payload.id) + 1;

  if (rowIndex > 0) {
    sheet.deleteRow(rowIndex);
    return 'Catatan berhasil dihapus.';
  }
  throw new Error('Catatan tidak ditemukan.');
}

