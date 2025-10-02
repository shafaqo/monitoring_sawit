// Ganti dengan ID Google Sheet Anda. ID bisa ditemukan di URL spreadsheet,
// contohnya: https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit
const SPREADSHEET_ID = "GANTI_DENGAN_ID_SPREADSHEET_ANDA"; 
const LOG_SAKLAR_SHEET_NAME = "LogSaklar";
const LOG_AKTIVITAS_SHEET_NAME = "LogAktivitas";

/**
 * Helper function to check for a sheet and its headers.
 * If the sheet or headers don't exist, they are created.
 * @param {Spreadsheet} spreadsheet The active spreadsheet object.
 * @param {string} sheetName The name of the sheet to check/create.
 * @param {string[]} headers An array of strings for the header row.
 */
function initializeSheet(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    // Set headers for the new sheet and freeze the first row
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    // If sheet exists, check if headers are missing
    const firstCell = sheet.getRange("A1").getValue();
    if (firstCell === "") {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
    }
  }
}

// Fungsi untuk menangani request POST (mengirim data ke sheet)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Otomatis buat sheet & header jika belum ada
    initializeSheet(spreadsheet, LOG_SAKLAR_SHEET_NAME, ['Timestamp', 'NamaAlat', 'Aksi', 'DurasiNyalaMenit']);
    initializeSheet(spreadsheet, LOG_AKTIVITAS_SHEET_NAME, ['Timestamp', 'Tanggal', 'Jam', 'NamaPelaksana', 'NamaKegiatan', 'Keterangan']);

    if (data.type === 'switch_log') {
      const sheet = spreadsheet.getSheetByName(LOG_SAKLAR_SHEET_NAME);
      sheet.appendRow([
        new Date(),
        data.switchName,
        data.action,
        data.durationMinutes || '' 
      ]);
    } else if (data.type === 'activity_log') {
      const sheet = spreadsheet.getSheetByName(LOG_AKTIVITAS_SHEET_NAME);
      sheet.appendRow([
        new Date(),
        data.date,
        data.time,
        data.pic,
        data.name,
        data.description
      ]);
    } else {
      throw new Error("Tipe data tidak valid.");
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Data berhasil ditambahkan.' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi untuk menangani request GET (mengambil data dari sheet)
function doGet(e) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Otomatis buat sheet & header jika belum ada
    initializeSheet(spreadsheet, LOG_SAKLAR_SHEET_NAME, ['Timestamp', 'NamaAlat', 'Aksi', 'DurasiNyalaMenit']);
    initializeSheet(spreadsheet, LOG_AKTIVITAS_SHEET_NAME, ['Timestamp', 'Tanggal', 'Jam', 'NamaPelaksana', 'NamaKegiatan', 'Keterangan']);

    // Ambil data log saklar
    const saklarSheet = spreadsheet.getSheetByName(LOG_SAKLAR_SHEET_NAME);
    const saklarData = saklarSheet.getDataRange().getValues();
    const saklarHeaders = saklarData.shift(); // Ambil header
    const saklarJson = saklarData.map(row => {
      let obj = {};
      saklarHeaders.forEach((header, i) => {
        if(header) obj[header] = row[i]
      });
      return obj;
    });

    // Ambil data log aktivitas
    const aktivitasSheet = spreadsheet.getSheetByName(LOG_AKTIVITAS_SHEET_NAME);
    const aktivitasData = aktivitasSheet.getDataRange().getValues();
    const aktivitasHeaders = aktivitasData.shift(); // Ambil header
    const aktivitasJson = aktivitasData.map(row => {
      let obj = {};
      aktivitasHeaders.forEach((header, i) => {
        if(header) obj[header] = row[i]
      });
      return obj;
    });
    
    const responseData = {
      status: 'success',
      data: {
        saklar: saklarJson,
        aktivitas: aktivitasJson
      }
    };
    
    return ContentService
      .createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

