// ============================================================
//  Kode.gs — Backend Google Apps Script
//  Tabungan Qurban 1448H · Masjid Ar-Rahman II Kelapa Gading
// ============================================================
//
//  CARA DEPLOY:
//  1. Buka https://script.google.com → buat project baru
//  2. Paste seluruh kode ini, ganti SPREADSHEET_ID di bawah
//  3. Klik Deploy → New Deployment → Web app
//     - Execute as : Me
//     - Who has access : Anyone
//  4. Copy URL deployment → paste ke variabel GAS_URL di file HTML
//
//  STRUKTUR SHEET (dibuat otomatis jika belum ada):
//  - Sheet "Groups"   : id | name | desc | maxMembers
//  - Sheet "Members"  : id | groupId | name | type | wa | active
//  - Sheet "Payments" : id | memberId | month | payDate | amount | proofUrl | note | timestamp
// ============================================================

// ⚠️ GANTI dengan ID Google Spreadsheet Anda
// Cara dapat ID: buka spreadsheet → lihat URL
// https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
const SPREADSHEET_ID = 'PASTE_SPREADSHEET_ID_DISINI';

// Nama sheet
const SHEET_GROUPS   = 'Groups';
const SHEET_MEMBERS  = 'Members';
const SHEET_PAYMENTS = 'Payments';

// Header kolom tiap sheet
const HEADERS = {
  Groups:   ['id', 'name', 'desc', 'maxMembers'],
  Members:  ['id', 'groupId', 'name', 'type', 'wa', 'active'],
  Payments: ['id', 'memberId', 'month', 'payDate', 'amount', 'proofUrl', 'note', 'timestamp'],
};

// ============================================================
// ENTRY POINT — menerima semua request dari HTML
// ============================================================
function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    let result;

    switch (action) {
      case 'getData':      result = getData();                        break;
      case 'saveGroup':    result = saveGroup(payload.data);         break;
      case 'deleteGroup':  result = deleteGroup(payload.id);         break;
      case 'saveMember':   result = saveMember(payload.data);        break;
      case 'deleteMember': result = deleteMember(payload.id);        break;
      case 'savePayment':  result = savePayment(payload.data);       break;
      case 'deletePayment':result = deletePayment(payload.id);       break;
      default:
        result = { error: 'Action tidak dikenal: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Tangani preflight OPTIONS (CORS)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Tabungan Qurban API aktif' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// INISIALISASI SPREADSHEET
// ============================================================

/**
 * Dapatkan atau buat sheet dengan nama tertentu.
 * Jika baru dibuat, tambahkan baris header.
 */
function getOrCreateSheet(sheetName) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(HEADERS[sheetName]);
    sheet.setFrozenRows(1);
    // Format header
    const headerRange = sheet.getRange(1, 1, 1, HEADERS[sheetName].length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a6b3c');
    headerRange.setFontColor('#ffffff');
  }
  return sheet;
}

/**
 * Inisialisasi semua sheet sekaligus.
 * Bisa dijalankan manual dari menu Apps Script untuk setup awal.
 */
function initSheets() {
  getOrCreateSheet(SHEET_GROUPS);
  getOrCreateSheet(SHEET_MEMBERS);
  getOrCreateSheet(SHEET_PAYMENTS);
  Logger.log('Semua sheet berhasil diinisialisasi.');
}

// ============================================================
// HELPER: baca sheet → array of objects
// ============================================================
function sheetToObjects(sheetName) {
  const sheet  = getOrCreateSheet(sheetName);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // hanya header atau kosong
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

/**
 * Cari nomor baris berdasarkan nilai kolom 'id'.
 * Mengembalikan nomor baris (1-based) atau -1 jika tidak ditemukan.
 */
function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // +1 karena 1-based
  }
  return -1;
}

/**
 * Konversi object ke array sesuai urutan header sheet.
 */
function objectToRow(sheetName, obj) {
  return HEADERS[sheetName].map(h => {
    const v = obj[h];
    if (v === undefined || v === null) return '';
    return v;
  });
}

// ============================================================
// GET DATA — ambil semua data sekaligus
// ============================================================
function getData() {
  const groups   = sheetToObjects(SHEET_GROUPS).map(normalizeGroup);
  const members  = sheetToObjects(SHEET_MEMBERS).map(normalizeMember);
  const payments = sheetToObjects(SHEET_PAYMENTS).map(normalizePayment);
  return { groups, members, payments };
}

// ============================================================
// NORMALISASI — pastikan tipe data konsisten
// ============================================================
function normalizeGroup(g) {
  return {
    id:         String(g.id),
    name:       String(g.name || ''),
    desc:       String(g.desc || ''),
    maxMembers: parseInt(g.maxMembers) || 7,
  };
}

function normalizeMember(m) {
  return {
    id:      String(m.id),
    groupId: String(m.groupId),
    name:    String(m.name || ''),
    type:    parseInt(m.type) || 0,
    wa:      String(m.wa || ''),
    active:  m.active === true || String(m.active).toLowerCase() === 'true',
  };
}

function normalizePayment(p) {
  return {
    id:        String(p.id),
    memberId:  String(p.memberId),
    month:     parseInt(p.month) || 0,
    payDate:   String(p.payDate || ''),
    amount:    parseInt(p.amount) || 0,
    proofUrl:  String(p.proofUrl || ''),
    note:      String(p.note || ''),
    timestamp: parseInt(p.timestamp) || 0,
  };
}

// ============================================================
// GROUPS CRUD
// ============================================================

/**
 * Simpan kelompok: INSERT jika baru, UPDATE jika sudah ada.
 */
function saveGroup(data) {
  if (!data || !data.id) return { error: 'Data kelompok tidak valid' };
  const sheet  = getOrCreateSheet(SHEET_GROUPS);
  const rowNum = findRowById(sheet, data.id);
  const row    = objectToRow(SHEET_GROUPS, data);

  if (rowNum === -1) {
    // INSERT baru
    sheet.appendRow(row);
  } else {
    // UPDATE baris yang sudah ada
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
  }
  return { ok: true };
}

/**
 * Hapus kelompok berdasarkan id.
 */
function deleteGroup(id) {
  if (!id) return { error: 'ID tidak valid' };
  const sheet  = getOrCreateSheet(SHEET_GROUPS);
  const rowNum = findRowById(sheet, id);
  if (rowNum === -1) return { error: 'Kelompok tidak ditemukan' };
  sheet.deleteRow(rowNum);
  return { ok: true };
}

// ============================================================
// MEMBERS CRUD
// ============================================================

/**
 * Simpan shohibul: INSERT jika baru, UPDATE jika sudah ada.
 */
function saveMember(data) {
  if (!data || !data.id) return { error: 'Data shohibul tidak valid' };
  const sheet  = getOrCreateSheet(SHEET_MEMBERS);
  const rowNum = findRowById(sheet, data.id);
  const row    = objectToRow(SHEET_MEMBERS, data);

  if (rowNum === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
  }
  return { ok: true };
}

/**
 * Hapus (nonaktifkan) shohibul berdasarkan id.
 * Menggunakan soft-delete: set active = false.
 */
function deleteMember(id) {
  if (!id) return { error: 'ID tidak valid' };
  const sheet  = getOrCreateSheet(SHEET_MEMBERS);
  const rowNum = findRowById(sheet, id);
  if (rowNum === -1) return { error: 'Shohibul tidak ditemukan' };

  // Temukan kolom 'active' (0-based index)
  const activeColIdx = HEADERS.Members.indexOf('active');
  if (activeColIdx !== -1) {
    sheet.getRange(rowNum, activeColIdx + 1).setValue(false);
  }

  // Hapus semua pembayaran milik shohibul ini
  deletePaymentsByMemberId(id);

  return { ok: true };
}

// ============================================================
// PAYMENTS CRUD
// ============================================================

/**
 * Simpan pembayaran: INSERT jika baru, UPDATE jika sudah ada.
 */
function savePayment(data) {
  if (!data || !data.id) return { error: 'Data pembayaran tidak valid' };
  const sheet  = getOrCreateSheet(SHEET_PAYMENTS);
  const rowNum = findRowById(sheet, data.id);
  const row    = objectToRow(SHEET_PAYMENTS, data);

  if (rowNum === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
  }
  return { ok: true };
}

/**
 * Hapus satu pembayaran berdasarkan id.
 */
function deletePayment(id) {
  if (!id) return { error: 'ID tidak valid' };
  const sheet  = getOrCreateSheet(SHEET_PAYMENTS);
  const rowNum = findRowById(sheet, id);
  if (rowNum === -1) return { error: 'Pembayaran tidak ditemukan' };
  sheet.deleteRow(rowNum);
  return { ok: true };
}

/**
 * Hapus semua pembayaran milik seorang shohibul.
 * Dipanggil saat deleteMember.
 * Hapus dari bawah ke atas agar indeks baris tidak bergeser.
 */
function deletePaymentsByMemberId(memberId) {
  const sheet = getOrCreateSheet(SHEET_PAYMENTS);
  const data  = sheet.getDataRange().getValues();
  const memberColIdx = HEADERS.Payments.indexOf('memberId'); // 0-based

  // Kumpulkan baris yang perlu dihapus, dari bawah ke atas
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][memberColIdx]) === String(memberId)) {
      rowsToDelete.push(i + 1); // 1-based
    }
  }
  rowsToDelete.forEach(r => sheet.deleteRow(r));
}

// ============================================================
// UTILITAS — jalankan manual untuk keperluan maintenance
// ============================================================

/**
 * Hapus semua data (groups, members, payments) — HATI-HATI!
 * Hanya header yang tersisa. Gunakan untuk reset total.
 */
function clearAllData() {
  [SHEET_GROUPS, SHEET_MEMBERS, SHEET_PAYMENTS].forEach(name => {
    const sheet = getOrCreateSheet(name);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
  });
  Logger.log('Semua data berhasil dihapus.');
}

/**
 * Tampilkan ringkasan data di Log — berguna untuk debug.
 */
function logSummary() {
  const data = getData();
  Logger.log('=== RINGKASAN DATA ===');
  Logger.log('Kelompok  : ' + data.groups.length);
  Logger.log('Shohibul  : ' + data.members.filter(m => m.active).length + ' aktif');
  Logger.log('Pembayaran: ' + data.payments.length + ' transaksi');
}
