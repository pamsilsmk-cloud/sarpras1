/**
 * Backend terpusat SARPRAS SMKN 1 Pamatang Silimahuta.
 * Database: Google Spreadsheet | Foto: Google Drive.
 *
 * Jalankan fungsi setupDatabase() sekali dari editor Apps Script untuk
 * membuat struktur database. Web App juga akan menjalankannya otomatis.
 */

const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1DHUGufSc8kBcphqLxk_hpWNV6rt6nzWactTHPfu-4Os',
  PHOTO_FOLDER_ID: '1fJvANLemEapp5_DWb9ggQ0QB3kN-BP-0',
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbzlYTzLaO4fynm5U7oLRupbkwWrt11ztTBweVwRdthlR4Tip9s4TFdBbDQ5G8WpZbSc/exec',
  SHEETS: ['barang', 'users', 'peminjaman', 'auditLogs']
});

// Header tetap menjaga struktur database walaupun tabel belum memiliki data.
const DEFAULT_ADMIN = Object.freeze({
  id: 'usr-1',
  nip_nis: '197501012000031001',
  nama_lengkap: 'Frangki Munthe , ST',
  username: 'admin',
  password: 'admin123',
  role: 'Super Admin',
  jurusan: 'ALL',
  status: 'Aktif',
  no_hp: ''
});

const SCHEMA = Object.freeze({
  barang: [
    'id', 'kode_barang', 'nama_barang', 'jenis', 'jurusan', 'lokasi_id',
    'merek', 'tipe', 'nomor_seri', 'jumlah', 'satuan', 'kondisi',
    'sumber_dana', 'tahun_anggaran', 'penanggung_jawab', 'status', 'photo_url', 'photo_name',
    'sha256', 'uploaded_at'
  ],
  users: [
    'id', 'nip_nis', 'nama_lengkap', 'username', 'password', 'role',
    'jurusan', 'status', 'no_hp'
  ],
  peminjaman: [
    'id', 'kode_peminjaman', 'user_id', 'user_name', 'barang_id',
    'barang_nama', 'jumlah', 'tanggal_pinjam', 'tanggal_rencana_kembali',
    'tanggal_kembali', 'keperluan', 'foto_sebelum_url', 'foto_sebelum_hash',
    'foto_sesudah_url', 'foto_sesudah_hash', 'status', 'kondisi_sebelum',
    'kondisi_sesudah', 'approved_by'
  ],
  auditLogs: [
    'id', 'waktu', 'user_name', 'aktivitas', 'tabel_terkait',
    'ip_address', 'keterangan'
  ]
});

function doGet() {
  setupDatabase();
  return json_({
    ok: true,
    service: 'SARPRAS API',
    version: '2.0',
    database: 'centralized',
    sheets: CONFIG.SHEETS
  });
}

function doPost(e) {
  try {
    setupDatabase();
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = body.action || '';

    if (action === 'bootstrap') {
      return json_({ ok: true, data: readAll_() });
    }

    if (action === 'saveState') {
      saveState_(body.data || {});
      return json_({ ok: true, data: readAll_() });
    }

    // Seed hanya boleh mengisi tabel yang benar-benar masih kosong.
    if (action === 'seedState') {
      seedState_(body.data || {});
      return json_({ ok: true, data: readAll_() });
    }

    if (action === 'uploadPhoto') {
      return json_({ ok: true, data: uploadPhoto_(body) });
    }

    if (action === 'downloadPhoto') {
      return json_({ ok: true, data: downloadPhoto_(body) });
    }

    if (action === 'setupDatabase') {
      return json_({ ok: true, data: databaseInfo_() });
    }

    return json_({ ok: false, error: 'Action tidak dikenal.' });
  } catch (error) {
    console.error(error);
    return json_({ ok: false, error: error.message });
  }
}

/**
 * Membuat database pusat secara otomatis.
 * Fungsi ini aman dijalankan berulang kali: data lama tidak dihapus dan
 * header yang sudah ada tetap dipertahankan.
 */
function setupDatabase() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    CONFIG.SHEETS.forEach(name => {
      let sheet = spreadsheet.getSheetByName(name);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(name);
      }

      const schemaHeaders = SCHEMA[name] || [];
      const columnCount = Math.max(sheet.getLastColumn(), 1);
      const existingHeaders = sheet.getRange(1, 1, 1, columnCount)
        .getDisplayValues()[0]
        .map(header => String(header).trim())
        .filter(Boolean);

      // Schema selalu berada di depan; kolom tambahan milik pengguna tetap ada.
      const headers = [...new Set(schemaHeaders.concat(existingHeaders))];
      const currentHeaders = sheet.getRange(1, 1, 1, headers.length)
        .getDisplayValues()[0]
        .map(header => String(header).trim());

      if (currentHeaders.join('|') !== headers.join('|')) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }

      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#0f172a')
        .setFontColor('#ffffff');
      sheet.autoResizeColumns(1, headers.length);
    });

    // Login awal harus selalu tersedia agar aplikasi dapat digunakan setelah
    // database baru dibuat. Data pengguna yang sudah ada tidak disentuh.
    const usersSheet = spreadsheet.getSheetByName('users');
    if (usersSheet) {
      const userRows = readSheetRows_(usersSheet);
      const hasAdmin = userRows.some(user =>
        String(user.username || '').trim().toLowerCase() === DEFAULT_ADMIN.username
      );
      if (!hasAdmin) {
        writeSheet_(usersSheet, userRows.concat([DEFAULT_ADMIN]));
      }
    }

    SpreadsheetApp.flush();
    return databaseInfo_();
  } catch (error) {
    throw new Error(`Gagal membuat database: ${error.message}`);
  } finally {
    lock.releaseLock();
  }
}

function databaseInfo_() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return CONFIG.SHEETS.reduce((result, name) => {
    const sheet = spreadsheet.getSheetByName(name);
    result[name] = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
    return result;
  }, {});
}

function readSheetRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values.shift().map(String);
  return values
    .filter(row => row.some(value => value !== ''))
    .map(row => headers.reduce((item, header, index) => {
      if (header) item[header] = normalizeValue_(row[index]);
      return item;
    }, {}));
}

function readAll_() {
  const result = {};
  CONFIG.SHEETS.forEach(name => {
    const sheet = getSheet_(name);
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 2 || lastColumn < 1) {
      result[name] = [];
      return;
    }

    const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    const headers = values.shift().map(String);
    result[name] = values
      .filter(row => row.some(value => value !== ''))
      .map(row => headers.reduce((item, header, index) => {
        if (header) item[header] = normalizeValue_(row[index]);
        return item;
      }, {}));
  });
  return result;
}

function saveState_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Object.keys(data).forEach(name => {
      if (CONFIG.SHEETS.includes(name) && Array.isArray(data[name])) {
        writeSheet_(getSheet_(name), data[name]);
      }
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function seedState_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Object.keys(data).forEach(name => {
      if (!CONFIG.SHEETS.includes(name) || !Array.isArray(data[name])) return;
      const sheet = getSheet_(name);
      if (sheet.getLastRow() <= 1 && data[name].length > 0) {
        writeSheet_(sheet, data[name]);
      }
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function writeSheet_(sheet, rows) {
  const schemaHeaders = SCHEMA[sheet.getName()] || [];
  const rowHeaders = rows.flatMap(row => Object.keys(row || {}));
  const headers = [...new Set(schemaHeaders.concat(rowHeaders))];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (!rows.length) return;

  const values = rows.map(row => headers.map(header => {
    const value = row[header];
    return value !== null && typeof value === 'object'
      ? JSON.stringify(value)
      : value ?? '';
  }));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0f172a')
    .setFontColor('#ffffff');
}

function normalizeValue_(value) {
  return value instanceof Date
    ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    : value;
}

function uploadPhoto_(body) {
  if (!body.base64 || !body.fileName) throw new Error('Foto dan nama file wajib diisi.');
  const folder = DriveApp.getFolderById(CONFIG.PHOTO_FOLDER_ID);
  const match = String(body.base64).match(/^data:(.*?);base64,(.*)$/);
  if (!match) throw new Error('Format foto tidak valid.');

  const requestedName = String(body.fileName).trim().replace(/[^a-zA-Z0-9._-]/g, '_');
  const safeName = requestedName.toLowerCase().endsWith('.jpg')
    ? requestedName
    : `${requestedName}.jpg`;
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], safeName);
  const file = folder.createFile(blob);
  file.setDescription(`SARPRAS | Kode: ${body.kodeBarang || safeName} | User: ${body.userName || 'unknown'}`);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    console.warn('Sharing foto tidak dapat diubah:', error.message);
  }
  return {
    id: file.getId(),
    name: file.getName(),
    url: `https://drive.google.com/uc?export=view&id=${file.getId()}`
  };
}

function downloadPhoto_(body) {
  const photoUrl = String(body.photoUrl || '').trim();
  if (!photoUrl) throw new Error('URL foto wajib diisi.');

  const fileIdMatch = photoUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/) ||
    photoUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!fileIdMatch) throw new Error('ID file Drive tidak ditemukan pada URL foto.');

  const file = DriveApp.getFileById(fileIdMatch[1]);
  const blob = file.getBlob();
  const bytes = blob.getBytes();

  return {
    fileName: file.getName(),
    mimeType: blob.getContentType(),
    // Encode byte array langsung agar file JPEG/PNG tidak rusak.
    base64: `data:${blob.getContentType()};base64,${Utilities.base64Encode(bytes)}`
  };
}

function getSheet_(name) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    setupDatabase();
    sheet = spreadsheet.getSheetByName(name);
  }
  if (!sheet) throw new Error(`Sheet database "${name}" tidak dapat dibuat.`);
  return sheet;
}

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
