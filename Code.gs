/**
 * =====================================================================
 * APLIKASI DATABASE BARANG & QUOTATION
 * Google Apps Script + Google Sheets + Google Drive
 * =====================================================================
 */

// ==================== KONFIGURASI ====================
var CONFIG = {
  SPREADSHEET_ID: '1doGHrGPfS4cEd1e5kTCf5LbVpX-DOO2uaHBFnaHJWoA', // MASUKKAN ID GOOGLE SPREADSHEET DI SINI
  ROOT_FOLDER_NAME: 'LEGEND_QUOTATION_ASSETS',
  NAMA_APLIKASI: 'Database Barang & Quotation'
};

// ==================== NAMA SHEET & HEADER ====================
var SHEET_BARANG = 'BARANG';
var SHEET_QUOTATION = 'QUOTATION';
var SHEET_QUOTATION_DETAIL = 'QUOTATION_DETAIL';
var SHEET_PENGATURAN = 'PENGATURAN';
var SHEET_JENIS_BARANG = 'JENIS_BARANG';

var HEADERS = {
  BARANG: ['id_barang', 'nama_barang', 'jenis', 'satuan', 'harga', 'created_at', 'updated_at', 'status'],
  QUOTATION: ['id_quotation', 'nomor_quotation', 'tanggal', 'tujuan', 'subtotal', 'total', 'created_at', 'updated_at'],
  QUOTATION_DETAIL: ['id_detail', 'id_quotation', 'id_barang', 'nama_barang', 'jenis', 'satuan', 'harga', 'qty', 'subtotal'],
  PENGATURAN: ['key', 'value'],
  JENIS_BARANG: ['jenis_barang']
};

var ALLOWED_MIME_TYPES = ['image/png', 'image/jpeg', 'image/jpg', 'image/webp'];
var MAX_UPLOAD_BYTES = 3 * 1024 * 1024; // 3MB

// ==================== ENTRY POINT ====================
function doGet(e) {
  try {
    initializeDatabase();
  } catch (err) {
    Logger.log(err);
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:24px;">' +
      '<h3>Konfigurasi Belum Lengkap</h3>' +
      '<p>' + escapeHtml(err.message) + '</p>' +
      '<p>Silakan isi <b>SPREADSHEET_ID</b> di bagian atas Code.gs, lalu deploy ulang.</p>' +
      '</div>'
    );
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.NAMA_APLIKASI)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== INISIALISASI DATABASE ====================
function initializeDatabase() {
  getOrCreateSheet(SHEET_BARANG, HEADERS.BARANG);
  getOrCreateSheet(SHEET_QUOTATION, HEADERS.QUOTATION);
  getOrCreateSheet(SHEET_QUOTATION_DETAIL, HEADERS.QUOTATION_DETAIL);
  getOrCreateSheet(SHEET_JENIS_BARANG, HEADERS.JENIS_BARANG);

  var pengaturanSheet = getOrCreateSheet(SHEET_PENGATURAN, HEADERS.PENGATURAN);
  if (pengaturanSheet.getLastRow() < 2) {
    var defaultKeys = [
      'nama_toko', 'nama_pemilik', 'jabatan_pemilik', 'alamat',
      'nomor_telepon', 'email', 'catatan_quotation',
      'logo_file_id', 'ttd_stempel_file_id'
    ];
    var rows = defaultKeys.map(function (k) { return [k, '']; });
    pengaturanSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  getOrCreateRootFolder();
  return { success: true, message: 'Database siap digunakan.' };
}

function getSpreadsheet() {
  if (!CONFIG.SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_ID belum diisi. Silakan isi di bagian atas Code.gs.');
  }
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateSheet(name, headers) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f4');
    forceTextColumns(sheet, name);
  } else if (sheet.getLastRow() === 0) {
    // sheet ada tapi kosong total -> pasang header, JANGAN hapus apapun (memang belum ada isi)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f4');
    forceTextColumns(sheet, name);
  }
  return sheet;
}

// Cegah Google Sheets auto-convert kolom tanggal/waktu (string ISO) jadi tipe Date,
// yang bikin serialisasi google.script.run gagal diam-diam (response jadi null di client).
var DATE_LIKE_COLUMNS = {
  BARANG: [6, 7],       // created_at, updated_at
  QUOTATION: [3, 7, 8]  // tanggal, created_at, updated_at
};
function forceTextColumns(sheet, sheetName) {
  var cols = DATE_LIKE_COLUMNS[sheetName];
  if (!cols) return;
  cols.forEach(function (col) {
    sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  });
}

function getOrCreateRootFolder() {
  var folders = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  var root = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  ['Logo', 'TandaTanganStempel', 'PDF'].forEach(function (sub) {
    getOrCreateSubFolder(root, sub);
  });
  return root;
}

function getOrCreateSubFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getAssetFolder(type) {
  var root = getOrCreateRootFolder();
  var map = { logo: 'Logo', ttdstempel: 'TandaTanganStempel', pdf: 'PDF' };
  return getOrCreateSubFolder(root, map[type]);
}

// ==================== HELPER UMUM ====================
function sheetToObjects(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function nowString() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function formatNumberID(num) {
  num = Math.round(Number(num) || 0);
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function formatTanggalIndo(dateStr) {
  if (!dateStr) return '';
  var months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  var str = String(dateStr).split('T')[0];
  var parts = str.split('-');
  if (parts.length === 3) {
    var year = parts[0];
    var monthIdx = parseInt(parts[1], 10) - 1;
    var day = parseInt(parts[2], 10);
    if (monthIdx >= 0 && monthIdx < 12) {
      return day + ' ' + months[monthIdx] + ' ' + year;
    }
  }
  return dateStr;
}

function escapeHtml(str) {
  return String(str === undefined || str === null ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function sanitizeForClient(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  if (Array.isArray(val)) return val.map(sanitizeForClient);
  if (val && typeof val === 'object') {
    var out = {};
    for (var k in val) out[k] = sanitizeForClient(val[k]);
    return out;
  }
  return val;
}

function isAktif(status) {
  // toleran: kolom status kosong/spasi/beda kapitalisasi tetap dianggap Aktif
  var s = String(status === undefined || status === null ? '' : status).trim().toLowerCase();
  return s === '' || s === 'aktif';
}

function formatRupiah(num) {
  num = Number(num) || 0;
  return 'Rp' + num.toLocaleString('id-ID', { maximumFractionDigits: 0 });
}

function getNextSequentialNumber(sheetName, colIndex) {
  var sheet = getOrCreateSheet(sheetName, HEADERS[sheetName]);
  var lastRow = sheet.getLastRow();
  var maxNum = 0;
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
    ids.forEach(function (row) {
      var id = row[0];
      if (id) {
        var parts = String(id).split('-');
        var num = parseInt(parts[parts.length - 1], 10);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    });
  }
  return maxNum + 1;
}

function generateSequentialIds(sheetName, colIndex, prefix, padLength, count) {
  var startNum = getNextSequentialNumber(sheetName, colIndex);
  var ids = [];
  for (var i = 0; i < count; i++) {
    var numStr = String(startNum + i);
    while (numStr.length < padLength) numStr = '0' + numStr;
    ids.push(prefix + '-' + numStr);
  }
  return ids;
}

function generateSequentialId(sheetName, colIndex, prefix, padLength) {
  return generateSequentialIds(sheetName, colIndex, prefix, padLength, 1)[0];
}

// ==================== DASHBOARD ====================
function getDashboardData() {
  try {
    var barang = sheetToObjects(getOrCreateSheet(SHEET_BARANG, HEADERS.BARANG));
    var aktif = barang.filter(function (b) { return isAktif(b.status); });
    var jenisSet = {};
    aktif.forEach(function (b) { if (b.jenis) jenisSet[b.jenis] = true; });

    var quotations = sheetToObjects(getOrCreateSheet(SHEET_QUOTATION, HEADERS.QUOTATION));
    var nowPrefix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    var bulanIni = quotations.filter(function (q) {
      return String(q.tanggal || '').indexOf(nowPrefix) === 0;
    });

    return {
      success: true,
      totalBarang: aktif.length,
      totalJenis: Object.keys(jenisSet).length,
      totalQuotation: quotations.length,
      quotationBulanIni: bulanIni.length
    };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal memuat data dashboard. Silakan coba lagi.' };
  }
}

// ==================== BARANG ====================
function getBarangList(filters) {
  try {
    filters = filters || {};
    var barang = sheetToObjects(getOrCreateSheet(SHEET_BARANG, HEADERS.BARANG));
    var result = barang.filter(function (b) { return isAktif(b.status); });

    if (filters.search) {
      var s = String(filters.search).toLowerCase();
      result = result.filter(function (b) { return String(b.nama_barang).toLowerCase().indexOf(s) !== -1; });
    }
    if (filters.jenis) {
      result = result.filter(function (b) { return b.jenis === filters.jenis; });
    }
    if (filters.satuan) {
      result = result.filter(function (b) { return b.satuan === filters.satuan; });
    }

    result.sort(function (a, b) {
      if (a.jenis === b.jenis) return String(a.nama_barang).localeCompare(String(b.nama_barang));
      return String(a.jenis).localeCompare(String(b.jenis));
    });

    return { success: true, data: sanitizeForClient(result) };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal memuat data barang. Silakan coba lagi.' };
  }
}

function getJenisList() {
  try {
    var sheet = getOrCreateSheet(SHEET_JENIS_BARANG, HEADERS.JENIS_BARANG);
    var list = sheetToObjects(sheet)
      .map(function (r) { return r.jenis_barang; })
      .filter(function (v) { return v; });
    list.sort();
    return { success: true, data: list };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal memuat daftar jenis barang.' };
  }
}

function addJenisBarang(nama) {
  try {
    nama = String(nama || '').trim();
    if (!nama) return { success: false, message: 'Nama jenis wajib diisi.' };
    var sheet = getOrCreateSheet(SHEET_JENIS_BARANG, HEADERS.JENIS_BARANG);
    var existing = sheetToObjects(sheet).map(function (r) { return String(r.jenis_barang).toLowerCase(); });
    if (existing.indexOf(nama.toLowerCase()) !== -1) return { success: false, message: 'Jenis "' + nama + '" sudah ada.' };
    sheet.appendRow([nama]);
    return { success: true, message: 'Jenis barang berhasil ditambahkan.' };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal menambahkan jenis barang. Silakan coba lagi.' };
  }
}

function deleteJenisBarang(nama) {
  try {
    var sheet = getOrCreateSheet(SHEET_JENIS_BARANG, HEADERS.JENIS_BARANG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Jenis tidak ditemukan.' };
    var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] === nama) {
        sheet.deleteRow(i + 2);
        return { success: true, message: 'Jenis barang berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Jenis tidak ditemukan.' };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal menghapus jenis barang. Silakan coba lagi.' };
  }
}

function saveBarang(data) {
  try {
    if (!data || !String(data.nama_barang || '').trim()) return { success: false, message: 'Nama barang wajib diisi.' };
    if (!String(data.jenis || '').trim()) return { success: false, message: 'Jenis wajib diisi.' };
    if (!String(data.satuan || '').trim()) return { success: false, message: 'Satuan wajib diisi.' };
    var harga = Number(data.harga);
    if (isNaN(harga) || harga < 0) return { success: false, message: 'Harga harus berupa angka dan tidak boleh negatif.' };

    var sheet = getOrCreateSheet(SHEET_BARANG, HEADERS.BARANG);
    var id = generateSequentialId(SHEET_BARANG, 1, 'BRG', 4);
    var now = nowString();
    sheet.appendRow([id, data.nama_barang.trim(), data.jenis.trim(), data.satuan.trim(), harga, now, now, 'Aktif']);
    return { success: true, message: 'Barang berhasil disimpan.', id: id };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal menyimpan barang. Silakan coba lagi.' };
  }
}

function updateBarang(data) {
  try {
    if (!data || !data.id_barang) return { success: false, message: 'ID barang tidak ditemukan.' };
    if (!String(data.nama_barang || '').trim() || !String(data.jenis || '').trim() || !String(data.satuan || '').trim()) {
      return { success: false, message: 'Semua field wajib diisi.' };
    }
    var harga = Number(data.harga);
    if (isNaN(harga) || harga < 0) return { success: false, message: 'Harga harus berupa angka dan tidak boleh negatif.' };

    var sheet = getOrCreateSheet(SHEET_BARANG, HEADERS.BARANG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Data barang tidak ditemukan.' };
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === data.id_barang) {
        var rowIndex = i + 2;
        sheet.getRange(rowIndex, 2, 1, 4).setValues([[data.nama_barang.trim(), data.jenis.trim(), data.satuan.trim(), harga]]);
        sheet.getRange(rowIndex, 7).setValue(nowString());
        return { success: true, message: 'Barang berhasil diperbarui.' };
      }
    }
    return { success: false, message: 'Barang tidak ditemukan.' };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal memperbarui barang. Silakan coba lagi.' };
  }
}

function deleteBarang(id) {
  try {
    if (!id) return { success: false, message: 'ID barang tidak valid.' };
    var sheet = getOrCreateSheet(SHEET_BARANG, HEADERS.BARANG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Data barang tidak ditemukan.' };
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === id) {
        var rowIndex = i + 2;
        sheet.getRange(rowIndex, 8).setValue('Nonaktif');
        sheet.getRange(rowIndex, 7).setValue(nowString());
        return { success: true, message: 'Barang berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Barang tidak ditemukan.' };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal menghapus barang. Silakan coba lagi.' };
  }
}

// ==================== QUOTATION ====================
function generateQuotationNumber(tanggalStr) {
  var sheet = getOrCreateSheet(SHEET_QUOTATION, HEADERS.QUOTATION);
  var dateCode = String(tanggalStr).replace(/-/g, '');
  var lastRow = sheet.getLastRow();
  var count = 0;
  if (lastRow >= 2) {
    var noms = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    noms.forEach(function (row) {
      if (row[0] && String(row[0]).indexOf('QUO-' + dateCode) === 0) count++;
    });
  }
  var numStr = String(count + 1);
  while (numStr.length < 3) numStr = '0' + numStr;
  return 'QUO-' + dateCode + '-' + numStr;
}

function getQuotationData(idQuotation) {
  try {
    var header = sheetToObjects(getOrCreateSheet(SHEET_QUOTATION, HEADERS.QUOTATION))
      .filter(function (q) { return q.id_quotation === idQuotation; })[0];
    if (!header) return { success: false, message: 'Quotation tidak ditemukan.' };
    var details = sheetToObjects(getOrCreateSheet(SHEET_QUOTATION_DETAIL, HEADERS.QUOTATION_DETAIL))
      .filter(function (d) { return d.id_quotation === idQuotation; });
    return { success: true, header: sanitizeForClient(header), details: sanitizeForClient(details) };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal memuat data quotation.' };
  }
}

function saveQuotation(data) {
  try {
    if (!data || !String(data.tujuan || '').trim()) return { success: false, message: 'Tujuan wajib diisi.' };
    if (!data.items || data.items.length < 1) return { success: false, message: 'Minimal harus ada satu barang dalam quotation.' };

    for (var i = 0; i < data.items.length; i++) {
      var qty = Number(data.items[i].qty);
      var harga = Number(data.items[i].harga);
      if (isNaN(qty) || qty <= 0) return { success: false, message: 'Qty barang tidak valid.' };
      if (isNaN(harga) || harga < 0) return { success: false, message: 'Harga barang tidak valid.' };
    }

    var tanggalStr = data.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var nomor = generateQuotationNumber(tanggalStr);
    var idQuotation = nomor;

    var subtotalList = data.items.map(function (it) { return Number(it.harga) * Number(it.qty); });
    var subtotal = subtotalList.reduce(function (a, b) { return a + b; }, 0);
    var total = subtotal;
    var now = nowString();

    var qSheet = getOrCreateSheet(SHEET_QUOTATION, HEADERS.QUOTATION);
    qSheet.appendRow([idQuotation, nomor, tanggalStr, data.tujuan.trim(), subtotal, total, now, now]);

    var dSheet = getOrCreateSheet(SHEET_QUOTATION_DETAIL, HEADERS.QUOTATION_DETAIL);
    var detailIds = generateSequentialIds(SHEET_QUOTATION_DETAIL, 1, 'QD', 6, data.items.length);
    var rows = data.items.map(function (it, idx) {
      return [detailIds[idx], idQuotation, it.id_barang, it.nama_barang, it.jenis, it.satuan, Number(it.harga), Number(it.qty), subtotalList[idx]];
    });
    dSheet.getRange(dSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    return { success: true, message: 'Quotation berhasil disimpan.', id_quotation: idQuotation, nomor_quotation: nomor };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal menyimpan quotation. Silakan coba lagi.' };
  }
}

// ==================== PENGATURAN ====================
function getPengaturan() {
  try {
    var sheet = getOrCreateSheet(SHEET_PENGATURAN, HEADERS.PENGATURAN);
    var rows = sheetToObjects(sheet);
    var data = {};
    rows.forEach(function (r) { data[r.key] = r.value; });

    var result = {
      nama_toko: data.nama_toko || '',
      nama_pemilik: data.nama_pemilik || '',
      jabatan_pemilik: data.jabatan_pemilik || '',
      alamat: data.alamat || '',
      nomor_telepon: data.nomor_telepon || '',
      email: data.email || '',
      catatan_quotation: data.catatan_quotation || '',
      logo_file_id: data.logo_file_id || '',
      ttd_stempel_file_id: data.ttd_stempel_file_id || ''
    };
    result.logo_data_url = result.logo_file_id ? getFileDataUrl(result.logo_file_id) : '';
    result.ttd_stempel_data_url = result.ttd_stempel_file_id ? getFileDataUrl(result.ttd_stempel_file_id) : '';

    return { success: true, data: result };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal memuat pengaturan.' };
  }
}

function getFileDataUrl(fileId) {
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (err) {
    return '';
  }
}

function savePengaturan(data) {
  try {
    if (!data) return { success: false, message: 'Data pengaturan tidak valid.' };
    var sheet = getOrCreateSheet(SHEET_PENGATURAN, HEADERS.PENGATURAN);
    var rows = sheetToObjects(sheet);
    var keyRowMap = {};
    rows.forEach(function (r, i) { keyRowMap[r.key] = i + 2; });

    var fields = ['nama_toko', 'nama_pemilik', 'jabatan_pemilik', 'alamat', 'nomor_telepon', 'email', 'catatan_quotation'];
    fields.forEach(function (key) {
      var value = data[key] !== undefined ? data[key] : '';
      if (keyRowMap[key]) {
        sheet.getRange(keyRowMap[key], 2).setValue(value);
      } else {
        sheet.appendRow([key, value]);
      }
    });
    return { success: true, message: 'Pengaturan berhasil disimpan.' };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal menyimpan pengaturan. Silakan coba lagi.' };
  }
}

// ==================== UPLOAD ASSET (LOGO/TTD/STEMPEL) ====================
function uploadAsset(base64Data, fileName, mimeType, assetType) {
  try {
    if (ALLOWED_MIME_TYPES.indexOf(mimeType) === -1) {
      return { success: false, message: 'Format file tidak didukung. Gunakan PNG, JPG, JPEG, atau WEBP.' };
    }
    var validTypes = ['logo', 'ttdstempel'];
    if (validTypes.indexOf(assetType) === -1) {
      return { success: false, message: 'Jenis asset tidak valid.' };
    }

    var bytes = Utilities.base64Decode(base64Data);
    if (bytes.length > MAX_UPLOAD_BYTES) {
      return { success: false, message: 'Ukuran file terlalu besar. Maksimal 3MB.' };
    }

    var safeExt = mimeType.split('/')[1].replace('jpeg', 'jpg');
    var blob = Utilities.newBlob(bytes, mimeType, assetType + '_' + new Date().getTime() + '.' + safeExt);
    var folder = getAssetFolder(assetType);
    var file = folder.createFile(blob);

    var keyMap = { logo: 'logo_file_id', ttdstempel: 'ttd_stempel_file_id' };
    var settingKey = keyMap[assetType];
    var sheet = getOrCreateSheet(SHEET_PENGATURAN, HEADERS.PENGATURAN);
    var rows = sheetToObjects(sheet);
    var found = false;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].key === settingKey) {
        var oldFileId = rows[i].value;
        sheet.getRange(i + 2, 2).setValue(file.getId());
        found = true;
        if (oldFileId) {
          try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) { /* abaikan */ }
        }
        break;
      }
    }
    if (!found) sheet.appendRow([settingKey, file.getId()]);

    return { success: true, message: 'File berhasil diupload.', fileId: file.getId(), dataUrl: getFileDataUrl(file.getId()) };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal mengupload file. Silakan coba lagi.' };
  }
}

// ==================== PDF: DATABASE BARANG ====================
function generateBarangPDF(filters) {
  try {
    filters = filters || {};
    var listResult = getBarangList(filters);
    if (!listResult.success) return listResult;

    var pengaturanResult = getPengaturan();
    var pengaturan = pengaturanResult.success ? pengaturanResult.data : {};

    var filterDesc = [];
    if (filters.search) filterDesc.push('Pencarian: ' + filters.search);
    if (filters.jenis) filterDesc.push('Jenis: ' + filters.jenis);
    if (filters.satuan) filterDesc.push('Satuan: ' + filters.satuan);
    var filterText = filterDesc.length ? filterDesc.join(' | ') : 'Semua Barang Aktif';
    var tanggalCetak = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    var html = buildBarangPdfHtml(listResult.data, pengaturan, filterText, tanggalCetak);
    var pdfBlob = HtmlService.createHtmlOutput(html).getAs('application/pdf')
      .setName('Database-Barang-' + new Date().getTime() + '.pdf');

    return { success: true, base64: Utilities.base64Encode(pdfBlob.getBytes()), fileName: pdfBlob.getName() };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal membuat PDF database barang. Silakan coba lagi.' };
  }
}

function buildBarangPdfHtml(items, pengaturan, filterText, tanggalCetak) {
  var logoImg = pengaturan.logo_data_url ? '<img src="' + pengaturan.logo_data_url + '" style="height:55px;">' : '';
  var rows = '';
  var no = 1;
  items.forEach(function (it) {
    rows += '<tr>' +
      '<td style="text-align:center;">' + (no++) + '</td>' +
      '<td>' + escapeHtml(it.nama_barang) + '</td>' +
      '<td>' + escapeHtml(it.jenis) + '</td>' +
      '<td>' + escapeHtml(it.satuan) + '</td>' +
      '<td style="text-align:right;">' + formatRupiah(it.harga) + '</td>' +
      '</tr>';
  });
  if (!items.length) {
    rows = '<tr><td colspan="5" style="text-align:center;">Tidak ada data.</td></tr>';
  }

  return '<html><head><meta charset="utf-8"><style>' +
    '@page { size: A4; margin: 16mm 14mm; }' +
    'body { font-family: Arial, sans-serif; font-size: 11px; color: #222; }' +
    '.header { text-align: center; margin-bottom: 8px; }' +
    '.header h2 { margin: 4px 0; font-size: 16px; }' +
    '.header h3 { margin: 2px 0; font-size: 13px; letter-spacing: 1px; }' +
    'table { width: 100%; border-collapse: collapse; margin-top: 8px; }' +
    'thead { display: table-header-group; }' +
    'tr { page-break-inside: avoid; }' +
    'th, td { border: 1px solid #999; padding: 5px 7px; font-size: 10.5px; }' +
    'th { background: #eef1f4; }' +
    '.info { margin-bottom: 6px; font-size: 10.5px; }' +
    '.footer { margin-top: 26px; font-size: 10.5px; }' +
    '</style></head><body>' +
    '<div class="header">' + logoImg +
    '<h2>' + escapeHtml(pengaturan.nama_toko || 'NAMA TOKO') + '</h2>' +
    '<h3>DAFTAR DATABASE BARANG</h3></div>' +
    '<div class="info">Tanggal cetak: ' + tanggalCetak + '<br>Filter: ' + escapeHtml(filterText) + '</div>' +
    '<table><thead><tr><th style="width:6%;">No</th><th>Nama Barang</th><th style="width:18%;">Jenis</th><th style="width:14%;">Satuan</th><th style="width:18%;">Harga</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>' +
    '<div class="footer">' + escapeHtml(pengaturan.nama_toko || '') + '<br>' + escapeHtml(pengaturan.nama_pemilik || '') +
    (pengaturan.ttd_stempel_data_url ? '<br><img src="' + pengaturan.ttd_stempel_data_url + '" style="height:55px;">' : '') + '</div>' +
    '</body></html>';
}

// ==================== PDF: QUOTATION ====================
function generateQuotationPDF(idQuotation) {
  try {
    var result = getQuotationData(idQuotation);
    if (!result.success) return result;

    var pengaturanResult = getPengaturan();
    var pengaturan = pengaturanResult.success ? pengaturanResult.data : {};

    var html = buildQuotationPdfHtml(result.header, result.details, pengaturan);
    var pdfBlob = HtmlService.createHtmlOutput(html).getAs('application/pdf')
      .setName('Quotation-' + result.header.nomor_quotation + '.pdf');

    return { success: true, base64: Utilities.base64Encode(pdfBlob.getBytes()), fileName: pdfBlob.getName() };
  } catch (err) {
    Logger.log(err);
    return { success: false, message: 'Gagal membuat PDF quotation. Silakan coba lagi.' };
  }
}

function splitBrandName(namaToko) {
  var name = String(namaToko || '').trim();
  if (!name) return { primary: 'QUOTATION', secondary: '' };
  var parts = name.split(/\s+/);
  if (parts.length === 1) {
    return { primary: parts[0].toUpperCase(), secondary: '' };
  }
  return {
    primary: parts[0].toUpperCase(),
    secondary: parts.slice(1).join(' ').toUpperCase()
  };
}

function formatRp(num) {
  return 'Rp ' + formatNumberID(num);
}

function buildQuotationNotesHtml(pengaturan) {
  if (pengaturan.catatan_quotation && String(pengaturan.catatan_quotation).trim()) {
    var cLines = String(pengaturan.catatan_quotation).replace(/\r/g, '').trim().split('\n');
    var html = '<ol class="notes-list">';
    cLines.forEach(function (line) {
      if (line.trim()) html += '<li>' + escapeHtml(line.trim()) + '</li>';
    });
    html += '</ol>';
    return html;
  }
  return '<ol class="notes-list">' +
    '<li><b>Pembayaran:</b> 50% di muka, 50% setelah barang diterima.</li>' +
    '<li><b>Pengiriman:</b> Barang akan dikirim dalam waktu 7 hari kerja setelah menerima pembayaran di muka.</li>' +
    '<li><b>Garansi:</b> 1 tahun untuk semua barang, kecuali ditentukan lain.</li>' +
    '<li><b>Masa Berlaku Penawaran:</b> Penawaran ini berlaku selama 30 hari dari tanggal penawaran.</li>' +
    '</ol>';
}

function buildQuotationPdfHtml(header, details, pengaturan) {
  pengaturan = pengaturan || {};
  header = header || {};
  details = details || [];

  var brand = splitBrandName(pengaturan.nama_toko);
  var brandHtml = '<div class="brand-logo"><div class="brand-primary">' + escapeHtml(brand.primary) + '</div>';
  if (brand.secondary) {
    brandHtml += '<div class="brand-secondary">' + escapeHtml(brand.secondary) + '</div>';
  }
  brandHtml += '</div>';

  var tokoInfoParts = [];
  if (pengaturan.alamat) tokoInfoParts.push(escapeHtml(pengaturan.alamat));
  var telpEmail = [];
  if (pengaturan.nomor_telepon) telpEmail.push('Telp: ' + escapeHtml(pengaturan.nomor_telepon));
  if (pengaturan.email) telpEmail.push(escapeHtml(pengaturan.email));
  if (telpEmail.length) tokoInfoParts.push(telpEmail.join(' | '));
  var tokoInfoHtml = tokoInfoParts.join('<br>');

  var tujuanLines = String(header.tujuan || '').replace(/\r/g, '').trim().split('\n');
  var customerName = escapeHtml(tujuanLines[0] || '-');
  var customerExtraParts = [];
  for (var t = 1; t < tujuanLines.length; t++) {
    if (tujuanLines[t].trim()) customerExtraParts.push(escapeHtml(tujuanLines[t]));
  }
  var customerExtraHtml = customerExtraParts.length
    ? '<div class="customer-extra">' + customerExtraParts.join('<br>') + '</div>'
    : '';

  var rows = '';
  var subtotal = 0;
  details.forEach(function (d, idx) {
    var itemSubtotal = Number(d.subtotal) || (Number(d.harga) * Number(d.qty));
    subtotal += itemSubtotal;
    rows += '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td><div class="item-name">' + escapeHtml(d.nama_barang) + '</div></td>' +
      '<td style="text-align:center;">' + escapeHtml(d.qty) + '</td>' +
      '<td style="text-align:right;">' + formatRp(d.harga) + '</td>' +
      '<td style="text-align:right;">' + formatRp(itemSubtotal) + '</td>' +
      '</tr>';
  });
  if (!details.length) {
    rows = '<tr><td colspan="5" style="text-align:center;color:#888;">Tidak ada item.</td></tr>';
  }

  var tanggalFormatted = formatTanggalIndo(header.tanggal);

  var ttdImg = pengaturan.ttd_stempel_data_url
    ? '<img src="' + pengaturan.ttd_stempel_data_url + '" class="signature-image">'
    : '';

  var watermarkHtml = pengaturan.logo_data_url
    ? '<div class="watermark-container"><img src="' + pengaturan.logo_data_url + '" class="watermark-logo" alt=""></div>'
    : '';

  var footerParts = [];
  if (pengaturan.alamat) footerParts.push(escapeHtml(pengaturan.alamat));
  if (pengaturan.nomor_telepon) footerParts.push('Telp: ' + escapeHtml(pengaturan.nomor_telepon));
  var footerText = footerParts.length ? footerParts.join('  |  ') : escapeHtml(pengaturan.nama_toko || '');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    '@page { size: A4; margin: 0; }' +
    'body { font-family: Helvetica, Arial, sans-serif; color: #2b2b2b; margin: 0; padding: 2.5cm 40px 2.5cm 40px; background-color: #ffffff; font-size: 13px; line-height: 1.4; }' +
    '.header { display: table; width: 100%; margin-bottom: 30px; }' +
    '.header-left { display: table-cell; width: 60%; vertical-align: top; }' +
    '.header-right { display: table-cell; width: 40%; vertical-align: top; text-align: right; }' +
    '.brand-logo { margin-bottom: 8px; }' +
    '.brand-primary { font-size: 24px; font-weight: 800; letter-spacing: 3px; color: #1a1a1a; line-height: 1; }' +
    '.brand-secondary { font-size: 13px; font-weight: 600; letter-spacing: 6px; color: #e3b853; margin-top: 2px; line-height: 1.2; }' +
    '.toko-info { font-size: 12px; color: #666666; margin: 0; line-height: 1.4; }' +
    '.invoice-title { font-size: 28px; font-weight: 500; letter-spacing: 2px; color: #c79850; margin: 0; text-transform: uppercase; }' +
    '.row-info { display: table; width: 100%; margin-bottom: 25px; }' +
    '.col-customer { display: table-cell; width: 55%; vertical-align: top; }' +
    '.col-meta { display: table-cell; width: 45%; vertical-align: top; }' +
    '.label-title { font-weight: bold; color: #1a1a1a; margin-bottom: 4px; font-size: 13px; }' +
    '.customer-name { font-size: 14px; font-weight: bold; color: #1a1a1a; }' +
    '.customer-extra { color: #555; margin-top: 2px; font-size: 12px; }' +
    '.meta-table { width: 100%; border-collapse: collapse; }' +
    '.meta-table td { padding: 2px 0; font-size: 13px; }' +
    '.meta-table td.meta-label { font-weight: bold; color: #1a1a1a; width: 50%; }' +
    '.meta-table td.meta-val { text-align: right; color: #333; }' +
    'table.items { width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 25px; }' +
    'table.items th { border-top: 1.5px solid #2c2c2c; border-bottom: 1.5px solid #2c2c2c; color: #1a1a1a; padding: 10px 6px; font-weight: bold; text-align: left; font-size: 12px; }' +
    'table.items td { border-bottom: 1px solid #e0e0e0; padding: 10px 6px; font-size: 12px; color: #333333; vertical-align: top; }' +
    'thead { display: table-header-group; }' +
    'tr { page-break-inside: avoid; }' +
    '.item-name { font-weight: bold; color: #1a1a1a; }' +
    '.item-sub { font-size: 11px; color: #888; margin-top: 2px; }' +
    '.row-bottom { display: table; width: 100%; table-layout: fixed; margin-top: 15px; }' +
    '.col-notes { display: table-cell; width: 55%; vertical-align: top; padding-right: 25px; }' +
    '.col-summary { display: table-cell; width: 45%; vertical-align: top; }' +
    '.payment-block { margin-bottom: 15px; }' +
    '.payment-content { min-height: 60px; }' +
    '.payment-block strong { display: block; font-size: 12px; text-transform: uppercase; color: #1a1a1a; letter-spacing: 0.5px; margin-bottom: 4px; }' +
    '.notes-list { margin: 0; padding-left: 18px; font-size: 12px; color: #555; line-height: 1.5; }' +
    '.notes-list li { margin-bottom: 3px; }' +
    '.summary-table { width: 100%; border-collapse: collapse; }' +
    '.summary-table td { padding: 4px 0; font-size: 13px; }' +
    '.summary-table td.sum-label { font-weight: bold; color: #1a1a1a; text-align: right; padding-right: 15px; }' +
    '.summary-table td.sum-val { text-align: right; color: #1a1a1a; }' +
    '.grand-total-row td { border-top: 1.5px solid #2c2c2c; padding-top: 8px !important; font-size: 16px !important; font-weight: bold; }' +
    '.signature-section { margin-top: 40px; display: table; width: 100%; page-break-inside: avoid; }' +
    '.thanks-note { display: table-cell; width: 50%; vertical-align: bottom; font-size: 14px; font-weight: bold; color: #2c2c2c; }' +
    '.signature-group { display: table-cell; width: 50%; text-align: right; vertical-align: bottom; }' +
    '.signature-box { display: inline-block; width: 220px; text-align: center; margin-left: 15px; vertical-align: top; overflow: visible; position: relative; }' +
    '.signature-image { width: 200px; height: auto; display: block; position: absolute; left: 8px; top: -35px; z-index: 10; }' +
    '.owner-name { margin-top: 60px; font-weight: bold; font-size: 15px; color: #1a1a1a; position: relative; z-index: 1; }' +
    '.signature-title { font-size: 11px; color: #666; margin-top: 3px; }' +
    '.watermark-container { position: fixed; bottom: 35px; left: 0; right: 0; width: 100%; text-align: center; z-index: 9; }' +
    '.watermark-logo { width: 200px; max-width: 300px; max-height: 70px; height: auto; opacity: 1; }' +
    '.footer-banner { position: fixed; bottom: 0; left: 0; right: 0; background-color: #e3b853; color: #ffffff; text-align: center; padding: 8px 20px; font-size: 11px; letter-spacing: 0.5px; z-index: 10; }' +
    '</style></head><body>' +
    '<div class="header">' +
      '<div class="header-left">' +
        brandHtml +
        '<p class="toko-info">' + tokoInfoHtml + '</p>' +
      '</div>' +
      '<div class="header-right">' +
        '<div class="invoice-title">QUOTATION</div>' +
      '</div>' +
    '</div>' +
    '<div class="row-info">' +
      '<div class="col-customer">' +
        '<div class="label-title">Quotation to:</div>' +
        '<div class="customer-name">' + customerName + '</div>' +
        customerExtraHtml +
      '</div>' +
      '<div class="col-meta">' +
        '<table class="meta-table">' +
          '<tr><td class="meta-label">Quotation#</td><td class="meta-val">' + escapeHtml(header.nomor_quotation || '-') + '</td></tr>' +
          '<tr><td class="meta-label">Date</td><td class="meta-val">' + escapeHtml(tanggalFormatted) + '</td></tr>' +
        '</table>' +
      '</div>' +
    '</div>' +
    '<table class="items">' +
      '<thead><tr>' +
        '<th width="8%">No</th>' +
        '<th width="42%">Deskripsi Barang</th>' +
        '<th width="12%" style="text-align:center;">Qty</th>' +
        '<th width="18%" style="text-align:right;">Harga Satuan</th>' +
        '<th width="20%" style="text-align:right;">Total</th>' +
      '</tr></thead>' +
      '<tbody>' + rows + '</tbody>' +
    '</table>' +
    '<div class="row-bottom">' +
      '<div class="col-notes">' +
        '<div class="payment-block">' +
          '<strong>Syarat &amp; Ketentuan</strong>' +
          '<div class="payment-content">' + buildQuotationNotesHtml(pengaturan) + '</div>' +
        '</div>' +
      '</div>' +
      '<div class="col-summary">' +
        '<table class="summary-table">' +
          '<tr><td class="sum-label">Subtotal</td><td class="sum-val">' + formatRp(subtotal) + '</td></tr>' +
          '<tr class="grand-total-row"><td class="sum-label">Total</td><td class="sum-val">' + formatRp(subtotal) + '</td></tr>' +
        '</table>' +
      '</div>' +
    '</div>' +
    '<div class="signature-section">' +
      '<div class="thanks-note">Thank you for your business!</div>' +
      '<div class="signature-group">' +
        '<div class="signature-box">' +
          ttdImg +
          '<div class="owner-name">' + escapeHtml(pengaturan.nama_pemilik || '') + '</div>' +
          '<div class="signature-title">' + escapeHtml(pengaturan.jabatan_pemilik || '') + '</div>' +
        '</div>' +
      '</div>' +
    '</div>' +
    watermarkHtml +
    '<div class="footer-banner">' + footerText + '</div>' +
    '</body></html>';
}
