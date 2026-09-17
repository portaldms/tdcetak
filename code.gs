/**
 * SILAP-BMN / SILAP-KDF - Backend Web API (Code.gs)
 * Version: 2.7 (Enhanced CRUD with Master Row Index & Dual GET/POST Support)
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var payload = {};
    
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        payload = {};
      }
    }

    if (params.data) {
      try {
        var parsedData = JSON.parse(params.data);
        for (var key in parsedData) {
          payload[key] = parsedData[key];
        }
      } catch(dErr) {}
    }

    var action = params.action || payload.action || 'getDatabaseData';
    var forceRefresh = (params.nocache === '1' || payload.nocache === 1);
    var result;

    switch (action) {
      case 'getDatabaseData':
        result = getDatabaseData(forceRefresh);
        break;
      case 'addTransaksiKeluar':
        result = addTransaksiKeluar(payload);
        break;
      case 'updateTransaksi':
        result = updateTransaksi(payload);
        break;
      case 'deleteTransaksi':
        result = deleteTransaksi(payload);
        break;
      default:
        result = { success: false, error: 'Aksi tidak dikenal: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getActiveSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('DATA_BMN') || ss.getSheets()[0];
}

function formatRupiah(number) {
  if (isNaN(number) || number === null) return "Rp 0";
  return "Rp " + Math.round(number).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function formatDate(dateVal) {
  if (!dateVal) return '';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(dateVal);
}

function fixDriveUrl(url) {
  if (!url || typeof url !== 'string') return '';
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return "https://lh3.googleusercontent.com/d/" + match[1];
  }
  return url;
}

function getDatabaseData(forceRefresh) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'SILAP_BMN_DB_DATA_V4';

    if (forceRefresh) {
      cache.remove(cacheKey);
    } else {
      var cached = cache.get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    }

    var sheet = getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, items: [] };
    }

    var rawData = sheet.getRange(2, 1, lastRow - 1, 26).getValues();
    var itemsList = [];
    var itemMap = {};
    var lastMasterItem = null;

    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      var actualRowIndex = i + 2;

      var no = row[0];
      var kodeSatker = row[1];
      var kodeBarang = row[2];
      var nup = String(row[3] || '').trim();
      var namaBarang = String(row[4] || '').trim();
      var tahun = row[5];
      var nilaiPerolehan = row[6];
      var merk = row[7];
      var tipe = row[8];
      var kondisi = row[9];
      var jenisBmn = row[10];
      var nopol = String(row[11] || '').trim();
      var linkDokumen = row[12];
      var foto = row[13];

      var currentItem = null;

      // Jika baris berisi data Master baru (Nama Barang / NUP / Nopol terisi)
      if (namaBarang !== "" || nup !== "" || nopol !== "") {
        var itemId = "ITEM_" + (nup || ('IDX_' + i)) + "_" + nopol.replace(/\s+/g, '');
        
        if (!itemMap[itemId]) {
          var numericNilai = 0;
          if (typeof nilaiPerolehan === 'number') {
            numericNilai = nilaiPerolehan;
          } else if (typeof nilaiPerolehan === 'string') {
            numericNilai = parseFloat(nilaiPerolehan.replace(/[^0-9.-]+/g, "")) || 0;
          }

          currentItem = {
            id: itemId,
            masterRowIndex: actualRowIndex,
            no: no || (itemsList.length + 1),
            kodeSatker: String(kodeSatker || '-'),
            kodeBarang: String(kodeBarang || '-'),
            nup: nup || '-',
            namaBarang: namaBarang || '-',
            tahun: String(tahun || '-'),
            nilaiPerolehan: numericNilai,
            nilaiPerolehanFormatted: formatRupiah(numericNilai),
            merk: String(merk || '-'),
            tipe: String(tipe || '-'),
            kondisi: String(kondisi || 'Baik'),
            jenisBmn: String(jenisBmn || '-'),
            nopol: nopol || '-',
            linkDokumen: String(linkDokumen || ''),
            foto: fixDriveUrl(foto),
            history: []
          };

          itemMap[itemId] = currentItem;
          itemsList.push(currentItem);
          lastMasterItem = currentItem;
        } else {
          currentItem = itemMap[itemId];
          lastMasterItem = currentItem;
        }
      } else {
        // Jika kolom master kosong, hubungkan ke master item terakhir yang aktif
        currentItem = lastMasterItem;
      }

      // Olah Kolom Transaksi (O-Z: Kolom Index 14 s/d 25)
      if (currentItem) {
        var tglKeluar = row[14] ? formatDate(row[14]) : '';
        var tglMasuk = row[20] ? formatDate(row[20]) : '';

        if (tglKeluar || row[15] || tglMasuk || row[21]) {
          currentItem.history.push({
            rowIndex: actualRowIndex,
            trxNo: currentItem.history.length + 1,
            keluarTgl: tglKeluar || '-',
            keluarPetugasSerah: String(row[15] || '-'),
            keluarPetugasTerima: String(row[16] || '-'),
            keluarBA: String(row[17] || '-'),
            keluarFoto: fixDriveUrl(row[18]),
            keluarFotoRaw: String(row[18] || ''),
            keluarLokasi: String(row[19] || '-'),
            masukTgl: tglMasuk || '-',
            masukPetugasTerima: String(row[21] || '-'),
            masukPetugasSerah: String(row[22] || '-'),
            masukBA: String(row[23] || '-'),
            masukFoto: fixDriveUrl(row[24]),
            masukFotoRaw: String(row[24] || ''),
            masukLokasi: String(row[25] || '-')
          });
        }
      }
    }

    var result = { success: true, items: itemsList };
    cache.put(cacheKey, JSON.stringify(result), 300); // Cache 5 Menit
    return result;

  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addTransaksiKeluar(payload) {
  try {
    var sheet = getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, error: "Data sheet kosong." };
    }

    var targetMasterIndex = -1;
    if (payload.masterRowIndex && !isNaN(parseInt(payload.masterRowIndex))) {
      targetMasterIndex = parseInt(payload.masterRowIndex);
    }

    if (targetMasterIndex === -1) {
      var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
      var targetNup = String(payload.nup || '').trim().toLowerCase();
      var targetNopol = String(payload.nopol || '').trim().toLowerCase();

      for (var i = 0; i < data.length; i++) {
        var rowNup = String(data[i][3] || '').trim().toLowerCase();
        var rowNopol = String(data[i][11] || '').trim().toLowerCase();
        
        if ((targetNup !== '' && targetNup !== '-' && rowNup === targetNup) || 
            (targetNopol !== '' && targetNopol !== '-' && rowNopol === targetNopol)) {
          targetMasterIndex = i + 2;
          break;
        }
      }
    }

    if (targetMasterIndex === -1 || targetMasterIndex > lastRow + 1) {
      return { success: false, error: "Master data barang tidak ditemukan di Spreadsheet." };
    }

    // Periksa apakah baris master kolom O-Z (trx) masih kosong
    var masterTrxValues = sheet.getRange(targetMasterIndex, 15, 1, 12).getValues()[0];
    var isMasterTrxEmpty = masterTrxValues.every(function(val) { return val === "" || val === null; });

    var targetRowNumber;
    if (isMasterTrxEmpty) {
      targetRowNumber = targetMasterIndex;
    } else {
      var dataFull = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 12).getValues();
      var insertRowIndex = targetMasterIndex;

      for (var j = targetMasterIndex - 1; j < dataFull.length; j++) {
        var hasMasterInfo = (String(dataFull[j][4] || '').trim() !== '' || 
                             String(dataFull[j][3] || '').trim() !== '' || 
                             String(dataFull[j][11] || '').trim() !== '');
        if (hasMasterInfo && (j + 2 !== targetMasterIndex)) {
          break;
        }
        insertRowIndex = j + 2;
      }

      sheet.insertRowAfter(insertRowIndex);
      targetRowNumber = insertRowIndex + 1;
    }

    var trxValues = [[
      payload.keluarTgl || '',
      payload.keluarPetugasSerah || '',
      payload.keluarPetugasTerima || '',
      payload.keluarBA || '',
      payload.keluarFoto || '',
      payload.keluarLokasi || '',
      payload.masukTgl || '',
      payload.masukPetugasTerima || '',
      payload.masukPetugasSerah || '',
      payload.masukBA || '',
      payload.masukFoto || '',
      payload.masukLokasi || ''
    ]];

    sheet.getRange(targetRowNumber, 15, 1, 12).setValues(trxValues);
    CacheService.getScriptCache().remove('SILAP_BMN_DB_DATA_V4');
    return { success: true, message: "Transaksi riwayat berhasil ditambahkan!" };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function updateTransaksi(payload) {
  try {
    var rowIndex = parseInt(payload.rowIndex);
    if (!rowIndex || rowIndex < 2) {
      return { success: false, error: "Index baris transaksi tidak valid." };
    }

    var sheet = getActiveSheet();
    var trxValues = [[
      payload.keluarTgl || '',
      payload.keluarPetugasSerah || '',
      payload.keluarPetugasTerima || '',
      payload.keluarBA || '',
      payload.keluarFoto || '',
      payload.keluarLokasi || '',
      payload.masukTgl || '',
      payload.masukPetugasTerima || '',
      payload.masukPetugasSerah || '',
      payload.masukBA || '',
      payload.masukFoto || '',
      payload.masukLokasi || ''
    ]];

    sheet.getRange(rowIndex, 15, 1, 12).setValues(trxValues);
    CacheService.getScriptCache().remove('SILAP_BMN_DB_DATA_V4');

    return { success: true, message: "Transaksi berhasil diperbarui!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteTransaksi(payload) {
  try {
    var rowIndex = parseInt(payload.rowIndex);
    if (!rowIndex || rowIndex < 2) {
      return { success: false, error: "Index baris tidak valid." };
    }

    var sheet = getActiveSheet();
    sheet.getRange(rowIndex, 15, 1, 12).clearContent();
    CacheService.getScriptCache().remove('SILAP_BMN_DB_DATA_V4');

    return { success: true, message: "Transaksi riwayat berhasil dihapus." };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
