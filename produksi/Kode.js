const CONFIG = {
  SOURCE_1: {
    name: 'WebData_Produk',
    spreadsheetId: '1P34UU9Q1pN4afM9L8Mn23LD90QdnbQ7Y',
    sheetName: 'WebData_Produk',
    range: 'A1:Z',
    firebasePath: 'produksi/webdata_produk',
    targetPlant: 'cp3'
  },
  SOURCE_2: {
    name: 'DASHBOARD_KHUSUS',
    spreadsheetId: '1OrmtFMggqx0j5uW_X5Nxfo23ty61F7m6ppKWEPvPu0s',
    sheetName: 'DASHBOARD KHUSUS',
    startRow: 2,
    range: 'A2:Z',
    firebasePath: 'produksi/dashboard_khusus'
  },
  TARGET: {
    spreadsheetId: '146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4',
    targetSheetGid: 1192757151,
    WRITE_TO_SEPARATE_TABS: false,
    CLEARANCE_COLS: 4
  },
  FIREBASE: {
    DATABASE_URL: 'https://project-produksi-default-rtdb.firebaseio.com',
    AUTH_SECRET: '',
    ENABLED: false
  }
};

function getCurrentEmail() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'unknown';
  } catch (e) {
    return 'unknown';
  }
}

function mintaIzinAkses() {
  DriveApp.getRootFolder();
  SpreadsheetApp.openById(CONFIG.TARGET.spreadsheetId);
  Logger.log('Izin Google Drive & Spreadsheet berhasil diberikan.');
}

function syncAllData() {
  const currentEmail = getCurrentEmail();
  Logger.log(`Menjalankan sinkronisasi dengan akun: ${currentEmail}`);

  const raw1 = fetchSheetDataWithFallback(CONFIG.SOURCE_1);
  const data1 = filterSource1(raw1, CONFIG.SOURCE_1.targetPlant);

  const raw2 = fetchSheetDataWithFallback(CONFIG.SOURCE_2);
  const data2 = processSource2WithHeaders(raw2);

  if (data1.length === 0 && data2.length === 0) {
    Logger.log('Tidak ada data yang berhasil diambil dari kedua sumber. Pastikan izin akses file sudah dibuka.');
    return { success: false, message: 'Semua sumber data gagal diakses' };
  }

  writeToTargetSpreadsheet(data1, data2);

  if (CONFIG.FIREBASE.ENABLED && CONFIG.FIREBASE.DATABASE_URL) {
    if (data1.length > 0) syncToFirebase(CONFIG.SOURCE_1.firebasePath, data1);
    if (data2.length > 0) syncToFirebase(CONFIG.SOURCE_2.firebasePath, data2, true);

    syncToFirebase('produksi/last_sync', {
      timestamp: new Date().toISOString(),
      source1_rows: data1.length,
      source2_rows: data2.length,
      status: 'SUCCESS'
    });
  }

  return { success: true };
}

function filterSource1(matrix, targetPlant) {
  if (!matrix || matrix.length === 0) return [];
  const header = matrix[0];
  const target = String(targetPlant || 'cp3').toLowerCase().replace(/\s+/g, '');
  const filtered = matrix.slice(1).filter(row => {
    const plant = String(row[2] || '').trim().toLowerCase().replace(/\s+/g, '');
    return plant === target;
  });
  return [header, ...filtered];
}

function processSource2WithHeaders(matrix) {
  if (!matrix || matrix.length < 5) return matrix;

  const headerRows = 4;

  for (let r = 0; r < headerRows; r++) {
    if (!matrix[r][0]) matrix[r][0] = 'ISOWEEK';
    if (!matrix[r][1]) matrix[r][1] = 'TANGGAL';
  }

  for (let r = 0; r < headerRows; r++) {
    let lastVal = '';
    for (let c = 2; c < matrix[r].length; c++) {
      const val = String(matrix[r][c] || '').trim();
      if (val !== '') {
        lastVal = val;
      } else if (lastVal !== '') {
        matrix[r][c] = lastVal;
      }
    }
  }

  return matrix;
}

function fetchSheetDataWithFallback(sourceConfig) {
  try {
    const ss = SpreadsheetApp.openById(sourceConfig.spreadsheetId);
    const sheet = ss.getSheetByName(sourceConfig.sheetName);
    if (!sheet) {
      throw new Error(`Sheet '${sourceConfig.sheetName}' tidak ditemukan.`);
    }

    let values;
    if (sourceConfig.startRow && sourceConfig.startRow > 1) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow >= sourceConfig.startRow && lastCol > 0) {
        const numRows = lastRow - sourceConfig.startRow + 1;
        values = sheet.getRange(sourceConfig.startRow, 1, numRows, lastCol).getValues();
      } else {
        values = [];
      }
    } else {
      const range = sheet.getRange(sourceConfig.range);
      values = range.getValues();
    }

    const cleaned = clean2DArray(values);
    if (cleaned.length > 0) {
      Logger.log(`[BERHASIL] ${sourceConfig.name} via SpreadsheetApp: ${cleaned.length} baris`);
      return cleaned;
    }
  } catch (errNative) {
    Logger.log(`SpreadsheetApp gagal untuk ${sourceConfig.name}: ${errNative.message}`);
  }

  const encodedSheet = encodeURIComponent(sourceConfig.sheetName);

  try {
    const exportUrl = `https://docs.google.com/spreadsheets/d/${sourceConfig.spreadsheetId}/export?format=csv&sheet=${encodedSheet}`;
    const resExport = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0'
      }
    });

    if (resExport.getResponseCode() === 200) {
      let parsed = Utilities.parseCsv(resExport.getContentText());
      if (sourceConfig.startRow && sourceConfig.startRow > 1 && parsed.length >= sourceConfig.startRow) {
        parsed = parsed.slice(sourceConfig.startRow - 1);
      }
      const cleaned = clean2DArray(parsed);
      if (cleaned.length > 0) {
        Logger.log(`[BERHASIL] ${sourceConfig.name} via Export CSV: ${cleaned.length} baris`);
        return cleaned;
      }
    }
  } catch (errExport) {
    Logger.log(`Export CSV gagal untuk ${sourceConfig.name}: ${errExport.message}`);
  }

  try {
    const gvizRange = sourceConfig.range ? `&range=${sourceConfig.range}` : '';
    const gvizUrl = `https://docs.google.com/spreadsheets/d/${sourceConfig.spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodedSheet}${gvizRange}`;
    const token = ScriptApp.getOAuthToken();
    const resGviz = UrlFetchApp.fetch(gvizUrl, {
      muteHttpExceptions: true,
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });

    if (resGviz.getResponseCode() === 200) {
      let parsed = Utilities.parseCsv(resGviz.getContentText());
      if (sourceConfig.startRow && sourceConfig.startRow > 1 && !sourceConfig.range && parsed.length >= sourceConfig.startRow) {
        parsed = parsed.slice(sourceConfig.startRow - 1);
      }
      const cleaned = clean2DArray(parsed);
      if (cleaned.length > 0) {
        Logger.log(`[BERHASIL] ${sourceConfig.name} via GViz Bearer: ${cleaned.length} baris`);
        return cleaned;
      }
    } else {
      Logger.log(`GViz untuk ${sourceConfig.name} gagal dengan status HTTP ${resGviz.getResponseCode()}`);
    }
  } catch (errGviz) {
    Logger.log(`GViz Bearer gagal untuk ${sourceConfig.name}: ${errGviz.message}`);
  }

  Logger.log(`[GAGAL TOTAL] Tidak dapat membaca data dari ${sourceConfig.name} (${sourceConfig.spreadsheetId}). Periksa izin share file tersebut.`);
  return [];
}

function writeToTargetSpreadsheet(data1, data2) {
  let targetSS;
  try {
    targetSS = SpreadsheetApp.openById(CONFIG.TARGET.spreadsheetId);
  } catch (e) {
    const activeEmail = getCurrentEmail();
    Logger.log(`[ERROR TARGET] Tidak dapat membuka Spreadsheet Target (${CONFIG.TARGET.spreadsheetId}).`);
    Logger.log(`Pastikan file target di Google Drive telah dibagikan dengan hak 'Editor' ke akun: ${activeEmail}`);
    throw new Error(`Akses ditolak pada Spreadsheet Target. Berikan hak akses 'Editor' ke ${activeEmail}`);
  }

  if (CONFIG.TARGET.WRITE_TO_SEPARATE_TABS) {
    writeMatrixToNamedSheet(targetSS, CONFIG.SOURCE_1.name, data1);
    writeMatrixToNamedSheet(targetSS, CONFIG.SOURCE_2.name, data2);
  } else {
    let targetSheet = null;
    const sheets = targetSS.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === CONFIG.TARGET.targetSheetGid) {
        targetSheet = sheets[i];
        break;
      }
    }

    if (!targetSheet) {
      targetSheet = sheets[0];
    }

    targetSheet.clearContents();

    const clearanceCols = CONFIG.TARGET.CLEARANCE_COLS || 4;
    let colSumber2 = 1;

    if (data1.length > 0) {
      targetSheet.getRange(2, 1).setValue(`SUMBER 1: ${CONFIG.SOURCE_1.name} (Khusus Plant CP 3) | Update: ${new Date().toLocaleString('id-ID')}`);
      targetSheet.getRange(2, 1).setFontWeight('bold').setFontColor('#0b57d0');

      targetSheet.getRange(3, 1, data1.length, data1[0].length).setValues(data1);
      targetSheet.getRange(3, 1, 1, data1[0].length).setBackground('#cfe2f3').setFontWeight('bold');

      colSumber2 = data1[0].length + 1 + clearanceCols;
    }

    if (data2.length > 0) {
      targetSheet.getRange(2, colSumber2).setValue(`SUMBER 2: ${CONFIG.SOURCE_2.name} (Lengkap Header & Terisi) | Update: ${new Date().toLocaleString('id-ID')}`);
      targetSheet.getRange(2, colSumber2).setFontWeight('bold').setFontColor('#0b57d0');

      targetSheet.getRange(3, colSumber2, data2.length, data2[0].length).setValues(data2);

      if (data2.length >= 4) {
        targetSheet.getRange(3, colSumber2, 1, data2[0].length).setBackground('#fff2a3').setFontWeight('bold');
        targetSheet.getRange(4, colSumber2, 1, data2[0].length).setBackground('#ffe599').setFontWeight('bold');
        targetSheet.getRange(5, colSumber2, 1, data2[0].length).setBackground('#d9ead3').setFontWeight('bold');
        targetSheet.getRange(6, colSumber2, 1, data2[0].length).setBackground('#cfe2f3').setFontWeight('bold');
      }
    }
  }
}

function writeMatrixToNamedSheet(spreadsheet, sheetName, data) {
  if (!data || data.length === 0) return;
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function syncToFirebase(endpointPath, data, isMultiHeader) {
  try {
    let cleanBaseUrl = CONFIG.FIREBASE.DATABASE_URL.replace(/\/+$/, '');
    let url = `${cleanBaseUrl}/${endpointPath.replace(/^\/+/, '')}.json`;

    if (CONFIG.FIREBASE.AUTH_SECRET) {
      url += `?auth=${CONFIG.FIREBASE.AUTH_SECRET}`;
    }

    const payload = isMultiHeader
      ? convertMultiHeader2DArrayToJson(data)
      : convert2DArrayToJsonObjects(data);

    const options = {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch(url, options);
  } catch (err) {
    Logger.log(err.message);
  }
}

function convertMultiHeader2DArrayToJson(matrix) {
  if (!matrix || matrix.length <= 4) return matrix;

  const rowProd = matrix[0];
  const rowKat = matrix[1];
  const rowShift = matrix[2];
  const rowSatuan = matrix[3];

  const headers = [];
  for (let c = 0; c < rowProd.length; c++) {
    if (c === 0) {
      headers.push('ISOWEEK');
    } else if (c === 1) {
      headers.push('TANGGAL');
    } else {
      const parts = [rowProd[c], rowKat[c], rowShift[c], rowSatuan[c]]
        .map(p => String(p || '').trim().replace(/[\.\$#\[\]\/\s+]/g, '_'))
        .filter(p => p !== '');
      headers.push(parts.join('_') || `col_${c + 1}`);
    }
  }

  const result = [];
  for (let r = 4; r < matrix.length; r++) {
    const row = matrix[r];
    const item = {};
    let hasData = false;
    for (let c = 0; c < headers.length; c++) {
      const val = row[c] !== undefined ? row[c] : '';
      item[headers[c]] = val;
      if (val !== '' && val !== null) hasData = true;
    }
    if (hasData) {
      result.push(item);
    }
  }
  return result;
}

function convert2DArrayToJsonObjects(matrix) {
  if (!matrix || matrix.length <= 1) return matrix;

  const headers = matrix[0].map((h, idx) => {
    let key = String(h || '').trim();
    key = key.replace(/[\.\$#\[\]\/]/g, '_');
    return key || `col_${idx + 1}`;
  });

  const result = [];
  for (let r = 1; r < matrix.length; r++) {
    const row = matrix[r];
    const item = {};
    let hasData = false;
    for (let c = 0; c < headers.length; c++) {
      const val = row[c] !== undefined ? row[c] : '';
      item[headers[c]] = val;
      if (val !== '' && val !== null) hasData = true;
    }
    if (hasData) {
      result.push(item);
    }
  }
  return result;
}

function clean2DArray(matrix) {
  if (!matrix || matrix.length === 0) return [];

  let lastNonEmptyRow = -1;
  for (let r = matrix.length - 1; r >= 0; r--) {
    const hasVal = matrix[r].some(cell => cell !== '' && cell !== null && cell !== undefined);
    if (hasVal) {
      lastNonEmptyRow = r;
      break;
    }
  }

  if (lastNonEmptyRow === -1) return [];

  const trimmedRows = matrix.slice(0, lastNonEmptyRow + 1);

  let maxCol = 0;
  for (let r = 0; r < trimmedRows.length; r++) {
    for (let c = trimmedRows[r].length - 1; c >= 0; c--) {
      const val = trimmedRows[r][c];
      if (val !== '' && val !== null && val !== undefined) {
        if (c + 1 > maxCol) maxCol = c + 1;
        break;
      }
    }
  }

  if (maxCol === 0) return [];

  return trimmedRows.map(row => {
    const newRow = row.slice(0, maxCol);
    while (newRow.length < maxCol) newRow.push('');
    return newRow;
  });
}

function setupTrigger15Min() {
  removeExistingTriggers();
  ScriptApp.newTrigger('syncAllData')
    .timeBased()
    .everyMinutes(15)
    .create();
}

function setupTrigger1Hour() {
  removeExistingTriggers();
  ScriptApp.newTrigger('syncAllData')
    .timeBased()
    .everyHours(1)
    .create();
}

function removeExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncAllData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
