const CONFIG = {
  TARGET_SPREADSHEET_ID: '146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4',
  WRITE_TO_SEPARATE_TABS: true,
  SOURCES: [
    {
      id: 'webdata_produk',
      name: 'WebData_Produk',
      spreadsheetId: '1P34UU9Q1pN4afM9L8Mn23LD90QdnbQ7Y',
      sheetName: 'WebData_Produk',
      range: 'A1:Z',
      targetTab: 'WebData_Produk',
      filterColIndex: 2,
      filterValue: 'cp3'
    },
    {
      id: 'dashboard_khusus',
      name: 'DASHBOARD_KHUSUS',
      spreadsheetId: '1OrmtFMggqx0j5uW_X5Nxfo23ty61F7m6ppKWEPvPu0s',
      sheetName: 'DASHBOARD KHUSUS',
      startRow: 2,
      range: 'A2:Z',
      targetTab: 'DASHBOARD KHUSUS',
      isMultiHeader: true,
      headerRows: 4
    }
  ]
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
  SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  for (let i = 0; i < CONFIG.SOURCES.length; i++) {
    try {
      SpreadsheetApp.openById(CONFIG.SOURCES[i].spreadsheetId);
    } catch (e) {}
  }
  Logger.log('Izin Google Drive & Spreadsheet berhasil diberikan.');
}

function syncAllData() {
  const currentEmail = getCurrentEmail();
  Logger.log(`[SYNC] Menjalankan sinkronisasi dengan akun: ${currentEmail}`);

  let targetSS;
  try {
    targetSS = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  } catch (e) {
    Logger.log(`[ERROR] Tidak dapat membuka Spreadsheet Target (${CONFIG.TARGET_SPREADSHEET_ID}). Pastikan akses Editor diberikan ke: ${currentEmail}`);
    throw new Error(`Akses ditolak pada Spreadsheet Target. Berikan hak akses 'Editor' ke ${currentEmail}`);
  }

  const results = [];

  for (let i = 0; i < CONFIG.SOURCES.length; i++) {
    const src = CONFIG.SOURCES[i];
    Logger.log(`[SYNC] Memproses sumber [${i + 1}/${CONFIG.SOURCES.length}]: ${src.name}`);

    let matrix = fetchSheetDataWithFallback(src);
    if (!matrix || matrix.length === 0) {
      Logger.log(`[WARN] Gagal mengambil data dari ${src.name}`);
      results.push({ name: src.name, status: 'FAILED', rows: 0 });
      continue;
    }

    if (src.filterValue !== undefined && src.filterColIndex !== undefined) {
      matrix = filterMatrix(matrix, src.filterColIndex, src.filterValue);
    }

    if (src.isMultiHeader) {
      matrix = processMultiHeader(matrix, src.headerRows || 4);
    }

    if (matrix.length > 0) {
      const tabName = src.targetTab || src.name;
      writeMatrixToSheet(targetSS, tabName, matrix);
      Logger.log(`[BERHASIL] ${src.name} -> Tab '${tabName}' (${matrix.length} baris)`);
      results.push({ name: src.name, status: 'SUCCESS', rows: matrix.length });
    }
  }

  Logger.log('[SYNC] Selesai sinkronisasi semua sumber: ' + JSON.stringify(results));
  return { success: true, timestamp: new Date().toISOString(), results: results };
}

function filterMatrix(matrix, colIndex, targetVal) {
  if (!matrix || matrix.length <= 1) return matrix;
  const header = matrix[0];
  const target = String(targetVal || '').toLowerCase().replace(/\s+/g, '');
  const filtered = matrix.slice(1).filter(row => {
    const val = String(row[colIndex] || '').trim().toLowerCase().replace(/\s+/g, '');
    return val === target;
  });
  return [header, ...filtered];
}

function processMultiHeader(matrix, headerRows) {
  if (!matrix || matrix.length < headerRows) return matrix;
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
      const range = sourceConfig.range ? sheet.getRange(sourceConfig.range) : sheet.getDataRange();
      values = range.getValues();
    }

    const cleaned = clean2DArray(values);
    if (cleaned.length > 0) {
      Logger.log(`[BERHASIL] ${sourceConfig.name} via Native: ${cleaned.length} baris`);
      return cleaned;
    }
  } catch (errNative) {
    Logger.log(`Native fetch gagal untuk ${sourceConfig.name}: ${errNative.message}`);
  }

  const encodedSheet = encodeURIComponent(sourceConfig.sheetName);

  try {
    const exportUrl = `https://docs.google.com/spreadsheets/d/${sourceConfig.spreadsheetId}/export?format=csv&sheet=${encodedSheet}`;
    const resExport = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });

    if (resExport.getResponseCode() === 200) {
      let parsed = Utilities.parseCsv(resExport.getContentText());
      if (sourceConfig.startRow && sourceConfig.startRow > 1 && parsed.length >= sourceConfig.startRow) {
        parsed = parsed.slice(sourceConfig.startRow - 1);
      }
      const cleaned = clean2DArray(parsed);
      if (cleaned.length > 0) {
        Logger.log(`[BERHASIL] ${sourceConfig.name} via CSV: ${cleaned.length} baris`);
        return cleaned;
      }
    }
  } catch (errExport) {
    Logger.log(`CSV export gagal untuk ${sourceConfig.name}: ${errExport.message}`);
  }

  try {
    const gvizRange = sourceConfig.range ? `&range=${sourceConfig.range}` : '';
    const gvizUrl = `https://docs.google.com/spreadsheets/d/${sourceConfig.spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodedSheet}${gvizRange}`;
    const token = ScriptApp.getOAuthToken();
    const resGviz = UrlFetchApp.fetch(gvizUrl, {
      muteHttpExceptions: true,
      headers: { 'Authorization': `Bearer ${token}` }
    });

    if (resGviz.getResponseCode() === 200) {
      let parsed = Utilities.parseCsv(resGviz.getContentText());
      if (sourceConfig.startRow && sourceConfig.startRow > 1 && !sourceConfig.range && parsed.length >= sourceConfig.startRow) {
        parsed = parsed.slice(sourceConfig.startRow - 1);
      }
      const cleaned = clean2DArray(parsed);
      if (cleaned.length > 0) {
        Logger.log(`[BERHASIL] ${sourceConfig.name} via GViz: ${cleaned.length} baris`);
        return cleaned;
      }
    }
  } catch (errGviz) {
    Logger.log(`GViz gagal untuk ${sourceConfig.name}: ${errGviz.message}`);
  }

  Logger.log(`[GAGAL] Tidak dapat membaca data dari ${sourceConfig.name} (${sourceConfig.spreadsheetId}).`);
  return [];
}

function writeMatrixToSheet(targetSS, tabName, data) {
  if (!data || data.length === 0) return;
  let sheet = targetSS.getSheetByName(tabName);
  if (!sheet) {
    sheet = targetSS.insertSheet(tabName);
  }
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
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
  Logger.log('[TRIGGER] Trigger 15 menit berhasil dibuat.');
}

function setupTrigger1Hour() {
  removeExistingTriggers();
  ScriptApp.newTrigger('syncAllData')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('[TRIGGER] Trigger 1 jam berhasil dibuat.');
}

function removeExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncAllData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
