const DASH_CONFIG = {
  MAIN_SPREADSHEET_ID: '146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4',
  DASHBOARD: {
    spreadsheetId: '1OrmtFMggqx0j5uW_X5Nxfo23ty61F7m6ppKWEPvPu0s',
    sheetName: 'DASHBOARD KHUSUS',
    startRow: 6,
    maxRows: 8
  },
  LEADERBOARD: {
    spreadsheetId: '146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4',
    sheetName: 'Leaderboard 5R'
  },
  WEBDATA: {
    spreadsheetId: '1P34UU9Q1pN4afM9L8Mn23LD90QdnbQ7Y',
    sheetName: 'WebData_Produk',
    targetPlant: 'cp3'
  },
  WEBDATA_AUDIT: {
    spreadsheetId: '1P34UU9Q1pN4afM9L8Mn23LD90QdnbQ7Y',
    sheetName: 'WebData',
    fallbackTab: 'WebData_Audit'
  },
  SR_AUDIT: {
    spreadsheetId: '1TpD67HaxPtkHgWfVRdQ-L5nKckYAIo_5SPEiJFzRBYI',
    sheetName: 'Update (CP-3)',
    fallbackTab: '5R_Audit'
  },
  PHOTO_FOLDER: 'Dashboard5R_Foto',
  ADMIN_KEY: 'cp3admin123',
  CACHE_SECONDS: 600
};

function doGet(e) {
  const params = (e && e.parameter) || {};
  const page = String(params.page || 'dash').toLowerCase();
  if (page === 'admin') {
    if (String(params.key || '') !== DASH_CONFIG.ADMIN_KEY) {
      return ContentService.createTextOutput('Akses ditolak: key salah.').setMimeType(ContentService.MimeType.TEXT);
    }
    return HtmlService.createTemplateFromFile('Admin').evaluate()
      .setTitle('Input Top 5 & Foreman')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Central Production 3 Majalengka')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getData() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('dash_all_v3');
  if (hit) return JSON.parse(hit);
  const payload = {
    ok: true,
    timestamp: new Date().toISOString(),
    dashboard: readDashboard(),
    leaderboard: readLeaderboard(),
    webdata: readWebData(),
    gmpScore: readGMPScore(),
    srScore: read5RScore()
  };
  try {
    cache.put('dash_all_v3', JSON.stringify(payload), DASH_CONFIG.CACHE_SECONDS);
  } catch (err) {}
  return payload;
}

function getLeaderboard() {
  return readLeaderboard();
}

function clearLeaderboard() {
  const cfg = DASH_CONFIG.LEADERBOARD;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = findSheet(ss, cfg.sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + cfg.sheetName);
  for (let i = 0; i < 5; i++) {
    sheet.getRange(2 + i, 2).setValue('');
    sheet.getRange(2 + i, 3).setValue('');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 8, lastRow - 1, 6).clearContent();
  }
  clearDashCache();
  return { ok: true };
}

function saveTop5(data) {
  const cfg = DASH_CONFIG.LEADERBOARD;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = findSheet(ss, cfg.sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + cfg.sheetName);
  const late = (data && data.late) || [];
  const overtime = (data && data.overtime) || [];
  for (let i = 0; i < 5; i++) {
    sheet.getRange(2 + i, 2).setValue(late[i] || '');
    sheet.getRange(2 + i, 3).setValue(overtime[i] || '');
  }
  const foremen = (data && data.foremen) || {};
  Object.keys(foremen).forEach(function (prod) {
    const row = findProdRow(sheet, prod);
    if (!row) return;
    const f = foremen[prod] || {};
    sheet.getRange(row, 8).setValue(f.f1 || '');
    sheet.getRange(row, 10).setValue(f.f2 || '');
    sheet.getRange(row, 12).setValue(f.f3 || '');
  });
  clearDashCache();
  return { ok: true };
}

function uploadPhoto(prod, slot, dataUrl, filename) {
  if (!prod || !slot || !dataUrl) throw new Error('Data foto tidak lengkap.');
  const s = [1, 2, 3].indexOf(Number(slot));
  if (s === -1) throw new Error('Slot harus 1/2/3.');
  
  let url = '';
  // Jika berupa URL web langsung
  if (String(dataUrl).startsWith('http://') || String(dataUrl).startsWith('https://')) {
    url = dataUrl;
  } else {
    // Jika base64 upload ke Drive
    const m = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
    if (!m) throw new Error('Format data foto salah.');
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], filename || ('foto_' + Date.now() + '.jpg'));
    const folder = getOrCreateFolder(DASH_CONFIG.PHOTO_FOLDER);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
  }

  const cfg = DASH_CONFIG.LEADERBOARD;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = findSheet(ss, cfg.sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + cfg.sheetName);
  const row = findProdRow(sheet, prod);
  if (!row) throw new Error('Produk tidak ditemukan di sheet: ' + prod);
  sheet.getRange(row, [9, 11, 13][s]).setValue(url);
  clearDashCache();
  return { ok: true, url: url };
}

function findSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  const want = String(name || '').toLowerCase().trim();
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    if (String(all[i].getName() || '').toLowerCase().trim() === want) return all[i];
  }
  return null;
}

function findProdRow(sheet, prod) {
  const want = String(prod || '').toUpperCase().trim();
  if (!want) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const vals = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').toUpperCase().trim() === want) return 2 + i;
  }
  return 0;
}

function getOrCreateFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function clearDashCache() {
  try {
    CacheService.getScriptCache().removeAll(['dash_all', 'dash_all_v2', 'dash_all_v3']);
  } catch (err) {}
}

function readDashboard() {
  const cfg = DASH_CONFIG.DASHBOARD;
  let ss = null;
  let sheet = null;

  try {
    const mainSs = SpreadsheetApp.openById(DASH_CONFIG.MAIN_SPREADSHEET_ID);
    const mainSheet = findSheet(mainSs, cfg.sheetName);
    if (mainSheet && mainSheet.getLastRow() >= (cfg.startRow || 6)) {
      ss = mainSs;
      sheet = mainSheet;
    }
  } catch (e) {}

  if (!sheet) {
    ss = SpreadsheetApp.openById(cfg.spreadsheetId);
    sheet = findSheet(ss, cfg.sheetName);
  }
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + cfg.sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // Deteksi dinamis baris header & baris awal data (toleran beda posisi baris 1 vs baris 2)
  const sampleRows = Math.min(lastRow, 12);
  const sample = sheet.getRange(1, 1, sampleRows, Math.min(lastCol, 5)).getValues();
  let headerStartRow = 1;
  let startRow = 6;

  for (let r = 0; r < sample.length; r++) {
    const c0 = String(sample[r][0] || '').toUpperCase();
    const c1 = String(sample[r][1] || '').toUpperCase();
    if (c0.indexOf('ISOWEEK') !== -1 || c1.indexOf('TANGGAL') !== -1 || c0.indexOf('WEEK') !== -1) {
      headerStartRow = r + 1;
      break;
    }
  }

  for (let r = headerStartRow; r < sample.length; r++) {
    const v1 = sample[r][1];
    if (v1 instanceof Date || (typeof v1 === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v1)) || (typeof v1 === 'string' && /^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}/.test(v1))) {
      startRow = r + 1;
      break;
    }
  }
  if (startRow <= headerStartRow) startRow = headerStartRow + 4;
  const numHeaderRows = startRow - headerStartRow;

  if (lastRow < startRow) return [];
  const headerMatrix = sheet.getRange(headerStartRow, 1, numHeaderRows, lastCol).getValues();
  for (let r = 0; r < headerMatrix.length; r++) {
    let lastVal = '';
    for (let c = 2; c < headerMatrix[r].length; c++) {
      const v = String(headerMatrix[r][c] || '').trim();
      if (v !== '') { lastVal = v; }
      else if (lastVal !== '') { headerMatrix[r][c] = lastVal; }
    }
    if (!headerMatrix[r][0]) headerMatrix[r][0] = 'ISOWEEK';
    if (!headerMatrix[r][1]) headerMatrix[r][1] = 'TANGGAL';
  }
  const colNames = [];
  for (let c = 0; c < lastCol; c++) {
    if (c === 0) { colNames.push('isoweek'); continue; }
    if (c === 1) { colNames.push('tanggal'); continue; }
    const parts = headerMatrix.map(function (r) { return String(r[c] || '').trim(); }).filter(function (p) { return p !== ''; });
    const unique = parts.filter(function (p, i) { return parts.indexOf(p) === i; });
    colNames.push(sanitizeKey(unique.join('__') || 'col_' + (c + 1)));
  }
  const endRow = findDataEndRow(sheet, startRow, lastRow);
  const wantRows = Math.min(endRow - startRow + 1, cfg.maxRows || 8);
  const safeFrom = Math.max(startRow, endRow - wantRows + 1);
  const dataMatrix = sheet.getRange(safeFrom, 1, endRow - safeFrom + 1, lastCol).getValues();
  const records = [];
  dataMatrix.forEach(function (row, idx) {
    const hasData = row.some(function (cell) { return cell !== '' && cell !== null && cell !== undefined; });
    if (!hasData) return;
    const obj = { row_index: safeFrom + idx };
    colNames.forEach(function (key, c) {
      const val = row[c];
      obj[key] = (val instanceof Date) ? val.toISOString() : ((val === '') ? null : val);
    });
    records.push(obj);
  });
  return records;
}

function readLeaderboard() {
  const cfg = DASH_CONFIG.LEADERBOARD;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = findSheet(ss, cfg.sheetName);
  if (!sheet) {
    const names = ss.getSheets().map(function (s) { return s.getName(); }).join(' | ');
    throw new Error('Sheet "' + cfg.sheetName + '" tidak ditemukan. Sheet yang ada: ' + names);
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { ranking: [], foremen: [] };
  const raw = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const norm = function (v) { return String(v || '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
  const header = raw[0].map(norm);
  const findCol = function (names, fallback) {
    for (let i = 0; i < header.length; i++) {
      if (names.indexOf(header[i]) !== -1) return i;
    }
    return fallback;
  };
  const cRank = findCol(['rank', 'no', 'ranking'], 0);
  const cTopLate = findCol(['toplate', 'topterlambat', 'late'], 1);
  const cTopOt = findCol(['topovertime', 'overtime', 'toplembur', 'normalovertime'], 2);
  const cProd = findCol(['prod', 'produk', 'product'], 6);
  const foremanCols = [];
  for (let i = 0; i < header.length; i++) {
    if (/^foreman\d*$/.test(header[i]) || header[i] === 'foreman') foremanCols.push(i);
  }
  const imgCols = [];
  for (let i = 0; i < header.length; i++) {
    if (/^(img|image|foto|photo|linkimg|urlimg)$/.test(header[i])) imgCols.push(i);
  }
  let pairs = [];
  if (foremanCols.length > 0) {
    foremanCols.slice(0, 3).forEach(function (fc, n) {
      let ic = -1;
      for (let k = 0; k < imgCols.length; k++) {
        if (imgCols[k] > fc && (k === imgCols.length - 1 || imgCols[k + 1] > (foremanCols[n + 1] || 999))) { ic = imgCols[k]; break; }
      }
      if (ic === -1 && imgCols[n] !== undefined) ic = imgCols[n];
      pairs.push({ f: fc, img: ic });
    });
  } else {
    pairs = [{ f: 7, img: 8 }, { f: 9, img: 10 }, { f: 11, img: 12 }];
  }
  while (pairs.length < 3) pairs.push({ f: 7 + pairs.length * 2, img: 8 + pairs.length * 2 });
  const cellStr = function (row, c) {
    if (c === -1 || c === undefined || c >= row.length) return null;
    const v = row[c];
    if (v === '' || v === null || v === undefined) return null;
    return String(v).trim() || null;
  };
  const ranking = [];
  for (let r = 1; r < raw.length; r++) {
    const row = raw[r];
    const topLate = cellStr(row, cTopLate);
    const topOt = cellStr(row, cTopOt);
    if (!topLate && !topOt) continue;
    const rankNum = parseInt(cellStr(row, cRank), 10);
    ranking.push({
      row_index: r + 1,
      rank: isNaN(rankNum) ? (ranking.length + 1) : rankNum,
      top_late: topLate,
      top_overtime: topOt
    });
    if (ranking.length >= 20) break;
  }
  const foremen = [];
  for (let r = 1; r < raw.length; r++) {
    const row = raw[r];
    const prod = cellStr(row, cProd);
    if (!prod) continue;
    foremen.push({
      row_index: r + 1,
      prod: prod.toUpperCase(),
      foreman_1: cellStr(row, pairs[0].f),
      img_1: cellStr(row, pairs[0].img),
      foreman_2: cellStr(row, pairs[1].f),
      img_2: cellStr(row, pairs[1].img),
      foreman_3: cellStr(row, pairs[2].f),
      img_3: cellStr(row, pairs[2].img)
    });
    if (foremen.length >= 50) break;
  }
  return { ranking: ranking, foremen: foremen };
}

function findDataEndRow(sheet, startRow, lastRow) {
  try {
    const t = new Date();
    t.setDate(t.getDate() - 2);
    const target = Utilities.formatDate(t, 'Asia/Jakarta', 'yyyyMMdd');
    const vals = sheet.getRange(startRow, 2, lastRow - startRow + 1, 1).getValues();
    let endRow = 0;
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i][0];
      if (v === '' || v === null || v === undefined) continue;
      const ymd = Utilities.formatDate(new Date(v), 'Asia/Jakarta', 'yyyyMMdd');
      if (ymd <= target) endRow = startRow + i;
    }
    if (endRow >= startRow) return endRow;
  } catch (err) {}
  return lastRow;
}

function readWebData() {
  try {
    const cfg = DASH_CONFIG.WEBDATA;
    let ss = null;
    let sheet = null;

    try {
      const mainSs = SpreadsheetApp.openById(DASH_CONFIG.MAIN_SPREADSHEET_ID);
      const mainSheet = findSheet(mainSs, cfg.sheetName);
      if (mainSheet && mainSheet.getLastRow() >= 2) {
        ss = mainSs;
        sheet = mainSheet;
      }
    } catch (e) {}

    if (!sheet) {
      ss = SpreadsheetApp.openById(cfg.spreadsheetId);
      sheet = findSheet(ss, cfg.sheetName);
    }
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return [];
    const raw = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    if (raw.length < 2) return [];
    const header = raw[0];
    const target = String(cfg.targetPlant || 'cp3').toLowerCase().replace(/\s+/g, '');
    const records = [];
    for (let r = 1; r < raw.length; r++) {
      const row = raw[r];
      const plant = String(row[2] || '').toLowerCase().trim().replace(/\s+/g, '');
      if (plant && plant !== target) continue;
      const obj = { row_index: r + 1 };
      header.forEach(function (h, c) {
        const key = sanitizeKey(String(h || 'col_' + (c + 1)));
        const val = row[c];
        obj[key] = (val instanceof Date) ? val.toISOString() : ((val === '') ? null : val);
      });
      records.push(obj);
      if (records.length >= 200) break;
    }
    return records;
  } catch (err) {
    return [];
  }
}

function readGMPScore() {
  try {
    const cfg = DASH_CONFIG.WEBDATA_AUDIT;
    let ss = null, sheet = null;
    try {
      const mainSs = SpreadsheetApp.openById(DASH_CONFIG.MAIN_SPREADSHEET_ID);
      sheet = findSheet(mainSs, cfg.fallbackTab) || findSheet(mainSs, cfg.sheetName);
      if (sheet && sheet.getLastRow() >= 2) ss = mainSs;
    } catch (e) {}

    if (!sheet) {
      ss = SpreadsheetApp.openById(cfg.spreadsheetId);
      sheet = findSheet(ss, cfg.sheetName);
    }
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 4) return null;

    const data = sheet.getRange(1, 1, lastRow, Math.min(lastCol, 10)).getValues();
    let gSum = 0, gCount = 0;
    for (let i = 1; i < data.length; i++) {
      const cat = String(data[i][2] || '').toLowerCase();
      const score = parseFloat(data[i][3]);
      if (!isNaN(score) && cat.indexOf('gmp') !== -1) {
        gSum += (score <= 1 && score > 0 ? score * 100 : score);
        gCount++;
      }
    }
    if (gCount > 0) return (gSum / gCount);
    return null;
  } catch (err) {
    return null;
  }
}

function read5RScore() {
  try {
    const cfg = DASH_CONFIG.SR_AUDIT;
    let ss = null, sheet = null;
    try {
      const mainSs = SpreadsheetApp.openById(DASH_CONFIG.MAIN_SPREADSHEET_ID);
      sheet = findSheet(mainSs, cfg.fallbackTab) || findSheet(mainSs, cfg.sheetName);
      if (sheet && sheet.getLastRow() >= 2) ss = mainSs;
    } catch (e) {}

    if (!sheet) {
      ss = SpreadsheetApp.openById(cfg.spreadsheetId);
      sheet = findSheet(ss, cfg.sheetName);
    }
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 7) return null;

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    let srSum = 0, srCount = 0;
    for (let i = 1; i < data.length; i++) {
      let score = parseFloat(data[i][6]);
      if (isNaN(score) && data[i].length > 48) {
        score = parseFloat(data[i][48]);
      }
      if (!isNaN(score) && score > 0) {
        srSum += (score <= 1 ? score * 100 : score);
        srCount++;
      }
    }
    if (srCount > 0) return (srSum / srCount);
    return null;
  } catch (err) {
    return null;
  }
}

function inspectAudit() {
  const ss = SpreadsheetApp.openById('1_Zw21JaDcsURJU0Zf7pyGxFqPYa5hB0-hkfgqUgGY-0');
  const out = [];
  const sheets = ss.getSheets();
  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    const lr = Math.min(sh.getLastRow(), 30);
    const lc = Math.min(sh.getLastColumn(), 30);
    if (lr < 1) {
      out.push({ sheet: sh.getName(), gid: sh.getSheetId(), rows: 0 });
      continue;
    }
    const vals = sh.getRange(1, 1, lr, lc).getValues();
    const preview = [];
    for (let r = 0; r < vals.length && preview.length < 6; r++) {
      const row = vals[r].map(function (v) { return String(v == null ? '' : v).trim().substring(0, 30); });
      if (row.some(function (c) { return c !== ''; })) preview.push(row);
    }
    out.push({ sheet: sh.getName(), gid: sh.getSheetId(), rows: sh.getLastRow(), cols: sh.getLastColumn(), preview: preview });
  }
  return { ok: true, sheets: out };
}

function sanitizeKey(str) {
  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9_]/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^_|_$/g, '')
    .substring(0, 63)
    || 'col';
}
