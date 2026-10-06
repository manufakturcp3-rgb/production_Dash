/**
 * ============================================================
 * BACKEND.JS — Web App penyedia data JSON untuk display Astro
 * ============================================================
 * Script ke-2 dari arsitektur tanpa Supabase:
 *   1. produksi/Kode.js = PENARIK (2 sumber -> 1 sheet target)
 *   2. backend/Backend.js = PENYEDIA (baca sheet -> JSON buat Astro)
 *
 * Cara pakai:
 *   1. Buat project Apps Script baru (standalone).
 *   2. Paste seluruh file ini -> Save.
 *   3. Deploy > New deployment > Web app:
 *        Execute as: Me
 *        Who has access: Anyone
 *      -> Copy URL .../exec (masukkan ke display sebagai
 *         PUBLIC_BACKEND_URL).
 *   4. Kalau edit kode lagi: Deploy > Manage deployments >
 *      Edit (icon pensil) > Version: New version (URL tetap sama).
 *
 * Endpoint:
 *   GET ?action=all          -> { dashboard, leaderboard }
 *   GET ?action=dashboard    -> { dashboard }
 *   GET ?action=leaderboard  -> { leaderboard }
 *   Tambah &refresh=1 untuk bypass cache (debug).
 * ============================================================
 */

const BACKEND_CONFIG = {
  DASHBOARD: {
    spreadsheetId: '1OrmtFMggqx0j5uW_X5Nxfo23ty61F7m6ppKWEPvPu0s',
    sheetName: 'DASHBOARD KHUSUS',
    startRow: 6,
    maxRows: 15 // hanya N baris terbaru (display cuma butuh baris terakhir)
  },
  LEADERBOARD: {
    spreadsheetId: '146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4',
    sheetName: 'Leaderboard 5R'
  },
  CACHE_SECONDS: 600 // 10 menit
};

// ============================================================
// ENTRY POINT WEB APP
// ============================================================
function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || 'all').toLowerCase();
  const noCache = String(params.refresh || '') === '1';

  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'backend_' + action;

    if (!noCache) {
      const hit = cache.get(cacheKey);
      if (hit) return jsonOut(hit);
    }

    const payload = buildPayload(action);
    const text = JSON.stringify(payload);

    try {
      cache.put(cacheKey, text, BACKEND_CONFIG.CACHE_SECONDS);
    } catch (cacheErr) {
      // Payload > 100KB (limit ScriptCache) -> sajikan tanpa cache
    }
    return jsonOut(text);
  } catch (err) {
    return jsonOut(JSON.stringify({ ok: false, error: String((err && err.message) || err) }));
  }
}

function jsonOut(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function buildPayload(action) {
  const payload = { ok: true, timestamp: new Date().toISOString() };
  if (action === 'all' || action === 'dashboard') payload.dashboard = readDashboard();
  if (action === 'all' || action === 'leaderboard') payload.leaderboard = readLeaderboard();
  return payload;
}

// ============================================================
// DASHBOARD KHUSUS — mulai baris 6, header multi-baris (2-5)
// Format key SAMA seperti Featch.js: produk__kategori__shift
// ============================================================
function readDashboard() {
  const cfg = BACKEND_CONFIG.DASHBOARD;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + cfg.sheetName);

  const startRow = cfg.startRow || 6;
  const headerStartRow = 2;
  const numHeaderRows = startRow - headerStartRow;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow || lastCol < 1) return [];

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

  // Ambil N baris TERBARU saja (payload kecil, display cuma butuh yang terakhir)
  const totalRows = lastRow - startRow + 1;
  const take = Math.min(totalRows, cfg.maxRows || 15);
  const fromRow = lastRow - take + 1;
  const dataMatrix = sheet.getRange(fromRow, 1, take, lastCol).getValues();

  const records = [];
  dataMatrix.forEach(function (row, idx) {
    const hasData = row.some(function (cell) { return cell !== '' && cell !== null && cell !== undefined; });
    if (!hasData) return;
    const obj = { row_index: fromRow + idx };
    colNames.forEach(function (key, c) {
      const val = row[c];
      obj[key] = (val instanceof Date) ? val.toISOString() : ((val === '') ? null : val);
    });
    records.push(obj);
  });
  return records;
}

// ============================================================
// LEADERBOARD 5R — Top Late/Overtime + Foreman per Produk + IMG URL
// Layout (baris 1 = header):
//   A=RANK | B=TOP_LATE | C=TOP_OVERTIME | ... | G=PROD |
//   H=FOREMAN_1 | I=IMG | J=FOREMAN_2 | K=IMG | L=FOREMAN_3 | M=IMG
// ============================================================
function readLeaderboard() {
  const cfg = BACKEND_CONFIG.LEADERBOARD;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);

  let sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) {
    const want = String(cfg.sheetName || '').toLowerCase().trim();
    const all = ss.getSheets();
    for (let i = 0; i < all.length; i++) {
      if (String(all[i].getName() || '').toLowerCase().trim() === want) { sheet = all[i]; break; }
    }
  }
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

// ============================================================
// UTIL
// ============================================================
function sanitizeKey(str) {
  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9_]/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^_|_$/g, '')
    .substring(0, 63)
    || 'col';
}

/** Hapus cache manual (debug) */
function clearBackendCache() {
  CacheService.getScriptCache().removeAll(['backend_all', 'backend_dashboard', 'backend_leaderboard']);
  Logger.log('Backend cache dibersihkan.');
}
